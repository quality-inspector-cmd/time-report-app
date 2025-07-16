import pandas as pd
import datetime
import os
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference, LineChart # Import LineChart
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF
from matplotlib import pyplot as plt
import tempfile
import re
import shutil

# Hàm hỗ trợ làm sạch tên file
def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)

def setup_paths():
    today = datetime.datetime.today().strftime('%Y%m%d')
    return {
        'template_file': "Time_report.xlsm",
        'output_file': f"Time_report_Standard_{today}.xlsx",
        'pdf_report': f"Time_report_Standard_{today}.pdf",
        'comparison_output_file': f"Time_report_Comparison_{today}.xlsx",
        'comparison_pdf_report': f"Time_report_Comparison_{today}.pdf"
    }

def read_configs(path_dict):
    try:
        # Load the configuration data from the template file
        year_mode_df = pd.read_excel(path_dict['template_file'], sheet_name='Config_Year_Mode', engine='openpyxl')
        project_filter_df = pd.read_excel(path_dict['template_file'], sheet_name='Config_Project_Filter', engine='openpyxl')

        # Extract mode, year, and months, handling potential NaNs
        mode_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'mode', 'Value']
        mode = str(mode_row.values[0]).strip().lower() if not mode_row.empty and pd.notna(mode_row.values[0]) else 'year'

        year_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'year', 'Value']
        # Set default year to current year if not found or NaN
        year = int(year_row.values[0]) if not year_row.empty and pd.notna(year_row.values[0]) else datetime.datetime.now().year
        
        months_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'months', 'Value']
        # Split months string and capitalize each, handle empty or NaN values
        months = [m.strip().capitalize() for m in str(months_row.values[0]).split(',')] if not months_row.empty and pd.notna(months_row.values[0]) else []
        
        # Ensure 'Include' column is string type for consistent comparison
        if 'Include' in project_filter_df.columns:
            project_filter_df['Include'] = project_filter_df['Include'].astype(str)

        return {
            'mode': mode,
            'year': year,
            'months': months,
            'project_filter_df': project_filter_df
        }
    except FileNotFoundError:
        print(f"Error: Template file not found at {path_dict['template_file']}")
        # Return default configuration if file not found
        return {'mode': 'year', 'year': datetime.datetime.now().year, 'months': [], 'project_filter_df': pd.DataFrame(columns=['Project Name', 'Include'])}
    except Exception as e:
        print(f"Error reading config: {e}")
        # Return default configuration on other errors
        return {'mode': 'year', 'year': datetime.datetime.now().year, 'months': [], 'project_filter_df': pd.DataFrame(columns=['Project Name', 'Include'])}


def load_raw_data(path_dict):
    try:
        df = pd.read_excel(path_dict['template_file'], sheet_name='Raw Data', engine='openpyxl')
        # Clean and standardize column names
        df.columns = df.columns.str.strip()
        df.rename(columns={'Hou': 'Hours', 'Team member': 'Employee', 'Project Name': 'Project name'}, inplace=True)
        
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df['Year'] = df['Date'].dt.year
        df['MonthName'] = df['Date'].dt.month_name()
        df['Week'] = df['Date'].dt.isocalendar().week.astype(int) # Ensure Week is int
        return df
    except Exception as e:
        print(f"Error loading raw data: {e}")
        return pd.DataFrame() # Return empty DataFrame on error

def apply_filters(df, config):
    df_filtered = df.copy()

    # Filter by year (either single year or multiple years, though UI uses single year for standard report)
    if 'years' in config and config['years']: # This key is used in comparison report config
        df_filtered = df_filtered[df_filtered['Year'].isin(config['years'])]
    elif 'year' in config and config['year']: # This key is used in standard report config
        df_filtered = df_filtered[df_filtered['Year'] == config['year']]

    # Filter by months
    if config['months']:
        df_filtered = df_filtered[df_filtered['MonthName'].isin(config['months'])]

    # Filter projects based on the 'project_filter_df' which is configured in Streamlit
    if not config['project_filter_df'].empty:
        selected_project_names = config['project_filter_df']['Project Name'].tolist()
        df_filtered = df_filtered[df_filtered['Project name'].isin(selected_project_names)]
    else:
        # If no projects are selected in the UI or all are excluded, return empty
        return pd.DataFrame(columns=df.columns) 

    return df_filtered

def export_report(df, config, path_dict):
    mode = config.get('mode', 'year')
    
    # Define required columns based on mode
    if mode == 'year':
        required_cols = ['Year', 'Project name', 'Hours']
        groupby_cols = ['Year', 'Project name']
    elif mode == 'month':
        required_cols = ['Year', 'MonthName', 'Project name', 'Hours']
        groupby_cols = ['Year', 'MonthName', 'Project name']
    else: # week mode
        required_cols = ['Year', 'Week', 'Project name', 'Hours']
        groupby_cols = ['Year', 'Week', 'Project name']

    # Check for missing columns before grouping
    for col in required_cols:
        if col not in df.columns:
            raise KeyError(f"Missing required column for grouping in '{mode}' mode: {col}")

    if df.empty:
        print("Warning: Filtered DataFrame is empty, no report will be generated.")
        return False # Return False to indicate failure

    # Group and summarize data
    summary = df.groupby(groupby_cols)['Hours'].sum().reset_index()

    # Create Excel report
    try:
        with pd.ExcelWriter(path_dict['output_file'], engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name='Summary', index=False)

        wb = load_workbook(path_dict['output_file'])
        ws = wb['Summary']
        max_row = ws.max_row
        
        # Determine column references for chart based on mode
        data_col_idx = summary.columns.get_loc('Hours') + 1
        cats_col_idx = summary.columns.get_loc('Project name') + 1

        # Data for chart needs to start from row 2 (after header)
        data_ref = Reference(ws, min_col=data_col_idx, min_row=2, max_row=max_row)
        cats_ref = Reference(ws, min_col=cats_col_idx, min_row=2, max_row=max_row)

        chart = BarChart()
        chart.title = f"Total Hours by Project ({mode})"
        chart.x_axis.title = "Project"
        chart.y_axis.title = "Hours"
        
        chart.add_data(data_ref, titles_from_data=False) 
        chart.set_categories(cats_ref)
        ws.add_chart(chart, "F2")

        # Create separate sheets for each project with task summaries and raw data
        for project in df['Project name'].unique():
            df_proj = df[df['Project name'] == project]
            # Sanitize project name for sheet title, truncate if too long
            sheet_title = sanitize_filename(project)[:31]
            ws_proj = wb.create_sheet(title=sheet_title)

            # Summary by Task
            summary_task = df_proj.groupby('Task')['Hours'].sum().reset_index().sort_values('Hours', ascending=False)
            
            # Write header for summary_task
            ws_proj.append(['Task', 'Hours'])
            # Write data for summary_task
            for row_data in dataframe_to_rows(summary_task, index=False, header=False):
                ws_proj.append(row_data)

            # Add chart for Task summary
            chart_task = BarChart()
            chart_task.title = f"{project} - Hours by Task"
            chart_task.x_axis.title = "Task"
            chart_task.y_axis.title = "Hours"
            task_len = len(summary_task)
            
            # Data and category references for task chart
            data_ref_task = Reference(ws_proj, min_col=2, min_row=1, max_row=task_len + 1) # Hours column, including header
            cats_ref_task = Reference(ws_proj, min_col=1, min_row=2, max_row=task_len + 1) # Task names, excluding header
            chart_task.add_data(data_ref_task, titles_from_data=True)
            chart_task.set_categories(cats_ref_task)
            ws_proj.add_chart(chart_task, f"E1")

            # Append raw data for the project below the task summary and chart
            start_row_raw_data = task_len + 5
            for r_idx, r in enumerate(dataframe_to_rows(df_proj, index=False, header=True)):
                for c_idx, cell_val in enumerate(r):
                    ws_proj.cell(row=start_row_raw_data + r_idx, column=c_idx + 1, value=cell_val)
        
        # Create Config_Info sheet
        ws_config = wb.create_sheet("Config_Info")
        ws_config['A1'], ws_config['B1'] = "Mode", config.get('mode', 'N/A').capitalize()
        # Ensure 'years' is handled for single year scenario correctly for display
        ws_config['A2'], ws_config['B2'] = "Year(s)", ', '.join(map(str, config.get('years', []))) if config.get('years') else str(config.get('year', 'N/A'))
        ws_config['A3'], ws_config['B3'] = "Months", ', '.join(config.get('months', [])) if config.get('months') else "All"
        
        if 'project_filter_df' in config and not config['project_filter_df'].empty:
            ws_config['A4'], ws_config['B4'] = "Projects", ', '.join(config['project_filter_df']['Project Name'])
        else:
            ws_config['A4'], ws_config['B4'] = "Projects", "No projects selected or found"

        # Remove 'Raw Data' sheet if it exists and is from the template
        if 'Raw Data' in wb.sheetnames:
            del wb['Raw Data']

        wb.save(path_dict['output_file'])
        return True
    except Exception as e:
        print(f"Error exporting standard report: {e}")
        return False

def export_pdf_report(df, config, path_dict):
    today_str = datetime.datetime.today().strftime("%Y-%m-%d")

    def create_pdf_from_charts(charts_data, output_path, title, config_info, logo_path="triac_logo.png"):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        # Use standard Helvetica font
        pdf.set_font('helvetica', 'B', 16) 
        
        pdf.add_page()
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=10, y=10, w=30)
        pdf.ln(40)
        pdf.cell(0, 10, title, ln=True, align='C')
        pdf.set_font("helvetica", '', 12) 
        pdf.ln(5)
        pdf.cell(0, 10, f"Generated on: {today_str}", ln=True, align='C')
        pdf.ln(10)
        pdf.set_font("helvetica", '', 11) 
        for key, value in config_info.items():
            pdf.cell(0, 7, f"{key}: {value}", ln=True, align='C')

        for img_path, chart_title, page_project_name in charts_data:
            if img_path and os.path.exists(img_path):
                pdf.add_page()
                if os.path.exists(logo_path):
                    pdf.image(logo_path, x=10, y=8, w=25)
                pdf.set_font("helvetica", 'B', 11) 
                pdf.set_y(35)
                if page_project_name:
                    pdf.cell(0, 10, f"Project: {page_project_name}", ln=True, align='C')
                pdf.cell(0, 10, chart_title, ln=True, align='C')
                pdf.image(img_path, x=10, y=45, w=190)

        pdf.output(output_path, "F")
        print(f"DEBUG: PDF report generated at {output_path}")

    tmp_dir = tempfile.mkdtemp()
    charts_for_pdf = []

    try:
        # Get unique projects from the filtered DataFrame
        projects = df['Project name'].unique() 

        config_info = {
            "Mode": config.get('mode', 'N/A').capitalize(),
            "Years": ', '.join(map(str, config.get('years', []))) or str(config.get('year', 'N/A')),
            "Months": ', '.join(config.get('months', [])) or "All",
            "Projects Included": ', '.join(config['project_filter_df']['Project Name']) if 'project_filter_df' in config and not config['project_filter_df'].empty else "No projects selected or found"
        }

        # Matplotlib configuration to avoid Vietnamese font issues
        plt.rcParams['font.family'] = 'sans-serif' # Use a generic sans-serif font
        plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'Liberation Sans'] # Prioritize common sans-serif fonts
        plt.rcParams['axes.unicode_minus'] = False # This is important for displaying minus signs correctly

        for project in projects:
            safe_project = sanitize_filename(project)
            df_proj = df[df['Project name'] == project]

            # Chart for Workcentre
            if 'Workcentre' in df_proj.columns and not df_proj['Workcentre'].empty:
                workcentre_summary = df_proj.groupby('Workcentre')['Hours'].sum().sort_values(ascending=False)
                if not workcentre_summary.empty:
                    fig, ax = plt.subplots(figsize=(10, 5))
                    workcentre_summary.plot(kind='barh', color='skyblue', ax=ax)
                    ax.set_title(f"{project} - Hours by Workcentre", fontsize=9)
                    ax.tick_params(axis='y', labelsize=8)
                    wc_img_path = os.path.join(tmp_dir, f"{safe_project}_wc.png")
                    plt.tight_layout()
                    fig.savefig(wc_img_path, dpi=150)
                    plt.close(fig)
                    charts_for_pdf.append((wc_img_path, f"{project} - Hours by Workcentre", project))

            # Chart for Task
            if 'Task' in df_proj.columns and not df_proj['Task'].empty:
                task_summary = df_proj.groupby('Task')['Hours'].sum().sort_values(ascending=False)
                if not task_summary.empty:
                    fig, ax = plt.subplots(figsize=(10, 6))
                    task_summary.plot(kind='barh', color='lightgreen', ax=ax)
                    ax.set_title(f"{project} - Hours by Task", fontsize=9)
                    ax.tick_params(axis='y', labelsize=8)
                    task_img_path = os.path.join(tmp_dir, f"{safe_project}_task.png")
                    plt.tight_layout()
                    fig.savefig(task_img_path, dpi=150)
                    plt.close(fig)
                    charts_for_pdf.append((task_img_path, f"{project} - Hours by Task", project))

        create_pdf_from_charts(charts_for_pdf, path_dict['pdf_report'], "TRIAC TIME REPORT - STANDARD", config_info)
        return True
    except Exception as e:
        print(f"Error creating PDF report: {e}")
        return False
    finally:
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)


# =========================================================================
# NEW FUNCTIONS FOR COMPARISON REPORT
# =========================================================================

def apply_comparison_filters(df_raw, comparison_config, comparison_mode):
    years = comparison_config.get('years', [])
    months = comparison_config.get('months', [])
    selected_projects = comparison_config.get('selected_projects', [])

    df_filtered = df_raw.copy()

    if years:
        df_filtered = df_filtered[df_filtered['Year'].isin(years)]
    
    if months:
        df_filtered = df_filtered[df_filtered['MonthName'].isin(months)]
    
    if selected_projects:
        df_filtered = df_filtered[df_filtered['Project name'].isin(selected_projects)]
    else: 
        return pd.DataFrame(), "Please select at least one project for comparison."

    if df_filtered.empty:
        return pd.DataFrame(), f"No data found for comparison mode: {comparison_mode} with current selections."

    title = "" # Initialize title

    if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
        if len(years) != 1 or len(months) != 1 or len(selected_projects) < 2:
            return pd.DataFrame(), "Please select ONE year, ONE month and at least TWO projects for this mode."
        
        df_comparison = df_filtered.groupby('Project name')['Hours'].sum().reset_index()
        df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
        title = f"Hours Comparison Between Projects in {months[0]}, Year {years[0]}"
        return df_comparison, title

    elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        if len(years) != 1 or len(selected_projects) < 2:
            return pd.DataFrame(), "Please select ONE year and at least TWO projects for this mode."
        
        df_comparison = df_filtered.groupby(['Project name', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
        
        # Ensure month columns are ordered correctly
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        existing_months = [m for m in month_order if m in df_comparison.columns]
        df_comparison = df_comparison[existing_months]

        df_comparison = df_comparison.reset_index().rename(columns={'index': 'Project Name'})
        
        title = f"Hours Comparison Between Projects in Year {years[0]} (by Month)"
        return df_comparison, title

    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
        if len(selected_projects) != 1 or (len(months) < 2 and len(years) < 2): 
            return pd.DataFrame(), "Please select ONE project and at least TWO months OR TWO years for comparison."
        
        # If only years are selected (and more than one year)
        if years and not months:
            df_comparison = df_filtered.groupby('Year')['Hours'].sum().reset_index()
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            df_comparison['Year'] = df_comparison['Year'].astype(str) # Convert year to string for consistent plotting index
            title = f"Hours for Project {selected_projects[0]} Over Years"
        # If months are selected (and more than one month), possibly across years
        elif months and not years: # Compare months across all selected years (if any) or all available years
            df_comparison = df_filtered.groupby(['MonthName'])['Hours'].sum().reset_index()
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            # Order months
            month_order_df = pd.DataFrame({'MonthName': ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']})
            df_comparison = pd.merge(month_order_df, df_comparison, on='MonthName', how='left').fillna(0)
            title = f"Hours for Project {selected_projects[0]} Over Months"
        # If both years and months are selected, compare for selected month(s) over selected year(s)
        elif years and months:
            df_comparison = df_filtered.groupby(['Year', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
            month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
            existing_months = [m for m in month_order if m in df_comparison.columns]
            df_comparison = df_comparison[existing_months]
            df_comparison = df_comparison.reset_index()
            title = f"Hours for Project {selected_projects[0]} Over Years and Months"
        else: # Should not happen with the previous checks, but for safety
            return pd.DataFrame(), "Invalid time selection for this comparison mode."

        # Add 'Total' row if there are multiple data columns (i.e., a pivoted table)
        if len(df_comparison.columns) > 2: # If there's a time column and multiple data columns
            # Calculate sum for numeric columns only, excluding the time identifier column
            numeric_cols = df_comparison.select_dtypes(include=['number']).columns
            df_comparison.loc['Total'] = df_comparison[numeric_cols].sum()
            if 'Year' in df_comparison.columns:
                df_comparison.loc['Total', 'Year'] = 'Total'
            elif 'MonthName' in df_comparison.columns:
                df_comparison.loc['Total', 'MonthName'] = 'Total'

        return df_comparison, title
        
    return pd.DataFrame(), "Invalid comparison mode."

def export_comparison_report(df_comparison, comparison_config, path_dict, comparison_mode):
    output_file = path_dict['comparison_output_file']
    
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Write a placeholder row if df_comparison is empty to prevent error during chart creation
            if df_comparison.empty:
                empty_df_for_excel = pd.DataFrame({"Message": ["No data to display for the selected filters."]})
                empty_df_for_excel.to_excel(writer, sheet_name='Comparison Report', index=False)
            else:
                df_comparison.to_excel(writer, sheet_name='Comparison Report', index=False)  

            wb = writer.book
            ws = wb['Comparison Report']

            # Add header and configuration info
            # Adjusting row indices as data is written from A1
            
            # Write configuration info slightly below the data if data is present, otherwise start from A1
            if not df_comparison.empty:
                # Find the actual last row of data in the sheet
                data_last_row = ws.max_row
                info_row = data_last_row + 2 # Start info 2 rows below the data
            else:
                info_row = 1 # If no data, start from row 1

            ws.merge_cells(start_row=info_row, start_column=1, end_row=info_row, end_column=4)
            ws.cell(row=info_row, column=1, value=f"COMPARISON REPORT: {comparison_mode}").font = ws.cell(row=info_row, column=1).font.copy(bold=True, size=14)
            info_row += 1

            ws.cell(row=info_row, column=1, value="Years:").font = ws.cell(row=info_row, column=1).font.copy(bold=True)
            ws.cell(row=info_row, column=2, value=', '.join(map(str, comparison_config.get('years', []))))
            info_row += 1
            ws.cell(row=info_row, column=1, value="Months:").font = ws.cell(row=info_row, column=1).font.copy(bold=True)
            ws.cell(row=info_row, column=2, value=', '.join(comparison_config.get('months', [])))
            info_row += 1
            ws.cell(row=info_row, column=1, value="Projects:").font = ws.cell(row=info_row, column=1).font.copy(bold=True)
            ws.cell(row=info_row, column=2, value=', '.join(comparison_config.get('selected_projects', [])))

            # Add Chart if df_comparison is not empty
            if not df_comparison.empty:
                chart = None
                data_start_row = 2 # Data starts from row 2 (after header)

                if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
                    chart = BarChart()
                    chart.title = "Hours Comparison by Project"
                    chart.x_axis.title = "Project"
                    chart.y_axis.title = "Hours"
                    
                    data_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Total Hours') + 1, min_row=data_start_row, max_row=data_start_row + len(df_comparison) - 1)
                    cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Project name') + 1, min_row=data_start_row, max_row=data_start_row + len(df_comparison) - 1) 
                    
                    chart.add_data(data_ref, titles_from_data=False) 
                    chart.set_categories(cats_ref)
                
                elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
                    chart = LineChart() # Use LineChart for time series comparison
                    chart.title = "Hours Comparison by Project and Month"
                    chart.x_axis.title = "Month"
                    chart.y_axis.title = "Hours"

                    # For data series, exclude the 'Project Name' column
                    value_cols_indices = [df_comparison.columns.get_loc(col) + 1 for col in df_comparison.columns if col != 'Project Name']
                    
                    # Add a series for each project row
                    # The titles for series will come from the 'Project Name' column
                    for r_idx, project_name in enumerate(df_comparison['Project Name']):
                        series_ref = Reference(ws, min_col=min(value_cols_indices), 
                                                min_row=data_start_row + r_idx, 
                                                max_col=max(value_cols_indices), 
                                                max_row=data_start_row + r_idx)
                        # Reference for series title (project name)
                        title_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Project Name') + 1, 
                                            min_row=data_start_row + r_idx, 
                                            max_row=data_start_row + r_idx)
                        chart.add_data(series_ref, titles_from_data=True)
                        chart.series[r_idx].title = title_ref # Assign title to series


                    # Categories (months) are the headers, from second column to last
                    cats_ref = Reference(ws, min_col=2, min_row=1, max_col=len(df_comparison.columns)) 
                    chart.set_categories(cats_ref)

                elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
                    chart = LineChart()
                    chart.title = f"Hours for Project {comparison_config['selected_projects'][0]} Over Time"
                    chart.y_axis.title = "Hours"

                    # Determine x-axis label and categories based on the data
                    if 'Year' in df_comparison.columns and 'MonthName' in df_comparison.columns:
                        chart.x_axis.title = "Year-Month"
                        # Categories from Year and MonthName columns
                        cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Year') + 1, 
                                            min_row=data_start_row, 
                                            max_col=df_comparison.columns.get_loc('MonthName') + 1,
                                            max_row=data_start_row + len(df_comparison) - 1)
                    elif 'Year' in df_comparison.columns:
                        chart.x_axis.title = "Year"
                        cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Year') + 1, min_row=data_start_row, max_row=data_start_row + len(df_comparison) - 1)
                    elif 'MonthName' in df_comparison.columns:
                        chart.x_axis.title = "Month"
                        cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('MonthName') + 1, min_row=data_start_row, max_row=data_start_row + len(df_comparison) - 1)
                    else:
                        raise ValueError("No valid time dimension found for chart categories.")

                    # Add data series, handling pivoted and non-pivoted data
                    value_cols = [col for col in df_comparison.columns if col not in ['Year', 'MonthName', 'Total Hours']]
                    if 'Total Hours' in df_comparison.columns: # For non-pivoted data (single series)
                        series_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Total Hours') + 1, min_row=data_start_row, max_row=data_start_row + len(df_comparison) - 1)
                        chart.add_data(series_ref, titles_from_data=False) 
                    else: # For pivoted data (multiple series, e.g., for different months in different years)
                        # Each month column is a series
                        for col_name in value_cols:
                            col_idx = df_comparison.columns.get_loc(col_name) + 1
                            series_ref = Reference(ws, min_col=col_idx, min_row=data_start_row, max_row=data_start_row + len(df_comparison) - 1)
                            chart.add_data(series_ref, titles_from_data=True) 

                    chart.set_categories(cats_ref)

                if chart: 
                    chart_placement_row = info_row + 2
                    ws.add_chart(chart, f"A{chart_placement_row}")

            wb.save(output_file)
            return True
    except Exception as e:
        print(f"Error exporting comparison report to Excel: {e}")
        return False

def export_comparison_pdf_report(df_comparison, comparison_config, path_dict, comparison_mode):
    pdf_file = path_dict['comparison_pdf_report']
    tmp_dir = tempfile.mkdtemp()
    charts_for_pdf = []

    def create_pdf_from_charts(charts_data, output_path, title, config_info, logo_path="triac_logo.png"):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font('helvetica', 'B', 16) 

        pdf.add_page()
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=10, y=10, w=30)
        pdf.ln(40)
        pdf.cell(0, 10, title, ln=True, align='C')
        pdf.set_font("helvetica", '', 12) 
        pdf.ln(5)
        pdf.cell(0, 10, f"Generated on: {datetime.datetime.today().strftime('%Y-%m-%d')}", ln=True, align='C')
        pdf.ln(10)
        pdf.set_font("helvetica", '', 11) 
        for key, value in config_info.items():
            pdf.cell(0, 7, f"{key}: {value}", ln=True, align='C')

        for img_path, chart_title, page_project_name in charts_data:
            if img_path and os.path.exists(img_path):
                pdf.add_page()
                if os.path.exists(logo_path):
                    pdf.image(logo_path, x=10, y=8, w=25)
                pdf.set_font("helvetica", 'B', 11) 
                pdf.set_y(35)
                if page_project_name:
                    pdf.cell(0, 10, f"Project: {page_project_name}", ln=True, align='C')
                pdf.cell(0, 10, chart_title, ln=True, align='C')
                pdf.image(img_path, x=10, y=45, w=190)

        pdf.output(output_path, "F")
        print(f"DEBUG: PDF report generated at {output_path}")

    def create_comparison_chart(df, mode, title, x_label, y_label, img_path):
        fig, ax = plt.subplots(figsize=(12, 7))  
        
        df_plot = df.copy()  
        
        if df_plot.empty:
            print(f"DEBUG: df_plot is empty for mode '{mode}'. Skipping chart creation.")
            plt.close(fig) 
            return None 

        ax.set_ylim(bottom=0)

        # Matplotlib configuration to avoid Vietnamese font issues
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'Liberation Sans']
        plt.rcParams['axes.unicode_minus'] = False

        if mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
            ax.bar(df_plot['Project name'], df_plot['Total Hours'], color='skyblue')
            ax.set_xticks(df_plot['Project name'])
            ax.tick_params(axis='x', rotation=45, ha='right')
            ax.tick_params(axis='y', labelsize=8)  
        elif mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
            # Line chart for multiple projects over months
            if 'Project Name' in df_plot.columns and 'Total' in df_plot['Project Name'].values:
                df_plot = df_plot[df_plot['Project Name'] != 'Total']
            
            month_columns = [col for col in df_plot.columns if col not in ['Project Name']]
            df_plot[month_columns] = df_plot[month_columns].apply(pd.to_numeric, errors='coerce').fillna(0)

            df_plot.set_index('Project Name', inplace=True) 
            
            # Transpose to plot months on X-axis and projects as lines
            df_plot.T.plot(kind='line', ax=ax, colormap='viridis', marker='o')  
            
            ax.set_xticks(range(len(df_plot.columns)))  
            ax.set_xticklabels(df_plot.columns, rotation=45, ha='right')
            
            ax.legend(title="Project", bbox_to_anchor=(1.05, 1), loc='upper left') 
            ax.tick_params(axis='y', labelsize=8)
            x_label = "Month" 
            
        elif mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
            if 'Total' in df_plot.index: # If 'Total' row was added, remove it for plotting
                df_plot = df_plot.drop('Total', errors='ignore')

            data_cols = [col for col in df_plot.columns if col not in ['Year', 'MonthName', 'Total Hours']]
            
            if 'Year' in df_plot.columns and 'MonthName' in df_plot.columns:
                month_to_num = {name: i for i, name in enumerate(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], 1)}
                df_plot['MonthNum'] = df_plot['MonthName'].map(month_to_num)
                df_plot['YearMonth'] = df_plot['Year'].astype(str) + '-' + df_plot['MonthNum'].astype(str).str.zfill(2)
                df_plot = df_plot.sort_values(by=['Year', 'MonthNum'])
                df_plot.set_index('YearMonth', inplace=True)
                x_label = "Time (Year-Month)"  
            elif 'Year' in df_plot.columns:
                df_plot = df_plot.sort_values(by='Year')
                df_plot.set_index('Year', inplace=True)
                x_label = "Year"  
            elif 'MonthName' in df_plot.columns:
                month_order_list = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
                df_plot['MonthName_ordered'] = pd.Categorical(df_plot['MonthName'], categories=month_order_list, ordered=True)
                df_plot = df_plot.sort_values(by='MonthName_ordered')
                df_plot.set_index('MonthName', inplace=True)
                x_label = "Month"  

            if 'Total Hours' in df_plot.columns: # If non-pivoted data
                ax.plot(df_plot.index, df_plot['Total Hours'], marker='o', color='salmon')  
            elif data_cols: # If pivoted data with multiple month columns
                df_plot[data_cols] = df_plot[data_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
                df_plot[data_cols].plot(kind='line', ax=ax, colormap='plasma', marker='o') 
                ax.legend(title="Month", bbox_to_anchor=(1.05, 1), loc='upper left')

            plt.xticks(rotation=45, ha='right')
            ax.tick_params(axis='y', labelsize=8)

        ax.set_title(title)
        ax.set_xlabel(x_label)  
        ax.set_ylabel(y_label)
        plt.tight_layout()
        fig.savefig(img_path, dpi=200)
        plt.close(fig) # Close the figure to free up memory
        return img_path

    try:
        config_info = {
            "Comparison Mode": comparison_mode,
            "Years": ', '.join(map(str, comparison_config.get('years', []))) or "N/A",
            "Months": ', '.join(comparison_config.get('months', [])) or "All",
            "Projects": ', '.join(comparison_config.get('selected_projects', [])) or "No projects selected"
        }

        # Determine chart title, x_label, y_label based on comparison_mode
        chart_title = ""
        x_label = ""
        y_label = "Hours"
        chart_data_to_plot = df_comparison.copy() # Use a copy for plotting

        if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
            chart_title = f"Hours Comparison Between Projects in {comparison_config['months'][0]}, Year {comparison_config['years'][0]}"
            x_label = "Project"
            page_project_name = None # Not applicable for this mode
        elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
            chart_title = f"Hours Comparison Between Projects in Year {comparison_config['years'][0]}"
            x_label = "Month"
            page_project_name = None # Not applicable for this mode
            # Ensure index is reset for plotting if 'Total' row was added for Excel export
            if 'Total' in chart_data_to_plot.index:
                chart_data_to_plot = chart_data_to_plot.drop('Total', errors='ignore')

        elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
            chart_title = f"Hours for Project {comparison_config['selected_projects'][0]} Over Time"
            page_project_name = comparison_config['selected_projects'][0] # Specific project name
            
            # Ensure index is reset for plotting if 'Total' row was added for Excel export
            if 'Total' in chart_data_to_plot.index:
                chart_data_to_plot = chart_data_to_plot.drop('Total', errors='ignore')

            if 'Year' in chart_data_to_plot.columns and 'MonthName' in chart_data_to_plot.columns:
                x_label = "Year-Month"
            elif 'Year' in chart_data_to_plot.columns:
                x_label = "Year"
            elif 'MonthName' in chart_data_to_plot.columns:
                x_label = "Month"
            else:
                x_label = "Time" # Fallback

        # Generate the chart and add to list
        chart_img_path = os.path.join(tmp_dir, "comparison_chart.png")
        generated_path = create_comparison_chart(chart_data_to_plot, comparison_mode, chart_title, x_label, y_label, chart_img_path)
        if generated_path:
            charts_for_pdf.append((generated_path, chart_title, page_project_name))

        create_pdf_from_charts(charts_for_pdf, pdf_file, "TRIAC TIME REPORT - COMPARISON", config_info)
        return True

    except Exception as e:
        print(f"Error exporting comparison PDF report: {e}")
        return False
    finally:
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)
