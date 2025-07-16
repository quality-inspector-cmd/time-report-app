import pandas as pd
import datetime
import os
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference, LineChart # Import LineChart explicitly
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF
from matplotlib import pyplot as plt
import matplotlib.font_manager as fm
import tempfile
import re
import shutil

# --- Cấu hình font cho Matplotlib (thêm vào đầu file) ---
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Liberation Sans', 'Helvetica', 'Arial']
plt.rcParams['axes.unicode_minus'] = False 

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
        year_mode_df = pd.read_excel(path_dict['template_file'], sheet_name='Config_Year_Mode', engine='openpyxl')
        project_filter_df = pd.read_excel(path_dict['template_file'], sheet_name='Config_Project_Filter', engine='openpyxl')

        mode = str(year_mode_df.loc[year_mode_df['Key'].str.lower() == 'mode', 'Value'].values[0]).strip().lower()
        year_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'year', 'Value']
        year = int(year_row.values[0]) if not year_row.empty and pd.notna(year_row.values[0]) else None
        months_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'months', 'Value']
        months = [m.strip().capitalize() for m in str(months_row.values[0]).split(',')] if not months_row.empty and pd.notna(months_row.values[0]) else []
        
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
        return {'mode': 'year', 'year': datetime.datetime.now().year, 'months': [], 'project_filter_df': pd.DataFrame(columns=['Project Name', 'Include'])}
    except Exception as e:
        print(f"Error reading config: {e}")
        return {'mode': 'year', 'year': datetime.datetime.now().year, 'months': [], 'project_filter_df': pd.DataFrame(columns=['Project Name', 'Include'])}


def load_raw_data(path_dict):
    df = pd.read_excel(path_dict['template_file'], sheet_name='Raw Data', engine='openpyxl')
    df.columns = df.columns.str.strip() 
    df.rename(columns={'Hou': 'Hours', 'Team member': 'Employee', 'Project Name': 'Project name'}, inplace=True)
    
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df['Year'] = df['Date'].dt.year
    df['MonthName'] = df['Date'].dt.month_name()
    df['Week'] = df['Date'].dt.isocalendar().week
    return df

def apply_filters(df, config):
    df_filtered = df.copy()
    if 'years' in config and config['years']:
        df_filtered = df_filtered[df_filtered['Year'].isin(config['years'])]
    elif 'year' in config and config['year']:
        df_filtered = df_filtered[df_filtered['Year'] == config['year']]

    if config['months']:
        df_filtered = df_filtered[df_filtered['MonthName'].isin(config['months'])]

    if not config['project_filter_df'].empty:
        selected_project_names = config['project_filter_df']['Project Name'].tolist()
        df_filtered = df_filtered[df_filtered['Project name'].isin(selected_project_names)]
    else:
        df_filtered = pd.DataFrame(columns=df.columns) 

    return df_filtered

def export_report(df, config, path_dict):
    mode = config.get('mode', 'year')
    
    required_cols = ['Year', 'Project name', 'Hours']
    if mode == 'month':
        required_cols.append('MonthName')
    
    for col in required_cols:
        if col not in df.columns:
            raise KeyError(f"Missing required column for grouping: {col}")

    if df.empty:
        print(f"DEBUG: df is empty in export_report for mode {mode}. Skipping export.")
        return

    if mode == 'year':
        summary = df.groupby(['Year', 'Project name'])['Hours'].sum().reset_index()
    elif mode == 'month':
        summary = df.groupby(['Year', 'MonthName', 'Project name'])['Hours'].sum().reset_index()
    else: # week mode
        if 'Week' not in df.columns:
            raise KeyError("Missing 'Week' column for 'week' mode report.")
        summary = df.groupby(['Year', 'Week', 'Project name'])['Hours'].sum().reset_index()

    with pd.ExcelWriter(path_dict['output_file'], engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='Summary', index=False)

    wb = load_workbook(path_dict['output_file'])

    ws = wb['Summary']
    max_row = ws.max_row
    
    if mode == 'year':
        data_col = summary.columns.get_loc('Hours') + 1
        cats_col = summary.columns.get_loc('Project name') + 1
    elif mode == 'month':
        data_col = summary.columns.get_loc('Hours') + 1
        cats_col = summary.columns.get_loc('Project name') + 1
    else: # week
        data_col = summary.columns.get_loc('Hours') + 1
        cats_col = summary.columns.get_loc('Project name') + 1


    data_ref = Reference(ws, min_col=data_col, min_row=1, max_row=max_row)
    cats_ref = Reference(ws, min_col=cats_col, min_row=2, max_row=max_row)

    chart = BarChart()
    chart.title = f"Total Hours by Project ({mode})"
    chart.x_axis.title = "Project"
    chart.y_axis.title = "Hours"
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws.add_chart(chart, "F2")

    for project in df['Project name'].unique():
        df_proj = df[df['Project name'] == project]
        ws_proj = wb.create_sheet(title=sanitize_filename(project)[:31])

        summary_task = df_proj.groupby('Task')['Hours'].sum().reset_index().sort_values('Hours', ascending=False)
        
        ws_proj.append(['Task', 'Hours'])
        for r_idx, row_data in enumerate(dataframe_to_rows(summary_task, index=False, header=False)):
            ws_proj.append(row_data)

        chart = BarChart()
        chart.title = f"{project} - Hours by Task"
        chart.x_axis.title = "Task"
        chart.y_axis.title = "Hours"
        task_len = len(summary_task)
        
        data_ref = Reference(ws_proj, min_col=2, min_row=1, max_row=task_len + 1)
        cats_ref = Reference(ws_proj, min_col=1, min_row=2, max_row=task_len + 1)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        ws_proj.add_chart(chart, f"E1")

        start_row = task_len + 5
        for r_idx, r in enumerate(dataframe_to_rows(df_proj, index=False, header=True)):
            for c_idx, cell_val in enumerate(r):
                ws_proj.cell(row=start_row + r_idx, column=c_idx + 1, value=cell_val)
        
    ws_config = wb.create_sheet("Config_Info")
    ws_config['A1'], ws_config['B1'] = "Mode", config.get('mode', 'N/A')
    ws_config['A2'], ws_config['B2'] = "Years", ', '.join(map(str, config.get('years', []))) if config.get('years') else str(config.get('year', 'N/A'))
    ws_config['A3'], ws_config['B3'] = "Months", ', '.join(config.get('months', [])) if config.get('months') else "All"
    
    if 'project_filter_df' in config and not config['project_filter_df'].empty:
        ws_config['A4'], ws_config['B4'] = "Projects", ', '.join(config['project_filter_df']['Project Name'])
    else:
        ws_config['A4'], ws_config['B4'] = "Projects", "No projects selected or found"

    if 'Raw_Data' in wb.sheetnames:
        del wb['Raw_Data']

    wb.save(path_dict['output_file'])

def export_pdf_report(df, config, path_dict):
    today_str = datetime.datetime.today().strftime("%Y-%m-%d")

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
        projects = df['Project name'].unique() 

        config_info = {
            "Mode": config.get('mode', 'N/A').capitalize(),
            "Years": ', '.join(map(str, config.get('years', []))) or str(config.get('year', 'N/A')),
            "Months": ', '.join(config.get('months', [])) or "All",
            "Projects Included": ', '.join(config['project_filter_df']['Project Name']) if 'project_filter_df' in config and not config['project_filter_df'].empty else "No projects selected or found"
        }

        for project in projects:
            safe_project = sanitize_filename(project)
            df_proj = df[df['Project name'] == project]

            fig, ax = plt.subplots(figsize=(10, 5)) 
            df_proj.groupby('Workcentre')['Hours'].sum().sort_values().plot(kind='barh', color='skyblue', ax=ax)
            ax.set_title(f"{project} - Hours by Workcentre", fontsize=9) 
            ax.tick_params(axis='y', labelsize=8) 
            wc_img_path = os.path.join(tmp_dir, f"{safe_project}_wc.png")
            plt.tight_layout()
            fig.savefig(wc_img_path, dpi=150)
            plt.close(fig)
            charts_for_pdf.append((wc_img_path, f"{project} - Hours by Workcentre", project))

            if 'Task' in df_proj.columns and not df_proj['Task'].empty:
                fig, ax = plt.subplots(figsize=(10, 6)) 
                df_proj.groupby('Task')['Hours'].sum().sort_values().plot(kind='barh', color='lightgreen', ax=ax)
                ax.set_title(f"{project} - Hours by Task", fontsize=9) 
                ax.tick_params(axis='y', labelsize=8) 
                task_img_path = os.path.join(tmp_dir, f"{safe_project}_task.png")
                plt.tight_layout()
                fig.savefig(task_img_path, dpi=150)
                plt.close(fig)
                charts_for_pdf.append((task_img_path, f"{project} - Hours by Task", project))

        create_pdf_from_charts(charts_for_pdf, path_dict['pdf_report'], "TRIAC TIME REPORT - STANDARD", config_info)

    finally:
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)


# =========================================================================
# CÁC HÀM MỚI CHO BÁO CÁO SO SÁNH
# =========================================================================

def apply_comparison_filters(df_raw, comparison_config, comparison_mode):
    years = comparison_config['years']
    months = comparison_config['months']
    selected_projects = comparison_config['selected_projects']

    print(f"\nDEBUG - apply_comparison_filters START:")
    print(f"  comparison_mode: {comparison_mode}")
    print(f"  years: {years}")
    print(f"  months: {months}")
    print(f"  selected_projects: {selected_projects}")

    df_filtered = df_raw.copy()

    if years:
        df_filtered = df_filtered[df_filtered['Year'].isin(years)]
    
    if months:
        df_filtered = df_filtered[df_filtered['MonthName'].isin(months)]
    
    if selected_projects:
        df_filtered = df_filtered[df_filtered['Project name'].isin(selected_projects)]
    else: 
        print(f"DEBUG - apply_comparison_filters: No selected projects. Returning empty DataFrame.")
        return pd.DataFrame(), "Vui lòng chọn ít nhất một dự án để so sánh."

    print(f"DEBUG - apply_comparison_filters: df_filtered head after initial filters:\n{df_filtered.head()}")
    print(f"DEBUG - apply_comparison_filters: df_filtered shape after initial filters: {df_filtered.shape}")

    if df_filtered.empty:
        print(f"DEBUG - apply_comparison_filters: df_filtered is empty after filtering. Returning empty DataFrame.")
        return pd.DataFrame(), f"Không có dữ liệu cho chế độ so sánh: {comparison_mode} với các lựa chọn hiện tại."

    df_comparison = pd.DataFrame()
    title = ""
    error_message = ""

    # Define standard month order for consistent sorting
    month_order_list = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

    if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
        if len(years) != 1 or len(months) != 1 or len(selected_projects) < 2:
            error_message = "Cần chọn MỘT năm, MỘT tháng và ít nhất HAI dự án cho chế độ này."
            print(f"DEBUG - apply_comparison_filters: Condition failed for 'Compare Projects in a Month'. {error_message}")
            return pd.DataFrame(), error_message
        
        df_comparison = df_filtered.groupby('Project name')['Hours'].sum().reset_index()
        df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
        title = f"So sánh giờ giữa các dự án trong {months[0]}, năm {years[0]}"
        print(f"DEBUG - apply_comparison_filters (Month Comparison): df_comparison head:\n{df_comparison.head()}")
        return df_comparison, title

    elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        if len(years) != 1 or len(selected_projects) < 2:
            error_message = "Cần chọn MỘT năm và ít nhất HAI dự án cho chế độ này."
            print(f"DEBUG - apply_comparison_filters: Condition failed for 'Compare Projects in a Year'. {error_message}")
            return pd.DataFrame(), error_message
        
        df_comparison = df_filtered.groupby(['Project name', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
        
        existing_months = [m for m in month_order_list if m in df_comparison.columns]
        df_comparison = df_comparison[existing_months]

        df_comparison = df_comparison.reset_index().rename(columns={'Project name': 'Project Name'})
        
        title = f"So sánh giờ giữa các dự án trong năm {years[0]} (theo tháng)"
        print(f"DEBUG - apply_comparison_filters (Year Comparison): df_comparison head:\n{df_comparison.head()}")
        return df_comparison, title

    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
        print(f"DEBUG - apply_comparison_filters: Inside 'Compare One Project Over Time' mode.")
        if len(selected_projects) != 1:
            error_message = "Chế độ 'So Sánh Một Dự Án Qua Các Tháng/Năm' yêu cầu chọn CHỈ MỘT dự án."
            print(f"DEBUG - apply_comparison_filters: More than one project selected. {error_message}")
            return pd.DataFrame(), error_message
        
        # Determine the time granularity for comparison
        is_month_comparison = bool(months) and len(years) <= 1 # If months are selected, and at most one year is specified
        is_year_comparison = bool(years) and not months # If years are selected but no specific months

        if is_month_comparison:
            print(f"DEBUG: Processing 1 project, specific months.")
            
            # Use all 12 months for the full range if no specific months are selected in config,
            # otherwise use only the selected months.
            target_months = months if months else month_order_list
            
            # Create a DataFrame with all target months, then merge
            full_time_range_df = pd.DataFrame({'MonthName': target_months})
            full_time_range_df['MonthNum'] = full_time_range_df['MonthName'].apply(lambda x: month_order_list.index(x) + 1)
            full_time_range_df = full_time_range_df.sort_values(by='MonthNum').drop(columns='MonthNum')

            df_grouped = df_filtered.groupby('MonthName')['Hours'].sum().reset_index()
            
            df_comparison = full_time_range_df.merge(
                df_grouped, on='MonthName', how='left'
            ).fillna(0)
            
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            
            # If a single year is selected, append it to the title
            if len(years) == 1:
                title = f"Giờ của dự án {selected_projects[0]} trong năm {years[0]} theo tháng"
                # Add year column for consistent structure, useful for debugging
                df_comparison['Year'] = years[0] 
            else:
                title = f"Giờ của dự án {selected_projects[0]} theo tháng (Tổng hợp các năm)"
            
            # Ensure correct categorical order for plotting
            df_comparison['MonthName'] = pd.Categorical(df_comparison['MonthName'], categories=target_months, ordered=True)
            df_comparison = df_comparison.sort_values(by='MonthName')
            
        elif is_year_comparison:
            print(f"DEBUG: Processing 1 project, multiple years.")
            
            target_years = sorted(years)
            full_time_range_df = pd.DataFrame({'Year': target_years})

            df_grouped = df_filtered.groupby('Year')['Hours'].sum().reset_index()

            df_comparison = full_time_range_df.merge(
                df_grouped, on='Year', how='left'
            ).fillna(0)
            
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            title = f"Giờ của dự án {selected_projects[0]} qua các năm"
            
        else: 
            error_message = "Cần chọn ít nhất HAI tháng HOẶC HAI năm để so sánh một dự án qua thời gian."
            print(f"DEBUG - apply_comparison_filters: No valid comparison combination for 'Compare One Project Over Time'. {error_message}")
            return pd.DataFrame(), error_message
        
        print(f"DEBUG - apply_comparison_filters (Compare One Project Over Time): Final df_comparison head:\n{df_comparison.head()}")
        print(f"DEBUG - apply_comparison_filters (Compare One Project Over Time): Final df_comparison columns: {df_comparison.columns.tolist()}")
        print(f"DEBUG - apply_comparison_filters (Compare One Project Over Time): df_comparison dtypes:\n{df_comparison.dtypes}")
        return df_comparison, title
    
    print(f"DEBUG - apply_comparison_filters END: No matching comparison mode.")
    return pd.DataFrame(), "Chế độ so sánh không hợp lệ."

def export_comparison_report(df_comparison, comparison_config, path_dict, comparison_mode):
    output_file = path_dict['comparison_output_file']
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_comparison.to_excel(writer, sheet_name='Comparison Report', index=False) 

        wb = writer.book
        ws = writer.sheets['Comparison Report']

        ws.merge_cells('A1:D1')
        ws['A1'].value = f"
