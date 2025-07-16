import pandas as pd
import datetime
import os
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
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

            # Bắt đầu đoạn code cho biểu đồ Workcentre
            fig, ax = plt.subplots(figsize=(10, 5)) 
            df_proj.groupby('Workcentre')['Hours'].sum().sort_values().plot(kind='barh', color='skyblue', ax=ax)
            ax.set_title(f"{project} - Hours by Workcentre", fontsize=9) 
            ax.tick_params(axis='y', labelsize=8) 
            wc_img_path = os.path.join(tmp_dir, f"{safe_project}_wc.png")
            plt.tight_layout()
            fig.savefig(wc_img_path, dpi=150)
            plt.close(fig)
            charts_for_pdf.append((wc_img_path, f"{project} - Hours by Workcentre", project))

            # Bắt đầu đoạn code cho biểu đồ Task
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

    if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
        if len(years) != 1 or len(months) != 1 or len(selected_projects) < 2:
            print(f"DEBUG - apply_comparison_filters: Condition failed for 'Compare Projects in a Month'. Years: {len(years)}, Months: {len(months)}, Projects: {len(selected_projects)}")
            return pd.DataFrame(), "Cần chọn MỘT năm, MỘT tháng và ít nhất HAI dự án cho chế độ này."
        
        df_comparison = df_filtered.groupby('Project name')['Hours'].sum().reset_index()
        df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
        title = f"So sánh giờ giữa các dự án trong {months[0]}, năm {years[0]}"
        print(f"DEBUG - apply_comparison_filters (Month Comparison): df_comparison head:\n{df_comparison.head()}")
        return df_comparison, title

    elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        if len(years) != 1 or len(selected_projects) < 2:
            print(f"DEBUG - apply_comparison_filters: Condition failed for 'Compare Projects in a Year'. Years: {len(years)}, Projects: {len(selected_projects)}")
            return pd.DataFrame(), "Cần chọn MỘT năm và ít nhất HAI dự án cho chế độ này."
        
        df_comparison = df_filtered.groupby(['Project name', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
        
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        existing_months = [m for m in month_order if m in df_comparison.columns]
        df_comparison = df_comparison[existing_months]

        df_comparison = df_comparison.reset_index().rename(columns={'index': 'Project Name'})
        
        title = f"So sánh giờ giữa các dự án trong năm {years[0]} (theo tháng)"
        print(f"DEBUG - apply_comparison_filters (Year Comparison): df_comparison head:\n{df_comparison.head()}")
        return df_comparison, title

    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
        print(f"DEBUG - apply_comparison_filters: Inside 'Compare One Project Over Time' mode.")
        if len(selected_projects) != 1:
            print(f"DEBUG - apply_comparison_filters: More than one project selected. Returning error.")
            return pd.DataFrame(), "Chế độ 'So Sánh Một Dự Án Qua Các Tháng/Năm' yêu cầu chọn CHỈ MỘT dự án."
        
        if not years and not months:
            print(f"DEBUG - apply_comparison_filters: No years or months selected. Returning error.")
            return pd.DataFrame(), "Cần chọn ít nhất một năm HOẶC một tháng để so sánh."

        df_comparison = pd.DataFrame()
        title = ""

        # Trường hợp 1: 1 dự án, 1 năm, nhiều tháng (chỉ filter tháng trong năm đó)
        if len(years) == 1 and len(months) >= 2: 
            print(f"DEBUG - apply_comparison_filters: Processing 1 project, 1 year ({years[0]}), multiple months ({months}).")
            df_comparison = df_filtered.groupby(['Year', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
            month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
            existing_months = [m for m in month_order if m in df_comparison.columns]
            df_comparison = df_comparison[existing_months] 
            df_comparison = df_comparison.reset_index() # Giữ Year là cột
            
        # Trường hợp 2: 1 dự án, nhiều năm, không chọn tháng cụ thể (tổng hợp theo năm)
        elif len(years) >= 2 and not months: 
            print(f"DEBUG - apply_comparison_filters: Processing 1 project, multiple years ({years}), no specific months.")
            df_comparison = df_filtered.groupby(['Year'])['Hours'].sum().reset_index()
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            
        # Trường hợp 3: 1 dự án, nhiều tháng, không chọn năm cụ thể (tổng hợp theo tháng trên tất cả các năm có dữ liệu)
        elif len(months) >= 2 and not years: 
            print(f"DEBUG - apply_comparison_filters: Processing 1 project, multiple months ({months}), no specific years.")
            df_comparison = df_filtered.groupby(['MonthName'])['Hours'].sum().reset_index()
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            month_order_df = pd.DataFrame({'MonthName': ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']})
            df_comparison = pd.merge(month_order_df, df_comparison, on='MonthName', how='left').fillna(0)
            
        # Trường hợp 4: 1 dự án, nhiều năm, và có chọn tháng cụ thể (VD: tháng 1 của 2023, tháng 1 của 2024, ...)
        elif len(years) >= 2 and len(months) >= 1: 
            print(f"DEBUG - apply_comparison_filters: Processing 1 project, multiple years ({years}), specific months ({months}).")
            df_comparison = df_filtered.groupby(['Year', 'MonthName'])['Hours'].sum().reset_index()
            month_to_num = {name: i for i, name in enumerate(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], 1)}
            df_comparison['MonthNum'] = df_comparison['MonthName'].map(month_to_num)
            df_comparison['YearMonth'] = df_comparison['Year'].astype(str) + '-' + df_comparison['MonthNum'].astype(str).str.zfill(2)
            df_comparison = df_comparison.sort_values(by=['Year', 'MonthNum'])
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            df_comparison = df_comparison[['YearMonth', 'Total Hours']] 
        
        else: # Các trường hợp không đủ dữ liệu hoặc không được hỗ trợ rõ ràng
            print(f"DEBUG - apply_comparison_filters: No valid comparison combination for 'Compare One Project Over Time'.")
            return pd.DataFrame(), "Cần chọn ít nhất HAI tháng HOẶC HAI năm để so sánh một dự án qua thời gian."
        
        print(f"DEBUG - apply_comparison_filters (Compare One Project Over Time): Final df_comparison head:\n{df_comparison.head()}")
        print(f"DEBUG - apply_comparison_filters (Compare One Project Over Time): Final df_comparison columns: {df_comparison.columns.tolist()}")

        title = f"So sánh giờ của dự án {selected_projects[0]} qua thời gian"
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
        ws['A1'].value = f"BÁO CÁO SO SÁNH: {comparison_mode}"
        ws['A1'].font = ws['A1'].font.copy(bold=True, size=14)

        info_row = 2
        ws.cell(row=info_row, column=1, value="Năm:").font = ws.cell(row=info_row, column=1).font.copy(bold=True)
        ws.cell(row=info_row, column=2, value=', '.join(map(str, comparison_config['years'])))
        info_row += 1
        ws.cell(row=info_row, column=1, value="Tháng:").font = ws.cell(row=info_row, column=1).font.copy(bold=True)
        ws.cell(row=info_row, column=2, value=', '.join(comparison_config['months']))
        info_row += 1
        ws.cell(row=info_row, column=1, value="Dự án:").font = ws.cell(row=info_row, column=1).font.copy(bold=True)
        ws.cell(row=info_row, column=2, value=', '.join(comparison_config['selected_projects']))

        if not df_comparison.empty:
            chart = BarChart()
            chart.y_axis.title = "Giờ"
            
            data_start_row = info_row + 2
            
            if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
                chart.title = "So sánh giờ theo Dự án"
                chart.x_axis.title = "Dự án"
                
                data_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Total Hours') + 1, min_row=data_start_row, max_row=data_start_row + len(df_comparison) -1)
                cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Project name') + 1, min_row=data_start_row + 1, max_row=data_start_row + len(df_comparison))
                
                chart.add_data(data_ref, titles_from_data=False) 
                chart.set_categories(cats_ref)
            
            elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
                chart.title = "So sánh giờ theo Dự án và Tháng"
                chart.x_axis.title = "Dự án"
                
                value_cols = [col for col in df_comparison.columns if col not in ['Project Name']]
                
                for idx, col_name in enumerate(value_cols):
                    series = Reference(ws, min_col=df_comparison.columns.get_loc(col_name) + 1, min_row=data_start_row, max_row=data_start_row + len(df_comparison)-1)
                    chart.add_data(series, titles_from_data=True)

                cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Project Name') + 1, min_row=data_start_row + 1, max_row=data_start_row + len(df_comparison))
                chart.set_categories(cats_ref)

            elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
                chart.title = "So sánh giờ của một dự án qua các tháng/năm"
                chart.x_axis.title = "Thời gian (Năm)"

                value_cols = [col for col in df_comparison.columns if col not in ['Year', 'MonthName', 'YearMonth']] 
                
                for idx, col_name in enumerate(value_cols):
                    series = Reference(ws, min_col=df_comparison.columns.get_loc(col_name) + 1, min_row=data_start_row, max_row=data_start_row + len(df_comparison)-1)
                    chart.add_data(series, titles_from_data=True)
                
                if 'Year' in df_comparison.columns:
                    cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Year') + 1, min_row=data_start_row + 1, max_row=data_start_row + len(df_comparison))
                elif 'MonthName' in df_comparison.columns:
                     cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('MonthName') + 1, min_row=data_start_row + 1, max_row=data_start_row + len(df_comparison))
                elif 'YearMonth' in df_comparison.columns:
                    cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('YearMonth') + 1, min_row=data_start_row + 1, max_row=data_start_row + len(df_comparison))
                else: 
                     cats_ref = Reference(ws, min_col=1, min_row=data_start_row + 1, max_row=data_start_row + len(df_comparison))

                chart.set_categories(cats_ref)


            ws.add_chart(chart, f"A{info_row + len(df_comparison) + 5}")

        wb.save(output_file)
        return True

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
        
        print(f"\nDEBUG - create_comparison_chart START for mode: '{mode}'")
        print(f"  Initial df_plot head:\n{df_plot.head()}")
        print(f"  Initial df_plot columns: {df_plot.columns.tolist()}")
        print(f"  df_plot shape: {df_plot.shape}")

        if df_plot.empty:
            print(f"DEBUG: df_plot is empty for mode '{mode}'. Skipping chart creation.")
            plt.close(fig) 
            return None 

        ax.set_ylim(bottom=0)

        if mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
            print(f"DEBUG: Plotting Bar Chart for 'Compare Projects in a Month'.")
            ax.bar(df_plot['Project name'], df_plot['Total Hours'], color='skyblue')
            ax.set_xticks(df_plot['Project name'])
            ax.tick_params(axis='x', rotation=45, ha='right')
            ax.tick_params(axis='y', labelsize=8) 
        elif mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
            print(f"DEBUG: Plotting Line Chart for 'Compare Projects in a Year'.")
            if 'Project Name' in df_plot.columns and 'Total' in df_plot['Project Name'].values:
                df_plot = df_plot[df_plot['Project Name'] != 'Total']
            
            month_columns = [col for col in df_plot.columns if col not in ['Project Name']]
            df_plot[month_columns] = df_plot[month_columns].apply(pd.to_numeric, errors='coerce').fillna(0)

            df_plot.set_index('Project Name', inplace=True)
            
            df_plot.T.plot(kind='line', ax=ax, colormap='viridis', marker='o') 
            
            ax.set_xticks(range(len(df_plot.columns))) 
            ax.set_xticklabels(df_plot.columns, rotation=45, ha='right')
            
            ax.legend(title="Dự án", bbox_to_anchor=(1.05, 1), loc='upper left') 
            ax.tick_params(axis='y', labelsize=8)
            x_label = "Tháng" 
            
        elif mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
            print(f"DEBUG: Inside 'Compare One Project Over Time' plotting logic.")
            if 'Year' in df_plot.columns and 'Total' in df_plot['Year'].values:
                df_plot = df_plot[df_plot['Year'] != 'Total']
            
            # Trường hợp 1: 1 dự án, 1 năm, nhiều tháng (df_plot có cột 'Year' và các cột tháng)
            # Sau unstack từ apply_comparison_filters: df_plot có dạng: Year | Month1 | Month2 | ...
            is_single_year_multi_month = 'Year' in df_plot.columns and len(df_plot['Year'].unique()) == 1 and any(col in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'] for col in df_plot.columns)

            if is_single_year_multi_month:
                print(f"DEBUG: Plotting Case 1: 1 project, 1 year, multiple months.")
                # Bỏ cột 'Year' và chuyển vị để các tháng thành index
                df_plot = df_plot.drop(columns='Year', errors='ignore').T
                df_plot.columns = ['Total Hours'] # Đặt tên cột là 'Total Hours' cho rõ ràng
                
                # Sắp xếp các tháng theo thứ tự tự nhiên
                month_order_list = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
                df_plot['MonthName_ordered'] = pd.Categorical(df_plot.index, categories=month_order_list, ordered=True)
                df_plot = df_plot.sort_values(by='MonthName_ordered')
                df_plot.drop(columns='MonthName_ordered', inplace=True)

                print(f"DEBUG: df_plot after transpose and sort:\n{df_plot.head()}")
                print(f"DEBUG: df_plot index (x-axis labels): {df_plot.index.tolist()}")

                ax.plot(df_plot.index, df_plot['Total Hours'], marker='o', color='salmon')
                x_label = f"Tháng ({df.iloc[0]['Year']})" if not df.empty and 'Year' in df.columns else "Tháng"
                ax.set_xticks(df_plot.index)
            
            # Trường hợp 2: 1 dự án, nhiều năm (df_plot có cột 'Year' và 'Total Hours')
            elif 'Year' in df_plot.columns and 'Total Hours' in df_plot.columns and not any(col in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'] for col in df_plot.columns):
                print(f"DEBUG: Plotting Case 2: 1 project, multiple years.")
                df_plot = df_plot.sort_values(by='Year')
                ax.plot(df_plot['Year'], df_plot['Total Hours'], marker='o', color='salmon')
                x_label = "Năm"
                ax.set_xticks(df_plot['Year'])

            # Trường hợp 3: 1 dự án, nhiều tháng qua nhiều năm (df_plot có cột 'YearMonth' và 'Total Hours')
            elif 'YearMonth' in df_plot.columns and 'Total Hours' in df_plot.columns:
                print(f"DEBUG: Plotting Case 3: 1 project, multiple months across multiple years.")
                df_plot = df_plot.sort_values(by='YearMonth')
                ax.plot(df_plot['YearMonth'], df_plot['Total Hours'], marker='o', color='salmon')
                x_label = "Thời gian (Năm-Tháng)"
                # Điều chỉnh tần suất hiển thị nhãn trục X nếu có quá nhiều điểm
                if len(df_plot['YearMonth']) > 12: 
                    step = max(1, len(df_plot['YearMonth']) // 6) 
                    ax.set_xticks(df_plot['YearMonth'].iloc[::step])
                else:
                    ax.set_xticks(df_plot['YearMonth'])

            # Trường hợp 4: 1 dự án, chỉ có tháng được chọn (tổng hợp tháng 1 qua tất cả các năm)
            elif 'MonthName' in df_plot.columns and 'Total Hours' in df_plot.columns:
                print(f"DEBUG: Plotting Case 4: 1 project, specific months (aggregated across years).")
                month_order_list = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
                df_plot['MonthName_ordered'] = pd.Categorical(df_plot['MonthName'], categories=month_order_list, ordered=True)
                df_plot = df_plot.sort_values(by='MonthName_ordered')
                ax.plot(df_plot['MonthName'], df_plot['Total Hours'], marker='o', color='salmon')
                x_label = "Tháng"
                ax.set_xticks(df_plot['MonthName'])
            
            else:
                print(f"Warning: No suitable plotting data structure found for 'Compare One Project Over Time' mode. df_plot columns: {df_plot.columns.tolist()}")
                ax.text(0.5, 0.5, "Không có dữ liệu để vẽ biểu đồ line.", horizontalalignment='center', verticalalignment='center', transform=ax.transAxes, fontsize=12)


            plt.xticks(rotation=45, ha='right')
            ax.tick_params(axis='y', labelsize=8)

        ax.set_title(title)
        ax.set_xlabel(x_label) 
        ax.set_ylabel(y_label)
        plt.tight_layout()
        fig.savefig(img_path, dpi=200)
        plt.close(fig)
        print(f"DEBUG: Chart saved to {img_path}")
        return img_path

    try:
        chart_title = f"Báo cáo so sánh: {comparison_mode}"
        x_label = ""
        y_label = "Tổng số giờ"

        if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
            x_label = "Dự án"
        elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
            x_label = "Tháng" 
        elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
            pass # x_label sẽ được set trong create_comparison_chart dựa trên case cụ thể
        
        comp_chart_path = os.path.join(tmp_dir, "comparison_chart.png")
        print(f"DEBUG: Attempting to create comparison chart at {comp_chart_path}")
        created_chart_path = create_comparison_chart(df_comparison.copy(), comparison_mode, chart_title, x_label, y_label, comp_chart_path)
        if created_chart_path:
            charts_for_pdf.append((created_chart_path, chart_title, None))
            print(f"DEBUG: Chart successfully created and added to charts_for_pdf.")
        else:
            print(f"DEBUG: create_comparison_chart returned None, chart not created.")


        config_info = {
            "Che do so sanh": comparison_mode, 
            "Nam": ', '.join(map(str, comparison_config['years'])) if comparison_config['years'] else "All", 
            "Thang": ', '.join(comparison_config['months']) if comparison_config['months'] else "All", 
            "Du an": ', '.join(comparison_config['selected_projects']) if comparison_config['selected_projects'] else "No projects selected"
        }
        
        print(f"DEBUG: Generating PDF with config_info: {config_info}")
        create_pdf_from_charts(charts_for_pdf, pdf_file, "TRIAC TIME REPORT - COMPARISON", config_info)
        print(f"DEBUG: PDF generation complete.")
    finally:
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)
            print(f"DEBUG: Cleaned up temporary directory: {tmp_dir}")
    return True
