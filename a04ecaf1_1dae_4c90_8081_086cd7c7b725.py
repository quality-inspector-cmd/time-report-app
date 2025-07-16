import pandas as pd
import datetime
import os
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference, LineChart
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF
from matplotlib import pyplot as plt
import tempfile
import re
import shutil

# Hàm hỗ trợ làm sạch tên file/sheet
def sanitize_filename(name):
    # Ký tự không hợp lệ trong tên file/sheet của Excel
    invalid_chars = re.compile(r'[\\/*?[\]:;|=,<>]')
    s = invalid_chars.sub("_", str(name))
    # Loại bỏ các ký tự điều khiển ASCII và các ký tự không an toàn khác
    s = ''.join(c for c in s if c.isprintable())
    return s[:31] # Giới hạn 31 ký tự cho tên sheet trong Excel

def setup_paths():
    """Thiết lập các đường dẫn file đầu vào và đầu ra."""
    today = datetime.datetime.today().strftime('%Y%m%d')
    return {
        'template_file': "Time_report.xlsm",
        'output_file': f"Time_report_Standard_{today}.xlsx",
        'pdf_report': f"Time_report_Standard_{today}.pdf",
        'comparison_output_file': f"Time_report_Comparison_{today}.xlsx",
        'comparison_pdf_report': f"Time_report_Comparison_{today}.pdf",
        'logo_path': "triac_logo.png" # Thêm đường dẫn logo
    }

def read_configs(template_file):
    """Đọc cấu hình từ file template Excel."""
    try:
        year_mode_df = pd.read_excel(template_file, sheet_name='Config_Year_Mode', engine='openpyxl')
        project_filter_df = pd.read_excel(template_file, sheet_name='Config_Project_Filter', engine='openpyxl')

        # Xử lý mode, year, months an toàn hơn
        mode_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'mode', 'Value']
        mode = str(mode_row.values[0]).strip().lower() if not mode_row.empty and pd.notna(mode_row.values[0]) else 'year'

        year_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'year', 'Value']
        year = int(year_row.values[0]) if not year_row.empty and pd.notna(year_row.values[0]) and pd.api.types.is_number(year_row.values[0]) else datetime.datetime.now().year
        
        months_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'months', 'Value']
        months = [m.strip().capitalize() for m in str(months_row.values[0]).split(',')] if not months_row.empty and pd.notna(months_row.values[0]) else []
        
        if 'Include' in project_filter_df.columns:
            project_filter_df['Include'] = project_filter_df['Include'].astype(str).str.lower()

        return {
            'mode': mode,
            'year': year,
            'months': months,
            'project_filter_df': project_filter_df
        }
    except FileNotFoundError:
        print(f"Lỗi: Không tìm thấy file template tại {template_file}")
        return {'mode': 'year', 'year': datetime.datetime.now().year, 'months': [], 'project_filter_df': pd.DataFrame(columns=['Project Name', 'Include'])}
    except Exception as e:
        print(f"Lỗi khi đọc cấu hình: {e}")
        return {'mode': 'year', 'year': datetime.datetime.now().year, 'months': [], 'project_filter_df': pd.DataFrame(columns=['Project Name', 'Include'])}

def load_raw_data(template_file):
    """Tải dữ liệu thô từ file template Excel."""
    try:
        df = pd.read_excel(template_file, sheet_name='Raw Data', engine='openpyxl')
        df.columns = df.columns.str.strip()
        df.rename(columns={'Hou': 'Hours', 'Team member': 'Employee', 'Project Name': 'Project name'}, inplace=True)
        
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.dropna(subset=['Date']) # Loại bỏ hàng không có ngày hợp lệ
        
        df['Year'] = df['Date'].dt.year
        df['MonthName'] = df['Date'].dt.month_name()
        df['Week'] = df['Date'].dt.isocalendar().week.astype(int)
        
        # Đảm bảo cột 'Hours' là số
        df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce').fillna(0)
        
        return df
    except Exception as e:
        print(f"Lỗi khi tải dữ liệu thô: {e}")
        return pd.DataFrame()

def apply_filters(df, config):
    """Áp dụng các bộ lọc dữ liệu dựa trên cấu hình."""
    df_filtered = df.copy()

    if 'years' in config and config['years']: # Dành cho so sánh nhiều năm
        df_filtered = df_filtered[df_filtered['Year'].isin(config['years'])]
    elif 'year' in config and config['year']: # Dành cho báo cáo tiêu chuẩn một năm
        df_filtered = df_filtered[df_filtered['Year'] == config['year']]

    if config['months']:
        df_filtered = df_filtered[df_filtered['MonthName'].isin(config['months'])]

    if not config['project_filter_df'].empty:
        selected_project_names = config['project_filter_df']['Project Name'].tolist()
        df_filtered = df_filtered[df_filtered['Project name'].isin(selected_project_names)]
    else:
        return pd.DataFrame(columns=df.columns) 

    return df_filtered

def export_report(df, config, output_file_path):
    """Xuất báo cáo tiêu chuẩn ra file Excel."""
    mode = config.get('mode', 'year')
    
    groupby_cols = []
    if mode == 'year':
        groupby_cols = ['Year', 'Project name']
    elif mode == 'month':
        groupby_cols = ['Year', 'MonthName', 'Project name']
    else: # week mode
        groupby_cols = ['Year', 'Week', 'Project name']

    for col in groupby_cols + ['Hours']:
        if col not in df.columns:
            print(f"Lỗi: Cột '{col}' không tồn tại trong DataFrame. Không thể tạo báo cáo.")
            return False

    if df.empty:
        print("Cảnh báo: DataFrame đã lọc trống, không có báo cáo nào được tạo.")
        return False

    summary = df.groupby(groupby_cols)['Hours'].sum().reset_index()

    try:
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name='Summary', index=False)

        wb = load_workbook(output_file_path)
        ws = wb['Summary']
        
        if len(summary) > 0:
            data_col_idx = summary.columns.get_loc('Hours') + 1
            cats_col_idx = summary.columns.get_loc('Project name') + 1

            data_ref = Reference(ws, min_col=data_col_idx, min_row=2, max_row=ws.max_row)
            cats_ref = Reference(ws, min_col=cats_col_idx, min_row=2, max_row=ws.max_row)

            chart = BarChart()
            chart.title = f"Total Hours by Project ({mode})"
            chart.x_axis.title = "Project"
            chart.y_axis.title = "Hours"
            
            chart.add_data(data_ref, titles_from_data=False) 
            chart.set_categories(cats_ref)
            ws.add_chart(chart, "F2")

        for project in df['Project name'].unique():
            df_proj = df[df['Project name'] == project]
            sheet_title = sanitize_filename(project)
            
            if sheet_title in wb.sheetnames:
                ws_proj = wb[sheet_title]
            else:
                ws_proj = wb.create_sheet(title=sheet_title)

            summary_task = df_proj.groupby('Task')['Hours'].sum().reset_index().sort_values('Hours', ascending=False)
            
            if not summary_task.empty:
                ws_proj.append(['Task', 'Hours'])
                for row_data in dataframe_to_rows(summary_task, index=False, header=False):
                    ws_proj.append(row_data)

                chart_task = BarChart()
                chart_task.title = f"{project} - Hours by Task"
                chart_task.x_axis.title = "Task"
                chart_task.y_axis.title = "Hours"
                task_len = len(summary_task)
                
                data_ref_task = Reference(ws_proj, min_col=2, min_row=1, max_row=task_len + 1)
                cats_ref_task = Reference(ws_proj, min_col=1, min_row=2, max_row=task_len + 1)
                chart_task.add_data(data_ref_task, titles_from_data=True)
                chart_task.set_categories(cats_ref_task)
                ws_proj.add_chart(chart_task, f"E1")

            start_row_raw_data = ws_proj.max_row + 2 if ws_proj.max_row > 1 else 1
            if not summary_task.empty:
                start_row_raw_data += 15

            for r_idx, r in enumerate(dataframe_to_rows(df_proj, index=False, header=True)):
                for c_idx, cell_val in enumerate(r):
                    ws_proj.cell(row=start_row_raw_data + r_idx, column=c_idx + 1, value=cell_val)
        
        ws_config = wb.create_sheet("Config_Info")
        ws_config['A1'], ws_config['B1'] = "Mode", config.get('mode', 'N/A').capitalize()
        ws_config['A2'], ws_config['B2'] = "Year(s)", ', '.join(map(str, config.get('years', []))) if config.get('years') else str(config.get('year', 'N/A'))
        ws_config['A3'], ws_config['B3'] = "Months", ', '.join(config.get('months', [])) if config.get('months') else "All"
        
        if 'project_filter_df' in config and not config['project_filter_df'].empty:
            selected_projects_display = config['project_filter_df'][config['project_filter_df']['Include'].astype(str).str.lower() == 'yes']['Project Name'].tolist()
            ws_config['A4'], ws_config['B4'] = "Projects Included", ', '.join(selected_projects_display)
        else:
            ws_config['A4'], ws_config['B4'] = "Projects Included", "No projects selected or found"

        # Remove template sheets
        for sheet_name in ['Raw Data', 'Config_Year_Mode', 'Config_Project_Filter']:
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]

        wb.save(output_file_path)
        return True
    except Exception as e:
        print(f"Lỗi khi xuất báo cáo tiêu chuẩn: {e}")
        return False

def export_pdf_report(df, config, pdf_report_path, logo_path):
    """Xuất báo cáo PDF tiêu chuẩn với các biểu đồ."""
    today_str = datetime.datetime.today().strftime("%Y-%m-%d")
    tmp_dir = tempfile.mkdtemp()
    charts_for_pdf = []

    def create_pdf_from_charts(charts_data, output_path, title, config_info, logo_path_inner):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font('helvetica', 'B', 16) 
        
        pdf.add_page()
        if os.path.exists(logo_path_inner):
            pdf.image(logo_path_inner, x=10, y=10, w=30)
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
                if os.path.exists(logo_path_inner):
                    pdf.image(logo_path_inner, x=10, y=8, w=25)
                pdf.set_font("helvetica", 'B', 11) 
                pdf.set_y(35)
                if page_project_name:
                    pdf.cell(0, 10, f"Project: {page_project_name}", ln=True, align='C')
                pdf.cell(0, 10, chart_title, ln=True, align='C')
                pdf.image(img_path, x=10, y=45, w=190)

        pdf.output(output_path, "F")
        print(f"DEBUG: PDF report generated at {output_path}")

    try:
        projects = df['Project name'].unique() 

        config_info = {
            "Mode": config.get('mode', 'N/A').capitalize(),
            "Years": ', '.join(map(str, config.get('years', []))) if config.get('years') else str(config.get('year', 'N/A')),
            "Months": ', '.join(config.get('months', [])) if config.get('months') else "All",
            "Projects Included": ', '.join(config['project_filter_df']['Project Name']) if 'project_filter_df' in config and not config['project_filter_df'].empty else "No projects selected or found"
        }

        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'Liberation Sans']
        plt.rcParams['axes.unicode_minus'] = False 

        for project in projects:
            safe_project = sanitize_filename(project)
            df_proj = df[df['Project name'] == project]

            if 'Workcentre' in df_proj.columns and not df_proj['Workcentre'].empty:
                workcentre_summary = df_proj.groupby('Workcentre')['Hours'].sum().sort_values(ascending=False)
                if not workcentre_summary.empty and workcentre_summary.sum() > 0:
                    fig, ax = plt.subplots(figsize=(10, 5))
                    workcentre_summary.plot(kind='barh', color='skyblue', ax=ax)
                    ax.set_title(f"{project} - Hours by Workcentre", fontsize=9)
                    ax.tick_params(axis='y', labelsize=8)
                    ax.set_xlabel("Hours")
                    ax.set_ylabel("Workcentre")
                    wc_img_path = os.path.join(tmp_dir, f"{safe_project}_wc.png")
                    plt.tight_layout()
                    fig.savefig(wc_img_path, dpi=150)
                    plt.close(fig)
                    charts_for_pdf.append((wc_img_path, f"{project} - Hours by Workcentre", project))

            if 'Task' in df_proj.columns and not df_proj['Task'].empty:
                task_summary = df_proj.groupby('Task')['Hours'].sum().sort_values(ascending=False)
                if not task_summary.empty and task_summary.sum() > 0:
                    fig, ax = plt.subplots(figsize=(10, 6))
                    task_summary.plot(kind='barh', color='lightgreen', ax=ax)
                    ax.set_title(f"{project} - Hours by Task", fontsize=9)
                    ax.tick_params(axis='y', labelsize=8)
                    ax.set_xlabel("Hours")
                    ax.set_ylabel("Task")
                    task_img_path = os.path.join(tmp_dir, f"{safe_project}_task.png")
                    plt.tight_layout()
                    fig.savefig(task_img_path, dpi=150)
                    plt.close(fig)
                    charts_for_pdf.append((task_img_path, f"{project} - Hours by Task", project))

        if not charts_for_pdf:
            print("Cảnh báo: Không có biểu đồ nào được tạo để đưa vào PDF. PDF có thể trống.")
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font('helvetica', 'B', 16)
            pdf.cell(0, 10, "TRIAC TIME REPORT - STANDARD", ln=True, align='C')
            pdf.set_font("helvetica", '', 12)
            pdf.cell(0, 10, f"Generated on: {today_str}", ln=True, align='C')
            pdf.ln(10)
            pdf.set_font("helvetica", '', 11)
            for key, value in config_info.items():
                pdf.cell(0, 7, f"{key}: {value}", ln=True, align='C')
            pdf.cell(0, 10, "No charts generated for this report.", ln=True, align='C')
            pdf.output(pdf_report_path, "F")
            return True
            
        create_pdf_from_charts(charts_for_pdf, pdf_report_path, "TRIAC TIME REPORT - STANDARD", config_info, logo_path)
        return True
    except Exception as e:
        print(f"Lỗi khi tạo báo cáo PDF: {e}")
        return False
    finally:
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)

def apply_comparison_filters(df_raw, comparison_config, comparison_mode):
    """Áp dụng bộ lọc và tạo DataFrame tóm tắt cho báo cáo so sánh."""
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
        return pd.DataFrame(), "Vui lòng chọn ít nhất một dự án để so sánh."

    if df_filtered.empty:
        return pd.DataFrame(), f"Không tìm thấy dữ liệu cho chế độ so sánh: {comparison_mode} với các lựa chọn hiện tại."

    title = ""

    if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
        if len(years) != 1 or len(months) != 1 or len(selected_projects) < 2:
            return pd.DataFrame(), "Vui lòng chọn MỘT năm, MỘT tháng và ít nhất HAI dự án cho chế độ này."
        
        df_comparison = df_filtered.groupby('Project name')['Hours'].sum().reset_index()
        df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
        title = f"So sánh giờ giữa các dự án trong {months[0]}, năm {years[0]}"
        return df_comparison, title

    elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        if len(years) != 1 or len(selected_projects) < 2:
            return pd.DataFrame(), "Vui lòng chọn MỘT năm và ít nhất HAI dự án cho chế độ này."
        
        df_comparison = df_filtered.groupby(['Project name', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
        
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        existing_months = [m for m in month_order if m in df_comparison.columns]
        df_comparison = df_comparison[existing_months]

        df_comparison = df_comparison.reset_index().rename(columns={'index': 'Project Name'})
        
        df_comparison['Total Hours'] = df_comparison[existing_months].sum(axis=1)

        df_comparison.loc['Total'] = df_comparison[existing_months + ['Total Hours']].sum()
        df_comparison.loc['Total', 'Project Name'] = 'Total'

        title = f"So sánh giờ giữa các dự án trong năm {years[0]} (theo tháng)"
        return df_comparison, title

    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
        if len(selected_projects) != 1 or (len(months) == 0 and len(years) == 0) or (len(months) == 1 and len(years) == 1): 
            return pd.DataFrame(), "Vui lòng chọn MỘT dự án và ít nhất HAI tháng HOẶC HAI năm để so sánh."
        
        if years and not months:
            df_comparison = df_filtered.groupby('Year')['Hours'].sum().reset_index()
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            df_comparison['Year'] = df_comparison['Year'].astype(str)
            title = f"Giờ cho dự án {selected_projects[0]} qua các năm"
        elif months and not years:
            df_comparison = df_filtered.groupby(['MonthName'])['Hours'].sum().reset_index()
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            
            month_order_df = pd.DataFrame({'MonthName': ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']})
            df_comparison = pd.merge(month_order_df, df_comparison, on='MonthName', how='left').fillna(0)
            title = f"Giờ cho dự án {selected_projects[0]} qua các tháng"
        elif years and months:
            df_comparison = df_filtered.groupby(['Year', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
            month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
            existing_months = [m for m in month_order if m in df_comparison.columns]
            df_comparison = df_comparison[existing_months]
            df_comparison = df_comparison.reset_index()
            
            df_comparison['Total Hours'] = df_comparison[existing_months].sum(axis=1)

            df_comparison.loc['Total'] = df_comparison[existing_months + ['Total Hours']].sum()
            df_comparison.loc['Total', 'Year'] = 'Total'
            
            title = f"Giờ cho dự án {selected_projects[0]} qua các năm và tháng"
        else:
            return pd.DataFrame(), "Lựa chọn thời gian không hợp lệ cho chế độ so sánh này."

        return df_comparison, title
        
    return pd.DataFrame(), "Chế độ so sánh không hợp lệ."

def export_comparison_report(df_comparison, comparison_config, output_file_path, comparison_mode):
    """Xuất báo cáo so sánh ra file Excel."""
    try:
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            if df_comparison.empty:
                empty_df_for_excel = pd.DataFrame({"Message": ["Không có dữ liệu để hiển thị với các bộ lọc đã chọn."]})
                empty_df_for_excel.to_excel(writer, sheet_name='Comparison Report', index=False)
            else:
                df_comparison.to_excel(writer, sheet_name='Comparison Report', index=False)  

            wb = writer.book
            ws = wb['Comparison Report']

            data_last_row = ws.max_row
            info_row = data_last_row + 2 

            ws.merge_cells(start_row=info_row, start_column=1, end_row=info_row, end_column=4)
            ws.cell(row=info_row, column=1, value=f"BÁO CÁO SO SÁNH: {comparison_mode}").font = ws.cell(row=info_row, column=1).font.copy(bold=True, size=14)
            info_row += 1

            ws.cell(row=info_row, column=1, value="Năm:").font = ws.cell(row=info_row, column=1).font.copy(bold=True)
            ws.cell(row=info_row, column=2, value=', '.join(map(str, comparison_config.get('years', []))))
            info_row += 1
            ws.cell(row=info_row, column=1, value="Tháng:").font = ws.cell(row=info_row, column=1).font.copy(bold=True)
            ws.cell(row=info_row, column=2, value=', '.join(comparison_config.get('months', [])))
            info_row += 1
            ws.cell(row=info_row, column=1, value="Dự án:").font = ws.cell(row=info_row, column=1).font.copy(bold=True)
            ws.cell(row=info_row, column=2, value=', '.join(comparison_config.get('selected_projects', [])))

            if not df_comparison.empty and len(df_comparison) > 0:
                chart = None
                data_start_row = 2 
                
                df_chart_data = df_comparison.copy()
                if 'Project Name' in df_chart_data.columns and 'Total' in df_chart_data['Project Name'].values:
                    df_chart_data = df_chart_data[df_chart_data['Project Name'] != 'Total']
                elif 'Year' in df_chart_data.columns and 'Total' in df_chart_data['Year'].values:
                    df_chart_data = df_chart_data[df_chart_data['Year'] != 'Total']
                
                if df_chart_data.empty: 
                    print("Không có đủ dữ liệu để vẽ biểu đồ so sánh sau khi loại bỏ hàng tổng.")
                    wb.save(output_file_path)
                    return True

                max_row_chart = data_start_row + len(df_chart_data) - 1

                if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
                    chart = BarChart()
                    chart.title = "So sánh giờ theo dự án"
                    chart.x_axis.title = "Dự án"
                    chart.y_axis.title = "Giờ"
                    
                    data_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Total Hours') + 1, min_row=data_start_row, max_row=max_row_chart)
                    cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Project name') + 1, min_row=data_start_row, max_row=max_row_chart) 
                    
                    chart.add_data(data_ref, titles_from_data=False) 
                    chart.set_categories(cats_ref)
                
                elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
                    chart = LineChart()
                    chart.title = "So sánh giờ theo dự án và tháng"
                    chart.x_axis.title = "Tháng"
                    chart.y_axis.title = "Giờ"

                    month_cols = [col for col in df_comparison.columns if col not in ['Project Name', 'Total Hours']]
                    
                    cats_ref = Reference(ws, min_col=ws[1][0].column + 1, 
                                            min_row=1, 
                                            max_col=ws[1][-2].column)
                    
                    for r_idx, project_name in enumerate(df_chart_data['Project Name']):
                        series_ref = Reference(ws, min_col=ws[1][0].column + 1, 
                                                min_row=data_start_row + r_idx, 
                                                max_col=ws[1][-2].column, 
                                                max_row=data_start_row + r_idx)
                        title_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Project Name') + 1, 
                                            min_row=data_start_row + r_idx, 
                                            max_row=data_start_row + r_idx)
                        chart.add_data(series_ref, titles_from_data=True)
                        chart.series[r_idx].title = title_ref
                    
                    chart.set_categories(cats_ref)

                elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
                    chart = LineChart()
                    chart.title = f"Giờ cho dự án {comparison_config['selected_projects'][0]} qua thời gian"
                    chart.y_axis.title = "Giờ"

                    if 'Year' in df_comparison.columns and 'MonthName' in df_comparison.columns:
                        chart.x_axis.title = "Năm-Tháng"
                        cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Year') + 1, 
                                            min_row=data_start_row, 
                                            max_col=df_comparison.columns.get_loc('MonthName') + 1,
                                            max_row=max_row_chart)
                    elif 'Year' in df_comparison.columns:
                        chart.x_axis.title = "Năm"
                        cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Year') + 1, min_row=data_start_row, max_row=max_row_chart)
                    elif 'MonthName' in df_comparison.columns:
                        chart.x_axis.title = "Tháng"
                        cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('MonthName') + 1, min_row=data_start_row, max_row=max_row_chart)
                    else:
                        raise ValueError("Không tìm thấy kích thước thời gian hợp lệ cho các danh mục biểu đồ.")

                    value_cols = [col for col in df_comparison.columns if col not in ['Year', 'MonthName', 'Total Hours', 'Project Name']]
                    if 'Total Hours' in df_comparison.columns:
                        series_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Total Hours') + 1, min_row=data_start_row, max_row=max_row_chart)
                        chart.add_data(series_ref, titles_from_data=False) 
                    elif value_cols:
                        for col_name in value_cols:
                            col_idx = df_comparison.columns.get_loc(col_name) + 1
                            series_ref = Reference(ws, min_col=col_idx, min_row=data_start_row, max_row=max_row_chart)
                            chart.add_data(series_ref, titles_from_data=True) 
                    
                    chart.set_categories(cats_ref)

                if chart: 
                    chart_placement_row = info_row + 2
                    ws.add_chart(chart, f"A{chart_placement_row}")

            wb.save(output_file_path)
            return True
    except Exception as e:
        print(f"Lỗi khi xuất báo cáo so sánh ra Excel: {e}")
        return False

def export_comparison_pdf_report(df_comparison, comparison_config, pdf_file_path, comparison_mode, logo_path):
    """Xuất báo cáo PDF so sánh với biểu đồ."""
    tmp_dir = tempfile.mkdtemp()
    charts_for_pdf = []

    def create_pdf_from_charts_comp(charts_data, output_path, title, config_info, logo_path_inner):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font('helvetica', 'B', 16) 

        pdf.add_page()
        if os.path.exists(logo_path_inner):
            pdf.image(logo_path_inner, x=10, y=10, w=30)
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
                if os.path.exists(logo_path_inner):
                    pdf.image(logo_path_inner, x=10, y=8, w=25)
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
        
        if 'Project Name' in df_plot.columns and 'Total' in df_plot['Project Name'].values:
            df_plot = df_plot[df_plot['Project Name'] != 'Total']
        elif 'Year' in df_plot.columns and 'Total' in df_plot['Year'].values:
            df_plot = df_plot[df_plot['Year'] != 'Total']
        
        if df_plot.empty:
            print(f"DEBUG: df_plot is empty for mode '{mode}' after dropping 'Total'. Skipping chart creation.")
            plt.close(fig) 
            return None 

        ax.set_ylim(bottom=0)
        
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'Liberation Sans']
        plt.rcParams['axes.unicode_minus'] = False

        if mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
            ax.bar(df_plot['Project name'], df_plot['Total Hours'], color='skyblue')
            ax.set_xticks(df_plot['Project name'])
            ax.tick_params(axis='x', rotation=45, ha='right', labelsize=9)
            ax.tick_params(axis='y', labelsize=8)  
        elif mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
            month_columns = [col for col in df_plot.columns if col not in ['Project Name', 'Total Hours']]
            df_plot[month_columns] = df_plot[month_columns].apply(pd.to_numeric, errors='coerce').fillna(0)
            df_plot.set_index('Project Name', inplace=True) 
            
            df_plot.T.plot(kind='line', ax=ax, colormap='viridis', marker='o')  
            
            ax.set_xticks(range(len(df_plot.columns)))  
            ax.set_xticklabels(df_plot.columns, rotation=45, ha='right', labelsize=9)
            
            ax.legend(title="Dự án", bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=8) 
            ax.tick_params(axis='y', labelsize=8)
            x_label = "Tháng" 
            
        elif mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
            data_cols = [col for col in df_plot.columns if col not in ['Year', 'MonthName', 'Total Hours']]
            
            if 'Year' in df_plot.columns and 'MonthName' in df_plot.columns:
                month_to_num = {name: i for i, name in enumerate(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], 1)}
                df_plot['MonthNum'] = df_plot['MonthName'].map(month_to_num)
                df_plot['YearMonth'] = df_plot['Year'].astype(str) + '-' + df_plot['MonthNum'].astype(str).str.zfill(2)
                df_plot = df_plot.sort_values(by=['Year', 'MonthNum'])
                df_plot.set_index('YearMonth', inplace=True)
                x_label = "Thời gian (Năm-Tháng)"  
            elif 'Year' in df_plot.columns:
                df_plot = df_plot.sort_values(by='Year')
                df_plot.set_index('Year', inplace=True)
                x_label = "Năm"  
            elif 'MonthName' in df_plot.columns:
                month_order_list = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
                df_plot['MonthName_ordered'] = pd.Categorical(df_plot['MonthName'], categories=month_order_list, ordered=True)
                df_plot = df_plot.sort_values(by='MonthName_ordered')
                df_plot.set_index('MonthName', inplace=True)
                x_label = "Tháng"  

            if 'Total Hours' in df_plot.columns:
                ax.plot(df_plot.index, df_plot['Total Hours'], marker='o', color='salmon')  
            elif data_cols:
                df_plot[data_cols] = df_plot[data_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
                df_plot[data_cols].plot(kind='line', ax=ax, colormap='plasma', marker='o') 
                ax.legend(title="Tháng", bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=8)

            plt.xticks(rotation=45, ha='right', labelsize=9)
            ax.tick_params(axis='y', labelsize=8)

        ax.set_title(title, fontsize=12)
        ax.set_xlabel(x_label, fontsize=10)  
        ax.set_ylabel(y_label, fontsize=10)
        plt.tight_layout()
        fig.savefig(img_path, dpi=200)
        plt.close(fig)
        return img_path

    try:
        config_info = {
            "Chế độ so sánh": comparison_mode,
            "Năm": ', '.join(map(str, comparison_config.get('years', []))) or "N/A",
            "Tháng": ', '.join(comparison_config.get('months', [])) or "Tất cả",
            "Dự án": ', '.join(comparison_config.get('selected_projects', [])) or "Không có dự án nào được chọn"
        }

        chart_title = ""
        x_label = ""
        y_label = "Giờ"
        chart_data_to_plot = df_comparison.copy() 
        page_project_name = None

        if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
            chart_title = f"So sánh giờ giữa các dự án trong {comparison_config['months'][0]}, năm {comparison_config['years'][0]}"
            x_label = "Dự án"
        elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
            chart_title = f"So sánh giờ giữa các dự án trong năm {comparison_config['years'][0]}"
            x_label = "Tháng"
        elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
            chart_title = f"Giờ cho dự án {comparison_config['selected_projects'][0]} qua thời gian"
            page_project_name = comparison_config['selected_projects'][0] 
            
            if 'Year' in chart_data_to_plot.columns and 'MonthName' in chart_data_to_plot.columns:
                x_label = "Năm-Tháng"
            elif 'Year' in chart_data_to_plot.columns:
                x_label = "Năm"
            elif 'MonthName' in chart_data_to_plot.columns:
                x_label = "Tháng"
            else:
                x_label = "Thời gian" 

        chart_img_path = os.path.join(tmp_dir, "comparison_chart.png")
        generated_path = create_comparison_chart(chart_data_to_plot, comparison_mode, chart_title, x_label, y_label, chart_img_path)
        
        if generated_path:
            charts_for_pdf.append((generated_path, chart_title, page_project_name))
        else:
            print("Không tạo được biểu đồ so sánh. PDF có thể sẽ trống biểu đồ.")
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font('helvetica', 'B', 16)
            pdf.cell(0, 10, "TRIAC TIME REPORT - COMPARISON", ln=True, align='C')
            pdf.set_font("helvetica", '', 12)
            pdf.cell(0, 10, f"Generated on: {datetime.datetime.today().strftime('%Y-%m-%d')}", ln=True, align='C')
            pdf.ln(10)
            pdf.set_font("helvetica", '', 11)
            for key, value in config_info.items():
                pdf.cell(0, 7, f"{key}: {value}", ln=True, align='C')
            pdf.cell(0, 10, "No charts generated for this report.", ln=True, align='C')
            pdf.output(pdf_file_path, "F")
            return True

        create_pdf_from_charts_comp(charts_for_pdf, pdf_file_path, "TRIAC TIME REPORT - COMPARISON", config_info, logo_path)
        return True

    except Exception as e:
        print(f"Lỗi khi xuất báo cáo PDF so sánh: {e}")
        return False
    finally:
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)
