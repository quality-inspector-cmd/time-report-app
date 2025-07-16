import pandas as pd
import datetime
import os
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF
from matplotlib import pyplot as plt
import matplotlib.font_manager as fm # Import the font_manager
import tempfile
import re
import shutil

# --- Cấu hình font cho Matplotlib (thêm vào đầu file) ---
# Đặt font mặc định cho matplotlib để tránh lỗi font Arial/DejaVuSans
# Sử dụng font 'sans-serif' và liệt kê các font có khả năng có sẵn trên môi trường Linux
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Liberation Sans', 'Helvetica', 'Arial']
# Đặt font cho ký tự unicode nếu cần, thường DejaVu Sans hỗ trợ tốt
plt.rcParams['axes.unicode_minus'] = False # Đảm bảo dấu trừ hiển thị đúng

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
        # Sử dụng font mặc định của FPDF để tránh lỗi font
        # Nếu bạn cần hỗ trợ tiếng Việt, bạn cần tải xuống một font unicode (ví dụ: DejaVuSansCondensed.ttf)
        # và upload nó cùng với ứng dụng của bạn, sau đó uncomment và sửa đường dẫn
        # try:
        #     pdf.add_font('DejaVuSans', '', 'DejaVuSansCondensed.ttf', uni=True)
        #     pdf.set_font('DejaVuSans', '', 16)
        # except RuntimeError:
        #     print("Warning: DejaVuSansCondensed.ttf not found. PDF might not display Vietnamese characters correctly.")
        #     # Fallback to Helvetica if custom font is not found
        pdf.set_font('helvetica', 'B', 16) # Sử dụng font mặc định của FPDF để tránh lỗi

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
                # Dùng Helvetica cho tiêu đề biểu đồ trong PDF
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

    df_comparison = pd.DataFrame()
    title = ""
    error_message = ""

    if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
        if len(years) != 1 or len(months) != 1 or len(selected_projects) < 2:
            error_message = "Cần chọn MỘT năm, MỘT tháng và ít nhất HAI dự án cho chế độ này."
            print(f"DEBUG - apply_comparison_filters: Condition failed for 'Compare Projects in a Month'. {error_message}")
            return pd.DataFrame(), error_message
        
        df_comparison = df_filtered.groupby('Project name')['Hours'].sum().reset_index()
        df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
        title = f"So sánh giờ giữa các dự án trong {months[0]}, năm {years[0]}"
        print(f"DEBUG - apply_comparison_filters (Month Comparison): df_comparison head:\n{df_comparison.head()}")
        print(f"DEBUG - apply_comparison_filters (Month Comparison): df_comparison dtypes:\n{df_comparison.dtypes}") # THÊM DÒNG NÀY
        print(f"DEBUG - apply_comparison_filters (Month Comparison): 'Total Hours' unique values: {df_comparison['Total Hours'].unique()}") # THÊM DÒNG NÀY
        return df_comparison, title

    elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        if len(years) != 1 or len(selected_projects) < 2:
            error_message = "Cần chọn MỘT năm và ít nhất HAI dự án cho chế độ này."
            print(f"DEBUG - apply_comparison_filters: Condition failed for 'Compare Projects in a Year'. {error_message}")
            return pd.DataFrame(), error_message
        
        df_comparison = df_filtered.groupby(['Project name', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
        
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        existing_months = [m for m in month_order if m in df_comparison.columns]
        df_comparison = df_comparison[existing_months]

        df_comparison = df_comparison.reset_index().rename(columns={'Project name': 'Project Name'}) # Đảm bảo tên cột là 'Project Name'
        
        title = f"So sánh giờ giữa các dự án trong năm {years[0]} (theo tháng)"
        print(f"DEBUG - apply_comparison_filters (Year Comparison): df_comparison head:\n{df_comparison.head()}")
        print(f"DEBUG - apply_comparison_filters (Year Comparison): df_comparison dtypes:\n{df_comparison.dtypes}") # THÊM DÒNG NÀY
        for col in df_comparison.columns:
            if col not in ['Project Name']:
                print(f"DEBUG - apply_comparison_filters (Year Comparison): '{col}' unique values: {df_comparison[col].unique()}") # THÊM DÒNG NÀY CHO TỪNG CỘT THÁNG
        return df_comparison, title

    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
        print(f"DEBUG - apply_comparison_filters: Inside 'Compare One Project Over Time' mode.")
        if len(selected_projects) != 1:
            error_message = "Chế độ 'So Sánh Một Dự Án Qua Các Tháng/Năm' yêu cầu chọn CHỈ MỘT dự án."
            print(f"DEBUG - apply_comparison_filters: More than one project selected. {error_message}")
            return pd.DataFrame(), error_message
        
        if not years and not months:
            error_message = "Cần chọn ít nhất một năm HOẶC một tháng để so sánh."
            print(f"DEBUG - apply_comparison_filters: No years or months selected. {error_message}")
            return pd.DataFrame(), error_message

        # Trường hợp 1: 1 dự án, 1 năm, nhiều tháng (chỉ filter tháng trong năm đó)
        if len(years) == 1 and len(months) >= 1: # Điều chỉnh: >= 1 tháng, vì có thể chỉ có 1 tháng nhưng vẫn muốn so sánh nếu user chọn
            print(f"DEBUG - apply_comparison_filters: Processing 1 project, 1 year ({years[0]}), multiple months ({months}).")
            df_comparison = df_filtered.groupby(['Year', 'MonthName'])['Hours'].sum().reset_index()
            # Đảm bảo tất cả các tháng được chọn đều có trong DataFrame, điền 0 nếu không có
            month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
            
            # Tạo DataFrame đầy đủ các tháng mong muốn
            full_month_df = pd.DataFrame({'MonthName': [m for m in month_order if m in months]})
            df_comparison = pd.merge(full_month_df, df_comparison, on='MonthName', how='left').fillna(0)
            df_comparison['Year'] = years[0] # Đảm bảo cột Year vẫn tồn tại sau merge/fill
            
            # Sắp xếp các tháng theo thứ tự tự nhiên
            df_comparison['MonthName'] = pd.Categorical(df_comparison['MonthName'], categories=month_order, ordered=True)
            df_comparison = df_comparison.sort_values(by='MonthName')
            
            # Đổi tên cột 'Hours' thành 'Total Hours' để nhất quán cho biểu đồ
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            
            title = f"Giờ của dự án {selected_projects[0]} trong năm {years[0]} theo tháng"
            
        # Trường hợp 2: 1 dự án, nhiều năm, không chọn tháng cụ thể (tổng hợp theo năm)
        elif len(years) >= 1 and not months: # Điều chỉnh: >= 1 năm, có thể so sánh 1 năm tổng hợp
            print(f"DEBUG - apply_comparison_filters: Processing 1 project, multiple years ({years}), no specific months.")
            df_comparison = df_filtered.groupby(['Year'])['Hours'].sum().reset_index()
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            title = f"Giờ của dự án {selected_projects[0]} qua các năm"
            
        # Trường hợp 3: 1 dự án, nhiều tháng, không chọn năm cụ thể (tổng hợp theo tháng trên tất cả các năm có dữ liệu)
        elif len(months) >= 1 and not years: # Điều chỉnh: >= 1 tháng
            print(f"DEBUG - apply_comparison_filters: Processing 1 project, multiple months ({months}), no specific years.")
            df_comparison = df_filtered.groupby(['MonthName'])['Hours'].sum().reset_index()
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
            full_month_df = pd.DataFrame({'MonthName': [m for m in month_order if m in months]})
            df_comparison = pd.merge(full_month_df, df_comparison, on='MonthName', how='left').fillna(0)
            df_comparison['MonthName'] = pd.Categorical(df_comparison['MonthName'], categories=month_order, ordered=True)
            df_comparison = df_comparison.sort_values(by='MonthName')
            title = f"Giờ của dự án {selected_projects[0]} theo tháng (Tổng hợp các năm)"
            
        else: # Các trường hợp không đủ dữ liệu hoặc không được hỗ trợ rõ ràng
            error_message = "Cần chọn ít nhất HAI tháng HOẶC HAI năm để so sánh một dự án qua thời gian."
            print(f"DEBUG - apply_comparison_filters: No valid comparison combination for 'Compare One Project Over Time'. {error_message}")
            return pd.DataFrame(), error_message
        
        print(f"DEBUG - apply_comparison_filters (Compare One Project Over Time): Final df_comparison head:\n{df_comparison.head()}")
        print(f"DEBUG - apply_comparison_filters (Compare One Project Over Time): Final df_comparison columns: {df_comparison.columns.tolist()}")
        print(f"DEBUG - apply_comparison_filters (Compare One Project Over Time): df_comparison dtypes:\n{df_comparison.dtypes}") # THÊM DÒNG NÀY
        for col in df_comparison.columns: # THÊM DÒNG NÀY ĐỂ KIỂM TRA TẤT CẢ CÁC CỘT
            print(f"DEBUG - apply_comparison_filters (Compare One Project Over Time): '{col}' unique values: {df_comparison[col].unique()}")

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
            chart = BarChart() # BarChart vẫn dùng được cho nhiều chế độ
            chart.y_axis.title = "Giờ"
            
            # Cập nhật data_start_row để tính toán động
            data_start_row = ws.max_row + 2 # Bắt đầu sau phần thông tin config và 1 dòng trống
            
            # Viết header của df_comparison vào Excel trước
            for col_idx, col_name in enumerate(df_comparison.columns):
                ws.cell(row=data_start_row, column=col_idx + 1, value=col_name)
            
            # Viết dữ liệu của df_comparison vào Excel
            for r_idx, r in enumerate(dataframe_to_rows(df_comparison, index=False, header=False)):
                for c_idx, cell_val in enumerate(r):
                    ws.cell(row=data_start_row + 1 + r_idx, column=c_idx + 1, value=cell_val)
            
            # Cập nhật max_row sau khi ghi dữ liệu
            current_max_row = data_start_row + len(df_comparison)
            
            if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
                chart.title = "So sánh giờ theo Dự án"
                chart.x_axis.title = "Dự án"
                
                # Cập nhật Reference dựa trên data_start_row và current_max_row
                data_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Total Hours') + 1, min_row=data_start_row + 1, max_row=current_max_row)
                cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Project name') + 1, min_row=data_start_row + 1, max_row=current_max_row)
                
                chart.add_data(data_ref, titles_from_data=False) 
                chart.set_categories(cats_ref)
            
            elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
                chart.title = "So sánh giờ giữa các dự án và Tháng"
                chart.x_axis.title = "Dự án" # Trong biểu đồ cột/line này, trục x là dự án
                
                value_cols = [col for col in df_comparison.columns if col not in ['Project Name']]
                
                # Cần thêm từng series cho mỗi tháng
                for idx, col_name in enumerate(value_cols):
                    # min_row là hàng chứa dữ liệu đầu tiên (không phải header)
                    series = Reference(ws, min_col=df_comparison.columns.get_loc(col_name) + 1, min_row=data_start_row + 1, max_row=current_max_row)
                    chart.add_data(series, titles_from_data=True, from_rows=False) # from_rows=False if data is in columns
                
                # categories là tên các dự án
                cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Project Name') + 1, min_row=data_start_row + 1, max_row=current_max_row)
                chart.set_categories(cats_ref)
                
                # Đặt loại biểu đồ là BarChart (stacked) hoặc ColumnChart (grouped)
                chart.type = "col" # Mặc định là column chart (bar dọc)
                chart.shape = 4 # Có thể thử các shape khác nhau
            
            elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
                chart.title = "So sánh giờ của một dự án qua thời gian"
                chart.x_axis.title = "Thời gian" # x_axis title sẽ được đặt cụ thể hơn trong chart_create function

                # Dữ liệu cho biểu đồ sẽ là Total Hours, category là Year/MonthName/YearMonth
                data_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Total Hours') + 1, min_row=data_start_row + 1, max_row=current_max_row)
                chart.add_data(data_ref, titles_from_data=False)

                if 'Year' in df_comparison.columns and len(df_comparison['Year'].unique()) > 1:
                    cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('Year') + 1, min_row=data_start_row + 1, max_row=current_max_row)
                    chart.x_axis.title = "Năm"
                elif 'MonthName' in df_comparison.columns:
                     cats_ref = Reference(ws, min_col=df_comparison.columns.get_loc('MonthName') + 1, min_row=data_start_row + 1, max_row=current_max_row)
                     chart.x_axis.title = "Tháng"
                else: # Fallback to the first column if no specific time column
                    cats_ref = Reference(ws, min_col=1, min_row=data_start_row + 1, max_row=current_max_row)
                    chart.x_axis.title = df_comparison.columns[0] # Tên cột đầu tiên
                
                chart.set_categories(cats_ref)
                chart.type = "col" # Vẫn có thể là col chart hoặc line chart tùy theo data

            ws.add_chart(chart, f"A{current_max_row + 5}") # Đặt biểu đồ sau dữ liệu và config
            
        wb.save(output_file)
        return True

def export_comparison_pdf_report(df_comparison, comparison_config, path_dict, comparison_mode):
    pdf_file = path_dict['comparison_pdf_report']
    tmp_dir = tempfile.mkdtemp()
    charts_for_pdf = []

    def create_pdf_from_charts(charts_data, output_path, title, config_info, logo_path="triac_logo.png"):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        # Sử dụng font mặc định của FPDF để tránh lỗi font
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
        print(f"  df_plot dtypes:\n{df_plot.dtypes}") # Thêm dòng này để kiểm tra dtype

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
            if 'Project Name' not in df_plot.columns: # KIỂM TRA ĐÂY
                print(f"ERROR: 'Project Name' column not found in df_plot for 'Compare Projects in a Year' mode. Columns: {df_plot.columns.tolist()}")
                plt.close(fig)
                return None

            # df_plot.set_index('Project Name', inplace=True) # KHÔNG set index ở đây nữa
            
            # Đảm bảo các cột tháng là số
            month_columns = [col for col in df_plot.columns if col not in ['Project Name']]
            for col in month_columns:
                df_plot[col] = pd.to_numeric(df_plot[col], errors='coerce').fillna(0) # Chuyển đổi an toàn hơn

            # Chuyển đổi định dạng để vẽ biểu đồ đường
            # Chuyển đổi df_plot từ wide format sang long format để plot hiệu quả hơn
            df_melted = df_plot.melt(id_vars=['Project Name'], var_name='Month', value_name='Hours')
            
            month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
            df_melted['Month'] = pd.Categorical(df_melted['Month'], categories=month_order, ordered=True)
            df_melted = df_melted.sort_values(by='Month')

            # Vẽ biểu đồ đường
            for project in df_melted['Project Name'].unique():
                df_project = df_melted[df_melted['Project Name'] == project]
                ax.plot(df_project['Month'], df_project['Hours'], marker='o', label=project)
            
            ax.set_xticks(df_melted['Month'].unique())
            ax.set_xticklabels(df_melted['Month'].unique(), rotation=45, ha='right')
            
            ax.legend(title="Dự án", bbox_to_anchor=(1.05, 1), loc='upper left') 
            ax.tick_params(axis='y', labelsize=8)
            x_label = "Tháng" 
            
        elif mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
            print(f"DEBUG: Inside 'Compare One Project Over Time' plotting logic.")
            # Đảm bảo cột 'Total Hours' là số
            df_plot['Total Hours'] = pd.to_numeric(df_plot['Total Hours'], errors='coerce').fillna(0)

            # Trường hợp 1: 1 dự án, 1 năm, nhiều tháng (df_plot có cột 'Year', 'MonthName', 'Total Hours')
            # Cột Year có thể là số (float hoặc int) hoặc object (sau khi reset_index)
            # Dòng code trong apply_comparison_filters đảm bảo df_comparison đã có Year và MonthName và Total Hours
            if 'Year' in df_plot.columns and 'MonthName' in df_plot.columns and len(df_plot['Year'].unique()) == 1 and len(df_plot['MonthName'].unique()) > 1:
                print(f"DEBUG: Plotting Case 1: 1 project, 1 year, multiple months.")
                
                # Sắp xếp các tháng theo thứ tự tự nhiên
                month_order_list = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
                df_plot['MonthName_ordered'] = pd.Categorical(df_plot['MonthName'], categories=month_order_list, ordered=True)
                df_plot = df_plot.sort_values(by='MonthName_ordered')
                
                ax.plot(df_plot['MonthName'], df_plot['Total Hours'], marker='o', color='salmon')
                x_label = f"Tháng ({df_plot['Year'].iloc[0]})"
                ax.set_xticks(df_plot['MonthName'])
            
            # Trường hợp 2: 1 dự án, nhiều năm (df_plot có cột 'Year' và 'Total Hours')
            elif 'Year' in df_plot.columns and 'Total Hours' in df_plot.columns and len(df_plot['Year'].unique()) > 1:
                print(f"DEBUG: Plotting Case 2: 1 project, multiple years.")
                df_plot['Year'] = pd.to_numeric(df_plot['Year'], errors='coerce') # Đảm bảo năm là số
                df_plot = df_plot.sort_values(by='Year')
                ax.plot(df_plot['Year'], df_plot['Total Hours'], marker='o', color='salmon')
                x_label = "Năm"
                ax.set_xticks(df_plot['Year'])

            # Trường hợp 3: 1 dự án, nhiều tháng, không chọn năm cụ thể (tổng hợp theo tháng trên tất cả các năm có dữ liệu)
            elif 'MonthName' in df_plot.columns and 'Total Hours' in df_plot.columns and len(df_plot['MonthName'].unique()) > 1:
                print(f"DEBUG: Plotting Case 3: 1 project, multiple months (aggregated across years).")
                month_order_list = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
                df_plot['MonthName_ordered'] = pd.Categorical(df_plot['MonthName'], categories=month_order_list, ordered=True)
                df_plot = df_plot.sort_values(by='MonthName_ordered')
                ax.plot(df_plot['MonthName'], df_plot['Total Hours'], marker='o', color='salmon')
                x_label = "Tháng"
                ax.set_xticks(df_plot['MonthName'])
            
            else:
                print(f"Warning: No suitable plotting data structure found for 'Compare One Project Over Time' mode based on current df_plot. Columns: {df_plot.columns.tolist()}")
                ax.text(0.5, 0.5, "Không có dữ liệu để vẽ biểu đồ line.", horizontalalignment='center', verticalalignment='center', transform=ax.transAxes, fontsize=12)
                plt.close(fig) # Đóng figure nếu không vẽ được gì
                return None


            plt.xticks(rotation=45, ha='right')
            ax.tick_params(axis='y', labelsize=8)

        ax.set_title(title)
        ax.set_xlabel(x_label) 
        ax.set_ylabel(y_label)
        plt.tight_layout()
        
        # Đảm bảo chỉ lưu ảnh nếu có dữ liệu để vẽ
        if not df_plot.empty:
            fig.savefig(img_path, dpi=200)
            print(f"DEBUG: Chart saved to {img_path}")
            plt.close(fig)
            return img_path
        else:
            plt.close(fig)
            return None


    try:
        chart_title = f"Báo cáo so sánh: {comparison_mode}"
        x_label = ""
        y_label = "Tổng số giờ"

        # Pass x_label and y_label to create_comparison_chart, it will adjust internally
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
