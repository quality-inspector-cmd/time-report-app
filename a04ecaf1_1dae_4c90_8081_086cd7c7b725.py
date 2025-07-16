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
    tmp_dir = tempfile.mkdtemp() # Corrected from tempfile.ktemp()
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
    print("DEBUG: apply_comparison_filters called with:")
    print(f" df_raw type: {type(df_raw)}")
    print(f" comparison_config type: {type(comparison_config)}")
    print(f" comparison_mode type: {type(comparison_mode)} value: {comparison_mode}")
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
        title = f"Total Hours by Project in {months[0]} {years[0]}"
        return df_comparison, title

    elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        if len(years) != 1 or len(selected_projects) < 2:
            return pd.DataFrame(), "Vui lòng chọn MỘT năm và ít nhất HAI dự án cho chế độ này."
        df_comparison = df_filtered.groupby('Project name')['Hours'].sum().reset_index()
        df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
        title = f"Total Hours by Project in {years[0]}"
        return df_comparison, title

    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
        if len(selected_projects) != 1:
            return pd.DataFrame(), "Vui lòng chọn CHỈ MỘT dự án cho chế độ này."
        
        project_name = selected_projects[0]
        
        if years and len(years) > 1 and not months: # So sánh qua nhiều năm
            df_comparison = df_filtered.groupby('Year')['Hours'].sum().reset_index()
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            title = f"Total Hours for {project_name} Across Years ({', '.join(map(str, years))})"
            return df_comparison, title
        elif len(years) == 1 and months: # So sánh qua nhiều tháng trong một năm
            df_comparison = df_filtered.groupby('MonthName')['Hours'].sum().reindex(
                ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
            ).dropna().reset_index()
            df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            title = f"Total Hours for {project_name} in {years[0]} by Month"
            return df_comparison, title
        else:
            return pd.DataFrame(), "Vui lòng chọn ít nhất một năm (hoặc nhiều năm) hoặc một năm và nhiều tháng để so sánh."
    
    return pd.DataFrame(), "Chế độ so sánh không hợp lệ hoặc thiếu tiêu chí."


def export_comparison_report(df_comparison, comparison_config, output_file_path, title):
    """Xuất báo cáo so sánh ra file Excel."""
    comparison_mode = comparison_config.get('comparison_mode', 'N/A')
    try:
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            df_comparison.to_excel(writer, sheet_name='Comparison Summary', index=False)

        wb = load_workbook(output_file_path)
        ws = wb['Comparison Summary']

        if not df_comparison.empty:
            chart = LineChart() if comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm" else BarChart()
            
            # Xử lý các trường hợp trục X và Y khác nhau
            if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "So Sánh Dự Án Trong Một Năm", "Compare Projects in a Month", "Compare Projects in a Year"]:
                data_ref = Reference(ws, min_col=2, min_row=1, max_col=2, max_row=ws.max_row)
                cats_ref = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
                chart.add_data(data_ref, titles_from_data=True)
                chart.set_categories(cats_ref)
                chart.x_axis.title = "Project"
                chart.y_axis.title = "Hours"
            elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
                data_ref = Reference(ws, min_col=2, min_row=1, max_col=2, max_row=ws.max_row)
                cats_ref_col = 1
                if 'Year' in df_comparison.columns:
                    cats_ref_col = df_comparison.columns.get_loc('Year') + 1
                    chart.x_axis.title = "Year"
                elif 'MonthName' in df_comparison.columns:
                    cats_ref_col = df_comparison.columns.get_loc('MonthName') + 1
                    chart.x_axis.title = "Month"
                
                cats_ref = Reference(ws, min_col=cats_ref_col, min_row=2, max_row=ws.max_row)
                chart.add_data(data_ref, titles_from_data=True)
                chart.set_categories(cats_ref)
                chart.y_axis.title = "Hours"

            chart.title = title
            ws.add_chart(chart, "D2")

        # Thêm sheet cấu hình cho báo cáo so sánh
        ws_config = wb.create_sheet("Comparison_Config_Info")
        ws_config['A1'], ws_config['B1'] = "Comparison Mode", comparison_mode
        ws_config['A2'], ws_config['B2'] = "Years Selected", ', '.join(map(str, comparison_config.get('years', []))) if comparison_config.get('years') else "N/A"
        ws_config['A3'], ws_config['B3'] = "Months Selected", ', '.join(comparison_config.get('months', [])) if comparison_config.get('months') else "All (if years selected)"
        ws_config['A4'], ws_config['B4'] = "Projects Selected", ', '.join(comparison_config.get('selected_projects', [])) if comparison_config.get('selected_projects') else "No projects selected"

        # Remove template sheets if they exist
        for sheet_name in ['Raw Data', 'Config_Year_Mode', 'Config_Project_Filter']:
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]

        wb.save(output_file_path)
        return True
    except Exception as e:
        print(f"Lỗi khi xuất báo cáo so sánh: {e}")
        return False

def export_comparison_pdf_report(df_comparison, comparison_config, pdf_file_path, title, logo_path):
    """Xuất báo cáo so sánh PDF với biểu đồ."""
    today_str = datetime.datetime.today().strftime("%Y-%m-%d")
    tmp_dir = tempfile.mkdtemp()
    charts_for_pdf = []
    comparison_mode = comparison_config.get('comparison_mode', 'N/A')

    config_info = {
        "Chế độ so sánh": comparison_mode,
        "Năm": ', '.join(map(str, comparison_config.get('years', []))) if comparison_config.get('years') else "N/A",
        "Tháng": ', '.join(comparison_config.get('months', [])) if comparison_config.get('months') else "Tất cả (nếu chọn năm)",
        "Dự án được chọn": ', '.join(comparison_config.get('selected_projects', [])) if comparison_config.get('selected_projects') else "Không có dự án nào được chọn"
    }

    try:
        if not df_comparison.empty:
            fig, ax = plt.subplots(figsize=(12, 7))
            
            if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "So Sánh Dự Án Trong Một Năm", "Compare Projects in a Month", "Compare Projects in a Year"]:
                df_comparison.plot(kind='bar', x='Project name', y='Total Hours', ax=ax, color='teal')
                ax.set_xlabel("Project Name")
                ax.set_ylabel("Total Hours")
                plt.xticks(rotation=45, ha='right')
            elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
                if 'Year' in df_comparison.columns:
                    df_comparison.plot(kind='line', x='Year', y='Total Hours', ax=ax, marker='o', color='purple')
                    ax.set_xlabel("Year")
                elif 'MonthName' in df_comparison.columns:
                    # Đảm bảo thứ tự tháng đúng
                    month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
                    df_comparison['MonthName'] = pd.Categorical(df_comparison['MonthName'], categories=month_order, ordered=True)
                    df_comparison = df_comparison.sort_values('MonthName')
                    df_comparison.plot(kind='line', x='MonthName', y='Total Hours', ax=ax, marker='o', color='purple')
                    ax.set_xlabel("Month")
                    plt.xticks(rotation=45, ha='right')

                ax.set_ylabel("Total Hours")
                ax.grid(True, linestyle='--', alpha=0.6)
            
            ax.set_title(title, fontsize=12, pad=20)
            plt.tight_layout()
            
            chart_img_path = os.path.join(tmp_dir, "comparison_chart.png")
            fig.savefig(chart_img_path, dpi=200)
            plt.close(fig)
            charts_for_pdf.append((chart_img_path, title, ''))
        
        if not charts_for_pdf:
            print("Cảnh báo: Không có biểu đồ nào được tạo để đưa vào PDF. PDF có thể trống.")
            # Tạo một PDF trống với thông báo nếu không có biểu đồ nào
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font('helvetica', 'B', 16)
            pdf.cell(0, 10, "BÁO CÁO SO SÁNH", ln=True, align='C')
            pdf.set_font("helvetica", '', 12)
            pdf.cell(0, 10, f"Ngày tạo: {datetime.datetime.today().strftime('%Y-%m-%d')}", ln=True, align='C')
            pdf.ln(10)
            pdf.set_font("helvetica", '', 11)
            for key, value in config_info.items():
                pdf.cell(0, 7, f"{key}: {value}", ln=True, align='C')
            pdf.cell(0, 10, "Không có biểu đồ nào được tạo cho báo cáo này.", ln=True, align='C')
            pdf.output(pdf_file_path, "F")
            return True
            
        create_pdf_from_charts_comparison(charts_for_pdf, pdf_file_path, "TRIAC TIME REPORT - COMPARISON", config_info, logo_path)
        return True
    except Exception as e:
        print(f"Lỗi khi tạo báo cáo PDF so sánh: {e}")
        return False
    finally:
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)

# Hàm mới để tạo PDF cho báo cáo so sánh, tương tự hàm gốc nhưng có thể tùy biến hơn
def create_pdf_from_charts_comparison(charts_data, output_path, title, config_info, logo_path_inner):
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
    pdf.cell(0, 10, f"Ngày tạo: {datetime.datetime.today().strftime('%Y-%m-%d')}", ln=True, align='C')
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
                pdf.cell(0, 10, f"Dự án: {page_project_name}", ln=True, align='C')
            pdf.cell(0, 10, chart_title, ln=True, align='C')
            pdf.image(img_path, x=10, y=45, w=190)

    pdf.output(output_path, "F")
    print(f"DEBUG: PDF báo cáo so sánh được tạo tại {output_path}")

def generate_reports_on_demand(
    selected_mode,
    selected_year,
    selected_months,
    selected_project_names_standard,
    comparison_config_years,
    comparison_config_months,
    comparison_config_projects,
    comparison_report_mode,
    export_excel_standard,
    export_pdf_standard,
    export_excel_comparison,
    export_pdf_comparison
):
    path_dict = setup_paths()
    template_file = path_dict['template_file']
    output_file_path = path_dict['output_file']
    pdf_report_path = path_dict['pdf_report']
    comparison_output_file_path = path_dict['comparison_output_file']
    comparison_pdf_report_path = path_dict['comparison_pdf_report']
    logo_path = path_dict['logo_path'] # Lấy đường dẫn logo

    df_raw = load_raw_data(template_file)
    if df_raw.empty:
        return {"status": "error", "message": "Không thể tải dữ liệu thô từ file template. Vui lòng kiểm tra file."}

    # =========================================================================
    # Xử lý Báo cáo tiêu chuẩn
    # =========================================================================
    if selected_mode:
        standard_report_config = {
            'mode': selected_mode,
            'year': selected_year,
            'months': selected_months,
            'project_filter_df': pd.DataFrame({'Project Name': selected_project_names_standard, 'Include': ['yes'] * len(selected_project_names_standard)})
        }
        df_standard_filtered = apply_filters(df_raw, standard_report_config)

        if not df_standard_filtered.empty:
            if export_excel_standard:
                if export_report(df_standard_filtered, standard_report_config, output_file_path):
                    print(f"Báo cáo Excel tiêu chuẩn đã tạo: {output_file_path}")
                else:
                    print("Lỗi khi tạo báo cáo Excel tiêu chuẩn.")
            
            if export_pdf_standard:
                if export_pdf_report(df_standard_filtered, standard_report_config, pdf_report_path, logo_path):
                    print(f"Báo cáo PDF tiêu chuẩn đã tạo: {pdf_report_path}")
                else:
                    print("Lỗi khi tạo báo cáo PDF tiêu chuẩn.")
        else:
            print("Không có dữ liệu sau khi lọc cho báo cáo tiêu chuẩn.")

    # =========================================================================
    # Xử lý Báo cáo so sánh
    # =========================================================================
    if comparison_report_mode:
        comparison_config = {
            'comparison_mode': comparison_report_mode,
            'years': comparison_config_years,
            'months': comparison_config_months,
            'selected_projects': comparison_config_projects
        }
        df_comparison_summary, message = apply_comparison_filters(df_raw, comparison_config, comparison_report_mode)

        if not df_comparison_summary.empty:
            chart_title_detail = ""
            if comparison_report_mode == "So Sánh Dự Án Trong Một Tháng":
                chart_title_detail = f"Total Hours by Project in {comparison_config_months[0]} {comparison_config_years[0]}"
            elif comparison_report_mode == "So Sánh Dự Án Trong Một Năm":
                chart_title_detail = f"Total Hours by Project in {comparison_config_years[0]}"
            elif comparison_report_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
                if comparison_config_years and len(comparison_config_years) > 1:
                    chart_title_detail = f"Total Hours for {comparison_config_projects[0]} Across Years ({', '.join(map(str, comparison_config_years))})"
                elif comparison_config_years and comparison_config_months:
                    chart_title_detail = f"Total Hours for {comparison_config_projects[0]} in {comparison_config_years[0]} by Month"

            if export_excel_comparison:
                if export_comparison_report(df_comparison_summary, comparison_config, comparison_output_file_path, chart_title_detail):
                    print(f"Báo cáo Excel so sánh đã tạo: {comparison_output_file_path}")
                else:
                    print("Lỗi khi tạo báo cáo Excel so sánh.")

            if export_pdf_comparison:
                if export_comparison_pdf_report(df_comparison_summary, comparison_config, comparison_pdf_report_path, chart_title_detail, logo_path):
                    print(f"Báo cáo PDF so sánh đã tạo: {comparison_pdf_report_path}")
                else:
                    print("Lỗi khi tạo báo cáo PDF so sánh.")
        else:
            print(f"Không có dữ liệu sau khi lọc cho báo cáo so sánh: {message}")

    return {"status": "success", "message": "Tiến trình tạo báo cáo đã hoàn tất. Vui lòng kiểm tra console để biết chi tiết."}


# Cấu hình ví dụ để chạy thử hàm generate_reports_on_demand
if __name__ == "__main__":
    # --- Cấu hình cho Báo cáo tiêu chuẩn ---
    standard_report_mode = "month" # Có thể là "year", "month", "week"
    standard_report_year = 2023
    standard_report_months = [] # Ví dụ: ['January', 'February'], để trống nếu muốn tất cả các tháng
    standard_report_projects = ["Project Alpha", "Project Beta"] # Thay thế bằng tên dự án của bạn

    # --- Cấu hình cho Báo cáo So sánh ---
    # comparison_report_mode = "So Sánh Dự Án Trong Một Tháng" # Có thể là:
    #   "So Sánh Dự Án Trong Một Tháng"
    #   "So Sánh Dự Án Trong Một Năm"
    #   "So Sánh Một Dự Án Qua Các Tháng/Năm"
    comparison_report_mode = "So Sánh Một Dự Án Qua Các Tháng/Năm" 
    
    comparison_years = [2022, 2023] # Ví dụ cho "So Sánh Một Dự Án Qua Các Tháng/Năm"
    comparison_months = [] # Để trống nếu so sánh theo năm, hoặc ['January', 'February'] nếu so sánh tháng trong một năm cụ thể.
    comparison_projects = ["Project Alpha"] # Ví dụ: ["Project Alpha", "Project Beta"] for "So Sánh Dự Án Trong Một Tháng/Năm"
                                          # Hoặc ["Project Alpha"] for "So Sánh Một Dự Án Qua Các Tháng/Năm"

    # Gọi hàm để tạo báo cáo
    generate_reports_on_demand(
        selected_mode=standard_report_mode,
        selected_year=standard_report_year,
        selected_months=standard_report_months,
        selected_project_names_standard=standard_report_projects,
        comparison_config_years=comparison_years,
        comparison_config_months=comparison_months,
        comparison_config_projects=comparison_projects,
        comparison_report_mode=comparison_report_mode,
        export_excel_standard=False,
        export_pdf_standard=False,
        export_excel_comparison=True,
        export_pdf_comparison=True
    )
