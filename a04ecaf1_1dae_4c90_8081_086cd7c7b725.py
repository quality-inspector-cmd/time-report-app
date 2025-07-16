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
import shutil # Import shutil để xóa thư mục tạm thời

# Hàm hỗ trợ làm sạch tên file
def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)

def setup_paths():
    today = datetime.datetime.today().strftime('%Y%m%d')
    return {
        'template_file': "Time_report.xlsm",
        'output_file': f"Time_report_Standard_{today}.xlsx", # Đổi tên để phân biệt
        'pdf_report': f"Time_report_Standard_{today}.pdf",    # Đổi tên để phân biệt
        'comparison_output_file': f"Time_report_Comparison_{today}.xlsx", # Đường dẫn mới cho Excel so sánh
        'comparison_pdf_report': f"Time_report_Comparison_{today}.pdf"    # Đường dẫn mới cho PDF so sánh
    }

def read_configs(path_dict):
    year_mode_df = pd.read_excel(path_dict['template_file'], sheet_name='Config_Year_Mode', engine='openpyxl')
    project_filter_df = pd.read_excel(path_dict['template_file'], sheet_name='Config_Project_Filter', engine='openpyxl')

    mode = str(year_mode_df.loc[year_mode_df['Key'].str.lower() == 'mode', 'Value'].values[0]).strip().lower()
    year_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'year', 'Value']
    year = int(year_row.values[0]) if not year_row.empty and pd.notna(year_row.values[0]) else None
    months_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'months', 'Value']
    months = [m.strip().capitalize() for m in str(months_row.values[0]).split(',')] if not months_row.empty else []

    return {
        'mode': mode,
        'year': year,
        'months': months,
        'project_filter_df': project_filter_df
    }

def load_raw_data(path_dict):
    df = pd.read_excel(path_dict['template_file'], sheet_name='Raw Data', engine='openpyxl')
    df.rename(columns={'Hou': 'Hours', 'Team member': 'Employee'}, inplace=True)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df['Year'] = df['Date'].dt.year
    df['MonthName'] = df['Date'].dt.month_name()
    df['Week'] = df['Date'].dt.isocalendar().week
    return df

def apply_filters(df, config):
    df_filtered = df.copy()
    if 'years' in config and config['years']: # Dành cho multiselect years
        df_filtered = df_filtered[df_filtered['Year'].isin(config['years'])]
    elif 'year' in config and config['year']: # Dành cho single select year
        df_filtered = df_filtered[df_filtered['Year'] == config['year']]

    if config['months']:
        df_filtered = df_filtered[df_filtered['MonthName'].isin(config['months'])]

    # Lọc dự án dựa trên project_filter_df đã được cấu hình trong Streamlit app
    # (chỉ bao gồm các dự án được chọn và có 'Include' là 'yes')
    if not config['project_filter_df'].empty:
        df_filtered = df_filtered.merge(
            config['project_filter_df'],
            how='inner',
            left_on='Project name',
            right_on='Project Name'
        )
    return df_filtered

def export_report(df, config, path_dict):
    mode = config['mode']
    if mode == 'year':
        summary = df.groupby(['Year', 'Project name'])['Hours'].sum().reset_index()
    elif mode == 'month':
        summary = df.groupby(['Year', 'MonthName', 'Project name'])['Hours'].sum().reset_index()
    else: # week mode
        summary = df.groupby(['Year', 'Week', 'Project name'])['Hours'].sum().reset_index()

    with pd.ExcelWriter(path_dict['output_file'], engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='Summary', index=False)
        # df.to_excel(writer, sheet_name='Raw_Data', index=False)  # Optional: can skip writing raw

    wb = load_workbook(path_dict['output_file'])

    # Add chart to Summary
    ws = wb['Summary']
    max_row = ws.max_row
    
    # Corrected data_col and cats_col logic
    if mode == 'year':
        data_col = summary.columns.get_loc('Hours') + 1 # +1 for 1-based indexing
        cats_col = summary.columns.get_loc('Project name') + 1
    elif mode == 'month':
        data_col = summary.columns.get_loc('Hours') + 1
        cats_col = summary.columns.get_loc('Project name') + 1 # Or MonthName if you want to chart by month
    else: # week
        data_col = summary.columns.get_loc('Hours') + 1
        cats_col = summary.columns.get_loc('Project name') + 1 # Or Week if you want to chart by week


    data_ref = Reference(ws, min_col=data_col, min_row=1, max_row=max_row)
    # Correctly set category reference. It should exclude the header if titles_from_data is True for data_ref
    cats_ref = Reference(ws, min_col=cats_col, min_row=2, max_row=max_row)

    chart = BarChart()
    chart.title = f"Total Hours by Project ({mode})"
    chart.x_axis.title = "Project"
    chart.y_axis.title = "Hours"
    chart.add_data(data_ref, titles_from_data=True) # Assuming first row of data_ref is the title 'Hours'
    chart.set_categories(cats_ref)
    ws.add_chart(chart, "F2")

    # Create project sheets
    for project in df['Project name'].unique():
        df_proj = df[df['Project name'] == project]
        ws_proj = wb.create_sheet(title=project[:31]) # Truncate project name for sheet title

        # Task summary
        summary_task = df_proj.groupby('Task')['Hours'].sum().reset_index().sort_values('Hours', ascending=False)
        
        # Write headers
        ws_proj.append(['Task', 'Hours'])
        # Write data
        for r_idx, row_data in enumerate(dataframe_to_rows(summary_task, index=False, header=False)):
            ws_proj.append(row_data)

        # Chart for Task
        chart = BarChart()
        chart.title = f"{project} - Hours by Task"
        chart.x_axis.title = "Task"
        chart.y_axis.title = "Hours"
        task_len = len(summary_task)
        
        # References for Task chart: data starts from row 1, categories from row 2
        data_ref = Reference(ws_proj, min_col=2, min_row=1, max_row=task_len + 1) # Hours column
        cats_ref = Reference(ws_proj, min_col=1, min_row=2, max_row=task_len + 1) # Task column
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        ws_proj.add_chart(chart, f"E1")

        # Add raw data below
        start_row = task_len + 5
        # Ensure that dataframe_to_rows correctly populates the worksheet
        for r_idx, r in enumerate(dataframe_to_rows(df_proj, index=False, header=True)):
            for c_idx, cell_val in enumerate(r):
                ws_proj.cell(row=start_row + r_idx, column=c_idx + 1, value=cell_val)
        
    # Config info
    ws_config = wb.create_sheet("Config_Info")
    ws_config['A1'], ws_config['B1'] = "Mode", config['mode']
    ws_config['A2'], ws_config['B2'] = "Years", ', '.join(map(str, config.get('years', [config.get('year')])))
    ws_config['A3'], ws_config['B3'] = "Months", ', '.join(config['months']) if config['months'] else "All"
    ws_config['A4'], ws_config['B4'] = "Projects", ', '.join(config['project_filter_df']['Project Name'])

    # Optionally remove 'Raw_Data' if exists
    if 'Raw_Data' in wb.sheetnames:
        del wb['Raw_Data']

    wb.save(path_dict['output_file'])

def export_pdf_report(df, config, path_dict):
    today_str = datetime.datetime.today().strftime("%Y-%m-%d")

    def create_pdf_from_charts(charts_data, output_path, title, config_info, logo_path="triac_logo.png"):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)

        # COVER PAGE
        pdf.add_page()
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=10, y=10, w=30)
        pdf.set_font("Arial", 'B', 16)
        pdf.ln(40)
        pdf.cell(0, 10, title, ln=True, align='C')
        pdf.set_font("Arial", '', 12)
        pdf.ln(5)
        pdf.cell(0, 10, f"Generated on: {today_str}", ln=True, align='C')
        pdf.ln(10)
        pdf.set_font("Arial", '', 11)
        for key, value in config_info.items():
            pdf.cell(0, 7, f"{key}: {value}", ln=True, align='C')

        # Charts pages
        for img_path, chart_title, page_project_name in charts_data:
            if img_path and os.path.exists(img_path):
                pdf.add_page()
                if os.path.exists(logo_path):
                    pdf.image(logo_path, x=10, y=8, w=25)
                pdf.set_font("Arial", 'B', 11)
                pdf.set_y(35)
                if page_project_name:
                    pdf.cell(0, 10, f"Project: {page_project_name}", ln=True, align='C')
                pdf.cell(0, 10, chart_title, ln=True, align='C')
                pdf.image(img_path, x=10, y=45, w=190) # Adjust w/h as needed

        pdf.output(output_path, "F")
        print(f"DEBUG: PDF report generated at {output_path}")

    tmp_dir = tempfile.mkdtemp()
    charts_for_pdf = []

    try: # Thêm khối try-finally để đảm bảo dọn dẹp thư mục tạm thời
        projects = df['Project name'].unique()

        config_info = {
            "Mode": config['mode'].capitalize(),
            "Years": ', '.join(map(str, config.get('years', []))) or str(config.get('year')),
            "Months": ', '.join(config.get('months', [])) or "All",
            "Projects Included": ', '.join(config['project_filter_df']['Project Name'])
        }

        # Generate Charts per Project (Standard Report)
        for project in projects:
            safe_project = sanitize_filename(project)
            df_proj = df[df['Project name'] == project]

            # Workcentre Chart
            fig, ax = plt.subplots(figsize=(8, 4))
            df_proj.groupby('Workcentre')['Hours'].sum().sort_values().plot(kind='barh', color='skyblue', ax=ax)
            ax.set_title(f"{project} - Hours by Workcentre", fontsize=10)
            wc_img_path = os.path.join(tmp_dir, f"{safe_project}_wc.png")
            plt.tight_layout()
            fig.savefig(wc_img_path, dpi=150)
            plt.close(fig)
            charts_for_pdf.append((wc_img_path, f"{project} - Hours by Workcentre", project))

            # Task Chart
            if 'Task' in df_proj.columns and not df_proj['Task'].empty: # Check if 'Task' column exists and is not empty
                fig, ax = plt.subplots(figsize=(8, 4))
                df_proj.groupby('Task')['Hours'].sum().sort_values().plot(kind='barh', color='lightgreen', ax=ax)
                ax.set_title(f"{project} - Hours by Task", fontsize=10)
                task_img_path = os.path.join(tmp_dir, f"{safe_project}_task.png")
                plt.tight_layout()
                fig.savefig(task_img_path, dpi=150)
                plt.close(fig)
                charts_for_pdf.append((task_img_path, f"{project} - Hours by Task", project))

        create_pdf_from_charts(charts_for_pdf, path_dict['pdf_report'], "TRIAC TIME REPORT - STANDARD", config_info)

    finally:
        # Dọn dẹp thư mục tạm thời
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)


# =========================================================================
# CÁC HÀM MỚI CHO BÁO CÁO SO SÁNH
# =========================================================================

def apply_comparison_filters(df_raw, comparison_config, comparison_mode):
    """
    Lọc và chuẩn bị dữ liệu cho báo cáo so sánh dựa trên comparison_mode.
    Trả về DataFrame đã được tổng hợp/pivot và một tiêu đề cho báo cáo.
    """
    years = comparison_config['years']
    months = comparison_config['months']
    selected_projects = comparison_config['selected_projects']

    df_filtered = df_raw[
        (df_raw['Year'].isin(years)) &
        (df_raw['MonthName'].isin(months)) &
        (df_raw['Project name'].isin(selected_projects))
    ].copy()

    if df_filtered.empty:
        return pd.DataFrame(), f"Không có dữ liệu cho chế độ so sánh: {comparison_mode}"

    if comparison_mode == "So Sánh Dự Án Trong Một Tháng" or comparison_mode == "Compare Projects in a Month":
        # Tổng hợp giờ theo Project cho một tháng/năm cụ thể
        # Yêu cầu 1 năm, 1 tháng, nhiều dự án
        if len(years) != 1 or len(months) != 1 or len(selected_projects) < 2:
            return pd.DataFrame(), "Cần chọn MỘT năm, MỘT tháng và ít nhất HAI dự án."
        
        df_comparison = df_filtered.groupby('Project name')['Hours'].sum().reset_index()
        df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
        title = f"So sánh giờ giữa các dự án trong {months[0]}, năm {years[0]}"
        return df_comparison, title

    elif comparison_mode == "So Sánh Dự Án Trong Một Năm" or comparison_mode == "Compare Projects in a Year":
        # Tổng hợp giờ theo Project và Month cho một năm
        # Yêu cầu 1 năm, nhiều dự án (tháng là tất cả các tháng được chọn)
        if len(years) != 1 or len(selected_projects) < 2:
            return pd.DataFrame(), "Cần chọn MỘT năm và ít nhất HAI dự án."
        
        # Ensure all selected months are considered for comparison across projects in a year
        df_comparison = df_filtered.groupby(['Project name', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
        # Sắp xếp cột tháng theo thứ tự tự nhiên (ví dụ: January, February...)
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        existing_months = [m for m in month_order if m in df_comparison.columns]
        df_comparison = df_comparison[existing_months]

        df_comparison.loc['Total'] = df_comparison.sum() # Thêm dòng tổng
        df_comparison = df_comparison.reset_index().rename(columns={'index': 'Project Name'})
        
        title = f"So sánh giờ giữa các dự án trong năm {years[0]} (theo tháng)"
        return df_comparison, title

    elif comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm" or comparison_mode == "Compare One Project Over Time (Months/Years)":
        # Tổng hợp giờ cho một Project theo Month và Year
        # Yêu cầu 1 dự án, nhiều tháng (và nhiều năm nếu chọn)
        if len(selected_projects) != 1 or len(months) < 2: # Ít nhất 2 tháng để so sánh
            return pd.DataFrame(), "Cần chọn MỘT dự án và ít nhất HAI tháng để so sánh."
        
        df_comparison = df_filtered.groupby(['Year', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
        
        # Sắp xếp cột tháng theo thứ tự tự nhiên
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        existing_months = [m for m in month_order if m in df_comparison.columns]
        df_comparison = df_comparison[existing_months]

        df_comparison.loc['Total'] = df_comparison.sum() # Thêm dòng tổng
        df_comparison = df_comparison.reset_index().rename(columns={'index': 'Year'})

        title = f"So sánh giờ của dự án {selected_projects[0]} qua các tháng"
        return df_comparison, title
    
    return pd.DataFrame(), "Chế độ so sánh không hợp lệ."

def export_comparison_report(df_comparison, comparison_config, path_dict, comparison_mode):
    """
    Xuất báo cáo so sánh ra file Excel.
    Sử dụng path_dict['comparison_output_file'] để lưu.
    """
    output_file = path_dict['comparison_output_file']
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_comparison.to_excel(writer, sheet_name='Comparison Report', index=False) # index=False nếu cột đầu tiên đã là tên cột

        wb = writer.book
        ws = writer.sheets['Comparison Report']

        # Tiêu đề báo cáo
        title_format = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
        ws.merge_cells('A1:D1') # Điều chỉnh phạm vi merge phù hợp
        ws['A1'].value = f"BÁO CÁO SO SÁNH: {comparison_mode}"

        # Thông tin cấu hình
        info_row = 2
        ws.cell(row=info_row, column=1, value="Năm:").font = wb.add_format({'bold': True})
        ws.cell(row=info_row, column=2, value=', '.join(map(str, comparison_config['years'])))
        info_row += 1
        ws.cell(row=info_row, column=1, value="Tháng:").font = wb.add_format({'bold': True})
        ws.cell(row=info_row, column=2, value=', '.join(comparison_config['months']))
        info_row += 1
        ws.cell(row=info_row, column=1, value="Dự án:").font = wb.add_format({'bold': True})
        ws.cell(row=info_row, column=2, value=', '.join(comparison_config['selected_projects']))

        # Đặt DataFrame sau thông tin cấu hình
        start_row_data = info_row + 2
        for r_idx, r in enumerate(dataframe_to_rows(df_comparison, index=False, header=True)):
            for c_idx, cell_val in enumerate(r):
                ws.cell(row=start_row_data + r_idx, column=c_idx + 1, value=cell_val)
        
        # Thêm biểu đồ so sánh (ví dụ)
        # Kiểm tra xem DataFrame có cột 'Total Hours' (cho mode 1) hoặc các cột tháng (cho mode 2, 3)
        if not df_comparison.empty:
            chart = BarChart()
            chart.y_axis.title = "Giờ"
            
            if comparison_mode == "So Sánh Dự Án Trong Một Tháng" or comparison_mode == "Compare Projects in a Month":
                # Data for 'Total Hours' (col 2), Categories for 'Project Name' (col 1)
                data_col_idx = df_comparison.columns.get_loc('Total Hours') + 1
                cat_col_idx = df_comparison.columns.get_loc('Project name') + 1
                chart.title = "So sánh giờ theo Dự án"
                chart.x_axis.title = "Dự án"
                
                data_ref = Reference(ws, min_col=data_col_idx, min_row=start_row_data, max_row=start_row_data + len(df_comparison) -1)
                cats_ref = Reference(ws, min_col=cat_col_idx, min_row=start_row_data + 1, max_row=start_row_data + len(df_comparison))
                chart.add_data(data_ref, titles_from_data=False) # Set to False if you explicitly provide titles or titles are not in the first row of data_ref
                chart.set_categories(cats_ref)
            
            elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"] :
                # Data for months, Categories for Project Name
                # Loại bỏ cột 'Project Name' và 'Total' khỏi data_cols_indices nếu có
                data_cols_indices = [df_comparison.columns.get_loc(col) + 1 for col in df_comparison.columns if col not in ['Project Name', 'Total']]
                cat_col_idx = df_comparison.columns.get_loc('Project Name') + 1
                chart.title = "So sánh giờ theo Dự án và Tháng"
                chart.x_axis.title = "Dự án"

                # For unstacked data, we need to iterate over series
                for col_idx in data_cols_indices:
                    series_ref = Reference(ws, min_col=col_idx, min_row=start_row_data, max_row=start_row_data + len(df_comparison) -1)
                    chart.add_data(series_ref, titles_from_data=True)
                
                cats_ref = Reference(ws, min_col=cat_col_idx, min_row=start_row_data + 1, max_row=start_row_data + len(df_comparison))
                chart.set_categories(cats_ref)

            elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"] :
                # Data for months, Categories for Year
                data_cols_indices = [df_comparison.columns.get_loc(col) + 1 for col in df_comparison.columns if col not in ['Year', 'Total']]
                cat_col_idx = df_comparison.columns.get_loc('Year') + 1
                chart.title = "So sánh giờ của một dự án qua các tháng/năm"
                chart.x_axis.title = "Thời gian (Năm)"

                # For unstacked data, we need to iterate over series
                for col_idx in data_cols_indices:
                    series_ref = Reference(ws, min_col=col_idx, min_row=start_row_data, max_row=start_row_data + len(df_comparison) -1)
                    chart.add_data(series_ref, titles_from_data=True)
                
                cats_ref = Reference(ws, min_col=cat_col_idx, min_row=start_row_data + 1, max_row=start_row_data + len(df_comparison))
                chart.set_categories(cats_ref)


            ws.add_chart(chart, f"A{start_row_data + len(df_comparison) + 2}") # Đặt biểu đồ bên dưới bảng

        wb.save(output_file)
        return True

def export_comparison_pdf_report(df_comparison, comparison_config, path_dict, comparison_mode):
    """
    Xuất báo cáo so sánh ra file PDF.
    Sử dụng path_dict['comparison_pdf_report'] để lưu.
    Sử dụng Matplotlib để tạo biểu đồ và nhúng vào PDF.
    """
    pdf_file = path_dict['comparison_pdf_report']
    tmp_dir = tempfile.mkdtemp()
    charts_for_pdf = []

    def create_pdf_from_charts(charts_data, output_path, title, config_info, logo_path="triac_logo.png"):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_font('DejaVuSans', '', 'DejaVuSansCondensed.ttf', uni=True) # Ensure font supports Vietnamese

        # COVER PAGE
        pdf.add_page()
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=10, y=10, w=30)
        pdf.set_font("DejaVuSans", 'B', 16)
        pdf.ln(40)
        pdf.cell(0, 10, title, ln=True, align='C')
        pdf.set_font("DejaVuSans", '', 12)
        pdf.ln(5)
        pdf.cell(0, 10, f"Generated on: {datetime.datetime.today().strftime('%Y-%m-%d')}", ln=True, align='C')
        pdf.ln(10)
        pdf.set_font("DejaVuSans", '', 11)
        for key, value in config_info.items():
            pdf.cell(0, 7, f"{key}: {value}", ln=True, align='C')

        # Charts pages
        for img_path, chart_title, page_project_name in charts_data:
            if img_path and os.path.exists(img_path):
                pdf.add_page()
                if os.path.exists(logo_path):
                    pdf.image(logo_path, x=10, y=8, w=25)
                pdf.set_font("DejaVuSans", 'B', 11)
                pdf.set_y(35)
                if page_project_name:
                    pdf.cell(0, 10, f"Project: {page_project_name}", ln=True, align='C')
                pdf.cell(0, 10, chart_title, ln=True, align='C')
                pdf.image(img_path, x=10, y=45, w=190) # Adjust w/h as needed

        pdf.output(output_path, "F")
        print(f"DEBUG: PDF report generated at {output_path}")

    def create_comparison_chart(df, mode, title, x_label, y_label, img_path):
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Tạo bản sao DataFrame để tránh SettingWithCopyWarning khi set_index
        df_plot = df.copy() 

        # Đảm bảo trục y luôn bắt đầu từ 0
        ax.set_ylim(bottom=0)

        if mode == "So Sánh Dự Án Trong Một Tháng" or mode == "Compare Projects in a Month":
            # Bar chart: Projects vs Total Hours
            ax.bar(df_plot['Project name'], df_plot['Total Hours'], color='skyblue')
            ax.set_xticks(df_plot['Project name'])
            ax.tick_params(axis='x', rotation=45, ha='right')
        elif mode == "So Sánh Dự Án Trong Một Năm" or mode == "Compare Projects in a Year":
            # Grouped Bar Chart: Projects vs Months
            # Loại bỏ dòng 'Total' nếu có để không vẽ vào biểu đồ
            if 'Project Name' in df_plot.columns and 'Total' in df_plot['Project Name'].values:
                df_plot = df_plot[df_plot['Project Name'] != 'Total']
            
            # Đảm bảo các cột tháng là số trước khi vẽ biểu đồ
            month_columns = [col for col in df_plot.columns if col not in ['Project Name']]
            df_plot[month_columns] = df_plot[month_columns].apply(pd.to_numeric, errors='coerce').fillna(0) # Chuyển sang số và điền NaN bằng 0

            df_plot.set_index('Project Name', inplace=True)
            df_plot.plot(kind='bar', ax=ax, figsize=(10,6), colormap='viridis') # Grouped bar chart
            ax.set_xticks(range(len(df_plot.index)))
            ax.set_xticklabels(df_plot.index, rotation=45, ha='right')
            ax.legend(title="Tháng", bbox_to_anchor=(1.05, 1), loc='upper left')
        elif mode == "So Sánh Một Dự Án Qua Các Tháng/Năm" or mode == "Compare One Project Over Time (Months/Years)":
            # Line chart or Bar chart: Year/Month vs Hours
            # Loại bỏ dòng 'Total' nếu có để không vẽ vào biểu đồ
            if 'Year' in df_plot.columns and 'Total' in df_plot['Year'].values:
                df_plot = df_plot[df_plot['Year'] != 'Total']
            
            # Đảm bảo các cột tháng là số trước khi vẽ biểu đồ
            month_columns = [col for col in df_plot.columns if col not in ['Year']]
            df_plot[month_columns] = df_plot[month_columns].apply(pd.to_numeric, errors='coerce').fillna(0) # Chuyển sang số và điền NaN bằng 0

            df_plot.set_index('Year', inplace=True)
            df_plot.plot(kind='bar', ax=ax, figsize=(10,6), colormap='plasma') # Bar chart for months within year
            ax.set_xticks(range(len(df_plot.index)))
            ax.set_xticklabels(df_plot.index, rotation=45, ha='right')
            ax.legend(title="Tháng", bbox_to_anchor=(1.05, 1), loc='upper left')

        ax.set_title(title)
        ax.set_xlabel(x_label)
        ax.set_ylabel(y_label)
        plt.tight_layout()
        fig.savefig(img_path, dpi=200)
        plt.close(fig)
        return img_path

    try:
        # Tạo biểu đồ chính cho báo cáo so sánh
        chart_title = f"Báo cáo so sánh: {comparison_mode}"
        x_label = ""
        y_label = "Tổng số giờ"

        # Định nghĩa x_label dựa trên mode so sánh
        if comparison_mode == "So Sánh Dự Án Trong Một Tháng" or comparison_mode == "Compare Projects in a Month":
            x_label = "Dự án"
        elif comparison_mode == "So Sánh Dự Án Trong Một Năm" or comparison_mode == "Compare Projects in a Year":
            x_label = "Dự án"
        elif comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm" or comparison_mode == "Compare One Project Over Time (Months/Years)":
            x_label = "Thời gian (Tháng/Năm)"
        
        comp_chart_path = os.path.join(tmp_dir, "comparison_chart.png")
        created_chart_path = create_comparison_chart(df_comparison.copy(), comparison_mode, chart_title, x_label, y_label, comp_chart_path)
        if created_chart_path:
            charts_for_pdf.append((created_chart_path, chart_title, None)) # page_project_name is None for main chart

        # Tạo PDF
        config_info = {
            "Chế độ so sánh": comparison_mode,
            "Năm": ', '.join(map(str, comparison_config['years'])),
            "Tháng": ', '.join(comparison_config['months']),
            "Dự án": ', '.join(comparison_config['selected_projects'])
        }
        
        # Sử dụng lại hàm create_pdf_from_charts từ export_pdf_report
        create_pdf_from_charts(charts_for_pdf, pdf_file, "TRIAC TIME REPORT - COMPARISON", config_info)
    finally:
        # Dọn dẹp thư mục tạm thời
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)
    return True
