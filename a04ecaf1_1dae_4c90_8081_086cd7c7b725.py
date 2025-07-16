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
    invalid_chars = re.compile(r'[\\\\/*?[\\]:;|=,<>]')
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
    """Đọc dữ liệu cấu hình từ file template Excel."""
    try:
        xls = pd.ExcelFile(template_file)
        config_year_mode = xls.parse('Config_Year_Mode').set_index('Key')['Value'].to_dict()
        config_project_filter = xls.parse('Config_Project_Filter')
        return {
            'config_year_mode': config_year_mode,
            'config_project_filter': config_project_filter
        }
    except FileNotFoundError:
        print(f"Lỗi: Không tìm thấy file template '{template_file}'. Vui lòng kiểm tra đường dẫn.")
        return None
    except Exception as e:
        print(f"Lỗi khi đọc file cấu hình: {e}")
        return None

def load_raw_data(template_file):
    """Tải dữ liệu thô từ sheet 'Raw Data' của file template Excel."""
    try:
        df_raw = pd.read_excel(template_file, sheet_name='Raw Data')
        # Đảm bảo các cột cần thiết có mặt
        required_columns = ['Year', 'MonthName', 'Project name', 'Workcenter', 'Task', 'Time (Hours)']
        if not all(col in df_raw.columns for col in required_columns):
            print(f"Lỗi: Thiếu một hoặc nhiều cột bắt buộc trong sheet 'Raw Data'. Các cột cần thiết: {required_columns}")
            return None
        
        # Chuyển đổi 'MonthName' sang dạng số nếu cần cho việc sắp xếp
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December']
        df_raw['Month'] = pd.Categorical(df_raw['MonthName'], categories=month_order, ordered=True)
        df_raw = df_raw.sort_values(by=['Year', 'Month'])
        
        return df_raw
    except FileNotFoundError:
        print(f"Lỗi: Không tìm thấy file template '{template_file}'. Vui lòng kiểm tra đường dẫn.")
        return None
    except Exception as e:
        print(f"Lỗi khi tải dữ liệu thô: {e}")
        return None

def apply_filters(df, selected_mode, selected_year, selected_months, selected_project_names):
    """Áp dụng các bộ lọc cho DataFrame."""
    df_filtered = df.copy()

    if selected_year:
        df_filtered = df_filtered[df_filtered['Year'] == selected_year]
    
    if selected_months:
        df_filtered = df_filtered[df_filtered['MonthName'].isin(selected_months)]

    if selected_project_names:
        df_filtered = df_filtered[df_filtered['Project name'].isin(selected_project_names)]

    return df_filtered

def export_report(df_filtered, selected_mode, selected_year, selected_months, selected_project_names, output_path):
    """Xuất báo cáo dưới dạng file Excel."""
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Tạo sheet "Summary"
            summary_data = df_filtered.groupby('Project name')['Time (Hours)'].sum().reset_index()
            summary_data.columns = ['Project Name', 'Total Hours']
            summary_data.to_excel(writer, sheet_name='Summary', index=False)

            # Vẽ biểu đồ và chèn vào Excel
            workbook = writer.book
            summary_sheet = writer.sheets['Summary']

            # Biểu đồ tổng giờ theo dự án
            chart1 = BarChart()
            chart1.type = "col"
            chart1.style = 10
            chart1.title = "Total Hours by Project"
            chart1.y_axis.title = 'Hours'
            chart1.x_axis.title = 'Project Name'
            
            data = Reference(summary_sheet, min_col=2, min_row=1, max_row=len(summary_data) + 1, max_col=2)
            categories = Reference(summary_sheet, min_col=1, min_row=2, max_row=len(summary_data) + 1)
            chart1.add_data(data, titles_from_data=True)
            chart1.set_categories(categories)
            summary_sheet.add_chart(chart1, "D2")

            # Tạo các sheet chi tiết cho từng dự án
            for project_name in selected_project_names:
                df_project = df_filtered[df_filtered['Project name'] == project_name]
                if not df_project.empty:
                    sheet_name = sanitize_filename(project_name) # Làm sạch tên sheet
                    df_project.to_excel(writer, sheet_name=sheet_name, index=False)

                    # Thêm biểu đồ Workcenter vào sheet dự án (nếu có đủ dữ liệu)
                    project_sheet = writer.sheets[sheet_name]
                    workcenter_summary = df_project.groupby('Workcenter')['Time (Hours)'].sum().reset_index()
                    if not workcenter_summary.empty:
                        # Ghi dữ liệu tóm tắt Workcenter vào sheet tạm thời
                        for r_idx, row in enumerate(dataframe_to_rows(workcenter_summary, index=False, header=True), 1):
                            project_sheet.append(row)

                        chart2 = BarChart()
                        chart2.type = "col"
                        chart2.style = 10
                        chart2.title = f"Hours by Workcenter for {project_name}"
                        chart2.y_axis.title = 'Hours'
                        chart2.x_axis.title = 'Workcenter'
                        
                        data = Reference(project_sheet, min_col=workcenter_summary.shape[1]+1, min_row=project_sheet.max_row - len(workcenter_summary), max_row=project_sheet.max_row, max_col=project_summary.shape[1]+1)
                        categories = Reference(project_sheet, min_col=workcenter_summary.shape[1], min_row=project_sheet.max_row - len(workcenter_summary)+1, max_row=project_sheet.max_row)
                        chart2.add_data(data, titles_from_data=True)
                        chart2.set_categories(categories)
                        project_sheet.add_chart(chart2, f"D{workcenter_summary.shape[0] + 5}") # Đặt biểu đồ bên dưới bảng

        return True, f"Báo cáo Excel đã được tạo thành công: {os.path.basename(output_path)}"
    except Exception as e:
        return False, f"Lỗi khi xuất báo cáo Excel: {e}"

def export_pdf_report(df_filtered, selected_mode, selected_year, selected_months, selected_project_names, output_path, logo_path):
    """Xuất báo cáo dưới dạng file PDF."""
    try:
        class PDF(FPDF):
            def header(self):
                if os.path.exists(logo_path):
                    self.image(logo_path, 10, 8, 33)
                self.set_font('Arial', 'B', 15)
                self.cell(80)
                self.cell(30, 10, 'Time Report', 1, 0, 'C')
                self.ln(20)

            def footer(self):
                self.set_y(-15)
                self.set_font('Arial', 'I', 8)
                self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', 0, 0, 'C')

        pdf = PDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        # Tiêu đề báo cáo
        report_title = f"Time Report for Year {selected_year}"
        if selected_months:
            report_title += f" - Months: {', '.join(selected_months)}"
        if selected_project_names:
            report_title += f" - Projects: {', '.join(selected_project_names[:3])}{'...' if len(selected_project_names) > 3 else ''}"
        
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, report_title, 0, 1, 'C')
        pdf.ln(5)

        # Tóm tắt tổng giờ
        total_hours = df_filtered['Time (Hours)'].sum()
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, f"Total Hours: {total_hours:.2f}", 0, 1)
        pdf.ln(5)

        # Bảng tóm tắt theo dự án
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Hours by Project:', 0, 1)
        pdf.ln(2)

        summary_data = df_filtered.groupby('Project name')['Time (Hours)'].sum().reset_index()
        for _, row in summary_data.iterrows():
            pdf.cell(80, 8, row['Project name'], 1)
            pdf.cell(40, 8, f"{row['Time (Hours)']:.2f}", 1, 1)
        pdf.ln(10)

        # Chèn biểu đồ (được lưu tạm thời)
        if not summary_data.empty:
            plt.figure(figsize=(10, 6))
            plt.bar(summary_data['Project name'], summary_data['Total Hours'])
            plt.title('Total Hours by Project')
            plt.xlabel('Project Name')
            plt.ylabel('Hours')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                plt.savefig(tmp_img.name)
                img_path = tmp_img.name
            plt.close() # Đóng biểu đồ matplotlib

            pdf.image(img_path, x=pdf.get_x(), y=pdf.get_y(), w=150)
            pdf.ln(10)
            os.unlink(img_path) # Xóa file tạm thời

        # Chi tiết cho từng dự án
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Details by Project:', 0, 1)
        pdf.ln(2)
        
        for project_name in selected_project_names:
            df_project = df_filtered[df_filtered['Project name'] == project_name]
            if not df_project.empty:
                pdf.add_page() # Trang mới cho mỗi dự án chi tiết
                pdf.set_font('Arial', 'B', 12)
                pdf.cell(0, 10, f"Project: {project_name}", 0, 1)
                pdf.ln(2)

                # Bảng chi tiết Task và Workcenter
                pdf.set_font('Arial', 'B', 10)
                pdf.cell(60, 7, 'Task', 1)
                pdf.cell(60, 7, 'Workcenter', 1)
                pdf.cell(40, 7, 'Hours', 1, 1)
                pdf.set_font('Arial', '', 10)
                for _, row in df_project.iterrows():
                    pdf.cell(60, 7, row['Task'], 1)
                    pdf.cell(60, 7, row['Workcenter'], 1)
                    pdf.cell(40, 7, f"{row['Time (Hours)']:.2f}", 1, 1)
                pdf.ln(5)

        pdf.output(output_path)
        return True, f"Báo cáo PDF đã được tạo thành công: {os.path.basename(output_path)}"
    except Exception as e:
        return False, f"Lỗi khi xuất báo cáo PDF: {e}"

def apply_comparison_filters(df, comparison_config_years, comparison_config_months, comparison_config_projects, comparison_report_mode):
    """Áp dụng các bộ lọc cho báo cáo so sánh."""
    df_filtered = df.copy()

    if comparison_report_mode == "So Sánh Dự Án Trong Một Tháng":
        # Yêu cầu 1 năm, 1 tháng, nhiều dự án
        if comparison_config_years and comparison_config_months and comparison_config_projects:
            df_filtered = df_filtered[
                (df_filtered['Year'] == comparison_config_years[0]) &
                (df_filtered['MonthName'] == comparison_config_months[0]) &
                (df_filtered['Project name'].isin(comparison_config_projects))
            ]
        else:
            return pd.DataFrame() # Trả về DataFrame rỗng nếu thiếu config
    
    elif comparison_report_mode == "So Sánh Dự Án Trong Một Năm":
        # Yêu cầu 1 năm, nhiều dự án (tất cả các tháng)
        if comparison_config_years and comparison_config_projects:
            df_filtered = df_filtered[
                (df_filtered['Year'] == comparison_config_years[0]) &
                (df_filtered['Project name'].isin(comparison_config_projects))
            ]
        else:
            return pd.DataFrame()

    elif comparison_report_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
        # Yêu cầu 1 dự án, và (nhiều tháng trong 1 năm HOẶC nhiều năm)
        if comparison_config_projects:
            df_filtered = df_filtered[df_filtered['Project name'] == comparison_config_projects[0]]
            if comparison_config_months: # So sánh qua các tháng trong 1 năm
                if comparison_config_years:
                     df_filtered = df_filtered[
                        (df_filtered['Year'] == comparison_config_years[0]) &
                        (df_filtered['MonthName'].isin(comparison_config_months))
                    ]
                else:
                    return pd.DataFrame() # Cần năm nếu so sánh theo tháng
            elif comparison_config_years: # So sánh qua các năm (tất cả tháng trong năm đó)
                df_filtered = df_filtered[df_filtered['Year'].isin(comparison_config_years)]
            else:
                return pd.DataFrame() # Cần ít nhất tháng hoặc năm
        else:
            return pd.DataFrame() # Trả về DataFrame rỗng nếu thiếu config

    else:
        return pd.DataFrame() # Chế độ so sánh không hợp lệ

    return df_filtered

def export_comparison_report(df_filtered, comparison_report_mode, output_path):
    """Xuất báo cáo so sánh dưới dạng file Excel."""
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if df_filtered.empty:
                # Tạo một sheet trống nếu không có dữ liệu
                empty_df = pd.DataFrame(["No data for selected criteria"], columns=["Message"])
                empty_df.to_excel(writer, sheet_name="Comparison Report", index=False)
                return True, "Không có dữ liệu cho tiêu chí so sánh đã chọn."

            if comparison_report_mode == "So Sánh Dự Án Trong Một Tháng":
                pivot_table = df_filtered.pivot_table(
                    values='Time (Hours)', 
                    index='Project name', 
                    aggfunc='sum'
                )
                pivot_table.to_excel(writer, sheet_name='Comparison - Projects in Month')
                
                # Thêm biểu đồ
                workbook = writer.book
                sheet = writer.sheets['Comparison - Projects in Month']
                chart = BarChart()
                chart.type = "col"
                chart.style = 10
                chart.title = "Comparison: Projects in a Month"
                chart.y_axis.title = 'Total Hours'
                chart.x_axis.title = 'Project Name'
                data = Reference(sheet, min_col=2, min_row=1, max_row=len(pivot_table)+1, max_col=2)
                categories = Reference(sheet, min_col=1, min_row=2, max_row=len(pivot_table)+1)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(categories)
                sheet.add_chart(chart, "D2")

            elif comparison_report_mode == "So Sánh Dự Án Trong Một Năm":
                pivot_table = df_filtered.pivot_table(
                    values='Time (Hours)', 
                    index='Project name', 
                    aggfunc='sum'
                )
                pivot_table.to_excel(writer, sheet_name='Comparison - Projects in Year')
                
                # Thêm biểu đồ
                workbook = writer.book
                sheet = writer.sheets['Comparison - Projects in Year']
                chart = BarChart()
                chart.type = "col"
                chart.style = 10
                chart.title = "Comparison: Projects in a Year"
                chart.y_axis.title = 'Total Hours'
                chart.x_axis.title = 'Project Name'
                data = Reference(sheet, min_col=2, min_row=1, max_row=len(pivot_table)+1, max_col=2)
                categories = Reference(sheet, min_col=1, min_row=2, max_row=len(pivot_table)+1)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(categories)
                sheet.add_chart(chart, "D2")

            elif comparison_report_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
                # Xác định so sánh theo tháng hay theo năm
                if df_filtered['MonthName'].nunique() > 1 and df_filtered['Year'].nunique() == 1:
                    # So sánh qua các tháng trong một năm cụ thể
                    pivot_table = df_filtered.pivot_table(
                        values='Time (Hours)', 
                        index=['MonthName'], # Sắp xếp theo Month để biểu đồ đúng thứ tự
                        columns='Project name', # Chỉ có 1 dự án nên cột này sẽ là tên dự án
                        aggfunc='sum'
                    ).reindex(pd.CategoricalIndex(['January', 'February', 'March', 'April', 'May', 'June',
                                                  'July', 'August', 'September', 'October', 'November', 'December'], ordered=True))
                    pivot_table = pivot_table.dropna(axis=0, how='all') # Bỏ các tháng không có dữ liệu
                    
                    sheet_name = sanitize_filename(f"Comparison - {df_filtered['Project name'].iloc[0]} by Month")
                    pivot_table.to_excel(writer, sheet_name=sheet_name)

                    # Thêm biểu đồ đường
                    workbook = writer.book
                    sheet = writer.sheets[sheet_name]
                    chart = LineChart()
                    chart.title = f"Comparison: {df_filtered['Project name'].iloc[0]} Hours by Month"
                    chart.y_axis.title = "Hours"
                    chart.x_axis.title = "Month"
                    data = Reference(sheet, min_col=2, min_row=1, max_col=2, max_row=len(pivot_table)+1)
                    categories = Reference(sheet, min_col=1, min_row=2, max_row=len(pivot_table)+1)
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(categories)
                    sheet.add_chart(chart, "D2")

                elif df_filtered['Year'].nunique() > 1:
                    # So sánh qua các năm
                    pivot_table = df_filtered.pivot_table(
                        values='Time (Hours)', 
                        index=['Year'], 
                        columns='Project name', # Chỉ có 1 dự án nên cột này sẽ là tên dự án
                        aggfunc='sum'
                    )
                    sheet_name = sanitize_filename(f"Comparison - {df_filtered['Project name'].iloc[0]} by Year")
                    pivot_table.to_excel(writer, sheet_name=sheet_name)
                    
                    # Thêm biểu đồ đường
                    workbook = writer.book
                    sheet = writer.sheets[sheet_name]
                    chart = LineChart()
                    chart.title = f"Comparison: {df_filtered['Project name'].iloc[0]} Hours by Year"
                    chart.y_axis.title = "Hours"
                    chart.x_axis.title = "Year"
                    data = Reference(sheet, min_col=2, min_row=1, max_col=2, max_row=len(pivot_table)+1)
                    categories = Reference(sheet, min_col=1, min_row=2, max_row=len(pivot_table)+1)
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(categories)
                    sheet.add_chart(chart, "D2")
                else:
                     return False, "Không đủ dữ liệu hoặc cấu hình cho chế độ so sánh này."
            else:
                return False, "Chế độ so sánh không hợp lệ hoặc thiếu dữ liệu."
        
        return True, f"Báo cáo so sánh Excel đã được tạo thành công: {os.path.basename(output_path)}"
    except Exception as e:
        return False, f"Lỗi khi xuất báo cáo so sánh Excel: {e}"

def export_comparison_pdf_report(df_filtered, comparison_report_mode, output_path, logo_path):
    """Xuất báo cáo so sánh dưới dạng file PDF."""
    try:
        class PDF(FPDF):
            def header(self):
                if os.path.exists(logo_path):
                    self.image(logo_path, 10, 8, 33)
                self.set_font('Arial', 'B', 15)
                self.cell(80)
                self.cell(30, 10, 'Comparison Report', 1, 0, 'C')
                self.ln(20)

            def footer(self):
                self.set_y(-15)
                self.set_font('Arial', 'I', 8)
                self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', 0, 0, 'C')

        pdf = PDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        if df_filtered.empty:
            pdf.cell(0, 10, "No data for selected comparison criteria.", 0, 1, 'C')
            pdf.output(output_path)
            return True, "Không có dữ liệu cho tiêu chí so sánh đã chọn."

        report_title = f"Comparison Report - Mode: {comparison_report_mode}"
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, report_title, 0, 1, 'C')
        pdf.ln(5)

        # Tạo biểu đồ và lưu tạm thời
        plt.figure(figsize=(12, 7))
        if comparison_report_mode == "So Sánh Dự Án Trong Một Tháng" or \
           comparison_report_mode == "So Sánh Dự Án Trong Một Năm":
            summary = df_filtered.groupby('Project name')['Time (Hours)'].sum().reset_index()
            plt.bar(summary['Project name'], summary['Time (Hours)'])
            plt.title('Total Hours by Project')
            plt.xlabel('Project Name')
            plt.ylabel('Hours')
            plt.xticks(rotation=45, ha='right')
        elif comparison_report_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
            if df_filtered['MonthName'].nunique() > 1 and df_filtered['Year'].nunique() == 1:
                # So sánh theo tháng
                summary = df_filtered.groupby('MonthName')['Time (Hours)'].sum().reindex(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']).dropna()
                plt.plot(summary.index, summary.values, marker='o')
                plt.title(f"{df_filtered['Project name'].iloc[0]} Hours by Month")
                plt.xlabel('Month')
                plt.ylabel('Hours')
                plt.xticks(rotation=45, ha='right')
            elif df_filtered['Year'].nunique() > 1:
                # So sánh theo năm
                summary = df_filtered.groupby('Year')['Time (Hours)'].sum()
                plt.plot(summary.index.astype(str), summary.values, marker='o')
                plt.title(f"{df_filtered['Project name'].iloc[0]} Hours by Year")
                plt.xlabel('Year')
                plt.ylabel('Hours')
            else:
                return False, "Không đủ dữ liệu cho biểu đồ PDF so sánh."
        else:
            return False, "Chế độ so sánh không hợp lệ cho PDF."

        plt.tight_layout()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
            plt.savefig(tmp_img.name)
            img_path = tmp_img.name
        plt.close() # Đóng biểu đồ matplotlib

        # Chèn biểu đồ vào PDF
        pdf.image(img_path, x=pdf.get_x(), y=pdf.get_y(), w=180)
        pdf.ln(10)
        os.unlink(img_path) # Xóa file tạm thời

        # Thêm bảng dữ liệu chi tiết
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Summary Data:', 0, 1)
        pdf.ln(2)

        # Chuyển dữ liệu thành bảng trong PDF
        # Tùy chỉnh hiển thị bảng tùy thuộc vào chế độ so sánh
        if 'summary' in locals(): # Đảm bảo summary đã được tạo
            data_for_pdf = summary.reset_index()
            # Đặt chiều rộng cột
            col_widths = [pdf.w / (len(data_for_pdf.columns) + 1)] * len(data_for_pdf.columns)
            
            # Tiêu đề bảng
            pdf.set_font('Arial', 'B', 10)
            for i, header in enumerate(data_for_pdf.columns):
                pdf.cell(col_widths[i], 7, str(header), 1, 0, 'C')
            pdf.ln()

            # Dữ liệu bảng
            pdf.set_font('Arial', '', 10)
            for index, row in data_for_pdf.iterrows():
                for i, value in enumerate(row):
                    pdf.cell(col_widths[i], 7, str(value), 1, 0, 'C')
                pdf.ln()

        pdf.output(output_path)
        return True, f"Báo cáo PDF so sánh đã được tạo thành công: {os.path.basename(output_path)}"
    except Exception as e:
        return False, f"Lỗi khi xuất báo cáo PDF so sánh: {e}"


def generate_reports_on_demand(
    df_raw,
    config_data,  # <-- ĐÃ THÊM THAM SỐ NÀY ĐỂ KHẮC PHỤC LỖI
    selected_mode,
    selected_year,
    selected_months,
    selected_project_names_standard,
    comparison_config_years,
    comparison_config_months,
    comparison_config_projects,
    comparison_report_mode,
    path_dict # <-- ĐÃ THÊM THAM SỐ NÀY ĐỂ KHẮC PHỤC LỖI
):
    """
    Hàm tổng hợp để tạo cả báo cáo tiêu chuẩn và báo cáo so sánh dựa trên các tham số.
    Trả về (success, message, excel_file_path, pdf_file_path)
    """
    excel_path = None
    pdf_path = None
    success = False
    message = "Unknown error."

    # Xác định loại báo cáo cần tạo
    is_standard_report = (selected_mode is not None and selected_mode != "")
    is_comparison_report = (comparison_report_mode is not None and comparison_report_mode != "")

    if is_standard_report:
        try:
            # Lấy thông tin cấu hình từ config_data (nếu cần cho việc lọc/xử lý)
            # Hiện tại các bộ lọc được truyền trực tiếp, nhưng config_data có thể hữu ích
            # cho các logic phức tạp hơn hoặc kiểm tra tính hợp lệ
            
            df_filtered = apply_filters(
                df_raw,
                selected_mode,
                selected_year,
                selected_months,
                selected_project_names_standard
            )

            if df_filtered.empty:
                return False, "Không có dữ liệu phù hợp với tiêu chí báo cáo tiêu chuẩn đã chọn.", None, None

            # Tạo báo cáo Excel tiêu chuẩn
            excel_success, excel_message = export_report(
                df_filtered,
                selected_mode,
                selected_year,
                selected_months,
                selected_project_names_standard,
                path_dict['output_file']
            )
            if excel_success:
                excel_path = path_dict['output_file']
            
            # Tạo báo cáo PDF tiêu chuẩn
            pdf_success, pdf_message = export_pdf_report(
                df_filtered,
                selected_mode,
                selected_year,
                selected_months,
                selected_project_names_standard,
                path_dict['pdf_report'],
                path_dict['logo_path']
            )
            if pdf_success:
                pdf_path = path_dict['pdf_report']

            if excel_success or pdf_success:
                success = True
                message = "Báo cáo tiêu chuẩn đã được tạo thành công."
                if excel_path and pdf_path:
                    message += f" (Excel: {os.path.basename(excel_path)}, PDF: {os.path.basename(pdf_path)})"
                elif excel_path:
                    message += f" (Excel: {os.path.basename(excel_path)})"
                elif pdf_path:
                    message += f" (PDF: {os.path.basename(pdf_path)})"
            else:
                success = False
                message = f"Lỗi khi tạo báo cáo tiêu chuẩn: Excel: {excel_message}, PDF: {pdf_message}"

        except Exception as e:
            success = False
            message = f"Lỗi hệ thống khi tạo báo cáo tiêu chuẩn: {e}"
        
        return success, message, excel_path, pdf_path

    elif is_comparison_report:
        try:
            df_filtered_comparison = apply_comparison_filters(
                df_raw,
                comparison_config_years,
                comparison_config_months,
                comparison_config_projects,
                comparison_report_mode
            )

            if df_filtered_comparison.empty:
                return False, "Không có dữ liệu phù hợp với tiêu chí báo cáo so sánh đã chọn.", None, None

            # Tạo báo cáo Excel so sánh
            excel_success, excel_message = export_comparison_report(
                df_filtered_comparison,
                comparison_report_mode,
                path_dict['comparison_output_file']
            )
            if excel_success:
                excel_path = path_dict['comparison_output_file']

            # Tạo báo cáo PDF so sánh
            pdf_success, pdf_message = export_comparison_pdf_report(
                df_filtered_comparison,
                comparison_report_mode,
                path_dict['comparison_pdf_report'],
                path_dict['logo_path']
            )
            if pdf_success:
                pdf_path = path_dict['comparison_pdf_report']
            
            if excel_success or pdf_success:
                success = True
                message = "Báo cáo so sánh đã được tạo thành công."
                if excel_path and pdf_path:
                    message += f" (Excel: {os.path.basename(excel_path)}, PDF: {os.path.basename(pdf_path)})"
                elif excel_path:
                    message += f" (Excel: {os.path.basename(excel_path)})"
                elif pdf_path:
                    message += f" (PDF: {os.path.basename(pdf_path)})"
            else:
                success = False
                message = f"Lỗi khi tạo báo cáo so sánh: Excel: {excel_message}, PDF: {pdf_message}"

        except Exception as e:
            success = False
            message = f"Lỗi hệ thống khi tạo báo cáo so sánh: {e}"
        
        return success, message, excel_path, pdf_path

    else:
        return False, "Vui lòng chọn chế độ báo cáo (Tiêu chuẩn hoặc So sánh).", None, None


# Đoạn code ví dụ này thường chỉ để kiểm tra cục bộ, không chạy trong môi trường Streamlit chính
# if __name__ == '__main__':
#     path_dict = setup_paths()
#     template_file = path_dict['template_file']

#     df_raw_example = load_raw_data(template_file)
#     config_data_example = read_configs(template_file)

#     if df_raw_example is None or config_data_example is None:
#         print("Không thể tải dữ liệu hoặc cấu hình mẫu. Vui lòng kiểm tra file template.")
#     else:
#         print("Dữ liệu và cấu hình mẫu đã tải thành công.")

#         # --- Cấu hình cho Báo cáo Tiêu chuẩn ---
#         standard_report_mode = "Năm" # Có thể là "Năm", "Tháng", "Tuần"
#         standard_report_year = 2023
#         standard_report_months = ['January', 'February'] # Ví dụ: ['January', 'February'], để trống nếu muốn tất cả các tháng
#         standard_report_projects = ["Project Alpha"] # Thay thế bằng tên dự án của bạn

#         # --- Cấu hình cho Báo cáo So sánh ---
#         # comparison_report_mode = "So Sánh Dự Án Trong Một Tháng" # Có thể là:
#         #   "So Sánh Dự Án Trong Một Tháng"
#         #   "So Sánh Dự Án Trong Một Năm"
#         #   "So Sánh Một Dự Án Qua Các Tháng/Năm"
#         # comparison_report_mode = "So Sánh Một Dự Án Qua Các Tháng/Năm" 
        
#         comparison_years = [2022] # Ví dụ cho "So Sánh Một Dự Án Qua Các Tháng/Năm"
#         comparison_months = ['January', 'February'] # Để trống nếu so sánh theo năm, hoặc ['January', 'February'] nếu so sánh tháng trong một năm cụ thể.
#         comparison_projects = ["Project Alpha", "Project Beta"] # Ví dụ: ["Project Alpha", "Project Beta"] for "So Sánh Dự Án Trong Một Tháng/Năm"
#                                               # Hoặc ["Project Alpha"] for "So Sánh Một Dự Án Qua Các Tháng/Năm"

#         # Gọi hàm để tạo báo cáo
#         print("\n--- Tạo báo cáo Tiêu chuẩn mẫu ---")
#         success_std, msg_std, excel_std_path, pdf_std_path = generate_reports_on_demand(
#             df_raw=df_raw_example,
#             config_data=config_data_example, # Truyền config_data vào đây
#             selected_mode=standard_report_mode,
#             selected_year=standard_report_year,
#             selected_months=standard_report_months,
#             selected_project_names_standard=standard_report_projects,
#             comparison_config_years=[],
#             comparison_config_months=[],
#             comparison_config_projects=[],
#             comparison_report_mode=None,
#             path_dict=path_dict # Truyền path_dict vào đây
#         )
#         print(f"Báo cáo Tiêu chuẩn: Thành công: {success_std}, Thông báo: {msg_std}")
#         if excel_std_path:
#             print(f"File Excel: {excel_std_path}")
#         if pdf_std_path:
#             print(f"File PDF: {pdf_std_path}")

#         print("\n--- Tạo báo cáo So sánh mẫu ---")
#         success_comp, msg_comp, excel_comp_path, pdf_comp_path = generate_reports_on_demand(
#             df_raw=df_raw_example,
#             config_data=config_data_example, # Truyền config_data vào đây
#             selected_mode=None,
#             selected_year=None,
#             selected_months=[],
#             selected_project_names_standard=[],
#             comparison_config_years=comparison_years,
#             comparison_config_months=comparison_months,
#             comparison_config_projects=comparison_projects,
#             comparison_report_mode=comparison_report_mode,
#             path_dict=path_dict # Truyền path_dict vào đây
#         )
#         print(f"Báo cáo So sánh: Thành công: {success_comp}, Thông báo: {msg_comp}")
#         if excel_comp_path:
#             print(f"File Excel: {excel_comp_path}")
#         if pdf_comp_path:
#             print(f"File PDF: {pdf_comp_path}")
