# a04ecaf1_1dae_4c90_8081_086cd7c7b725.py
import pandas as pd
import os
# import các thư viện khác bạn cần cho việc xử lý và xuất báo cáo (ví dụ: openpyxl, reportlab, matplotlib/plotly)

def setup_paths():
    # ... (logic hiện tại của bạn) ...
    # Có thể thêm các đường dẫn riêng cho báo cáo so sánh nếu cần
    return {
        'output_file': 'TimeReport_Standard.xlsx',
        'pdf_report': 'TimeReport_Standard.pdf',
        'comparison_output_file': 'TimeReport_Comparison.xlsx', # Đường dẫn mới
        'comparison_pdf_report': 'TimeReport_Comparison.pdf' # Đường dẫn mới
    }

def load_raw_data(path_dict):
    # ... (logic hiện tại của bạn) ...
    pass

def read_configs(path_dict):
    # ... (logic hiện tại của bạn) ...
    pass

def apply_filters(df_raw, config):
    # ... (logic hiện tại của bạn) ...
    pass

def export_report(df_filtered, config, path_dict):
    # ... (logic hiện tại của bạn để xuất Excel tiêu chuẩn) ...
    pass

def export_pdf_report(df_filtered, config, path_dict):
    # ... (logic hiện tại của bạn để xuất PDF tiêu chuẩn) ...
    pass

# =========================================================================
# CÁC HÀM MỚI CHO BÁO CÁO SO SÁNH - BẠN CẦN TRIỂN KHAI LOGIC TẠI ĐÂY
# =========================================================================

def apply_comparison_filters(df_raw, comparison_config, comparison_mode):
    """
    Lọc và chuẩn bị dữ liệu cho báo cáo so sánh dựa trên comparison_mode.
    Trả về DataFrame đã được tổng hợp/pivot.
    """
    years = comparison_config['years']
    months = comparison_config['months']
    selected_projects = comparison_config['selected_projects']

    df_filtered = df_raw[
        df_raw['Year'].isin(years) &
        df_raw['MonthName'].isin(months) &
        df_raw['Project Name'].isin(selected_projects)
    ].copy() # Luôn tạo bản sao khi lọc để tránh SettingWithCopyWarning

    if df_filtered.empty:
        return pd.DataFrame()

    if comparison_mode == "So Sánh Dự Án Trong Một Tháng" or comparison_mode == "Compare Projects in a Month":
        # Tổng hợp giờ theo Project cho một tháng/năm cụ thể
        # Giả sử bạn có cột 'Hours' chứa số giờ làm việc
        df_comparison = df_filtered.groupby('Project Name')['Hours'].sum().reset_index()
        df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
        return df_comparison

    elif comparison_mode == "So Sánh Dự Án Trong Một Năm" or comparison_mode == "Compare Projects in a Year":
        # Tổng hợp giờ theo Project và Month cho một năm
        df_comparison = df_filtered.groupby(['Project Name', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
        df_comparison.loc['Total'] = df_comparison.sum() # Thêm dòng tổng
        return df_comparison

    elif comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm" or comparison_mode == "Compare One Project Over Time (Months/Years)":
        # Tổng hợp giờ cho một Project theo Month và Year
        if len(selected_projects) == 1:
            df_comparison = df_filtered.groupby(['Year', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
            df_comparison.loc['Total'] = df_comparison.sum() # Thêm dòng tổng
            return df_comparison
        else:
            return pd.DataFrame() # Hoặc xử lý lỗi nếu không phải một dự án

    return pd.DataFrame() # Trả về DataFrame rỗng nếu không khớp chế độ nào

def export_comparison_report(df_comparison, comparison_config, path_dict, comparison_mode):
    """
    Xuất báo cáo so sánh ra file Excel.
    Sử dụng path_dict['comparison_output_file'] để lưu.
    """
    output_file = path_dict['comparison_output_file']

    # Ví dụ đơn giản: ghi DataFrame vào Excel
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df_comparison.to_excel(writer, sheet_name='Comparison Report', index=True)
        workbook = writer.book
        worksheet = writer.sheets['Comparison Report']

        # Thêm tiêu đề báo cáo vào Excel
        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
        worksheet.merge_range('A1:C1', 'BÁO CÁO SO SÁNH', title_format) # Điều chỉnh phạm vi merge

        # Thêm thông tin cấu hình
        info_format = workbook.add_format({'font_size': 10})
        worksheet.write('A2', f"Chế độ so sánh: {comparison_mode}", info_format)
        worksheet.write('A3', f"Năm: {comparison_config['years']}", info_format)
        worksheet.write('A4', f"Tháng: {comparison_config['months']}", info_format)
        worksheet.write('A5', f"Dự án: {comparison_config['selected_projects']}", info_format)

        # Tự động điều chỉnh độ rộng cột (ví dụ)
        for i, col in enumerate(df_comparison.columns):
            max_len = max(df_comparison[col].astype(str).map(len).max(), len(str(col)))
            worksheet.set_column(i + 1, i + 1, max_len + 2) # +1 vì cột A là index

    print(f"DEBUG: Comparison Excel report generated at {output_file}")
    return True # Trả về True nếu thành công

def export_comparison_pdf_report(df_comparison, comparison_config, path_dict, comparison_mode):
    """
    Xuất báo cáo so sánh ra file PDF.
    Sử dụng path_dict['comparison_pdf_report'] để lưu.
    Bạn có thể dùng ReportLab, FPDF, hoặc Matplotlib để vẽ biểu đồ và xuất PDF.
    """
    pdf_file = path_dict['comparison_pdf_report']

    # Đây là một placeholder đơn giản.
    # Thực tế bạn sẽ cần thư viện như ReportLab hoặc fpdf để tạo PDF phức tạp hơn,
    # hoặc matplotlib/seaborn để tạo biểu đồ rồi lưu thành ảnh, nhúng vào PDF.

    with open(pdf_file, 'w', encoding='utf-8') as f:
        f.write(f"BÁO CÁO SO SÁNH - {comparison_mode}\n\n")
        f.write(f"Cấu hình: Năm={comparison_config['years']}, Tháng={comparison_config['months']}, Dự án={comparison_config['selected_projects']}\n\n")
        f.write("Dữ liệu:\n")
        f.write(df_comparison.to_string()) # Chuyển DataFrame thành chuỗi để ghi vào PDF đơn giản

    print(f"DEBUG: Comparison PDF report generated at {pdf_file}")
    return True # Trả về True nếu thành công
