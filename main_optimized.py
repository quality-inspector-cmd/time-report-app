import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ==============================================================================
# ĐẢM BẢO FILE 'a04ecaf1_1dae_4c90_8081_086cd7c7b725.py' NẰNG CÙNG THƯ MỤC
# HOẶNG THAY THẾ TÊN FILE NẾU BẠN ĐÃ ĐỔI TÊN NÓ.
# ==============================================================================
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    generate_reports_on_demand
)
# ==============================================================================

script_dir = os.path.dirname(__file__)
csv_file_path = os.path.join(script_dir, "invited_emails.csv")

# Gọi hàm setup_paths ngay từ đầu để path_dict có sẵn
path_dict = setup_paths()

# ==============================================================================
# KHỞI TẠO CÁC BIẾN TRẠNG THÁI PHIÊN (SESSION STATE VARIABLES)
# ==============================================================================
if 'comparison_selected_years' not in st.session_state:
    st.session_state.comparison_selected_years = []
if 'comparison_selected_months' not in st.session_state:
    st.session_state.comparison_selected_months = []
if 'comparison_selected_projects' not in st.session_state:
    st.session_state.comparison_selected_projects = []
if 'comparison_mode' not in st.session_state:
    st.session_state.comparison_mode = "So Sánh Dự Án Trong Một Tháng" # Giá trị mặc định

# ==============================================================================\
# HÀM HỖ TRỢ NGÔN NGỮ
# ==============================================================================\
LANGUAGES = {
    "en": {
        "title": "Time Report Automation Tool",
        "intro_message": "Welcome to the Time Report Automation Tool! Please navigate through the tabs to generate your reports.",
        "tabs": ["Standard Report", "Comparison Report", "Data Preview", "Instructions"],
        "upload_file_header": "Upload Time Report Template (Time_report.xlsm)",
        "upload_file_help": "Please upload your .xlsm file containing 'Raw Data', 'Config_Year_Mode', and 'Config_Project_Filter' sheets.",
        "loading_data": "Loading data and configurations...",
        "data_loaded_success": "Data and configurations loaded successfully!",
        "file_not_found_error": "Error: Template file not found.",
        "error_loading_data": "Error loading data or configurations.",
        "select_analysis_mode": "Select Analysis Mode:",
        "modes": ["Year", "Month", "Week"],
        "select_year": "Select Year:",
        "all_years": "All Years",
        "select_months": "Select Months:",
        "all_months": "All Months",
        "select_projects": "Select Projects:",
        "all_projects": "All Projects",
        "generate_standard_report_btn": "Generate Standard Report",
        "no_year_selected_error": "Please select a year.",
        "no_project_selected_warning_standard": "Please select at least one project for the standard report.",
        "export_options": "Export Options",
        "export_excel_option": "Export to Excel (.xlsx)",
        "export_pdf_option": "Export to PDF (.pdf)",
        "warning_select_export_format": "Please select at least one export format (Excel or PDF).",
        "generating_excel_report": "Generating Excel report...",
        "generating_pdf_report": "Generating PDF report...",
        "report_done": "Report generated successfully!",
        "error_generating_report": "Error occurred while generating report. Please try again.",
        "download_excel": "Download Excel Report",
        "download_pdf": "Download PDF Report",
        "download_comparison_excel": "Download Comparison Excel Report",
        "download_comparison_pdf": "Download Comparison PDF Report",
        "comparison_report_mode": "Select Comparison Mode:",
        "comparison_modes": {
            "So Sánh Dự Án Trong Một Tháng": "Compare Projects Within A Month",
            "So Sánh Dự Án Trong Một Năm": "Compare Projects Within A Year",
            "So Sánh Một Dự Án Qua Các Tháng/Năm": "Compare One Project Across Months/Years"
        },
        "select_comparison_years": "Select Years for Comparison:",
        "select_comparison_months": "Select Months for Comparison:",
        "select_comparison_projects": "Select Projects for Comparison:",
        "generate_comparison_report_btn": "Generate Comparison Report",
        "no_year_selected_comparison": "Please select a year for comparison.",
        "no_month_selected_comparison": "Please select a month for comparison.",
        "no_project_selected_comparison": "Please select at least one project for comparison.",
        "multiple_years_warning": "Please select only ONE year when comparing projects within a month or for one project across months.",
        "multiple_months_warning": "Please select only ONE month when comparing projects within a month.",
        "one_project_required": "For 'Compare One Project Across Months/Years', please select exactly ONE project.",
        "preview_data_header": "Raw Data Preview (First 100 Rows)",
        "no_data_loaded": "No raw data loaded. Please upload your template file.",
        "instructions_header": "Instructions",
        "instructions_intro": "This tool helps automate time report generation.",
        "instructions_standard_report": "### 1. Standard Report",
        "instructions_standard_report_desc": "Generate a detailed report based on selected year, months, and projects.",
        "instructions_comparison_report": "### 2. Comparison Report",
        "instructions_comparison_report_desc": "Generate a comparative analysis based on different modes:",
        "compare_projects_in_month_desc": "Compare multiple projects within a single selected month of a specific year.",
        "compare_projects_in_year_desc": "Compare multiple projects within a single selected year (all months included).",
        "compare_one_project_across_desc": "Compare a single project across multiple months within one year, OR across multiple years.",
        "instructions_data_preview": "### 3. Data Preview",
        "instructions_data_preview_desc": "This tab allows you to preview the first 100 rows of the loaded raw data, helping you verify the data format and content.",
        "instructions_template_config": "### 4. Template File Configuration (External to Application)",
        "instructions_template_config_desc": "The tool reads data and configurations from an Excel template file (typically `Time_report.xlsm`). Ensure that:",
        "template_sheet_data": "The 'Raw Data' sheet contains your raw time data with necessary columns like 'Year', 'MonthName', 'Project name', etc.",
        "template_sheet_config": "The 'Config_Year_Mode' and 'Config_Project_Filter' sheets can be used to set default configurations, but the selections on the interface will override them.",
        "common_errors_header": "### Common Errors:",
        "template_not_found": "Template file not found: Ensure `Time_report.xlsm` is in the same directory as this application.",
        "data_load_failure": "Raw data load failure: Check data format and column names in the 'Raw Data' sheet.",
        "language_select": "Select Language:"
    },
    "vi": {
        "title": "Công Cụ Tự Động Hóa Báo Cáo Giờ Làm Việc",
        "intro_message": "Chào mừng bạn đến với Công Cụ Tự Động Hóa Báo Cáo Giờ Làm Việc! Vui lòng điều hướng qua các tab để tạo báo cáo của bạn.",
        "tabs": ["Báo cáo Tiêu chuẩn", "Báo cáo So sánh", "Xem trước Dữ liệu", "Hướng dẫn"],
        "upload_file_header": "Tải lên File Template Báo cáo Giờ làm việc (Time_report.xlsm)",
        "upload_file_help": "Vui lòng tải lên file .xlsm của bạn có chứa các sheet 'Raw Data', 'Config_Year_Mode' và 'Config_Project_Filter'.",
        "loading_data": "Đang tải dữ liệu và cấu hình...",
        "data_loaded_success": "Dữ liệu và cấu hình đã tải thành công!",
        "file_not_found_error": "Lỗi: Không tìm thấy file template.",
        "error_loading_data": "Lỗi khi tải dữ liệu hoặc cấu hình.",
        "select_analysis_mode": "Chọn Chế độ phân tích:",
        "modes": ["Năm", "Tháng", "Tuần"],
        "select_year": "Chọn Năm:",
        "all_years": "Tất cả các Năm",
        "select_months": "Chọn Tháng:",
        "all_months": "Tất cả các Tháng",
        "select_projects": "Chọn Dự án:",
        "all_projects": "Tất cả các Dự án",
        "generate_standard_report_btn": "Tạo Báo cáo Tiêu chuẩn",
        "no_year_selected_error": "Vui lòng chọn một năm.",
        "no_project_selected_warning_standard": "Vui lòng chọn ít nhất một dự án cho báo cáo tiêu chuẩn.",
        "export_options": "Tùy chọn Xuất",
        "export_excel_option": "Xuất ra Excel (.xlsx)",
        "export_pdf_option": "Xuất ra PDF (.pdf)",
        "warning_select_export_format": "Vui lòng chọn ít nhất một định dạng xuất (Excel hoặc PDF).",
        "generating_excel_report": "Đang tạo báo cáo Excel...",
        "generating_pdf_report": "Đang tạo báo cáo PDF...",
        "report_done": "Báo cáo đã tạo thành công!",
        "error_generating_report": "Có lỗi xảy ra khi tạo báo cáo. Vui lòng thử lại.",
        "download_excel": "Tải xuống Báo cáo Excel",
        "download_pdf": "Tải xuống Báo cáo PDF",
        "download_comparison_excel": "Tải xuống Báo cáo So sánh Excel",
        "download_comparison_pdf": "Tải xuống Báo cáo So sánh PDF",
        "comparison_report_mode": "Chọn Chế độ So sánh:",
        "comparison_modes": {
            "So Sánh Dự Án Trong Một Tháng": "So Sánh Dự Án Trong Một Tháng",
            "So Sánh Dự Án Trong Một Năm": "So Sánh Dự Án Trong Một Năm",
            "So Sánh Một Dự Án Qua Các Tháng/Năm": "So Sánh Một Dự Án Qua Các Tháng/Năm"
        },
        "select_comparison_years": "Chọn Năm để So sánh:",
        "select_comparison_months": "Chọn Tháng để So sánh:",
        "select_comparison_projects": "Chọn Dự án để So sánh:",
        "generate_comparison_report_btn": "Tạo Báo cáo So sánh",
        "no_year_selected_comparison": "Vui lòng chọn một năm để so sánh.",
        "no_month_selected_comparison": "Vui lòng chọn một tháng để so sánh.",
        "no_project_selected_comparison": "Vui lòng chọn ít nhất một dự án để so sánh.",
        "multiple_years_warning": "Vui lòng chọn CHỈ MỘT năm khi so sánh các dự án trong một tháng hoặc cho một dự án qua các tháng.",
        "multiple_months_warning": "Vui lòng chọn CHỈ MỘT tháng khi so sánh các dự án trong một tháng.",
        "one_project_required": "Đối với 'So Sánh Một Dự Án Qua Các Tháng/Năm', vui lòng chọn CHỈ MỘT dự án.",
        "preview_data_header": "Xem trước Dữ liệu Thô (100 hàng đầu tiên)",
        "no_data_loaded": "Chưa có dữ liệu thô được tải. Vui lòng tải lên file template của bạn.",
        "instructions_header": "Hướng dẫn",
        "instructions_intro": "Công cụ này giúp tự động hóa việc tạo báo cáo giờ làm việc.",
        "instructions_standard_report": "### 1. Báo cáo Tiêu chuẩn",
        "instructions_standard_report_desc": "Tạo báo cáo chi tiết dựa trên năm, tháng và dự án đã chọn.",
        "instructions_comparison_report": "### 2. Báo cáo So sánh",
        "instructions_comparison_report_desc": "Tạo phân tích so sánh dựa trên các chế độ khác nhau:",
        "compare_projects_in_month_desc": "So sánh nhiều dự án trong một tháng được chọn của một năm cụ thể.",
        "compare_projects_in_year_desc": "So sánh nhiều dự án trong một năm được chọn (bao gồm tất cả các tháng).",
        "compare_one_project_across_desc": "So sánh một dự án duy nhất qua nhiều tháng trong cùng một năm, HOẶC so sánh qua nhiều năm.",
        "instructions_data_preview": "### 3. Xem trước Dữ liệu",
        "instructions_data_preview_desc": "Tab này cho phép bạn xem 100 hàng đầu tiên của dữ liệu thô đã tải, giúp bạn kiểm tra định dạng và nội dung dữ liệu.",
        "instructions_template_config": "### 4. Cấu hình File Template (Bên ngoài ứng dụng)",
        "instructions_template_config_desc": "Công cụ đọc dữ liệu và cấu hình từ một file Excel template (thường là `Time_report.xlsm`). Đảm bảo rằng:",
        "template_sheet_data": "Sheet 'Raw Data' chứa dữ liệu thời gian thô của bạn với các cột cần thiết như 'Year', 'MonthName', 'Project name', v.v.",
        "template_sheet_config": "Sheet 'Config_Year_Mode' và 'Config_Project_Filter' có thể được sử dụng để đặt cấu hình mặc định, nhưng các lựa chọn trên giao diện sẽ ghi đè lên chúng.",
        "common_errors_header": "### Lỗi thường gặp:",
        "template_not_found": "Không tìm thấy file template: Đảm bảo `Time_report.xlsm` nằm cùng thư mục với ứng dụng này.",
        "data_load_failure": "Không tải được dữ liệu thô: Kiểm tra định dạng dữ liệu và tên cột trong sheet 'Raw Data'.",
        "language_select": "Chọn Ngôn ngữ:"
    }
}

# Khởi tạo hoặc lấy ngôn ngữ từ session state
if 'language' not in st.session_state:
    st.session_state.language = "vi" # Mặc định là tiếng Việt

def get_text(key):
    return LANGUAGES[st.session_state.language].get(key, f"<{key} not found>")

# ==============================================================================\
# TIÊU ĐỀ ỨNG DỤNG
# ==============================================================================\
st.set_page_config(layout="wide")
st.title(get_text('title'))
st.write(get_text('intro_message'))

# Thanh sidebar để chọn ngôn ngữ
st.sidebar.header(get_text('language_select'))
selected_language = st.sidebar.radio("", ["Tiếng Việt", "English"], 
                                     index=0 if st.session_state.language == "vi" else 1)

if selected_language == "Tiếng Việt":
    st.session_state.language = "vi"
else:
    st.session_state.language = "en"

# ==============================================================================\
# TẢI DỮ LIỆU ĐẦU VÀO
# ==============================================================================\
@st.cache_data(ttl=3600) # Cache dữ liệu trong 1 giờ
def cached_load():
    template_file = path_dict['template_file']
    if not os.path.exists(template_file):
        st.error(get_text('file_not_found_error'))
        return None, None
    
    df_raw = load_raw_data(template_file)
    config_data = read_configs(template_file)
    return df_raw, config_data

df_raw, config_data = cached_load()

if df_raw is None or config_data is None:
    st.warning(get_text('error_loading_data'))
else:
    st.sidebar.success(get_text('data_loaded_success'))

# ==============================================================================\
# CÁC TAB CHÍNH CỦA ỨNG DỤNG
# ==============================================================================\
tab_names = [get_text('tabs')[0], get_text('tabs')[1], get_text('tabs')[2], get_text('tabs')[3]]
tab_standard_report_main, tab_comparison_report_main, tab_data_preview, tab_instructions = st.tabs(tab_names)

# ==============================================================================\
# TAB 1: BÁO CÁO TIÊU CHUẨN
# ==============================================================================\
with tab_standard_report_main:
    st.header(get_text('tabs')[0])

    if df_raw is not None and config_data is not None:
        all_years = sorted(df_raw['Year'].unique().tolist(), reverse=True)
        all_months = ['January', 'February', 'March', 'April', 'May', 'June', 
                      'July', 'August', 'September', 'October', 'November', 'December']
        all_projects = sorted(df_raw['Project name'].unique().tolist())

        analysis_mode = st.radio(get_text('select_analysis_mode'), get_text('modes'))

        # Lựa chọn Năm
        selected_year_option = st.selectbox(get_text('select_year'), [get_text('all_years')] + all_years, index=0)
        selected_year = selected_year_option if selected_year_option != get_text('all_years') else None

        # Lựa chọn Tháng (chỉ hiển thị nếu chế độ là "Tháng")
        selected_months_standard = []
        if analysis_mode == get_text('modes')[1]: # Tháng
            selected_months_options = st.multiselect(get_text('select_months'), all_months, default=all_months)
            selected_months_standard = selected_months_options if selected_months_options else []
        elif analysis_mode == get_text('modes')[0]: # Năm
            selected_months_standard = all_months # Mặc định chọn tất cả các tháng cho chế độ "Năm"
        # Chế độ "Tuần" sẽ không lọc theo tháng cụ thể, sẽ xử lý logic tuần bên trong generate_reports_on_demand

        # Lựa chọn Dự án
        selected_projects_standard = st.multiselect(get_text('select_projects'), all_projects, default=all_projects)
        if not selected_projects_standard:
            st.warning(get_text('no_project_selected_warning_standard'))


        if st.button(get_text('generate_standard_report_btn'), key='generate_std_report_btn'):
            if not selected_year and selected_year_option != get_text('all_years'):
                st.warning(get_text('no_year_selected_error'))
            elif not selected_projects_standard:
                st.warning(get_text('no_project_selected_warning_standard'))
            else:
                export_excel, export_pdf = False, False
                with st.expander(get_text('export_options')):
                    export_excel = st.checkbox(get_text('export_excel_option'), value=True, key='export_excel_std')
                    export_pdf = st.checkbox(get_text('export_pdf_option'), value=True, key='export_pdf_std')

                if not export_excel and not export_pdf:
                    st.warning(get_text('warning_select_export_format'))
                else:
                    with st.spinner(get_text('generating_excel_report') if export_excel else get_text('generating_pdf_report')):
                        try:
                            # Gọi hàm generate_reports_on_demand
                            success, message, excel_file_path, pdf_file_path = generate_reports_on_demand(
                                df_raw=df_raw,
                                config_data=config_data, # Truyền config_data vào đây
                                selected_mode=analysis_mode,
                                selected_year=selected_year,
                                selected_months=selected_months_standard,
                                selected_project_names_standard=selected_projects_standard,
                                comparison_config_years=[], # Không dùng cho báo cáo tiêu chuẩn
                                comparison_config_months=[], # Không dùng cho báo cáo tiêu chuẩn
                                comparison_config_projects=[], # Không dùng cho báo cáo tiêu chuẩn
                                comparison_report_mode=None, # Không dùng cho báo cáo tiêu chuẩn
                                path_dict=path_dict # Truyền path_dict vào đây
                            )

                            if success:
                                st.success(get_text('report_done'))
                                if excel_file_path and os.path.exists(excel_file_path):
                                    with open(excel_file_path, "rb") as file:
                                        st.download_button(
                                            label=get_text('download_excel'),
                                            data=file,
                                            file_name=os.path.basename(excel_file_path),
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            key='download_std_excel'
                                        )
                                if pdf_file_path and os.path.exists(pdf_file_path):
                                    with open(pdf_file_path, "rb") as file:
                                        st.download_button(
                                            label=get_text('download_pdf'),
                                            data=file,
                                            file_name=os.path.basename(pdf_file_path),
                                            mime="application/pdf",
                                            key='download_std_pdf'
                                        )
                            else:
                                st.error(f"{get_text('error_generating_report')} {message}")

                        except Exception as e:
                            st.error(f"{get_text('error_generating_report')} {e}")
                            print(f"Error generating standard report: {e}")
    else:
        st.info(get_text('no_data_loaded'))


# ==============================================================================\
# TAB 2: BÁO CÁO SO SÁNH
# ==============================================================================\
with tab_comparison_report_main:
    st.header(get_text('tabs')[1])

    if df_raw is not None and config_data is not None:
        all_years = sorted(df_raw['Year'].unique().tolist(), reverse=True)
        all_months = ['January', 'February', 'March', 'April', 'May', 'June', 
                      'July', 'August', 'September', 'October', 'November', 'December']
        all_projects = sorted(df_raw['Project name'].unique().tolist())

        # Chọn chế độ so sánh
        comparison_mode_key = st.radio(get_text('comparison_report_mode'), 
                                       list(get_text('comparison_modes').keys()),
                                       key='comp_mode_radio')
        st.session_state.comparison_mode = comparison_mode_key # Cập nhật session state

        # Lựa chọn Năm cho so sánh
        if st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Tháng" or \
           st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Năm":
            comparison_selected_years_options = st.selectbox(get_text('select_comparison_years'), 
                                                             [get_text('all_years')] + all_years,
                                                             key='comp_year_select')
            st.session_state.comparison_selected_years = [comparison_selected_years_options] if comparison_selected_years_options != get_text('all_years') else []
            if st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Tháng" and \
               len(st.session_state.comparison_selected_years) != 1:
                st.warning(get_text('multiple_years_warning'))
        elif st.session_state.comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
             comparison_selected_years_options = st.multiselect(get_text('select_comparison_years'), 
                                                                all_years, 
                                                                default=st.session_state.comparison_selected_years,
                                                                key='comp_years_multi')
             st.session_state.comparison_selected_years = comparison_selected_years_options
             if len(st.session_state.comparison_selected_years) == 0 and len(st.session_state.comparison_selected_months) == 0:
                 st.warning("Vui lòng chọn ít nhất một năm hoặc một tháng để so sánh.")


        # Lựa chọn Tháng cho so sánh
        if st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Tháng":
            comparison_selected_months_options = st.selectbox(get_text('select_comparison_months'), 
                                                              all_months,
                                                              key='comp_month_select')
            st.session_state.comparison_selected_months = [comparison_selected_months_options] if comparison_selected_months_options else []
            if len(st.session_state.comparison_selected_months) != 1:
                st.warning(get_text('multiple_months_warning'))
        elif st.session_state.comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
            # Chỉ cho phép chọn tháng nếu chỉ có 1 năm được chọn
            if len(st.session_state.comparison_selected_years) == 1:
                comparison_selected_months_options = st.multiselect(get_text('select_comparison_months'), 
                                                                    all_months, 
                                                                    default=st.session_state.comparison_selected_months,
                                                                    key='comp_months_multi')
                st.session_state.comparison_selected_months = comparison_selected_months_options
            elif len(st.session_state.comparison_selected_years) > 1:
                st.info("Khi so sánh một dự án qua nhiều năm, lựa chọn tháng sẽ bị bỏ qua.")
                st.session_state.comparison_selected_months = []
            else:
                st.session_state.comparison_selected_months = []


        # Lựa chọn Dự án cho so sánh
        if st.session_state.comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
            comparison_selected_projects_options = st.selectbox(get_text('select_comparison_projects'), 
                                                                all_projects,
                                                                key='comp_project_select_single')
            st.session_state.comparison_selected_projects = [comparison_selected_projects_options] if comparison_selected_projects_options else []
            if len(st.session_state.comparison_selected_projects) != 1:
                st.warning(get_text('one_project_required'))
        else:
            comparison_selected_projects_options = st.multiselect(get_text('select_comparison_projects'), 
                                                                  all_projects, 
                                                                  default=st.session_state.comparison_selected_projects,
                                                                  key='comp_project_select_multi')
            st.session_state.comparison_selected_projects = comparison_selected_projects_options
            if not st.session_state.comparison_selected_projects:
                st.warning(get_text('no_project_selected_comparison'))


        if st.button(get_text('generate_comparison_report_btn'), key='generate_comp_report_btn'):
            # Kiểm tra các điều kiện dựa trên chế độ so sánh
            can_generate = True
            if st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Tháng":
                if not st.session_state.comparison_selected_years or len(st.session_state.comparison_selected_years) != 1:
                    st.warning(get_text('multiple_years_warning'))
                    can_generate = False
                if not st.session_state.comparison_selected_months or len(st.session_state.comparison_selected_months) != 1:
                    st.warning(get_text('multiple_months_warning'))
                    can_generate = False
                if not st.session_state.comparison_selected_projects:
                    st.warning(get_text('no_project_selected_comparison'))
                    can_generate = False
            elif st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Năm":
                if not st.session_state.comparison_selected_years or len(st.session_state.comparison_selected_years) != 1:
                    st.warning(get_text('multiple_years_warning')) # Ở đây là cần 1 năm, không phải nhiều năm
                    can_generate = False
                if not st.session_state.comparison_selected_projects:
                    st.warning(get_text('no_project_selected_comparison'))
                    can_generate = False
            elif st.session_state.comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
                if not st.session_state.comparison_selected_projects or len(st.session_state.comparison_selected_projects) != 1:
                    st.warning(get_text('one_project_required'))
                    can_generate = False
                if not st.session_state.comparison_selected_years and not st.session_state.comparison_selected_months:
                    st.warning("Vui lòng chọn ít nhất một năm hoặc một tháng để so sánh.")
                    can_generate = False
                if len(st.session_state.comparison_selected_years) > 1 and len(st.session_state.comparison_selected_months) > 0:
                     st.warning("Bạn không thể so sánh nhiều năm VÀ nhiều tháng cùng lúc. Vui lòng chọn chỉ năm HOẶC chỉ tháng.")
                     can_generate = False

            if can_generate:
                export_excel_comp, export_pdf_comp = False, False
                with st.expander(get_text('export_options')):
                    export_excel_comp = st.checkbox(get_text('export_excel_option'), value=True, key='export_excel_comp')
                    export_pdf_comp = st.checkbox(get_text('export_pdf_option'), value=True, key='export_pdf_comp')

                if not export_excel_comp and not export_pdf_comp:
                    st.warning(get_text('warning_select_export_format'))
                else:
                    with st.spinner(get_text('generating_comparison_excel') if export_excel_comp else get_text('generating_comparison_pdf')):
                        try:
                            # Gọi hàm generate_reports_on_demand
                            success, message, excel_file_path, pdf_file_path = generate_reports_on_demand(
                                df_raw=df_raw,
                                config_data=config_data, # THÊM DÒNG NÀY
                                selected_mode=None, # Không dùng cho báo cáo so sánh
                                selected_year=None, # Không dùng cho báo cáo so sánh
                                selected_months=[], # Không dùng cho báo cáo so sánh
                                selected_project_names_standard=[], # Không dùng cho báo cáo so sánh
                                comparison_config_years=st.session_state.comparison_selected_years,
                                comparison_config_months=st.session_state.comparison_selected_months,
                                comparison_config_projects=st.session_state.comparison_selected_projects,
                                comparison_report_mode=st.session_state.comparison_mode,
                                path_dict=path_dict # ĐẢM BẢO DÒNG NÀY CÓ MẶT
                            )
                            
                            if success:
                                st.success(get_text('report_done'))
                                if excel_file_path and os.path.exists(excel_file_path):
                                    with open(excel_file_path, "rb") as file:
                                        st.download_button(
                                            label=get_text('download_comparison_excel'),
                                            data=file,
                                            file_name=os.path.basename(excel_file_path),
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            key='download_comp_excel'
                                        )
                                if pdf_file_path and os.path.exists(pdf_file_path):
                                    with open(pdf_file_path, "rb") as file:
                                        st.download_button(
                                            label=get_text('download_comparison_pdf'),
                                            data=file,
                                            file_name=os.path.basename(pdf_file_path),
                                            mime="application/pdf",
                                            key='download_comp_pdf'
                                        )
                            else:
                                st.error(f"{get_text('error_generating_report')} {message}")

                        except Exception as e:
                            st.error(f"{get_text('error_generating_report')} {e}")
                            print(f"Error generating comparison report: {e}")
    else:
        st.info(get_text('no_data_loaded'))

# ==============================================================================\
# TAB 3: XEM TRƯỚC DỮ LIỆU
# ==============================================================================\
with tab_data_preview:
    st.header(get_text('preview_data_header'))
    if df_raw is not None:
        st.dataframe(df_raw.head(100))
    else:
        st.info(get_text('no_data_loaded'))

# ==============================================================================\
# TAB 4: HƯỚNG DẪN
# ==============================================================================\
with tab_instructions:
    st.header(get_text('instructions_header'))
    st.write(get_text('instructions_intro'))

    st.markdown(get_text('instructions_standard_report'))
    st.write(get_text('instructions_standard_report_desc'))

    st.markdown(get_text('instructions_comparison_report'))
    st.write(get_text('instructions_comparison_report_desc'))
    st.markdown(f"- **{get_text('comparison_modes')['So Sánh Dự Án Trong Một Tháng']}:** {get_text('compare_projects_in_month_desc')}")
    st.markdown(f"- **{get_text('comparison_modes')['So Sánh Dự Án Trong Một Năm']}:** {get_text('compare_projects_in_year_desc')}")
    st.markdown(f"- **{get_text('comparison_modes')['So Sánh Một Dự Án Qua Các Tháng/Năm']}:** {get_text('compare_one_project_across_desc')}")
    
    st.markdown(get_text('instructions_data_preview'))
    st.write(get_text('instructions_data_preview_desc'))

    st.markdown(get_text('instructions_template_config'))
    st.write(get_text('instructions_template_config_desc'))
    st.markdown(f"- {get_text('template_sheet_data')}")
    st.markdown(f"- {get_text('template_sheet_config')}")

    st.markdown(get_text('common_errors_header'))
    st.markdown(f"- **{get_text('template_not_found')}")
    st.markdown(f"- **{get_text('data_load_failure')}")
