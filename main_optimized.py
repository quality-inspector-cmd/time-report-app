import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ==============================================================================
# ĐẢM BẢO FILE 'a04ecaf1_1dae_4c90_8081_086cd7c7b725.py' NẰNG CÙNG THƯ MỤC
# HOẶC THAY THẾ TÊN FILE NẾU BẠN ĐÃ ĐỔI TÊN NÓ.
# ==============================================================================
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    # apply_filters, export_report, export_pdf_report, # Các hàm này sẽ được gọi bên trong generate_reports_on_demand
    # apply_comparison_filters, export_comparison_report, export_comparison_pdf_report, # Các hàm này cũng vậy
    generate_reports_on_demand # <--- THÊM DÒNG NÀY ĐỂ IMPORT HÀM MỚI
)
# ==============================================================================

script_dir = os.path.dirname(__file__)
csv_file_path = os.path.join(script_dir, "invited_emails.csv")

# Gọi hàm setup_paths ngay từ đầu để path_dict có sẵn
path_dict = setup_paths()

# ==============================================================================
# KHỞI TẠO CÁC BIẾN TRẠNG THÁI PHIÊN (SESSION STATE VARIABLES)
# ==============================================================================
if 'comparison_mode' not in st.session_state:
    st.session_state.comparison_mode = "So Sánh Dự Án Trong Một Tháng" # Hoặc giá trị mặc định phù hợp

if 'comparison_selected_years' not in st.session_state:
    st.session_state.comparison_selected_years = [datetime.now().year] # Hoặc giá trị mặc định phù hợp

if 'comparison_selected_months' not in st.session_state:
    st.session_state.comparison_selected_months = [] # Hoặc giá trị mặc định phù hợp

if 'comparison_selected_projects' not in st.session_state:
    st.session_state.comparison_selected_projects = [] # Hoặc giá trị mặc định phù hợp

if 'comparison_selected_months_over_time' not in st.session_state:
    st.session_state.comparison_selected_months_over_time = [] # Khởi tạo là một danh sách rỗng hoặc giá trị mặc định phù hợp

if 'selected_years' not in st.session_state: # Ví dụ cho bộ lọc báo cáo tiêu chuẩn
    st.session_state.selected_years = [datetime.now().year]

if 'selected_months' not in st.session_state: # Ví dụ cho bộ lọc báo cáo tiêu chuẩn
    st.session_state.selected_months = []

# Thêm dòng này để mặc định ngôn ngữ là tiếng Anh
if 'selected_language' not in st.session_state:
    st.session_state.selected_language = "English"

# ---------------------------
# PHẦN XÁC THỰC TRUY CẬP
# ---------------------------

@st.cache_data
def load_invited_emails():
    try:
        df = pd.read_csv(csv_file_path, header=None, encoding='utf-8')
        # Sửa lỗi: Thêm .str trước .strip()
        emails = df.iloc[:, 0].astype(str).str.strip().str.lower().tolist()
        return emails
    except FileNotFoundError:
        st.error(f"Lỗi: Không tìm thấy file invited_emails.csv tại {csv_file_path}. Vui lòng kiểm tra đường dẫn.")
        return []
    except Exception as e:
        st.error(f"Lỗi khi tải file invited_emails.csv: {e}")
        return []

# Tải danh sách email được mời một lần
INVITED_EMAILS = load_invited_emails()

# Hàm ghi log truy cập
def log_user_access(email):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = {"Time": timestamp, "Email": email}
    if "access_log" not in st.session_state:
        st.session_state.access_log = []
    st.session_state.access_log.append(log_entry)

# Logic xác thực người dùng
if "user_email" not in st.session_state:
    st.set_page_config(page_title="Triac Time Report", layout="wide")
    st.title("🔐 Access authentication")
    email_input = st.text_input("📧 Enter the invited email to access:")

    if email_input:
        email = email_input.strip().lower()
        if email in INVITED_EMAILS:
            st.session_state.user_email = email
            log_user_access(email)
            st.success("✅ Valid email! Entering application...")
            st.rerun()
        else:
            st.error("❌ Email is not on the invitation list.")
    st.stop() # Dừng thực thi nếu chưa xác thực

# ---------------------------
# PHẦN GIAO DIỆN CHÍNH CỦA ỨNG DỤNG
# ---------------------------
# Sử dụng session_state để lưu trữ lựa chọn ngôn ngữ
if 'lang' not in st.session_state:
    st.session_state.lang = 'en' # Mặc định là tiếng Anh

# Cấu hình trang (chỉ chạy một lần sau khi xác thực)
st.set_page_config(page_title="Triac Time Report", layout="wide")

st.markdown("""
    <style>
        .report-title {font-size: 30px; color: #003366; font-weight: bold;}
        .report-subtitle {font-size: 14px; color: gray;}
        footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# =====================================
# Khởi tạo ngôn ngữ và từ điển văn bản
# =====================================
# Từ điển cho các chuỗi văn bản
TEXTS = {
    'vi': {
        'app_title': "📊 Công cụ tạo báo cáo thời gian",
        'lang_select': "Chọn ngôn ngữ:",
        'language_vi': "Tiếng Việt",
        'language_en': "English",
        'template_not_found': "❌ Không tìm thấy file template: {}. Vui lòng đảm bảo file nằm cùng thư mục với ứng dụng.",
        'failed_to_load_raw_data': "⚠️ Không thể tải dữ liệu thô. Vui lòng kiểm tra sheet 'Raw Data' trong file template và định dạng dữ liệu.",
        'loading_data': "🔄 Đang tải dữ liệu và cấu hình...",
        'tab_standard_report': "Báo cáo tiêu chuẩn",
        'tab_comparison_report': "Báo cáo so sánh",
        'tab_data_preview': "Xem trước dữ liệu",
        'standard_report_header': "Cấu hình báo cáo thời gian tiêu chuẩn",
        'select_analysis_mode': "Chọn chế độ phân tích:",
        'select_year': "Chọn năm:",
        'select_months': "Chọn tháng(các tháng):",
        'standard_project_selection_header': "Lựa chọn dự án cho báo cáo tiêu chuẩn",
        'standard_project_selection_text': "Chọn dự án để bao gồm (mặc định chỉ bao gồm các dự án 'yes' từ cấu hình template):",
        'generate_standard_report_btn': "🚀 Tạo báo cáo tiêu chuẩn",
        'no_year_selected_error': "Vui lòng chọn một năm hợp lệ để tạo báo cáo.",
        'no_project_selected_warning_standard': "Vui lòng chọn ít nhất một dự án để tạo báo cáo tiêu chuẩn.",
        'no_data_after_filter_standard': "⚠️ Không có dữ liệu sau khi lọc cho báo cáo tiêu chuẩn. Vui lòng kiểm tra các lựa chọn của bạn.",
        'generating_excel_report': "Đang tạo báo cáo Excel...",
        'excel_report_generated': "✅ Báo cáo Excel đã được tạo: {}",
        'download_excel_report': "📥 Tải báo cáo Excel",
        'generating_pdf_report': "Đang tạo báo cáo PDF...",
        'pdf_report_generated': "✅ Báo cáo PDF đã được tạo: {}",
        'download_pdf_report': "📥 Tải báo cáo PDF",
        'failed_to_generate_excel': "❌ Đã xảy ra lỗi khi tạo báo cáo Excel.",
        'failed_to_generate_pdf': "❌ Đã xảy ra lỗi khi tạo báo cáo PDF.",
        'comparison_report_header': "Cấu hình báo cáo so sánh",
        'select_comparison_mode': "Chọn chế độ so sánh:",
        'compare_projects_month': "So Sánh Dự Án Trong Một Tháng",
        'compare_projects_year': "So Sánh Dự Án Trong Một Năm",
        'compare_one_project_over_time': "So Sánh Một Dự Án Qua Các Tháng/Năm",
        'filter_data_for_comparison': "Lọc dữ liệu để so sánh",
        'select_years': "Chọn năm(các năm):", # Dùng chung cho các mode
        'select_months_comp': "Chọn tháng(các tháng):", # Dùng chung cho các mode
        'select_projects_comp': "Chọn dự án(các dự án):", # Dùng chung cho các mode
        'generate_comparison_report_btn': "🚀 Tạo báo cáo so sánh",
        'no_data_after_filter_comparison': "⚠️ {}",
        'data_filtered_success': "✅ Dữ liệu đã được lọc thành công cho so sánh.",
        'comparison_data_preview': "Xem trước dữ liệu so sánh",
        'generating_comparison_excel': "Đang tạo báo cáo Excel so sánh...",
        'comparison_excel_generated': "✅ Báo cáo Excel so sánh đã được tạo: {}",
        'download_comparison_excel': "📥 Tải báo cáo Excel so sánh",
        'generating_comparison_pdf': "Đang tạo báo cáo PDF so sánh...",
        'comparison_pdf_generated': "✅ Báo cáo PDF so sánh đã được tạo: {}",
        'download_comparison_pdf': "📥 Tải báo cáo PDF so sánh",
        'failed_to_generate_comparison_excel': "❌ Đã xảy ra lỗi khi tạo báo cáo Excel so sánh.",
        'failed_to_generate_comparison_pdf': "❌ Đã xảy ra lỗi khi tạo báo cáo PDF so sánh.",
        'raw_data_preview_header': "Dữ liệu đầu vào thô (100 hàng đầu)",
        'no_raw_data': "Không có dữ liệu thô được tải.",
        'no_year_in_data': "Không có năm nào trong dữ liệu để chọn.",
        'user_guide': "Hướng dẫn sử dụng",
        'export_options': "Tùy chọn xuất báo cáo",
        'export_excel_option': "Xuất ra Excel (.xlsx)",
        'export_pdf_option': "Xuất ra PDF (.pdf)",
        'report_button': "Tạo báo cáo",
        'no_data': "Không có dữ liệu sau khi lọc",
        'report_done': "Đã tạo báo cáo",
        'download_excel': "Tải Excel",
        'download_pdf': "Tải PDF",
        'warning_select_export_format': "Vui lòng chọn ít nhất một định dạng xuất báo cáo (Excel hoặc PDF).",
        'error_generating_report': "Có lỗi xảy ra khi tạo báo cáo. Vui lòng thử lại.",
        # Thêm các tin nhắn mới cho mode "So Sánh Một Dự Án Qua Các Tháng/Năm"
        'select_single_project_warning': "Vui lòng chọn CHỈ MỘT dự án cho chế độ này.",
        'select_years_for_over_time_months': "Chọn năm (hoặc các năm) bạn muốn so sánh:",
        'select_months_for_single_year': "Chọn tháng(các tháng) trong năm đã chọn:",
        'comparison_over_years_note': "Lưu ý: Bạn đã chọn nhiều năm. Báo cáo sẽ so sánh dữ liệu của dự án qua các năm đã chọn. Lựa chọn tháng sẽ bị bỏ qua.",
        'comparison_over_months_note': "Lưu ý: Báo cáo sẽ so sánh dữ liệu của dự án qua các tháng đã chọn trong năm {}.",
        'no_comparison_criteria_selected': "Vui lòng chọn ít nhất một năm hoặc một tháng để so sánh.",
        'no_month_selected_for_single_year': "Vui lòng chọn ít nhất một tháng khi so sánh một dự án trong một năm cụ thể."
    },
    'en': {
        'app_title': "📊 Time Report Generator",
        'lang_select': "Select language:",
        'language_vi': "Tiếng Việt",
        'language_en': "English",
        'template_not_found': "❌ Template file not found: {}. Please ensure the file is in the same directory as the application.",
        'failed_to_load_raw_data': "⚠️ Failed to load raw data. Please check the 'Raw Data' sheet in the template file and data format.",
        'loading_data': "🔄 Loading data and configurations...",
        'tab_standard_report': "Standard Report",
        'tab_comparison_report': "Comparison Report",
        'tab_data_preview': "Data Preview",
        'standard_report_header': "Standard Time Report Configuration",
        'select_analysis_mode': "Select analysis mode:",
        'select_year': "Select year:",
        'select_months': "Select month(s):",
        'standard_project_selection_header': "Project Selection for Standard Report",
        'standard_project_selection_text': "Select projects to include (only 'yes' projects from template config will be included by default):",
        'generate_standard_report_btn': "🚀 Generate Standard Report",
        'no_year_selected_error': "Please select a valid year to generate the report.",
        'no_project_selected_warning_standard': "Please select at least one project to generate the standard report.",
        'no_data_after_filter_standard': "⚠️ No data after filtering for the standard report. Please check your selections.",
        'generating_excel_report': "Generating Excel report...",
        'excel_report_generated': "✅ Excel Report generated: {}",
        'download_excel_report': "📥 Download Excel Report",
        'generating_pdf_report': "Generating PDF report...",
        'pdf_report_generated': "✅ PDF Report generated: {}",
        'download_pdf_report': "📥 Download PDF Report",
        'failed_to_generate_excel': "❌ Failed to generate Excel report.",
        'failed_to_generate_pdf': "❌ Failed to generate PDF report.",
        'comparison_report_header': "Comparison Report Configuration",
        'select_comparison_mode': "Select comparison mode:",
        'compare_projects_month': "Compare Projects in a Month",
        'compare_projects_year': "Compare Projects in a Year",
        'compare_one_project_over_time': "Compare One Project Over Time (Months/Years)",
        'filter_data_for_comparison': "Filter Data for Comparison",
        'select_years': "Select Year(s):",
        'select_months_comp': "Select Month(s):",
        'select_projects_comp': "Select Project(s):",
        'generate_comparison_report_btn': "🚀 Generate Comparison Report",
        'no_data_after_filter_comparison': "⚠️ {}",
        'data_filtered_success': "✅ Data filtered successfully for comparison.",
        'comparison_data_preview': "Comparison Data Preview",
        'generating_comparison_excel': "Generating Comparison Excel Report...",
        'comparison_excel_generated': "✅ Comparison Excel Report generated: {}",
        'download_comparison_excel': "📥 Download Comparison Excel",
        'generating_comparison_pdf': "Generating Comparison PDF Report...",
        'comparison_pdf_generated': "✅ PDF Report generated: {}",
        'download_comparison_pdf': "📥 Download Comparison PDF",
        'failed_to_generate_comparison_excel': "❌ Failed to generate Comparison Excel report.",
        'failed_to_generate_comparison_pdf': "❌ Failed to generate Comparison PDF report.",
        'raw_data_preview_header': "Raw Input Data (First 100 rows)",
        'no_raw_data': "No raw data loaded.",
        'no_year_in_data': "No years in data to select.",
        'user_guide': "User Guide",
        'export_options': "Export Options",
        'export_excel_option': "Export as Excel (.xlsx)",
        'export_pdf_option': "Export as PDF (.pdf)",
        'report_button': "Generate report",
        'no_data': "No data after filtering",
        'report_done': "Report created successfully",
        'download_excel': "Download Excel",
        'download_pdf': "Download PDF",
        'warning_select_export_format': "Please select at least one report export format (Excel or PDF).",
        'error_generating_report': "An error occurred while generating the report. Please try again.",
        # Add new messages for "Compare One Project Over Time" mode
        'select_single_project_warning': "Please select ONLY ONE project for this mode.",
        'select_years_for_over_time_months': "Select the year(s) for comparison:",
        'select_months_for_single_year': "Select month(s) within the chosen year:",
        'comparison_over_years_note': "Note: You have selected multiple years. The report will compare the project's data across the selected years. Month selection will be ignored.",
        'comparison_over_months_note': "Note: The report will compare the project's data across the selected months in year {}.",
        'no_comparison_criteria_selected': "Please select at least one year or month for comparison.",
        'no_month_selected_for_single_year': "Please select at least one month when comparing a single project within a specific year."
    }
}

# Lấy từ điển văn bản dựa trên lựa chọn ngôn ngữ hiện tại
def get_text(key):
    return TEXTS[st.session_state.lang].get(key, f"Missing text for {key}")

# Header của ứng dụng
col_logo_title, col_lang = st.columns([0.8, 0.2])
with col_logo_title:
    st.image("triac_logo.png", width=110) # Logo cố định
    st.markdown("<div class='report-title'>Triac Time Report Generator</div>", unsafe_allow_html=True) # Tiêu đề cố định
    st.markdown("<div class='report-subtitle'>Reporting tool for time tracking and analysis</div>", unsafe_allow_html=True) # Phụ đề cố định

with col_lang:
    st.session_state.lang = st.radio(
        get_text('lang_select'),
        options=['vi', 'en'],
        format_func=lambda x: get_text('language_' + x),
        key='language_selector_main'
    )


# Check if template file exists
if not os.path.exists(path_dict['template_file']):
    st.error(get_text('template_not_found').format(path_dict['template_file']))
    st.stop()

# Load raw data and configurations once
@st.cache_data(ttl=1800)
def cached_load():
    df_raw = load_raw_data(path_dict['template_file'])
    config_data = read_configs(path_dict['template_file'])
    return df_raw, config_data

with st.spinner(get_text('loading_data')):
    df_raw, config_data = cached_load()

if df_raw.empty:
    st.error(get_text('failed_to_load_raw_data'))
    st.stop()

# Get unique years, months, and projects from raw data for selectbox options
all_years = sorted(df_raw['Year'].dropna().unique().astype(int).tolist())
month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
all_months = [m for m in month_order if m in df_raw['MonthName'].dropna().unique()]
all_projects = sorted(df_raw['Project name'].dropna().unique().tolist())


# Main interface tabs
tab_standard_report_main, tab_comparison_report_main, tab_data_preview_main, tab_user_guide_main = st.tabs([
    get_text('tab_standard_report'),
    get_text('tab_comparison_report'),
    get_text('tab_data_preview'),
    get_text('user_guide')
])

# =========================================================================
# STANDARD REPORT TAB
# =========================================================================
with tab_standard_report_main:
    st.header(get_text('standard_report_header'))

    col1_std, col2_std, col3_std = st.columns(3)
    with col1_std:
        # State management for standard analysis mode
        if 'standard_analysis_mode' not in st.session_state:
            st.session_state.standard_analysis_mode = config_data['mode'] if config_data['mode'] in ['year', 'month', 'week'] else 'year'

        mode_options = ['year', 'month', 'week']
        try:
            mode_index = mode_options.index(st.session_state.standard_analysis_mode)
        except ValueError:
            mode_index = 0
            st.session_state.standard_analysis_mode = mode_options[0]

        mode = st.selectbox(
            get_text('select_analysis_mode'),
            options=mode_options,
            index=mode_index,
            key='standard_mode_tab'
        )
        st.session_state.standard_analysis_mode = mode # Update state

    with col2_std:
        # State management for standard selected year
        if 'standard_selected_year' not in st.session_state:
            st.session_state.standard_selected_year = config_data['year'] if config_data['year'] in all_years else (all_years[0] if all_years else None)
            
        default_std_year_index = 0
        if st.session_state.standard_selected_year in all_years:
            default_std_year_index = all_years.index(st.session_state.standard_selected_year)
        elif all_years:
            st.session_state.standard_selected_year = all_years[0] # Fallback
            default_std_year_index = 0
        elif st.session_state.standard_selected_year is None: # No years available at all
            default_std_year_index = None


        selected_year = st.selectbox(
            get_text('select_year'),
            options=all_years,
            index=default_std_year_index,
            key='standard_year_tab'
        )
        st.session_state.standard_selected_year = selected_year # Update state

        if selected_year is None:
            st.warning(get_text('no_year_in_data'))

    with col3_std:
        # State management for standard selected months
        if 'standard_selected_months' not in st.session_state:
            st.session_state.standard_selected_months = config_data['months'] if config_data['months'] else all_months
            
        # Ensure default months are valid in current all_months
        valid_default_months = [m for m in st.session_state.standard_selected_months if m in all_months]
        if not valid_default_months and all_months: # Fallback if no valid default or if default is empty but options exist
            valid_default_months = all_months # Select all months as default if nothing is selected

        selected_months = st.multiselect(
            get_text('select_months'),
            options=all_months,
            default=valid_default_months,
            key='standard_months_tab'
        )
        st.session_state.standard_selected_months = selected_months # Update state


    st.subheader(get_text('standard_project_selection_header'))

    # Determine initial included projects based on config for default
    initial_included_projects_config = config_data['project_filter_df'][
        config_data['project_filter_df']['Include'].astype(str).str.lower() == 'yes'
    ]['Project Name'].tolist()

    # State management for standard project selection
    if 'standard_selected_projects' not in st.session_state:
        default_standard_projects = [p for p in initial_included_projects_config if p in all_projects]
        if not default_standard_projects and all_projects:
            default_standard_projects = all_projects # Default to all if config is empty
        st.session_state.standard_selected_projects = default_standard_projects
        
    # Ensure default value for multiselect is valid
    current_std_projects_default = [p for p in st.session_state.standard_selected_projects if p in all_projects]
    if not current_std_projects_default and all_projects: # Fallback if selected projects are no longer valid or empty
        current_std_projects_default = all_projects

    standard_project_selection = st.multiselect(
        get_text('standard_project_selection_text'),
        options=all_projects,
        default=current_std_projects_default,
        key='standard_project_selection_tab'
    )
    st.session_state.standard_selected_projects = standard_project_selection # Update state


    st.markdown("---")
    st.subheader(get_text("export_options"))
    export_excel_std = st.checkbox(get_text("export_excel_option"), value=True, key='export_excel_std')
    export_pdf_std = st.checkbox(get_text("export_pdf_option"), value=False, key='export_pdf_std')

    if st.button(get_text('generate_standard_report_btn'), key='generate_standard_report_btn_tab'):
        if not export_excel_std and not export_pdf_std:
            st.warning(get_text("warning_select_export_format"))
        elif selected_year is None:
            st.error(get_text('no_year_selected_error'))
        elif not standard_project_selection:
            st.warning(get_text('no_project_selected_warning_standard'))
        else:
            # Gọi hàm generate_reports_on_demand cho báo cáo tiêu chuẩn
            try:
                with st.spinner(get_text('generating_excel_report')): # Dùng chung spinner cho cả 2 loại báo cáo
                    standard_report_output_paths = generate_reports_on_demand(
                        df_raw=df_raw,
                        config_data=config_data, # Truyền config_data hiện có nếu cần bên trong hàm
                        selected_mode=mode,
                        selected_year=selected_year,
                        selected_months=selected_months,
                        selected_project_names_standard=standard_project_selection,
                        comparison_config_years=None, # Không áp dụng cho báo cáo tiêu chuẩn
                        comparison_config_months=None, # Không áp dụng cho báo cáo tiêu chuẩn
                        comparison_config_projects=None, # Không áp dụng cho báo cáo tiêu chuẩn
                        comparison_report_mode=None, # Không áp dụng cho báo cáo tiêu chuẩn
                        export_excel=export_excel_std,
                        export_pdf=export_pdf_std,
                        path_dict=path_dict
                    )

                if standard_report_output_paths:
                    st.success(get_text('report_done'))
                    if export_excel_std and standard_report_output_paths.get('standard_excel_path') and os.path.exists(standard_report_output_paths['standard_excel_path']):
                        with open(standard_report_output_paths['standard_excel_path'], "rb") as f:
                            st.download_button(
                                get_text("download_excel"),
                                data=f,
                                file_name=os.path.basename(standard_report_output_paths['standard_excel_path']),
                                use_container_width=True,
                                key='download_excel_std_btn'
                            )
                    if export_pdf_std and standard_report_output_paths.get('standard_pdf_path') and os.path.exists(standard_report_output_paths['standard_pdf_path']):
                        with open(standard_report_output_paths['standard_pdf_path'], "rb") as f:
                            st.download_button(
                                get_text("download_pdf"),
                                data=f,
                                file_name=os.path.basename(standard_report_output_paths['standard_pdf_path']),
                                use_container_width=True,
                                key='download_pdf_std_btn'
                            )
                else:
                    st.warning(get_text('no_data_after_filter_standard'))

            except Exception as e:
                st.error(f"{get_text('error_generating_report')}: {e}")
                st.exception(e) # Hiển thị chi tiết lỗi để debug


# =========================================================================
# COMPARISON REPORT TAB
# =========================================================================
with tab_comparison_report_main:
    st.header(get_text('comparison_report_header'))

    # Define the mapping from text key to (Vietnamese_internal_string, English_internal_string)
    # This ensures the correct internal string is passed to backend, regardless of UI language
    internal_comparison_modes_map = {
        'compare_projects_month': ("So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"),
        'compare_projects_year': ("So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"),
        'compare_one_project_over_time': ("So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)")
    }

    # Khởi tạo giá trị mặc định nếu chưa có trong session_state
    if 'selected_comparison_mode_key' not in st.session_state:
        # Mặc định chọn key đầu tiên trong danh sách
        st.session_state.selected_comparison_mode_key = list(internal_comparison_modes_map.keys())[0]

    # Tạo list các options để hiển thị trong selectbox
    # và một map để tìm key từ display text
    display_options = []
    display_to_key_map = {}
    for key in internal_comparison_modes_map.keys():
        display_text = get_text(key)
        display_options.append(display_text)
        display_to_key_map[display_text] = key

    # Lấy giá trị hiển thị mặc định dựa trên key đã lưu
    default_display_value = get_text(st.session_state.selected_comparison_mode_key)
    
    # Đảm bảo giá trị mặc định tồn tại trong display_options để tránh lỗi
    # Nếu không tìm thấy, fallback về mục đầu tiên và cập nhật session_state
    try:
        current_index = display_options.index(default_display_value)
    except ValueError:
        # Giá trị mặc định không tìm thấy trong options hiện tại, fallback về đầu tiên
        current_index = 0
        st.session_state.selected_comparison_mode_key = display_to_key_map[display_options[0]]
        default_display_value = display_options[0] # Cập nhật lại default_display_value cho đúng

    selected_comparison_display = st.selectbox(
        get_text('select_comparison_mode'),
        options=display_options,
        index=current_index, # Đặt index dựa trên giá trị mặc định đã được kiểm tra
        key='comparison_mode_select_tab_main'
    )
    
    # Cập nhật key lựa chọn vào session_state khi người dùng thay đổi
    current_selected_key = display_to_key_map[selected_comparison_display]
    if st.session_state.selected_comparison_mode_key != current_selected_key:
        st.session_state.selected_comparison_mode_key = current_selected_key


    # Lấy giá trị chuỗi nội bộ (internal string) để truyền vào backend
    # Dựa trên key đã lưu và ngôn ngữ hiện tại
    vi_val, en_val = internal_comparison_modes_map[st.session_state.selected_comparison_mode_key]
    if st.session_state.lang == 'vi':
        comparison_mode = vi_val
    else: # 'en'
        comparison_mode = en_val

    st.subheader(get_text('filter_data_for_comparison'))

    comp_years = []
    comp_months = []
    comp_projects = []
    validation_error = False # Flag to check input errors

    # State management for comparison projects
    if 'comparison_selected_projects' not in st.session_state:
        st.session_state.comparison_selected_projects = [] # Default to empty

    comp_projects = st.multiselect(
        get_text('select_projects_comp'),
        options=all_projects,
        default=[p for p in st.session_state.comparison_selected_projects if p in all_projects], # Ensure default is valid
        key='comp_projects_select_tab_common'
    )
    st.session_state.comparison_selected_projects = comp_projects # Update state


    if comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm" or comparison_mode == "Compare One Project Over Time (Months/Years)":
        if len(comp_projects) != 1:
            st.warning(get_text('select_single_project_warning'))
            validation_error = True

        # State management for selected years in "Over Time" mode
        if 'comparison_selected_years_over_time' not in st.session_state:
            st.session_state.comparison_selected_years_over_time = []

        selected_years_over_time = st.multiselect(
            get_text('select_years_for_over_time_months'),
            options=all_years,
            default=[y for y in st.session_state.comparison_selected_years_over_time if y in all_years], # Ensure default is valid
            key='comp_years_select_tab_over_time'
        )
        st.session_state.comparison_selected_years_over_time = selected_years_over_time # Update state
        comp_years = selected_years_over_time # Assign to comp_years for config

        # State management for selected months in "Over Time" mode (if single year selected)
        if 'comparison_selected_months_over_time' not in st.session_state:
            st.session_state.comparison_selected_months_over_time = []


        if len(selected_years_over_time) == 1:
            st.info(get_text('comparison_over_months_note').format(selected_years_over_time[0]))
            comp_months = st.multiselect(
                get_text('select_months_for_single_year'),
                options=all_months,
                default=[m for m in st.session_state.comparison_selected_months_over_time if m in all_months], # Ensure default is valid
                key='comp_months_select_tab_over_time'
            )
            st.session_state.comparison_selected_months_over_time = comp_months # Update state

            if not comp_months:
                st.warning(get_text('no_month_selected_for_single_year'))
                validation_error = True

        elif len(selected_years_over_time) > 1:
            st.info(get_text('comparison_over_years_note'))
            comp_months = [] # Months are ignored for multi-year comparison
            st.session_state.comparison_selected_months_over_time = [] # Clear months state
        else:
            st.warning(get_text('no_comparison_criteria_selected'))
            validation_error = True
            comp_months = [] # Ensure empty
            st.session_state.comparison_selected_months_over_time = [] # Clear months state

    elif comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month", "So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        col_comp1, col_comp2 = st.columns(2)
        with col_comp1:
            # State management for general comparison years
            if 'comparison_selected_years_general' not in st.session_state:
                st.session_state.comparison_selected_years_general = []

            comp_years = st.multiselect(
                get_text('select_years'),
                options=all_years,
                default=[y for y in st.session_state.comparison_selected_years_general if y in all_years], # Ensure default is valid
                key='comp_years_select_tab_general'
            )
            st.session_state.comparison_selected_years_general = comp_years # Update state

        with col_comp2:
            # State management for general comparison months
            if 'comparison_selected_months_general' not in st.session_state:
                st.session_state.comparison_selected_months_general = []

            if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
                comp_months = st.multiselect(
                    get_text('select_months_comp'),
                    options=all_months,
                    default=[m for m in st.session_state.comparison_selected_months_general if m in all_months], # Ensure default is valid
                    key='comp_months_select_tab_general'
                )
                st.session_state.comparison_selected_months_general = comp_months # Update state
            else:
                # Clear months selection if mode changes
                comp_months = []
                st.session_state.comparison_selected_months_general = []

        # Validation for "Compare Projects in a Month"
        if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
            if len(comp_years) != 1:
                st.warning(get_text('no_year_selected_error'))
                validation_error = True
            if len(comp_months) != 1:
                st.warning("Vui lòng chọn CHỈ MỘT tháng cho chế độ 'So Sánh Dự Án Trong Một Tháng'.")
                validation_error = True
            if not comp_projects:
                st.warning(get_text('no_project_selected_warning_standard'))
                validation_error = True
        # Validation for "Compare Projects in a Year"
        elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
            if len(comp_years) != 1:
                st.warning(get_text('no_year_selected_error'))
                validation_error = True
            if not comp_projects:
                st.warning(get_text('no_project_selected_warning_standard'))
                validation_error = True


    st.markdown("---")
    st.subheader(get_text("export_options"))
    export_excel_comp = st.checkbox(get_text("export_excel_option"), value=True, key='export_excel_comp')
    export_pdf_comp = st.checkbox(get_text("export_pdf_option"), value=False, key='export_pdf_comp')

    if st.button(get_text('generate_comparison_report_btn'), key='generate_comparison_report_btn_tab'):
        if not export_excel_comp and not export_pdf_comp:
            st.warning(get_text("warning_select_export_format"))
        elif validation_error: # Check if any validation error occurred above
            st.error(get_text('error_generating_report')) # General error message for user input
        else:
            # Gọi hàm generate_reports_on_demand cho báo cáo so sánh
            try:
                with st.spinner(get_text('generating_comparison_excel')): # Dùng chung spinner
                    comparison_report_output_paths = generate_reports_on_demand(
                        df_raw=df_raw,
                        config_data=config_data,
                        selected_mode=None, # Không áp dụng cho báo cáo so sánh
                        selected_year=None, # Không áp dụng cho báo cáo so sánh
                        selected_months=None, # Không áp dụng cho báo cáo so sánh
                        selected_project_names_standard=None, # Không áp dụng cho báo cáo so sánh
                        comparison_config_years=comp_years,
                        comparison_config_months=comp_months,
                        comparison_config_projects=comp_projects,
                        comparison_report_mode=comparison_mode,
                        export_excel=export_excel_comp,
                        export_pdf=export_pdf_comp,
                        path_dict=path_dict
                    )

                if comparison_report_output_paths:
                    st.success(get_text('report_done'))
                    if export_excel_comp and comparison_report_output_paths.get('comparison_excel_path') and os.path.exists(comparison_report_output_paths['comparison_excel_path']):
                        with open(comparison_report_output_paths['comparison_excel_path'], "rb") as f:
                            st.download_button(
                                get_text("download_comparison_excel"),
                                data=f,
                                file_name=os.path.basename(comparison_report_output_paths['comparison_excel_path']),
                                use_container_width=True,
                                key='download_comparison_excel_btn'
                            )
                    if export_pdf_comp and comparison_report_output_paths.get('comparison_pdf_path') and os.path.exists(comparison_report_output_paths['comparison_pdf_path']):
                        with open(comparison_report_output_paths['comparison_pdf_path'], "rb") as f:
                            st.download_button(
                                get_text("download_comparison_pdf"),
                                data=f,
                                file_name=os.path.basename(comparison_report_output_paths['comparison_pdf_path']),
                                use_container_width=True,
                                key='download_comparison_pdf_btn'
                            )
                else:
                    st.warning(get_text('no_data_after_filter_comparison').format("Không có dữ liệu sau khi lọc cho báo cáo so sánh."))

            except Exception as e:
                st.error(f"{get_text('error_generating_report')}: {e}")
                st.exception(e) # Hiển thị chi tiết lỗi để debug

# =========================================================================
# DATA PREVIEW TAB
# =========================================================================
with tab_data_preview_main:
    st.header(get_text('raw_data_preview_header'))
    if not df_raw.empty:
        st.dataframe(df_raw.head(100))
    else:
        st.info(get_text('no_raw_data'))

# =========================================================================
# USER GUIDE TAB
# =========================================================================
with tab_user_guide_main:
    st.header(get_text('user_guide'))
    st.markdown("""
    **Chào mừng bạn đến với Công cụ tạo báo cáo thời gian Triac!**

    Công cụ này giúp bạn tạo ra các báo cáo thời gian chi tiết và so sánh từ dữ liệu thô của mình.

    ### 1. Báo cáo tiêu chuẩn
    * **Chế độ phân tích:** Chọn cách bạn muốn phân tích dữ liệu (theo năm, tháng hoặc tuần).
    * **Chọn năm/tháng:** Lọc dữ liệu theo năm và tháng cụ thể.
    * **Lựa chọn dự án:** Chọn các dự án bạn muốn bao gồm trong báo cáo. Theo mặc định, công cụ sẽ chỉ bao gồm các dự án được đánh dấu 'yes' trong sheet 'Config_Project_Filter' của file template. Bạn có thể thay đổi lựa chọn này tại đây.
    * **Tạo báo cáo:** Nhấn nút 'Tạo báo cáo tiêu chuẩn' để tạo file Excel và/hoặc PDF.

    ### 2. Báo cáo so sánh
    * **Chế độ so sánh:**
        * **So Sánh Dự Án Trong Một Tháng:** So sánh hiệu suất của nhiều dự án trong một tháng cụ thể. Chọn một năm và một tháng, cùng với các dự án muốn so sánh.
        * **So Sánh Dự Án Trong Một Năm:** So sánh hiệu suất của nhiều dự án trong một năm cụ thể. Chọn một năm và các dự án muốn so sánh. Lựa chọn tháng sẽ bị bỏ qua.
        * **So Sánh Một Dự Án Qua Các Tháng/Năm:** So sánh hiệu suất của MỘT dự án duy nhất qua nhiều tháng trong cùng một năm, HOẶC so sánh qua nhiều năm.
            * Nếu bạn chọn **một năm và nhiều tháng**: Báo cáo sẽ so sánh dự án đó qua các tháng đã chọn trong năm đó.
            * Nếu bạn chọn **nhiều năm**: Báo cáo sẽ so sánh dự án đó qua các năm đã chọn. Lựa chọn tháng sẽ bị bỏ qua.
    * **Tạo báo cáo:** Nhấn nút 'Tạo báo cáo so sánh' để tạo file Excel và/hoặc PDF.

    ### 3. Xem trước dữ liệu
    Tab này cho phép bạn xem 100 hàng đầu tiên của dữ liệu thô đã tải, giúp bạn kiểm tra định dạng và nội dung dữ liệu.

    ### 4. Cấu hình file template (Bên ngoài ứng dụng)
    Công cụ đọc dữ liệu và cấu hình từ một file Excel template (thường là `Timesheet_Template.xlsx`). Đảm bảo rằng:
    * Sheet 'Raw Data' chứa dữ liệu thời gian thô của bạn với các cột cần thiết như 'Year', 'MonthName', 'Project name', v.v.
    * Sheet 'Config_Year_Mode' và 'Config_Project_Filter' có thể được sử dụng để đặt cấu hình mặc định, nhưng các lựa chọn trên giao diện sẽ ghi đè lên chúng.

    ### Lỗi thường gặp:
    * **File template không tìm thấy:** Đảm bảo `Timesheet_Template.xlsx` nằm cùng thư mục với ứng dụng này.
    * **Không tải được dữ liệu thô:** Kiểm tra định dạng dữ liệu trong sheet 'Raw Data' của template. Đảm bảo các cột như 'Year', 'MonthName', 'Project name' tồn tại và có định dạng hợp lệ.

    ---
    **Chúc bạn sử dụng hiệu quả!**
    """, unsafe_allow_html=True)
