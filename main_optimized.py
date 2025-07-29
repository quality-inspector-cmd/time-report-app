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
    apply_filters, export_report, export_pdf_report,
    apply_comparison_filters, export_comparison_report, export_comparison_pdf_report,send_email_to_admin
)
# ==============================================================================

script_dir = os.path.dirname(__file__)
csv_file_path = os.path.join(script_dir, "invited_emails.csv")

# Gọi hàm setup_paths ngay từ đầu để path_dict có sẵn
path_dict = setup_paths()

# ==============================================================================
# KHỞI TẠO CÁC BIẾN TRẠNG THÁI PHIÊN (SESSION STATE VARIABLES)
# ==============================================================================
if 'comparison_mode_select_tab_main' not in st.session_state:
    st.session_state.comparison_mode_select_tab_main = "Compare Projects in a Month"  # hoặc giá trị phù hợp với `display_options`
    
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
        'compare_projects_over_time': "Compare Projects Over Time (Months/Years)",
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
        'select_at_least_two_projects_warning': "Please select at least two projects for comparison.",
        'select_years_for_over_time_months': "Select the year(s) for comparison:",
        'select_months_for_single_year': "Select month(s) within the chosen year:",
        'comparison_over_years_note': "Note: You have selected multiple years. The report will compare the project's data across the selected years. Month selection will be ignored.",
        'comparison_over_months_note': "Note: The report will compare the project's data across the selected months in year {}.",
        'no_comparison_criteria_selected': "Please select at least one year or month for comparison.",
        'no_month_selected_for_single_year': "Please select at least one month when comparing a single project within a specific year.",
        'tab_help': "🔧 Need Help?",
        'help_tab_description': "If you experience issues or have questions, please describe your problem below. The system will notify the administrator.",
        'your_email': "Your email",
        'describe_issue': "Describe the issue you're facing",
        "help_title": ("Gửi Yêu Cầu Trợ Giúp", "Submit Help Request"),
        "help_instruction": "Please enter your email to receive detailed instructions.",
        "help_input_label": "Enter your email",
        "help_input_placeholder": ("Nhập mô tả sự cố ở đây...", "Type your issue here..."),
        "help_submit_button": ("Gửi yêu cầu", "Submit request"),
        "help_submit_success": ("Yêu cầu đã được gửi thành công!", "Your request has been successfully sent!"),
        "help_submit_fail": ("Không thể gửi yêu cầu. Vui lòng thử lại.", "Failed to send the request. Please try again."),
        "help_submit_warning": ("Vui lòng nhập nội dung trước khi gửi.", "Please enter some text before submitting."),
        'send_help_request_button': "Send Help Request",
        'issue_required_warning': "Please enter the issue description.",
        'email_required_warning': "Please enter your email.",
        'email_sent_success': "✅ Request sent successfully. Admin will respond soon.",
        'email_sent_error': "❌ Failed to send your help request",
        'select_all_projects_checkbox': "Select all projects"
    },
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
        'compare_projects_over_time': "So Sánh Nhiều Dự Án Qua Các Tháng/Năm",
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
        # Thêm các tin nhắn mới cho mode "So Sánh Các Dự Án Qua Các Tháng/Năm"
        'select_at_least_two_projects_warning': "Vui lòng chọn ít nhất hai dự án để so sánh.",
        'select_years_for_over_time_months': "Chọn năm (hoặc các năm) bạn muốn so sánh:",
        'select_months_for_single_year': "Chọn tháng(các tháng) trong năm đã chọn:",
        'comparison_over_years_note': "Lưu ý: Bạn đã chọn nhiều năm. Báo cáo sẽ so sánh dữ liệu của dự án qua các năm đã chọn. Lựa chọn tháng sẽ bị bỏ qua.",
        'comparison_over_months_note': "Lưu ý: Báo cáo sẽ so sánh dữ liệu của dự án qua các tháng đã chọn trong năm {}.",
        'no_comparison_criteria_selected': "Vui lòng chọn ít nhất một năm hoặc một tháng để so sánh.",
        'no_month_selected_for_single_year': "Vui lòng chọn ít nhất một tháng khi so sánh một dự án trong một năm cụ thể.",
        'tab_help': "🔧 Cần trợ giúp?",
        "help_title": "Gửi Yêu Cầu Trợ Giúp",
        'help_tab_description': "Nếu bạn gặp sự cố hoặc có thắc mắc, vui lòng mô tả vấn đề bên dưới. Hệ thống sẽ thông báo cho quản trị viên.",
        'your_email': "Email của bạn",
        "help_instruction": "Vui lòng nhập email để nhận hướng dẫn chi tiết qua email.",
        'help_input_label': "Nhập nội dung cần trợ giúp",
        'help_input_placeholder': "Nhập mô tả sự cố ở đây...",
        'describe_issue': "Mô tả vấn đề bạn gặp phải",
        'help_submit_button': "Gửi yêu cầu",
        'issue_required_warning': "Vui lòng nhập mô tả vấn đề.",
        'email_required_warning': "Vui lòng nhập email của bạn.",
        'email_sent_success': "✅ Đã gửi yêu cầu thành công. Quản trị viên sẽ phản hồi sớm.",
        'email_sent_error': "❌ Gửi yêu cầu trợ giúp thất bại",
        'select_all_projects_checkbox': "Chọn tất cả dự án"
    }
}

# Lấy từ điển văn bản dựa trên lựa chọn ngôn ngữ hiện tại
def get_text(key, lang=None):
    lang = lang or st.session_state.get("lang", "vi")
    val = TEXTS.get(lang, {}).get(key)

    if val is None:
        return f"Missing text for {key}"

    # ✅ Nếu là tuple, chọn theo lang
    if isinstance(val, tuple):
        return val[0] if lang == 'vi' else val[1]

    return val

# Header của ứng dụng
col_logo_title, col_lang = st.columns([0.8, 0.2])
with col_logo_title:
    st.image("triac_logo.png", width=110) # Logo cố định
    st.markdown("<div class='report-title'>Triac Time Report Generator</div>", unsafe_allow_html=True) # Tiêu đề cố định
    st.markdown("<div class='report-subtitle'>Reporting tool for time tracking and analysis</div>", unsafe_allow_html=True) # Phụ đề cố định

with col_lang:
    selected_lang = st.radio(
        "Select language:",
        options=['vi', 'en'],
        format_func=lambda x: "Tiếng Việt" if x == "vi" else "English",
        key='language_selector_main'
    )
    if st.session_state.lang != selected_lang:
        st.session_state.lang = selected_lang
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
tab_standard_report_main, tab_comparison_report_main, tab_data_preview_main, tab_user_guide_main, tab_help_main = st.tabs([
    get_text('tab_standard_report'),
    get_text('tab_comparison_report'),
    get_text('tab_data_preview'),
    get_text('user_guide'),
    get_text("tab_help")   # <-- phải đúng key trong text_dict
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

# 🟩 Hỗ trợ chọn tất cả dự án
    select_all_std_projects = st.checkbox(get_text("select_all_projects_checkbox"), value=True, key="select_all_std_projects_checkbox")

    if select_all_std_projects:
        standard_project_selection = all_projects
    else:
        current_std_projects_default = [p for p in st.session_state.standard_selected_projects if p in all_projects]
        if not current_std_projects_default and all_projects:
            current_std_projects_default = all_projects
        # ✅ Chèn dòng hiển thị số lượng đang chọn
        st.caption(f"Đang chọn {len(current_std_projects_default)} dự án")
        
        standard_project_selection = st.multiselect(
            get_text('standard_project_selection_text'),
            options=all_projects,
            default=current_std_projects_default,
            key='standard_project_selection_tab'
        )
        # ✅ CHỈ cập nhật nếu có thay đổi → tránh Streamlit reload vô ích
    if st.session_state.standard_selected_projects != standard_project_selection:
        st.session_state.standard_selected_projects = standard_project_selection
        
    st.markdown("---")
    st.subheader(get_text("export_options"))
    export_excel = st.checkbox(get_text("export_excel_option"), value=True, key='export_excel_std')
    export_pdf = st.checkbox(get_text("export_pdf_option"), value=False, key='export_pdf_std')

    if st.button(get_text('generate_standard_report_btn'), key='generate_standard_report_btn_tab'):
        if not export_excel and not export_pdf:
            st.warning(get_text("warning_select_export_format"))
        elif selected_year is None:
            st.error(get_text('no_year_selected_error'))
        elif not standard_project_selection:
            st.warning(get_text('no_project_selected_warning_standard'))
        else:
            temp_project_filter_df_standard = pd.DataFrame({
                'Project Name': standard_project_selection,
                'Include': ['yes'] * len(standard_project_selection)
            })

            standard_report_config = {
                'mode': mode,
                'year': selected_year,
                'months': selected_months,
                'project_filter_df': temp_project_filter_df_standard
            }

            df_filtered_standard = apply_filters(df_raw, standard_report_config)
            # Tự động loại bỏ dự án không có dữ liệu sau khi lọc
            project_col = 'Project name'  # <-- Đúng tên cột trong df_raw, sửa nếu cần
            valid_projects_in_filtered = df_filtered_standard[project_col].unique().tolist()

            # Giữ lại các dự án có dữ liệu
            standard_project_selection = [p for p in standard_project_selection if p in valid_projects_in_filtered]

            # Nếu không còn dự án nào hợp lệ, cảnh báo và dừng
            if not standard_project_selection:
                st.warning("Không có dự án nào có dữ liệu trong năm và tháng đã chọn.")
                st.stop()

            # Cập nhật lại config và project_filter_df
            temp_project_filter_df_standard = pd.DataFrame({
                'Project Name': standard_project_selection,
                'Include': ['yes'] * len(standard_project_selection)
            })
            standard_report_config['project_filter_df'] = temp_project_filter_df_standard
            if 'Date' in df_filtered_standard.columns:
                df_filtered_standard['MonthName'] = pd.to_datetime(df_filtered_standard['Date']).dt.strftime('%B')
            if df_filtered_standard.empty:
                st.warning(get_text('no_data_after_filter_standard'))
            else:
                today_str = datetime.today().strftime("%Y-%m-%d")  # ✅ Đúng cú pháp
                path_dict = {                                        # ✅ Bổ sung cần thiết
                'output_file': f'outputs/standard/Time_report_Standard_{today_str}.xlsx',
                'pdf_report': f'outputs/standard/Time_report_Standard_{today_str}.pdf',
                'logo_path': 'triac_logo.png'
                } 
                report_generated = False
                if export_excel:
                    with st.spinner(get_text('generating_excel_report')):
                        excel_success = export_report(df_filtered_standard, standard_report_config, path_dict['output_file'])
                    if excel_success:
                        st.success(get_text('excel_report_generated').format(os.path.basename(path_dict['output_file'])))
                        report_generated = True
                    else:
                        st.error(get_text('failed_to_generate_excel'))
                if export_pdf:
                    pdf_report_path = path_dict['pdf_report']  # ✅ thêm dòng này trước khi dùng biến
                    # ✅ Kiểm tra trước khi gọi
                    if not pdf_report_path:
                        raise ValueError("❌ pdf_report_path is empty. Please check where it's defined.")
                    with st.spinner(get_text('generating_pdf_report')):
                        print(f"[DEBUG] path_dict['pdf_report'] = {path_dict['pdf_report']}")
                        pdf_success = export_pdf_report(df_filtered_standard, standard_report_config, path_dict['pdf_report'], path_dict['logo_path'])
                    if pdf_success:
                        st.success(get_text('pdf_report_generated').format(os.path.basename(path_dict['pdf_report'])))
                        report_generated = True
                    else:
                        st.error(get_text('failed_to_generate_pdf'))

                if report_generated:
                    if export_excel and os.path.exists(path_dict['output_file']):
                        with open(path_dict['output_file'], "rb") as f:
                            st.download_button(get_text("download_excel"), data=f, file_name=os.path.basename(path_dict['output_file']), use_container_width=True, key='download_excel_std_btn')
                    if export_pdf and os.path.exists(path_dict['pdf_report']):
                        with open(path_dict['pdf_report'], "rb") as f:
                            st.download_button(get_text("download_pdf"), data=f, file_name=os.path.basename(path_dict['pdf_report']), use_container_width=True, key='download_pdf_std_btn')
                else:
                    st.error(get_text('error_generating_report'))


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
        'compare_projects_over_time': ("So Sánh Nhiều Dự Án Qua Các Tháng/Năm", "Compare Projects Over Time (Months/Years)")
    }
    # Tạo display options và ánh xạ ngược lại
    display_options = [get_text(k) for k in internal_comparison_modes_map]
    display_to_key_map = {get_text(k): k for k in internal_comparison_modes_map}

    # Lấy giá trị mặc định từ session
    default_key = st.session_state.get('selected_comparison_mode_key', list(internal_comparison_modes_map.keys())[0])
    default_display = get_text(default_key)

    try:
        current_index = display_options.index(default_display)
    except ValueError:
        # Giá trị mặc định không tìm thấy trong options hiện tại, fallback về đầu tiên
        current_index = 0
        default_key = list(internal_comparison_modes_map.keys())[0]
        st.session_state.selected_comparison_mode_key = default_key
        default_display = get_text(default_key)  # cập nhật lại display nếu fallback
    # Hiển thị selectbox (dùng chính session key để giữ đồng bộ)
    selected_display = st.selectbox(
        get_text('select_comparison_mode'),
        options=display_options,
        index=current_index,
        key='selected_comparison_display_key'  # 👈 đổi key khác
    )
    # Ánh xạ ngược lại key gốc
    selected_key = display_to_key_map[selected_display]
    st.session_state.selected_comparison_mode_key = selected_key
    # Lấy giá trị chuỗi nội bộ (internal string) để truyền vào backend
    # Dựa trên key đã lưu và ngôn ngữ hiện tại
    vi_val, en_val = internal_comparison_modes_map[selected_key]
    comparison_mode = vi_val if st.session_state.lang == 'vi' else en_val
    
    st.subheader(get_text('filter_data_for_comparison'))
    # Display options và mapping display → internal
    if st.session_state.lang == 'vi':
        filter_mode_display_options = ["Theo Tổng Giờ", "Theo Task", "Theo Workcentre"]
    else:
        filter_mode_display_options = ["By Total hour", "By Task", "By Workcentre"]

    # ✅ Map display string → internal string
    display_to_internal = {
        "Theo Tổng Giờ": "Total",
        "Theo Task": "Task",
        "Theo Workcentre": "Workcentre",
        "By Total hour": "Total",
        "By Task": "Task",
        "By Workcentre": "Workcentre"
    }

    # Lấy giá trị hiện tại từ session hoặc mặc định
    current_display = st.session_state.get("selected_filter_display", filter_mode_display_options[0])
    
    # ✅ Fallback nếu không hợp lệ sau khi đổi ngôn ngữ
    if current_display not in filter_mode_display_options:
        current_display = filter_mode_display_options[0]
        st.session_state.selected_filter_display = current_display
        st.session_state.selected_filter_mode = display_to_internal[current_display]

    # Hiển thị selectbox
    selected_filter_display = st.selectbox(
        "Comparison filter mode",
        options=filter_mode_display_options,
        index=filter_mode_display_options.index(current_display),
        key="filter_mode_selectbox"
    )

    # Nếu người dùng thay đổi lựa chọn
    if selected_filter_display != current_display:
        st.session_state.selected_filter_display = selected_filter_display
        st.session_state.selected_filter_mode = display_to_internal[selected_filter_display]
        st.rerun()

      # ✅ Luôn lấy filter_mode (chuẩn hóa) từ session
    filter_mode = st.session_state.get("selected_filter_mode", display_to_internal[current_display])

    # State management for comparison projects
    if 'comparison_selected_projects' not in st.session_state:
        st.session_state.comparison_selected_projects = [] # Default to empty
    # Đặt ở đây, trước khi bắt đầu kiểm tra từng chế độ
    validation_error = False

    select_all_projects = st.checkbox(
    get_text("select_all_projects_checkbox"),
    value=True,
    key="select_all_projects_checkbox"
    )
    if select_all_projects:
        comp_projects = all_projects
    else:
        comp_projects = st.multiselect(
            get_text('select_projects_comp'),
            options=all_projects,
            default=[p for p in st.session_state.comparison_selected_projects if p in all_projects],
            key='comp_projects_select_tab_common'
        )
    st.session_state.comparison_selected_projects = comp_projects # Update state


    if comparison_mode == "So Sánh Nhiều Dự Án Qua Các Tháng/Năm" or comparison_mode == "Compare Projects Over Time (Months/Years)":
        if len(comp_projects) < 1:
            st.warning(get_text('no_project_selected_warning_standard'))  # báo lỗi nếu không chọn gì
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
                default=[y for y in st.session_state.comparison_selected_years_general if y in all_years],
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
                    default=[m for m in st.session_state.comparison_selected_months_general if m in all_months],
                    key='comp_months_select_tab_general'
                )
                st.session_state.comparison_selected_months_general = comp_months # Update state
            else:
                comp_months = [] # Months are not relevant for yearly comparison
                st.session_state.comparison_selected_months_general = [] # Clear months state

        if not comp_years:
            st.warning(get_text('no_comparison_criteria_selected'))
            validation_error = True
        
        if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"] and not comp_months:
            st.warning(get_text('no_comparison_criteria_selected'))
            validation_error = True

        if not comp_projects:
            st.warning(get_text('no_project_selected_warning_standard')) # Reusing standard report message
            validation_error = True
            

    st.markdown("---")
    st.subheader(get_text("export_options"))
    export_excel_comp = st.checkbox(get_text("export_excel_option"), value=True, key='export_excel_comp')
    export_pdf_comp = st.checkbox(get_text("export_pdf_option"), value=False, key='export_pdf_comp')

    if st.button(get_text('generate_comparison_report_btn'), key='generate_comparison_report_btn_tab'):
        if not export_excel_comp and not export_pdf_comp:
            st.warning(get_text("warning_select_export_format"))
        elif validation_error:
            # Error messages already displayed by specific conditions
            pass
        else:
            # DEBUG print statements (giữ lại để chẩn đoán vấn đề dự án)
            print(f"DEBUG: Comparison Mode selected before filter: {comparison_mode}")
            print(f"DEBUG: Selected Projects before filter: {comp_projects}")
            print(f"DEBUG: Selected Years before filter: {comp_years}")
            print(f"DEBUG: Selected Months before filter: {comp_months}")


            comparison_config = {
                'years': comp_years,
                'months': comp_months,
                'selected_projects': comp_projects,
                'filter_mode': filter_mode   # ✅ THÊM DÒNG NÀY
                # 'selected_months_over_time' không cần truyền riêng nếu đã gán vào comp_months
                # nó đã được xử lý trong logic trên
            }
            print("✅ DEBUG - comparison_config:", comparison_config)
            # Print the final config before calling the function
            comparison_output_folder = "outputs/comparison"
            comparison_path_dict = path_dict.copy()  # path_dict được định nghĩa trước đó ở đầu chương trình
            # Thêm các key cho báo cáo so sánh
            comparison_path_dict.update({
                "comparison_output_excel": os.path.join(comparison_output_folder, "comparison_result.xlsx"),
                "comparison_output_file": os.path.join(comparison_output_folder, "comparison_export.xlsx"),
                "comparison_pdf_output": os.path.join(comparison_output_folder, "comparison_chart.png"),
                "comparison_pdf_report": os.path.join(comparison_output_folder, "comparison_report.pdf"),
                "logo": path_dict["logo_path"]  # ✅ đảm bảo tồn tại
            })
            print(f"DEBUG: Final comparison_config sent to filter: {comparison_config}")
            print(f"DEBUG: comparison_path_dict = {comparison_path_dict}")
            # ✅ Thêm dòng này sau khi path_dict đã tạo
            # Áp dụng filter
            df_filtered_comparison, comparison_filter_message, filtered_projects = apply_comparison_filters(
            df_raw, comparison_config, comparison_mode, filter_mode
            )
            # ✅ Cảnh báo nếu có dự án được chọn nhưng không có dữ liệu thực tế
            original_projects = comparison_config.get("selected_projects", [])
            if len(filtered_projects) < len(original_projects):
                removed = set(original_projects) - set(filtered_projects)
                st.warning(f"⚠️ Một số dự án không có dữ liệu thực tế và đã bị loại khỏi báo cáo: {', '.join(removed)}")

            if df_filtered_comparison.empty:
                # Đảm bảo thư mục chứa file output tồn tại
                for key in ["comparison_output_excel", "comparison_pdf_output", "comparison_output_file", "comparison_pdf_report"]:
                    folder = os.path.dirname(comparison_path_dict[key])
                    if folder:
                        os.makedirs(folder, exist_ok=True)

                st.warning(get_text('no_data_after_filter_comparison').format(comparison_filter_message))
            else:
                st.success(get_text('data_filtered_success'))
                st.subheader(get_text('comparison_data_preview'))
                st.dataframe(df_filtered_comparison)

                report_generated_comp = False
                if export_excel_comp:
                    with st.spinner(get_text('generating_comparison_excel')):
                        try:
                            excel_success_comp = export_comparison_report(
                                df_filtered_comparison,
                                comparison_config,
                                comparison_path_dict['comparison_output_file'],
                                comparison_mode,
                                filter_mode
                                )
                        except Exception as e:
                            excel_success_comp = False
                            st.error(f"❌ Lỗi khi xuất Excel: {e}")
                    # ✅ Kiểm tra file có thực sự được tạo ra không
                    if os.path.exists(comparison_path_dict['comparison_output_file']):
                        st.success("✅ File Excel đã được tạo đúng tại: " + comparison_path_dict['comparison_output_file'])
                        report_generated_comp = True
                    else:
                        st.error("❌ File Excel KHÔNG được tạo ra: " + comparison_path_dict['comparison_output_file'])
                        st.code("Current working directory: " + os.getcwd(), language="text")
                        st.code("Expected path: " + os.path.abspath(comparison_path_dict['comparison_output_file']), language="text")
                    if excel_success_comp:
                        st.success(get_text('comparison_excel_generated').format(os.path.basename(comparison_path_dict['comparison_output_file'])))
                        report_generated_comp = True
                    else:
                        st.error(get_text('failed_to_generate_comparison_excel'))

                if export_pdf_comp:
                    with st.spinner(get_text('generating_comparison_pdf')):
                        try:
                            pdf_path = comparison_path_dict['comparison_pdf_report']
                            print("▶️ Gọi export_comparison_pdf_report...")
                            pdf_success_comp = export_comparison_pdf_report(
                                df_filtered_comparison,
                                comparison_config,
                                pdf_path,
                                comparison_mode,
                                comparison_path_dict['logo'],                   # ✅ thêm logo_path
                                filter_mode
                            )
                            print("✅ PDF Success?", pdf_success_comp)
                            print("📁 File tồn tại?", os.path.exists(pdf_path))
                        except Exception as e:
                            pdf_success_comp = False
                            st.error(f"❌ Lỗi khi xuất PDF: {e}")
                            print("❌ Exception khi xuất PDF:", e)
                    if pdf_success_comp:
                        st.success(get_text('comparison_pdf_generated').format(os.path.basename(comparison_path_dict['comparison_pdf_report'])))
                        report_generated_comp = True
                    else:
                        st.error(get_text('failed_to_generate_comparison_pdf'))
                        st.warning(f"⚠️ PDF không được tạo tại: {pdf_path}")
                
                if report_generated_comp:
                # ======= HIỆN NÚT TẢI PDF/EXCEL SAU KHI XUẤT =========
                    with st.expander("📥 Tải báo cáo PDF/Excel so sánh"):
                        st.write("🪵 DEBUG path dict:", comparison_path_dict)

                        excel_path = comparison_path_dict.get("comparison_output_file")
                        pdf_path = comparison_path_dict.get("comparison_pdf_report")
                        # ⬇️ Tải Excel
                        if export_excel_comp and excel_path and os.path.exists(excel_path):
                            with open(excel_path, "rb") as f_excel:
                                excel_data = f_excel.read()  # ✅ đọc nội dung
                            st.download_button(
                                label="📄 Tải Excel So sánh",
                                data=excel_data,
                                file_name=os.path.basename(comparison_path_dict["comparison_output_file"]),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key="exp_excel_comp_btn"
                            )
                        else:
                            st.warning(f"⚠️ File Excel không tồn tại: {excel_path}")
                        # ⬇️ Tải PDF
                        if export_pdf_comp and pdf_path and os.path.exists(pdf_path):
                            with open(pdf_path, "rb") as f_pdf:
                                pdf_data = f_pdf.read()  # ✅ đọc nội dung
                            st.download_button(
                                label="🖨️ Tải PDF So sánh",
                                data=pdf_data,
                                file_name=os.path.basename(comparison_path_dict["comparison_pdf_report"]),
                                mime="application/pdf",
                                use_container_width=True,
                                key="exp_pdf_comp_btn"        
                            )
                        else:
                            st.warning(f"⚠️ File PDF không tồn tại: {pdf_path}")
                else:
                    st.error(get_text("⚠️ error_generating_report"))
# =========================================================================
# DATA PREVIEW TAB
# =========================================================================
with tab_data_preview_main:
    st.subheader(get_text('raw_data_preview_header'))
    if not df_raw.empty:
        st.dataframe(df_raw.head(100))
    else:
        st.info(get_text('no_raw_data'))

# =========================================================================
# USER GUIDE TAB
# =========================================================================
with tab_user_guide_main:
    st.markdown(f"### {get_text('user_guide')}")
    st.markdown("""
    - Select filters: mode, year, month, project
    - Select report export format (Excel, PDF or both)
    - Click "Create report"
    - Download generated report
    """)

    # 👇 Thêm đoạn này để xem access log (nếu có)
    if "access_log" in st.session_state:
        st.write("📜 Current session access log:")
        st.dataframe(pd.DataFrame(st.session_state.access_log))
# HELP TAB
# =========================================================================
with tab_help_main:
    lang = st.session_state.get("lang", "vi")

    st.markdown(f"### {get_text('help_title', lang)}")
    st.markdown(get_text("help_instruction", lang))

    user_issue = st.text_area(
        label=get_text("describe_issue", lang),
        placeholder=get_text("help_input_placeholder", lang),
        key="help_user_input"
    )

    if st.button(get_text("help_submit_button", lang)):
        if user_issue.strip():
            user_email = st.session_state.get('user_email', 'unknown@triaccomposites.com')
            sent = send_email_via_emailjs(user_issue, user_email)
            
            if sent:
                st.success(get_text("help_submit_success", lang))
            else:
                st.error(get_text("help_submit_fail", lang))
    else:
        st.warning(get_text("help_submit_warning", lang))

            if sent:
                st.success(get_text("help_submit_success", lang))
            else:
                st.error(get_text("help_submit_fail", lang))
        else:
            st.warning(get_text("help_submit_warning", lang))
