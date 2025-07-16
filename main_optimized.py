import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ==============================================================================
# ĐẢM BẢO FILE 'a04ecaf1_1dae_4c90_8081_086cd7c7b725.py' NẰM CÙNG THƯ MỤC
# HOẶC THAY THẾ TÊN FILE NẾU BẠN ĐÃ ĐỔI TÊN NÓ.
# ==============================================================================
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    apply_filters, export_report, export_pdf_report,
    apply_comparison_filters, export_comparison_report, export_comparison_pdf_report
)
# ==============================================================================

script_dir = os.path.dirname(__file__)
csv_file_path = os.path.join(script_dir, "invited_emails.csv")

# Gọi hàm setup_paths ngay từ đầu để path_dict có sẵn
path_dict = setup_paths()

# ---------------------------
# PHẦN XÁC THỰC TRUY CẬP
# ---------------------------

@st.cache_data
def load_invited_emails():
    try:
        df = pd.read_csv(csv_file_path, header=None, encoding='utf-8')
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

# Sử dụng session_state để lưu trữ lựa chọn ngôn ngữ
if 'lang' not in st.session_state:
    st.session_state.lang = 'vi' # Mặc định là tiếng Việt

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
        mode = st.selectbox(
            get_text('select_analysis_mode'),
            options=['year', 'month', 'week'],
            index=['year', 'month', 'week'].index(config_data['mode']) if config_data['mode'] in ['year', 'month', 'week'] else 0,
            key='standard_mode_tab'
        )
    with col2_std:
        selected_year = st.selectbox(
            get_text('select_year'),
            options=all_years,
            index=all_years.index(config_data['year']) if config_data['year'] in all_years else (0 if all_years else None),
            key='standard_year_tab'
        )
        if selected_year is None:
            st.warning(get_text('no_year_in_data'))

    with col3_std:
        default_months_standard = config_data['months'] if config_data['months'] else all_months
        selected_months = st.multiselect(
            get_text('select_months'),
            options=all_months,
            default=default_months_standard,
            key='standard_months_tab'
        )

    st.subheader(get_text('standard_project_selection_header'))

    initial_included_projects_config = config_data['project_filter_df'][
        config_data['project_filter_df']['Include'].astype(str).str.lower() == 'yes'
    ]['Project Name'].tolist()

    default_standard_projects = [p for p in initial_included_projects_config if p in all_projects]
    if not default_standard_projects and all_projects:
        default_standard_projects = all_projects

    standard_project_selection = st.multiselect(
        get_text('standard_project_selection_text'),
        options=all_projects,
        default=default_standard_projects,
        key='standard_project_selection_tab'
    )

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

            if df_filtered_standard.empty:
                st.warning(get_text('no_data_after_filter_standard'))
            else:
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
                    with st.spinner(get_text('generating_pdf_report')):
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
    if default_display_value not in display_options:
        default_display_value = display_options[0] # Fallback về mục đầu tiên
        # Cập nhật session_state để phản ánh fallback
        st.session_state.selected_comparison_mode_key = display_to_key_map[default_display_value]


    selected_comparison_display = st.selectbox(
        get_text('select_comparison_mode'),
        options=display_options,
        index=display_options.index(default_display_value), # Đặt index dựa trên giá trị mặc định
        key='comparison_mode_select_tab_main'
    )
    
    # Cập nhật key lựa chọn vào session_state khi người dùng thay đổi
    # Đảm bảo chỉ cập nhật nếu có sự thay đổi thực sự để tránh loop không cần thiết
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

    # Lựa chọn dự án chung cho tất cả các mode so sánh
    comp_projects = st.multiselect(
        get_text('select_projects_comp'),
        options=all_projects,
        default=[],
        key='comp_projects_select_tab_common'
    )

    # Note: Use the 'comparison_mode' variable derived from the dynamic map
    # for the conditional logic.
    if comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm" or comparison_mode == "Compare One Project Over Time (Months/Years)":
        if len(comp_projects) != 1:
            st.warning(get_text('select_single_project_warning'))
            validation_error = True

        # Lựa chọn năm cho mode "Qua Các Tháng/Năm"
        selected_years_over_time = st.multiselect(
            get_text('select_years_for_over_time_months'),
            options=all_years,
            default=[],
            key='comp_years_select_tab_over_time'
        )
        comp_years = selected_years_over_time # Gán vào comp_years cho config

        if len(selected_years_over_time) == 1:
            # So sánh qua các tháng trong một năm cụ thể
            st.info(get_text('comparison_over_months_note').format(selected_years_over_time[0]))
            comp_months = st.multiselect(
                get_text('select_months_for_single_year'),
                options=all_months,
                default=[],
                key='comp_months_select_tab_over_time'
            )
            # Nếu chọn một năm, thì tháng là bắt buộc
            if not comp_months:
                st.warning(get_text('no_month_selected_for_single_year'))
                validation_error = True

        elif len(selected_years_over_time) > 1:
            # So sánh qua nhiều năm, tháng không liên quan
            st.info(get_text('comparison_over_years_note'))
            # Đảm bảo months trống vì sẽ bị bỏ qua
            comp_months = []
        else: # Chưa chọn năm nào
            st.warning(get_text('no_comparison_criteria_selected'))
            validation_error = True
            comp_months = [] # Đảm bảo trống

    elif comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month", "So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        # Giao diện hiện có cho các mode so sánh khác
        col_comp1, col_comp2 = st.columns(2)
        with col_comp1:
            comp_years = st.multiselect(get_text('select_years'), options=all_years, default=[], key='comp_years_select_tab')
        with col_comp2:
            comp_months = st.multiselect(get_text('select_months_comp'), options=all_months, default=[], key='comp_months_select_tab')
        # comp_projects đã được chọn ở trên, không cần chọn lại

    st.markdown("---")
    st.subheader(get_text("export_options"))
    export_excel_comp = st.checkbox(get_text("export_excel_option"), value=True, key='export_excel_comp')
    export_pdf_comp = st.checkbox(get_text("export_pdf_option"), value=False, key='export_pdf_comp')

    if st.button(get_text('generate_comparison_report_btn'), key='generate_comparison_report_btn_tab'):
        if validation_error: # Kiểm tra nếu có lỗi xác thực từ UI
            st.error("Vui lòng sửa các lựa chọn được đánh dấu phía trên.")
        elif not export_excel_comp and not export_pdf_comp:
            st.warning(get_text("warning_select_export_format"))
        else:
            comparison_config = {
                'years': comp_years,
                'months': comp_months,
                'selected_projects': comp_projects,
            }

            # Hàm backend apply_comparison_filters và export_comparison_report
            # sẽ cần diễn giải đúng comparison_mode và comparison_config.
            df_comparison, message = apply_comparison_filters(df_raw, comparison_config, comparison_mode)

            if df_comparison.empty:
                st.warning(get_text('no_data_after_filter_comparison').format(message))
            else:
                st.success(get_text('data_filtered_success'))
                st.subheader(get_text('comparison_data_preview'))
                st.dataframe(df_comparison)

                report_generated_comp = False
                if export_excel_comp:
                    with st.spinner(get_text('generating_comparison_excel')):
                        excel_success_comp = export_comparison_report(df_comparison, comparison_config, path_dict['comparison_output_file'], comparison_mode)
                    if excel_success_comp:
                        st.success(get_text('comparison_excel_generated').format(os.path.basename(path_dict['comparison_output_file'])))
                        report_generated_comp = True
                    else:
                        st.error(get_text('failed_to_generate_comparison_excel'))

                if export_pdf_comp:
                    with st.spinner(get_text('generating_comparison_pdf')):
                        pdf_success_comp = export_comparison_pdf_report(df_comparison, comparison_config, path_dict['comparison_pdf_report'], comparison_mode, path_dict['logo_path'])
                    if pdf_success_comp:
                        st.success(get_text('comparison_pdf_generated').format(os.path.basename(path_dict['comparison_pdf_report'])))
                        report_generated_comp = True
                    else:
                        st.error(get_text('failed_to_generate_comparison_pdf'))

                if report_generated_comp:
                    if export_excel_comp and os.path.exists(path_dict['comparison_output_file']):
                        with open(path_dict['comparison_output_file'], "rb") as f:
                            st.download_button(get_text("download_comparison_excel"), data=f, file_name=os.path.basename(path_dict['comparison_output_file']), use_container_width=True, key='download_excel_comp_btn')
                    if export_pdf_comp and os.path.exists(path_dict['comparison_pdf_report']):
                        with open(path_dict['comparison_pdf_report'], "rb") as f:
                            st.download_button(get_text("download_comparison_pdf"), data=f, file_name=os.path.basename(path_dict['comparison_pdf_report']), use_container_width=True, key='download_pdf_comp_btn')
                else:
                    st.error(get_text('error_generating_report'))


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
