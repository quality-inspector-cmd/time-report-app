import streamlit as st
import pandas as pd
import os
from datetime import datetime
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import setup_paths, load_raw_data, read_configs, apply_filters, export_report, apply_comparison_filters, export_comparison_report, export_comparison_pdf_report

# Đặt cấu hình trang
st.set_page_config(page_title="Time Report Generator", layout="wide")

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
        'system_explanation_title': "Giới thiệu về hệ thống báo cáo thời gian",
        'system_explanation_text': """
        <span style='color:blue;'>
        Đây là một ứng dụng Python Streamlit được thiết kế để phân tích và tạo báo cáo về dữ liệu thời gian làm việc.
        Nó giúp người dùng dễ dàng tạo các báo cáo chi tiết theo năm, tháng, tuần và so sánh hiệu suất giữa các dự án hoặc qua các khoảng thời gian khác nhau.
        Ứng dụng này đọc dữ liệu từ file Excel template, áp dụng các bộ lọc cấu hình và xuất ra các báo cáo dạng Excel và PDF.
        </span>
        """,
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
        'select_years': "Chọn năm(các năm):",
        'select_months_comp': "Chọn tháng(các tháng):",
        'select_projects_comp': "Chọn dự án(các dự án):",
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
    },
    'en': {
        'app_title': "📊 Time Report Generator",
        'lang_select': "Select language:",
        'language_vi': "Tiếng Việt",
        'language_en': "English",
        'system_explanation_title': "About the Time Reporting System",
        'system_explanation_text': """
        <span style='color:blue;'>
        This is a Streamlit Python application designed to analyze and generate reports on work time data.
        It helps users easily create detailed reports by year, month, week, and compare performance between projects or over different time periods.
        The application reads data from an Excel template file, applies configured filters, and exports reports in both Excel and PDF formats.
        </span>
        """,
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
        'comparison_pdf_generated': "✅ Comparison PDF Report generated: {}",
        'download_comparison_pdf': "📥 Download Comparison PDF",
        'failed_to_generate_comparison_excel': "❌ Failed to generate Comparison Excel report.",
        'failed_to_generate_comparison_pdf': "❌ Failed to generate Comparison PDF report.",
        'raw_data_preview_header': "Raw Input Data (First 100 rows)",
        'no_raw_data': "No raw data loaded.",
        'no_year_in_data': "No years in data to select.",
    }
}

# Lấy từ điển văn bản dựa trên lựa chọn ngôn ngữ hiện tại
def get_text(key):
    return TEXTS[st.session_state.lang].get(key, f"Missing text for {key}")

# Header của ứng dụng
st.title(get_text('app_title'))

# Logo và lựa chọn ngôn ngữ trên cùng
col_logo, col_lang = st.columns([0.8, 0.2])
with col_logo:
    # Hiển thị logo Triac nếu tồn tại
    logo_path = path_dict['logo_path'] # Lấy từ setup_paths
    if os.path.exists(logo_path):
        st.image(logo_path, width=150)
    else:
        st.warning(f"Không tìm thấy logo tại: {logo_path}. Vui lòng đảm bảo file logo ở đúng vị trí.")

with col_lang:
    st.session_state.lang = st.radio(
        get_text('lang_select'),
        options=['vi', 'en'],
        format_func=lambda x: get_text('language_' + x),
        key='language_selector'
    )

# Diễn giải hệ thống
st.subheader(get_text('system_explanation_title'))
st.markdown(get_text('system_explanation_text'), unsafe_allow_html=True)


# Setup paths (e.g., template, output files)
path_dict = setup_paths()

# Check if template file exists
if not os.path.exists(path_dict['template_file']):
    st.error(get_text('template_not_found').format(path_dict['template_file']))
    st.stop()

# Load raw data and configurations once
@st.cache_data
def load_data_and_configs():
    df_raw = load_raw_data(path_dict['template_file'])
    config_data = read_configs(path_dict['template_file'])
    return df_raw, config_data

with st.spinner(get_text('loading_data')):
    df_raw, config_data = load_data_and_configs()

if df_raw.empty:
    st.error(get_text('failed_to_load_raw_data'))
    st.stop()

# Get unique years, months, and projects from raw data for selectbox options
all_years = sorted(df_raw['Year'].dropna().unique().astype(int).tolist())
month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
all_months = [m for m in month_order if m in df_raw['MonthName'].dropna().unique()]
all_projects = sorted(df_raw['Project name'].dropna().unique().tolist())

# Main interface tabs
tab_standard, tab_comparison, tab_data_preview = st.tabs([
    get_text('tab_standard_report'), 
    get_text('tab_comparison_report'), 
    get_text('tab_data_preview')
])

# =========================================================================
# STANDARD REPORT TAB
# =========================================================================
with tab_standard:
    st.header(get_text('standard_report_header'))

    col1, col2, col3 = st.columns(3)
    with col1:
        mode = st.selectbox(
            get_text('select_analysis_mode'), 
            options=['year', 'month', 'week'], 
            index=['year', 'month', 'week'].index(config_data['mode']) if config_data['mode'] in ['year', 'month', 'week'] else 0,
            key='standard_mode'
        )
    with col2:
        selected_year = st.selectbox(
            get_text('select_year'), 
            options=all_years, 
            index=all_years.index(config_data['year']) if config_data['year'] in all_years else (0 if all_years else None),
            key='standard_year'
        )
        if selected_year is None:
            st.warning(get_text('no_year_in_data'))
            
    with col3:
        default_months_standard = config_data['months'] if config_data['months'] else all_months
        selected_months = st.multiselect(
            get_text('select_months'), 
            options=all_months, 
            default=default_months_standard,
            key='standard_months'
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
        key='standard_project_selection'
    )

    if st.button(get_text('generate_standard_report_btn'), key='generate_standard_report_btn_main'):
        if selected_year is None:
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
                with st.spinner(get_text('generating_excel_report')):
                    excel_success = export_report(df_filtered_standard, standard_report_config, path_dict['output_file'])
                
                if excel_success:
                    st.success(get_text('excel_report_generated').format(os.path.basename(path_dict['output_file'])))
                    with open(path_dict['output_file'], "rb") as f:
                        st.download_button(get_text('download_excel_report'), data=f, file_name=os.path.basename(path_dict['output_file']), key='download_excel_standard')
                    
                    with st.spinner(get_text('generating_pdf_report')):
                        # Truyền logo_path vào hàm export_pdf_report
                        pdf_success = export_pdf_report(df_filtered_standard, standard_report_config, path_dict['pdf_report'], path_dict['logo_path'])
                    
                    if pdf_success:
                        st.success(get_text('pdf_report_generated').format(os.path.basename(path_dict['pdf_report'])))
                        with open(path_dict['pdf_report'], "rb") as f:
                            st.download_button(get_text('download_pdf_report'), data=f, file_name=os.path.basename(path_dict['pdf_report']), key='download_pdf_standard')
                    else:
                        st.error(get_text('failed_to_generate_pdf'))
                else:
                    st.error(get_text('failed_to_generate_excel'))


# =========================================================================
# COMPARISON REPORT TAB
# =========================================================================
with tab_comparison:
    st.header(get_text('comparison_report_header'))

    comparison_mode_options = {
        get_text('compare_projects_month'): "So Sánh Dự Án Trong Một Tháng",
        get_text('compare_projects_year'): "So Sánh Dự Án Trong Một Năm",
        get_text('compare_one_project_over_time'): "So Sánh Một Dự Án Qua Các Tháng/Năm"
    }
    
    # Tạo mapping ngược để lấy key tiếng Việt/Anh từ value
    reverse_comparison_mode_options = {v: k for k, v in comparison_mode_options.items()}

    selected_comparison_display = st.selectbox(
        get_text('select_comparison_mode'),
        options=list(comparison_mode_options.keys()),
        key='comparison_mode_select_tab'
    )
    # Lấy giá trị thực của chế độ so sánh (tiếng Việt) để truyền vào hàm backend
    comparison_mode = comparison_mode_options[selected_comparison_display]

    st.subheader(get_text('filter_data_for_comparison'))
    
    col_comp1, col_comp2 = st.columns(2)
    with col_comp1:
        comp_years = st.multiselect(get_text('select_years'), options=all_years, default=[], key='comp_years_select')
    with col_comp2:
        comp_months = st.multiselect(get_text('select_months_comp'), options=all_months, default=[], key='comp_months_select')
    
    comp_projects = st.multiselect(get_text('select_projects_comp'), options=all_projects, default=[], key='comp_projects_select')

    if st.button(get_text('generate_comparison_report_btn'), key='generate_comparison_report_btn_tab'):
        comparison_config = {
            'years': comp_years,
            'months': comp_months,
            'selected_projects': comp_projects,
        }

        df_comparison, message = apply_comparison_filters(df_raw, comparison_config, comparison_mode)

        if df_comparison.empty:
            st.warning(get_text('no_data_after_filter_comparison').format(message))
        else:
            st.success(get_text('data_filtered_success'))
            st.subheader(get_text('comparison_data_preview'))
            st.dataframe(df_comparison)

            with st.spinner(get_text('generating_comparison_excel')):
                excel_success_comp = export_comparison_report(df_comparison, comparison_config, path_dict['comparison_output_file'], comparison_mode)
            
            if excel_success_comp:
                st.success(get_text('comparison_excel_generated').format(os.path.basename(path_dict['comparison_output_file'])))
                with open(path_dict['comparison_output_file'], "rb") as f:
                    st.download_button(get_text('download_comparison_excel'), data=f, file_name=os.path.basename(path_dict['comparison_output_file']), key='download_excel_comparison')
                
                with st.spinner(get_text('generating_comparison_pdf')):
                    # Truyền logo_path vào hàm export_comparison_pdf_report
                    pdf_success_comp = export_comparison_pdf_report(df_comparison, comparison_config, path_dict['comparison_pdf_report'], comparison_mode, path_dict['logo_path'])
                
                if pdf_success_comp:
                    st.success(get_text('comparison_pdf_generated').format(os.path.basename(path_dict['comparison_pdf_report'])))
                    with open(path_dict['comparison_pdf_report'], "rb") as f:
                        st.download_button(get_text('download_comparison_pdf'), data=f, file_name=os.path.basename(path_dict['comparison_pdf_report']), key='download_pdf_comparison')
                else:
                    st.error(get_text('failed_to_generate_comparison_pdf'))
            else:
                st.error(get_text('failed_to_generate_comparison_excel'))

# =========================================================================
# DATA PREVIEW TAB
# =========================================================================
with tab_data_preview:
    st.subheader(get_text('raw_data_preview_header'))
    if not df_raw.empty:
        st.dataframe(df_raw.head(100))
    else:
        st.info(get_text('no_raw_data'))
