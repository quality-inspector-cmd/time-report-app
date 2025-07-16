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
    apply_comparison_filters, export_comparison_report, export_comparison_pdf_report
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
    st.session_state.comparison_selected_years = []

if 'comparison_selected_months' not in st.session_state:
    st.session_state.comparison_selected_months = []

if 'comparison_selected_projects' not in st.session_state:
    st.session_state.comparison_selected_projects = []

if 'standard_report_mode' not in st.session_state:
    st.session_state.standard_report_mode = "year" # Giá trị mặc định

if 'standard_report_year' not in st.session_state:
    st.session_state.standard_report_year = datetime.now().year # Giá trị mặc định

if 'standard_report_months' not in st.session_state:
    st.session_state.standard_report_months = [] # Giá trị mặc định

if 'standard_report_projects' not in st.session_state:
    st.session_state.standard_report_projects = [] # Giá trị mặc định

if 'raw_data_loaded' not in st.session_state:
    st.session_state.raw_data_loaded = False
    st.session_state.raw_data = pd.DataFrame()
    st.session_state.all_projects = [] # Danh sách tất cả các dự án

# =============================================================================
# CÁC HÀM HỖ TRỢ
# =============================================================================
def get_text(key):
    """
    Hàm trả về văn bản dựa trên ngôn ngữ đã chọn.
    Sử dụng từ điển đơn giản cho ví dụ này.
    Trong ứng dụng thực tế, bạn sẽ tải từ các file ngôn ngữ.
    """
    texts = {
        'en': {
            'app_title': "Triac Time Report Generator",
            'select_language': "Select Language:",
            'standard_report_header': "Standard Report Generation",
            'report_mode': "Select Report Mode:",
            'select_year': "Select Year:",
            'select_months': "Select Months (optional):",
            'all_months': "All Months",
            'select_projects': "Select Projects:",
            'export_format': "Export Format:",
            'create_report': "Create Standard Report",
            'download_excel': "Download Excel Report",
            'download_pdf': "Download PDF Report",
            'comparison_report_header': "Comparison Report Generation",
            'comparison_mode': "Select Comparison Mode:",
            'compare_projects_month': "Compare Projects in a Month",
            'compare_projects_year': "Compare Projects in a Year",
            'compare_one_project_over_time': "Compare One Project Over Time (Months/Years)",
            'select_years_comp': "Select Years for Comparison:",
            'select_months_comp': "Select Months for Comparison (optional):",
            'select_projects_comp': "Select Projects for Comparison:",
            'create_comparison_report': "Create Comparison Report",
            'download_comparison_excel': "Download Comparison Excel Report",
            'download_comparison_pdf': "Download Comparison PDF Report",
            'data_preview_header': "Raw Data Preview",
            'no_raw_data': "No raw data loaded. Please ensure 'Time_report.xlsm' is in the same directory.",
            'user_guide': "User Guide",
            'app_description': "This application allows you to generate time reports from your raw time data. You can create standard reports filtered by year, month, and project, or generate comparison reports to analyze trends across projects or over time.",
            'how_to_use_standard': "How to Use Standard Report:",
            'how_to_use_comparison': "How to Use Comparison Report:",
            'how_to_use_data_preview': "How to Use Data Preview:",
            'template_config_info': "Template File Configuration (External to App):",
            'common_errors': "Common Errors:",
            'template_not_found': "Template file not found: Ensure 'Time_report.xlsm' is in the same directory as this application.",
            'data_load_error': "Raw data could not be loaded: Check data format and column names in 'Raw Data' sheet.",
            'success_generating_report': "Report generation process completed. Check console for details and download links below.",
            'error_generating_report': "Error generating report. Please check the console for details.",
            'no_data_for_standard_report': "No data after filtering for standard report.",
            'no_data_for_comparison_report': "No data after filtering for comparison report.",
            'select_at_least_two_projects': "Please select at least TWO projects for this comparison mode.",
            'select_one_year_one_month': "Please select ONE year and ONE month for this comparison mode.",
            'select_only_one_project': "Please select ONLY ONE project for this comparison mode.",
            'select_years_or_months': "Please select at least one year (or multiple years) or one year and multiple months for comparison."

        },
        'vn': {
            'app_title': "Công cụ tạo báo cáo thời gian Triac",
            'select_language': "Chọn ngôn ngữ:",
            'standard_report_header': "Tạo báo cáo tiêu chuẩn",
            'report_mode': "Chọn chế độ báo cáo:",
            'select_year': "Chọn năm:",
            'select_months': "Chọn tháng (tùy chọn):",
            'all_months': "Tất cả các tháng",
            'select_projects': "Chọn dự án:",
            'export_format': "Định dạng xuất:",
            'create_report': "Tạo báo cáo tiêu chuẩn",
            'download_excel': "Tải báo cáo Excel",
            'download_pdf': "Tải báo cáo PDF",
            'comparison_report_header': "Tạo báo cáo so sánh",
            'comparison_mode': "Chọn chế độ so sánh:",
            'compare_projects_month': "So Sánh Dự Án Trong Một Tháng",
            'compare_projects_year': "So Sánh Dự Án Trong Một Năm",
            'compare_one_project_over_time': "So Sánh Một Dự Án Qua Các Tháng/Năm",
            'select_years_comp': "Chọn năm để so sánh:",
            'select_months_comp': "Chọn tháng để so sánh (tùy chọn):",
            'select_projects_comp': "Chọn dự án để so sánh:",
            'create_comparison_report': "Tạo báo cáo so sánh",
            'download_comparison_excel': "Tải báo cáo so sánh Excel",
            'download_comparison_pdf': "Tải báo cáo so sánh PDF",
            'data_preview_header': "Xem trước dữ liệu thô",
            'no_raw_data': "Chưa có dữ liệu thô nào được tải. Vui lòng đảm bảo 'Time_report.xlsm' nằm cùng thư mục.",
            'user_guide': "Hướng dẫn sử dụng",
            'app_description': "Ứng dụng này cho phép bạn tạo báo cáo thời gian từ dữ liệu thời gian thô của mình. Bạn có thể tạo các báo cáo tiêu chuẩn được lọc theo năm, tháng và dự án, hoặc tạo báo cáo so sánh để phân tích xu hướng giữa các dự án hoặc theo thời gian.",
            'how_to_use_standard': "Cách sử dụng Báo cáo tiêu chuẩn:",
            'how_to_use_comparison': "Cách sử dụng Báo cáo so sánh:",
            'how_to_use_data_preview': "Cách sử dụng Xem trước dữ liệu:",
            'template_config_info': "Cấu hình file Template (Bên ngoài ứng dụng):",
            'common_errors': "Các lỗi thường gặp:",
            'template_not_found': "Không tìm thấy file template: Đảm bảo 'Time_report.xlsm' nằm cùng thư mục với ứng dụng này.",
            'data_load_error': "Không thể tải dữ liệu thô: Kiểm tra định dạng dữ liệu và tên cột trong sheet 'Raw Data'.",
            'success_generating_report': "Quá trình tạo báo cáo đã hoàn tất. Kiểm tra console để biết chi tiết và liên kết tải xuống bên dưới.",
            'error_generating_report': "Lỗi khi tạo báo cáo. Vui lòng kiểm tra console để biết chi tiết.",
            'no_data_for_standard_report': "Không có dữ liệu sau khi lọc cho báo cáo tiêu chuẩn.",
            'no_data_for_comparison_report': "Không có dữ liệu sau khi lọc cho báo cáo so sánh.",
            'select_at_least_two_projects': "Vui lòng chọn ít nhất HAI dự án cho chế độ so sánh này.",
            'select_one_year_one_month': "Vui lòng chọn MỘT năm và MỘT tháng cho chế độ so sánh này.",
            'select_only_one_project': "Vui lòng chọn CHỈ MỘT dự án cho chế độ so sánh này.",
            'select_years_or_months': "Vui lòng chọn ít nhất một năm (hoặc nhiều năm) hoặc một năm và nhiều tháng để so sánh."
        }
    }
    return texts[st.session_state.language].get(key, key)

# =============================================================================
# KHU VỰC CHÍNH CỦA ỨNG DỤNG STREAMLIT
# =============================================================================

st.set_page_config(layout="wide", page_title=get_text('app_title'))

# Bộ chọn ngôn ngữ ở thanh bên
st.sidebar.selectbox(
    get_text('select_language'),
    ['English', 'Tiếng Việt'],
    key='language_selector',
    on_change=lambda: st.session_state.__setitem__('language', 'en' if st.session_state.language_selector == 'English' else 'vn')
)
if 'language' not in st.session_state:
    st.session_state.language = 'en' # Mặc định tiếng Anh

st.title(get_text('app_title'))

# Load initial data and configurations if not already loaded
if not st.session_state.raw_data_loaded:
    try:
        template_file = path_dict['template_file']
        st.session_state.raw_data = load_raw_data(template_file)
        if not st.session_state.raw_data.empty:
            st.session_state.all_projects = sorted(st.session_state.raw_data['Project name'].unique().tolist())
        st.session_state.raw_data_loaded = True
    except Exception as e:
        st.error(f"{get_text('data_load_error')} {e}")
        st.session_state.raw_data_loaded = False # Đảm bảo trạng thái là False nếu có lỗi

df_raw = st.session_state.raw_data
all_projects = st.session_state.all_projects

# Tạo các tab
tab_standard_report_main, tab_comparison_report_main, tab_data_preview_main, tab_user_guide_main = st.tabs([
    get_text('standard_report_header'),
    get_text('comparison_report_header'),
    get_text('data_preview_header'),
    get_text('user_guide')
])

# =========================================================================
# STANDARD REPORT TAB
# =========================================================================
with tab_standard_report_main:
    st.header(get_text('standard_report_header'))

    col1, col2 = st.columns(2)
    with col1:
        st.session_state.standard_report_mode = st.radio(
            get_text('report_mode'),
            ['year', 'month', 'week'],
            index=['year', 'month', 'week'].index(st.session_state.standard_report_mode),
            key='standard_report_mode_radio'
        )
    
    with col2:
        available_years = sorted(df_raw['Year'].unique().tolist()) if not df_raw.empty else [datetime.now().year]
        if st.session_state.standard_report_year not in available_years and available_years:
            st.session_state.standard_report_year = available_years[0] # Đặt lại nếu năm mặc định không có
        
        st.session_state.standard_report_year = st.selectbox(
            get_text('select_year'),
            available_years,
            index=available_years.index(st.session_state.standard_report_year) if st.session_state.standard_report_year in available_years else 0,
            key='standard_report_year_select'
        )
        
        all_months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        st.session_state.standard_report_months = st.multiselect(
            get_text('select_months'),
            all_months,
            default=st.session_state.standard_report_months,
            placeholder=get_text('all_months'),
            key='standard_report_months_multiselect'
        )
        
        # Lọc danh sách dự án có sẵn cho năm và tháng đã chọn
        df_for_project_selection = df_raw[df_raw['Year'] == st.session_state.standard_report_year]
        if st.session_state.standard_report_months:
            df_for_project_selection = df_for_project_selection[df_for_project_selection['MonthName'].isin(st.session_state.standard_report_months)]
        
        available_projects_for_standard = sorted(df_for_project_selection['Project name'].unique().tolist()) if not df_for_project_selection.empty else []

        if not st.session_state.standard_report_projects or not all(p in available_projects_for_standard for p in st.session_state.standard_report_projects):
            # Nếu các dự án đã chọn không còn khả dụng hoặc không có, đặt lại mặc định
            st.session_state.standard_report_projects = []
        
        st.session_state.standard_report_projects = st.multiselect(
            get_text('select_projects'),
            available_projects_for_standard,
            default=st.session_state.standard_report_projects,
            key='standard_report_projects_multiselect'
        )

    st.subheader(get_text('export_format'))
    col_export_std1, col_export_std2 = st.columns(2)
    with col_export_std1:
        export_excel_standard = st.checkbox("Excel", value=True, key='export_excel_standard_checkbox')
    with col_export_std2:
        export_pdf_standard = st.checkbox("PDF", value=True, key='export_pdf_standard_checkbox')

    if st.button(get_text('create_report'), key='create_standard_report_button'):
        if not st.session_state.standard_report_projects:
            st.warning(get_text('no_data_for_standard_report'))
        else:
            with st.spinner('Generating standard report...'):
                report_status = generate_reports_on_demand(
                    selected_mode=st.session_state.standard_report_mode,
                    selected_year=st.session_state.standard_report_year,
                    selected_months=st.session_state.standard_report_months,
                    selected_project_names_standard=st.session_state.standard_report_projects,
                    comparison_config_years=[], # Không liên quan đến báo cáo tiêu chuẩn
                    comparison_config_months=[], # Không liên quan đến báo cáo tiêu chuẩn
                    comparison_config_projects=[], # Không liên quan đến báo cáo tiêu chuẩn
                    comparison_report_mode=None, # Không liên quan đến báo cáo tiêu chuẩn
                    export_excel_standard=export_excel_standard,
                    export_pdf_standard=export_pdf_standard,
                    export_excel_comparison=False,
                    export_pdf_comparison=False
                )
            if report_status and report_status.get('status') == "success":
                st.success(get_text('success_generating_report'))
                # Hiển thị nút tải xuống
                if export_excel_standard and os.path.exists(path_dict['output_file']):
                    with open(path_dict['output_file'], "rb") as f:
                        st.download_button(get_text('download_excel'), data=f, file_name=os.path.basename(path_dict['output_file']), use_container_width=True, key='download_excel_std_btn')
                if export_pdf_standard and os.path.exists(path_dict['pdf_report']):
                    with open(path_dict['pdf_report'], "rb") as f:
                        st.download_button(get_text('download_pdf'), data=f, file_name=os.path.basename(path_dict['pdf_report']), use_container_width=True, key='download_pdf_std_btn')
            else:
                st.error(get_text('error_generating_report'))


# =========================================================================
# COMPARISON REPORT TAB
# =========================================================================
with tab_comparison_report_main:
    st.header(get_text('comparison_report_header'))

    st.session_state.comparison_mode = st.radio(
        get_text('comparison_mode'),
        [get_text('compare_projects_month'), get_text('compare_projects_year'), get_text('compare_one_project_over_time')],
        index=[get_text('compare_projects_month'), get_text('compare_projects_year'), get_text('compare_one_project_over_time')].index(st.session_state.comparison_mode),
        key='comparison_mode_radio'
    )
    
    available_years_comp = sorted(df_raw['Year'].unique().tolist()) if not df_raw.empty else [datetime.now().year]
    if not st.session_state.comparison_selected_years and available_years_comp:
        st.session_state.comparison_selected_years = [available_years_comp[0]] # Mặc định chọn năm đầu tiên

    st.session_state.comparison_selected_years = st.multiselect(
        get_text('select_years_comp'),
        available_years_comp,
        default=st.session_state.comparison_selected_years,
        key='comparison_years_multiselect'
    )

    all_months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    st.session_state.comparison_selected_months = st.multiselect(
        get_text('select_months_comp'),
        all_months,
        default=st.session_state.comparison_selected_months,
        placeholder=get_text('all_months'),
        key='comparison_months_multiselect'
    )

    # Lọc danh sách dự án có sẵn cho các năm và tháng đã chọn trong phần so sánh
    df_for_comp_project_selection = df_raw.copy()
    if st.session_state.comparison_selected_years:
        df_for_comp_project_selection = df_for_comp_project_selection[df_for_comp_project_selection['Year'].isin(st.session_state.comparison_selected_years)]
    if st.session_state.comparison_selected_months:
        df_for_comp_project_selection = df_for_comp_project_selection[df_for_comp_project_selection['MonthName'].isin(st.session_state.comparison_selected_months)]
    
    available_projects_for_comparison = sorted(df_for_comp_project_selection['Project name'].unique().tolist()) if not df_for_comp_project_selection.empty else []

    if not st.session_state.comparison_selected_projects or not all(p in available_projects_for_comparison for p in st.session_state.comparison_selected_projects):
        st.session_state.comparison_selected_projects = []

    st.session_state.comparison_selected_projects = st.multiselect(
        get_text('select_projects_comp'),
        available_projects_for_comparison,
        default=st.session_state.comparison_selected_projects,
        key='comparison_projects_multiselect'
    )

    st.subheader(get_text('export_format'))
    col_export_comp1, col_export_comp2 = st.columns(2)
    with col_export_comp1:
        export_excel_comp = st.checkbox("Excel", value=True, key='export_excel_comparison_checkbox')
    with col_export_comp2:
        export_pdf_comp = st.checkbox("PDF", value=True, key='export_pdf_comparison_checkbox')

    if st.button(get_text('create_comparison_report'), key='create_comparison_report_button'):
        validation_message = ""
        if st.session_state.comparison_mode == get_text('compare_projects_month'):
            if len(st.session_state.comparison_selected_years) != 1 or len(st.session_state.comparison_selected_months) != 1 or len(st.session_state.comparison_selected_projects) < 2:
                validation_message = get_text('select_one_year_one_month') + " " + get_text('select_at_least_two_projects')
        elif st.session_state.comparison_mode == get_text('compare_projects_year'):
            if len(st.session_state.comparison_selected_years) != 1 or len(st.session_state.comparison_selected_projects) < 2:
                validation_message = get_text('select_at_least_two_projects') + " " + get_text('select_one_year_one_month') # Re-using part of message for clarity
        elif st.session_state.comparison_mode == get_text('compare_one_project_over_time'):
            if len(st.session_state.comparison_selected_projects) != 1:
                validation_message = get_text('select_only_one_project')
            elif not (len(st.session_state.comparison_selected_years) > 1 or (len(st.session_state.comparison_selected_years) == 1 and len(st.session_state.comparison_selected_months) > 0)):
                validation_message = get_text('select_years_or_months')
        
        if validation_message:
            st.warning(validation_message)
        elif not st.session_state.comparison_selected_projects:
             st.warning(get_text('no_data_for_comparison_report'))
        else:
            with st.spinner('Generating comparison report...'):
                report_status = generate_reports_on_demand(
                    selected_mode=None, # Không liên quan đến báo cáo so sánh
                    selected_year=None, # Không liên quan đến báo cáo so sánh
                    selected_months=[], # Không liên quan đến báo cáo so sánh
                    selected_project_names_standard=[], # Không liên quan đến báo cáo so sánh
                    comparison_config_years=st.session_state.comparison_selected_years,
                    comparison_config_months=st.session_state.comparison_selected_months,
                    comparison_config_projects=st.session_state.comparison_selected_projects,
                    comparison_report_mode=st.session_state.comparison_mode,
                    export_excel_standard=False,
                    export_pdf_standard=False,
                    export_excel_comparison=export_excel_comp,
                    export_pdf_comparison=export_pdf_comp
                )
            if report_status and report_status.get('status') == "success":
                st.success(get_text('success_generating_report'))
                # Hiển thị nút tải xuống
                if export_excel_comp and os.path.exists(path_dict['comparison_output_file']):
                    with open(path_dict['comparison_output_file'], "rb") as f:
                        st.download_button(get_text('download_comparison_excel'), data=f, file_name=os.path.basename(path_dict['comparison_output_file']), use_container_width=True, key='download_excel_comp_btn')
                if export_pdf_comp and os.path.exists(path_dict['comparison_pdf_report']):
                    with open(path_dict['comparison_pdf_report'], "rb") as f:
                        st.download_button(get_text('download_comparison_pdf'), data=f, file_name=os.path.basename(path_dict['comparison_pdf_report']), use_container_width=True, key='download_pdf_comp_btn')
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
    st.markdown(f"{get_text('app_description')}")

    st.markdown(f"#### {get_text('how_to_use_standard')}")
    st.markdown("""
    * Chọn chế độ báo cáo (Năm, Tháng, hoặc Tuần).
    * Chọn năm và (tùy chọn) các tháng bạn muốn đưa vào báo cáo.
    * Chọn các dự án bạn muốn đưa vào.
    * Chọn định dạng xuất (Excel, PDF hoặc cả hai).
    * Nhấn nút 'Tạo báo cáo tiêu chuẩn'.
    * Sau khi báo cáo được tạo, bạn có thể tải xuống.
    """)

    st.markdown(f"#### {get_text('how_to_use_comparison')}")
    st.markdown("""
    * **Chọn chế độ so sánh:**
        * **So Sánh Dự Án Trong Một Tháng:** So sánh tổng số giờ của các dự án khác nhau trong MỘT năm và MỘT tháng cụ thể. Bạn phải chọn một năm, một tháng và ít nhất hai dự án.
        * **So Sánh Dự Án Trong Một Năm:** So sánh tổng số giờ của các dự án khác nhau trong MỘT năm cụ thể. Bạn phải chọn một năm và ít nhất hai dự án. Lựa chọn tháng sẽ bị bỏ qua.
        * **So Sánh Một Dự Án Qua Các Tháng/Năm:** So sánh hiệu suất của MỘT dự án duy nhất qua nhiều tháng trong cùng một năm, HOẶC so sánh qua nhiều năm.
            * Nếu bạn chọn **một năm và nhiều tháng**: Báo cáo sẽ so sánh dự án đó qua các tháng đã chọn trong năm đó.
            * Nếu bạn chọn **nhiều năm**: Báo cáo sẽ so sánh dự án đó qua các năm đã chọn. Lựa chọn tháng sẽ bị bỏ qua.
    * **Tạo báo cáo:** Nhấn nút 'Tạo báo cáo so sánh' để tạo file Excel và/hoặc PDF.

    ### 3. Xem trước dữ liệu
    Tab này cho phép bạn xem 100 hàng đầu tiên của dữ liệu thô đã tải, giúp bạn kiểm tra định dạng và nội dung dữ liệu.

    ### 4. Cấu hình file template (Bên ngoài ứng dụng)
    Công cụ đọc dữ liệu và cấu hình từ một file Excel template (thường là `Time_report.xlsm`). Đảm bảo rằng:
    * Sheet 'Raw Data' chứa dữ liệu thời gian thô của bạn với các cột cần thiết như 'Year', 'MonthName', 'Project name', v.v.
    * Sheet 'Config_Year_Mode' và 'Config_Project_Filter' có thể được sử dụng để đặt cấu hình mặc định, nhưng các lựa chọn trên giao diện sẽ ghi đè lên chúng.

    ### Lỗi thường gặp:
    * **File template không tìm thấy:** Đảm bảo `Time_report.xlsm` nằm cùng thư mục với ứng dụng này.
    * **Không tải được dữ liệu thô:** Kiểm tra định dạng dữ liệu và tên cột trong sheet 'Raw Data'.
    """)
