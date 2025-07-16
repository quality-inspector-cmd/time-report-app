import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ==============================================================================
# ĐẢM BẢO TÊN FILE LOGIC DƯỚI ĐÂY CHÍNH XÁC VỚI FILE BẠN ĐÃ LƯU
# VÀ NÓ NẰM CÙNG THƯ MỤC VỚI main_optimized.py
# ==============================================================================
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    apply_filters, export_report, export_pdf_report,
    apply_comparison_filters, export_comparison_report, export_comparison_pdf_report
)
# ==============================================================================

script_dir = os.path.dirname(__file__)
csv_file_path = os.path.join(script_dir, "invited_emails.csv")

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

INVITED_EMAILS = load_invited_emails()

def log_user_access(email):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = {"Time": timestamp, "Email": email}
    if "access_log" not in st.session_state:
        st.session_state.access_log = []
    st.session_state.access_log.append(log_entry)

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
    st.stop()

# ---------------------------
# PHẦN GIAO DIỆN CHÍNH CỦA ỨNG DỤNG
# ---------------------------

st.set_page_config(
    page_title="Triac Time Report",
    page_icon="⏰",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
        .report-title {font-size: 30px; color: #003366; font-weight: bold;}
        .report-subtitle {font-size: 14px; color: gray;}
        footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([0.12, 0.88])
with col1:
    logo_path = "triac_logo.png"
    if os.path.exists(logo_path):
        st.image(logo_path, width=110)
    else:
        st.warning(f"File logo '{logo_path}' không tìm thấy. Vui lòng kiểm tra đường dẫn.")
with col2:
    st.markdown("<div class='report-title'>Triac Time Report Generator</div>", unsafe_allow_html=True)
    st.markdown("<div class='report-subtitle'>Reporting tool for time tracking and analysis</div>", unsafe_allow_html=True)

translations = {
    "English": {
        "report_tab": "Standard Report",
        "compare_report_tab": "Comparison Report",
        "data_preview": "Data Preview",
        "user_guide": "User Guide",

        "mode": "Select mode",
        "year": "Select year(s)",
        "month": "Select month(s)",
        "project": "Select project(s)",
        "report_button": "Generate Report",
        "no_data": "No data after filtering",
        "report_done": "Report created successfully",
        "download_excel": "Download Excel",
        "download_pdf": "Download PDF",
        "export_options": "Export Options",
        "export_excel_option": "Export as Excel (.xlsx)",
        "export_pdf_option": "Export as PDF (.pdf)",
        "no_project_data_for_timeframe": "Some previously selected projects do not have data in the chosen time period and have been automatically deselected: {removed_projects}",


        # Comparison specific translations
        "comparison_mode": "Select Comparison Mode",
        "comp_proj_month": "Compare Projects in a Month",
        "comp_proj_year": "Compare Projects in a Year",
        "comp_one_proj_time": "Compare One Project Over Time (Months/Years)",
        "comp_years": "Select Year(s)",
        "comp_months": "Select Month(s)",
        "comp_projects": "Select Project(s)",
        "generate_comp_report": "Generate Comparison Report",
        "comp_data_preview": "Comparison Data Preview",
        "no_comp_data": "No data for selected comparison criteria or invalid selection.",
        "download_comp_excel": "Download Comparison Excel",
        "download_comp_pdf": "Download Comparison PDF",
        "comp_report_done": "Comparison Report created successfully",
        "select_criteria": "Please select enough criteria for comparison."
    },
    "Tiếng Việt": {
        "report_tab": "Báo Cáo Tiêu Chuẩn",
        "compare_report_tab": "Báo Cáo So Sánh",
        "data_preview": "Xem Dữ Liệu",
        "user_guide": "Hướng Dẫn Sử Dụng",

        "mode": "Chọn chế độ",
        "year": "Chọn năm",
        "month": "Chọn tháng",
        "project": "Chọn dự án",
        "report_button": "Tạo báo cáo",
        "no_data": "Không có dữ liệu sau khi lọc",
        "report_done": "Đã tạo báo cáo thành công",
        "download_excel": "Tải Excel",
        "download_pdf": "Tải PDF",
        "export_options": "Tùy chọn xuất báo cáo",
        "export_excel_option": "Xuất ra Excel (.xlsx)",
        "export_pdf_option": "Xuất ra PDF (.pdf)",
        "no_project_data_for_timeframe": "Một số dự án đã chọn trước đó không có dữ liệu trong khoảng thời gian đã chọn và đã được tự động bỏ chọn: {removed_projects}",

        # Comparison specific translations
        "comparison_mode": "Chọn Chế Độ So Sánh",
        "comp_proj_month": "So Sánh Dự Án Trong Một Tháng",
        "comp_proj_year": "So Sánh Dự Án Trong Một Năm",
        "comp_one_proj_time": "So Sánh Một Dự Án Qua Các Tháng/Năm",
        "comp_years": "Chọn Năm",
        "comp_months": "Chọn Tháng",
        "comp_projects": "Chọn Dự Án",
        "generate_comp_report": "Tạo Báo Cáo So Sánh",
        "comp_data_preview": "Xem Dữ Liệu So Sánh",
        "no_comp_data": "Không có dữ liệu phù hợp với tiêu chí so sánh hoặc lựa chọn không hợp lệ.",
        "download_comp_excel": "Tải Excel So Sánh",
        "download_comp_pdf": "Tải PDF So Sánh",
        "comp_report_done": "Đã tạo báo cáo so sánh thành công",
        "select_criteria": "Vui lòng chọn đủ tiêu chí để so sánh."
    }
}

lang = st.sidebar.selectbox("Language / Ngôn ngữ", ["English", "Tiếng Việt"])
T = translations[lang]

path_dict = setup_paths()

@st.cache_data(ttl=1800)
def cached_load_data():
    return load_raw_data(path_dict)

@st.cache_data(ttl=1800)
def cached_read_configs():
    return read_configs(path_dict)

with st.spinner("Đang tải dữ liệu..."):
    df_raw = cached_load_data()
    default_config_data = cached_read_configs()

all_years = sorted(df_raw['Year'].dropna().unique().astype(int).tolist(), reverse=True)
all_months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
all_projects = sorted(df_raw['Project name'].dropna().unique().tolist()) # Danh sách tất cả dự án gốc

tab1, tab_comparison, tab2, tab3 = st.tabs([
    T["report_tab"],
    T["compare_report_tab"],
    T["data_preview"],
    T["user_guide"]
])

with tab1: # Báo cáo tiêu chuẩn
    st.header(T["report_tab"])
    col1, col2, col3 = st.columns(3)
    with col1:
        default_years = [default_config_data['year']] if default_config_data['year'] in all_years else []
        selected_years_standard = st.multiselect(
            T["year"],
            options=all_years,
            default=default_years,
            key="standard_report_years" # Thêm key để Streamlit quản lý trạng thái tốt hơn
        )
    with col2:
        selected_months_standard = st.multiselect(
            T["month"],
            options=all_months,
            default=default_config_data['months'],
            key="standard_report_months" # Thêm key
        )
    
    # --- LOGIC MỚI: Lọc dự án khả dụng dựa trên năm và tháng đã chọn ---
    available_projects_for_timeframe = []
    if selected_years_standard:
        df_temp = df_raw[df_raw['Year'].isin(selected_years_standard)]
        if selected_months_standard:
            df_temp = df_temp[df_temp['MonthName'].isin(selected_months_standard)]
        
        available_projects_for_timeframe = sorted(df_temp['Project name'].dropna().unique().tolist())
    else:
        # Nếu chưa chọn năm nào, tất cả dự án đều có thể có
        available_projects_for_timeframe = all_projects

    # Lấy các dự án mặc định từ config (nếu có) và lọc chúng qua danh sách khả dụng
    default_included_projects_from_config = default_config_data['project_filter_df'][
        default_config_data['project_filter_df']['Include'].str.lower() == 'yes'
    ]['Project Name'].tolist()
    
    # Lấy lựa chọn dự án hiện tại của người dùng (nếu có trong session_state)
    # Đây là nơi chúng ta sẽ lưu trữ và lấy lại lựa chọn của người dùng
    if 'selected_projects_standard' not in st.session_state:
        st.session_state.selected_projects_standard = default_included_projects_from_config
    
    # Kiểm tra các dự án đã chọn trước đó (nếu có) có còn trong danh sách khả dụng không
    current_selected_projects = st.session_state.selected_projects_standard
    
    projects_to_keep = [proj for proj in current_selected_projects if proj in available_projects_for_timeframe]
    projects_removed = [proj for proj in current_selected_projects if proj not in available_projects_for_timeframe]

    if projects_removed:
        st.warning(T["no_project_data_for_timeframe"].format(removed_projects=", ".join(projects_removed)))
        st.session_state.selected_projects_standard = projects_to_keep # Cập nhật session_state

    with col3:
        # Sử dụng st.session_state.selected_projects_standard làm default
        selected_projects_standard = st.multiselect(
            T["project"],
            options=available_projects_for_timeframe, # CHỈ HIỂN THỊ CÁC DỰ ÁN CÓ DỮ LIỆU
            default=st.session_state.selected_projects_standard, # SỬ DỤNG TRẠNG THÁI HIỆN TẠI
            key="standard_report_projects" # Thêm key
        )
        # Cập nhật lại session state khi người dùng thay đổi lựa chọn
        st.session_state.selected_projects_standard = selected_projects_standard

    st.markdown("---") # Đường phân cách
    st.subheader(T["export_options"]) # Tiêu đề cho tùy chọn xuất
    export_excel_standard = st.checkbox(T["export_excel_option"], value=True, key="excel_standard_chk")
    export_pdf_standard = st.checkbox(T["export_pdf_option"], value=True, key="pdf_standard_chk")

    if st.button(T["report_button"], use_container_width=True, key="generate_standard_report_btn"):
        if not selected_years_standard:
            st.warning("Vui lòng chọn ít nhất một năm cho báo cáo tiêu chuẩn.")
        elif not selected_projects_standard:
            st.warning("Vui lòng chọn ít nhất một dự án cho báo cáo tiêu chuẩn.")
        elif not export_excel_standard and not export_pdf_standard:
            st.warning("Vui lòng chọn ít nhất một định dạng xuất báo cáo (Excel hoặc PDF).")
        else:
            with st.spinner("Đang tạo báo cáo..."):
                user_project_filter_df = pd.DataFrame({
                    'Project Name': selected_projects_standard,
                    'Include': ['yes'] * len(selected_projects_standard)
                })

                # Kiểm tra user_project_filter_df.empty LÀ KHÔNG CẦN THIẾT NỮA
                # vì đã kiểm tra selected_projects_standard ở trên
                
                config_standard = {
                    'mode': 'year' if not selected_months_standard else 'month',
                    'years': selected_years_standard,
                    'months': selected_months_standard,
                    'project_filter_df': user_project_filter_df
                }
                
                df_filtered_standard = apply_filters(df_raw, config_standard)

                if df_filtered_standard.empty:
                    st.warning(T["no_data"])
                else:
                    report_generated_standard = False
                    if export_excel_standard:
                        export_report(df_filtered_standard, config_standard, path_dict)
                        report_generated_standard = True
                    if export_pdf_standard:
                        export_pdf_report(df_filtered_standard, config_standard, path_dict)
                        report_generated_standard = True
                    
                    if report_generated_standard:
                        st.success(f"{T['report_done']}.")
                        if export_excel_standard:
                            with open(path_dict['output_file'], "rb") as f:
                                st.download_button(T["download_excel"], f, file_name=os.path.basename(path_dict['output_file']), use_container_width=True)
                        if export_pdf_standard:
                            with open(path_dict['pdf_report'], "rb") as f:
                                st.download_button(T["download_pdf"], f, file_name=os.path.basename(path_dict['pdf_report']), use_container_width=True)
                    else:
                        st.error("Có lỗi xảy ra khi tạo báo cáo. Vui lòng thử lại.")


with tab_comparison: # Báo cáo so sánh
    st.header(T["compare_report_tab"])

    comparison_mode = st.selectbox(
        T["comparison_mode"],
        [
            T["comp_proj_month"],
            T["comp_proj_year"],
            T["comp_one_proj_time"]
        ],
        key="comp_mode_select"
    )

    col1_comp, col2_comp, col3_comp = st.columns(3)
    with col1_comp:
        comp_years = st.multiselect(T["comp_years"], options=all_years, default=[datetime.now().year] if datetime.now().year in all_years else [], key="comp_years_select")
    with col2_comp:
        comp_months = st.multiselect(T["comp_months"], options=all_months, default=[], key="comp_months_select")
    
    # --- LOGIC TƯƠNG TỰ CHO TAB SO SÁNH (chỉ ví dụ, bạn có thể chỉnh sửa thêm) ---
    available_projects_for_comp_timeframe = []
    if comp_years:
        df_temp_comp = df_raw[df_raw['Year'].isin(comp_years)]
        if comp_months:
            df_temp_comp = df_temp_comp[df_temp_comp['MonthName'].isin(comp_months)]
        
        available_projects_for_comp_timeframe = sorted(df_temp_comp['Project name'].dropna().unique().tolist())
    else:
        available_projects_for_comp_timeframe = all_projects

    if 'selected_projects_comparison' not in st.session_state:
        st.session_state.selected_projects_comparison = [] # Khởi tạo rỗng
    
    current_selected_projects_comp = st.session_state.selected_projects_comparison
    projects_to_keep_comp = [proj for proj in current_selected_projects_comp if proj in available_projects_for_comp_timeframe]
    projects_removed_comp = [proj for proj in current_selected_projects_comp if proj not in available_projects_for_comp_timeframe]

    if projects_removed_comp:
        st.warning(T["no_project_data_for_timeframe"].format(removed_projects=", ".join(projects_removed_comp)))
        st.session_state.selected_projects_comparison = projects_to_keep_comp

    with col3_comp:
        comp_projects = st.multiselect(T["comp_projects"], options=available_projects_for_comp_timeframe, default=st.session_state.selected_projects_comparison, key="comp_projects_select")
        st.session_state.selected_projects_comparison = comp_projects # Cập nhật session state
    # --- HẾT LOGIC TƯƠNG TỰ CHO TAB SO SÁNH ---


    comparison_config = {
        'years': comp_years,
        'months': comp_months,
        'selected_projects': comp_projects
    }

    st.markdown("---")
    st.subheader(T["export_options"])
    export_excel_comp = st.checkbox(T["export_excel_option"], value=True, key="excel_comp_chk")
    export_pdf_comp = st.checkbox(T["export_pdf_option"], value=True, key="pdf_comp_chk")

    if st.button(T["generate_comp_report"], use_container_width=True, key="generate_comparison_report_btn"):
        if not comp_years or not comp_months or not comp_projects:
            st.warning(T["select_criteria"])
        elif not export_excel_comp and not export_pdf_comp:
            st.warning("Vui lòng chọn ít nhất một định dạng xuất báo cáo so sánh (Excel hoặc PDF).")
        else:
            with st.spinner("Đang tạo báo cáo so sánh..."):
                df_comparison, message = apply_comparison_filters(df_raw, comparison_config, comparison_mode)

                if df_comparison.empty:
                    st.warning(message)
                else:
                    st.subheader(T["comp_data_preview"])
                    st.dataframe(df_comparison)

                    report_generated_comp = False
                    if export_excel_comp:
                        export_comparison_report(df_comparison, comparison_config, path_dict, comparison_mode)
                        report_generated_comp = True
                    if export_pdf_comp:
                        export_comparison_pdf_report(df_comparison, comparison_config, path_dict, comparison_mode)
                        report_generated_comp = True

                    if report_generated_comp:
                        st.success(f"{T['comp_report_done']}.")
                        if export_excel_comp:
                            with open(path_dict['comparison_output_file'], "rb") as f:
                                st.download_button(T["download_comp_excel"], f, file_name=os.path.basename(path_dict['comparison_output_file']), use_container_width=True)
                        if export_pdf_comp:
                            with open(path_dict['comparison_pdf_report'], "rb") as f:
                                st.download_button(T["download_comp_pdf"], f, file_name=os.path.basename(path_dict['comparison_pdf_report']), use_container_width=True)
                    else:
                        st.error("Có lỗi xảy ra khi tạo báo cáo so sánh. Vui lòng thử lại.")


with tab2: # Xem dữ liệu
    st.subheader(T["data_preview"])
    st.dataframe(df_raw.head(100), use_container_width=True)

with tab3: # Hướng dẫn sử dụng
    st.markdown(f"### {T['user_guide']}")
    if lang == "English":
        st.markdown("""
        #### Standard Report:
        - Select desired "Mode" (Year, Month, or Week).
        - Choose specific "Year(s)", "Month(s)", and "Project(s)" to filter the data. Note: The list of available projects will update based on selected years and months.
        - Select desired export formats (Excel, PDF, or both).
        - Click "Generate Report" and then download the generated files.

        #### Comparison Report:
        - Select a "Comparison Mode":
            - **Compare Projects in a Month:** Requires one year, one month, and multiple projects.
            - **Compare Projects in a Year:** Requires one year and multiple projects.
            - **Compare One Project Over Time (Months/Years):** Requires one project and multiple months/years.
        - Select desired Years, Months, and Projects based on the chosen comparison mode. The list of available projects will update based on selected years and months.
        - Select desired export formats (Excel, PDF, or both).
        - Click "Generate Comparison Report" and then download the generated files.
        """)
    else: # Tiếng Việt
        st.markdown("""
        #### Báo Cáo Tiêu Chuẩn:
        - Chọn "Chế độ" (Năm, Tháng, hoặc Tuần) mong muốn.
        - Chọn "Năm", "Tháng", và "Dự án" cụ thể để lọc dữ liệu. Lưu ý: Danh sách dự án khả dụng sẽ được cập nhật dựa trên năm và tháng đã chọn.
        - Chọn định dạng xuất báo cáo (Excel, PDF hoặc cả hai).
        - Nhấp vào "Tạo báo cáo" và sau đó tải các file đã tạo.

        #### Báo Cáo So Sánh:
        - Chọn "Chế Độ So Sánh":
            - **So Sánh Dự Án Trong Một Tháng:** Yêu cầu một năm, một tháng và nhiều dự án.
            - **So Sánh Dự Án Trong Một Năm:** Yêu cầu một năm và nhiều dự án (tháng là tất cả các tháng được chọn).
            - **So Sánh Một Dự Án Qua Các Tháng/Năm:** Yêu cầu một dự án và nhiều tháng/năm.
        - Chọn Năm, Tháng và Dự án dựa trên chế độ so sánh đã chọn. Danh sách dự án khả dụng sẽ được cập nhật dựa trên năm và tháng đã chọn.
        - Chọn định dạng xuất báo cáo (Excel, PDF hoặc cả hai).
        - Nhấp vào "Tạo Báo Cáo So Sánh" và sau đó tải các file đã tạo.
        """)
