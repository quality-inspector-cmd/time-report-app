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
# TIÊU ĐỀ ỨNG DỤNG
# ==============================================================================\
st.set_page_config(layout="wide")
st.title("Công Cụ Tự Động Hóa Báo Cáo Giờ Làm Việc")
st.write("Chào mừng bạn đến với Công Cụ Tự Động Hóa Báo Cáo Giờ Làm Việc! Vui lòng điều hướng qua các tab để tạo báo cáo của bạn.")

# ==============================================================================\
# TẢI DỮ LIỆU ĐẦU VÀO
# ==============================================================================\
@st.cache_data(ttl=3600) # Cache dữ liệu trong 1 giờ
def cached_load():
    template_file = path_dict['template_file']
    if not os.path.exists(template_file):
        st.error("Lỗi: Không tìm thấy file template.")
        return None, None
    
    df_raw = load_raw_data(template_file)
    config_data = read_configs(template_file)
    return df_raw, config_data

df_raw, config_data = cached_load()

if df_raw is None or config_data is None:
    st.warning("Lỗi khi tải dữ liệu hoặc cấu hình.")
else:
    st.sidebar.success("Dữ liệu và cấu hình đã tải thành công!")

# ==============================================================================\
# CÁC TAB CHÍNH CỦA ỨNG DỤNG
# ==============================================================================\
tab_standard_report_main, tab_comparison_report_main, tab_data_preview, tab_instructions = st.tabs(["Báo cáo Tiêu chuẩn", "Báo cáo So sánh", "Xem trước Dữ liệu", "Hướng dẫn"])

# ==============================================================================\
# TAB 1: BÁO CÁO TIÊU CHUẨN
# ==============================================================================\
with tab_standard_report_main:
    st.header("Báo cáo Tiêu chuẩn")

    if df_raw is not None and config_data is not None:
        all_years = sorted(df_raw['Year'].unique().tolist(), reverse=True)
        all_months = ['January', 'February', 'March', 'April', 'May', 'June', 
                      'July', 'August', 'September', 'October', 'November', 'December']
        all_projects = sorted(df_raw['Project name'].unique().tolist())

        analysis_mode = st.radio("Chọn Chế độ phân tích:", ["Năm", "Tháng", "Tuần"])

        # Lựa chọn Năm
        selected_year_option = st.selectbox("Chọn Năm:", ["Tất cả các Năm"] + all_years, index=0)
        selected_year = selected_year_option if selected_year_option != "Tất cả các Năm" else None

        # Lựa chọn Tháng (chỉ hiển thị nếu chế độ là "Tháng")
        selected_months_standard = []
        if analysis_mode == "Tháng":
            selected_months_options = st.multiselect("Chọn Tháng:", all_months, default=all_months)
            selected_months_standard = selected_months_options if selected_months_options else []
        elif analysis_mode == "Năm":
            selected_months_standard = all_months # Mặc định chọn tất cả các tháng cho chế độ "Năm"
        # Chế độ "Tuần" sẽ không lọc theo tháng cụ thể, sẽ xử lý logic tuần bên trong generate_reports_on_demand

        # Lựa chọn Dự án
        selected_projects_standard = st.multiselect("Chọn Dự án:", all_projects, default=all_projects)
        if not selected_projects_standard:
            st.warning("Vui lòng chọn ít nhất một dự án cho báo cáo tiêu chuẩn.")


        if st.button("Tạo Báo cáo Tiêu chuẩn", key='generate_std_report_btn'):
            if not selected_year and selected_year_option != "Tất cả các Năm":
                st.warning("Vui lòng chọn một năm.")
            elif not selected_projects_standard:
                st.warning("Vui lòng chọn ít nhất một dự án cho báo cáo tiêu chuẩn.")
            else:
                export_excel, export_pdf = False, False
                with st.expander("Tùy chọn Xuất"):
                    export_excel = st.checkbox("Xuất ra Excel (.xlsx)", value=True, key='export_excel_std')
                    export_pdf = st.checkbox("Xuất ra PDF (.pdf)", value=True, key='export_pdf_std')

                if not export_excel and not export_pdf:
                    st.warning("Vui lòng chọn ít nhất một định dạng xuất (Excel hoặc PDF).")
                else:
                    with st.spinner("Đang tạo báo cáo Excel..." if export_excel else "Đang tạo báo cáo PDF..."):
                        try:
                            # Gọi hàm generate_reports_on_demand
                            success, message, excel_file_path, pdf_file_path = generate_reports_on_demand(
                                df_raw=df_raw,
                                config_data=config_data, # <--- ĐÃ CHỈNH SỬA
                                selected_mode=analysis_mode,
                                selected_year=selected_year,
                                selected_months=selected_months_standard,
                                selected_project_names_standard=selected_projects_standard,
                                comparison_config_years=[], # Không dùng cho báo cáo tiêu chuẩn
                                comparison_config_months=[], # Không dùng cho báo cáo tiêu chuẩn
                                comparison_config_projects=[], # Không dùng cho báo cáo tiêu chuẩn
                                comparison_report_mode=None, # Không dùng cho báo cáo tiêu chuẩn
                                path_dict=path_dict # <--- ĐÃ CHỈNH SỬA
                            )

                            if success:
                                st.success("Báo cáo đã tạo thành công!")
                                if excel_file_path and os.path.exists(excel_file_path):
                                    with open(excel_file_path, "rb") as file:
                                        st.download_button(
                                            label="Tải xuống Báo cáo Excel",
                                            data=file,
                                            file_name=os.path.basename(excel_file_path),
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            key='download_std_excel'
                                        )
                                if pdf_file_path and os.path.exists(pdf_file_path):
                                    with open(pdf_file_path, "rb") as file:
                                        st.download_button(
                                            label="Tải xuống Báo cáo PDF",
                                            data=file,
                                            file_name=os.path.basename(pdf_file_path),
                                            mime="application/pdf",
                                            key='download_std_pdf'
                                        )
                            else:
                                st.error(f"Có lỗi xảy ra khi tạo báo cáo. Vui lòng thử lại. {message}")

                        except Exception as e:
                            st.error(f"Có lỗi xảy ra khi tạo báo cáo. Vui lòng thử lại. {e}")
                            print(f"Error generating standard report: {e}")
    else:
        st.info("Chưa có dữ liệu thô được tải. Vui lòng tải lên file template của bạn.")


# ==============================================================================\
# TAB 2: BÁO CÁO SO SÁNH
# ==============================================================================\
with tab_comparison_report_main:
    st.header("Báo cáo So sánh")

    if df_raw is not None and config_data is not None:
        all_years = sorted(df_raw['Year'].unique().tolist(), reverse=True)
        all_months = ['January', 'February', 'March', 'April', 'May', 'June', 
                      'July', 'August', 'September', 'October', 'November', 'December']
        all_projects = sorted(df_raw['Project name'].unique().tolist())

        # Chọn chế độ so sánh
        comparison_mode_key = st.radio("Chọn Chế độ So sánh:", 
                                       ["So Sánh Dự Án Trong Một Tháng", 
                                        "So Sánh Dự Án Trong Một Năm", 
                                        "So Sánh Một Dự Án Qua Các Tháng/Năm"],
                                       key='comp_mode_radio')
        st.session_state.comparison_mode = comparison_mode_key # Cập nhật session state

        # Lựa chọn Năm cho so sánh
        if st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Tháng" or \
           st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Năm":
            comparison_selected_years_options = st.selectbox("Chọn Năm để So sánh:", 
                                                             ["Tất cả các Năm"] + all_years,
                                                             key='comp_year_select')
            st.session_state.comparison_selected_years = [comparison_selected_years_options] if comparison_selected_years_options != "Tất cả các Năm" else []
            if st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Tháng" and \
               len(st.session_state.comparison_selected_years) != 1:
                st.warning("Vui lòng chọn CHỈ MỘT năm khi so sánh các dự án trong một tháng hoặc cho một dự án qua các tháng.")
        elif st.session_state.comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
             comparison_selected_years_options = st.multiselect("Chọn Năm để So sánh:", 
                                                                all_years, 
                                                                default=st.session_state.comparison_selected_years,
                                                                key='comp_years_multi')
             st.session_state.comparison_selected_years = comparison_selected_years_options
             if len(st.session_state.comparison_selected_years) == 0 and len(st.session_state.comparison_selected_months) == 0:
                 st.warning("Vui lòng chọn ít nhất một năm hoặc một tháng để so sánh.")


        # Lựa chọn Tháng cho so sánh
        if st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Tháng":
            comparison_selected_months_options = st.selectbox("Chọn Tháng để So sánh:", 
                                                              all_months,
                                                              key='comp_month_select')
            st.session_state.comparison_selected_months = [comparison_selected_months_options] if comparison_selected_months_options else []
            if len(st.session_state.comparison_selected_months) != 1:
                st.warning("Vui lòng chọn CHỈ MỘT tháng khi so sánh các dự án trong một tháng.")
        elif st.session_state.comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
            # Chỉ cho phép chọn tháng nếu chỉ có 1 năm được chọn
            if len(st.session_state.comparison_selected_years) == 1:
                comparison_selected_months_options = st.multiselect("Chọn Tháng để So sánh:", 
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
            comparison_selected_projects_options = st.selectbox("Chọn Dự án để So sánh:", 
                                                                all_projects,
                                                                key='comp_project_select_single')
            st.session_state.comparison_selected_projects = [comparison_selected_projects_options] if comparison_selected_projects_options else []
            if len(st.session_state.comparison_selected_projects) != 1:
                st.warning("Đối với 'So Sánh Một Dự Án Qua Các Tháng/Năm', vui lòng chọn CHỈ MỘT dự án.")
        else:
            comparison_selected_projects_options = st.multiselect("Chọn Dự án để So sánh:", 
                                                                  all_projects, 
                                                                  default=st.session_state.comparison_selected_projects,
                                                                  key='comp_project_select_multi')
            st.session_state.comparison_selected_projects = comparison_selected_projects_options
            if not st.session_state.comparison_selected_projects:
                st.warning("Vui lòng chọn ít nhất một dự án để so sánh.")


        if st.button("Tạo Báo cáo So sánh", key='generate_comp_report_btn'):
            # Kiểm tra các điều kiện dựa trên chế độ so sánh
            can_generate = True
            if st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Tháng":
                if not st.session_state.comparison_selected_years or len(st.session_state.comparison_selected_years) != 1:
                    st.warning("Vui lòng chọn CHỈ MỘT năm khi so sánh các dự án trong một tháng hoặc cho một dự án qua các tháng.")
                    can_generate = False
                if not st.session_state.comparison_selected_months or len(st.session_state.comparison_selected_months) != 1:
                    st.warning("Vui lòng chọn CHỈ MỘT tháng khi so sánh các dự án trong một tháng.")
                    can_generate = False
                if not st.session_state.comparison_selected_projects:
                    st.warning("Vui lòng chọn ít nhất một dự án để so sánh.")
                    can_generate = False
            elif st.session_state.comparison_mode == "So Sánh Dự Án Trong Một Năm":
                if not st.session_state.comparison_selected_years or len(st.session_state.comparison_selected_years) != 1:
                    st.warning("Vui lòng chọn CHỈ MỘT năm khi so sánh các dự án trong một tháng hoặc cho một dự án qua các tháng.")
                    can_generate = False
                if not st.session_state.comparison_selected_projects:
                    st.warning("Vui lòng chọn ít nhất một dự án để so sánh.")
                    can_generate = False
            elif st.session_state.comparison_mode == "So Sánh Một Dự Án Qua Các Tháng/Năm":
                if not st.session_state.comparison_selected_projects or len(st.session_state.comparison_selected_projects) != 1:
                    st.warning("Đối với 'So Sánh Một Dự Án Qua Các Tháng/Năm', vui lòng chọn CHỈ MỘT dự án.")
                    can_generate = False
                if not st.session_state.comparison_selected_years and not st.session_state.comparison_selected_months:
                    st.warning("Vui lòng chọn ít nhất một năm hoặc một tháng để so sánh.")
                    can_generate = False
                if len(st.session_state.comparison_selected_years) > 1 and len(st.session_state.comparison_selected_months) > 0:
                     st.warning("Bạn không thể so sánh nhiều năm VÀ nhiều tháng cùng lúc. Vui lòng chọn chỉ năm HOẶC chỉ tháng.")
                     can_generate = False

            if can_generate:
                export_excel_comp, export_pdf_comp = False, False
                with st.expander("Tùy chọn Xuất"):
                    export_excel_comp = st.checkbox("Xuất ra Excel (.xlsx)", value=True, key='export_excel_comp')
                    export_pdf_comp = st.checkbox("Xuất ra PDF (.pdf)", value=True, key='export_pdf_comp')

                if not export_excel_comp and not export_pdf_comp:
                    st.warning("Vui lòng chọn ít nhất một định dạng xuất (Excel hoặc PDF).")
                else:
                    with st.spinner("Đang tạo báo cáo so sánh Excel..." if export_excel_comp else "Đang tạo báo cáo so sánh PDF..."):
                        try:
                            # Gọi hàm generate_reports_on_demand
                            success, message, excel_file_path, pdf_file_path = generate_reports_on_demand(
                                df_raw=df_raw,
                                config_data=config_data, # <--- ĐÃ CHỈNH SỬA
                                selected_mode=None, # Không dùng cho báo cáo so sánh
                                selected_year=None, # Không dùng cho báo cáo so sánh
                                selected_months=[], # Không dùng cho báo cáo so sánh
                                selected_project_names_standard=[], # Không dùng cho báo cáo so sánh
                                comparison_config_years=st.session_state.comparison_selected_years,
                                comparison_config_months=st.session_state.comparison_selected_months,
                                comparison_config_projects=st.session_state.comparison_selected_projects,
                                comparison_report_mode=st.session_state.comparison_mode,
                                path_dict=path_dict # <--- ĐÃ CHỈNH SỬA
                            )
                            
                            if success:
                                st.success("Báo cáo so sánh đã được tạo thành công!")
                                if excel_file_path and os.path.exists(excel_file_path):
                                    with open(excel_file_path, "rb") as file:
                                        st.download_button(
                                            label="Tải xuống Báo cáo So sánh Excel",
                                            data=file,
                                            file_name=os.path.basename(excel_file_path),
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            key='download_comp_excel'
                                        )
                                if pdf_file_path and os.path.exists(pdf_file_path):
                                    with open(pdf_file_path, "rb") as file:
                                        st.download_button(
                                            label="Tải xuống Báo cáo So sánh PDF",
                                            data=file,
                                            file_name=os.path.basename(pdf_file_path),
                                            mime="application/pdf",
                                            key='download_comp_pdf'
                                        )
                            else:
                                st.error(f"Có lỗi xảy ra khi tạo báo cáo. Vui lòng thử lại. {message}")

                        except Exception as e:
                            st.error(f"Có lỗi xảy ra khi tạo báo cáo. Vui lòng thử lại. {e}")
                            print(f"Error generating comparison report: {e}")
    else:
        st.info("Chưa có dữ liệu thô được tải. Vui lòng tải lên file template của bạn.")

# ==============================================================================\
# TAB 3: XEM TRƯỚC DỮ LIỆU
# ==============================================================================\
with tab_data_preview:
    st.header("Xem trước Dữ liệu Thô (100 hàng đầu tiên)")
    if df_raw is not None:
        st.dataframe(df_raw.head(100))
    else:
        st.info("Chưa có dữ liệu thô được tải. Vui lòng tải lên file template của bạn.")

# ==============================================================================\
# TAB 4: HƯỚNG DẪN
# ==============================================================================\
with tab_instructions:
    st.header("Hướng dẫn")
    st.write("Công cụ này giúp tự động hóa việc tạo báo cáo giờ làm việc.")

    st.markdown("### 1. Báo cáo Tiêu chuẩn")
    st.write("Tạo báo cáo chi tiết dựa trên năm, tháng và dự án đã chọn.")

    st.markdown("### 2. Báo cáo So sánh")
    st.write("Tạo phân tích so sánh dựa trên các chế độ khác nhau:")
    st.markdown("- **So Sánh Dự Án Trong Một Tháng:** So sánh nhiều dự án trong một tháng được chọn của một năm cụ thể.")
    st.markdown("- **So Sánh Dự Án Trong Một Năm:** So sánh nhiều dự án trong một năm được chọn (bao gồm tất cả các tháng).")
    st.markdown("- **So Sánh Một Dự Án Qua Các Tháng/Năm:** So sánh một dự án duy nhất qua nhiều tháng trong cùng một năm, HOẶC so sánh qua nhiều năm.")
    
    st.markdown("### 3. Xem trước Dữ liệu")
    st.write("Tab này cho phép bạn xem 100 hàng đầu tiên của dữ liệu thô đã tải, giúp bạn kiểm tra định dạng và nội dung dữ liệu.")

    st.markdown("### 4. Cấu hình File Template (Bên ngoài ứng dụng)")
    st.write("Công cụ đọc dữ liệu và cấu hình từ một file Excel template (thường là `Time_report.xlsm`). Đảm bảo rằng:")
    st.markdown("- Sheet 'Raw Data' chứa dữ liệu thời gian thô của bạn với các cột cần thiết như 'Year', 'MonthName', 'Project name', v.v.")
    st.markdown("- Sheet 'Config_Year_Mode' và 'Config_Project_Filter' có thể được sử dụng để đặt cấu hình mặc định, nhưng các lựa chọn trên giao diện sẽ ghi đè lên chúng.")

    st.markdown("### Lỗi thường gặp:")
    st.markdown("- **File template không tìm thấy:** Đảm bảo `Time_report.xlsm` nằm cùng thư mục với ứng dụng này.")
    st.markdown("- **Không tải được dữ liệu thô:** Kiểm tra định dạng dữ liệu và tên cột trong sheet 'Raw Data'.")
