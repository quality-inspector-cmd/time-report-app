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

# Khởi tạo trạng thái phiên (session state) nếu chưa có
if "user_email" not in st.session_state:
    st.session_state.user_email = None

# Màn hình xác thực email
if st.session_state.user_email is None:
    st.set_page_config(page_title="Triac Time Report", layout="wide")
    st.title("🔐 Access authentication")
    email_input = st.text_input("📧 Enter the invited email to access:")

    if email_input:
        email = email_input.strip().lower()
        # Đọc danh sách email từ invited_emails.csv
        try:
            invited_emails_df = pd.read_csv("invited_emails.csv")
            INVITED_EMAILS = invited_emails_df['email'].str.strip().str.lower().tolist()
        except FileNotFoundError:
            st.error("Lỗi: Không tìm thấy file 'invited_emails.csv'. Vui lòng kiểm tra lại.")
            st.stop()
        except Exception as e:
            st.error(f"Lỗi khi đọc file 'invited_emails.csv': {e}")
            st.stop()

        if email in INVITED_EMAILS:
            st.session_state.user_email = email
            st.success("✅ Email hợp lệ! Đang vào ứng dụng...")
            st.rerun() # Refresh app after successful login
        else:
            st.error("❌ Email không có trong danh sách mời.")
    st.stop() # Stop rendering the rest of the app if not authenticated

# ==============================================================================
# Phần chính của ứng dụng sau khi xác thực
# ==============================================================================

st.set_page_config(
    page_title="Triac Time Report",
    page_icon="⏰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Debug: Display current user email (can be removed in production)
# st.sidebar.write(f"Logged in as: {st.session_state.user_email}")

st.title("⏰ TRIAC Time Report Generator")

# --- Setup Paths and Load Data ---
path_dict = setup_paths()

try:
    df_raw = load_raw_data(path_dict)
    all_projects = sorted(df_raw['Project name'].dropna().unique().tolist())
    all_years = sorted(df_raw['Year'].dropna().unique().astype(int).tolist(), reverse=True)
    all_months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
except FileNotFoundError:
    st.error(f"Lỗi: Không tìm thấy file dữ liệu '{path_dict['template_file']}'. Vui lòng đảm bảo nó nằm trong cùng thư mục.")
    st.stop()
except KeyError as e:
    st.error(f"Lỗi dữ liệu: Cột '{e}' không tìm thấy trong file '{path_dict['template_file']}'. Vui lòng kiểm tra tên cột trong sheet 'Raw Data'.")
    st.stop()
except Exception as e:
    st.error(f"Đã xảy ra lỗi khi tải hoặc xử lý dữ liệu thô: {e}")
    st.stop()

# --- Sidebar for Navigation ---
st.sidebar.title("Navigation")
report_type = st.sidebar.radio("Chọn loại báo cáo:", ["Báo Cáo Tiêu Chuẩn", "Báo Cáo So Sánh"])

# --- Standard Report Section ---
if report_type == "Báo Cáo Tiêu Chuẩn":
    st.header("Báo Cáo Tiêu Chuẩn")

    st.subheader("Cấu hình báo cáo")
    col1, col2 = st.columns(2)
    with col1:
        selected_years = st.multiselect(
            "Chọn năm:",
            options=all_years,
            default=[datetime.now().year] if datetime.now().year in all_years else []
        )
    with col2:
        selected_months = st.multiselect(
            "Chọn tháng (để trống cho tất cả):",
            options=all_months,
            default=[]
        )

    st.subheader("Lọc dự án")
    # Tạo một DataFrame giả định cho Project Filter để người dùng nhập/chọn
    if 'project_filter_df' not in st.session_state:
        st.session_state.project_filter_df = pd.DataFrame(columns=['Project Name', 'Include'])

    # Hiển thị các dự án đã có trong df_raw
    st.write("Chọn các dự án để bao gồm trong báo cáo:")
    selected_projects_for_standard = st.multiselect(
        "Chọn dự án:",
        options=all_projects,
        default=st.session_state.project_filter_df['Project Name'].tolist() if not st.session_state.project_filter_df.empty else all_projects
    )

    # Cập nhật st.session_state.project_filter_df dựa trên lựa chọn
    temp_df = pd.DataFrame({'Project Name': selected_projects_for_standard})
    if not temp_df.empty:
        temp_df['Include'] = 'yes'
    st.session_state.project_filter_df = temp_df

    st.write("---")

    # Xác định mode báo cáo dựa trên lựa chọn năm và tháng
    report_mode = 'year'
    if selected_months:
        report_mode = 'month'
        if len(selected_months) == 12: # Nếu chọn đủ 12 tháng, coi như là báo cáo năm
            report_mode = 'year' # hoặc có thể gọi là annual-by-month

    # Tạo config dictionary
    config = {
        'mode': report_mode,
        'year': selected_years[0] if selected_years else None, # For single year mode
        'years': selected_years, # For multi-year mode
        'months': selected_months,
        'project_filter_df': st.session_state.project_filter_df
    }

    if st.button("Tạo Báo Cáo Tiêu Chuẩn"):
        if not selected_years:
            st.warning("Vui lòng chọn ít nhất một năm để tạo báo cáo.")
        elif st.session_state.project_filter_df.empty:
            st.warning("Vui lòng chọn ít nhất một dự án để tạo báo cáo.")
        else:
            try:
                # Áp dụng bộ lọc
                df_filtered = apply_filters(df_raw, config)

                if df_filtered.empty:
                    st.warning("Không có dữ liệu phù hợp với các tiêu chí lọc đã chọn. Vui lòng thử các lựa chọn khác.")
                else:
                    # Xuất báo cáo Excel
                    export_report(df_filtered, config, path_dict)
                    st.success(f"Đã tạo báo cáo Excel thành công: {path_dict['output_file']}")

                    # Xuất báo cáo PDF
                    export_pdf_report(df_filtered, config, path_dict)
                    st.success(f"Đã tạo báo cáo PDF thành công: {path_dict['pdf_report']}")

                    # Hiển thị nút tải về
                    st.download_button(
                        label="Tải xuống báo cáo Excel",
                        data=open(path_dict['output_file'], "rb").read(),
                        file_name=os.path.basename(path_dict['output_file']),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.download_button(
                        label="Tải xuống báo cáo PDF",
                        data=open(path_dict['pdf_report'], "rb").read(),
                        file_name=os.path.basename(path_dict['pdf_report']),
                        mime="application/pdf"
                    )
            except Exception as e:
                st.error(f"Đã xảy ra lỗi khi tạo báo cáo tiêu chuẩn: {e}")
                # Optional: st.exception(e) # để hiển thị traceback đầy đủ

# --- Comparison Report Section ---
elif report_type == "Báo Cáo So Sánh":
    st.header("Báo Cáo So Sánh")

    comparison_mode = st.selectbox(
        "Chọn chế độ so sánh:",
        [
            "So Sánh Dự Án Trong Một Tháng",
            "So Sánh Dự Án Trong Một Năm",
            "So Sánh Một Dự Án Qua Các Tháng/Năm"
        ]
    )

    st.subheader("Cấu hình so sánh")
    col1, col2, col3 = st.columns(3)
    with col1:
        comp_years = st.multiselect("Chọn năm:", options=all_years, default=[datetime.now().year] if datetime.now().year in all_years else [])
    with col2:
        comp_months = st.multiselect("Chọn tháng:", options=all_months, default=[])
    with col3:
        comp_projects = st.multiselect("Chọn dự án:", options=all_projects, default=[])

    comparison_config = {
        'years': comp_years,
        'months': comp_months,
        'selected_projects': comp_projects
    }

    st.write("---")

    if st.button("Tạo Báo Cáo So Sánh"):
        if not comp_years or not comp_months or not comp_projects:
             st.warning("Vui lòng chọn đủ Năm, Tháng và Dự án cho báo cáo so sánh.")
        else:
            try:
                df_comparison, message = apply_comparison_filters(df_raw, comparison_config, comparison_mode)

                if df_comparison.empty:
                    st.warning(message)
                else:
                    st.subheader("Dữ liệu so sánh:")
                    st.dataframe(df_comparison)

                    # Xuất báo cáo Excel so sánh
                    export_comparison_report(df_comparison, comparison_config, path_dict, comparison_mode)
                    st.success(f"Đã tạo báo cáo Excel so sánh thành công: {path_dict['comparison_output_file']}")

                    # Xuất báo cáo PDF so sánh
                    export_comparison_pdf_report(df_comparison, comparison_config, path_dict, comparison_mode)
                    st.success(f"Đã tạo báo cáo PDF so sánh thành công: {path_dict['comparison_pdf_report']}")

                    # Nút tải về báo cáo so sánh
                    st.download_button(
                        label="Tải xuống báo cáo Excel so sánh",
                        data=open(path_dict['comparison_output_file'], "rb").read(),
                        file_name=os.path.basename(path_dict['comparison_output_file']),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.download_button(
                        label="Tải xuống báo cáo PDF so sánh",
                        data=open(path_dict['comparison_pdf_report'], "rb").read(),
                        file_name=os.path.basename(path_dict['comparison_pdf_report']),
                        mime="application/pdf"
                    )
            except Exception as e:
                st.error(f"Đã xảy ra lỗi khi tạo báo cáo so sánh: {e}")
                # Optional: st.exception(e)
