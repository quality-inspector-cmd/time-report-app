import streamlit as st
import pandas as pd
import datetime
import os

# ĐẢM BẢO TÊN FILE LOGIC DƯỚI ĐÂY CHÍNH XÁC VỚI FILE BẠN ĐÃ LƯU
# VÀ NÓ NẰM CÙNG THƯ MỤC VỚI main_optimized.py
# ====================================================================
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    apply_filters, export_report, export_pdf_report,
    apply_comparison_filters, export_comparison_report, export_comparison_pdf_report
)
# ====================================================================

st.set_page_config(layout="wide", page_title="TRIAC Time Report App")

# Cài đặt paths
path_dict = setup_paths()

# Header
st.title("TRIAC Time Report Dashboard")
st.markdown("---")

# Cố gắng đọc cấu hình và dữ liệu thô
try:
    config = read_configs(path_dict)
    df_raw = load_raw_data(path_dict)
    st.sidebar.success("Đã tải dữ liệu và cấu hình thành công!")
except Exception as e:
    st.sidebar.error(f"Lỗi khi tải dữ liệu hoặc cấu hình: {e}")
    st.stop() # Dừng ứng dụng nếu không thể tải dữ liệu ban đầu

# Tạo sidebar cho các tùy chọn lọc
st.sidebar.header("Cấu hình báo cáo")

report_type = st.sidebar.radio(
    "Chọn loại báo cáo:",
    ("Báo cáo tiêu chuẩn", "Báo cáo so sánh")
)

# Lấy danh sách duy nhất của các năm và tháng từ dữ liệu thô
all_years = sorted(df_raw['Year'].unique().tolist())
all_months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
all_projects = sorted(df_raw['Project name'].unique().tolist())

if report_type == "Báo cáo tiêu chuẩn":
    st.sidebar.markdown("### Lọc cho Báo cáo tiêu chuẩn")
    mode_options = ["year", "month", "week"]
    selected_mode = st.sidebar.selectbox("Chế độ báo cáo", mode_options, index=mode_options.index(config.get('mode', 'year')))

    selected_year_standard = st.sidebar.selectbox("Chọn năm", all_years, index=all_years.index(config.get('year', datetime.datetime.now().year)) if config.get('year', datetime.datetime.now().year) in all_years else 0)
    
    selected_months_standard = []
    if selected_mode == "month":
        selected_months_standard = st.sidebar.multiselect(
            "Chọn tháng (chỉ áp dụng cho chế độ tháng)", 
            all_months, 
            default=config.get('months', [])
        )
    
    # Sử dụng 'selected_projects_config' từ biến config chính
    default_selected_projects_standard = config.get('selected_projects_config', [])
    selected_projects_standard = st.sidebar.multiselect(
        "Lọc theo dự án (từ Config_Project_Filter)", 
        all_projects, 
        default=default_selected_projects_standard
    )

    format_options_standard = st.sidebar.multiselect(
        "Chọn định dạng xuất báo cáo", 
        ["Excel", "PDF"], 
        default=["Excel"]
    )

    if st.sidebar.button("Tạo Báo Cáo Tiêu Chuẩn"):
        if not format_options_standard:
            st.warning("Vui lòng chọn ít nhất một định dạng xuất báo cáo.")
        else:
            with st.spinner("Đang tạo báo cáo tiêu chuẩn..."):
                current_config = {
                    'mode': selected_mode,
                    'year': selected_year_standard,
                    'months': selected_months_standard,
                    # Đảm bảo selected_projects_config luôn được truyền vào
                    'selected_projects_config': selected_projects_standard 
                }
                
                df_filtered = apply_filters(df_raw, current_config)
                
                if df_filtered.empty:
                    st.warning("Không có dữ liệu phù hợp với các lựa chọn lọc hiện tại.")
                else:
                    if "Excel" in format_options_standard:
                        export_report(df_filtered, current_config, path_dict)
                        st.success(f"Báo cáo tiêu chuẩn Excel đã được tạo tại {path_dict['output_file']}")
                    
                    if "PDF" in format_options_standard:
                        export_pdf_report(df_filtered, current_config, path_dict)
                        st.success(f"Báo cáo tiêu chuẩn PDF đã được tạo tại {path_dict['pdf_report']}")

elif report_type == "Báo cáo so sánh":
    st.sidebar.markdown("### Lọc cho Báo cáo so sánh")
    comparison_mode_options = [
        "So Sánh Dự Án Trong Một Tháng",
        "So Sánh Dự Án Trong Một Năm",
        "So Sánh Một Dự Án Qua Các Tháng/Năm"
    ]
    comparison_mode = st.sidebar.selectbox("Chọn chế độ so sánh", comparison_mode_options)

    selected_years_comparison = st.sidebar.multiselect("Chọn năm", all_years, default=[datetime.datetime.now().year])
    selected_months_comparison = st.sidebar.multiselect("Chọn tháng", all_months)
    
    # Sử dụng 'selected_projects_config' từ biến config chính
    default_selected_projects_comparison = config.get('selected_projects_config', [])
    selected_projects_comparison = st.sidebar.multiselect(
        "Chọn dự án để so sánh (từ Config_Project_Filter)", 
        all_projects, 
        default=default_selected_projects_comparison
    )

    format_options_comparison = st.sidebar.multiselect(
        "Chọn định dạng xuất báo cáo so sánh", 
        ["Excel", "PDF"], 
        default=["Excel"]
    )

    if st.sidebar.button("Tạo Báo Cáo So Sánh"):
        if not format_options_comparison:
            st.warning("Vui lòng chọn ít nhất một định dạng xuất báo cáo so sánh.")
        else:
            with st.spinner("Đang tạo báo cáo so sánh..."):
                # Đảm bảo comparison_config_for_functions chứa 'selected_projects_config'
                comparison_config_for_functions = {
                    'years': selected_years_comparison,
                    'months': selected_months_comparison,
                    'selected_projects_config': selected_projects_comparison # Đảm bảo dòng này
                }
                
                df_comparison, message = apply_comparison_filters(df_raw, comparison_config_for_functions, comparison_mode)
                
                if df_comparison.empty:
                    st.warning(message)
                else:
                    if "Excel" in format_options_comparison:
                        export_success_excel = export_comparison_report(df_comparison, comparison_config_for_functions, path_dict, comparison_mode)
                        if export_success_excel:
                            st.success(f"Báo cáo so sánh Excel đã được tạo tại {path_dict['comparison_output_file']}")
                        else:
                            st.error("Có lỗi xảy ra khi tạo báo cáo so sánh Excel.")
                    
                    if "PDF" in format_options_comparison:
                        export_success_pdf = export_comparison_pdf_report(df_comparison, comparison_config_for_functions, path_dict, comparison_mode)
                        if export_success_pdf:
                            st.success(f"Báo cáo so sánh PDF đã được tạo tại {path_dict['comparison_pdf_report']}")
                        else:
                            st.error("Có lỗi xảy ra khi tạo báo cáo so sánh PDF.")

# Hiển thị DataFrame thô (tùy chọn, chỉ để debug)
# st.subheader("Dữ liệu thô (Chỉ để debug)")
# st.dataframe(df_raw.head())

# st.subheader("Cấu hình đã đọc (Chỉ để debug)")
# st.json(config)
