import streamlit as st
import pandas as pd
import os
from datetime import datetime
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    apply_filters, export_report, export_pdf_report
)

# --- PAGE CONFIG ---
st.set_page_config(page_title="Triac Time Report", layout="wide", page_icon="📊")

# --- HEADER ---
st.markdown("""
    <style>
        .report-title {font-size: 32px; color: #003366; font-weight: bold; margin-bottom: 0;}
        .report-subtitle {font-size: 16px; color: gray; margin-top: 4px;}
        .block-container {padding-top: 1rem;}
        footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([0.12, 0.88])
with col1:
    st.image("triac_logo.png", width=110)
with col2:
    st.markdown("<p class='report-title'>Triac Time Report Generator</p>", unsafe_allow_html=True)
    st.markdown("<p class='report-subtitle'>Professional reporting tool for time tracking and analysis.</p>", unsafe_allow_html=True)

# --- LANGUAGE ---
translations = {
    "English": {
        "mode": "Select analysis mode:",
        "year": "Select year(s):",
        "month": "Select month(s):",
        "project": "Select project(s):",
        "report_button": "\ud83d\ude80 Generate report",
        "no_data": "\u26a0\ufe0f No data after filtering.",
        "report_done": "\u2705 Report created successfully",
        "download_excel": "\ud83d\udcc5 Download Excel",
        "download_pdf": "\ud83d\udcc4 Download PDF",
        "data_preview": "\ud83d\udcc2 Data preview",
        "user_guide": "\ud83d\udcd8 User Guide",
    },
    "Tiếng Việt": {
        "mode": "Chọn chế độ phân tích:",
        "year": "Chọn năm:",
        "month": "Chọn tháng:",
        "project": "Chọn dự án:",
        "report_button": "\ud83d\ude80 Tạo báo cáo",
        "no_data": "\u26a0\ufe0f Không có dữ liệu sau khi lọc.",
        "report_done": "\u2705 Đã tạo báo cáo",
        "download_excel": "\ud83d\udcc5 Tải Excel",
        "download_pdf": "\ud83d\udcc4 Tải PDF",
        "data_preview": "\ud83d\udcc2 Xem dữ liệu",
        "user_guide": "\ud83d\udcd8 Hướng dẫn sử dụng",
    }
}
lang = st.sidebar.selectbox("\ud83c\udf10 Language / Ngôn ngữ", ["English", "Tiếng Việt"])
T = translations[lang]

# --- PATHS ---
path_dict = setup_paths()

# --- LOAD DATA ---
@st.cache_data(ttl=1800)
def cached_load_raw_data():
    return load_raw_data(path_dict)

@st.cache_data(ttl=1800)
def cached_read_configs():
    return read_configs(path_dict)

with st.spinner("\ud83d\udd04 Loading data..."):
    df_raw = cached_load_raw_data()
    config_data = cached_read_configs()

# --- TABS ---
tab1, tab2, tab3 = st.tabs(["⚙️ Report Generator", T["data_preview"], T["user_guide"]])

with tab1:
    mode = st.selectbox(T["mode"], options=['year', 'month', 'week'], index=['year', 'month', 'week'].index(config_data['mode']))
    years = st.multiselect(T["year"], sorted(df_raw['Year'].dropna().unique()), default=[config_data['year']])
    months = st.multiselect(T["month"], list(df_raw['MonthName'].dropna().unique()), default=config_data['months'])

    project_df = config_data['project_filter_df']
    included_projects = project_df[project_df['Include'].str.lower() == 'yes']['Project Name'].tolist()
    project_selection = st.multiselect(T["project"], sorted(project_df['Project Name'].unique()), default=included_projects)

    st.markdown("---")
    if st.button(T["report_button"], use_container_width=True):
        with st.spinner("\ud83d\udcca Generating report..."):
            config = {
                'mode': mode,
                'years': years,
                'months': months,
                'project_filter_df': project_df[
                    project_df['Project Name'].isin(project_selection) &
                    (project_df['Include'].str.lower() == 'yes')
                ]
            }
            df_filtered = apply_filters(df_raw, config)
            if df_filtered.empty:
                st.warning(T["no_data"])
            else:
                export_report(df_filtered, config, path_dict)
                export_pdf_report(df_filtered, config, path_dict)
                st.success(f"{T['report_done']}: `{os.path.basename(path_dict['output_file'])}`")

                with open(path_dict['output_file'], "rb") as f:
                    st.download_button(T["download_excel"], f, file_name=os.path.basename(path_dict['output_file']), use_container_width=True)
                with open(path_dict['pdf_report'], "rb") as f:
                    st.download_button(T["download_pdf"], f, file_name=os.path.basename(path_dict['pdf_report']), use_container_width=True)

with tab2:
    st.subheader(T["data_preview"])
    st.dataframe(df_raw.head(100), use_container_width=True)

with tab3:
    st.markdown(f"### {T['user_guide']}")
    st.markdown("""
    - 🗂 Select filters: Mode, year, month, project
    - 🚀 Click **Generate report**
    - 📥 Download the Excel or PDF report from the buttons
    """)
