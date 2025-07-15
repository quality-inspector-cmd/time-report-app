import streamlit as st
import pandas as pd
import os
from datetime import datetime
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    apply_filters, export_report
)

# --- PAGE CONFIG ---
st.set_page_config(page_title="Triac Time Report", layout="wide", page_icon="📊")

# --- HEADER SECTION ---
st.markdown("""
    <style>
        .report-title {
            font-size: 32px;
            color: #003366;
            font-weight: bold;
            margin-bottom: 0;
        }
        .report-subtitle {
            font-size: 16px;
            color: gray;
            margin-top: 4px;
        }
        .block-container {
            padding-top: 1rem;
        }
        footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([0.12, 0.88])
with col1:
    st.image("triac_logo.png", width=110)
with col2:
    st.markdown("<p class='report-title'>Triac Time Report Generator</p>", unsafe_allow_html=True)
    st.markdown("<p class='report-subtitle'>Professional reporting tool for time tracking and analysis.</p>", unsafe_allow_html=True)

# --- LANGUAGE SELECT ---
translations = {
    "English": {
        "mode": "Select analysis mode:",
        "year": "Select year(s):",
        "month": "Select month(s):",
        "project": "Select project(s):",
        "report_button": "🚀 Generate report",
        "no_data": "⚠️ No data after filtering.",
        "report_done": "✅ Report created successfully",
        "download_excel": "📥 Download Excel report",
        "data_preview": "📂 Input data preview",
        "user_guide": "📘 User Guide",
    },
    "Tiếng Việt": {
        "mode": "Chọn chế độ phân tích:",
        "year": "Chọn năm:",
        "month": "Chọn tháng:",
        "project": "Chọn dự án:",
        "report_button": "🚀 Tạo báo cáo",
        "no_data": "⚠️ Không có dữ liệu sau khi lọc.",
        "report_done": "✅ Đã tạo báo cáo",
        "download_excel": "📥 Tải báo cáo Excel",
        "data_preview": "📂 Xem trước dữ liệu",
        "user_guide": "📘 Hướng dẫn sử dụng",
    }
}

lang = st.sidebar.selectbox("🌐 Language / Ngôn ngữ", ["English", "Tiếng Việt"])
T = translations[lang]

# --- LOAD DATA ---
path_dict = setup_paths()

if not os.path.exists(path_dict['template_file']):
    st.error(f"❌ Template file not found: {path_dict['template_file']}")
    st.stop()

@st.cache_data(ttl=60)
def cached_load_raw_data(path_dict):
    return load_raw_data(path_dict)

@st.cache_data(ttl=60)
def cached_read_configs(path_dict):
    return read_configs(path_dict)

with st.spinner("🔄 Loading data..."):
    df_raw = cached_load_raw_data(path_dict)
    config_data = cached_read_configs(path_dict)

# --- TABS ---
tab1, tab2, tab3 = st.tabs(["⚙️ Report Generator", T["data_preview"], T["user_guide"]])

# --- REPORT TAB ---
with tab1:
    mode = st.selectbox(T["mode"], options=['year', 'month', 'week'], index=['year', 'month', 'week'].index(config_data['mode']))
    all_years = sorted(df_raw['Year'].dropna().unique())
    years = st.multiselect(T["year"], options=all_years, default=[config_data['year']] if config_data['year'] else all_years)

    all_months = list(df_raw['MonthName'].dropna().unique())
    months = st.multiselect(T["month"], options=all_months, default=config_data['months'] if config_data['months'] else all_months)

    project_df = config_data['project_filter_df']
    included_projects = project_df[project_df['Include'].str.lower() == 'yes']['Project Name'].tolist()
    project_selection = st.multiselect(T["project"], options=sorted(project_df['Project Name'].unique()), default=included_projects)

    st.markdown("---")

    if st.button(T["report_button"], use_container_width=True):
        with st.spinner("📊 Generating report..."):
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
                st.success(f"{T['report_done']}: `{os.path.basename(path_dict['output_file'])}`")

                with open(path_dict['output_file'], "rb") as f:
                    st.download_button(T["download_excel"], data=f, file_name=os.path.basename(path_dict['output_file']), use_container_width=True)

# --- DATA PREVIEW TAB ---
with tab2:
    st.subheader(T["data_preview"])
    st.dataframe(df_raw.head(100), use_container_width=True)

# --- USER GUIDE TAB ---
with tab3:
    st.markdown(f"### {T['user_guide']}")
    st.markdown("""
    - 🗂 Select filters: Mode, year, month, project
    - 🚀 Click **Generate report**
    - 📥 Download the Excel report from the button
    """)
