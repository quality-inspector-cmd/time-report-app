import streamlit as st
import pandas as pd
import os
from datetime import datetime
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    apply_filters, export_report, export_summary_pdf
)

st.set_page_config(page_title="Triac Time Report", layout="wide", page_icon="📊")

# --- HEADER ---
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

# --- TRANSLATION ---
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
        "download_pdf": "📥 Download PDF report",
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
        "download_pdf": "📥 Tải báo cáo PDF",
        "data_preview": "📂 Xem trước dữ liệu",
        "user_guide": "📘 Hướng dẫn sử dụng",
    }
}

lang = st.sidebar.selectbox("🌐 Language / Ngôn ngữ", ["English", "Tiếng Việt"])
T = translations[lang]

# --- LOAD DATA ---
path_dict = setup_paths()

@st.cache_data(ttl=3600)
def cached_load():
    return load_raw_data(path_dict), read_configs(path_dict)

with st.spinner("🔄 Loading data..."):
    df_raw, config_data = cached_load()

# --- UI ---
tab1, tab2, tab3 = st.tabs(["⚙️ Report", T["data_preview"], T["user_guide"]])

with tab1:
    mode = st.selectbox(T["mode"], options=['year', 'month', 'week'], index=['year', 'month', 'week'].index(config_data['mode']))
    years = st.multiselect(T["year"], sorted(df_raw['Year'].dropna().unique()), default=[config_data['year']] if config_data['year'] else [])
    months = st.multiselect(T["month"], list(df_raw['MonthName'].dropna().unique()), default=config_data['months'])

    project_df = config_data['project_filter_df']
    included_projects = project_df[project_df['Include'].str.lower() == 'yes']['Project Name'].tolist()
    project_selection = st.multiselect(T["project"], sorted(project_df['Project Name'].unique()), default=included_projects)

    if st.button(T["report_button"], use_container_width=True):
        with st.spinner("📊 Generating report..."):
            config = {
                'mode': mode,
                'years': years,
                'months': months,
                'project_filter_df': project_df[project_df['Project Name'].isin(project_selection) & (project_df['Include'].str.lower() == 'yes')]
            }
            df_filtered = apply_filters(df_raw, config)
            if df_filtered.empty:
                st.warning(T["no_data"])
            else:
                export_report(df_filtered, config, path_dict)
                export_summary_pdf(df_filtered, config, path_dict)

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
    - 📥 Download the Excel or PDF report
    """)
