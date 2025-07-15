import streamlit as st
import pandas as pd
import os
from datetime import datetime
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    apply_filters, export_report, export_all_charts_to_pdf
)

# Set page and branding
st.set_page_config(page_title="Time Report Generator", layout="wide")
col1, col2 = st.columns([0.15, 0.85])
with col1:
    st.image("triac_logo.png", width=120)
with col2:
    st.markdown("""
    <div style='padding-top:10px'>
        <h1 style='color:#003366; font-size:32px;'>📊 Time Report Generator</h1>
        <p style='color:gray; font-size:16px;'>Generate customized time tracking reports for Triac Composites.</p>
    </div>
    """, unsafe_allow_html=True)

# Hide Streamlit footer
st.markdown("""
    <style>
        footer {visibility: hidden;}
        .block-container {padding-top: 2rem;}
    </style>
""", unsafe_allow_html=True)

# Multilingual dictionary
translations = {
    "English": {
        "mode": "Select analysis mode:",
        "year": "Select year(s):",
        "month": "Select month(s):",
        "project": "Select project(s):",
        "report_button": "🚀 Generate report",
        "no_data": "⚠️ No data after filtering.",
        "report_done": "✅ Report created",
        "download_excel": "📥 Download Excel report",
        "download_pdf": "📥 Download PDF report",
        "data_preview": "📂 Input data (first 100 rows)",
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
        "data_preview": "📂 Dữ liệu đầu vào (100 dòng đầu)",
        "user_guide": "📘 Hướng dẫn sử dụng",
    }
}

lang = st.sidebar.selectbox("🌐 Language / Ngôn ngữ", ["English", "Tiếng Việt"])
T = translations[lang]

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

tab1, tab2, tab3 = st.tabs(["⚙️ Report", T["data_preview"], T["user_guide"]])

with tab1:
    mode = st.selectbox(T["mode"], options=['year', 'month', 'week'], index=['year', 'month', 'week'].index(config_data['mode']))
    all_years = sorted(df_raw['Year'].dropna().unique())
    years = st.multiselect(T["year"], options=all_years, default=[config_data['year']] if config_data['year'] else all_years)
    all_months = list(df_raw['MonthName'].dropna().unique())
    months = st.multiselect(T["month"], options=all_months, default=config_data['months'] if config_data['months'] else all_months)

    project_df = config_data['project_filter_df']
    included_projects = project_df[project_df['Include'].str.lower() == 'yes']['Project Name'].tolist()
    project_selection = st.multiselect(T["project"], options=sorted(project_df['Project Name'].unique()), default=included_projects)

    if st.button(T["report_button"]):
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
                export_all_charts_to_pdf(path_dict)
                st.success(f"✅ {T['report_done']}: {os.path.basename(path_dict['output_file'])}")

                with open(path_dict['output_file'], "rb") as f:
                    st.download_button(T["download_excel"], data=f, file_name=os.path.basename(path_dict['output_file']))
                with open(path_dict['pdf_report'], "rb") as f:
                    st.download_button(T["download_pdf"], data=f, file_name=os.path.basename(path_dict['pdf_report']))

with tab2:
    st.subheader(T["data_preview"])
    with st.expander("Click to view raw data sample"):
        st.dataframe(df_raw.head(100))

with tab3:
    st.markdown(f"### {T['user_guide']}")
    st.markdown("""
    - 📤 Upload `Raw_data.xlsx` and `Time_report.xlsm`
    - 🗂 Chọn bộ lọc: Chế độ, năm, tháng, dự án
    - 🚀 Nhấn **Tạo báo cáo**
    - 📥 Tải báo cáo Excel hoặc PDF
    """)
