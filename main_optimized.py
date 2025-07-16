import streamlit as st
import pandas as pd
import os
from datetime import datetime
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    apply_filters, export_report, export_pdf_report
)
script_dir = os.path.dirname(__file__) 
csv_file_path = os.path.join(script_dir, "invited_emails.csv")

@st.cache_data
def load_invited_emails():
    try:
        df = pd.read_csv(csv_file_path) 
        print(f"DEBUG: File path being read: {csv_file_path}")
        print(f"DEBUG: Columns detected by pandas: {df.columns.tolist()}")
        print(f"DEBUG: First 5 rows of DataFrame:\n{df.head()}")
        if "email" in df.columns:
            emails = df["email"].astype(str).str.strip().str.lower().tolist()
            print(f"DEBUG: Loaded invited emails list: {emails}") # In ra danh sách email đã xử lý
            return emails
        else:
            print("ERROR: 'email' column not found in invited_emails.csv.")
            st.error("Lỗi: Không tìm thấy cột 'email' trong file invited_emails.csv. Vui lòng đảm bảo file có header 'email' (viết thường).")
            return []
    except FileNotFoundError:
        print(f"ERROR: invited_emails.csv not found at {csv_file_path}")
        st.error(f"Lỗi: Không tìm thấy file invited_emails.csv tại {csv_file_path}. Vui lòng kiểm tra đường dẫn.")
        return []
    except Exception as e:
        print(f"ERROR: An error occurred while loading invited_emails.csv: {e}")
        st.error(f"Lỗi khi tải file invited_emails.csv: {e}")
        return []

INVITED_EMAILS = load_invited_emails()

if "user_email" not in st.session_state:
    st.set_page_config(page_title="Triac Time Report", layout="wide")
    st.title("🔐 Access authentication")
    email_input = st.text_input("📧 Enter the invited email to access:")

    if email_input:
        email = email_input.strip().lower()
        print(f"DEBUG: User input email (processed): '{email}'")
        if email in INVITED_EMAILS:
            st.session_state.user_email = email
            log_user_access(email)
            st.success("✅ Email hợp lệ! Đang vào ứng dụng...")
            st.experimental_rerun()
        else:
            st.error("❌ Email không có trong danh sách mời.")
    st.stop()

# ---------------------------
# ---------------------------

st.set_page_config(page_title="Triac Time Report", layout="wide")

st.markdown("""
    <style>
        .report-title {font-size: 30px; color: #003366; font-weight: bold;}
        .report-subtitle {font-size: 14px; color: gray;}
        footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([0.12, 0.88])
with col1:
    st.image("triac_logo.png", width=110)
with col2:
    st.markdown("<div class='report-title'>Triac Time Report Generator</div>", unsafe_allow_html=True)
    st.markdown("<div class='report-subtitle'>Reporting tool for time tracking and analysis</div>", unsafe_allow_html=True)

translations = {
    "English": {
        "mode": "Select mode",
        "year": "Select year(s)",
        "month": "Select month(s)",
        "project": "Select project(s)",
        "report_button": "Generate report",
        "no_data": "No data after filtering",
        "report_done": "Report created successfully",
        "download_excel": "Download Excel",
        "download_pdf": "Download PDF",
        "data_preview": "Data preview",
        "user_guide": "User Guide"
    },
    "Tiếng Việt": {
        "mode": "Chọn chế độ",
        "year": "Chọn năm",
        "month": "Chọn tháng",
        "project": "Chọn dự án",
        "report_button": "Tạo báo cáo",
        "no_data": "Không có dữ liệu sau khi lọc",
        "report_done": "Đã tạo báo cáo",
        "download_excel": "Tải Excel",
        "download_pdf": "Tải PDF",
        "data_preview": "Xem dữ liệu",
        "user_guide": "Hướng dẫn"
    }
}

lang = st.sidebar.selectbox("Language / Ngôn ngữ", ["English", "Tiếng Việt"])
T = translations[lang]

path_dict = setup_paths()

@st.cache_data(ttl=1800)
def cached_load():
    return load_raw_data(path_dict), read_configs(path_dict)

with st.spinner("Loading data..."):
    df_raw, config_data = cached_load()

tab1, tab2, tab3 = st.tabs(["Report", T["data_preview"], T["user_guide"]])

with tab1:
    mode = st.selectbox(T["mode"], ['year', 'month', 'week'], index=['year', 'month', 'week'].index(config_data['mode']))
    years = st.multiselect(T["year"], sorted(df_raw['Year'].dropna().unique()), default=[config_data['year']])
    months = st.multiselect(T["month"], df_raw['MonthName'].dropna().unique(), default=config_data['months'])

    project_df = config_data['project_filter_df']
    included = project_df[project_df['Include'].str.lower() == 'yes']['Project Name'].tolist()
    selected_projects = st.multiselect(T["project"], sorted(project_df['Project Name'].unique()), default=included)

    if st.button(T["report_button"], use_container_width=True):
        with st.spinner("Generating report..."):
            config = {
                'mode': mode,
                'years': years,
                'months': months,
                'project_filter_df': project_df[
                    project_df['Project Name'].isin(selected_projects) &
                    (project_df['Include'].str.lower() == 'yes')
                ]
            }
            df_filtered = apply_filters(df_raw, config)
            if df_filtered.empty:
                st.warning(T["no_data"])
            else:
                export_report(df_filtered, config, path_dict)
                export_pdf_report(df_filtered, config, path_dict)
                st.success(f"{T['report_done']}: {os.path.basename(path_dict['output_file'])}")

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
    - Chọn bộ lọc: chế độ, năm, tháng, dự án
    - Nhấn "Tạo báo cáo"
    - Tải về Excel hoặc PDF
    """)
