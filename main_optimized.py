import streamlit as st
import pandas as pd
from datetime import datetime
import base64
import plotly.graph_objects as go

# Import functions from the separate file
from a04ecaf1_1dae_4c90_8081_086cd7c7b725.py import (
    load_and_preprocess_data,
    get_unique_values,
    get_min_max_years,
    calculate_monthly_summary,
    calculate_project_summary,
    create_monthly_summary_chart,
    create_project_summary_chart,
    create_raw_data_table,
    create_overall_summary_table,
    apply_filters,
    apply_comparison_filters,
    create_comparison_chart,
    create_comparison_table,
    export_pdf_report,
    MONTH_NAMES_EN # Import month names for consistent ordering
)

# --- Streamlit UI Setup ---
st.set_page_config(layout="wide", page_title="Báo cáo giờ làm việc dự án")

# Custom CSS for styling
st.markdown("""
<style>
    .reportview-container .main .block-container {
        padding-top: 2rem;
        padding-right: 2rem;
        padding-left: 2rem;
        padding-bottom: 2rem;
    }
    .css-1d391kg e16zpvmm1 { /* This targets the main content area */
        padding-top: 1rem;
        padding-right: 1rem;
        padding-left: 1rem;
        padding-bottom: 1rem;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        padding: 10px 20px;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        font-size: 16px;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .stDownloadButton>button {
        background-color: #008CBA;
        color: white;
        padding: 10px 20px;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        font-size: 16px;
    }
    .stDownloadButton>button:hover {
        background-color: #007bb5;
    }
    h1, h2, h3 {
        color: #2E86C1; /* A nice blue for headers */
    }
    .stTextInput>div>div>input {
        border-radius: 5px;
        border: 1px solid #ccc;
        padding: 8px;
    }
    .stMultiSelect div[data-baseweb="select"] {
        border-radius: 5px;
        border: 1px solid #ccc;
    }
    .stCheckbox > label {
        font-size: 1rem;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Ứng dụng báo cáo giờ làm việc dự án")

# --- Session State Initialization ---
if 'df_raw' not in st.session_state:
    st.session_state['df_raw'] = None
if 'df_filtered' not in st.session_state:
    st.session_state['df_filtered'] = None
if 'selected_year' not in st.session_state:
    st.session_state['selected_year'] = None
if 'selected_month_name' not in st.session_state:
    st.session_state['selected_month_name'] = None
if 'selected_project_name' not in st.session_state:
    st.session_state['selected_project_name'] = None
if 'comparison_config' not in st.session_state:
    st.session_state['comparison_config'] = {'years': [], 'months': [], 'selected_projects': []}
if 'comparison_mode' not in st.session_state:
    st.session_state['comparison_mode'] = "So Sánh Dự Án Trong Một Tháng"

# --- File Upload ---
uploaded_file = st.sidebar.file_uploader("Tải lên file Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    with st.spinner("Đang tải và xử lý dữ liệu..."):
        st.session_state['df_raw'] = load_and_preprocess_data(uploaded_file)
    st.sidebar.success("Tải file thành công!")

if st.session_state['df_raw'] is not None and not st.session_state['df_raw'].empty:
    df_raw = st.session_state['df_raw']

    # --- Sidebar Filters ---
    st.sidebar.header("Bộ lọc")
    
    unique_years = get_unique_values(df_raw, 'Year')
    min_year, max_year = get_min_max_years(df_raw)
    
    # Add an "All Years" option
    all_years_option = "Tất cả các năm"
    year_options = [all_years_option] + unique_years
    
    selected_year_filter = st.sidebar.selectbox("Chọn năm", year_options, 
                                                index=year_options.index(st.session_state['selected_year']) if st.session_state['selected_year'] in year_options else 0)
    
    if selected_year_filter == all_years_option:
        st.session_state['selected_year'] = None
        unique_months = MONTH_NAMES_EN # All months if no specific year is selected
    else:
        st.session_state['selected_year'] = selected_year_filter
        # Filter months based on the selected year
        df_for_month_filter = df_raw[df_raw['Year'] == st.session_state['selected_year']]
        unique_months = [m for m in MONTH_NAMES_EN if m in df_for_month_filter['MonthName'].unique()]


    all_months_option = "Tất cả các tháng"
    month_options = [all_months_option] + unique_months
    selected_month_filter = st.sidebar.selectbox("Chọn tháng", month_options,
                                                 index=month_options.index(st.session_state['selected_month_name']) if st.session_state['selected_month_name'] in month_options else 0)

    if selected_month_filter == all_months_option:
        st.session_state['selected_month_name'] = None
    else:
        st.session_state['selected_month_name'] = selected_month_filter

    all_projects_option = "Tất cả các dự án"
    unique_projects = get_unique_values(df_raw, 'Project name')
    project_options = [all_projects_option] + unique_projects
    selected_project_filter = st.sidebar.selectbox("Chọn dự án", project_options,
                                                   index=project_options.index(st.session_state['selected_project_name']) if st.session_state['selected_project_name'] in project_options else 0)

    if selected_project_filter == all_projects_option:
        st.session_state['selected_project_name'] = None
    else:
        st.session_state['selected_project_name'] = selected_project_filter

    # Apply main filters
    st.session_state['df_filtered'] = apply_filters(
        df_raw,
        st.session_state['selected_year'],
        st.session_state['selected_month_name'],
        st.session_state['selected_project_name']
    )

    df_filtered = st.session_state['df_filtered']

    # --- Main Content Tabs ---
    tab1, tab2 = st.tabs(["Tổng quan báo cáo", "So sánh"])

    with tab1:
        st.header("Tổng quan báo cáo")

        if df_filtered.empty:
            st.warning("Không có dữ liệu để hiển thị với các bộ lọc đã chọn.")
        else:
            overall_summary_fig = create_overall_summary_table(df_filtered)
            st.plotly_chart(overall_summary_fig, use_container_width=True)

            monthly_summary = calculate_monthly_summary(df_filtered)
            monthly_summary_fig = create_monthly_summary_chart(monthly_summary, st.session_state['selected_year'] if st.session_state['selected_year'] else "Tất cả các năm")
            st.plotly_chart(monthly_summary_fig, use_container_width=True)

            project_summary = calculate_project_summary(df_filtered)
            project_summary_fig = create_project_summary_chart(project_summary, st.session_state['selected_year'] if st.session_state['selected_year'] else "Tất cả các năm")
            st.plotly_chart(project_summary_fig, use_container_width=True)

            # Raw data table is often very large, consider only showing a sample or making it expandable
            st.subheader("Dữ liệu thô đã lọc")
            # Using st.dataframe for better interactive table display in Streamlit
            st.dataframe(df_filtered[['Project name', 'Start date', 'End date', 'Hours', 'Total cost (USD)']].style.format({
                'Hours': "{:,.0f}",
                'Total cost (USD)': "{:,.2f} USD",
                'Start date': lambda x: x.strftime('%Y-%m-%d') if pd.notnull(x) else '',
                'End date': lambda x: x.strftime('%Y-%m-%d') if pd.notnull(x) else ''
            }), use_container_width=True)

            st.markdown("---")
            st.subheader("Xuất báo cáo PDF")

            pdf_report = export_pdf_report(
                overall_summary_fig, monthly_summary_fig, project_summary_fig, 
                create_raw_data_table(df_filtered), # Pass a figure for the raw data, even if not directly used in PDF for now
                go.Figure(), # Empty figure for comparison chart
                go.Figure(), # Empty figure for comparison table
                st.session_state['selected_year'], 
                st.session_state['selected_month_name'], 
                st.session_state['selected_project_name'],
                "", # No comparison mode text for full report
                "full_report"
            )
            
            download_filename = f"BaoCaoTongQuan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            st.download_button(
                label="Tải xuống báo cáo tổng quan PDF",
                data=pdf_report,
                file_name=download_filename,
                mime="application/pdf"
            )

    with tab2:
        st.header("So sánh dữ liệu")

        comparison_mode_options = {
            "So Sánh Dự Án Trong Một Tháng": "Compare Projects in a Month",
            "So Sánh Dự Án Trong Một Năm": "Compare Projects in a Year",
            "So Sánh Một Dự Án Qua Các Tháng/Năm": "Compare One Project Over Time (Months/Years)"
        }
        
        selected_comparison_mode_text = st.selectbox(
            "Chọn chế độ so sánh",
            list(comparison_mode_options.keys()),
            key='comparison_mode_select',
            index=list(comparison_mode_options.keys()).index(st.session_state['comparison_mode'])
        )
        st.session_state['comparison_mode'] = selected_comparison_mode_text
        current_comparison_mode = comparison_mode_options[selected_comparison_mode_text]

        st.markdown("---")
        st.subheader("Cấu hình so sánh")

        # Dynamic filter options based on comparison mode
        if current_comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
            st.info("Chọn MỘT năm, MỘT tháng và ít nhất HAI dự án.")
            
            # Year selection - must be single
            comp_year_options = get_unique_values(df_raw, 'Year')
            comp_selected_years = st.multiselect("Chọn Năm (chỉ 1)", comp_year_options, 
                                                 default=st.session_state['comparison_config']['years'],
                                                 key='comp_years_month_project')
            if len(comp_selected_years) > 1:
                st.warning("Vui lòng chọn CHỈ MỘT năm cho chế độ này.")
                comp_selected_years = comp_selected_years[:1] # Keep only the first one
            st.session_state['comparison_config']['years'] = comp_selected_years
            
            # Month selection - must be single
            if comp_selected_years:
                df_for_comp_month = df_raw[df_raw['Year'].isin(comp_selected_years)]
                comp_month_options = [m for m in MONTH_NAMES_EN if m in df_for_comp_month['MonthName'].unique()]
            else:
                comp_month_options = MONTH_NAMES_EN # Show all if no year is selected yet

            comp_selected_months = st.multiselect("Chọn Tháng (chỉ 1)", comp_month_options, 
                                                  default=st.session_state['comparison_config']['months'],
                                                  key='comp_months_month_project')
            if len(comp_selected_months) > 1:
                st.warning("Vui lòng chọn CHỈ MỘT tháng cho chế độ này.")
                comp_selected_months = comp_selected_months[:1] # Keep only the first one
            st.session_state['comparison_config']['months'] = comp_selected_months

            # Projects selection - multiple
            comp_project_options = get_unique_values(df_raw, 'Project name')
            comp_selected_projects = st.multiselect("Chọn Các Dự Án (ít nhất 2)", comp_project_options, 
                                                    default=st.session_state['comparison_config']['selected_projects'],
                                                    key='comp_projects_month_project')
            if len(comp_selected_projects) < 2:
                st.warning("Vui lòng chọn ít nhất HAI dự án cho chế độ này.")
            st.session_state['comparison_config']['selected_projects'] = comp_selected_projects

        elif current_comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
            st.info("Chọn MỘT năm và ít nhất HAI dự án. Dữ liệu sẽ được so sánh theo tháng.")
            
            # Year selection - must be single
            comp_year_options = get_unique_values(df_raw, 'Year')
            comp_selected_years = st.multiselect("Chọn Năm (chỉ 1)", comp_year_options, 
                                                 default=st.session_state['comparison_config']['years'],
                                                 key='comp_years_year_project')
            if len(comp_selected_years) > 1:
                st.warning("Vui lòng chọn CHỈ MỘT năm cho chế độ này.")
                comp_selected_years = comp_selected_years[:1]
            st.session_state['comparison_config']['years'] = comp_selected_years
            st.session_state['comparison_config']['months'] = [] # Clear months for this mode

            # Projects selection - multiple
            comp_project_options = get_unique_values(df_raw, 'Project name')
            comp_selected_projects = st.multiselect("Chọn Các Dự Án (ít nhất 2)", comp_project_options, 
                                                    default=st.session_state['comparison_config']['selected_projects'],
                                                    key='comp_projects_year_project')
            if len(comp_selected_projects) < 2:
                st.warning("Vui lòng chọn ít nhất HAI dự án cho chế độ này.")
            st.session_state['comparison_config']['selected_projects'] = comp_selected_projects
        
        elif current_comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
            st.info("Chọn CHỈ MỘT dự án. Sau đó, chọn một năm (để so sánh các tháng trong năm đó) HOẶC nhiều năm (để so sánh tổng giờ qua các năm).")
            
            # Projects selection - must be single
            comp_project_options = get_unique_values(df_raw, 'Project name')
            comp_selected_projects = st.multiselect("Chọn Một Dự Án (chỉ 1)", comp_project_options, 
                                                    default=st.session_state['comparison_config']['selected_projects'],
                                                    key='comp_project_single_project_time')
            if len(comp_selected_projects) > 1:
                st.warning("Vui lòng chọn CHỈ MỘT dự án cho chế độ này.")
                comp_selected_projects = comp_selected_projects[:1]
            st.session_state['comparison_config']['selected_projects'] = comp_selected_projects

            # Year/Month selection - dynamic based on user intent
            st.markdown("Chọn 'Năm' để xem so sánh các tháng trong năm đó, HOẶC chọn nhiều 'Năm' để xem tổng giờ qua các năm.")

            comp_year_options = get_unique_values(df_raw, 'Year')
            comp_selected_years = st.multiselect("Chọn Năm(s)", comp_year_options, 
                                                 default=st.session_state['comparison_config']['years'],
                                                 key='comp_years_single_project_time')
            st.session_state['comparison_config']['years'] = comp_selected_years

            if len(comp_selected_years) == 1:
                # If one year is selected, allow selecting months for that year
                df_for_comp_month = df_raw[df_raw['Year'].isin(comp_selected_years)]
                comp_month_options = [m for m in MONTH_NAMES_EN if m in df_for_comp_month['MonthName'].unique()]
                comp_selected_months = st.multiselect("Chọn Tháng(s) (trong năm đã chọn)", comp_month_options, 
                                                      default=st.session_state['comparison_config']['months'],
                                                      key='comp_months_single_project_time')
                st.session_state['comparison_config']['months'] = comp_selected_months
            else:
                # If multiple years selected, clear months
                st.session_state['comparison_config']['months'] = []

            # Display warning if configuration is invalid
            if len(comp_selected_projects) != 1:
                 st.warning("Vui lòng chọn CHỈ MỘT dự án.")
            elif not comp_selected_years:
                st.warning("Vui lòng chọn ít nhất một năm.")
            elif len(comp_selected_years) == 1 and not st.session_state['comparison_config']['months']:
                st.warning("Vui lòng chọn ít nhất một tháng nếu bạn chỉ chọn một năm để so sánh theo tháng.")
            elif len(comp_selected_years) > 1 and st.session_state['comparison_config']['months']:
                st.warning("Không thể so sánh nhiều năm VÀ các tháng cùng lúc. Vui lòng xóa lựa chọn tháng nếu bạn muốn so sánh nhiều năm.")


        # Button to generate comparison
        if st.button("Tạo báo cáo so sánh", key='generate_comparison_button'):
            with st.spinner("Đang tạo báo cáo so sánh..."):
                df_comparison, msg = apply_comparison_filters(
                    df_raw, 
                    st.session_state['comparison_config'], 
                    current_comparison_mode
                )
                
                if not df_comparison.empty:
                    chart_title = msg # The message from apply_comparison_filters is the chart title
                    comparison_chart_fig = create_comparison_chart(df_comparison, current_comparison_mode, chart_title)
                    st.plotly_chart(comparison_chart_fig, use_container_width=True)

                    comparison_table_fig = create_comparison_table(df_comparison.copy(), current_comparison_mode)
                    st.plotly_chart(comparison_table_fig, use_container_width=True)

                    pdf_comparison_report = export_pdf_report(
                        create_overall_summary_table(df_filtered), # Still include overall summary from current filter
                        go.Figure(), # Empty figure for monthly summary
                        go.Figure(), # Empty figure for project summary
                        go.Figure(), # Empty figure for raw data
                        comparison_chart_fig, 
                        comparison_table_fig,
                        st.session_state['comparison_config']['years'], # Pass years for naming
                        st.session_state['comparison_config']['months'], # Pass months for naming
                        st.session_state['comparison_config']['selected_projects'], # Pass projects for naming
                        selected_comparison_mode_text, # Use the user-friendly text for PDF
                        "comparison_report"
                    )

                    comp_filename_parts = ["BaoCaoSoSanh"]
                    if st.session_state['comparison_config']['selected_projects']:
                        if len(st.session_state['comparison_config']['selected_projects']) == 1:
                            comp_filename_parts.append(st.session_state['comparison_config']['selected_projects'][0].replace(" ", "_"))
                        else:
                            comp_filename_parts.append("NhieuDuAn")
                    
                    if st.session_state['comparison_config']['years']:
                        if len(st.session_state['comparison_config']['years']) == 1:
                            comp_filename_parts.append(str(st.session_state['comparison_config']['years'][0]))
                        else:
                            comp_filename_parts.append("NhieuNam")
                    
                    if st.session_state['comparison_config']['months']:
                        if len(st.session_state['comparison_config']['months']) == 1:
                            comp_filename_parts.append(st.session_state['comparison_config']['months'][0].replace(" ", "_"))
                        else:
                            comp_filename_parts.append("NhieuThang")

                    comp_filename = "_".join(comp_filename_parts) + f"_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

                    st.download_button(
                        label="Tải xuống báo cáo so sánh PDF",
                        data=pdf_comparison_report,
                        file_name=comp_filename,
                        mime="application/pdf"
                    )
                else:
                    st.warning(msg) # Display the error message from apply_comparison_filters
else:
    st.info("Vui lòng tải lên file Excel để bắt đầu.")
