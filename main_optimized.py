import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ==============================================================================
# ĐẢM BẢO FILE 'a04ecaf1_1dae_4c90_8081_086cd7c7b725.py' NẰM CÙNG THƯ MỤC
# HOẶC THAY THẾ TÊN FILE NẾU BẠN ĐÃ ĐỔI TÊN NÓ.
# ==============================================================================
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    apply_filters, export_report, export_pdf_report,
    apply_comparison_filters, export_comparison_report, export_comparison_pdf_report,
    initialize_language_data, get_text # Add initialize_language_data and get_text
)
# ==============================================================================

script_dir = os.path.dirname(__file__)
csv_file_path = os.path.join(script_dir, "invited_emails.csv")

# Gọi hàm setup_paths ngay từ đầu để path_dict có sẵn
path_dict = setup_paths()

# ---------------------------\
# PHẦN XÁC THỰC TRUY CẬP
# ---------------------------\

@st.cache_data
def load_invited_emails():
    try:
        df = pd.read_csv(csv_file_path, header=None, encoding='utf-8')
        emails = df.iloc[:, 0].astype(str).str.strip().str.lower().tolist()
        return emails
    except FileNotFoundError:
        st.error("File invited_emails.csv không tìm thấy. Vui lòng đảm bảo file này có trong cùng thư mục với script.")
        return []

# --- Language Setup ---
# Initialize language data
lang_data = initialize_language_data()

# Language selection in sidebar
language_options = {
    "Tiếng Việt": "vi",
    "English": "en"
}
# Set default to English (index=1)
selected_language_name = st.sidebar.selectbox("Select Language / Chọn ngôn ngữ", list(language_options.keys()), index=1) 
st.session_state['language'] = language_options[selected_language_name]

# Now, use get_text throughout the app by passing the language data
# This partial function ensures we don't need to pass lang_data every time
get_text_func = lambda key: get_text(key, st.session_state['language'], lang_data)

st.set_page_config(layout="wide", page_title=get_text_func("app_title"))
st.title(get_text_func("app_title"))

# Authentication logic (assuming it's present but not shown in snippet)
# You need to uncomment and use your actual authentication logic here
# invited_emails = load_invited_emails()
# if 'authenticated' not in st.session_state:
#     st.session_state['authenticated'] = False
# if not st.session_state['authenticated']:
#     email = st.text_input("Enter your email to access:", key="auth_email")
#     if st.button("Access"):
#         if email.strip().lower() in invited_emails:
#             st.session_state['authenticated'] = True
#             st.success("Authentication successful!")
#             st.rerun()
#         else:
#             st.error("Access denied. Please check your email or contact support.")
#     st.stop() # Stop execution if not authenticated
authenticated = True # Temporarily set to True for development/testing

if authenticated:
    # --- File Upload ---
    uploaded_file = st.sidebar.file_uploader(get_text_func("upload_excel_file"), type=["xlsx"])

    df_raw = pd.DataFrame()
    if uploaded_file:
        with st.spinner(get_text_func("loading_data_spinner")):
            df_raw = load_raw_data(uploaded_file, path_dict)
        if not df_raw.empty:
            st.sidebar.success(get_text_func("file_upload_success"))
            st.session_state['df_raw'] = df_raw
        else:
            st.sidebar.error(get_text_func("file_upload_error"))
            st.session_state['df_raw'] = None
    elif 'df_raw' in st.session_state and st.session_state['df_raw'] is not None:
        df_raw = st.session_state['df_raw']

    if df_raw.empty:
        st.info(get_text_func("upload_file_to_start"))
        st.stop() # Stop execution if no data

    # --- Sidebar Filters ---
    st.sidebar.header(get_text_func("filters_header"))

    unique_years = sorted(df_raw['Year'].unique().tolist())
    all_years_option = get_text_func("all_years_option")
    year_options = [all_years_option] + unique_years
    selected_year_filter = st.sidebar.selectbox(get_text_func("select_year"), year_options)

    selected_year = None
    if selected_year_filter != all_years_option:
        selected_year = selected_year_filter

    # Get unique months based on selected year for month filter
    df_for_month_filter = df_raw
    if selected_year is not None:
        df_for_month_filter = df_raw[df_raw['Year'] == selected_year]
    unique_months = sorted(df_for_month_filter['MonthName'].unique().tolist())
    
    # Ensure month names are consistently ordered (e.g., Jan, Feb, Mar...)
    month_name_order = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    # Filter and sort unique_months according to month_name_order
    ordered_unique_months = [m for m in month_name_order if m in unique_months]


    all_months_option = get_text_func("all_months_option")
    month_options = [all_months_option] + ordered_unique_months
    selected_month_filter = st.sidebar.selectbox(get_text_func("select_month"), month_options)
    
    selected_month_name = None
    if selected_month_filter != all_months_option:
        selected_month_name = selected_month_filter

    unique_projects = sorted(df_raw['Project name'].unique().tolist())
    all_projects_option = get_text_func("all_projects_option")
    project_options = [all_projects_option] + unique_projects
    selected_project_filter = st.sidebar.selectbox(get_text_func("select_project"), project_options)

    selected_project_name = None
    if selected_project_filter != all_projects_option:
        selected_project_name = selected_project_filter

    df_filtered = apply_filters(df_raw, selected_year, selected_month_name, selected_project_name)

    # --- Main Content Tabs ---
    tab_overview_main, tab_comparison_main, tab_data_preview_main, tab_user_guide_main = st.tabs([
        get_text_func("overview_report_tab"),
        get_text_func("comparison_tab"),
        get_text_func("data_preview_tab"),
        get_text_func("user_guide_tab")
    ])

    with tab_overview_main:
        st.header(get_text_func("overview_report_header"))

        if df_filtered.empty:
            st.warning(get_text_func("no_data_for_filters"))
        else:
            # Display summary tables and charts
            st.subheader(get_text_func("overall_summary"))
            overall_summary_df = pd.DataFrame({
                get_text_func("metric_column"): [get_text_func("total_hours"), get_text_func("total_cost_usd")],
                get_text_func("value_column"): [f"{df_filtered['Hours'].sum():,.0f}", f"{df_filtered['Total cost (USD)'].sum():,.2f}"]
            })
            st.table(overall_summary_df)

            st.subheader(get_text_func("monthly_summary_header"))
            monthly_summary_df = df_filtered.groupby('MonthName').agg(
                total_hours=('Hours', 'sum'),
                total_cost=('Total cost (USD)', 'sum')
            ).reset_index()
            # Order months for display
            monthly_summary_df['MonthName'] = pd.Categorical(monthly_summary_df['MonthName'], categories=ordered_unique_months, ordered=True)
            monthly_summary_df = monthly_summary_df.sort_values('MonthName')
            st.dataframe(monthly_summary_df.style.format({"total_hours": "{:,.0f}", "total_cost": "{:,.2f}"}), use_container_width=True)

            st.subheader(get_text_func("project_summary_header"))
            project_summary_df = df_filtered.groupby('Project name').agg(
                total_hours=('Hours', 'sum'),
                total_cost=('Total cost (USD)', 'sum')
            ).reset_index().sort_values(by='total_hours', ascending=False)
            st.dataframe(project_summary_df.style.format({"total_hours": "{:,.0f}", "total_cost": "{:,.2f}"}), use_container_width=True)

            st.markdown("---")
            st.subheader(get_text_func("export_report_header"))

            export_excel = st.checkbox(get_text_func("export_excel"), value=True)
            export_pdf = st.checkbox(get_text_func("export_pdf"), value=True)

            if st.button(get_text_func("create_report_button")):
                if export_excel or export_pdf:
                    with st.spinner(get_text_func("generating_report_spinner")):
                        report_created = export_report(
                            df_filtered,
                            path_dict['template_file'],
                            path_dict['output_file'],
                            selected_year,
                            selected_month_name,
                            selected_project_name,
                            get_text_func # Pass get_text_func to the export function
                        )

                        pdf_created = False
                        if export_pdf:
                            pdf_created = export_pdf_report(
                                df_filtered,
                                path_dict['pdf_report'],
                                selected_year,
                                selected_month_name,
                                selected_project_name,
                                path_dict['logo_path'],
                                get_text_func # Pass get_text_func to the PDF export function
                            )
                        
                    if report_created or pdf_created: # Check if either was created successfully
                        if export_excel and os.path.exists(path_dict['output_file']):
                            with open(path_dict['output_file'], "rb") as f:
                                st.download_button(get_text_func("download_excel_report"), data=f, file_name=os.path.basename(path_dict['output_file']), use_container_width=True, key='download_excel_btn')
                        if export_pdf and pdf_created and os.path.exists(path_dict['pdf_report']):
                            with open(path_dict['pdf_report'], "rb") as f:
                                st.download_button(get_text_func("download_pdf_report"), data=f, file_name=os.path.basename(path_dict['pdf_report']), use_container_width=True, key='download_pdf_btn')
                        elif export_pdf and not pdf_created:
                             st.error(get_text_func('error_generating_report') + " " + get_text_func('pdf_specific_error')) # Specific PDF error
                    else:
                        st.error(get_text_func('error_generating_report'))
                else:
                    st.warning(get_text_func('select_export_format'))

    with tab_comparison_main:
        st.header(get_text_func("comparison_header"))

        comparison_mode_options = {
            get_text_func("compare_projects_in_month"): "month_project",
            get_text_func("compare_projects_in_year"): "year_project",
            get_text_func("compare_one_project_over_time"): "project_over_time"
        }
        
        selected_comparison_mode_text = st.selectbox(
            get_text_func("select_comparison_mode"),
            list(comparison_mode_options.keys()),
            key='comparison_mode_select'
        )
        current_comparison_mode = comparison_mode_options[selected_comparison_mode_text]

        st.markdown("---")
        st.subheader(get_text_func("comparison_config_header"))

        comparison_config = {'years': [], 'months': [], 'selected_projects': []}

        # Dynamically populate month options for comparison based on selected year(s)
        # For simplicity, getting all unique months and letting the backend filter
        # The specific ordering for selectbox is handled by the `ordered_unique_months` which uses MONTH_NAME_ORDER
        
        comp_year_options = sorted(df_raw['Year'].unique().tolist())
        comp_month_all_options = sorted(df_raw['MonthName'].unique().tolist())
        ordered_comp_month_all_options = [m for m in month_name_order if m in comp_month_all_options]
        comp_project_options = sorted(df_raw['Project name'].unique().tolist())


        if current_comparison_mode == "month_project":
            st.info(get_text_func("comp_month_project_info"))
            
            comp_selected_years = st.multiselect(get_text_func("select_year_single"), comp_year_options, key='comp_years_month_project')
            if len(comp_selected_years) > 1:
                st.warning(get_text_func("select_only_one_year"))
                comp_selected_years = comp_selected_years[:1]
            comparison_config['years'] = comp_selected_years
            
            comp_selected_months = st.multiselect(get_text_func("select_month_single"), ordered_comp_month_all_options, key='comp_months_month_project')
            if len(comp_selected_months) > 1:
                st.warning(get_text_func("select_only_one_month"))
                comp_selected_months = comp_selected_months[:1]
            comparison_config['months'] = comp_selected_months

            comp_selected_projects = st.multiselect(get_text_func("select_projects_multiple"), comp_project_options, key='comp_projects_month_project')
            if len(comp_selected_projects) < 2:
                st.warning(get_text_func("select_at_least_two_projects"))
            comparison_config['selected_projects'] = comp_selected_projects

        elif current_comparison_mode == "year_project":
            st.info(get_text_func("comp_year_project_info"))
            
            comp_selected_years = st.multiselect(get_text_func("select_year_single"), comp_year_options, key='comp_years_year_project')
            if len(comp_selected_years) > 1:
                st.warning(get_text_func("select_only_one_year"))
                comp_selected_years = comp_selected_years[:1]
            comparison_config['years'] = comp_selected_years
            comparison_config['months'] = [] # Clear months for this mode

            comp_selected_projects = st.multiselect(get_text_func("select_projects_multiple"), comp_project_options, key='comp_projects_year_project')
            if len(comp_selected_projects) < 2:
                st.warning(get_text_func("select_at_least_two_projects"))
            comparison_config['selected_projects'] = comp_selected_projects
        
        elif current_comparison_mode == "project_over_time":
            st.info(get_text_func("comp_project_over_time_info"))
            
            comp_selected_projects = st.multiselect(get_text_func("select_project_single"), comp_project_options, key='comp_project_single_project_time')
            if len(comp_selected_projects) > 1:
                st.warning(get_text_func("select_only_one_project"))
                comp_selected_projects = comp_selected_projects[:1]
            comparison_config['selected_projects'] = comp_selected_projects

            st.markdown(get_text_func("select_years_or_months_info"))

            comp_selected_years = st.multiselect(get_text_func("select_years"), comp_year_options, key='comp_years_single_project_time')
            comparison_config['years'] = comp_selected_years

            if len(comp_selected_years) == 1:
                # If one year is selected, allow selecting months for that year
                df_for_comp_month = df_raw[df_raw['Year'].isin(comp_selected_years)]
                comp_month_options_for_year = sorted(df_for_comp_month['MonthName'].unique().tolist())
                ordered_comp_month_options_for_year = [m for m in month_name_order if m in comp_month_options_for_year]

                comp_selected_months = st.multiselect(get_text_func("select_months_in_year"), ordered_comp_month_options_for_year, key='comp_months_single_project_time')
                comparison_config['months'] = comp_selected_months
            else:
                # If multiple years selected, clear months
                comparison_config['months'] = []

            # Display warning if configuration is invalid
            if len(comp_selected_projects) != 1:
                 st.warning(get_text_func("select_only_one_project_again"))
            # Refined check for time selection: must have at least one year OR (one year AND at least one month)
            elif not comp_selected_years and not comparison_config['months']:
                st.warning(get_text_func("select_at_least_one_year_or_month"))
            elif len(comp_selected_years) == 1 and not comparison_config['months']:
                st.warning(get_text_func("select_at_least_one_month_if_one_year"))
            elif len(comp_selected_years) > 1 and comparison_config['months']:
                st.warning(get_text_func("cannot_compare_multiple_years_and_months"))


        export_excel_comp = st.checkbox(get_text_func("export_excel"), value=True, key='export_excel_comp')
        export_pdf_comp = st.checkbox(get_text_func("export_pdf"), value=True, key='export_pdf_comp')

        if st.button(get_text_func("create_comparison_report"), key='generate_comparison_button'):
            # Validate selected options before calling backend
            is_valid_config = True
            if current_comparison_mode == "month_project":
                if len(comparison_config['years']) != 1 or len(comparison_config['months']) != 1 or len(comparison_config['selected_projects']) < 2:
                    is_valid_config = False
            elif current_comparison_mode == "year_project":
                if len(comparison_config['years']) != 1 or len(comparison_config['selected_projects']) < 2:
                    is_valid_config = False
            elif current_comparison_mode == "project_over_time":
                if len(comparison_config['selected_projects']) != 1:
                    is_valid_config = False
                elif not comparison_config['years'] and not comparison_config['months']:
                    is_valid_config = False
                elif len(comparison_config['years']) == 1 and not comparison_config['months']:
                    is_valid_config = False
                elif len(comparison_config['years']) > 1 and comparison_config['months']:
                    is_valid_config = False
            
            if not is_valid_config:
                st.error(get_text_func("invalid_config_please_check_messages"))
            elif not (export_excel_comp or export_pdf_comp):
                st.warning(get_text_func("select_export_format"))
            else:
                with st.spinner(get_text_func("generating_comparison_report_spinner")):
                    # df_comparison is the main DataFrame for table display
                    # chart_df is prepared for chart drawing
                    # table_df is specifically for table formatting in PDF (might be same as df_comparison)
                    df_comparison, msg, chart_df_for_plot, table_df_for_pdf = apply_comparison_filters(df_raw, comparison_config, current_comparison_mode, get_text_func)
                    
                    if not df_comparison.empty:
                        st.subheader(get_text_func("comparison_chart_header"))
                        # The `create_chart_image` is used by the PDF export.
                        # For Streamlit display, you'd typically use Plotly Express here.
                        # Since your original code didn't have live chart rendering here,
                        # I'll keep the placeholder text. If you want live charts,
                        # we'd need to add Plotly code here.
                        st.info(get_text_func("chart_placeholder")) # Display placeholder
                        
                        st.subheader(get_text_func("comparison_table_header"))
                        # Adjust format based on the structure of df_comparison
                        if current_comparison_mode == "month_project":
                            st.dataframe(df_comparison.style.format({get_text_func("total_hours_col"): "{:,.0f}"}), use_container_width=True)
                        elif current_comparison_mode == "year_project":
                            # Dynamic formatting for months and Total Hours
                            format_dict = {col: "{:,.0f}" for col in df_comparison.columns if col in ordered_unique_months or col == get_text_func("total_hours_col")}
                            st.dataframe(df_comparison.style.format(format_dict), use_container_width=True)
                        elif current_comparison_mode == "project_over_time":
                             # The hours column name can vary based on project name
                            hours_col_name = df_comparison.columns[1] if len(df_comparison.columns) > 1 else None
                            format_dict = {}
                            if hours_col_name:
                                format_dict[hours_col_name] = "{:,.0f}"
                            st.dataframe(df_comparison.style.format(format_dict), use_container_width=True)

                        comparison_report_created = export_comparison_report(
                            df_comparison, # This is the data used for Excel table
                            current_comparison_mode,
                            path_dict['comparison_output_file'],
                            comparison_config,
                            get_text_func # Pass get_text_func
                        )

                        comparison_pdf_created = False
                        if export_pdf_comp:
                            comparison_pdf_created = export_comparison_pdf_report(
                                chart_df_for_plot, # Data for chart generation (matplotlib)
                                table_df_for_pdf, # Data for table in PDF
                                current_comparison_mode,
                                path_dict['comparison_pdf_report'],
                                path_dict['logo_path'],
                                comparison_config,
                                get_text_func # Pass get_text_func
                            )

                        if export_excel_comp and os.path.exists(path_dict['comparison_output_file']):
                            with open(path_dict['comparison_output_file'], "rb") as f:
                                st.download_button(get_text_func("download_comparison_excel"), data=f, file_name=os.path.basename(path_dict['comparison_output_file']), use_container_width=True, key='download_excel_comp_btn')
                        if export_pdf_comp and comparison_pdf_created and os.path.exists(path_dict['comparison_pdf_report']):
                            with open(path_dict['comparison_pdf_report'], "rb") as f:
                                st.download_button(get_text_func("download_comparison_pdf"), data=f, file_name=os.path.basename(path_dict['comparison_pdf_report']), use_container_width=True, key='download_pdf_comp_btn')
                        elif export_pdf_comp and not comparison_pdf_created:
                            st.error(get_text_func('error_generating_report') + " " + get_text_func('pdf_specific_error_comp')) # Specific PDF error
                    else:
                        st.error(msg) # Display error message from comparison filter

# =========================================================================
# DATA PREVIEW TAB
# =========================================================================
with tab_data_preview_main:
    st.subheader(get_text_func('raw_data_preview_header'))
    if not df_raw.empty:
        st.dataframe(df_raw.head(100))
    else:
        st.info(get_text_func('no_raw_data'))

# =========================================================================
# USER GUIDE TAB
# =========================================================================
with tab_user_guide_main:
    st.markdown(f"### {get_text_func('user_guide')}")
    st.markdown(get_text_func("user_guide_content")) # Use get_text_func for content
