import streamlit as st
import pandas as pd
import os
from datetime import datetime
# Import all functions from the specified module
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import setup_paths, load_raw_data, read_configs, apply_filters, export_report, apply_comparison_filters, export_comparison_report, export_comparison_pdf_report

st.set_page_config(page_title="Time Report Generator", layout="wide") # Changed to wide layout
st.title("📊 Time Report Generator")

# Setup paths (e.g., template, output files)
path_dict = setup_paths()

# Check if template file exists
if not os.path.exists(path_dict['template_file']):
    st.error(f"❌ Template file not found: {path_dict['template_file']}")
    st.stop()

# Load raw data and configurations once
with st.spinner("🔄 Loading data and configurations..."):
    df_raw = load_raw_data(path_dict)
    if df_raw.empty:
        st.error("⚠️ Failed to load raw data. Please check 'Raw Data' sheet in the template file.")
        st.stop()

    config_data = read_configs(path_dict)

# Get unique years, months, and projects from raw data for selectbox options
all_years = sorted(df_raw['Year'].dropna().unique().astype(int).tolist())
all_months = list(df_raw['MonthName'].dropna().unique())
all_projects = sorted(df_raw['Project name'].dropna().unique().tolist())

# Main interface tabs
tab_standard, tab_comparison, tab_data_preview = st.tabs(["Standard Report", "Comparison Report", "Data Preview"])

# =========================================================================
# STANDARD REPORT TAB
# =========================================================================
with tab_standard:
    st.header("Standard Time Report Configuration")

    col1, col2, col3 = st.columns(3)
    with col1:
        mode = st.selectbox(
            "Select analysis mode:", 
            options=['year', 'month', 'week'], 
            index=['year', 'month', 'week'].index(config_data['mode']) if config_data['mode'] in ['year', 'month', 'week'] else 0,
            key='standard_mode'
        )
    with col2:
        year_options = [y for y in all_years if y is not None] # Filter out None/NaN
        selected_year = st.selectbox(
            "Select year:", 
            options=year_options, 
            index=year_options.index(config_data['year']) if config_data['year'] in year_options else (0 if year_options else 0), # Ensure valid index
            key='standard_year'
        )
    with col3:
        # Default months to all if config_data['months'] is empty, otherwise use config
        default_months_standard = config_data['months'] if config_data['months'] else all_months
        selected_months = st.multiselect(
            "Select month(s):", 
            options=all_months, 
            default=default_months_standard,
            key='standard_months'
        )

    # Project selection for standard report
    st.subheader("Project Selection for Standard Report")
    # Initialize included_projects based on config_data's project_filter_df
    # Only include projects marked 'yes' in the config
    initial_included_projects_config = config_data['project_filter_df'][
        config_data['project_filter_df']['Include'].astype(str).str.lower() == 'yes'
    ]['Project Name'].tolist()
    
    # Ensure all initial projects from config are actually present in all_projects
    initial_included_projects_valid = [p for p in initial_included_projects_config if p in all_projects]
    
    # Use the valid initial projects as default, or all_projects if none are configured/valid
    default_standard_projects = initial_included_projects_valid if initial_included_projects_valid else all_projects

    standard_project_selection = st.multiselect(
        "Select projects to include (only 'yes' projects from template config will be included by default):", 
        options=all_projects, 
        default=default_standard_projects,
        key='standard_project_selection'
    )

    if st.button("🚀 Generate Standard Report", key='generate_standard_report_btn'):
        # Create a temporary project_filter_df for the standard report based on user selection
        # This simulates the 'Include' column logic for selected projects
        temp_project_filter_df_standard = pd.DataFrame({
            'Project Name': all_projects,
            'Include': ['yes' if p in standard_project_selection else 'no' for p in all_projects]
        })
        
        standard_report_config = {
            'mode': mode,
            'year': selected_year,
            'months': selected_months,
            'project_filter_df': temp_project_filter_df_standard[temp_project_filter_df_standard['Project Name'].isin(standard_project_selection)]
        }

        df_filtered_standard = apply_filters(df_raw, standard_report_config)

        if df_filtered_standard.empty:
            st.warning("⚠️ No data after filtering for the standard report. Please check your selections.")
        else:
            with st.spinner("Generating Excel report..."):
                excel_success = export_report(df_filtered_standard, standard_report_config, path_dict)
            
            if excel_success:
                st.success(f"✅ Excel Report generated: {os.path.basename(path_dict['output_file'])}")
                with open(path_dict['output_file'], "rb") as f:
                    st.download_button("📥 Download Excel Report", data=f, file_name=os.path.basename(path_dict['output_file']), key='download_excel_standard')
                
                with st.spinner("Generating PDF report..."):
                    pdf_success = export_pdf_report(df_filtered_standard, standard_report_config, path_dict)
                
                if pdf_success:
                    st.success(f"✅ PDF Report generated: {os.path.basename(path_dict['pdf_report'])}")
                    with open(path_dict['pdf_report'], "rb") as f:
                        st.download_button("📥 Download PDF Report", data=f, file_name=os.path.basename(path_dict['pdf_report']), key='download_pdf_standard')
                else:
                    st.error("❌ Failed to generate PDF report.")
            else:
                st.error("❌ Failed to generate Excel report.")


# =========================================================================
# COMPARISON REPORT TAB
# =========================================================================
with tab_comparison:
    st.header("Comparison Report Configuration")

    comparison_mode = st.selectbox(
        "Select comparison mode:",
        options=[
            "So Sánh Dự Án Trong Một Tháng", # Compare Projects in a Month
            "So Sánh Dự Án Trong Một Năm",   # Compare Projects in a Year (by Month)
            "So Sánh Một Dự Án Qua Các Tháng/Năm" # Compare One Project Over Time (Months/Years)
        ],
        key='comparison_mode_select'
    )

    # Input for Years, Months, Projects (for comparison)
    st.subheader("Filter Data for Comparison")
    
    col_comp1, col_comp2 = st.columns(2)
    with col_comp1:
        comp_years = st.multiselect("Select Year(s):", options=all_years, default=[all_years[0]] if all_years else [], key='comp_years')
    with col_comp2:
        comp_months = st.multiselect("Select Month(s):", options=all_months, default=[], key='comp_months')
    
    comp_projects = st.multiselect("Select Project(s):", options=all_projects, default=[], key='comp_projects')

    if st.button("🚀 Generate Comparison Report", key='generate_comparison_report_btn'):
        comparison_config = {
            'years': comp_years,
            'months': comp_months,
            'selected_projects': comp_projects,
            # For comparison report, project_filter_df is not used in the same way as standard
            # We directly pass 'selected_projects' to apply_comparison_filters
        }

        df_comparison, message = apply_comparison_filters(df_raw, comparison_config, comparison_mode)

        if df_comparison.empty:
            st.warning(f"⚠️ {message}")
        else:
            st.success("✅ Data filtered successfully for comparison.")
            st.subheader("Comparison Data Preview")
            st.dataframe(df_comparison)

            with st.spinner("Generating Comparison Excel Report..."):
                excel_success_comp = export_comparison_report(df_comparison, comparison_config, path_dict, comparison_mode)
            
            if excel_success_comp:
                st.success(f"✅ Comparison Excel Report generated: {os.path.basename(path_dict['comparison_output_file'])}")
                with open(path_dict['comparison_output_file'], "rb") as f:
                    st.download_button("📥 Download Comparison Excel", data=f, file_name=os.path.basename(path_dict['comparison_output_file']), key='download_excel_comparison')
                
                with st.spinner("Generating Comparison PDF Report..."):
                    pdf_success_comp = export_comparison_pdf_report(df_comparison, comparison_config, path_dict, comparison_mode)
                
                if pdf_success_comp:
                    st.success(f"✅ Comparison PDF Report generated: {os.path.basename(path_dict['comparison_pdf_report'])}")
                    with open(path_dict['comparison_pdf_report'], "rb") as f:
                        st.download_button("📥 Download Comparison PDF", data=f, file_name=os.path.basename(path_dict['comparison_pdf_report']), key='download_pdf_comparison')
                else:
                    st.error("❌ Failed to generate Comparison PDF report.")
            else:
                st.error("❌ Failed to generate Comparison Excel report.")

# =========================================================================
# DATA PREVIEW TAB
# =========================================================================
with tab_data_preview:
    st.subheader("Raw Input Data (First 100 rows)")
    if not df_raw.empty:
        st.dataframe(df_raw.head(100))
    else:
        st.info("No raw data loaded.")
