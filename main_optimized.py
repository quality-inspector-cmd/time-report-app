import streamlit as st 
import pandas as pd
import os
from datetime import datetime
from a04ecaf1_1dae_4c90_8081_086cd7c7b725 import (
    setup_paths, load_raw_data, read_configs,
    apply_filters, export_report
)

st.set_page_config(page_title="Time Report Generator (v2)", layout="centered")
st.title("📊 Time Report Generator (v2.1)")

path_dict = setup_paths()

if not os.path.exists(path_dict['template_file']):
    st.error(f"❌ Template file not found: {path_dict['template_file']}")
    st.stop()

@st.cache_data
def cached_load_raw_data(path_dict):
    return load_raw_data(path_dict)

@st.cache_data
def cached_read_configs(path_dict):
    return read_configs(path_dict)

with st.spinner("🔄 Loading data..."):
    df_raw = cached_load_raw_data(path_dict)
    config_data = cached_read_configs(path_dict)

tab1, tab2 = st.tabs(["Report configuration", "Data preview"])

with tab1:
    mode = st.selectbox("Select analysis mode:", options=['year', 'month', 'week'],
                        index=['year', 'month', 'week'].index(config_data['mode']))

    all_years = sorted(df_raw['Year'].dropna().unique())
    default_year = config_data['year']
    years = st.multiselect("Select year(s):", options=all_years,
                           default=[default_year] if default_year else all_years)

    all_months = list(df_raw['MonthName'].dropna().unique())
    months = st.multiselect("Select month:", options=all_months,
                            default=config_data['months'] if config_data['months'] else all_months)

    project_df = config_data['project_filter_df']
    included_projects = project_df[project_df['Include'].str.lower() == 'yes']['Project Name'].tolist()
    project_selection = st.multiselect("Select project:",
                                       options=sorted(project_df['Project Name'].unique()),
                                       default=included_projects)

    if st.button("🚀 Generate report"):
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
                st.warning("⚠️ No data after filtering. Please check your selection again.")
            else:
                export_report(df_filtered, config, path_dict)
                st.success(f"✅ Report created: {os.path.basename(path_dict['output_file'])}")
                with open(path_dict['output_file'], "rb") as f:
                    st.download_button("📥 Download Excel report", data=f,
                                       file_name=os.path.basename(path_dict['output_file']))

with tab2:
    st.subheader("📂 Input data (first 100 rows)")
    with st.expander("Click to view raw data sample"):
        st.dataframe(df_raw.head(100))
