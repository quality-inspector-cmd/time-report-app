import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import datetime
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF

sns.set(style="whitegrid")

def setup_paths():
    today = datetime.datetime.today().strftime('%Y%m%d')
    return {
        'template_file': "Time_report.xlsm",
        'output_file': f"Time_report_{today}.xlsx",
        'pdf_report': f"Time_report_{today}.pdf"
    }

def read_configs(path_dict):
    year_mode_df = pd.read_excel(path_dict['template_file'], sheet_name='Config_Year_Mode', engine='openpyxl')
    project_filter_df = pd.read_excel(path_dict['template_file'], sheet_name='Config_Project_Filter', engine='openpyxl')

    mode = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'mode', 'Value'].values[0].strip().lower()
    year_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'year', 'Value']
    year = int(year_row.values[0]) if not year_row.empty and pd.notna(year_row.values[0]) else None

    months_row = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'months', 'Value']
    months = [m.strip().capitalize() for m in str(months_row.values[0]).split(',')] if not months_row.empty else []

    return {
        'mode': mode,
        'year': year,
        'months': months,
        'project_filter_df': project_filter_df
    }

def load_raw_data(path_dict):
    df = pd.read_excel(path_dict['template_file'], sheet_name='Raw Data', engine='openpyxl')
    df.rename(columns={'Hou': 'Hours', 'Team member': 'Employee'}, inplace=True)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df['Year'] = df['Date'].dt.year
    df['MonthName'] = df['Date'].dt.month_name()
    df['Week'] = df['Date'].dt.isocalendar().week
    return df

def apply_filters(df, config):
    df_filtered = df.copy()
    if 'years' in config and config['years']:
        df_filtered = df_filtered[df_filtered['Year'].isin(config['years'])]
    elif 'year' in config and config['year']:
        df_filtered = df_filtered[df_filtered['Year'] == config['year']]
    if config['months']:
        df_filtered = df_filtered[df_filtered['MonthName'].isin(config['months'])]

    df_filtered = df_filtered.merge(
        config['project_filter_df'][config['project_filter_df']['Include'].str.lower() == 'yes'],
        how='inner',
        left_on='Project name',
        right_on='Project Name'
    )
    return df_filtered

def export_report(df, config, path_dict):
    mode = config['mode']
    if mode == 'year':
        summary = df.groupby(['Year', 'Project name'])['Hours'].sum().reset_index()
    elif mode == 'month':
        summary = df.groupby(['Year', 'MonthName', 'Project name'])['Hours'].sum().reset_index()
    else:
        summary = df.groupby(['Year', 'Week', 'Project name'])['Hours'].sum().reset_index()

    with pd.ExcelWriter(path_dict['output_file'], engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='Summary', index=False)

    wb = load_workbook(path_dict['output_file'])
    ws_summary = wb['Summary']
    max_row = ws_summary.max_row

    data_ref = Reference(ws_summary, min_col=3 if mode == 'year' else 4, min_row=1, max_row=max_row)
    cats_ref = Reference(ws_summary, min_col=2 if mode == 'year' else 3, min_row=2, max_row=max_row)

    chart = BarChart()
    chart.title = f"Total Hours by Project ({mode})"
    chart.x_axis.title = "Project"
    chart.y_axis.title = "Hours"
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws_summary.add_chart(chart, "F2")

    for project in df['Project name'].unique():
        df_proj = df[df['Project name'] == project]
        ws_proj = wb.create_sheet(title=project[:31])

        # --- Workcentre Summary ---
        summary_wc = df_proj.groupby('Workcentre')['Hours'].sum().reset_index()
        ws_proj.append(["Workcentre", "Hours"])
        for r in dataframe_to_rows(summary_wc, index=False, header=False):
            ws_proj.append(r)

        chart_wc = BarChart()
        chart_wc.title = f"{project} - Hours by Workcentre"
        chart_wc.x_axis.title = "Workcentre"
        chart_wc.y_axis.title = "Hours"
        wc_rows = len(summary_wc)
        data_ref = Reference(ws_proj, min_col=2, min_row=2, max_row=wc_rows + 1)
        cats_ref = Reference(ws_proj, min_col=1, min_row=2, max_row=wc_rows + 1)
        chart_wc.add_data(data_ref, titles_from_data=True)
        chart_wc.set_categories(cats_ref)
        ws_proj.add_chart(chart_wc, f"E2")

        # --- Task Summary ---
        if 'Task' in df_proj.columns:
            start_row = wc_rows + 6
            summary_task = df_proj.groupby('Task')['Hours'].sum().reset_index()
            ws_proj.cell(row=start_row, column=1, value="Task")
            ws_proj.cell(row=start_row, column=2, value="Hours")
            for i, row in enumerate(summary_task.itertuples(index=False), start=start_row + 1):
                ws_proj.cell(row=i, column=1, value=row.Task)
                ws_proj.cell(row=i, column=2, value=row.Hours)

            chart_task = BarChart()
            chart_task.title = f"{project} - Hours by Task"
            chart_task.x_axis.title = "Task"
            chart_task.y_axis.title = "Hours"
            data_ref = Reference(ws_proj, min_col=2, min_row=start_row, max_row=start_row + len(summary_task))
            cats_ref = Reference(ws_proj, min_col=1, min_row=start_row + 1, max_row=start_row + len(summary_task))
            ws_proj.add_chart(chart_task, f"E{start_row}")

        # --- Detailed Data ---
        start_row = ws_proj.max_row + 3
        for col_num, col_name in enumerate(df_proj.columns, start=1):
            ws_proj.cell(row=start_row, column=col_num, value=col_name)
        for i, row in enumerate(df_proj.itertuples(index=False), start=start_row + 1):
            for j, val in enumerate(row, start=1):
                ws_proj.cell(row=i, column=j, value=val)

    # --- Config Info Sheet ---
    ws_config = wb.create_sheet("Config_Info")
    ws_config.append(["Mode", config['mode']])
    ws_config.append(["Years", ', '.join(map(str, config['years'])) if 'years' in config else str(config['year'])])
    ws_config.append(["Months", ', '.join(config['months']) if config['months'] else "All"])
    ws_config.append(["Projects", ', '.join(config['project_filter_df']['Project Name'].dropna())])

    # --- Remove Raw_Data Sheet (if exists) ---
    if 'Raw_Data' in wb.sheetnames:
        wb.remove(wb['Raw_Data'])

    wb.save(path_dict['output_file'])

def export_summary_pdf(summary_df, path_dict, logo_path="triac_logo.png"):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Logo + Title
    if os.path.exists(logo_path):
        pdf.image(logo_path, x=10, y=8, w=25)
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Time Report Summary", ln=True, align="C")
    pdf.set_font("Arial", '', 12)
    pdf.ln(10)

    # Table Header
    col_names = summary_df.columns.tolist()
    col_widths = [40] * len(col_names)
    for i, col in enumerate(col_names):
        pdf.cell(col_widths[i], 10, str(col), border=1, ln=0, align='C')
    pdf.ln()

    # Table Rows
    for row in summary_df.itertuples(index=False):
        for i, val in enumerate(row):
            pdf.cell(col_widths[i], 10, str(val), border=1, ln=0)
        pdf.ln()

    pdf.output(path_dict['pdf_report'])

# Optional: Hook for CLI use
if __name__ == "__main__":
    paths = setup_paths()
    config = read_configs(paths)
    data = load_raw_data(paths)
    filtered = apply_filters(data, config)
    export_report(filtered, config, paths)

    # Optional: generate PDF summary
    if config['mode'] == 'year':
        summary = filtered.groupby(['Year', 'Project name'])['Hours'].sum().reset_index()
    elif config['mode'] == 'month':
        summary = filtered.groupby(['Year', 'MonthName', 'Project name'])['Hours'].sum().reset_index()
    else:
        summary = filtered.groupby(['Year', 'Week', 'Project name'])['Hours'].sum().reset_index()
    export_summary_pdf(summary, paths)
