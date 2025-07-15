import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import datetime
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference

sns.set(style="whitegrid")

def setup_paths():
    today = datetime.datetime.today().strftime('%Y%m%d')
    return {
        'template_file': "Time_report.xlsm",
        'output_file': f"Time_report_{today}.xlsx",
        'pdf_report': f"Time_report_{today}.pdf"  # giữ nguyên cho tương thích, không dùng đến
    }

def read_configs(path_dict):
    year_mode_df = pd.read_excel(path_dict['template_file'], sheet_name='Config_Year_Mode', engine='openpyxl')
    project_filter_df = pd.read_excel(path_dict['template_file'], sheet_name='Config_Project_Filter', engine='openpyxl')

    mode_raw = year_mode_df.loc[year_mode_df['Key'].str.lower() == 'mode', 'Value'].values[0]
    mode = str(mode_raw).strip().lower()

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
    if 'Task' not in df.columns:
        df['Task'] = "Unknown"
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
        df.to_excel(writer, sheet_name='Raw_Data', index=False)
        summary.to_excel(writer, sheet_name='Summary', index=False)

    wb = load_workbook(path_dict['output_file'])
    ws = wb['Summary']
    max_row = ws.max_row

    if mode == 'year':
        data_ref = Reference(ws, min_col=3, min_row=1, max_row=max_row)
        cats_ref = Reference(ws, min_col=2, min_row=2, max_row=max_row)
    elif mode == 'month':
        data_ref = Reference(ws, min_col=4, min_row=1, max_row=max_row)
        cats_ref = Reference(ws, min_col=3, min_row=2, max_row=max_row)
    else:
        data_ref = Reference(ws, min_col=4, min_row=1, max_row=max_row)
        cats_ref = Reference(ws, min_col=3, min_row=2, max_row=max_row)

    chart = BarChart()
    chart.title = f"Total Hours by Project ({mode})"
    chart.x_axis.title = "Project"
    chart.y_axis.title = "Hours"
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws.add_chart(chart, "F2")

    for project in df['Project name'].unique():
        df_proj = df[df['Project name'] == project]
        ws_proj = wb.create_sheet(title=project[:31])

        # === 1. Task Summary Table ===
        task_summary = df_proj.groupby('Task')['Hours'].sum().reset_index()
        ws_proj.cell(row=1, column=1, value="Task")
        ws_proj.cell(row=1, column=2, value="Hours")
        for i, row in enumerate(task_summary.itertuples(index=False), start=2):
            ws_proj.cell(row=i, column=1, value=row.Task)
            ws_proj.cell(row=i, column=2, value=row.Hours)

        # === 2. Task Chart ===
        task_chart = BarChart()
        task_chart.title = f"{project} - Hours by Task"
        task_chart.x_axis.title = "Task"
        task_chart.y_axis.title = "Hours"
        task_data_ref = Reference(ws_proj, min_col=2, min_row=1, max_row=1 + len(task_summary))
        task_cats_ref = Reference(ws_proj, min_col=1, min_row=2, max_row=1 + len(task_summary))
        task_chart.add_data(task_data_ref, titles_from_data=True)
        task_chart.set_categories(task_cats_ref)
        ws_proj.add_chart(task_chart, "E2")

        # === 3. Workcentre Summary Table ===
        wc_start_row = len(task_summary) + 5
        wc_summary = df_proj.groupby('Workcentre')['Hours'].sum().reset_index()
        ws_proj.cell(row=wc_start_row, column=1, value="Workcentre")
        ws_proj.cell(row=wc_start_row, column=2, value="Hours")
        for i, row in enumerate(wc_summary.itertuples(index=False), start=wc_start_row + 1):
            ws_proj.cell(row=i, column=1, value=row.Workcentre)
            ws_proj.cell(row=i, column=2, value=row.Hours)

        # === 4. Workcentre Chart ===
        wc_chart = BarChart()
        wc_chart.title = f"{project} - Hours by Workcentre"
        wc_chart.x_axis.title = "Workcentre"
        wc_chart.y_axis.title = "Hours"
        wc_data_ref = Reference(ws_proj, min_col=2, min_row=wc_start_row, max_row=wc_start_row + len(wc_summary))
        wc_cats_ref = Reference(ws_proj, min_col=1, min_row=wc_start_row + 1, max_row=wc_start_row + len(wc_summary))
        wc_chart.add_data(wc_data_ref, titles_from_data=True)
        wc_chart.set_categories(wc_cats_ref)
        ws_proj.add_chart(wc_chart, f"E{wc_start_row}")

        # === 5. Raw Data (after summaries) ===
        data_start_row = wc_start_row + len(wc_summary) + 5
        for col_idx, col_name in enumerate(df_proj.columns, start=1):
            ws_proj.cell(row=data_start_row, column=col_idx, value=col_name)
        for r_idx, row in enumerate(df_proj.itertuples(index=False), start=data_start_row + 1):
            for c_idx, value in enumerate(row, start=1):
                ws_proj.cell(row=r_idx, column=c_idx, value=value)

    ws_config = wb.create_sheet("Config_Info")
    ws_config['A1'], ws_config['B1'] = "Mode", config['mode']
    ws_config['A2'], ws_config['B2'] = "Years", ', '.join(map(str, config['years'])) if 'years' in config else str(config['year'])
    ws_config['A3'], ws_config['B3'] = "Months", ', '.join(config['months']) if config['months'] else "All"
    project_names = config['project_filter_df']['Project Name'].dropna()
    ws_config['A4'], ws_config['B4'] = "Projects", ', '.join(map(str, project_names))

    wb.save(path_dict['output_file'])

if __name__ == "__main__":
    paths = setup_paths()
    config = read_configs(paths)
    data = load_raw_data(paths)
    filtered_data = apply_filters(data, config)
    export_report(filtered_data, config, paths)
