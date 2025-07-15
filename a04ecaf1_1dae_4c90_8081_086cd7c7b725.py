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
        'pdf_report': f"Time_report_{today}.pdf"
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

        for r_idx, row in enumerate(df_proj.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                ws_proj.cell(row=r_idx, column=c_idx, value=value)

        summary_wc = df_proj.groupby('Workcentre')['Hours'].sum().reset_index()
        start_row = len(df_proj) + 4
        ws_proj.cell(row=start_row, column=1, value="Workcentre")
        ws_proj.cell(row=start_row, column=2, value="Hours")
        for i, row in enumerate(summary_wc.itertuples(index=False), start=start_row + 1):
            ws_proj.cell(row=i, column=1, value=row.Workcentre)
            ws_proj.cell(row=i, column=2, value=row.Hours)

        chart = BarChart()
        chart.title = f"{project} - Hours by Workcentre"
        chart.x_axis.title = "Workcentre"
        chart.y_axis.title = "Hours"
        data_ref = Reference(ws_proj, min_col=2, min_row=start_row, max_row=start_row + len(summary_wc))
        cats_ref = Reference(ws_proj, min_col=1, min_row=start_row + 1, max_row=start_row + len(summary_wc))
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        ws_proj.add_chart(chart, f"E{start_row}")

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
