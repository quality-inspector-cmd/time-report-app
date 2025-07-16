import pandas as pd
import datetime
import os
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from fpdf import FPDF

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

    mode = str(year_mode_df.loc[year_mode_df['Key'].str.lower() == 'mode', 'Value'].values[0]).strip().lower()
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
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.dataframe import dataframe_to_rows

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
        # df.to_excel(writer, sheet_name='Raw_Data', index=False)  # Optional: can skip writing raw

    wb = load_workbook(path_dict['output_file'])

    # Add chart to Summary
    ws = wb['Summary']
    max_row = ws.max_row
    data_col = 4 if mode in ['month', 'week'] else 3
    cats_col = 3 if mode in ['month', 'week'] else 2

    data_ref = Reference(ws, min_col=data_col, min_row=1, max_row=max_row)
    cats_ref = Reference(ws, min_col=cats_col, min_row=2, max_row=max_row)

    chart = BarChart()
    chart.title = f"Total Hours by Project ({mode})"
    chart.x_axis.title = "Project"
    chart.y_axis.title = "Hours"
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws.add_chart(chart, "F2")

    # Create project sheets
    for project in df['Project name'].unique():
        df_proj = df[df['Project name'] == project]
        ws_proj = wb.create_sheet(title=project[:31])

        # Task summary
        summary_task = df_proj.groupby('Task')['Hours'].sum().reset_index().sort_values('Hours', ascending=False)
        ws_proj.append(['Task', 'Hours'])
        for row in summary_task.itertuples(index=False):
            ws_proj.append([row.Task, row.Hours])

        # Chart for Task
        chart = BarChart()
        chart.title = f"{project} - Hours by Task"
        chart.x_axis.title = "Task"
        chart.y_axis.title = "Hours"
        task_len = len(summary_task)
        data_ref = Reference(ws_proj, min_col=2, min_row=1, max_row=task_len + 1)
        cats_ref = Reference(ws_proj, min_col=1, min_row=2, max_row=task_len + 1)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        ws_proj.add_chart(chart, f"E1")

        # Add raw data below
        start_row = task_len + 5
        for r in dataframe_to_rows(df_proj, index=False, header=True):
            ws_proj.append(r) if ws_proj.max_row < start_row else ws_proj.append([''] * len(r))
            for i, cell_val in enumerate(r, start=1):
                ws_proj.cell(row=ws_proj.max_row, column=i, value=cell_val)

    # Config info
    ws_config = wb.create_sheet("Config_Info")
    ws_config['A1'], ws_config['B1'] = "Mode", config['mode']
    ws_config['A2'], ws_config['B2'] = "Years", ', '.join(map(str, config['years'])) if 'years' in config else str(config['year'])
    ws_config['A3'], ws_config['B3'] = "Months", ', '.join(config['months']) if config['months'] else "All"
    ws_config['A4'], ws_config['B4'] = "Projects", ', '.join(config['project_filter_df']['Project Name'])

    # Optionally remove 'Raw_Data' if exists
    if 'Raw_Data' in wb.sheetnames:
        del wb['Raw_Data']

    wb.save(path_dict['output_file'])

def export_pdf_report(df, config, path_dict):
    from matplotlib import pyplot as plt
    from fpdf import FPDF
    import tempfile
    import os
    import re

    def sanitize_filename(name):
        return re.sub(r'[\\/*?:"<>|]', "_", name)

    tmp_dir = tempfile.mkdtemp()
    pdf_images = []

    logo_path = "triac_logo.png"
    mode = config['mode']
    today_str = datetime.datetime.today().strftime("%Y-%m-%d")

    projects = df['Project name'].unique()
    for project in projects:
        safe_project = sanitize_filename(project)
        df_proj = df[df['Project name'] == project]

        # --- Workcentre Chart ---
        fig, ax = plt.subplots(figsize=(8, 4))
        df_proj.groupby('Workcentre')['Hours'].sum().sort_values().plot(kind='barh', color='skyblue', ax=ax)
        ax.set_title(f"{project} - Hours by Workcentre", fontsize=10)
        wc_img_path = os.path.join(tmp_dir, f"{safe_project}_wc.png")
        plt.tight_layout()
        fig.savefig(wc_img_path, dpi=150)
        plt.close(fig)
        pdf_images.append(wc_img_path)

        # --- Task Chart ---
        if 'Task' in df_proj.columns:
            fig, ax = plt.subplots(figsize=(8, 4))
            df_proj.groupby('Task')['Hours'].sum().sort_values().plot(kind='barh', color='lightgreen', ax=ax)
            ax.set_title(f"{project} - Hours by Task", fontsize=10)
            task_img_path = os.path.join(tmp_dir, f"{safe_project}_task.png")
            plt.tight_layout()
            fig.savefig(task_img_path, dpi=150)
            plt.close(fig)
            pdf_images.append(task_img_path)

    # --- Build PDF ---
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    for img_path in pdf_images:
        pdf.add_page()
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=10, y=8, w=25)
        pdf.set_font("Arial", size=10)
        pdf.set_y(35)
        pdf.image(img_path, x=10, y=40, w=190)
    pdf.output(path_dict['pdf_report'], "F")

