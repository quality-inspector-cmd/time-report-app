import pandas as pd
import numpy as np
import os
import io
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.font_manager as fm # Thêm dòng này để quản lý font
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.units import inch
from reportlab.lib import colors
from datetime import datetime

# ==============================================================================
# CẤU HÌNH FONT CHO MATPLOTLIB VÀ REPORTLAB
# Đảm bảo sử dụng DejaVu Sans hoặc bất kỳ font nào có sẵn trên môi trường.
# ==============================================================================
import matplotlib
# Đặt font mặc định cho Matplotlib
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['DejaVu Sans'] # DejaVu Sans thường có sẵn

# Đảm bảo PDF nhúng font đúng cách để hiển thị trên mọi máy
matplotlib.rcParams['pdf.fonttype'] = 42 # Type 42 (TrueType)
matplotlib.rcParams['ps.fonttype'] = 42 # Type 42 (TrueType)

# Cấu hình ReportLab để sử dụng font có sẵn hoặc tự định nghĩa nếu cần
# ReportLab sử dụng font riêng, và DejaVu Sans không phải là font mặc định của nó.
# Cần đăng ký font TrueType nếu muốn dùng các font không phải Helvetica/Times/Courier.
# Tuy nhiên, đối với các font cơ bản, ReportLab có thể tự fallback.
# Nếu bạn muốn đảm bảo DejaVu Sans cho ReportLab, bạn sẽ cần file .ttf và đăng ký:
# from reportlab.pdfbase import pdfmetrics
# from reportlab.pdfbase.ttfonts import TTFont
# try:
#     # Thay thế 'path/to/DejaVuSans.ttf' bằng đường dẫn thực tế của font nếu bạn có
#     # Hoặc tải nó từ internet và đặt vào thư mục dự án
#     pdfmetrics.registerFont(TTFont('DejaVuSans', 'DejaVuSans.ttf'))
#     # Sau đó sử dụng 'DejaVuSans' trong ParagraphStyle
# except Exception as e:
#     print(f"Could not register DejaVuSans for ReportLab: {e}. Falling back to default.")

# Mặc định, ReportLab sử dụng Helvetica, Times-Roman, Courier.
# Chúng ta sẽ sử dụng Helvetica làm font mặc định cho ReportLab để đảm bảo tính tương thích
# và tránh lỗi font nếu DejaVu Sans .ttf không được tìm thấy/đăng ký.
# Các style mặc định của ReportLab đã sử dụng font này.

# ==============================================================================


def setup_paths():
    script_dir = os.path.dirname(__file__)
    return {
        "template_file": os.path.join(script_dir, "Time Report Template.xlsm"),
        "output_file": os.path.join(script_dir, "Triac_Time_Report.xlsx"),
        "pdf_report": os.path.join(script_dir, "Triac_Time_Report.pdf"),
        "comparison_output_file": os.path.join(script_dir, "Triac_Comparison_Report.xlsx"),
        "comparison_pdf_report": os.path.join(script_dir, "Triac_Comparison_Report.pdf"),
        "logo_path": os.path.join(script_dir, "triac_logo.png"), # Đảm bảo file logo tồn tại
    }

def load_raw_data(template_path):
    try:
        df = pd.read_excel(template_path, sheet_name="Raw Data")
        # Chuyển đổi cột 'Date' thành datetime và trích xuất năm, tháng
        df['Date'] = pd.to_datetime(df['Date'])
        df['Year'] = df['Date'].dt.year
        df['Month'] = df['Date'].dt.month
        df['MonthName'] = df['Date'].dt.strftime('%B') # Tên tháng đầy đủ
        df['Week'] = df['Date'].dt.isocalendar().week.astype(int) # Số tuần
        df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce').fillna(0)
        return df
    except FileNotFoundError:
        print(f"Error: Template file not found at {template_path}")
        return pd.DataFrame()
    except Exception as e:
        print(f"Error loading raw data: {e}")
        return pd.DataFrame()

def read_configs(template_path):
    config = {}
    try:
        df_config = pd.read_excel(template_path, sheet_name="Config", header=None)
        df_config.set_index(0, inplace=True)
        config['mode'] = df_config.loc['Analysis Mode', 1]
        config['year'] = int(df_config.loc['Year', 1])
        
        months_str = df_config.loc['Month(s)', 1]
        if isinstance(months_str, str):
            config['months'] = [m.strip() for m in months_str.split(',')]
        else:
            config['months'] = [] # Default to empty list if not string

        df_project_filter = pd.read_excel(template_path, sheet_name="Project Filter")
        config['project_filter_df'] = df_project_filter
    except Exception as e:
        print(f"Error reading configurations: {e}")
        # Cung cấp giá trị mặc định an toàn nếu không đọc được config
        config['mode'] = 'year'
        config['year'] = datetime.now().year
        config['months'] = []
        config['project_filter_df'] = pd.DataFrame(columns=['Project Name', 'Include'])
    return config

def apply_filters(df_raw, config):
    df_filtered = df_raw.copy()

    # Filter by year
    if config.get('year'):
        df_filtered = df_filtered[df_filtered['Year'] == config['year']]

    # Filter by months
    if config.get('months') and config['months'] != ['All']:
        df_filtered = df_filtered[df_filtered['MonthName'].isin(config['months'])]

    # Filter by projects based on config['project_filter_df']
    if not config['project_filter_df'].empty:
        included_projects = config['project_filter_df'][
            config['project_filter_df']['Include'].astype(str).str.lower() == 'yes'
        ]['Project Name'].tolist()
        df_filtered = df_filtered[df_filtered['Project name'].isin(included_projects)]
    
    return df_filtered

def export_report(df_filtered, config, output_path):
    try:
        # Load the template workbook
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # Write raw data
            df_filtered.to_excel(writer, sheet_name='Raw Data Filtered', index=False)

            # Create pivot table based on mode
            if config['mode'] == 'year':
                pivot_df = pd.pivot_table(df_filtered,
                                          values='Hours',
                                          index=['Project name'],
                                          aggfunc='sum').reset_index()
                pivot_df.rename(columns={'Hours': 'Total Hours'}, inplace=True)
                title_suffix = f"for Year {config['year']}"
            elif config['mode'] == 'month':
                pivot_df = pd.pivot_table(df_filtered,
                                          values='Hours',
                                          index=['Project name'],
                                          columns=['MonthName'],
                                          aggfunc='sum').reset_index()
                # Ensure month order
                pivot_df = pivot_df[['Project name'] + [m for m in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'] if m in pivot_df.columns]]
                title_suffix = f"for Month(s) {', '.join(config['months'])} in Year {config['year']}"
            elif config['mode'] == 'week':
                pivot_df = pd.pivot_table(df_filtered,
                                          values='Hours',
                                          index=['Project name'],
                                          columns=['Week'],
                                          aggfunc='sum').reset_index()
                pivot_df = pivot_df.set_index('Project name').sort_index(axis=1).reset_index() # Sort weeks numerically
                title_suffix = f"for Weeks in Year {config['year']}"
            else: # Default to year if mode is invalid
                pivot_df = pd.pivot_table(df_filtered,
                                          values='Hours',
                                          index=['Project name'],
                                          aggfunc='sum').reset_index()
                pivot_df.rename(columns={'Hours': 'Total Hours'}, inplace=True)
                title_suffix = f"for Year {config['year']}"


            # Write pivot table to a new sheet
            report_sheet_name = 'Time Report'
            pivot_df.to_excel(writer, sheet_name=report_sheet_name, index=False)

            workbook = writer.book
            worksheet = writer.sheets[report_sheet_name]

            # Add a title
            title_format = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'})
            subtitle_format = workbook.add_format({'bold': False, 'font_size': 12, 'align': 'center', 'valign': 'vcenter'})
            
            # Merge cells for title
            max_col = len(pivot_df.columns) - 1
            worksheet.merge_range(0, 0, 0, max_col, f"Time Report {title_suffix}", title_format)
            
            # Format columns
            currency_format = workbook.add_format({'num_format': '#,##0.00'})
            for col_num, value in enumerate(pivot_df.columns):
                max_len = max(pivot_df[value].astype(str).map(len).max(), len(str(value)))
                worksheet.set_column(col_num, col_num, max_len + 2) # Set column width
                if pivot_df[value].dtype in ['float64', 'int64'] and value != 'Year' and value != 'Week':
                    worksheet.set_column(col_num, col_num, None, currency_format)

            # Adjust column width for 'Project name'
            max_project_name_len = pivot_df['Project name'].astype(str).map(len).max()
            worksheet.set_column(0, 0, max_project_name_len + 2)
            
            # Make header bold
            header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            for col_num, value in enumerate(pivot_df.columns):
                worksheet.write(1, col_num, value, header_format)

            # Ensure data starts from row 2
            for r_idx, row_data in enumerate(pivot_df.values):
                for c_idx, cell_data in enumerate(row_data):
                    worksheet.write(r_idx + 2, c_idx, cell_data)

        return True
    except Exception as e:
        print(f"Error exporting report: {e}")
        return False

def export_pdf_report(df_filtered, config, output_pdf_path, logo_path):
    try:
        # Create pivot table for PDF report
        if config['mode'] == 'year':
            pivot_df = pd.pivot_table(df_filtered,
                                      values='Hours',
                                      index=['Project name'],
                                      aggfunc='sum').reset_index()
            pivot_df.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            title_suffix = f"for Year {config['year']}"
            plot_title = f"Total Hours per Project for Year {config['year']}"
            x_label = "Project Name"
            y_label = "Total Hours"
        elif config['mode'] == 'month':
            pivot_df = pd.pivot_table(df_filtered,
                                      values='Hours',
                                      index=['Project name'],
                                      columns=['MonthName'],
                                      aggfunc='sum').reset_index()
            # Ensure month order
            pivot_df = pivot_df[['Project name'] + [m for m in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'] if m in pivot_df.columns]]
            title_suffix = f"for Month(s) {', '.join(config['months'])} in Year {config['year']}"
            plot_title = f"Hours per Project by Month for Year {config['year']}"
            x_label = "Project Name"
            y_label = "Hours"
        elif config['mode'] == 'week':
            pivot_df = pd.pivot_table(df_filtered,
                                      values='Hours',
                                      index=['Project name'],
                                      columns=['Week'],
                                      aggfunc='sum').reset_index()
            pivot_df = pivot_df.set_index('Project name').sort_index(axis=1).reset_index()
            title_suffix = f"for Weeks in Year {config['year']}"
            plot_title = f"Hours per Project by Week for Year {config['year']}"
            x_label = "Project Name"
            y_label = "Hours"
        else: # Default to year
            pivot_df = pd.pivot_table(df_filtered,
                                      values='Hours',
                                      index=['Project name'],
                                      aggfunc='sum').reset_index()
            pivot_df.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            title_suffix = f"for Year {config['year']}"
            plot_title = f"Total Hours per Project for Year {config['year']}"
            x_label = "Project Name"
            y_label = "Total Hours"
        
        if pivot_df.empty:
            print("No data to generate PDF report.")
            return False

        # Generate a plot (Bar Chart)
        plt.figure(figsize=(10, 6))
        
        # Determine how to plot based on columns
        if config['mode'] == 'year':
            plt.bar(pivot_df['Project name'], pivot_df['Total Hours'])
            plt.xticks(rotation=45, ha="right")
        else:
            # For month/week reports, we might need a grouped bar chart
            # Or just sum up all columns for a single bar chart for simplicity in PDF
            # For this example, let's sum up for a single bar chart if multiple columns exist,
            # otherwise, the plot would become too complex to generalize easily for PDF.
            # A more advanced solution would be to generate multiple charts or a stacked bar chart.
            
            # Exclude 'Project name' from numeric columns
            numeric_cols = pivot_df.select_dtypes(include=np.number).columns.tolist()
            if 'Year' in numeric_cols: # Remove Year if it appears as a column
                numeric_cols.remove('Year')
            
            if not numeric_cols:
                print("No numeric data columns to plot for PDF.")
                return False

            if len(numeric_cols) > 1: # Multiple months/weeks
                # Sum across all months/weeks for a total per project for the bar chart
                pivot_df['Total Hours for Plot'] = pivot_df[numeric_cols].sum(axis=1)
                plt.bar(pivot_df['Project name'], pivot_df['Total Hours for Plot'])
                plt.xticks(rotation=45, ha="right")
                plot_title = f"Total Hours per Project {title_suffix}"
            else: # Single month/week, or 'Total Hours' if it was a year report
                plt.bar(pivot_df['Project name'], pivot_df[numeric_cols[0]])
                plt.xticks(rotation=45, ha="right")
                plot_title = f"Hours per Project {title_suffix}"
                
        plt.title(plot_title)
        plt.xlabel(x_label)
        plt.ylabel(y_label)
        plt.tight_layout()

        # Save plot to a BytesIO object
        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png', bbox_inches='tight')
        img_buffer.seek(0)
        plt.close() # Close the plot to free memory

        # Set up ReportLab
        doc = SimpleDocTemplate(output_pdf_path, pagesize=landscape(letter))
        styles = getSampleStyleSheet()

        # Custom styles for ReportLab (using Helvetica for better compatibility on various systems)
        # Nếu muốn dùng DejaVu Sans trong ReportLab, bạn cần đăng ký font.
        # Ở đây, tôi tạm thời sử dụng Helvetica, font mặc định của ReportLab.
        
        styles.add(ParagraphStyle(name='TitleStyle',
                                  parent=styles['h1'],
                                  fontName='Helvetica-Bold', # Sử dụng font Helvetica
                                  fontSize=24,
                                  spaceAfter=14,
                                  alignment=TA_CENTER))
        styles.add(ParagraphStyle(name='SubtitleStyle',
                                  parent=styles['h2'],
                                  fontName='Helvetica', # Sử dụng font Helvetica
                                  fontSize=14,
                                  spaceAfter=12,
                                  alignment=TA_CENTER))
        styles.add(ParagraphStyle(name='TableTitleStyle',
                                  parent=styles['h3'],
                                  fontName='Helvetica-Bold', # Sử dụng font Helvetica
                                  fontSize=16,
                                  spaceAfter=6,
                                  alignment=TA_LEFT))
        styles.add(ParagraphStyle(name='Normal',
                                  parent=styles['Normal'],
                                  fontName='Helvetica', # Sử dụng font Helvetica
                                  fontSize=10,
                                  leading=12))

        elements = []

        # Add logo
        if os.path.exists(logo_path):
            img = Image(logo_path)
            img.drawHeight = 0.8 * inch * img.drawHeight / img.drawWidth
            img.drawWidth = 0.8 * inch
            elements.append(img)
            elements.append(Spacer(1, 0.2 * inch))

        # Add title and subtitle
        elements.append(Paragraph("Triac Time Report", styles['TitleStyle']))
        elements.append(Paragraph(f"Summary Report {title_suffix}", styles['SubtitleStyle']))
        elements.append(Spacer(1, 0.4 * inch))

        # Add plot image
        elements.append(Image(img_buffer, width=7.5*inch, height=4.5*inch))
        elements.append(Spacer(1, 0.4 * inch))

        # Prepare data for table
        data = [pivot_df.columns.tolist()] + pivot_df.values.tolist()

        # Table Style
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#003366')), # Header background
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), # Header text color
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), # Header font
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige), # Data background
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'), # Data font
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('RIGHTPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ])
        
        # Apply conditional formatting for numbers (right align, format as currency if applicable)
        for i, col_name in enumerate(pivot_df.columns):
            if pivot_df[col_name].dtype in ['float64', 'int64'] and col_name not in ['Year', 'Week']:
                table_style.add('ALIGN', (i, 1), (i, -1), 'RIGHT')
                # Custom format for numbers in table, ReportLab doesn't have direct number format like Excel
                # You'd need to pre-format the data in the 'data' list if precise formatting is needed.
                # For simplicity, we just align right here.

        table = Table(data, hAlign='CENTER')
        table.setStyle(table_style)
        elements.append(Paragraph("Detailed Data:", styles['TableTitleStyle']))
        elements.append(table)

        doc.build(elements)
        return True
    except Exception as e:
        print(f"Error generating PDF report: {e}")
        return False


def apply_comparison_filters(df_raw, comparison_config, comparison_mode_internal_string):
    df_filtered = df_raw.copy()
    message = ""

    # Filter by selected projects first
    if comparison_config.get('selected_projects'):
        df_filtered = df_filtered[df_filtered['Project name'].isin(comparison_config['selected_projects'])]
    else:
        message = "No projects selected for comparison."
        return pd.DataFrame(), message

    if comparison_mode_internal_string == "So Sánh Dự Án Trong Một Tháng" or comparison_mode_internal_string == "Compare Projects in a Month":
        if not comparison_config.get('months') or not comparison_config.get('years') or len(comparison_config['years']) != 1:
            message = "For 'Compare Projects in a Month' mode, please select exactly one year and one or more months."
            return pd.DataFrame(), message
        
        selected_year = comparison_config['years'][0]
        selected_months = comparison_config['months']
        df_filtered = df_filtered[
            (df_filtered['Year'] == selected_year) & 
            (df_filtered['MonthName'].isin(selected_months))
        ]
        
        # Aggregate by Project and Month
        df_pivot = pd.pivot_table(df_filtered, 
                                  values='Hours', 
                                  index='Project name', 
                                  columns='MonthName', 
                                  aggfunc='sum',
                                  fill_value=0)
        # Ensure month order
        df_pivot = df_pivot[[m for m in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'] if m in df_pivot.columns]]
        df_pivot['Total Hours'] = df_pivot.sum(axis=1) # Add total column
        df_pivot = df_pivot.reset_index()
        message = f"Comparing projects for year {selected_year}, month(s): {', '.join(selected_months)}"

    elif comparison_mode_internal_string == "So Sánh Dự Án Trong Một Năm" or comparison_mode_internal_string == "Compare Projects in a Year":
        if not comparison_config.get('years') or len(comparison_config['years']) != 1:
            message = "For 'Compare Projects in a Year' mode, please select exactly one year."
            return pd.DataFrame(), message
        
        selected_year = comparison_config['years'][0]
        df_filtered = df_filtered[df_filtered['Year'] == selected_year]

        # Aggregate by Project and Month (all months of the year)
        df_pivot = pd.pivot_table(df_filtered, 
                                  values='Hours', 
                                  index='Project name', 
                                  columns='MonthName', 
                                  aggfunc='sum',
                                  fill_value=0)
        # Ensure month order
        df_pivot = df_pivot[[m for m in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'] if m in df_pivot.columns]]
        df_pivot['Total Hours'] = df_pivot.sum(axis=1) # Add total column
        df_pivot = df_pivot.reset_index()
        message = f"Comparing projects for year {selected_year}"

    elif comparison_mode_internal_string == "So Sánh Một Dự Án Qua Các Tháng/Năm" or comparison_mode_internal_string == "Compare One Project Over Time (Months/Years)":
        if len(comparison_config.get('selected_projects', [])) != 1:
            message = "Please select exactly one project for 'Compare One Project Over Time' mode."
            return pd.DataFrame(), message

        selected_project = comparison_config['selected_projects'][0]
        df_filtered = df_filtered[df_filtered['Project name'] == selected_project]

        if comparison_config.get('years') and len(comparison_config['years']) > 1:
            # Compare across multiple years (total hours per year for the selected project)
            df_pivot = pd.pivot_table(df_filtered,
                                      values='Hours',
                                      index='Year',
                                      aggfunc='sum',
                                      fill_value=0).reset_index()
            df_pivot.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            message = f"Comparing '{selected_project}' across years: {', '.join(map(str, comparison_config['years']))}"
        elif comparison_config.get('years') and len(comparison_config['years']) == 1:
            selected_year = comparison_config['years'][0]
            df_filtered = df_filtered[df_filtered['Year'] == selected_year]
            if not comparison_config.get('months'):
                message = "Please select months for 'Compare One Project Over Time' within a single year."
                return pd.DataFrame(), message
            
            # Compare across months within a single year for the selected project
            df_filtered = df_filtered[df_filtered['MonthName'].isin(comparison_config['months'])]
            df_pivot = pd.pivot_table(df_filtered,
                                      values='Hours',
                                      index='MonthName',
                                      aggfunc='sum',
                                      fill_value=0).reset_index()
            df_pivot.rename(columns={'Hours': 'Total Hours'}, inplace=True)
            df_pivot['MonthName'] = pd.Categorical(df_pivot['MonthName'], categories=df_raw['MonthName'].cat.categories, ordered=True)
            df_pivot = df_pivot.sort_values('MonthName')
            message = f"Comparing '{selected_project}' across months {', '.join(comparison_config['months'])} in year {selected_year}"
        else:
            message = "Please select at least one year for 'Compare One Project Over Time' mode."
            return pd.DataFrame(), message
    else:
        message = "Invalid comparison mode selected."
        return pd.DataFrame(), message

    if df_pivot.empty:
        message = message + " No data found after filtering."
    return df_pivot, message


def export_comparison_report(df_comparison, comparison_config, output_path, comparison_mode_internal_string):
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            report_sheet_name = 'Comparison Report'
            df_comparison.to_excel(writer, sheet_name=report_sheet_name, index=False)

            workbook = writer.book
            worksheet = writer.sheets[report_sheet_name]

            title_format = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'})
            
            report_title = ""
            if comparison_mode_internal_string == "So Sánh Dự Án Trong Một Tháng" or comparison_mode_internal_string == "Compare Projects in a Month":
                report_title = f"Project Comparison for Year {comparison_config['years'][0]}, Month(s): {', '.join(comparison_config['months'])}"
            elif comparison_mode_internal_string == "So Sánh Dự Án Trong Một Năm" or comparison_mode_internal_string == "Compare Projects in a Year":
                report_title = f"Project Comparison for Year {comparison_config['years'][0]}"
            elif comparison_mode_internal_string == "So Sánh Một Dự Án Qua Các Tháng/Năm" or comparison_mode_internal_string == "Compare One Project Over Time (Months/Years)":
                project_name = comparison_config['selected_projects'][0]
                if comparison_config.get('years') and len(comparison_config['years']) > 1:
                    report_title = f"Hours for Project '{project_name}' across Years: {', '.join(map(str, comparison_config['years']))}"
                elif comparison_config.get('years') and len(comparison_config['years']) == 1:
                    report_title = f"Hours for Project '{project_name}' in Year {comparison_config['years'][0]} across Months: {', '.join(comparison_config['months'])}"

            max_col = len(df_comparison.columns) - 1
            worksheet.merge_range(0, 0, 0, max_col, report_title, title_format)

            # Format columns
            currency_format = workbook.add_format({'num_format': '#,##0.00'})
            for col_num, value in enumerate(df_comparison.columns):
                max_len = max(df_comparison[value].astype(str).map(len).max(), len(str(value)))
                worksheet.set_column(col_num, col_num, max_len + 2)
                if df_comparison[value].dtype in ['float64', 'int64'] and value not in ['Year', 'Week']:
                    worksheet.set_column(col_num, col_num, None, currency_format)
            
            # Make header bold
            header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            for col_num, value in enumerate(df_comparison.columns):
                worksheet.write(1, col_num, value, header_format)

            # Ensure data starts from row 2
            for r_idx, row_data in enumerate(df_comparison.values):
                for c_idx, cell_data in enumerate(row_data):
                    worksheet.write(r_idx + 2, c_idx, cell_data)

        return True
    except Exception as e:
        print(f"Error exporting comparison report: {e}")
        return False

def export_comparison_pdf_report(df_comparison, comparison_config, output_pdf_path, comparison_mode_internal_string, logo_path):
    try:
        # Generate plot based on comparison mode
        plt.figure(figsize=(10, 6))
        plot_title = ""
        x_label = ""
        y_label = "Hours"

        if comparison_mode_internal_string == "So Sánh Dự Án Trong Một Tháng" or comparison_mode_internal_string == "Compare Projects in a Month":
            plot_title = f"Project Comparison for Year {comparison_config['years'][0]}, Month(s): {', '.join(comparison_config['months'])}"
            x_label = "Project Name"
            # For multiple months, plot total hours, or plot multiple bars if needed
            numeric_cols = df_comparison.select_dtypes(include=np.number).columns.tolist()
            if 'Total Hours' in numeric_cols:
                plt.bar(df_comparison['Project name'], df_comparison['Total Hours'])
            elif numeric_cols:
                # If no 'Total Hours', sum up whatever numeric columns are present
                df_comparison['Summed Hours'] = df_comparison[numeric_cols].sum(axis=1)
                plt.bar(df_comparison['Project name'], df_comparison['Summed Hours'])
            else:
                 print("No numeric data to plot for this comparison mode.")
                 return False
            plt.xticks(rotation=45, ha="right")

        elif comparison_mode_internal_string == "So Sánh Dự Án Trong Một Năm" or comparison_mode_internal_string == "Compare Projects in a Year":
            plot_title = f"Project Comparison for Year {comparison_config['years'][0]}"
            x_label = "Project Name"
            numeric_cols = df_comparison.select_dtypes(include=np.number).columns.tolist()
            if 'Total Hours' in numeric_cols:
                plt.bar(df_comparison['Project name'], df_comparison['Total Hours'])
            elif numeric_cols:
                df_comparison['Summed Hours'] = df_comparison[numeric_cols].sum(axis=1)
                plt.bar(df_comparison['Project name'], df_comparison['Summed Hours'])
            else:
                 print("No numeric data to plot for this comparison mode.")
                 return False
            plt.xticks(rotation=45, ha="right")

        elif comparison_mode_internal_string == "So Sánh Một Dự Án Qua Các Tháng/Năm" or comparison_mode_internal_string == "Compare One Project Over Time (Months/Years)":
            project_name = comparison_config['selected_projects'][0]
            if comparison_config.get('years') and len(comparison_config['years']) > 1:
                plot_title = f"Hours for Project '{project_name}' across Years"
                x_label = "Year"
                plt.bar(df_comparison['Year'].astype(str), df_comparison['Total Hours'])
                plt.xticks(rotation=45, ha="right")
            elif comparison_config.get('years') and len(comparison_config['years']) == 1:
                plot_title = f"Hours for Project '{project_name}' in Year {comparison_config['years'][0]} across Months"
                x_label = "Month"
                plt.bar(df_comparison['MonthName'], df_comparison['Total Hours'])
                plt.xticks(rotation=45, ha="right")

        plt.title(plot_title)
        plt.xlabel(x_label)
        plt.ylabel(y_label)
        plt.tight_layout()

        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png', bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()

        # ReportLab setup
        doc = SimpleDocTemplate(output_pdf_path, pagesize=landscape(letter))
        styles = getSampleStyleSheet()

        # Using Helvetica for ReportLab as DejaVuSans TTF is not registered by default
        styles.add(ParagraphStyle(name='TitleStyle', parent=styles['h1'],
                                  fontName='Helvetica-Bold', fontSize=24, spaceAfter=14, alignment=TA_CENTER))
        styles.add(ParagraphStyle(name='SubtitleStyle', parent=styles['h2'],
                                  fontName='Helvetica', fontSize=14, spaceAfter=12, alignment=TA_CENTER))
        styles.add(ParagraphStyle(name='TableTitleStyle', parent=styles['h3'],
                                  fontName='Helvetica-Bold', fontSize=16, spaceAfter=6, alignment=TA_LEFT))
        styles.add(ParagraphStyle(name='Normal', parent=styles['Normal'],
                                  fontName='Helvetica', fontSize=10, leading=12))

        elements = []

        # Add logo
        if os.path.exists(logo_path):
            img = Image(logo_path)
            img.drawHeight = 0.8 * inch * img.drawHeight / img.drawWidth
            img.drawWidth = 0.8 * inch
            elements.append(img)
            elements.append(Spacer(1, 0.2 * inch))

        elements.append(Paragraph("Triac Time Report - Comparison", styles['TitleStyle']))
        elements.append(Paragraph(plot_title, styles['SubtitleStyle']))
        elements.append(Spacer(1, 0.4 * inch))

        elements.append(Image(img_buffer, width=7.5*inch, height=4.5*inch))
        elements.append(Spacer(1, 0.4 * inch))

        # Prepare data for table
        data = [df_comparison.columns.tolist()] + df_comparison.values.tolist()

        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#003366')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('RIGHTPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ])

        for i, col_name in enumerate(df_comparison.columns):
            if df_comparison[col_name].dtype in ['float64', 'int64'] and col_name not in ['Year', 'Week']:
                table_style.add('ALIGN', (i, 1), (i, -1), 'RIGHT')

        table = Table(data, hAlign='CENTER')
        table.setStyle(table_style)
        elements.append(Paragraph("Detailed Data:", styles['TableTitleStyle']))
        elements.append(table)

        doc.build(elements)
        return True
    except Exception as e:
        print(f"Error generating comparison PDF report: {e}")
        return False
