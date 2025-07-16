import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors
import base64
from PIL import Image as PILImage
from datetime import datetime

# --- Constants for localization (You can expand this if needed) ---
MONTH_NAMES_EN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
MONTH_NAMES_VI = ['Tháng 1', 'Tháng 2', 'Tháng 3', 'Tháng 4', 'Tháng 5', 'Tháng 6', 'Tháng 7', 'Tháng 8', 'Tháng 9', 'Tháng 10', 'Tháng 11', 'Tháng 12']

def load_and_preprocess_data(uploaded_file):
    """
    Tải và tiền xử lý dữ liệu từ file Excel đã tải lên.
    Chuyển đổi các cột 'Start date', 'End date' sang datetime và trích xuất tháng/năm.
    """
    df = pd.read_excel(uploaded_file)

    # Convert date columns to datetime, handling potential errors
    for col in ['Start date', 'End date']:
        # Attempt to convert, coerce errors will turn invalid dates into NaT
        df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Drop rows where 'Start date' or 'End date' could not be parsed
    df.dropna(subset=['Start date', 'End date'], inplace=True)

    df['Month'] = df['Start date'].dt.month
    df['Year'] = df['Start date'].dt.year
    df['MonthName'] = df['Start date'].dt.strftime('%B') # English month names
    df['Week'] = df['Start date'].dt.isocalendar().week.astype(int)

    # Ensure 'Hours' column is numeric
    df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce')
    df.dropna(subset=['Hours'], inplace=True)

    return df

def get_unique_values(df, column):
    """Lấy danh sách các giá trị duy nhất từ một cột."""
    return sorted(df[column].unique().tolist())

def get_min_max_years(df):
    """Lấy năm nhỏ nhất và lớn nhất từ dữ liệu."""
    if 'Year' in df.columns and not df.empty:
        return int(df['Year'].min()), int(df['Year'].max())
    return None, None

def calculate_monthly_summary(df_filtered):
    """Tính toán tổng số giờ và chi phí hàng tháng."""
    monthly_summary = df_filtered.groupby('MonthName').agg(
        total_hours=('Hours', 'sum'),
        total_cost=('Total cost (USD)', 'sum')
    ).reset_index()

    # Ensure month order for plotting
    month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    
    # Filter month_order to only include months actually present in the data to avoid errors if some months have no data
    present_months = [m for m in month_order if m in monthly_summary['MonthName'].unique()]
    
    if present_months: # Only apply if there are months to categorize
        monthly_summary['MonthName'] = pd.Categorical(monthly_summary['MonthName'], categories=present_months, ordered=True)
        monthly_summary = monthly_summary.sort_values('MonthName')
    
    return monthly_summary

def calculate_project_summary(df_filtered):
    """Tính toán tổng số giờ và chi phí cho từng dự án."""
    project_summary = df_filtered.groupby('Project name').agg(
        total_hours=('Hours', 'sum'),
        total_cost=('Total cost (USD)', 'sum')
    ).reset_index()
    project_summary = project_summary.sort_values(by='total_hours', ascending=False)
    return project_summary

def create_monthly_summary_chart(monthly_summary, year):
    """Tạo biểu đồ tổng quan hàng tháng (giờ và chi phí)."""
    if monthly_summary.empty:
        fig = go.Figure()
        fig.add_annotation(text="Không có dữ liệu để hiển thị biểu đồ tổng quan hàng tháng.",
                           xref="paper", yref="paper", showarrow=False,
                           font=dict(size=16, color="grey"))
        fig.update_layout(height=400)
        return fig

    fig = make_subplots(rows=1, cols=2, shared_yaxes=False,
                        subplot_titles=[f'Tổng giờ theo tháng ({year})', f'Tổng chi phí theo tháng ({year})'])

    # Bar chart for Total Hours
    fig.add_trace(
        go.Bar(name='Tổng giờ', x=monthly_summary['MonthName'], y=monthly_summary['total_hours'],
               marker_color='skyblue', hovertemplate='Tháng: %{x}<br>Tổng giờ: %{y:,.0f}<extra></extra>'),
        row=1, col=1
    )

    # Bar chart for Total Cost
    fig.add_trace(
        go.Bar(name='Tổng chi phí', x=monthly_summary['MonthName'], y=monthly_summary['total_cost'],
               marker_color='lightcoral', hovertemplate='Tháng: %{x}<br>Tổng chi phí: %{y:,.2f} USD<extra></extra>'),
        row=1, col=2
    )

    fig.update_layout(
        title_text=f"Biểu đồ tổng quan hàng tháng cho năm {year}",
        title_x=0.5,
        height=400,
        showlegend=False,
        margin=dict(l=20, r=20, t=60, b=20)
    )
    fig.update_xaxes(title_text="Tháng", row=1, col=1)
    fig.update_yaxes(title_text="Tổng giờ", row=1, col=1)
    fig.update_xaxes(title_text="Tháng", row=1, col=2)
    fig.update_yaxes(title_text="Tổng chi phí (USD)", row=1, col=2)

    return fig

def create_project_summary_chart(project_summary, year):
    """Tạo biểu đồ tổng quan theo dự án (giờ và chi phí)."""
    if project_summary.empty:
        fig = go.Figure()
        fig.add_annotation(text="Không có dữ liệu để hiển thị biểu đồ tổng quan dự án.",
                           xref="paper", yref="paper", showarrow=False,
                           font=dict(size=16, color="grey"))
        fig.update_layout(height=400)
        return fig

    fig = make_subplots(rows=1, cols=2, shared_yaxes=False,
                        subplot_titles=[f'Tổng giờ theo dự án ({year})', f'Tổng chi phí theo dự án ({year})'])

    # Bar chart for Total Hours
    fig.add_trace(
        go.Bar(name='Tổng giờ', x=project_summary['Project name'], y=project_summary['total_hours'],
               marker_color='teal', hovertemplate='Dự án: %{x}<br>Tổng giờ: %{y:,.0f}<extra></extra>'),
        row=1, col=1
    )

    # Bar chart for Total Cost
    fig.add_trace(
        go.Bar(name='Tổng chi phí', x=project_summary['Project name'], y=project_summary['total_cost'],
               marker_color='orange', hovertemplate='Dự án: %{x}<br>Tổng chi phí: %{y:,.2f} USD<extra></extra>'),
        row=1, col=2
    )

    fig.update_layout(
        title_text=f"Biểu đồ tổng quan dự án cho năm {year}",
        title_x=0.5,
        height=400,
        showlegend=False,
        margin=dict(l=20, r=20, t=60, b=20)
    )
    fig.update_xaxes(title_text="Dự án", row=1, col=1, tickangle=45)
    fig.update_yaxes(title_text="Tổng giờ", row=1, col=1)
    fig.update_xaxes(title_text="Dự án", row=1, col=2, tickangle=45)
    fig.update_yaxes(title_text="Tổng chi phí (USD)", row=1, col=2)

    return fig

def create_raw_data_table(df_filtered):
    """Tạo bảng dữ liệu thô."""
    if df_filtered.empty:
        return go.Figure().add_annotation(text="Không có dữ liệu để hiển thị.",
                                           xref="paper", yref="paper", showarrow=False,
                                           font=dict(size=16, color="grey")).update_layout(height=100)

    # Select relevant columns for display in the raw data table
    display_df = df_filtered[['Project name', 'Start date', 'End date', 'Hours', 'Total cost (USD)']].copy()
    
    # Format dates for better readability in the table
    display_df['Start date'] = display_df['Start date'].dt.strftime('%Y-%m-%d')
    display_df['End date'] = display_df['End date'].dt.strftime('%Y-%m-%d')

    fig = go.Figure(data=[go.Table(
        header=dict(values=list(display_df.columns),
                    fill_color='paleturquoise',
                    align='left'),
        cells=dict(values=[display_df[col] for col in display_df.columns],
                   fill_color='lavender',
                   align='left'))
    ])
    fig.update_layout(title_text="Dữ liệu thô đã lọc", title_x=0.5, height=400)
    return fig

def create_overall_summary_table(df_filtered):
    """Tạo bảng tổng quan chung về giờ và chi phí."""
    total_hours = df_filtered['Hours'].sum()
    total_cost = df_filtered['Total cost (USD)'].sum()

    summary_data = {
        "Metric": ["Tổng giờ làm việc", "Tổng chi phí (USD)"],
        "Value": [f"{total_hours:,.0f}", f"{total_cost:,.2f}"]
    }
    summary_df = pd.DataFrame(summary_data)

    fig = go.Figure(data=[go.Table(
        header=dict(values=list(summary_df.columns),
                    fill_color='paleturquoise',
                    align='left'),
        cells=dict(values=[summary_df[col] for col in summary_df.columns],
                   fill_color='lavender',
                   align='left'))
    ])
    fig.update_layout(title_text="Tổng quan chung", title_x=0.5, height=150)
    return fig

def apply_filters(df_raw, year=None, month_name=None, project_name=None):
    """
    Áp dụng bộ lọc cho DataFrame.
    """
    df_filtered = df_raw.copy()

    if year:
        df_filtered = df_filtered[df_filtered['Year'] == year]
    if month_name:
        df_filtered = df_filtered[df_filtered['MonthName'] == month_name]
    if project_name:
        df_filtered = df_filtered[df_filtered['Project name'] == project_name]
    
    return df_filtered

def apply_comparison_filters(df_raw, comparison_config, comparison_mode):
    """Áp dụng bộ lọc và tạo DataFrame tóm tắt cho báo cáo so sánh."""
    years = comparison_config.get('years', [])
    months = comparison_config.get('months', [])
    selected_projects = comparison_config.get('selected_projects', [])

    df_filtered = df_raw.copy()

    if years:
        df_filtered = df_filtered[df_filtered['Year'].isin(years)]
    
    if months:
        df_filtered = df_filtered[df_filtered['MonthName'].isin(months)]
    
    if selected_projects:
        df_filtered = df_filtered[df_filtered['Project name'].isin(selected_projects)]
    else: 
        return pd.DataFrame(), "Vui lòng chọn ít nhất một dự án để so sánh."

    if df_filtered.empty:
        return pd.DataFrame(), f"Không tìm thấy dữ liệu cho chế độ so sánh: {comparison_mode} với các lựa chọn hiện tại."

    title = ""

    if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
        if len(years) != 1 or len(months) != 1 or len(selected_projects) < 2:
            return pd.DataFrame(), "Vui lòng chọn MỘT năm, MỘT tháng và ít nhất HAI dự án cho chế độ này."
        
        df_comparison = df_filtered.groupby('Project name')['Hours'].sum().reset_index()
        df_comparison.rename(columns={'Hours': 'Total Hours'}, inplace=True)
        title = f"So sánh giờ giữa các dự án trong {months[0]}, năm {years[0]}"
        return df_comparison, title

    elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        if len(years) != 1 or len(selected_projects) < 2:
            return pd.DataFrame(), "Vui lòng chọn MỘT năm và ít nhất HAI dự án cho chế độ này."
        
        df_comparison = df_filtered.groupby(['Project name', 'MonthName'])['Hours'].sum().unstack(fill_value=0)
        
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'July', 'August', 'September', 'October', 'November', 'December']
        existing_months = [m for m in month_order if m in df_comparison.columns]
        df_comparison = df_comparison[existing_months]

        df_comparison = df_comparison.reset_index().rename(columns={'index': 'Project Name'})
        
        df_comparison['Total Hours'] = df_comparison[existing_months].sum(axis=1)

        # Kiểm tra trước khi thêm hàng 'Total'
        if not df_comparison.empty:
            df_comparison.loc['Total'] = df_comparison[existing_months + ['Total Hours']].sum()
            df_comparison.loc['Total', 'Project Name'] = 'Total'

        title = f"So sánh giờ giữa các dự án trong năm {years[0]} (theo tháng)"
        return df_comparison, title

    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
        if len(selected_projects) != 1:
            return pd.DataFrame(), "Lỗi: Internal - Vui lòng chọn CHỈ MỘT dự án cho chế độ này."

        selected_project_name = selected_projects[0]

        if len(years) == 1 and len(months) > 0:
            df_comparison = df_filtered.groupby('MonthName')['Hours'].sum().reset_index()
            df_comparison.rename(columns={'Hours': f'Total Hours for {selected_project_name}'}, inplace=True)
            
            # Đảm bảo thứ tự tháng đúng cho biểu đồ, chỉ khi cột 'MonthName' tồn tại và không rỗng
            if not df_comparison.empty and 'MonthName' in df_comparison.columns:
                month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
                # Filter month_order to only include months actually present in the data to avoid errors if some months have no data
                present_months = [m for m in month_order if m in df_comparison['MonthName'].unique()]
                if present_months: # Only apply if there are months to categorize
                    df_comparison['MonthName'] = pd.Categorical(df_comparison['MonthName'], categories=present_months, ordered=True)
                    df_comparison = df_comparison.sort_values('MonthName').reset_index(drop=True)
            else:
                print("Cảnh báo: df_comparison trống hoặc không có cột MonthName khi cố gắng sắp xếp theo tháng.")
                return pd.DataFrame(), f"Không có dữ liệu tháng nào cho dự án '{selected_project_name}' trong năm {years[0]}."

            df_comparison['Project Name'] = selected_project_name
            title = f"Tổng giờ dự án {selected_project_name} qua các tháng trong năm {years[0]}"
            return df_comparison, title

        elif len(years) > 1 and not months:
            df_comparison = df_filtered.groupby('Year')['Hours'].sum().reset_index()
            df_comparison.rename(columns={'Hours': f'Total Hours for {selected_project_name}'}, inplace=True)
            df_comparison['Year'] = df_comparison['Year'].astype(str)
            
            df_comparison['Project Name'] = selected_project_name
            title = f"Tổng giờ dự án {selected_project_name} qua các năm"
            return df_comparison, title

        else:
            return pd.DataFrame(), "Cấu hình so sánh dự án qua thời gian không hợp lệ. Vui lòng chọn một năm với nhiều tháng, HOẶC nhiều năm."
        
    return pd.DataFrame(), "Chế độ so sánh không hợp lệ."

def create_comparison_chart(df_comparison, comparison_mode, chart_title):
    """
    Tạo biểu đồ so sánh dựa trên chế độ so sánh.
    """
    if df_comparison.empty:
        fig = go.Figure()
        fig.add_annotation(text="Không có dữ liệu để hiển thị biểu đồ so sánh.",
                           xref="paper", yref="paper", showarrow=False,
                           font=dict(size=16, color="grey"))
        fig.update_layout(height=400)
        return fig

    if comparison_mode in ["So Sánh Dự Án Trong Một Tháng", "Compare Projects in a Month"]:
        fig = px.bar(df_comparison, x='Project name', y='Total Hours', 
                     title=chart_title,
                     labels={'Project name': 'Tên dự án', 'Total Hours': 'Tổng giờ'},
                     color_discrete_sequence=px.colors.qualitative.Pastel)
        fig.update_layout(xaxis_title="Tên dự án", yaxis_title="Tổng giờ")

    elif comparison_mode in ["So Sánh Dự Án Trong Một Năm", "Compare Projects in a Year"]:
        # Loại bỏ hàng 'Total' nếu có để tránh lỗi với plotly.express
        df_plot = df_comparison[df_comparison['Project Name'] != 'Total'].copy()
        
        # Melt the DataFrame to long format for Plotly Express
        month_columns = [col for col in df_plot.columns if col in MONTH_NAMES_EN]
        df_melted = df_plot.melt(id_vars=['Project Name'], value_vars=month_columns, 
                                 var_name='MonthName', value_name='Hours')
        
        # Ensure correct month order
        month_order = [m for m in MONTH_NAMES_EN if m in df_melted['MonthName'].unique()]
        df_melted['MonthName'] = pd.Categorical(df_melted['MonthName'], categories=month_order, ordered=True)
        df_melted = df_melted.sort_values('MonthName')

        fig = px.bar(df_melted, x='MonthName', y='Hours', color='Project Name', 
                     barmode='group',
                     title=chart_title,
                     labels={'MonthName': 'Tháng', 'Hours': 'Tổng giờ', 'Project Name': 'Tên dự án'},
                     color_discrete_sequence=px.colors.qualitative.Pastel)
        fig.update_layout(xaxis_title="Tháng", yaxis_title="Tổng giờ")

    elif comparison_mode in ["So Sánh Một Dự Án Qua Các Tháng/Năm", "Compare One Project Over Time (Months/Years)"]:
        if 'MonthName' in df_comparison.columns:
            # So sánh qua các tháng trong một năm
            fig = px.line(df_comparison, x='MonthName', y=df_comparison.columns[1], # Lấy cột giờ tổng dựa vào index
                          title=chart_title,
                          labels={'MonthName': 'Tháng', df_comparison.columns[1]: 'Tổng giờ'},
                          markers=True)
            fig.update_layout(xaxis_title="Tháng", yaxis_title="Tổng giờ")
        elif 'Year' in df_comparison.columns:
            # So sánh qua các năm
            fig = px.line(df_comparison, x='Year', y=df_comparison.columns[1], # Lấy cột giờ tổng dựa vào index
                          title=chart_title,
                          labels={'Year': 'Năm', df_comparison.columns[1]: 'Tổng giờ'},
                          markers=True)
            fig.update_layout(xaxis_title="Năm", yaxis_title="Tổng giờ")
        else:
            fig = go.Figure()
            fig.add_annotation(text="Không thể tạo biểu đồ cho chế độ so sánh này. Thiếu cột thời gian (Tháng hoặc Năm).",
                            xref="paper", yref="paper", showarrow=False,
                            font=dict(size=16, color="grey"))
            fig.update_layout(height=400)
    else:
        fig = go.Figure()
        fig.add_annotation(text="Chế độ so sánh không được hỗ trợ để tạo biểu đồ.",
                           xref="paper", yref="paper", showarrow=False,
                           font=dict(size=16, color="grey"))
        fig.update_layout(height=400)
    
    fig.update_layout(title_x=0.5, height=400, margin=dict(l=20, r=20, t=60, b=20))
    return fig

def create_comparison_table(df_comparison, comparison_mode):
    """
    Tạo bảng so sánh dựa trên chế độ so sánh.
    """
    if df_comparison.empty:
        return go.Figure().add_annotation(text="Không có dữ liệu để hiển thị bảng so sánh.",
                                           xref="paper", yref="paper", showarrow=False,
                                           font=dict(size=16, color="grey")).update_layout(height=100)

    # Format numeric columns
    for col in df_comparison.columns:
        if df_comparison[col].dtype in ['float64', 'int64'] and col not in ['Year']:
            df_comparison[col] = df_comparison[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
    
    fig = go.Figure(data=[go.Table(
        header=dict(values=list(df_comparison.columns),
                    fill_color='paleturquoise',
                    align='left'),
        cells=dict(values=[df_comparison[col] for col in df_comparison.columns],
                   fill_color='lavender',
                   align='left'))
    ])
    fig.update_layout(title_text="Dữ liệu bảng so sánh", title_x=0.5, height=300)
    return fig


def fig_to_image_bytes(fig):
    """Chuyển đổi Plotly figure sang bytes hình ảnh PNG."""
    img_bytes = fig.to_image(format="png", engine="kaleido")
    return io.BytesIO(img_bytes)

def export_pdf_report(overall_summary_fig, monthly_summary_fig, project_summary_fig, raw_data_fig,
                       comparison_chart_fig, comparison_table_fig,
                       year, selected_month_name, selected_project_name, 
                       comparison_mode_text, report_type):
    """
    Xuất báo cáo PDF từ các biểu đồ và bảng.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter))
    styles = getSampleStyleSheet()
    story = []

    # Title
    if report_type == "full_report":
        title_text = f"Báo cáo tổng quan dự án"
        if year:
            title_text += f" năm {year}"
        if selected_month_name:
            title_text += f" tháng {selected_month_name}"
        if selected_project_name:
            title_text += f" dự án {selected_project_name}"
    elif report_type == "comparison_report":
        title_text = f"Báo cáo so sánh dự án"
        title_text += f" ({comparison_mode_text})"

    story.append(Paragraph(f"<h1 align='center'>{title_text}</h1>", styles['h1']))
    story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph(f"<i>Ngày tạo báo cáo: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</i>", styles['Normal']))
    story.append(Spacer(1, 0.2 * inch))

    # Add Overall Summary Table
    story.append(Paragraph("<h3>1. Tổng quan chung</h3>", styles['h3']))
    overall_summary_img = fig_to_image_bytes(overall_summary_fig)
    story.append(Image(overall_summary_img, width=4*inch, height=1.5*inch))
    story.append(Spacer(1, 0.2 * inch))

    if report_type == "full_report":
        # Add Monthly Summary Chart
        story.append(Paragraph("<h3>2. Tổng quan theo tháng</h3>", styles['h3']))
        monthly_img = fig_to_image_bytes(monthly_summary_fig)
        story.append(Image(monthly_img, width=10*inch, height=4*inch))
        story.append(Spacer(1, 0.2 * inch))

        # Add Project Summary Chart
        story.append(Paragraph("<h3>3. Tổng quan theo dự án</h3>", styles['h3']))
        project_img = fig_to_image_bytes(project_summary_fig)
        story.append(Image(project_img, width=10*inch, height=4*inch))
        story.append(Spacer(1, 0.2 * inch))
    
        # Add Raw Data Table (optional, might be too big for PDF)
        # story.append(Paragraph("<h3>4. Dữ liệu thô đã lọc</h3>", styles['h3']))
        # raw_data_img = fig_to_image_bytes(raw_data_fig)
        # story.append(Image(raw_data_img, width=10*inch, height=4*inch))
        # story.append(Spacer(1, 0.2 * inch))
    
    elif report_type == "comparison_report":
        story.append(Paragraph("<h3>2. Biểu đồ so sánh</h3>", styles['h3']))
        comparison_chart_img = fig_to_image_bytes(comparison_chart_fig)
        story.append(Image(comparison_chart_img, width=10*inch, height=4*inch))
        story.append(Spacer(1, 0.2 * inch))

        story.append(Paragraph("<h3>3. Bảng dữ liệu so sánh</h3>", styles['h3']))
        comparison_table_img = fig_to_image_bytes(comparison_table_fig)
        story.append(Image(comparison_table_img, width=10*inch, height=3*inch)) # Adjust height as needed
        story.append(Spacer(1, 0.2 * inch))


    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()
