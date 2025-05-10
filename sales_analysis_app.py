import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import io
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet
import xlsxwriter
from sklearn.linear_model import LinearRegression
import numpy as np
import statsmodels.api as sm
import os
import plotly.io as pio

# Función auxiliar para generar botones de descarga y reset de gráficas
def add_graph_controls(fig, fig_name):
    col1, col2 = st.columns(2)
    with col1:
        # Botón de descarga como PNG
        img_bytes = pio.to_image(fig, format="png", scale=2)
        st.download_button(
            label="Descargar Gráfica (PNG)",
            data=img_bytes,
            file_name=f"{fig_name}.png",
            mime="image/png"
        )
    with col2:
        # Botón para restablecer zoom/vista
        if st.button("Restablecer Vista", key=f"reset_{fig_name}"):
            fig.update_layout(
                xaxis=dict(autorange=True),
                yaxis=dict(autorange=True)
            )
            st.rerun()

# Funciones auxiliares
def generate_pdf(data: pd.DataFrame, title: str, filename: str, _data_hash: str) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()

    try:
        logo = Image("app/data/logo.png", width=100, height=50)
        elements.append(logo)
    except Exception as e:
        elements.append(Paragraph("Logo no disponible", styles['Normal']))

    elements.append(Paragraph(title, styles['Title']))
    elements.append(Paragraph(" ", styles['Normal']))

    data_list = [data.columns.tolist()] + data.values.tolist()
    table = Table(data_list)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

def generate_excel(data: pd.DataFrame, sheet_name: str, _data_hash: str) -> io.BytesIO:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        data.to_excel(writer, sheet_name=sheet_name, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
        for col_num, value in enumerate(data.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
        worksheet.autofit()
    buffer.seek(0)
    return buffer

def load_data():
    def load_excel():
        try:
            df = pd.read_excel("app/data/Órdenes del punto de venta (pos.order).xlsx", engine='openpyxl')
            return df
        except Exception as e:
            st.error(f"Error al cargar los datos: {str(e)}")
            return pd.DataFrame()

    def validate_data(df):
        required_cols = ['Cliente/Código de barras', 'Cliente/Nombre', 'Centro de Costos Aseavna', 'Fecha', 'Número de recibo', 
                        'Cliente/Nombre principal', 'Precio total colaborador', 'Comision Aseavna', 'Cuentas por a cobrar aseavna', 
                        'Cuentas por a Cobrar Avna', 'Ventas Totales', 'Líneas de la orden', 'Líneas de la orden/Cantidad']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.error(f"Faltan las columnas: {', '.join(missing_cols)}")
            return False
        return True

    def map_columns(df):
        df_columns = {col.strip().lower(): col for col in df.columns}
        for expected_col, search_col in CONFIG['columns'].items():
            found_col = next((col for col_name, col in df_columns.items() if col_name == search_col.strip().lower()), None)
            df[expected_col] = df[found_col] if found_col else ('Desconocido' if 'Cliente' in expected_col or 'Líneas' in expected_col else 0)
        return df

    def calculate_total(df):
        # Calcular Total Final como suma de Cuentas por a cobrar aseavna y Cuentas por a Cobrar Avna
        df['Total Final'] = pd.to_numeric(df['Cuentas por a cobrar aseavna'], errors='coerce').fillna(0) + \
                           pd.to_numeric(df['Cuentas por a Cobrar Avna'], errors='coerce').fillna(0)
        return df

    def clean_data(df):
        defaults = {
            'Cliente/Código de barras': 'Desconocido',
            'Cliente/Nombre': 'Desconocido',
            'Centro de Costos Aseavna': 'Desconocido',
            'Cliente/Nombre principal': 'Desconocido',
            'Líneas de la orden': 'Desconocido'
        }
        for col, default in defaults.items():
            df[col] = df[col].fillna(default)
        numeric_cols = ['Líneas de la orden/Cantidad', 'Total Final', 'Comision Aseavna', 'Cuentas por a cobrar aseavna', 
                       'Cuentas por a Cobrar Avna', 'Precio total colaborador']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        if 'Cliente/Nombre' in df.columns:
            df['Cliente/Nombre'] = df['Cliente/Nombre'].astype(str).str.strip().str.lower()
        if 'Centro de Costos Aseavna' in df.columns:
            df['Centro de Costos Aseavna'] = df['Centro de Costos Aseavna'].astype(str).str.strip().str.lower()
        return df

    def add_day_of_week(df):
        original_rows = len(df)
        st.sidebar.write(f"Filas cargadas inicialmente del archivo Excel: {original_rows}")

        if pd.api.types.is_numeric_dtype(df['Fecha']):
            df['Fecha'] = pd.to_datetime(df['Fecha'], unit='D', origin='1899-12-30', errors='coerce') - timedelta(days=2)
        else:
            df['Fecha'] = pd.to_datetime(df['Fecha'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
        
        df['Fecha_Valida'] = df['Fecha'].notna()
        invalid_dates = df['Fecha'].isna().sum()
        if invalid_dates > 0:
            st.warning(f"Se encontraron {invalid_dates} fechas no válidas que se excluirán del análisis.")
        
        duplicates = df.duplicated().sum()
        if duplicates > 0:
            st.warning(f"Se encontraron {duplicates} filas duplicadas en el archivo Excel. Se eliminarán.")
            df = df.drop_duplicates()
            st.sidebar.write(f"Filas después de eliminar duplicados: {len(df)}")
        
        df = df.dropna(subset=['Fecha'])
        st.sidebar.write(f"Filas después de eliminar fechas no válidas: {len(df)}")

        df['Día de la Semana'] = df['Fecha'].dt.day_name()
        day_translation = {
            'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
            'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
        }
        df['Día de la Semana'] = df['Día de la Semana'].map(day_translation).fillna(df['Día de la Semana'])
        return df

    df = load_excel()
    if df.empty or not validate_data(df):
        return pd.DataFrame()
    df = map_columns(df)
    df = calculate_total(df)
    df = clean_data(df)
    df = add_day_of_week(df)
    return df

# Configuración centralizada
CONFIG = {
    'columns': {
        'Cliente/Código de barras': 'Cliente/Código de barras',
        'Cliente/Nombre': 'Cliente/Nombre',
        'Centro de Costos Aseavna': 'Centro de Costos Aseavna',
        'Fecha': 'Fecha',
        'Número de recibo': 'Número de recibo',
        'Cliente/Nombre principal': 'Cliente/Nombre principal',
        'Precio total colaborador': 'Precio total colaborador',
        'Comision': 'Comision Aseavna',
        'Cuentas por a cobrar aseavna': 'Cuentas por a cobrar aseavna',
        'Cuentas por a Cobrar Avna': 'Cuentas por a Cobrar Avna',
        'Líneas de la orden': 'Líneas de la orden',
        'Líneas de la orden/Cantidad': 'Líneas de la orden/Cantidad'
    },
    'styles': {
        'metric_box': 'border: 1px solid #d3d3d3; padding: 10px; border-radius: 5px; background-color: white; margin: 5px auto; text-align: center; width: 90%; display: flex; flex-direction: column; justify-content: center; align-items: center;',
        'alert_box': 'background-color: #ff4d4d; padding: 10px; border-radius: 5px; margin: 10px auto; color: white; text-align: center; width: 90%;'
    },
    'colors': {
        'primary': '#4CAF50',
        'secondary': '#2c3e50',
        'warning': '#ffeb3b'
    }
}

# Soporte multi-idioma
TRANSLATIONS = {
    'es': {
        'title': '📊 Dashboard de Análisis de Ventas - ASEAVNA',
        'description': 'Análisis de órdenes de venta del sistema POS.',
        'filters_header': 'Filtros de Análisis',
        'date_range': 'Rango de Fechas',
        'select_period': 'Seleccionar Período',
        'product_type': 'Tipo de Producto',
        'client_group': 'Grupo de Clientes',
        'day_of_week': 'Día de la Semana',
        'specific_client': 'Cliente Específico',
        'reset_filters': 'Restablecer Filtros',
        'metrics': 'Métricas Generales',
        'duplicates': 'Almuerzos Duplicados',
        'client_sales': 'Ventas por Cliente',
        'predictive': 'Análisis Predictivo',
        'visualizations': 'Visualizaciones',
        'export': 'Exportar Resumen',
        'raw_data': 'Datos Crudos',
        'no_data': 'No se encontraron datos. Asegúrese de que el archivo "Órdenes del punto de venta (pos.order).xlsx" esté disponible en app/data/.',
        'metrics_summary': 'Resumen de Métricas Principales',
        'orders': 'Órdenes Totales',
        'lines': 'Líneas Totales',
        'commission': 'Comisión Total',
        'accounts_aseavna': 'Ctas. por Cobrar Aseavna',
        'accounts_avna': 'Ctas. por Cobrar Avna',
        'top_product': 'Producto Más Vendido',
        'unique_clients': 'Clientes Únicos',
        'daily_sales': 'Resumen de Ingresos Diarios',
        'duplicates_detected': '⚠️ Se detectaron almuerzos ejecutivos duplicados:',
        'no_duplicates': '✅ No se detectaron almuerzos ejecutivos duplicados en el mismo día.',
        'download_excel': 'Descargar Duplicados (Excel)',
        'download_pdf': 'Descargar Duplicados (PDF)',
        'unusual_sales': '⚠️ Clientes con volumen de consumo inusual:',
        'export_client_sales': 'Exportar Reporte de Consumo por Cliente',
        'download_csv': 'Descargar CSV',
        'download_excel_client': 'Descargar Excel',
        'download_pdf_client': 'Descargar PDF',
        'predictive_subheader': 'Predicción de Ingresos Totales para los Próximos 7 Días',
        'growth_subheader': 'Productos con Potencial de Crecimiento (Basado en Ingresos)',
        'no_predictive_data': 'No hay suficientes datos históricos para predicción (se requieren al menos 2 días).',
        'no_monthly_data': 'No hay suficientes datos mensuales para calcular el crecimiento de productos (se requieren al menos dos meses).',
        'predictive_error': 'Error en el análisis predictivo: {error}',
        'top_products': 'Top 10 Productos por Ingresos',
        '/yyyy-MM-ddaily_trend': 'Tendencia Diaria de Ingresos',
        'sales_by_group': 'Ingresos por Grupo de Clientes',
        'export_summary': 'Exportar Resumen de Métricas',
        'download_summary_csv': 'Descargar Resumen (CSV)',
        'download_summary_excel': 'Descargar Resumen (Excel)',
        'download_summary_pdf': 'Descargar Resumen (PDF)',
        'show_raw_data': 'Mostrar Datos Crudos',
        'footer': 'Desarrollado por Wilfredos para ASEAVNA | Fuente de Datos: Órdenes del Punto de Venta (POS) | 2025'
    },
    'en': {
        'title': '📊 Sales Analysis Dashboard - ASEAVNA',
        'description': 'Advanced analysis of POS sales orders, with metrics, predictions, and downloadable client reports.',
        'filters_header': 'Analysis Filters',
        'date_range': 'Date Range',
        'select_period': 'Select Period',
        'product_type': 'Product Type',
        'client_group': 'Client Group',
        'day_of_week': 'Day of the Week',
        'specific_client': 'Specific Client',
        'reset_filters': 'Reset Filters',
        'metrics': 'General Metrics',
        'duplicates': 'Duplicate Lunches',
        'client_sales': 'Sales by Client',
        'predictive': 'Predictive Analysis',
        'visualizations': 'Visualizations',
        'export': 'Export Summary',
        'raw_data': 'Raw Data',
        'no_data': 'No data found. Ensure the file "Órdenes del punto de venta (pos.order).xlsx" is available in app/data/.',
        'metrics_summary': 'Key Metrics Summary',
        'orders': 'Total Orders',
        'lines': 'Total Lines',
        'commission': 'Total Commission',
        'accounts_aseavna': 'Accounts Receivable Aseavna',
        'accounts_avna': 'Accounts Receivable Avna',
        'top_product': 'Top Selling Product',
        'unique_clients': 'Unique Clients',
        'daily_sales': 'Daily Revenue Summary',
        'duplicates_detected': '⚠️ Duplicate executive lunches detected:',
        'no_duplicates': '✅ No duplicate executive lunches detected on the same day.',
        'download_excel': 'Download Duplicates (Excel)',
        'download_pdf': 'Download Duplicates (PDF)',
        'unusual_sales': '⚠️ Clients with unusual consumption volume:',
        'export_client_sales': 'Export Client Consumption Report',
        'download_csv': 'Download CSV',
        'download_excel_client': 'Download Excel',
        'download_pdf_client': 'Download PDF',
        'predictive_subheader': 'Total Revenue Forecast for the Next 7 Days',
        'growth_subheader': 'Products with Growth Potential (Based on Revenue)',
        'no_predictive_data': 'Not enough historical data for prediction (at least 2 days required).',
        'no_monthly_data': 'Not enough monthly data to calculate product growth (at least two months required).',
        'predictive_error': 'Error in predictive analysis: {error}',
        'top_products': 'Top 10 Products by Revenue',
        'daily_trend': 'Daily Revenue Trend',
        'sales_by_group': 'Revenue by Client Group',
        'export_summary': 'Export Metrics Summary',
        'download_summary_csv': 'Download Summary (CSV)',
        'download_summary_excel': 'Download Summary (Excel)',
        'download_summary_pdf': 'Download Summary (PDF)',
        'show_raw_data': 'Show Raw Data',
        'footer': 'Developed by Wilfredos for ASEAVNA | Data Source: Point of Sale (POS) Orders | 2025'
    }
}

# Configuración de la página
st.set_page_config(
    page_title="Análisis de Ventas - ASEAVNA",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo personalizado
st.markdown(f"""
<style>
.main {{background-color: #f5f7fa; padding: 10px;}}
.stButton>button {{background-color: {CONFIG['colors']['primary']}; color: white; border-radius: 5px;}}
.stSidebar {{background-color: #e8ecef; padding: 5px;}}
h1, h2, h3 {{color: {CONFIG['colors']['secondary']}; text-align: center;}}
.metric-box {{border: 1px solid #d3d3d3; padding: 10px; border-radius: 5px; background-color: white; margin: 5px auto; text-align: center; width: 90%; display: flex; flex-direction: column; justify-content: center; align-items: center;}}
.metric-box .title {{font-size: 10px; color: {CONFIG['colors']['primary']}; margin-bottom: 2px;}}
.metric-box .value {{font-size: 14px; color: {CONFIG['colors']['secondary']}; font-weight: bold;}}
.alert-box {{background-color: #ff4d4d; padding: 10px; border-radius: 5px; margin: 10px auto; color: white; text-align: center; width: 90%;}}
.logo-container {{text-align: center; margin: 10px 0;}}
/* Ajustes para dispositivos móviles */
@media (max-width: 600px) {{
    .metric-box {{width: 95%; padding: 8px;}}
    .metric-box .title {{font-size: 9px;}}
    .metric-box .value {{font-size: 12px;}}
    .plotly-graph-div {{width: 100% !important; height: auto !important;}}
    .modebar {{display: block !important;}}
    .stButton>button {{font-size: 12px; padding: 8px;}}
}}
</style>
""", unsafe_allow_html=True)

# Selección de idioma
language = st.sidebar.selectbox("Idioma / Language", ["Español", "English"])
lang_code = 'es' if language == "Español" else 'en'

# Mostrar el logo
st.markdown('<div class="logo-container">', unsafe_allow_html=True)
try:
    if os.path.exists("app/data/logo.png"):
        st.image("app/data/logo.png", use_container_width=False, width=200)
    else:
        st.warning("El archivo 'logo.png' no se encuentra en app/data/. Por favor, asegúrese de que el archivo esté en la ruta correcta.")
except Exception as e:
    st.warning(f"No se pudo cargar el logo debido a un error: {str(e)}. Asegúrese de que 'logo.png' esté en app/data/.")
st.markdown('</div>', unsafe_allow_html=True)

# Título y descripción
st.title(TRANSLATIONS[lang_code]['title'])
st.markdown(TRANSLATIONS[lang_code]['description'], unsafe_allow_html=True)

# Carga de datos
df = load_data()

if df.empty:
    st.warning(TRANSLATIONS[lang_code]['no_data'])
else:
    # Sidebar: filtros
    st.sidebar.header(TRANSLATIONS[lang_code]['filters_header'])
    with st.sidebar.expander(TRANSLATIONS[lang_code]['date_range'], expanded=True):
        date_option = st.selectbox(
            TRANSLATIONS[lang_code]['select_period'],
            ["Personalizado", "Última Semana", "Último Mes", "Todo el Período"],
            key="date_option"
        )
        if date_option == "Última Semana":
            end_date = df['Fecha'].max().date()
            start_date = end_date - timedelta(days=7)
        elif date_option == "Último Mes":
            end_date = df['Fecha'].max().date()
            start_date = end_date - timedelta(days=30)
        elif date_option == "Todo el Período":
            start_date = df['Fecha'].min().date()
            end_date = df['Fecha'].max().date()
        else:
            start_date = df['Fecha'].min().date()
            end_date = df['Fecha'].max().date()

        date_range = st.date_input(
            TRANSLATIONS[lang_code]['date_range'],
            [start_date, end_date],
            min_value=df['Fecha'].min().date(),
            max_value=df['Fecha'].max().date(),
            key="date_range"
        )

    with st.sidebar.expander("Filtros de Categorías"):
        product_types = ['Todos'] + sorted(df['Líneas de la orden'].dropna().astype(str).unique().tolist())
        selected_product = st.selectbox(TRANSLATIONS[lang_code]['product_type'], product_types, key="product")
        
        client_groups = ['Todos'] + sorted(df['Cliente/Nombre principal'].dropna().astype(str).unique().tolist())
        selected_client_grp = st.selectbox(TRANSLATIONS[lang_code]['client_group'], client_groups, key="client_group")
        
        days_of_week = ['Todos'] + sorted(df['Día de la Semana'].dropna().astype(str).unique().tolist())
        selected_day = st.selectbox(TRANSLATIONS[lang_code]['day_of_week'], days_of_week, key="day")
        
        clients = ['Todos'] + sorted(df['Cliente/Nombre'].dropna().astype(str).unique().tolist())
        selected_client = st.selectbox(TRANSLATIONS[lang_code]['specific_client'], clients, key="client")
        
        centros_costos = ['Todos'] + sorted(df['Centro de Costos Aseavna'].dropna().astype(str).unique().tolist())
        selected_centro = st.selectbox("Centro de Costos", centros_costos, key="centro_costos")

    if st.sidebar.button(TRANSLATIONS[lang_code]['reset_filters']):
        st.rerun()

    # Aplicar filtros
    filtered_df = df.copy()
    total_lines = len(filtered_df)
    if len(date_range) == 2:
        sd, ed = date_range
        sd = pd.to_datetime(sd)
        ed = pd.to_datetime(ed) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        filtered_df = filtered_df[
            (filtered_df['Fecha'] >= sd) &
            (filtered_df['Fecha'] <= ed)
        ]
        st.sidebar.write(f"Filas totales antes del filtro: {total_lines}")
        st.sidebar.write(f"Filas después de filtrar por fechas ({sd.date()} a {ed.date()}): {len(filtered_df)}")
        st.sidebar.write(f"Órdenes únicas después del filtro: {filtered_df['Número de recibo'].nunique()}")
    else:
        st.warning("Por favor, selecciona un rango de fechas válido.")

    if selected_product != 'Todos':
        filtered_df = filtered_df[filtered_df['Líneas de la orden'] == selected_product]
    if selected_client_grp != 'Todos':
        filtered_df = filtered_df[filtered_df['Cliente/Nombre principal'] == selected_client_grp]
    if selected_day != 'Todos':
        filtered_df = filtered_df[filtered_df['Día de la Semana'] == selected_day]
    if selected_client != 'Todos':
        selected_client_normalized = selected_client.strip().lower()
        filtered_df = filtered_df[filtered_df['Cliente/Nombre'] == selected_client_normalized]
        st.sidebar.write(f"Filas después de filtrar por cliente '{selected_client}': {len(filtered_df)}")
    if selected_centro != 'Todos':
        selected_centro_normalized = selected_centro.strip().lower()
        filtered_df = filtered_df[filtered_df['Centro de Costos Aseavna'] == selected_centro_normalized]
        st.sidebar.write(f"Filas después de filtrar por centro de costos '{selected_centro}': {len(filtered_df)}")

    product_types = ['Todos'] + sorted(filtered_df['Líneas de la orden'].dropna().astype(str).unique().tolist())
    client_groups = ['Todos'] + sorted(filtered_df['Cliente/Nombre principal'].dropna().astype(str).unique().tolist())
    days_of_week = ['Todos'] + sorted(filtered_df['Día de la Semana'].dropna().astype(str).unique().tolist())
    clients = ['Todos'] + sorted(filtered_df['Cliente/Nombre'].dropna().astype(str).unique().tolist())
    centros_costos = ['Todos'] + sorted(filtered_df['Centro de Costos Aseavna'].dropna().astype(str).unique().tolist())

    # Panel de métricas principales
    st.subheader(TRANSLATIONS[lang_code]['metrics_summary'])
    col1, col2, col3, col4, col5 = st.columns([1, 1, 1, 1, 1])
    total_orders = filtered_df['Número de recibo'].nunique()
    total_lines_filtered = len(filtered_df)
    total_commission = filtered_df['Comision Aseavna'].sum()
    total_cuentas_cobrar_aseavna = filtered_df['Cuentas por a cobrar aseavna'].sum()
    total_cuentas_cobrar_avna = filtered_df['Cuentas por a Cobrar Avna'].sum()

    with col1:
        st.markdown(f'<div class="metric-box"><span class="title">{TRANSLATIONS[lang_code]["orders"]}</span><span class="value">{total_orders:,}</span></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="metric-box"><span class="title">{TRANSLATIONS[lang_code]["lines"]}</span><span class="value">{total_lines_filtered:,}</span></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="metric-box"><span class="title">{TRANSLATIONS[lang_code]["commission"]}</span><span class="value">₡{total_commission:,.2f}</span></div>', unsafe_allow_html=True)
    with col4:
        st.markdown(f'<div class="metric-box"><span class="title">{TRANSLATIONS[lang_code]["accounts_aseavna"]}</span><span class="value">₡{total_cuentas_cobrar_aseavna:,.2f}</span></div>', unsafe_allow_html=True)
    with col5:
        st.markdown(f'<div class="metric-box"><span class="title">{TRANSLATIONS[lang_code]["accounts_avna"]}</span><span class="value">₡{total_cuentas_cobrar_avna:,.2f}</span></div>', unsafe_allow_html=True)

    # Crear pestañas
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        TRANSLATIONS[lang_code]['metrics'],
        TRANSLATIONS[lang_code]['duplicates'],
        TRANSLATIONS[lang_code]['client_sales'],
        TRANSLATIONS[lang_code]['predictive'],
        TRANSLATIONS[lang_code]['visualizations'],
        TRANSLATIONS[lang_code]['export'],
        TRANSLATIONS[lang_code]['raw_data']
    ])

    # Tab 1: Métricas Generales
    with tab1:
        st.header(TRANSLATIONS[lang_code]['metrics'])
        most_sold = filtered_df.groupby('Líneas de la orden')['Total Final'].sum().idxmax() if not filtered_df.empty else "N/A"
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f'<div class="metric-box"><span class="title">{TRANSLATIONS[lang_code]["top_product"]}</span><span class="value">{most_sold}</span></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-box"><span class="title">{TRANSLATIONS[lang_code]["unique_clients"]}</span><span class="value">{len(filtered_df["Cliente/Nombre"].unique())}</span></div>', unsafe_allow_html=True)
        
        daily_summary = filtered_df.groupby(filtered_df['Fecha'].dt.date)['Total Final'].sum().reset_index()
        if not daily_summary.empty:
            fig_summary = px.line(
                daily_summary, x='Fecha', y='Total Final',
                labels={'Total Final': 'Ingresos (₡)', 'Fecha': 'Fecha'},
                title=TRANSLATIONS[lang_code]['daily_sales'],
                template="plotly_white",
                color_discrete_sequence=["#4CAF50"]
            )
            fig_summary.update_layout(
                margin=dict(l=20, r=20, t=60, b=20),
                xaxis_title_font_size=14,
                yaxis_title_font_size=14,
                title_x=0.5,
                showlegend=True,
                xaxis=dict(tickformat="%Y-%m-%d", gridcolor='lightgray'),
                yaxis=dict(gridcolor='lightgray'),
                dragmode='zoom',  # Habilitar zoom
                modebar=dict(
                    bgcolor='rgba(0,0,0,0)',
                    color='rgba(0,0,0,0.5)',
                    activecolor=CONFIG['colors']['primary']
                )
            )
            fig_summary.update_xaxes(
                rangeslider_visible=True,  # Agregar control deslizante para zoom
                rangeselector=dict(
                    buttons=list([
                        dict(count=7, label="1w", step="day", stepmode="backward"),
                        dict(count=1, label="1m", step="month", stepmode="backward"),
                        dict(step="all", label="Todo")
                    ])
                )
            )
            st.plotly_chart(fig_summary, use_container_width=True)
            add_graph_controls(fig_summary, "daily_sales")
        else:
            st.warning("No hay datos suficientes para mostrar la tendencia diaria.")

    # Tab 2: Verificación de Almuerzos Ejecutivos Duplicados
    with tab2:
        st.header(TRANSLATIONS[lang_code]['duplicates'])
        lunch_df = filtered_df[filtered_df['Líneas de la orden'] == 'Almuerzo Ejecutivo Aseavna'].copy()
        lunch_df['Fecha_Dia'] = lunch_df['Fecha'].dt.date
        dup = lunch_df.groupby(['Cliente/Nombre', 'Fecha_Dia']).filter(lambda x: len(x) > 1)
        
        if not dup.empty:
            st.markdown(f'<div class="alert-box">{TRANSLATIONS[lang_code]["duplicates_detected"]}</div>', unsafe_allow_html=True)
            st.balloons()
            summary = dup.groupby(['Cliente/Nombre', 'Fecha_Dia']).size().reset_index(name='Cantidad')
            st.dataframe(summary)
            st.subheader("Detalles de Duplicados")
            st.dataframe(dup[['Cliente/Nombre', 'Fecha', 'Número de recibo', 'Líneas de la orden']])
            c1, c2 = st.columns(2)
            with c1:
                buf_xl = generate_excel(dup, "Duplicados", dup.to_string())
                st.download_button(
                    TRANSLATIONS[lang_code]['download_excel'],
                    data=buf_xl,
                    file_name="almuerzos_duplicados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with c2:
                buf_pdf = generate_pdf(dup, "Reporte de Almuerzos Duplicados", "almuerzos_duplicados.pdf", dup.to_string())
                st.download_button(
                    TRANSLATIONS[lang_code]['download_pdf'],
                    data=buf_pdf,
                    file_name="almuerzos_duplicados.pdf",
                    mime="application/pdf"
                )
        else:
            st.success(TRANSLATIONS[lang_code]['no_duplicates'])

    # Tab 3: Análisis de Consumo por Cliente
    with tab3:
        st.header(TRANSLATIONS[lang_code]['client_sales'])
        client_sales = filtered_df.groupby('Cliente/Nombre').agg({
            'Total Final': 'sum',
            'Número de recibo': 'nunique',
            'Comision Aseavna': 'sum',
            'Cuentas por a cobrar aseavna': 'sum',
            'Cuentas por a Cobrar Avna': 'sum',
            'Líneas de la orden': lambda x: x.mode()[0] if not x.empty else 'N/A'
        }).reset_index()
        client_sales.columns = [
            'Cliente',
            'Ingresos Totales (₡)',
            'Número de Órdenes',
            'Comisión Total (₡)',
            'Ctas. por Cobrar Aseavna (₡)',
            'Ctas. por Cobrar Avna (₡)',
            'Producto Más Comprado'
        ]
        
        if not client_sales.empty and client_sales['Ingresos Totales (₡)'].sum() > 0:
            threshold = client_sales['Ingresos Totales (₡)'].quantile(0.95)
            unusual = client_sales[client_sales['Ingresos Totales (₡)'] > threshold]
            if not unusual.empty:
                st.markdown(
                    f'<div class="alert-box" style="background-color: {CONFIG["colors"]["warning"]}; color: black;">'
                    f'{TRANSLATIONS[lang_code]["unusual_sales"]} (Ingresos Totales > ₡{threshold:,.2f})'
                    f'</div>',
                    unsafe_allow_html=True
                )
                unusual_display = unusual[['Cliente', 'Ingresos Totales (₡)', 'Número de Órdenes']].copy()
                unusual_display['Ingresos Totales (₡)'] = unusual_display['Ingresos Totales (₡)'].apply(lambda x: f"₡{x:,.2f}")
                st.dataframe(unusual_display)
            else:
                st.info("No se encontraron clientes con volumen de ingresos inusual.")
        else:
            st.warning("No hay datos suficientes para identificar clientes con ingresos inusuales.")
        
        client_sales_display = client_sales.copy()
        client_sales_display['Ingresos Totales (₡)'] = client_sales_display['Ingresos Totales (₡)'].apply(lambda x: f"₡{x:,.2f}")
        client_sales_display['Comisión Total (₡)'] = client_sales_display['Comisión Total (₡)'].apply(lambda x: f"₡{x:,.2f}")
        client_sales_display['Ctas. por Cobrar Aseavna (₡)'] = client_sales_display['Ctas. por Cobrar Aseavna (₡)'].apply(lambda x: f"₡{x:,.2f}")
        client_sales_display['Ctas. por Cobrar Avna (₡)'] = client_sales_display['Ctas. por Cobrar Avna (₡)'].apply(lambda x: f"₡{x:,.2f}")
        st.dataframe(client_sales_display)
        
        st.subheader(TRANSLATIONS[lang_code]['export_client_sales'])
        c1, c2, c3 = st.columns(3)
        with c1:
            csv_bytes = client_sales.to_csv(index=False).encode('utf-8')
            st.download_button(
                TRANSLATIONS[lang_code]['download_csv'],
                data=csv_bytes,
                file_name="ingresos_por_cliente.csv",
                mime="text/csv"
            )
        with c2:
            buf_xl2 = generate_excel(client_sales, "Ingresos por Cliente", client_sales.to_string())
            st.download_button(
                TRANSLATIONS[lang_code]['download_excel_client'],
                data=buf_xl2,
                file_name="ingresos_por_cliente.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c3:
            buf_pdf2 = generate_pdf(client_sales, "Reporte de Ingresos por Cliente - ASEAVNA", "ingresos_por_cliente.pdf", client_sales.to_string())
            st.download_button(
                TRANSLATIONS[lang_code]['download_pdf_client'],
                data=buf_pdf2,
                file_name="ingresos_por_cliente.pdf",
                mime="application/pdf"
            )

    # Tab 4: Análisis Predictivo
    with tab4:
        st.header(TRANSLATIONS[lang_code]['predictive'])
        try:
            # Agrupar por fecha para obtener los ingresos totales diarios (Total Final)
            daily = filtered_df.groupby(filtered_df['Fecha'].dt.date)['Total Final'].sum().reset_index(name='Total')
            daily['Days'] = (pd.to_datetime(daily['Fecha']) - pd.to_datetime(daily['Fecha'].min())).dt.days
            
            # Validar que haya suficientes datos para la predicción
            if len(daily) < 2:
                st.warning(TRANSLATIONS[lang_code]['no_predictive_data'])
            else:
                # Asegurar que no haya valores nulos en 'Total'
                daily = daily.dropna(subset=['Total'])
                if daily.empty:
                    st.warning("No hay datos válidos para realizar la predicción.")
                else:
                    # Modelo de regresión lineal para predecir ingresos futuros
                    X = sm.add_constant(daily['Days'])
                    model = sm.OLS(daily['Total'], X).fit()
                    future_days = np.array([daily['Days'].iloc[-1] + i for i in range(1, 8)])
                    future_X = sm.add_constant(future_days)
                    preds = model.predict(future_X)
                    conf_int = model.get_prediction(future_X).conf_int()
                    
                    # Crear DataFrame de predicciones con intervalos de confianza
                    pred_df = pd.DataFrame({
                        'Fecha': [pd.to_datetime(daily['Fecha']).max() + timedelta(days=i) for i in range(1, 8)],
                        'Total': preds,
                        'Lower': np.maximum(conf_int[:, 0], 0),  # Evitar valores negativos
                        'Upper': np.maximum(conf_int[:, 1], 0),  # Evitar valores negativos
                        'Tipo': 'Predicción'
                    })
                    hist_df = pd.DataFrame({
                        'Fecha': pd.to_datetime(daily['Fecha']),
                        'Total': daily['Total'],
                        'Tipo': 'Histórico'
                    })
                    combined = pd.concat([hist_df, pred_df]).reset_index(drop=True)
                    
                    # Gráfica de predicción de ingresos
                    st.subheader(TRANSLATIONS[lang_code]['predictive_subheader'])
                    fig_pred = px.line(
                        combined, 
                        x='Fecha', 
                        y='Total', 
                        color='Tipo',
                        labels={'Total': 'Ingresos Totales (₡)', 'Fecha': 'Fecha'},
                        title="Tendencia Histórica y Predicción de Ingresos Totales con Intervalos de Confianza",
                        template="plotly_white",
                        color_discrete_sequence=["#4CAF50", "#FF5733"]
                    )
                    # Añadir intervalos de confianza
                    fig_pred.add_scatter(
                        x=pred_df['Fecha'], 
                        y=pred_df['Upper'], 
                        mode='lines', 
                        line=dict(dash='dash', color='gray'), 
                        name='Límite Superior',
                        showlegend=True
                    )
                    fig_pred.add_scatter(
                        x=pred_df['Fecha'], 
                        y=pred_df['Lower'], 
                        mode='lines', 
                        line=dict(dash='dash', color='gray'), 
                        name='Límite Inferior',
                        showlegend=True
                    )
                    # Personalizar el formato del eje Y para mostrar comas
                    fig_pred.update_layout(
                        margin=dict(l=20, r=20, t=60, b=20),
                        xaxis_title_font_size=14,
                        yaxis_title_font_size=14,
                        title_x=0.5,
                        yaxis=dict(
                            tickformat=",.0f",  # Formato con comas para miles
                            gridcolor='lightgray'
                        ),
                        xaxis=dict(
                            tickformat="%Y-%m-%d",
                            gridcolor='lightgray'
                        ),
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=-0.3,
                            xanchor="center",
                            x=0.5
                        ),
                        dragmode='zoom',  # Habilitar zoom
                        modebar=dict(
                            bgcolor='rgba(0,0,0,0)',
                            color='rgba(0,0,0,0.5)',
                            activecolor=CONFIG['colors']['primary']
                        )
                    )
                    fig_pred.update_xaxes(
                        rangeslider_visible=True,  # Agregar control deslizante para zoom
                        rangeselector=dict(
                            buttons=list([
                                dict(count=7, label="1w", step="day", stepmode="backward"),
                                dict(count=1, label="1m", step="month", stepmode="backward"),
                                dict(step="all", label="Todo")
                            ])
                        )
                    )
                    st.plotly_chart(fig_pred, use_container_width=True)
                    add_graph_controls(fig_pred, "predictive_trend")
                    
                    # Análisis de crecimiento de productos basado en ingresos (Total Final)
                    trends = filtered_df.groupby(['Líneas de la orden', filtered_df['Fecha'].dt.to_period('M')])['Total Final'].sum().unstack(fill_value=0)
                    if trends.shape[1] >= 2:
                        growth = ((trends.iloc[:, -1] - trends.iloc[:, -2]) / trends.iloc[:, -2].replace(0, np.nan) * 100).replace([np.inf, -np.inf], 0).dropna().sort_values(ascending=False)
                        top5 = growth.head(5).reset_index()
                        top5.columns = ['Producto', 'Crecimiento (%)']
                        st.subheader(TRANSLATIONS[lang_code]['growth_subheader'])
                        st.dataframe(top5)
                    else:
                        st.warning(TRANSLATIONS[lang_code]['no_monthly_data'])
        except Exception as e:
            st.error(TRANSLATIONS[lang_code]['predictive_error'].format(error=str(e)))

    # Tab 5: Visualizaciones Detalladas
    with tab5:
        st.header(TRANSLATIONS[lang_code]['visualizations'])
        viz_df = filtered_df.copy()
        if selected_centro != 'Todos':
            viz_df = viz_df[viz_df['Centro de Costos Aseavna'] == selected_centro_normalized]

        top10 = viz_df.groupby('Líneas de la orden')['Total Final'].sum().nlargest(10).reset_index()
        if not top10.empty and top10['Total Final'].sum() > 0:
            top10['Total Final'] = top10['Total Final'].clip(upper=1e7)
            fig1 = px.bar(
                top10, 
                x='Líneas de la orden', 
                y='Total Final',
                title=TRANSLATIONS[lang_code]['top_products'],
                labels={'Total Final': 'Ingresos (₡)', 'Líneas de la orden': 'Producto'},
                template="plotly_white",
                color_discrete_sequence=["#4CAF50"],
                hover_data={'Total Final': ':,.2f'}
            )
            fig1.update_layout(
                margin=dict(l=40, r=40, t=80, b=100),
                xaxis_tickangle=45,
                xaxis_title_font_size=14,
                yaxis_title_font_size=14,
                title_x=0.5,
                showlegend=False,
                xaxis=dict(tickmode='linear', gridcolor='lightgray'),
                yaxis=dict(gridcolor='lightgray'),
                dragmode='zoom',  # Habilitar zoom
                modebar=dict(
                    bgcolor='rgba(0,0,0,0)',
                    color='rgba(0,0,0,0.5)',
                    activecolor=CONFIG['colors']['primary']
                )
            )
            st.plotly_chart(fig1, use_container_width=True)
            add_graph_controls(fig1, "top_products")
        else:
            st.warning("No hay datos suficientes o válidos para mostrar los top 10 productos por ingresos.")

        daily_summary = viz_df.groupby(viz_df['Fecha'].dt.date)['Total Final'].sum().reset_index()
        if not daily_summary.empty and daily_summary['Total Final'].sum() > 0:
            fig2 = px.line(
                daily_summary, 
                x='Fecha', 
                y='Total Final',
                labels={'Total Final': 'Ingresos (₡)', 'Fecha': 'Fecha'},
                title=TRANSLATIONS[lang_code]['daily_trend'],
                template="plotly_white",
                color_discrete_sequence=["#4CAF50"],
                markers=True
            )
            fig2.update_layout(
                margin=dict(l=40, r=40, t=80, b=40),
                xaxis_title_font_size=14,
                yaxis_title_font_size=14,
                title_x=0.5,
                xaxis=dict(tickformat="%Y-%m-%d", gridcolor='lightgray'),
                yaxis=dict(gridcolor='lightgray'),
                dragmode='zoom',  # Habilitar zoom
                modebar=dict(
                    bgcolor='rgba(0,0,0,0)',
                    color='rgba(0,0,0,0.5)',
                    activecolor=CONFIG['colors']['primary']
                )
            )
            fig2.update_xaxes(
                rangeslider_visible=True,  # Agregar control deslizante para zoom
                rangeselector=dict(
                    buttons=list([
                        dict(count=7, label="1w", step="day", stepmode="backward"),
                        dict(count=1, label="1m", step="month", stepmode="backward"),
                        dict(step="all", label="Todo")
                    ])
                )
            )
            st.plotly_chart(fig2, use_container_width=True)
            add_graph_controls(fig2, "daily_trend")
        else:
            st.warning("No hay datos suficientes o válidos para mostrar la tendencia diaria de ingresos.")

        grp = viz_df.groupby('Cliente/Nombre principal')['Total Final'].sum().reset_index()
        if not grp.empty and grp['Total Final'].sum() > 0:
            # Limitar a los 10 grupos con mayores ingresos
            grp = grp.nlargest(10, 'Total Final')
            fig3 = px.pie(
                grp, 
                names='Cliente/Nombre principal', 
                values='Total Final',
                title=TRANSLATIONS[lang_code]['sales_by_group'],
                template="plotly_white",
                color_discrete_sequence=px.colors.sequential.Viridis
            )
            # Ajustar el espaciado y las etiquetas para evitar superposición
            fig3.update_traces(
                textinfo='percent+label',
                pull=[0.1 if i == 0 else 0 for i in range(len(grp))],  # Separar ligeramente la primera sección
                textposition='auto',  # Permitir que Plotly ajuste automáticamente la posición
                textfont=dict(size=10),  # Reducir tamaño de fuente para mejor ajuste
                insidetextorientation='radial'  # Asegurar que el texto no interfiera con el círculo
            )
            fig3.update_layout(
                margin=dict(l=40, r=150, t=80, b=40),  # Más espacio a la derecha para la leyenda
                title_x=0.5,
                legend=dict(
                    orientation="v",  # Leyenda vertical
                    x=1.1,  # Colocar a la derecha
                    y=0.5,
                    xanchor="left",
                    yanchor="middle",
                    font=dict(size=10)
                ),
                height=600,
                width=800,  # Reducir ancho para dejar espacio a la leyenda
                dragmode=False,  # Desactivar drag para evitar interacciones accidentales
                modebar=dict(
                    bgcolor='rgba(0,0,0,0)',
                    color='rgba(0,0,0,0.5)',
                    activecolor=CONFIG['colors']['primary']
                )
            )
            st.plotly_chart(fig3, use_container_width=True)
            add_graph_controls(fig3, "sales_by_group")
        else:
            st.warning("No hay datos suficientes o válidos para mostrar los ingresos por grupo de clientes.")

    # Tab 6: Resumen de Métricas para Exportar
    with tab6:
        st.header(TRANSLATIONS[lang_code]['export'])
        most_sold = filtered_df.groupby('Líneas de la orden')['Total Final'].sum().idxmax() if not filtered_df.empty else "N/A"
        least_sold = filtered_df.groupby('Líneas de la orden')['Total Final'].sum().idxmin() if not filtered_df.empty else "N/A"
        
        report = {
            "Número de Órdenes": total_orders,
            "Líneas Totales": total_lines_filtered,
            "Comisión Total (₡)": total_commission,
            "Ctas. por Cobrar Aseavna (₡)": total_cuentas_cobrar_aseavna,
            "Ctas. por Cobrar Avna (₡)": total_cuentas_cobrar_avna,
            "Clientes Únicos": len(clients) - 1,
            "Producto Más Vendido": most_sold,
            "Producto Menos Vendido": least_sold
        }
        report_df = pd.DataFrame([report])
        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button(
                TRANSLATIONS[lang_code]['download_summary_csv'],
                data=report_df.to_csv(index=False).encode('utf-8'),
                file_name="resumen_ventas_aseavna.csv",
                mime="text/csv"
            )
        with c2:
            buf_xl3 = generate_excel(report_df, "Resumen", report_df.to_string())
            st.download_button(
                TRANSLATIONS[lang_code]['download_summary_excel'],
                data=buf_xl3,
                file_name="resumen_ventas_aseavna.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c3:
            buf_pdf3 = generate_pdf(report_df, "Resumen de Ventas - ASEAVNA", "resumen_ventas_aseavna.pdf", report_df.to_string())
            st.download_button(
                TRANSLATIONS[lang_code]['download_summary_pdf'],
                data=buf_pdf3,
                file_name="resumen_ventas_aseavna.pdf",
                mime="application/pdf"
            )

    # Tab 7: Datos Crudos
    with tab7:
        st.header(TRANSLATIONS[lang_code]['raw_data'])
        if st.checkbox(TRANSLATIONS[lang_code]['show_raw_data']):
            st.dataframe(df.drop(columns=['Fecha_Valida'], errors='ignore'))

# Pie de página
st.markdown("---")
st.markdown(TRANSLATIONS[lang_code]['footer'])