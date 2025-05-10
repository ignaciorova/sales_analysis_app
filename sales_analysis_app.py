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
        required_cols = ['Fecha', 'Cliente/Nombre', 'Líneas de la orden']
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
        total_cols = ['Precio total colaborador', 'Comision Aseavna', 'Cuentas por a cobrar aseavna', 'Cuentas por a Cobrar Avna']
        df['Total'] = 0
        for col in total_cols:
            if col in df.columns:
                df['Total'] += pd.to_numeric(df[col], errors='coerce').fillna(0)
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
        numeric_cols = ['Líneas de la orden/Cantidad', 'Total', 'Comision', 'Cuentas por a cobrar aseavna', 'Cuentas por a Cobrar Avna', 'Precio total colaborador']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
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
            st.warning(f"Se encontraron {duplicates} filas duplicadas en el archivo Excel.")
        
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
        'description': 'Análisis avanzado de órdenes de venta del sistema POS, con métricas, predicciones y reportes descargables por cliente.',
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
        'daily_sales': 'Resumen de Ventas Diarias',
        'duplicates_detected': '⚠️ Se detectaron almuerzos ejecutivos duplicados:',
        'no_duplicates': '✅ No se detectaron almuerzos ejecutivos duplicados en el mismo día.',
        'download_excel': 'Descargar Duplicados (Excel)',
        'download_pdf': 'Descargar Duplicados (PDF)',
        'unusual_sales': '⚠️ Clientes con volumen de compras inusual:',
        'export_client_sales': 'Exportar Reporte de Ventas por Cliente',
        'download_csv': 'Descargar CSV',
        'download_excel_client': 'Descargar Excel',
        'download_pdf_client': 'Descargar PDF',
        'predictive_subheader': 'Predicción de Ventas para los Próximos 7 Días',
        'growth_subheader': 'Productos con Potencial de Crecimiento',
        'no_predictive_data': 'No hay suficientes datos históricos para predicción (se requieren al menos 2 días).',
        'no_monthly_data': 'No hay suficientes datos mensuales para calcular el crecimiento de productos (se requieren al menos dos meses).',
        'predictive_error': 'Error en el análisis predictivo: {error}',
        'top_products': 'Top 10 Productos por Ventas',
        'daily_trend': 'Tendencia Diaria de Ventas',
        'sales_by_group': 'Ventas por Grupo de Clientes',
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
        'daily_sales': 'Daily Sales Summary',
        'duplicates_detected': '⚠️ Duplicate executive lunches detected:',
        'no_duplicates': '✅ No duplicate executive lunches detected on the same day.',
        'download_excel': 'Download Duplicates (Excel)',
        'download_pdf': 'Download Duplicates (PDF)',
        'unusual_sales': '⚠️ Clients with unusual purchase volume:',
        'export_client_sales': 'Export Client Sales Report',
        'download_csv': 'Download CSV',
        'download_excel_client': 'Download Excel',
        'download_pdf_client': 'Download PDF',
        'predictive_subheader': 'Sales Forecast for the Next 7 Days',
        'growth_subheader': 'Products with Growth Potential',
        'no_predictive_data': 'Not enough historical data for prediction (at least 2 days required).',
        'no_monthly_data': 'Not enough monthly data to calculate product growth (at least two months required).',
        'predictive_error': 'Error in predictive analysis: {error}',
        'top_products': 'Top 10 Products by Sales',
        'daily_trend': 'Daily Sales Trend',
        'sales_by_group': 'Sales by Client Group',
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
.metric-box .stMetric > div:first-child {{font-size: 14px; color: {CONFIG['colors']['secondary']};}} /* Título */
.metric-box .stMetric > div:last-child {{font-size: 12px;}} /* Valor */
.alert-box {{background-color: #ff4d4d; padding: 10px; border-radius: 5px; margin: 10px auto; color: white; text-align: center; width: 90%;}}
.logo-container {{text-align: center; margin: 10px 0;}}
</style>
""", unsafe_allow_html=True)

# Selección de idioma
language = st.sidebar.selectbox("Idioma / Language", ["Español", "English"])
lang_code = 'es' if language == "Español" else 'en'

# Mostrar el logo
st.markdown('<div class="logo-container">', unsafe_allow_html=True)
try:
    import os
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
        filtered_df = filtered_df[filtered_df['Cliente/Nombre'] == selected_client]
    if selected_centro != 'Todos':
        filtered_df = filtered_df[filtered_df['Centro de Costos Aseavna'] == selected_centro]

    product_types = ['Todos'] + sorted(filtered_df['Líneas de la orden'].dropna().astype(str).unique().tolist())
    client_groups = ['Todos'] + sorted(filtered_df['Cliente/Nombre principal'].dropna().astype(str).unique().tolist())
    days_of_week = ['Todos'] + sorted(filtered_df['Día de la Semana'].dropna().astype(str).unique().tolist())
    clients = ['Todos'] + sorted(filtered_df['Cliente/Nombre'].dropna().astype(str).unique().tolist())
    centros_costos = ['Todos'] + sorted(filtered_df['Centro de Costos Aseavna'].dropna().astype(str).unique().tolist())

    # Panel de métricas principales
    st.subheader(TRANSLATIONS[lang_code]['metrics_summary'])
    col1, col2, col3, col4, col5 = st.columns([1, 1, 1, 1, 1])  # Espacio equitativo
    total_orders = filtered_df['Número de recibo'].nunique()
    total_lines_filtered = len(filtered_df)
    total_commission = filtered_df['Comision'].sum()
    total_cuentas_cobrar_aseavna = filtered_df['Cuentas por a cobrar aseavna'].sum()
    total_cuentas_cobrar_avna = filtered_df['Cuentas por a Cobrar Avna'].sum()

    with col1:
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.metric(TRANSLATIONS[lang_code]['orders'], f"{total_orders:,}")
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.metric(TRANSLATIONS[lang_code]['lines'], f"{total_lines_filtered:,}")
        st.markdown('</div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.metric(TRANSLATIONS[lang_code]['commission'], f"₡{total_commission:,.2f}")
        st.markdown('</div>', unsafe_allow_html=True)
    with col4:
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.metric(TRANSLATIONS[lang_code]['accounts_aseavna'], f"₡{total_cuentas_cobrar_aseavna:,.2f}")
        st.markdown('</div>', unsafe_allow_html=True)
    with col5:
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.metric(TRANSLATIONS[lang_code]['accounts_avna'], f"₡{total_cuentas_cobrar_avna:,.2f}")
        st.markdown('</div>', unsafe_allow_html=True)

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
        most_sold = filtered_df.groupby('Líneas de la orden')['Total'].sum().idxmax() if not filtered_df.empty else "N/A"
        col1, col2 = st.columns(2)
        with col1:
            st.markdown('<div class="metric-box">', unsafe_allow_html=True)
            st.metric(TRANSLATIONS[lang_code]['top_product'], most_sold)
            st.markdown('</div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div class="metric-box">', unsafe_allow_html=True)
            st.metric(TRANSLATIONS[lang_code]['unique_clients'], len(filtered_df['Cliente/Nombre'].unique()))
            st.markdown('</div>', unsafe_allow_html=True)
        
        daily_summary = filtered_df.groupby(filtered_df['Fecha'].dt.date)['Total'].sum().reset_index()
        if not daily_summary.empty:
            fig_summary = px.line(
                daily_summary, x='Fecha', y='Total',
                labels={'Total': 'Ventas (₡)', 'Fecha': 'Fecha'},
                title=TRANSLATIONS[lang_code]['daily_sales'],
                template="plotly_white",
                color_discrete_sequence=["#4CAF50"]
            )
            fig_summary.update_layout(
                margin=dict(l=20, r=20, t=60, b=20),
                xaxis_title_font_size=14,
                yaxis_title_font_size=14,
                title_x=0.5
            )
            st.plotly_chart(fig_summary, use_container_width=True)
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

    # Tab 3: Análisis de Ventas por Cliente
    with tab3:
        st.header(TRANSLATIONS[lang_code]['client_sales'])
        client_sales = filtered_df.groupby('Cliente/Nombre').agg({
            'Total': 'sum',
            'Número de recibo': 'nunique',
            'Comision': 'sum',
            'Cuentas por a cobrar aseavna': 'sum',
            'Cuentas por a Cobrar Avna': 'sum',
            'Líneas de la orden': lambda x: x.mode()[0] if not x.empty else 'N/A'
        }).reset_index()
        client_sales.columns = [
            'Cliente',
            'Ventas Totales (₡)',
            'Número de Órdenes',
            'Comisión Total (₡)',
            'Ctas. por Cobrar Aseavna (₡)',
            'Ctas. por Cobrar Avna (₡)',
            'Producto Más Comprado'
        ]
        
        avg_client_sales = client_sales['Ventas Totales (₡)'].mean()
        unusual = client_sales[client_sales['Ventas Totales (₡)'] > avg_client_sales * 2]
        if not unusual.empty:
            st.markdown(f'<div class="alert-box" style="background-color: {CONFIG["colors"]["warning"]}; color: black;">{TRANSLATIONS[lang_code]["unusual_sales"]}</div>', unsafe_allow_html=True)
            st.dataframe(unusual[['Cliente', 'Ventas Totales (₡)']])
        
        st.dataframe(client_sales)
        
        st.subheader(TRANSLATIONS[lang_code]['export_client_sales'])
        c1, c2, c3 = st.columns(3)
        with c1:
            csv_bytes = client_sales.to_csv(index=False).encode('utf-8')
            st.download_button(
                TRANSLATIONS[lang_code]['download_csv'],
                data=csv_bytes,
                file_name="ventas_por_cliente.csv",
                mime="text/csv"
            )
        with c2:
            buf_xl2 = generate_excel(client_sales, "Ventas por Cliente", client_sales.to_string())
            st.download_button(
                TRANSLATIONS[lang_code]['download_excel_client'],
                data=buf_xl2,
                file_name="ventas_por_cliente.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c3:
            buf_pdf2 = generate_pdf(client_sales, "Reporte de Ventas por Cliente - ASEAVNA", "ventas_por_cliente.pdf", client_sales.to_string())
            st.download_button(
                TRANSLATIONS[lang_code]['download_pdf_client'],
                data=buf_pdf2,
                file_name="ventas_por_cliente.pdf",
                mime="application/pdf"
            )

    # Tab 4: Análisis Predictivo
    with tab4:
        st.header(TRANSLATIONS[lang_code]['predictive'])
        try:
            daily = filtered_df.groupby(filtered_df['Fecha'].dt.date)['Total'].sum().reset_index(name='Total')
            daily['Days'] = (pd.to_datetime(daily['Fecha']) - pd.to_datetime(daily['Fecha'].min())).dt.days
            
            if len(daily) > 1:
                X = sm.add_constant(daily['Days'])
                model = sm.OLS(daily['Total'], X).fit()
                future_days = np.array([daily['Days'].iloc[-1] + i for i in range(1, 8)])
                future_X = sm.add_constant(future_days)
                preds = model.predict(future_X)
                conf_int = model.get_prediction(future_X).conf_int()
                
                pred_df = pd.DataFrame({
                    'Fecha': [pd.to_datetime(daily['Fecha']).max() + timedelta(days=i) for i in range(1, 8)],
                    'Total': preds,
                    'Lower': conf_int[:, 0],
                    'Upper': conf_int[:, 1],
                    'Tipo': 'Predicción'
                })
                hist_df = pd.DataFrame({
                    'Fecha': pd.to_datetime(daily['Fecha']),
                    'Total': daily['Total'],
                    'Tipo': 'Histórico'
                })
                combined = pd.concat([hist_df, pred_df])
                
                st.subheader(TRANSLATIONS[lang_code]['predictive_subheader'])
                fig_pred = px.line(
                    combined, x='Fecha', y='Total', color='Tipo',
                    labels={'Total': 'Ventas (₡)', 'Fecha': 'Fecha'},
                    title="Tendencia Histórica y Predicción de Ventas con Intervalos de Confianza",
                    template="plotly_white",
                    color_discrete_sequence=["#4CAF50", "#FF5733"]
                )
                fig_pred.add_scatter(
                    x=pred_df['Fecha'], y=pred_df['Upper'], mode='lines', line=dict(dash='dash', color='gray'), name='Límite Superior'
                )
                fig_pred.add_scatter(
                    x=pred_df['Fecha'], y=pred_df['Lower'], mode='lines', line=dict(dash='dash', color='gray'), name='Límite Inferior'
                )
                fig_pred.update_layout(
                    margin=dict(l=20, r=20, t=60, b=20),
                    xaxis_title_font_size=14,
                    yaxis_title_font_size=14,
                    title_x=0.5
                )
                st.plotly_chart(fig_pred, use_container_width=True)
                
                trends = filtered_df.groupby(['Líneas de la orden', filtered_df['Fecha'].dt.to_period('M')])['Total'].sum().unstack(fill_value=0)
                if trends.shape[1] >= 2:
                    growth = ((trends.iloc[:, -1] - trends.iloc[:, -2]) / trends.iloc[:, -2].replace(0, np.nan) * 100).replace([np.inf, -np.inf], 0).dropna().sort_values(ascending=False)
                    top5 = growth.head(5).reset_index()
                    top5.columns = ['Producto', 'Crecimiento (%)']
                    st.subheader(TRANSLATIONS[lang_code]['growth_subheader'])
                    st.dataframe(top5)
                else:
                    st.warning(TRANSLATIONS[lang_code]['no_monthly_data'])
            else:
                st.warning(TRANSLATIONS[lang_code]['no_predictive_data'])
        except Exception as e:
            st.error(TRANSLATIONS[lang_code]['predictive_error'].format(error=str(e)))

    # Tab 5: Visualizaciones Detalladas
    with tab5:
        st.header(TRANSLATIONS[lang_code]['visualizations'])
        # Filtrar por Centro de Costos para visualizaciones
        viz_df = filtered_df.copy()
        if selected_centro != 'Todos':
            viz_df = viz_df[viz_df['Centro de Costos Aseavna'] == selected_centro]
        
        # Top 10 Productos por Ventas con validación de datos
        top10 = viz_df.groupby('Líneas de la orden')['Total'].sum().nlargest(10).reset_index()
        if not top10.empty:
            top10['Total'] = top10['Total'].clip(upper=1000000)
            fig1 = px.bar(
                top10, x='Líneas de la orden', y='Total',
                title=TRANSLATIONS[lang_code]['top_products'],
                labels={'Total': 'Ventas (₡)', 'Líneas de la orden': 'Producto'},
                template="plotly_white",
                color_discrete_sequence=["#4CAF50"],
                hover_data={'Total': ':,.2f'}
            )
            fig1.update_layout(
                margin=dict(l=20, r=20, t=60, b=80),
                xaxis_tickangle=45,
                xaxis_title_font_size=14,
                yaxis_title_font_size=14,
                title_x=0.5
            )
            st.plotly_chart(fig1, use_container_width=True)
        else:
            st.warning("No hay datos suficientes para mostrar los top 10 productos.")

        daily_summary = viz_df.groupby(viz_df['Fecha'].dt.date)['Total'].sum().reset_index()
        if not daily_summary.empty:
            fig2 = px.line(
                daily_summary, x='Fecha', y='Total',
                labels={'Total': 'Ventas (₡)', 'Fecha': 'Fecha'},
                title=TRANSLATIONS[lang_code]['daily_trend'],
                template="plotly_white",
                color_discrete_sequence=["#4CAF50"]
            )
            fig2.update_layout(
                margin=dict(l=20, r=20, t=60, b=20),
                xaxis_title_font_size=14,
                yaxis_title_font_size=14,
                title_x=0.5
            )
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.warning("No hay datos suficientes para mostrar la tendencia diaria.")
        
        grp = viz_df.groupby('Cliente/Nombre principal')['Total'].sum().reset_index()
        if not grp.empty:
            fig3 = px.pie(
                grp, names='Cliente/Nombre principal', values='Total',
                title=TRANSLATIONS[lang_code]['sales_by_group'],
                template="plotly_white",
                color_discrete_sequence=px.colors.sequential.Viridis
            )
            fig3.update_layout(
                margin=dict(l=20, r=20, t=60, b=20),
                title_x=0.5
            )
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.warning("No hay datos suficientes para mostrar las ventas por grupo de clientes.")

    # Tab 6: Resumen de Métricas para Exportar
    with tab6:
        st.header(TRANSLATIONS[lang_code]['export'])
        most_sold = filtered_df.groupby('Líneas de la orden')['Total'].sum().idxmax() if not filtered_df.empty else "N/A"
        least_sold = filtered_df.groupby('Líneas de la orden')['Total'].sum().idxmin() if not filtered_df.empty else "N/A"
        
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