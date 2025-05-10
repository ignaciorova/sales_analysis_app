import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import locale
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import xlsxwriter
from sklearn.linear_model import LinearRegression
import numpy as np
import uuid
import io

# Configuración inicial
st.set_page_config(page_title="Dashboard de Ventas ASEAVNA", layout="wide")
try:
    locale.setlocale(locale.LC_TIME, 'es_ES')
except:
    locale.setlocale(locale.LC_TIME, '')

# Estilos CSS
st.markdown("""
<style>
.metric-box {
    background-color: #f0f2f6;
    padding: 20px;
    border-radius: 10px;
    text-align: center;
    box-shadow: 2px 2px 8px rgba(0,0,0,0.1);
}
.metric-box h3 {
    margin: 0;
    font-size: 1.2em;
    color: #333;
}
.metric-box p {
    font-size: 1.8em;
    font-weight: bold;
    color: #1f77b4;
    margin: 10px 0 0 0;
}
.sidebar .sidebar-content {
    background-color: #f8f9fa;
}
</style>
""", unsafe_allow_html=True)

# Función para cargar y limpiar datos
@st.cache_data
def load_data(file, file_name=None):
    """
    Carga y limpia el archivo de datos Excel desde un path o un objeto de archivo cargado.
    """
    try:
        if isinstance(file, str):
            if not os.path.exists(file):
                raise FileNotFoundError(f"El archivo {file} no se encuentra en el directorio {os.getcwd()}")
            df = pd.read_excel(file)
        else:
            df = pd.read_excel(file, engine='openpyxl')
        
        # Verificar columnas esperadas
        expected_columns = ['Fecha', 'Cliente/Nombre', 'Cliente/Nombre principal', 'Líneas de la orden', 
                           'Precio total colaborador', 'Comision Aseavna', 'Líneas de la orden/Cantidad', 'Número de recibo']
        missing_columns = [col for col in expected_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Faltan las siguientes columnas en el archivo: {missing_columns}")
        
        # Manejo de la columna Fecha
        if pd.api.types.is_datetime64_any_dtype(df['Fecha']):
            # Si ya es datetime, no necesita conversión
            st.write("La columna 'Fecha' ya está en formato datetime.")
        else:
            # Intentar convertir como número serial de Excel
            try:
                df['Fecha'] = pd.to_datetime(df['Fecha'], origin='1899-12-30', unit='D', errors='coerce')
            except:
                # Si falla, intentar como texto o formato de fecha
                df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        
        # Verificar si hay fechas nulas después de la conversión
        if df['Fecha'].isna().any():
            st.warning(f"Se encontraron {df['Fecha'].isna().sum()} filas con fechas inválidas. Estas se ignorarán en algunos análisis.")
        
        # Limpieza de datos
        df['Cliente/Nombre'] = df['Cliente/Nombre'].fillna('Desconocido')
        df['Cliente/Nombre principal'] = df['Cliente/Nombre principal'].fillna('Desconocido')
        df['Líneas de la orden'] = df['Líneas de la orden'].fillna('Sin Producto')
        df['Precio total colaborador'] = pd.to_numeric(df['Precio total colaborador'], errors='coerce').fillna(0)
        df['Comision Aseavna'] = pd.to_numeric(df['Comision Aseavna'], errors='coerce').fillna(0)
        df['Líneas de la orden/Cantidad'] = pd.to_numeric(df['Líneas de la orden/Cantidad'], errors='coerce').fillna(1)
        df['Número de recibo'] = df['Número de recibo'].fillna('Sin Recibo')
        
        # Mostrar información de depuración
        st.write(f"Filas cargadas: {len(df)}")
        st.write(f"Primeras fechas en la columna 'Fecha': {df['Fecha'].head().tolist()}")
        
        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo: {str(e)}")
        return pd.DataFrame()

# Función para calcular métricas
def calculate_metrics(df):
    if df.empty:
        return 0, 0, 0, 0
    total_sales = df['Precio total colaborador'].sum()
    num_orders = df['Número de recibo'].nunique()
    avg_order_value = total_sales / num_orders if num_orders > 0 else 0
    total_commission = df['Comision Aseavna'].sum()
    return total_sales, num_orders, avg_order_value, total_commission

# Función para generar PDF
def generate_pdf(data, title, filename):
    try:
        doc = SimpleDocTemplate(filename, pagesize=letter)
        elements = []
        styles = getSampleStyleSheet()
        elements.append(Paragraph(title, styles['Heading1']))
        data = data.astype(str)
        table_data = [data.columns.tolist()] + data.values.tolist()
        table = Table(table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
        ]))
        elements.append(table)
        doc.build(elements)
    except Exception as e:
        st.error(f"Error al generar PDF: {str(e)}")

# Función para generar Excel
def generate_excel(data, filename):
    try:
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            data.to_excel(writer, index=False, sheet_name='Reporte')
            worksheet = writer.sheets['Reporte']
            for idx, col in enumerate(data.columns):
                max_len = max(data[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, max_len)
    except Exception as e:
        st.error(f"Error al generar Excel: {str(e)}")

# Función para detectar duplicados
def detect_duplicates(df):
    duplicates = df[df['Líneas de la orden'] == 'Almuerzo Ejecutivo Aseavna']
    duplicates = duplicates[duplicates.duplicated(subset=['Cliente/Nombre', 'Fecha'], keep=False)]
    return duplicates

# Función para análisis por cliente
def client_analysis(df):
    if df.empty:
        return pd.DataFrame(columns=['Cliente', 'Ventas Totales', 'Número de Órdenes', 'Comisión Total', 'Producto Más Comprado'])
    analysis = df.groupby('Cliente/Nombre').agg({
        'Precio total colaborador': 'sum',
        'Número de recibo': 'nunique',
        'Comision Aseavna': 'sum',
        'Líneas de la orden': lambda x: x.value_counts().index[0] if not x.empty else 'N/A'
    }).reset_index()
    analysis.columns = ['Cliente', 'Ventas Totales', 'Número de Órdenes', 'Comisión Total', 'Producto Más Comprado']
    return analysis

# Función para predicciones de ventas
def predict_sales(daily_sales):
    if len(daily_sales) < 2:
        return pd.DataFrame()
    X = np.arange(len(daily_sales)).reshape(-1, 1)
    y = daily_sales['Precio total colaborador'].values
    model = LinearRegression()
    model.fit(X, y)
    future_days = 7
    future_X = np.arange(len(daily_sales), len(daily_sales) + future_days).reshape(-1, 1)
    future_predictions = model.predict(future_X)
    future_dates = [daily_sales['Fecha'].iloc[-1] + timedelta(days=i+1) for i in range(future_days)]
    return pd.DataFrame({'Fecha': future_dates, 'Ventas Pronosticadas': future_predictions})

# Función para calcular crecimiento de productos
def calculate_product_growth(df):
    df['Mes'] = df['Fecha'].dt.to_period('M')
    trends = df.pivot_table(values='Precio total colaborador', 
                           index='Líneas de la orden', 
                           columns='Mes', 
                           aggfunc='sum', 
                           fill_value=0)
    if trends.shape[1] >= 2:
        growth = ((trends.iloc[:, -1] - trends.iloc[:, -2]) / trends.iloc[:, -2].replace(0, np.nan) * 100).dropna()
        return growth.sort_values(ascending=False).head(5)
    return pd.Series()

# Mostrar directorio de trabajo actual
st.write(f"**Directorio de trabajo actual**: {os.getcwd()}")

# Intento de carga automática del archivo
file_path = "Órdenes del punto de venta (pos.order).xlsx"
df = pd.DataFrame()

if os.path.exists(file_path):
    df = load_data(file_path)
else:
    st.warning(f"El archivo {file_path} no se encuentra en el directorio {os.getcwd()}. Por favor, carga el archivo manualmente.")

# Selector de archivo manual
uploaded_file = st.file_uploader("Carga el archivo Excel (.xlsx)", type=["xlsx"], key="file_uploader")
if uploaded_file is not None:
    df = load_data(uploaded_file, uploaded_file.name)

# Procesar datos si se cargaron correctamente
if not df.empty:
    # Sidebar con filtros
    st.sidebar.header("Filtros")
    date_filter = st.sidebar.selectbox("Rango de Fechas", ["Personalizado", "Última Semana", "Último Mes"], key="date_filter")
    
    min_date = df['Fecha'].min()
    max_date = df['Fecha'].max()
    
    if date_filter == "Personalizado":
        start_date, end_date = st.sidebar.date_input("Selecciona el rango de fechas", 
                                                    [min_date, max_date], 
                                                    min_value=min_date, 
                                                    max_value=max_date, 
                                                    key="date_range")
    elif date_filter == "Última Semana":
        end_date = max_date
        start_date = end_date - timedelta(days=7)
    else:
        end_date = max_date
        start_date = end_date - timedelta(days=30)
    
    product_types = ['Todos'] + sorted(df['Líneas de la orden'].unique().tolist())
    selected_product = st.sidebar.selectbox("Tipo de Producto", product_types, key="product_filter")
    
    client_groups = ['Todos'] + sorted(df['Cliente/Nombre principal'].unique().tolist())
    selected_group = st.sidebar.selectbox("Grupo de Clientes", client_groups, key="group_filter")
    
    days = ['Todos', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
    selected_day = st.sidebar.selectbox("Día de la Semana", days, key="day_filter")
    
    clients = ['Todos'] + sorted(df['Cliente/Nombre'].unique().tolist())
    selected_client = st.sidebar.selectbox("Cliente", clients, key="client_filter")
    
    # Filtrado de datos
    filtered_df = df.copy()
    filtered_df = filtered_df[(filtered_df['Fecha'] >= pd.to_datetime(start_date)) & 
                             (filtered_df['Fecha'] <= pd.to_datetime(end_date))]
    
    if selected_product != 'Todos':
        filtered_df = filtered_df[filtered_df['Líneas de la orden'] == selected_product]
    
    if selected_group != 'Todos':
        filtered_df = filtered_df[filtered_df['Cliente/Nombre principal'] == selected_group]
    
    if selected_day != 'Todos':
        filtered_df = filtered_df[filtered_df['Fecha'].dt.day_name(locale='es_ES') == selected_day]
    
    if selected_client != 'Todos':
        filtered_df = filtered_df[filtered_df['Cliente/Nombre'] == selected_client]
    
    # Título principal
    st.title("Dashboard de Ventas ASEAVNA")
    
    # Análisis de duplicados
    st.header("Análisis de Almuerzos Duplicados")
    duplicates = detect_duplicates(filtered_df)
    
    if not duplicates.empty:
        st.write(f"Se encontraron {len(duplicates)} órdenes duplicadas de Almuerzo Ejecutivo.")
        st.dataframe(duplicates[['Cliente/Nombre', 'Fecha', 'Número de recibo', 'Precio total colaborador', 'Líneas de la orden']])
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Descargar Duplicados (Excel)", key="download_duplicates_excel"):
                excel_file = f"duplicados_{uuid.uuid4()}.xlsx"
                generate_excel(duplicates, excel_file)
                with open(excel_file, "rb") as f:
                    st.download_button("Descargar Excel", f, file_name=excel_file, key="download_excel_button")
        with col2:
            if st.button("Descargar Duplicados (PDF)", key="download_duplicates_pdf"):
                pdf_file = f"duplicados_{uuid.uuid4()}.pdf"
                generate_pdf(duplicates[['Cliente/Nombre', 'Fecha', 'Número de recibo', 'Precio total colaborador']], 
                            "Reporte de Duplicados", pdf_file)
                with open(pdf_file, "rb") as f:
                    st.download_button("Descargar PDF", f, file_name=pdf_file, key="download_pdf_button")
    else:
        st.write("No se encontraron órdenes duplicadas de Almuerzo Ejecutivo.")
    
    # Métricas generales
    st.header("Métricas Generales")
    total_sales, num_orders, avg_order_value, total_commission = calculate_metrics(filtered_df)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class="metric-box">
            <h3>Ventas Totales</h3>
            <p>₡{total_sales:,.2f}</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-box">
            <h3>Número de Órdenes</h3>
            <p>{num_orders}</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="metric-box">
            <h3>Valor Promedio por Orden</h3>
            <p>₡{avg_order_value:,.2f}</p>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="metric-box">
            <h3>Comisión Total</h3>
            <p>₡{total_commission:,.2f}</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Análisis por cliente
    st.header("Análisis por Cliente")
    client_data = client_analysis(filtered_df)
    
    if not client_data.empty:
        st.dataframe(client_data)
        
        avg_sales = client_data['Ventas Totales'].mean()
        unusual_clients = client_data[client_data['Ventas Totales'] > 2 * avg_sales]
        
        if not unusual_clients.empty:
            st.write("Clientes con compras inusuales (ventas > 2x promedio):")
            st.dataframe(unusual_clients)
        else:
            st.write("No se encontraron clientes con compras inusuales.")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("Descargar Análisis por Cliente (CSV)", key="download_client_csv"):
                csv = client_data.to_csv(index=False)
                st.download_button("Descargar CSV", csv, file_name="analisis_clientes.csv", mime="text/csv", key="download_csv_button")
        with col2:
            if st.button("Descargar Análisis por Cliente (Excel)", key="download_client_excel"):
                excel_file = f"analisis_clientes_{uuid.uuid4()}.xlsx"
                generate_excel(client_data, excel_file)
                with open(excel_file, "rb") as f:
                    st.download_button("Descargar Excel", f, file_name=excel_file, key="download_client_excel_button")
        with col3:
            if st.button("Descargar Análisis por Cliente (PDF)", key="download_client_pdf"):
                pdf_file = f"analisis_clientes_{uuid.uuid4()}.pdf"
                generate_pdf(client_data, "Análisis por Cliente", pdf_file)
                with open(pdf_file, "rb") as f:
                    st.download_button("Descargar PDF", f, file_name=pdf_file, key="download_client_pdf_button")
    else:
        st.write("No hay datos de clientes para mostrar.")
    
    # Visualizaciones
    st.header("Visualizaciones")
    
    product_sales = filtered_df.groupby('Líneas de la orden')['Precio total colaborador'].sum().reset_index()
    product_sales = product_sales.sort_values('Precio total colaborador', ascending=False).head(10)
    fig1 = px.bar(product_sales, x='Líneas de la orden', y='Precio total colaborador', 
                  title="Top 10 Productos por Ventas",
                  labels={'Precio total colaborador': 'Ventas (₡)', 'Líneas de la orden': 'Producto'},
                  text_auto='.2s')
    fig1.update_traces(textfont_size=12, textangle=0, textposition="outside", cliponaxis=False)
    fig1.update_layout(xaxis_tickangle=45)
    st.plotly_chart(fig1, use_container_width=True)
    
    daily_sales = filtered_df.groupby(filtered_df['Fecha'].dt.date)['Precio total colaborador'].sum().reset_index()
    daily_sales['Fecha'] = pd.to_datetime(daily_sales['Fecha'])
    fig2 = px.line(daily_sales, x='Fecha', y='Precio total colaborador', 
                   title="Tendencia Diaria de Ventas",
                   labels={'Precio total colaborador': 'Ventas (₡)', 'Fecha': 'Fecha'})
    fig2.update_layout(xaxis_title="Fecha", yaxis_title="Ventas (₡)")
    st.plotly_chart(fig2, use_container_width=True)
    
    group_sales = filtered_df.groupby('Cliente/Nombre principal')['Precio total colaborador'].sum().reset_index()
    fig3 = px.pie(group_sales, names='Cliente/Nombre principal', values='Precio total colaborador', 
                  title="Distribución de Ventas por Grupo de Clientes",
                  labels={'Precio total colaborador': 'Ventas (₡)'})
    fig3.update_traces(textinfo='percent+label')
    st.plotly_chart(fig3, use_container_width=True)
    
    product_type_sales = filtered_df.groupby('Líneas de la orden')['Precio total colaborador'].sum().reset_index()
    fig4 = px.pie(product_type_sales, names='Líneas de la orden', values='Precio total colaborador', 
                  title="Distribución de Ventas por Tipo de Producto",
                  labels={'Precio total colaborador': 'Ventas (₡)'})
    fig4.update_traces(textinfo='percent+label')
    st.plotly_chart(fig4, use_container_width=True)
    
    # Análisis predictivo
    st.header("Análisis Predictivo")
    predictions_df = predict_sales(daily_sales)
    
    if not predictions_df.empty:
        fig5 = px.line(predictions_df, x='Fecha', y='Ventas Pronosticadas', 
                       title="Pronóstico de Ventas (7 días)",
                       labels={'Ventas Pronosticadas': 'Ventas (₡)', 'Fecha': 'Fecha'})
        fig5.update_layout(xaxis_title="Fecha", yaxis_title="Ventas Pronosticadas (₡)")
        st.plotly_chart(fig5, use_container_width=True)
        
        top_growth = calculate_product_growth(filtered_df)
        if not top_growth.empty:
            st.write("Productos con mayor crecimiento mensual (%):")
            st.write(top_growth)
        else:
            st.warning("No hay suficientes datos mensuales para calcular el crecimiento.")
    else:
        st.warning("No hay suficientes datos para realizar predicciones.")
    
    if st.checkbox("Mostrar datos crudos", key="show_raw_data"):
        st.header("Datos Filtrados")
        st.dataframe(filtered_df)
else:
    st.error("No se pudieron cargar los datos. Verifica que el archivo esté en el directorio correcto o cárgalo manualmente usando el selector de archivos.")