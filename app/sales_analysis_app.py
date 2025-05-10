import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import io
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import xlsxwriter
from sklearn.linear_model import LinearRegression
import numpy as np

# Configuración de la página
st.set_page_config(
    page_title="Análisis de Ventas - ASEAVNA",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo personalizado
st.markdown("""
    <style>
    .main {background-color: #f5f7fa;}
    .stButton>button {background-color: #4CAF50; color: white; border-radius: 5px;}
    .stSidebar {background-color: #e8ecef;}
    h1, h2, h3 {color: #2c3e50;}
    .metric-box {border: 1px solid #d3d3d3; padding: 10px; border-radius: 5px; background-color: white;}
    .alert-box {background-color: #ffeb3b; padding: 10px; border-radius: 5px; margin: 10px 0;}
    </style>
""", unsafe_allow_html=True)

# Título y descripción
st.title("📊 Dashboard de Análisis de Ventas - ASEAVNA")
st.markdown("Análisis avanzado de órdenes de venta del sistema POS, con métricas, predicciones y reportes descargables por cliente.")

@st.cache_data
def load_data(uploaded_file):
    try:
        if uploaded_file is None:
            raise FileNotFoundError("No se ha cargado ningún archivo.")
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Mostrar columnas detectadas para depuración
        st.write("Columnas detectadas en el archivo:", df.columns.tolist())
        
        # Detectar si 'Fecha' es serial de Excel (numérico) o ya datetime
        if pd.api.types.is_numeric_dtype(df['Fecha']):
            df['Fecha'] = pd.to_datetime(df['Fecha'], unit='D', origin='1899-12-30') - timedelta(days=2)
        else:
            df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        
        # Mapeo flexible de columnas
        column_mapping = {
            'Cliente/Código de barras': 'Cliente/Código de barras',
            'Cliente/Nombre principal': 'Cliente/Nombre principal',
            'Cliente/Nombre': 'Cliente/Nombre',
            'Líneas de la orden/Cantidad': 'Líneas de la orden/Cantidad',
            'Total': 'Precio total colaborador',  # Mapeo a la columna existente
            'Comision': 'Comision Aseavna',      # Nueva métrica
            'Líneas de la orden': 'Líneas de la orden',
            'Número de recibo': 'Número de recibo'
        }
        
        # Asignar columnas con fallback
        for expected_col, actual_col in column_mapping.items():
            if actual_col in df.columns:
                df[expected_col] = df[actual_col]
            else:
                df[expected_col] = 'Desconocido' if 'Cliente' in expected_col or 'Líneas' in expected_col else 0
        
        # Limpieza de datos
        df['Cliente/Código de barras'] = df['Cliente/Código de barras'].fillna('Desconocido')
        df['Cliente/Nombre principal'] = df['Cliente/Nombre principal'].fillna('Desconocido')
        df['Cliente/Nombre'] = df['Cliente/Nombre'].fillna('Desconocido')
        df['Líneas de la orden/Cantidad'] = df['Líneas de la orden/Cantidad'].fillna(0)
        df['Total'] = pd.to_numeric(df['Total'], errors='coerce').fillna(0)
        df['Comision'] = pd.to_numeric(df['Comision'], errors='coerce').fillna(0)
        df['Líneas de la orden'] = df['Líneas de la orden'].fillna('Desconocido')
        
        # Añadir día de la semana
        df['Día de la Semana'] = df['Fecha'].dt.day_name(locale='es_ES')
        
        # Depuración
        st.write(f"Filas cargadas: {len(df)}")
        st.write(f"Primeras fechas: {df['Fecha'].head().tolist()}")
        
        return df

    except Exception as e:
        st.error(f"Error al cargar los datos: {str(e)}")
        return pd.DataFrame()

def generate_pdf(data: pd.DataFrame, title: str, filename: str) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()
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

def generate_excel(data: pd.DataFrame, sheet_name: str) -> io.BytesIO:
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

# Carga de datos
uploaded_file = st.file_uploader("Carga el archivo Excel (.xlsx)", type=["xlsx"], key="file_uploader")
df = load_data(uploaded_file)

if df.empty:
    st.warning("No se encontraron datos. Por favor, carga el archivo 'Órdenes del punto de venta (pos.order).xlsx' usando el selector de archivos.")
else:
    # Sidebar: filtros de fecha y categorías
    st.sidebar.header("Filtros de Análisis")
    st.sidebar.subheader("Rango de Fechas")
    date_option = st.sidebar.selectbox(
        "Seleccionar Período",
        ["Personalizado", "Última Semana", "Último Mes"]
    )
    if date_option == "Última Semana":
        end_date = df['Fecha'].max().date()
        start_date = end_date - timedelta(days=7)
    elif date_option == "Último Mes":
        end_date = df['Fecha'].max().date()
        start_date = end_date - timedelta(days=30)
    else:
        start_date = df['Fecha'].min().date()
        end_date = df['Fecha'].max().date()
    
    date_range = st.sidebar.date_input(
        "Rango de Fechas",
        [start_date, end_date],
        min_value=df['Fecha'].min().date(),
        max_value=df['Fecha'].max().date()
    )
    
    # Listas de filtros convirtiendo todo a str para evitar mezcla de tipos
    product_types = ['Todos'] + sorted(df['Líneas de la orden'].dropna().astype(str).unique().tolist())
    selected_product = st.sidebar.selectbox("Tipo de Producto", product_types)
    
    client_groups = ['Todos'] + sorted(df['Cliente/Nombre principal'].dropna().astype(str).unique().tolist())
    selected_client_grp = st.sidebar.selectbox("Grupo de Clientes", client_groups)
    
    days_of_week = ['Todos'] + sorted(df['Día de la Semana'].dropna().astype(str).unique().tolist())
    selected_day = st.sidebar.selectbox("Día de la Semana", days_of_week)
    
    clients = ['Todos'] + sorted(df['Cliente/Nombre'].dropna().astype(str).unique().tolist())
    selected_client = st.sidebar.selectbox("Cliente Específico", clients)
    
    # Aplicar filtros
    filtered_df = df.copy()
    if len(date_range) == 2:
        sd, ed = date_range
        filtered_df = filtered_df[
            (filtered_df['Fecha'] >= pd.to_datetime(sd)) &
            (filtered_df['Fecha'] <= pd.to_datetime(ed))
        ]
    if selected_product != 'Todos':
        filtered_df = filtered_df[filtered_df['Líneas de la orden'] == selected_product]
    if selected_client_grp != 'Todos':
        filtered_df = filtered_df[filtered_df['Cliente/Nombre principal'] == selected_client_grp]
    if selected_day != 'Todos':
        filtered_df = filtered_df[filtered_df['Día de la Semana'] == selected_day]
    if selected_client != 'Todos':
        filtered_df = filtered_df[filtered_df['Cliente/Nombre'] == selected_client]
    
    # Verificación de almuerzos ejecutivos duplicados
    st.header("Verificación de Almuerzos Ejecutivos Duplicados")
    lunch_df = filtered_df[filtered_df['Líneas de la orden'] == 'Almuerzo Ejecutivo Aseavna'].copy()
    lunch_df['Fecha_Dia'] = lunch_df['Fecha'].dt.date
    dup = lunch_df.groupby(['Cliente/Nombre', 'Fecha_Dia']).filter(lambda x: len(x) > 1)
    
    if not dup.empty:
        st.markdown('<div class="alert-box">⚠️ Se detectaron almuerzos ejecutivos duplicados:</div>', unsafe_allow_html=True)
        summary = dup.groupby(['Cliente/Nombre', 'Fecha_Dia']).size().reset_index(name='Cantidad')
        st.dataframe(summary)
        st.subheader("Detalles de Duplicados")
        st.dataframe(dup[['Cliente/Nombre', 'Fecha', 'Número de recibo', 'Líneas de la orden']])
        c1, c2 = st.columns(2)
        with c1:
            buf_xl = generate_excel(dup, "Duplicados")
            st.download_button(
                "Descargar Duplicados (Excel)",
                data=buf_xl,
                file_name="almuerzos_duplicados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c2:
            buf_pdf = generate_pdf(dup, "Reporte de Almuerzos Duplicados", "almuerzos_duplicados.pdf")
            st.download_button(
                "Descargar Duplicados (PDF)",
                data=buf_pdf,
                file_name="almuerzos_duplicados.pdf",
                mime="application/pdf"
            )
    else:
        st.success("✅ No se detectaron almuerzos ejecutivos duplicados en el mismo día.")
    
    # Métricas Generales
    st.header("Métricas Generales")
    col1, col2, col3, col4 = st.columns(4)
    total_sales = filtered_df['Total'].sum()
    num_orders = filtered_df['Número de recibo'].nunique()
    avg_order_value = total_sales / num_orders if num_orders > 0 else 0
    total_commission = filtered_df['Comision'].sum()
    
    with col1:
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.metric("Ventas Totales", f"₡{total_sales:,.2f}")
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.metric("Número de Órdenes", f"{num_orders:,}")
        st.markdown('</div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.metric("Valor Promedio por Orden", f"₡{avg_order_value:,.2f}")
        st.markdown('</div>', unsafe_allow_html=True)
    with col4:
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.metric("Comisión Total", f"₡{total_commission:,.2f}")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Análisis de Ventas por Cliente
    st.header("Análisis de Ventas por Cliente")
    client_sales = filtered_df.groupby('Cliente/Nombre').agg({
        'Total': 'sum',
        'Número de recibo': 'nunique',
        'Comision': 'sum',
        'Líneas de la orden': lambda x: x.mode()[0] if not x.empty else 'N/A'
    }).reset_index()
    client_sales.columns = [
        'Cliente',
        'Ventas Totales (₡)',
        'Número de Órdenes',
        'Comisión Total (₡)',
        'Producto Más Comprado'
    ]
    
    # Clientes con compras inusuales
    avg_client_sales = client_sales['Ventas Totales (₡)'].mean()
    unusual = client_sales[client_sales['Ventas Totales (₡)'] > avg_client_sales * 2]
    if not unusual.empty:
        st.markdown('<div class="alert-box">⚠️ Clientes con volumen de compras inusual:</div>', unsafe_allow_html=True)
        st.dataframe(unusual[['Cliente', 'Ventas Totales (₡)']])
    
    st.dataframe(client_sales)
    
    # Exportar reporte por cliente
    st.subheader("Exportar Reporte de Ventas por Cliente")
    c1, c2, c3 = st.columns(3)
    with c1:
        csv_bytes = client_sales.to_csv(index=False).encode('utf-8')
        st.download_button(
            "Descargar CSV",
            data=csv_bytes,
            file_name="ventas_por_cliente.csv",
            mime="text/csv"
        )
    with c2:
        buf_xl2 = generate_excel(client_sales, "Ventas por Cliente")
        st.download_button(
            "Descargar Excel",
            data=buf_xl2,
            file_name="ventas_por_cliente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with c3:
        buf_pdf2 = generate_pdf(client_sales, "Reporte de Ventas por Cliente - ASEAVNA", "ventas_por_cliente.pdf")
        st.download_button(
            "Descargar PDF",
            data=buf_pdf2,
            file_name="ventas_por_cliente.pdf",
            mime="application/pdf"
        )
    
    # Análisis Predictivo
    st.header("Análisis Predictivo")
    daily = filtered_df.groupby(filtered_df['Fecha'].dt.date)['Total'].sum().reset_index(name='Total')
    daily['Days'] = (pd.to_datetime(daily['Fecha']) - pd.to_datetime(daily['Fecha'].min())).dt.days
    
    if len(daily) > 1:
        X = daily[['Days']]
        y = daily['Total']
        model = LinearRegression().fit(X, y)
        
        future_X = np.array([X['Days'].iloc[-1] + i for i in range(1, 8)]).reshape(-1, 1)
        preds = model.predict(future_X)
        future_dates = [pd.to_datetime(daily['Fecha']).max() + timedelta(days=i) for i in range(1, 8)]
        
        pred_df = pd.DataFrame({
            'Fecha': future_dates,
            'Total': preds,
            'Tipo': 'Predicción'
        })
        hist_df = pd.DataFrame({
            'Fecha': pd.to_datetime(daily['Fecha']),
            'Total': daily['Total'],
            'Tipo': 'Histórico'
        })
        combined = pd.concat([hist_df, pred_df])
        
        st.subheader("Predicción de Ventas para los Próximos 7 Días")
        fig_pred = px.line(
            combined, x='Fecha', y='Total', color='Tipo',
            labels={'Total': 'Ventas (₡)', 'Fecha': 'Fecha'},
            title="Tendencia Histórica y Predicción de Ventas"
        )
        st.plotly_chart(fig_pred, use_container_width=True)
        
        # Crecimiento mensual productos
        trends = filtered_df.groupby(['Líneas de la orden', filtered_df['Fecha'].dt.to_period('M')])['Total'].sum().unstack(fill_value=0)
        
        if trends.shape[1] >= 2:
            growth = ((trends.iloc[:, -1] - trends.iloc[:, -2]) / trends.iloc[:, -2].replace(0, np.nan) * 100).replace([np.inf, -np.inf], 0).dropna().sort_values(ascending=False)
            top5 = growth.head(5).reset_index()
            top5.columns = ['Producto', 'Crecimiento (%)']
            st.subheader("Productos con Potencial de Crecimiento")
            st.dataframe(top5)
        else:
            st.warning("No hay suficientes datos mensuales para calcular el crecimiento de productos (se requieren al menos dos meses).")
    else:
        st.warning("No hay suficientes datos históricos para predicción.")
    
    # Visualizaciones Detalladas
    st.header("Visualizaciones Detalladas")
    # Top 10 productos
    top10 = filtered_df.groupby('Líneas de la orden')['Total'].sum().nlargest(10).reset_index()
    fig1 = px.bar(
        top10, x='Líneas de la orden', y='Total',
        title="Top 10 Productos por Ventas",
        labels={'Total': 'Ventas (₡)', 'Líneas de la orden': 'Producto'}
    )
    fig1.update_layout(xaxis_tickangle=45)
    st.plotly_chart(fig1, use_container_width=True)
    
    # Tendencia diaria
    fig2 = px.line(
        x=pd.to_datetime(daily['Fecha']),
        y=daily['Total'],
        labels={'x': 'Fecha', 'y': 'Ventas (₡)'},
        title="Tendencia Diaria de Ventas"
    )
    st.plotly_chart(fig2, use_container_width=True)
    
    # Pie de ventas por grupo
    grp = filtered_df.groupby('Cliente/Nombre principal')['Total'].sum().reset_index()
    fig3 = px.pie(
        grp, names='Cliente/Nombre principal', values='Total',
        title="Ventas por Grupo de Clientes"
    )
    st.plotly_chart(fig3, use_container_width=True)
    
    # Resumen de Métricas para exportar
    most_sold = filtered_df.groupby('Líneas de la orden')['Total'].sum().idxmax() if not filtered_df.empty else "N/A"
    least_sold = filtered_df.groupby('Líneas de la orden')['Total'].sum().idxmin() if not filtered_df.empty else "N/A"
    
    st.header("Exportar Resumen de Métricas")
    report = {
        "Ventas Totales (₡)": total_sales,
        "Número de Órdenes": num_orders,
        "Valor Promedio por Orden (₡)": avg_order_value,
        "Comisión Total (₡)": total_commission,
        "Clientes Únicos": len(clients) - 1,
        "Producto Más Vendido": most_sold,
        "Producto Menos Vendido": least_sold
    }
    report_df = pd.DataFrame([report])
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            "Descargar Resumen (CSV)",
            data=report_df.to_csv(index=False).encode('utf-8'),
            file_name="resumen_ventas_aseavna.csv",
            mime="text/csv"
        )
    with c2:
        buf_xl3 = generate_excel(report_df, "Resumen")
        st.download_button(
            "Descargar Resumen (Excel)",
            data=buf_xl3,
            file_name="resumen_ventas_aseavna.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with c3:
        buf_pdf3 = generate_pdf(report_df, "Resumen de Ventas - ASEAVNA", "resumen_ventas_aseavna.pdf")
        st.download_button(
            "Descargar Resumen (PDF)",
            data=buf_pdf3,
            file_name="resumen_ventas_aseavna.pdf",
            mime="application/pdf"
        )
    
    # Mostrar datos crudos
    st.header("Datos Crudos")
    if st.checkbox("Mostrar Datos Crudos"):
        st.dataframe(filtered_df)
    
    # Pie de página
    st.markdown("---")
    st.markdown("Desarrollado por Wilfredos para ASEAVNA | Fuente de Datos: Órdenes del Punto de Venta (POS) | 2025")