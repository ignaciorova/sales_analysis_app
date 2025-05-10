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
import os
import smtplib
from email.mime.text import MIMEText
from streamlit_aggrid import AgGrid, GridOptionsBuilder  # Mantener este import
from pathlib import Path
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()
EMAIL_FROM = os.getenv("EMAIL_FROM")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Configuración de la página
st.set_page_config(
    page_title="Análisis de Ventas - ASEAVNA",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo personalizado con paleta de colores ASEAVNA (tonos verdes)
st.markdown("""
    <style>
    .main {background-color: #f5f7fa;}
    .stButton>button {background-color: #2E7D32; color: white; border-radius: 5px;}
    .stSidebar {background-color: #E8F5E9;}
    h1, h2, h3 {color: #1B5E20;}
    .metric-box {border: 1px solid #C8E6C9; padding: 15px; border-radius: 5px; background-color: #E8F5E9; margin: 5px 0; text-align: center;}
    .alert-box {background-color: #FFECB3; padding: 10px; border-radius: 5px; margin: 10px 0;}
    .stMetric {font-size: 18px;}
    </style>
""", unsafe_allow_html=True)

# Cargar logo
logo_path = Path("app/assets/logo.png")
if logo_path.exists():
    st.sidebar.image(str(logo_path), width=200)
else:
    st.sidebar.warning("Logo no encontrado en app/assets/logo.png")

@st.cache_data
def load_data():
    try:
        data_path = Path("app/data/Orders_pos.xlsx")
        if not data_path.exists():
            st.error(f"Archivo no encontrado: {data_path}")
            return pd.DataFrame()
        df = pd.read_excel(data_path, engine='openpyxl')
        st.write("Columnas cargadas:", df.columns.tolist())  # Depuración

        if pd.api.types.is_numeric_dtype(df['Fecha']):
            df['Fecha'] = pd.to_datetime(df['Fecha'], unit='D', origin='1899-12-30') - timedelta(days=2)
        else:
            df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')

        df_columns = {col.strip().lower(): col for col in df.columns}
        column_mapping = {
            'Cliente/Código de barras': 'cliente/código de barras',
            'Cliente/Nombre': 'cliente/nombre',
            'Centro de Costos Aseavna': 'centro de costos aseavna',
            'Fecha': 'fecha',
            'Número de recibo': 'número de recibo',
            'Cliente/Nombre principal': 'cliente/nombre principal',
            'Total': 'precio total colaborador',
            'Comision': 'comision aseavna',
            'Cuentas por a cobrar aseavna': 'cuentas por a cobrar aseavna',
            'Cuentas por a Cobrar Avna': 'cuentas por a cobrar avna',
            'Líneas de la orden': 'líneas de la orden',
            'Líneas de la orden/Cantidad': 'líneas de la orden/cantidad'
        }

        for expected_col, search_col in column_mapping.items():
            found_col = next((col for col_name, col in df_columns.items() if col_name == search_col.strip().lower()), None)
            if found_col:
                df[expected_col] = df[found_col]
            else:
                df[expected_col] = 'Desconocido' if 'Cliente' in expected_col or 'Líneas' in expected_col else 0

        df['Cliente/Código de barras'] = df['Cliente/Código de barras'].fillna('Desconocido')
        df['Cliente/Nombre'] = df['Cliente/Nombre'].fillna('Desconocido')
        df['Centro de Costos Aseavna'] = df['Centro de Costos Aseavna'].fillna('Desconocido')
        df['Cliente/Nombre principal'] = df['Cliente/Nombre principal'].fillna('Desconocido')
        df['Líneas de la orden'] = df['Líneas de la orden'].fillna('Desconocido')
        df['Líneas de la orden/Cantidad'] = pd.to_numeric(df['Líneas de la orden/Cantidad'], errors='coerce').fillna(0)
        df['Total'] = pd.to_numeric(df['Total'], errors='coerce').fillna(0)
        df['Comision'] = pd.to_numeric(df['Comision'], errors='coerce').fillna(0)
        df['Cuentas por a cobrar aseavna'] = pd.to_numeric(df['Cuentas por a cobrar aseavna'], errors='coerce').fillna(0)
        df['Cuentas por a Cobrar Avna'] = pd.to_numeric(df['Cuentas por a Cobrar Avna'], errors='coerce').fillna(0)

        df['Día de la Semana'] = df['Fecha'].dt.day_name()
        day_translation = {
            'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
            'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
        }
        df['Día de la Semana'] = df['Día de la Semana'].map(day_translation).fillna(df['Día de la Semana'])

        return df
    except Exception as e:
        st.error(f"Error al cargar los datos: {e}")
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

def send_alert(email_to, subject, message):
    if not EMAIL_FROM or not EMAIL_PASSWORD:
        st.error("Credenciales de correo no configuradas en .env")
        return
    msg = MIMEText(message)
    msg['Subject'] = subject
    msg['From'] = EMAIL_FROM
    msg['To'] = email_to
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.sendmail(EMAIL_FROM, email_to, msg.as_string())
        st.success("Correo enviado exitosamente")
    except Exception as e:
        st.error(f"Error al enviar correo: {e}")

# Carga de datos automática
df = load_data()

if df.empty:
    st.warning("No se encontraron datos. Asegúrese de que el archivo 'Orders_pos.xlsx' esté disponible en app/data/.")
else:
    # Sidebar: filtros avanzados
    st.sidebar.header("Filtros de Análisis")
    st.sidebar.subheader("Rango de Fechas")
    date_option = st.sidebar.selectbox(
        "Seleccionar Período", ["Personalizado", "Última Semana", "Último Mes", "Todo el Período"]
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

    date_range = st.sidebar.date_input(
        "Rango de Fechas", [start_date, end_date],
        min_value=df['Fecha'].min().date(), max_value=df['Fecha'].max().date()
    )

    # Filtros avanzados
    min_amount = st.sidebar.number_input("Monto Mínimo (₡)", min_value=0.0, value=0.0, step=1000.0)
    max_amount = st.sidebar.number_input("Monto Máximo (₡)", min_value=0.0, value=df['Total'].max(), step=1000.0)
    custom_category = st.sidebar.text_input("Categoría Personalizada", "")
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
        filtered_df = filtered_df[(filtered_df['Fecha'] >= pd.to_datetime(sd)) & (filtered_df['Fecha'] <= pd.to_datetime(ed))]
    if selected_product != 'Todos':
        filtered_df = filtered_df[filtered_df['Líneas de la orden'] == selected_product]
    if selected_client_grp != 'Todos':
        filtered_df = filtered_df[filtered_df['Cliente/Nombre principal'] == selected_client_grp]
    if selected_day != 'Todos':
        filtered_df = filtered_df[filtered_df['Día de la Semana'] == selected_day]
    if selected_client != 'Todos':
        filtered_df = filtered_df[filtered_df['Cliente/Nombre'] == selected_client]
    if min_amount > 0 or max_amount < df['Total'].max():
        filtered_df = filtered_df[(filtered_df['Total'] >= min_amount) & (filtered_df['Total'] <= max_amount)]
    if custom_category:
        filtered_df = filtered_df[filtered_df['Líneas de la orden'].str.contains(custom_category, case=False, na=False)]

    # Crear pestañas
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "Métricas Generales", "Almuerzos Duplicados", "Ventas por Cliente",
        "Análisis Predictivo", "Visualizaciones", "Exportar Resumen", "Datos Crudos"
    ])

    # Tab 1: Métricas Generales
    with tab1:
        st.header("Métricas Generales")
        col1, col2, col3, col4 = st.columns(4)
        total_orders = df['Número de recibo'].nunique()
        total_commission = df['Comision'].sum()
        total_cuentas_cobrar_aseavna = df['Cuentas por a cobrar aseavna'].sum()
        total_cuentas_cobrar_avna = df['Cuentas por a Cobrar Avna'].sum()

        with col1:
            st.markdown('<div class="metric-box">', unsafe_allow_html=True)
            st.metric("Número de Órdenes", f"{total_orders:,}")
            st.markdown('</div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div class="metric-box">', unsafe_allow_html=True)
            st.metric("Comisión Total", f"₡{total_commission:,.2f}")
            st.markdown('</div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="metric-box">', unsafe_allow_html=True)
            st.metric("Ctas. por Cobrar Aseavna", f"₡{total_cuentas_cobrar_aseavna:,.2f}")
            st.markdown('</div>', unsafe_allow_html=True)
        with col4:
            st.markdown('<div class="metric-box">', unsafe_allow_html=True)
            st.metric("Ctas. por Cobrar Avna", f"₡{total_cuentas_cobrar_avna:,.2f}")
            st.markdown('</div>', unsafe_allow_html=True)

    # Tab 2: Verificación de Almuerzos Duplicados
    with tab2:
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
            # Alerta de duplicados
            st.warning("Se detectaron duplicados. Revisa los detalles arriba.")
            if st.button("Enviar Alerta por Correo"):
                send_alert("tu_email@gmail.com", "Alerta: Duplicados Detectados", "Se encontraron almuerzos ejecutivos duplicados.")
        else:
            st.success("✅ No se detectaron almuerzos ejecutivos duplicados en el mismo día.")

    # Tab 3: Análisis de Ventas por Cliente
    with tab3:
        st.header("Análisis de Ventas por Cliente")
        client_sales = filtered_df.groupby('Cliente/Nombre').agg({
            'Total': 'sum', 'Número de recibo': 'nunique', 'Comision': 'sum',
            'Cuentas por a cobrar aseavna': 'sum', 'Cuentas por a Cobrar Avna': 'sum',
            'Líneas de la orden': lambda x: x.mode()[0] if not x.empty else 'N/A'
        }).reset_index()
        client_sales.columns = [
            'Cliente', 'Ventas Totales (₡)', 'Número de Órdenes', 'Comisión Total (₡)',
            'Ctas. por Cobrar Aseavna (₡)', 'Ctas. por Cobrar Avna (₡)', 'Producto Más Comprado'
        ]
        avg_client_sales = client_sales['Ventas Totales (₡)'].mean()
        unusual = client_sales[client_sales['Ventas Totales (₡)'] > avg_client_sales * 2]
        if not unusual.empty:
            st.markdown('<div class="alert-box">⚠️ Clientes con volumen de compras inusual:</div>', unsafe_allow_html=True)
            st.dataframe(unusual[['Cliente', 'Ventas Totales (₡)']])
            st.warning("Se detectaron clientes con compras inusuales. Revisa los detalles arriba.")
            if st.button("Enviar Alerta por Correo"):
                send_alert("tu_email@gmail.com", "Alerta: Compras Inusuales", "Se detectaron clientes con compras inusuales.")
        st.dataframe(client_sales)
        st.subheader("Exportar Reporte de Ventas por Cliente")
        c1, c2, c3 = st.columns(3)
        with c1:
            csv_bytes = client_sales.to_csv(index=False).encode('utf-8')
            st.download_button("Descargar CSV", data=csv_bytes, file_name="ventas_por_cliente.csv", mime="text/csv")
        with c2:
            buf_xl2 = generate_excel(client_sales, "Ventas por Cliente")
            st.download_button("Descargar Excel", data=buf_xl2, file_name="ventas_por_cliente.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c3:
            buf_pdf2 = generate_pdf(client_sales, "Reporte de Ventas por Cliente - ASEAVNA", "ventas_por_cliente.pdf")
            st.download_button("Descargar PDF", data=buf_pdf2, file_name="ventas_por_cliente.pdf", mime="application/pdf")

    # Tab 4: Análisis Predictivo
    with tab4:
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
            pred_df = pd.DataFrame({'Fecha': future_dates, 'Total': preds, 'Tipo': 'Predicción'})
            hist_df = pd.DataFrame({'Fecha': pd.to_datetime(daily['Fecha']), 'Total': daily['Total'], 'Tipo': 'Histórico'})
            combined = pd.concat([hist_df, pred_df])
            st.subheader("Predicción de Ventas para los Próximos 7 Días")
            fig_pred = px.line(combined, x='Fecha', y='Total', color='Tipo', labels={'Total': 'Ventas (₡)', 'Fecha': 'Fecha'}, title="Tendencia Histórica y Predicción de Ventas")
            fig_pred.update_layout(dragmode='zoom', hovermode='x unified')
            st.plotly_chart(fig_pred, use_container_width=True)
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

    # Tab 5: Visualizaciones Detalladas
    with tab5:
        st.header("Visualizaciones Detalladas")
        top10 = filtered_df.groupby('Líneas de la orden')['Total'].sum().nlargest(10).reset_index()
        fig1 = px.bar(top10, x='Líneas de la orden', y='Total', title="Top 10 Productos por Ventas", labels={'Total': 'Ventas (₡)', 'Líneas de la orden': 'Producto'})
        fig1.update_layout(xaxis_tickangle=45, dragmode='zoom', hovermode='x unified')
        st.plotly_chart(fig1, use_container_width=True)
        fig2 = px.line(x=pd.to_datetime(daily['Fecha']), y=daily['Total'], labels={'x': 'Fecha', 'y': 'Ventas (₡)'}, title="Tendencia Diaria de Ventas")
        fig2.update_layout(dragmode='zoom', hovermode='x unified')
        st.plotly_chart(fig2, use_container_width=True)
        grp = filtered_df.groupby('Cliente/Nombre principal')['Total'].sum().reset_index()
        fig3 = px.pie(grp, names='Cliente/Nombre principal', values='Total', title="Ventas por Grupo de Clientes")
        fig3.update_layout(dragmode='zoom', hovermode='x unified')
        st.plotly_chart(fig3, use_container_width=True)

    # Tab 6: Exportar Resumen de Métricas
    with tab6:
        st.header("Exportar Resumen de Métricas")
        most_sold = filtered_df.groupby('Líneas de la orden')['Total'].sum().idxmax() if not filtered_df.empty else "N/A"
        least_sold = filtered_df.groupby('Líneas de la orden')['Total'].sum().idxmin() if not filtered_df.empty else "N/A"
        report = {
            "Número de Órdenes": total_orders,
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
            st.download_button("Descargar Resumen (CSV)", data=report_df.to_csv(index=False).encode('utf-8'), file_name="resumen_ventas_aseavna.csv", mime="text/csv")
        with c2:
            buf_xl3 = generate_excel(report_df, "Resumen")
            st.download_button("Descargar Resumen (Excel)", data=buf_xl3, file_name="resumen_ventas_aseavna.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c3:
            buf_pdf3 = generate_pdf(report_df, "Resumen de Ventas - ASEAVNA", "resumen_ventas_aseavna.pdf")
            st.download_button("Descargar Resumen (PDF)", data=buf_pdf3, file_name="resumen_ventas_aseavna.pdf", mime="application/pdf")

        # Exportación automática programada
        if st.button("Programar Exportación Diaria"):
            output_dir = "app/reports"
            Path(output_dir).mkdir(parents=True, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            with open(f"{output_dir}/resumen_ventas_{timestamp}.csv", "wb") as f:
                f.write(report_df.to_csv(index=False).encode('utf-8'))
            with open(f"{output_dir}/resumen_ventas_{timestamp}.xlsx", "wb") as f:
                buf_xl3.seek(0)
                f.write(buf_xl3.read())
            with open(f"{output_dir}/resumen_ventas_{timestamp}.pdf", "wb") as f:
                buf_pdf3.seek(0)
                f.write(buf_pdf3.read())
            st.success(f"Reportes guardados en {output_dir} con timestamp {timestamp}")

    # Tab 7: Datos Crudos
    with tab7:
        st.header("Datos Crudos")
        gb = GridOptionsBuilder.from_dataframe(filtered_df)
        gb.configure_pagination(paginationAutoPageSize=True)
        gb.configure_side_bar()
        grid_options = gb.build()
        AgGrid(filtered_df, gridOptions=grid_options, height=400, width='100%')

    # Pie de página
    st.markdown("---")
    st.markdown("Desarrollado por Wilfredos para ASEAVNA | Fuente de Datos: Órdenes del Punto de Venta (POS) | 2025")