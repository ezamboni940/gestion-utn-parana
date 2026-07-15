import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import zipfile
import openpyxl
from PIL import Image
import textwrap

# --- 1. CONFIGURACIÓN GENERAL ---
st.set_page_config(page_title="Sistema de Gestión - UTN Paraná", layout="wide", page_icon="🏛️")
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .stMetric { background-color: white; padding: 15px; border-radius: 10px; border-left: 5px solid #005088; }
    .sidebar .sidebar-content { background-color: #f1f5f9; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. MENÚ LATERAL (SIDEBAR) ---
st.sidebar.title("🏛️ UTN Paraná")
st.sidebar.subheader("Sistema Integrado")

menu = st.sidebar.radio(
    "Navegación del Tablero",
    [
        "📊 Presupuesto General",
        "💸 Flujo de Fondos",
        "🔌 Servicios y Otros",
        "📜 Resoluciones Rectorado",
        "💼 Producidos Propios",
        "📱 Presentación Redes"
    ]
)

st.sidebar.divider()
st.sidebar.header("Carga de Datos")
uploaded_file = st.sidebar.file_uploader("Subir Archivo Excel", type="xlsx")

# --- 3. LÓGICA DE LAS SECCIONES ---

if uploaded_file is not None:
    
    # ==========================================
    # SECCIÓN 1: PRESUPUESTO GENERAL
    # ==========================================
    if menu == "📊 Presupuesto General":
        st.title("📊 Tablero de Control - Presupuesto General")
        try:
            xls = pd.ExcelFile(uploaded_file)
            hojas_disponibles = xls.sheet_names
            nombre_hoja_presupuesto = next((h for h in hojas_disponibles if h.strip().lower() in ['presupuesto', 'presupuesto general']), None)
            
            if nombre_hoja_presupuesto is None:
                st.error(f"❌ No se encontró la hoja de Presupuesto. Las hojas actuales son: {', '.join(hojas_disponibles)}. Asegúrese de que la hoja se llame 'Presupuesto'.")
            else:
                uploaded_file.seek(0)
                df = pd.read_excel(uploaded_file, sheet_name=nombre_hoja_presupuesto, skiprows=5)
                df.columns = df.columns.str.strip()
                
                if 'Imputación Presupuestaria' not in df.columns:
                    st.warning("El archivo subido no parece tener la estructura de Presupuesto.")
                else:
                    df['Imputación Presupuestaria'] = df['Imputación Presupuestaria'].astype(str).str.replace("'", "")
                    df['Fuente'] = df['Imputación Presupuestaria'].str.split('.').str[0]
                    df['Inciso'] = df['Imputación Presupuestaria'].str.split('.').str[6]
                    
                    mapeo_fuentes = {'11': 'Tesoro Nacional', '12': 'Producido Propio', '16': 'Tesoro Nacional'}
                    df['Fuente_Nombre'] = df['Fuente'].map(mapeo_fuentes).fillna('Otras')

                    fuentes_disponibles = df['Fuente_Nombre'].dropna().unique()
                    incisos_disponibles = sorted(df['Inciso'].dropna().unique())
                    
                    c_filt1, c_filt2 = st.columns(2)
                    with c_filt1: fuente_sel = st.multiselect("Fuente de Financiamiento", options=fuentes_disponibles, default=fuentes_disponibles)
                    with c_filt2: inciso_sel = st.multiselect("Filtrar por Inciso", options=incisos_disponibles, default=incisos_disponibles)

                    mask = (df['Fuente_Nombre'].isin(fuente_sel)) & (df['Inciso'].isin(inciso_sel))
                    df_filtrado = df[mask]

                    c1, c2, c3, c4, c5, c6 = st.columns(6)
                    credito_total = df_filtrado['Crédito'].sum()
                    preventivo_total = df_filtrado['Preventivos'].sum()
                    compromiso_total = df_filtrado['Compromisos'].sum()
                    devengado_total = df_filtrado['Devengados'].sum()
                    saldo_total = df_filtrado['Saldo'].sum()
                    porcentaje_ejecucion = (devengado_total / credito_total * 100) if credito_total > 0 else 0.0
                    
                    c1.metric("Crédito Total", f"$ {credito_total:,.0f}")
                    c2.metric("Preventivo", f"$ {preventivo_total:,.0f}")
                    c3.metric("Compromiso", f"$ {compromiso_total:,.0f}")
                    c4.metric("Devengado", f"$ {devengado_total:,.0f}")
                    c5.metric("Saldo Disp.", f"$ {saldo_total:,.0f}")
                    c6.metric("% Ejecución", f"{porcentaje_ejecucion:.1f}%")

                    st.divider()
                    col_left, col_right = st.columns(2)
                    with col_left:
                        st.subheader("Ejecución por Inciso")
                        if not df_filtrado.empty:
                            fig_inciso = px.bar(df_filtrado.groupby('Inciso')['Devengados'].sum().reset_index(), x='Inciso', y='Devengados', color_discrete_sequence=['#005088'])
                            st.plotly_chart(fig_inciso, use_container_width=True)
                    with col_right:
                        st.subheader("Distribución por Fuente")
                        if not df_filtrado.empty and devengado_total > 0:
                            fig_fuente = px.pie(df_filtrado, values='Devengados', names='Fuente_Nombre', color_discrete_sequence=['#005088', '#0ea5e9', '#1e293b'])
                            st.plotly_chart(fig_fuente, use_container_width=True)

                    st.divider()
                    st.subheader("Detalle Presupuestario y Alertas")
                    columnas_ideales = ['Nombre Dep', 'Descripción', 'Inciso', 'Fuente_Nombre', 'Crédito', 'Preventivos', 'Compromisos', 'Devengados', 'Saldo']
                    columnas_mostrar = [col for col in columnas_ideales if col in df.columns]
                    
                    def aplicar_alertas(row):
                        estilos = [''] * len(row)
                        if 'Crédito' in row and 'Saldo' in row and row['Crédito'] > 0:
                            porc_saldo = row['Saldo'] / row['Crédito']
                            idx_saldo = row.index.get_loc('Saldo')
                            if porc_saldo <= 0.10: estilos[idx_saldo] = 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
                            elif porc_saldo <= 0.25: estilos[idx_saldo] = 'background-color: #fef9c3; color: #854d0e; font-weight: bold;'
                        return estilos

                    st.dataframe(df_filtrado[columnas_mostrar].style.apply(aplicar_alertas, axis=1), use_container_width=True)

        except Exception as e:
            st.error(f"Error procesando Presupuesto: {e}")

    # ==========================================
    # SECCIÓN 2: FLUJO DE FONDOS
    # ==========================================
    elif menu == "💸 Flujo de Fondos":
        st.title("💸 Tablero de Flujo de Fondos Proyectado")
        try:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, sheet_name='Análisis Proyectado', header=None)
            meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
            
            df_egresos = df.iloc[5:48, [1, 2] + list(range(3, 15))].copy()
            df_egresos.columns = ['Concepto', 'Remanente 2025'] + meses
            df_egresos['Tipo'] = 'Egreso'
            
            df_ingresos = df.iloc[55:68, [1, 2] + list(range(3, 15))].copy()
            df_ingresos.columns = ['Concepto', 'Remanente 2025'] + meses
            df_ingresos['Tipo'] = 'Ingreso'
            
            df_flujo = pd.concat([df_egresos, df_ingresos], ignore_index=True)
            df_flujo = df_flujo.dropna(subset=['Concepto'])
            
            columnas_numericas = ['Remanente 2025'] + meses
            df_flujo[columnas_numericas] = df_flujo[columnas_numericas].fillna(0).apply(pd.to_numeric, errors='coerce').fillna(0)
            df_flujo['Total General'] = df_flujo[columnas_numericas].sum(axis=1)
            df_flujo = df_flujo[df_flujo['Total General'] > 0] 
            
            try: saldos = df[df[1] == 'Saldo'].iloc[0, 3:15].fillna(0).apply(pd.to_numeric, errors='coerce').values
            except: saldos = [0]*12

            remanente_total = df_flujo['Remanente 2025'].sum()
            total_ingresos = df_flujo[df_flujo['Tipo'] == 'Ingreso']['Total General'].sum()
            total_egresos = df_flujo[df_flujo['Tipo'] == 'Egreso']['Total General'].sum()
            saldo_final = saldos[-1] if len(saldos) > 0 and saldos[-1] != 0 else total_ingresos - total_egresos

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Remanente 2025", f"$ {remanente_total:,.0f}")
            c2.metric("Ingresos (Inc. Remanente)", f"$ {total_ingresos:,.0f}")
            c3.metric("Egresos Proyectados", f"$ {total_egresos:,.0f}")
            c4.metric("Saldo Caja a Diciembre", f"$ {saldo_final:,.0f}", delta="Cierre de Ejercicio", delta_color="normal" if saldo_final > 0 else "inverse")

            st.divider()
            col_g1, col_g2 = st.columns(2)
            totales_mensuales = df_flujo.groupby('Tipo')[meses].sum().T
            totales_mensuales.reset_index(inplace=True)
            totales_mensuales.columns = ['Mes', 'Egreso', 'Ingreso']
            
            with col_g1:
                st.subheader("Ingresos vs Egresos Mensuales")
                fig_bar = go.Figure()
                fig_bar.add_trace(go.Bar(x=totales_mensuales['Mes'], y=totales_mensuales['Ingreso'], name='Ingresos', marker_color='#0ea5e9'))
                fig_bar.add_trace(go.Bar(x=totales_mensuales['Mes'], y=totales_mensuales['Egreso'], name='Egresos', marker_color='#991b1b'))
                fig_bar.update_layout(barmode='group')
                st.plotly_chart(fig_bar, use_container_width=True)

            with col_g2:
                st.subheader("Evolución del Saldo de Caja")
                fig_line = px.line(x=meses, y=saldos, markers=True, color_discrete_sequence=['#005088'])
                fig_line.add_hline(y=0, line_dash="dash", line_color="red")
                st.plotly_chart(fig_line, use_container_width=True)

            st.divider()
            st.subheader("Resultado Neto Mensual (Superávit vs. Déficit)")
            totales_mensuales['Neto'] = totales_mensuales['Ingreso'] - totales_mensuales['Egreso']
            totales_mensuales['Color'] = totales_mensuales['Neto'].apply(lambda x: '#16a34a' if x >= 0 else '#dc2626')
            
            fig_neto = go.Figure(go.Bar(
                x=totales_mensuales['Mes'], y=totales_mensuales['Neto'], marker_color=totales_mensuales['Color'],
                text=totales_mensuales['Neto'], texttemplate='$ %{text:,.0f}', textposition='outside', cliponaxis=False
            ))
            fig_neto.update_layout(margin=dict(t=40, b=40))
            fig_neto.add_hline(y=0, line_width=2, line_color="black")
            st.plotly_chart(fig_neto, use_container_width=True)

            st.divider()
            st.subheader("🛒 Análisis de Compras Directas Simplificadas (CDS)")
            df_cds = df_flujo[(df_flujo['Tipo'] == 'Egreso') & (df_flujo['Concepto'].str.contains('CDS', case=False, na=False))]
            if not df_cds.empty:
                total_cds = df_cds['Total General'].sum()
                porcentaje_cds = (total_cds / total_egresos) * 100 if total_egresos > 0 else 0
                st.markdown(f"**Total proyectado en CDS:** $ {total_cds:,.0f} (*Representa el {porcentaje_cds:.1f}% de los egresos proyectados*)")
                df_cds_sorted = df_cds.sort_values(by='Total General', ascending=True)
                fig_cds = px.bar(df_cds_sorted, x='Total General', y='Concepto', orientation='h', color_discrete_sequence=['#f59e0b'], text='Total General')
                fig_cds.update_traces(texttemplate='$ %{text:,.0f}', textposition='outside', cliponaxis=False)
                fig_cds.update_layout(xaxis_title="Monto Anual Asignado ($)", yaxis_title="", margin=dict(r=50))
                st.plotly_chart(fig_cds, use_container_width=True)
            else:
                st.info("No se detectaron Compras Directas Simplificadas (CDS) en los registros de egresos.")

        except Exception as e:
            st.error(f"Error procesando Flujo de Fondos. Detalle: {e}")

    # ==========================================
    # SECCIÓN 3: SERVICIOS Y OTROS
    # ==========================================
    elif menu == "🔌 Servicios y Otros":
        st.title("🔌 Tablero de Egresos por Servicios")
        try:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, sheet_name='Servicios y otros', header=None)
            meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
            idx_total = df[df[1] == 'Total'].index[0]
            df_servicios = df.iloc[6:idx_total, [1] + list(range(2, 14))].copy()
            df_servicios.columns = ['Servicio'] + meses
            df_servicios[meses] = df_servicios[meses].fillna(0).apply(pd.to_numeric, errors='coerce').fillna(0)
            df_servicios['Total Anual'] = df_servicios[meses].sum(axis=1)
            df_servicios = df_servicios[df_servicios['Total Anual'] > 0].reset_index(drop=True)
            
            gasto_total = df_servicios['Total Anual'].sum()
            df_mensual = df_servicios[meses].sum()
            promedio_mensual = df_mensual[df_mensual > 0].mean() if (df_mensual > 0).sum() > 0 else 0

            c1, c2, c3 = st.columns(3)
            c1.metric("Gasto Total Proyectado", f"$ {gasto_total:,.0f}")
            c2.metric("Promedio Mensual (Activo)", f"$ {promedio_mensual:,.0f}")
            c3.metric("Mayor Gasto", df_servicios.loc[df_servicios['Total Anual'].idxmax()]['Servicio'])

            st.divider()
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.subheader("Gasto Mensual Total")
                fig_bar_mes = px.bar(x=meses, y=df_mensual.values, labels={'x': 'Mes', 'y': 'Gasto Total ($)'}, text_auto='.2s', color_discrete_sequence=['#0ea5e9'])
                fig_bar_mes.update_traces(textposition='outside')
                st.plotly_chart(fig_bar_mes, use_container_width=True)

            with col_g2:
                st.subheader("Distribución Anual")
                fig_pie = px.pie(df_servicios, values='Total Anual', names='Servicio', hole=0.4, color_discrete_sequence=px.colors.sequential.Blues_r)
                fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_pie, use_container_width=True)

            st.divider()
            st.subheader("Composición Mensual del Gasto (Barras Apiladas)")
            df_long = df_servicios.melt(id_vars=['Servicio'], value_vars=meses, var_name='Mes', value_name='Monto')
            df_long = df_long[df_long['Monto'] > 0] 
            fig_stack = px.bar(df_long, x='Mes', y='Monto', color='Servicio', labels={'Monto': 'Monto Gastado ($)', 'Mes': 'Meses'}, color_discrete_sequence=px.colors.qualitative.Prism)
            fig_stack.update_layout(barmode='stack')
            st.plotly_chart(fig_stack, use_container_width=True)

        except Exception as e:
            st.error(f"Error procesando Servicios. Detalle: {e}")

    # ==========================================
    # SECCIÓN 4: RESOLUCIONES RECTORADO
    # ==========================================
    elif menu == "📜 Resoluciones Rectorado":
        st.title("📜 Tablero de Créditos por Resoluciones")
        try:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, sheet_name='resoluciones de presupesto', skiprows=6)
            df_res = df.iloc[:, 1:9].copy()
            incisos = ['Inciso 2', 'Inciso 3', 'Inciso 4.3', 'Inciso 5.1']
            df_res.columns = ['Fecha'] + incisos + ['Total', 'Resolución', 'Tipo']
            df_res = df_res.dropna(subset=['Fecha']).copy()
            df_res[incisos + ['Total']] = df_res[incisos + ['Total']].fillna(0).apply(pd.to_numeric, errors='coerce').fillna(0)
            df_res['Total Calculado'] = df_res[incisos].sum(axis=1)
            df_res['Resolución'] = df_res['Resolución'].fillna('S/N').astype(str)
            df_res['Fecha'] = pd.to_datetime(df_res['Fecha'], errors='coerce')
            df_res['Mes_Orden'] = df_res['Fecha'].dt.strftime('%Y-%m')
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Crédito Total Recibido", f"$ {df_res['Total Calculado'].sum():,.0f}")
            c2.metric("Resoluciones Emitidas", str(len(df_res)))
            c3.metric("Inciso de Mayor Peso", df_res[incisos].sum().idxmax())

            st.divider()
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.subheader("Distribución por Inciso")
                fig_pie = px.pie(names=df_res[incisos].sum().index, values=df_res[incisos].sum().values, hole=0.4, color_discrete_sequence=['#005088', '#0ea5e9', '#38bdf8', '#7dd3fc'])
                st.plotly_chart(fig_pie, use_container_width=True)
            with col_g2:
                st.subheader("Ingresos por Fecha")
                df_mes = df_res.groupby('Mes_Orden')['Total Calculado'].sum().reset_index().sort_values('Mes_Orden')
                fig_bar_tiempo = px.bar(df_mes, x='Mes_Orden', y='Total Calculado', color_discrete_sequence=['#0ea5e9'])
                st.plotly_chart(fig_bar_tiempo, use_container_width=True)

        except Exception as e:
            st.error(f"Error procesando Resoluciones. Detalle: {e}")

    # ==========================================
    # SECCIÓN 5: PRODUCIDOS PROPIOS
    # ==========================================
    elif menu == "💼 Producidos Propios":
        st.title("💼 Tablero de Producidos Propios")
        try:
            uploaded_file.seek(0)
            df_pp = pd.read_excel(uploaded_file, sheet_name='Producidos Propios', header=None)
            df_pp = df_pp.dropna(how='all').reset_index(drop=True)
            trimesters = {}
            for col in range(1, len(df_pp.columns), 3):
                trim_name = df_pp.iloc[0, col]
                if pd.notna(trim_name) and "Trimestre" in str(trim_name): trimesters[trim_name] = [col, col+1, col+2]
            data = []
            for idx in range(2, len(df_pp)):
                area = df_pp.iloc[idx, 0]
                if pd.notna(area) and area != "Ingresos": 
                    for trim, cols in trimesters.items():
                        data.append({
                            'Área': str(area).strip(), 'Trimestre': str(trim).strip(),
                            'Ingresos': pd.to_numeric(df_pp.iloc[idx, cols[0]], errors='coerce') if pd.notna(df_pp.iloc[idx, cols[0]]) else 0,
                            'Egresos': pd.to_numeric(df_pp.iloc[idx, cols[1]], errors='coerce') if pd.notna(df_pp.iloc[idx, cols[1]]) else 0,
                            'Saldo': pd.to_numeric(df_pp.iloc[idx, cols[2]], errors='coerce') if pd.notna(df_pp.iloc[idx, cols[2]]) else 0
                        })
            df_clean = pd.DataFrame(data)
            df_clean['Saldo Calculado'] = df_clean['Ingresos'] - df_clean['Egresos']
            total_ingresos = df_clean['Ingresos'].sum()

            c1, c2, c3 = st.columns(3)
            c1.metric("Total Ingresos Generados", f"$ {total_ingresos:,.0f}")
            c2.metric("Total Gastos (Egresos)", f"$ {df_clean['Egresos'].sum():,.0f}")
            c3.metric("Saldo Neto (Rentabilidad)", f"$ {total_ingresos - df_clean['Egresos'].sum():,.0f}")

            st.divider()
            df_area = df_clean.groupby('Área')[['Ingresos', 'Egresos']].sum().reset_index()
            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(x=df_area['Área'], y=df_area['Ingresos'], name='Ingresos', marker_color='#0ea5e9', text=df_area['Ingresos'], texttemplate='$ %{text:,.0f}', textposition='outside', cliponaxis=False))
            fig_bar.add_trace(go.Bar(x=df_area['Área'], y=df_area['Egresos'], name='Egresos', marker_color='#991b1b', text=df_area['Egresos'], texttemplate='$ %{text:,.0f}', textposition='outside', cliponaxis=False))
            fig_bar.update_layout(barmode='group', margin=dict(t=40))
            st.plotly_chart(fig_bar, use_container_width=True)

            st.divider()
            df_trim = df_clean.groupby('Trimestre')[['Ingresos', 'Egresos', 'Saldo Calculado']].sum().reset_index()
            orden_dict = {'Primer Trimestre': 1, 'Segundo Trimestre': 2, 'Tercer Trimestre': 3, 'Cuarto Trimestre': 4}
            df_trim['Orden'] = df_trim['Trimestre'].map(orden_dict).fillna(5)
            df_trim = df_trim.sort_values('Orden')
            if not df_trim.empty and total_ingresos > 0:
                df_trim['Color'] = df_trim['Saldo Calculado'].apply(lambda x: '#16a34a' if x >= 0 else '#dc2626')
                fig_trim = go.Figure(go.Bar(x=df_trim['Trimestre'].astype(str).tolist(), y=df_trim['Saldo Calculado'], marker_color=df_trim['Color'], text=df_trim['Saldo Calculado'], texttemplate='$ %{text:,.0f}', textposition='outside', cliponaxis=False))
                fig_trim.update_layout(margin=dict(t=40, b=40))
                st.plotly_chart(fig_trim, use_container_width=True)

        except Exception as e:
            st.error(f"Error procesando Producidos Propios. Detalle: {e}")

    # ==========================================
    # SECCIÓN 6: PRESENTACIÓN PARA REDES Y DESCARGA
    # ==========================================
    elif menu == "📱 Presentación Redes":
        st.title("📱 Contenido Gráfico para Redes Sociales")
        st.write("Generador de carrusel (5 filminas) destacando eficiencia, autogestión e inversión. Formato 1080x1080px.")
        
        try:
            # --- EXTRACCIÓN DEL LOGO DEL EXCEL ---
            uploaded_file.seek(0)
            logo_img = None
            try:
                wb = openpyxl.load_workbook(uploaded_file, data_only=True)
                for sheetname in wb.sheetnames:
                    sheet = wb[sheetname]
                    if hasattr(sheet, '_images') and len(sheet._images) > 0:
                        img_data = sheet._images[0]._data()
                        logo_img = Image.open(io.BytesIO(img_data))
                        break
            except:
                pass 
            
            # --- PROCESAMIENTO PRESUPUESTO GENERAL ---
            uploaded_file.seek(0)
            xls = pd.ExcelFile(uploaded_file)
            hojas_disponibles = xls.sheet_names
            nombre_hoja_presupuesto = next((h for h in hojas_disponibles if h.strip().lower() in ['presupuesto', 'presupuesto general']), None)
            
            df_pres = pd.read_excel(uploaded_file, sheet_name=nombre_hoja_presupuesto, skiprows=5)
            df_pres.columns = df_pres.columns.str.strip()
            df_pres['Imputación Presupuestaria'] = df_pres['Imputación Presupuestaria'].astype(str).str.replace("'", "")
            df_pres['Fuente'] = df_pres['Imputación Presupuestaria'].str.split('.').str[0]
            mapeo_fuentes = {'11': 'Tesoro Nacional', '12': 'Producido Propio', '16': 'Tesoro Nacional'}
            df_pres['Fuente_Nombre'] = df_pres['Fuente'].map(mapeo_fuentes).fillna('Otras')

            credito_total = df_pres['Crédito'].sum()
            devengado_total = df_pres['Devengados'].sum()
            porcentaje_ejecucion = (devengado_total / credito_total * 100) if credito_total > 0 else 0.0
            
            fuentes_agrupadas = df_pres.groupby('Fuente_Nombre')['Crédito'].sum()
            producido_propio = fuentes_agrupadas.get('Producido Propio', 0)
            tesoro_nacional = fuentes_agrupadas.get('Tesoro Nacional', 0)
            top_3_inversiones = df_pres.groupby('Descripción')['Devengados'].sum().sort_values(ascending=True).tail(3).reset_index()

            # --- PROCESAMIENTO FLUJO DE FONDOS ---
            uploaded_file.seek(0)
            df_ff = pd.read_excel(uploaded_file, sheet_name='Análisis Proyectado', header=None)
            meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
            df_ff_egresos = df_ff.iloc[5:48, [1, 2] + list(range(3, 15))]
            df_ff_ingresos = df_ff.iloc[55:68, [1, 2] + list(range(3, 15))]
            
            cols_num = [2] + list(range(3, 15))
            total_ingresos_ff = df_ff_ingresos[cols_num].fillna(0).apply(pd.to_numeric, errors='coerce').sum().sum()
            total_egresos_ff = df_ff_egresos[cols_num].fillna(0).apply(pd.to_numeric, errors='coerce').sum().sum()

            # --- PROCESAMIENTO PRODUCIDOS PROPIOS ---
            uploaded_file.seek(0)
            df_pp_raw = pd.read_excel(uploaded_file, sheet_name='Producidos Propios', header=None)
            df_pp_raw = df_pp_raw.dropna(how='all').reset_index(drop=True)
            trimesters = {}
            for col in range(1, len(df_pp_raw.columns), 3):
                trim_name = df_pp_raw.iloc[0, col]
                if pd.notna(trim_name) and "Trimestre" in str(trim_name): trimesters[trim_name] = [col, col+1, col+2]
            
            data_pp = []
            for idx in range(2, len(df_pp_raw)):
                area = df_pp_raw.iloc[idx, 0]
                if pd.notna(area) and area != "Ingresos": 
                    for trim, cols in trimesters.items():
                        data_pp.append({
                            'Área': str(area).strip(),
                            'Ingresos': pd.to_numeric(df_pp_raw.iloc[idx, cols[0]], errors='coerce') if pd.notna(df_pp_raw.iloc[idx, cols[0]]) else 0,
                            'Egresos': pd.to_numeric(df_pp_raw.iloc[idx, cols[1]], errors='coerce') if pd.notna(df_pp_raw.iloc[idx, cols[1]]) else 0
                        })
            df_pp_clean = pd.DataFrame(data_pp)
            df_pp_clean['Saldo'] = df_pp_clean['Ingresos'] - df_pp_clean['Egresos']
            df_area_saldo = df_pp_clean.groupby('Área')['Saldo'].sum().reset_index()
            df_area_saldo = df_area_saldo[df_area_saldo['Saldo'] > 0].sort_values(by='Saldo', ascending=True)

            # --- CONFIGURACIÓN DEL LOGO ---
            if logo_img:
                logo_dict = dict(source=logo_img, xref="paper", yref="paper", x=0.5, y=1.15, sizex=0.22, sizey=0.22, xanchor="center", yanchor="bottom")
            else:
                logo_dict = dict(source="https://upload.wikimedia.org/wikipedia/commons/thumb/6/67/UTN_logo.jpg/320px-UTN_logo.jpg", xref="paper", yref="paper", x=0.5, y=1.15, sizex=0.22, sizey=0.22, xanchor="center", yanchor="bottom")

            st.divider()
            
            # --- FILA 1 ---
            col1, col2 = st.columns([1, 1])

            fig1 = go.Figure(go.Indicator(
                mode="gauge+number", value=porcentaje_ejecucion,
                number={'suffix': "%", 'font': {'size': 80, 'color': '#005088'}},
                title={'text': f"Presupuesto Invertido<br><span style='font-size:0.6em;color:gray'>Sobre un total de $ {credito_total/1000000:.1f} Millones</span>", 'font': {'size': 30}},
                gauge={'axis': {'range': [None, 100], 'tickwidth': 1, 'tickcolor': "darkblue"}, 'bar': {'color': "#0ea5e9"}, 'bgcolor': "white", 'borderwidth': 2, 'bordercolor': "gray"}
            ))
            fig1.update_layout(width=700, height=700, font=dict(family="Arial", size=18), margin=dict(t=160))
            fig1.add_layout_image(logo_dict)
            with col1:
                st.subheader("Filmina 1: Ejecución Total")
                st.plotly_chart(fig1, use_container_width=False)

            fig2 = px.pie(
                names=['Autogestión (Producido Propio)', 'Aporte Estatal (Tesoro)'],
                values=[producido_propio, tesoro_nacional], hole=0.5, color_discrete_sequence=['#0ea5e9', '#005088']
            )
            fig2.update_layout(
                title={'text': "El Valor de nuestra Autogestión", 'font': {'size': 26}, 'x': 0.5, 'y': 0.95},
                width=700, height=700, font=dict(family="Arial", size=20),
                legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5),
                annotations=[dict(text='Fuentes de<br>Ingreso', x=0.5, y=0.5, font_size=20, showarrow=False)], margin=dict(t=160)
            )
            fig2.update_traces(textposition='inside', textinfo='percent')
            fig2.add_layout_image(logo_dict)
            with col2:
                st.subheader("Filmina 2: Autogestión")
                st.plotly_chart(fig2, use_container_width=False)

            # --- FILA 2 ---
            st.divider()
            col3, col4 = st.columns([1, 1])
            
            # --- FILMINA 3: CORRECCIÓN DE MÁRGENES Y TEXTO ---
            # Aplicamos "textwrap" para convertir oraciones largas en párrafos de múltiples líneas (separados por <br>)
            top_3_inversiones['Desc_Wrap'] = top_3_inversiones['Descripción'].apply(lambda x: '<br>'.join(textwrap.wrap(str(x), width=25)))
            
            fig3 = px.bar(
                top_3_inversiones, x='Devengados', y='Desc_Wrap', orientation='h', text='Devengados', color_discrete_sequence=['#005088']
            )
            fig3.update_traces(texttemplate='$ %{text:,.0f}', textposition='outside', textfont_size=20, cliponaxis=False)
            fig3.update_layout(
                title={'text': "Top 3: Inversión que Transforma", 'font': {'size': 26}, 'x': 0.5, 'y': 0.95},
                width=700, height=700, font=dict(family="Arial", size=16),
                yaxis=dict(title='', tickfont=dict(size=15)), xaxis={'visible': False}, 
                margin=dict(l=280, r=150, t=160, b=50) # l=280 asegura que haya suficiente espacio izquierdo
            )
            fig3.add_layout_image(logo_dict)
            with col3:
                st.subheader("Filmina 3: Inversión Real")
                st.plotly_chart(fig3, use_container_width=False)
            
            fig4 = go.Figure()
            fig4.add_trace(go.Bar(name='Ingresos', x=['Proyección Anual'], y=[total_ingresos_ff], marker_color='#16a34a', text=[total_ingresos_ff], texttemplate='$ %{text:,.0f}', textposition='auto'))
            fig4.add_trace(go.Bar(name='Egresos', x=['Proyección Anual'], y=[total_egresos_ff], marker_color='#dc2626', text=[total_egresos_ff], texttemplate='$ %{text:,.0f}', textposition='auto'))
            fig4.update_layout(
                title={'text': "Equilibrio y Solidez Financiera", 'font': {'size': 26}, 'x': 0.5, 'y': 0.95},
                width=700, height=700, font=dict(family="Arial", size=20),
                barmode='group', yaxis={'visible': False}, margin=dict(l=10, r=10, t=160),
                legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5)
            )
            fig4.add_layout_image(logo_dict)
            with col4:
                st.subheader("Filmina 4: Solidez Financiera")
                st.plotly_chart(fig4, use_container_width=False)

            # --- FILA 3 ---
            st.divider()
            col5, col6 = st.columns([1, 1])

            # --- FILMINA 5: CORRECCIÓN DE MÁRGENES Y TEXTO ---
            # Aplicamos "textwrap" a las áreas también
            df_area_saldo['Área_Wrap'] = df_area_saldo['Área'].apply(lambda x: '<br>'.join(textwrap.wrap(str(x), width=25)))
            
            fig5 = px.bar(
                df_area_saldo, x='Saldo', y='Área_Wrap', orientation='h', text='Saldo', color_discrete_sequence=['#0ea5e9']
            )
            fig5.update_traces(texttemplate='$ %{text:,.0f}', textposition='outside', textfont_size=20, cliponaxis=False)
            fig5.update_layout(
                title={'text': "Áreas Impulsoras (Superávit Neto)", 'font': {'size': 26}, 'x': 0.5, 'y': 0.95},
                width=700, height=700, font=dict(family="Arial", size=18),
                yaxis=dict(title='', tickfont=dict(size=15)), xaxis={'visible': False}, 
                margin=dict(l=280, r=150, t=160, b=50) # l=280 asegura que haya suficiente espacio izquierdo
            )
            fig5.add_layout_image(logo_dict)
            with col5:
                st.subheader("Filmina 5: Motor de Crecimiento")
                st.plotly_chart(fig5, use_container_width=False)
            
            # --- ZONA DE DESCARGA Y COPY ---
            with col6:
                st.subheader("📦 Descargar Presentación Completa")
                st.write("Obtén las **5 imágenes en alta calidad (1080x1080px)**, con tu logo incrustado, listas para publicar en Instagram o Facebook.")
                
                try:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        zip_file.writestr("Filmina_1_Ejecucion.png", fig1.to_image(format="png", width=1080, height=1080))
                        zip_file.writestr("Filmina_2_Autogestion.png", fig2.to_image(format="png", width=1080, height=1080))
                        zip_file.writestr("Filmina_3_Inversiones.png", fig3.to_image(format="png", width=1080, height=1080))
                        zip_file.writestr("Filmina_4_Solidez.png", fig4.to_image(format="png", width=1080, height=1080))
                        zip_file.writestr("Filmina_5_Superavit.png", fig5.to_image(format="png", width=1080, height=1080))
                    
                    st.download_button(
                        label="📥 Descargar las 5 Imágenes (Archivo ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="Presentacion_Integral_Redes_UTN.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                except Exception as ex:
                    st.error("🚨 Falta la librería gráfica para exportar. Ejecuta `pip install -U kaleido` en tu terminal.")
                
                st.write("---")
                st.subheader("📝 Texto sugerido para la publicación (Copy)")
                st.info(f"""
                En la UTN Regional Paraná creemos que la educación pública se defiende con gestión responsable, cuentas claras y trabajo duro. ⚙️ Hoy queremos compartir con nuestra comunidad un resumen de nuestra solidez financiera.
                
                No solo proyectamos cerrar el ciclo con nuestras cuentas en **equilibrio**, sino que nos enorgullece destacar que gran parte de nuestro presupuesto se genera gracias al esfuerzo emprendedor de la propia Facultad (Posgrados, Servicios a Terceros, UVT). ¡Esto nos permite seguir mejorando las aulas y laboratorios para vos! 
                
                Deslizá para conocer el trabajo que hay detrás de nuestra facultad ➡️ 
                
                #UTN #UTNParaná #GestiónUniversitaria #Transparencia #UniversidadPública #Educación
                """)

        except Exception as e:
            st.error(f"Error procesando los datos para las redes. Es posible que falte alguna pestaña en el Excel. Detalle: {e}")

else:
    st.info("👋 ¡Bienvenido al Sistema Integrado! Por favor, seleccione un módulo en el menú lateral y suba su archivo de Excel para comenzar.")