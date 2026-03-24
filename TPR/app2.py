import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from thefuzz import process
import re

# 1. CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(page_title="Inmovision - Dashboard Corporativo", layout="wide")

# --- INICIALIZACIÓN DE LINKS DINÁMICOS (SESSION STATE) ---
if 'url_ventas' not in st.session_state:
    st.session_state.url_ventas = "https://docs.google.com/spreadsheets/d/18WS22r1Fml5a9qW3fJOj40d0h7aldJVj/edit?usp=sharing"
if 'url_instalaciones' not in st.session_state:
    st.session_state.url_instalaciones = "https://docs.google.com/spreadsheets/d/1QqH8lGktix5YLV7seIL9cXUxWCaANauM/edit?usp=sharing"
if 'url_gestion' not in st.session_state:
    st.session_state.url_gestion = "https://docs.google.com/spreadsheets/d/1ZlKyiraV_jVoK5Eoq3ZBYjDsSm_6tx8x/edit?usp=sharing"

# 2. CSS UNIFICADO
st.markdown("""
    <style>
    [data-testid="bundle-lib-sidebar-close-icon"], 
    [data-testid="stSidebarCollapseButton"] { display: none !important; }
    .stApp { background-color: #0e1117 !important; }
    .asesor-header { color: #ffffff; font-size: 26px; font-weight: bold; margin-bottom: 20px; border-bottom: 2px solid #4CAF50; padding-bottom: 10px; }
    .main-header { color: #ffffff; font-size: 28px; font-weight: bold; padding-bottom: 10px; border-bottom: 3px solid #4CAF50; margin-bottom: 20px; }
    .card-container { display: flex; gap: 15px; margin-bottom: 25px; margin-top: 10px; }
    .card-total, .card-style, .card-image-style {
        background-color: #ffffff; border-radius: 12px; padding: 20px; text-align: center;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1); border: 1px solid #e1e4e8; margin-bottom: 15px; flex: 1;
    }
    .card-total { background-color: #f1f3f6; border-left: 5px solid #4CAF50; color: #1a1c24; }
    .total-label, .image-label { font-weight: 800; font-size: 13px; text-transform: uppercase; color: #7f8c8d; margin-bottom: 5px; }
    .total-amount, .image-value-number, .image-value-text { font-weight: 900; color: #1a1c24; font-size: 28px; }
    .border-green { border-left: 8px solid #2ecc71; }
    .border-blue { border-left: 8px solid #3498db; }
    .section-title { color: #ffffff; font-size: 18px; font-weight: bold; margin-top: 25px; margin-bottom: 15px; background: #34495e; padding: 8px 15px; border-radius: 5px; }
    [data-testid="stSidebar"] { background-color: #1a1c24; }
    .logo-sidebar { display: block; margin: auto; width: 150px; margin-bottom: 20px; }
    .detalle-titulo { color: #ffffff; font-size: 20px; font-weight: bold; margin-top: 30px; margin-bottom: 10px; }
    header, footer {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# 3. FUNCIONES DE APOYO
def limpiar_monto(valor):
    if pd.isna(valor): return 0.0
    texto = re.sub(r'[^\d,.]', '', str(valor))
    if ',' in texto and '.' not in texto: texto = texto.replace(',', '.')
    elif ',' in texto and '.' in texto: texto = texto.replace(',', '')
    try: return float(texto)
    except: return 0.0

def corregir_nombre(nombre_sucio, lista_maestra, umbral=90):
    if not nombre_sucio or pd.isna(nombre_sucio): return "DESCONOCIDO"
    nombre_sucio = str(nombre_sucio).strip().upper()
    if "MENDD" in nombre_sucio or "ANDRE M" in nombre_sucio: return "ANDREA MENDOZA"
    if "PACURUCO" in nombre_sucio: return "SUSANA PACURUCU"
    if "AYORA GLENDA" in nombre_sucio: return "GLENDA RAMOS AYORA"
    match, score = process.extractOne(nombre_sucio, lista_maestra)
    return match if score >= umbral else nombre_sucio

@st.cache_data(ttl=60)
def cargar_datos(url, nombre_hoja, limpiar_precios=False):
    try:
        file_id = url.split('/')[-2]
        export_url = f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'
        df = pd.read_excel(export_url, sheet_name=nombre_hoja)
        df.columns = [str(c).strip().upper() for c in df.columns]
        if limpiar_precios:
            for col in ['VALOR MENSUAL A PAGAR INCLUIDO IVA', 'VALOR MENSUAL A PAGAR SIN IVA']:
                if col in df.columns:
                    df[col] = df[col].apply(limpiar_monto)
        return df
    except: return None

# 4. RECURSOS
URL_LOGO = "https://lh4.googleusercontent.com/proxy/SeW7l23MFgElfFnJzA8WsomRRdBeiXYsMuQMdiB6_m4J0N0j7RGAB09PNGAO-uUPhKMPITGfAgagRh76fzbODUl3jU3utoz20hT2W99Q7BODxV-g" 

VENDEDORES_PERMITIDOS = [
    "ALEXANDRA REINO", "ANDREA MENDOZA", "CESAR VERA", "DIANA RIVERA", 
    "EDISON SACA", "FRANKLIN QUEZADA", "GLENDA RAMOS AYORA", "JENNIFER ATANCURI", 
    "JORGE GARCIA", "LAURA MORAN", "MANCHENO KARLA", "MARIA JOSE PEÑAFIEL", 
    "MELANY GUZHÑAY", "NANCY JARAMA", "PRISCILA RAMOS", "SILVIA YUNGA", 
    "STALIN ROJAS", "SUSANA PACURUCU", "VERONICA MALO", 
    "WILLIAM BRITO", "WILLIAN MOLINA"
]

PALABRAS_FILTRO = [
    'SUMATORIA', 'RESUMEN', 'TOTAL', 'DIARIA', 'MARZO', 'ABRIL', 'MAYO', 
    'JUNIO', 'TÉCNICO', 'TECNICO', 'PERSONAL NO COMERCIAL',"EMBAJADORA MAYRA MACHUCA", "EMBAJADORA NAYELI JUELA"
]

# --- CONFIGURACIÓN DE BARRA LATERAL ---
st.sidebar.markdown(f'<img src="{URL_LOGO}" class="logo-sidebar">', unsafe_allow_html=True)

# NUEVA FUNCIÓN: Cambio de Links
with st.sidebar.expander("⚙️ Configuración de Google Sheets"):
    st.session_state.url_ventas = st.text_input("Link Control Ventas:", st.session_state.url_ventas)
    st.session_state.url_instalaciones = st.text_input("Link Instalaciones:", st.session_state.url_instalaciones)
    st.session_state.url_gestion = st.text_input("Link Gestión Asesores:", st.session_state.url_gestion)
    if st.button("🔄 Actualizar y Recargar"):
        st.cache_data.clear()
        st.rerun()

st.sidebar.markdown("---")
seccion = st.sidebar.radio("Módulo:", ["📊 Control de Ventas", "🛠️ Informe de Instalaciones", "📈 Gestión de Asesores","📈 Gestión Diaria"])

# Carga inicial de meses (usando el link dinámico de Ventas)
try:
    file_id_v = st.session_state.url_ventas.split('/')[-2]
    xls_v = pd.ExcelFile(f'https://docs.google.com/spreadsheets/d/{file_id_v}/export?format=xlsx')
    mapa_meses = {str(h).lower(): h for h in xls_v.sheet_names if str(h).upper() not in ['VARIABLES', 'VENTAS']}
except: mapa_meses = {}

# ==========================================
# MÓDULO 1: CONTROL DE VENTAS
# ==========================================
if seccion == "📊 Control de Ventas":
    try:
        mes_sel_display = st.sidebar.selectbox("📅 Mes a consultar:", list(mapa_meses.keys()))
        df_ventas = cargar_datos(st.session_state.url_ventas, mapa_meses[mes_sel_display])
        df_inst_conteo = cargar_datos(st.session_state.url_instalaciones, "Instalaciones")

        if df_ventas is not None:
            col_vendedor = df_ventas.columns[0]
            vendedores_en_hoja = df_ventas[col_vendedor].dropna().unique()
            v_unicos = []
            for v in vendedores_en_hoja:
                v_str = str(v).upper()
                if not any(palabra in v_str for palabra in PALABRAS_FILTRO):
                    nombre_limpio = corregir_nombre(v, VENDEDORES_PERMITIDOS)
                    if nombre_limpio != "DESCONOCIDO":
                        v_unicos.append(nombre_limpio)
            v_unicos = sorted(list(set(v_unicos)))
            
            ver_todo = st.sidebar.checkbox("Ver Resumen General del Mes")
            vendedor_sel = st.sidebar.selectbox("👤 Seleccionar Asesor:", v_unicos)

            if ver_todo:
                st.markdown('<div class="asesor-header">📊 Resumen General de Ventas</div>', unsafe_allow_html=True)
                columnas_total = [c for c in df_ventas.columns if 'TOTAL' in str(c).upper() and c != col_vendedor]
                if columnas_total:
                    col_total_name = columnas_total[0]
                    resumen = df_ventas[[col_vendedor, col_total_name]].copy()
                    resumen.columns = ['Vendedor', 'Monto Total']
                    resumen['V_UPPER'] = resumen['Vendedor'].apply(lambda x: str(x).upper())
                    mask_basura = resumen['V_UPPER'].apply(lambda x: any(p in x for p in PALABRAS_FILTRO))
                    resumen = resumen[~mask_basura]
                    resumen['Vendedor'] = resumen['Vendedor'].apply(lambda x: corregir_nombre(x, VENDEDORES_PERMITIDOS))
                    resumen = resumen[resumen['Vendedor'] != "DESCONOCIDO"]
                    resumen['Monto Total'] = pd.to_numeric(resumen['Monto Total'], errors='coerce').fillna(0)
                    resumen = resumen[resumen['Monto Total'] > 0].sort_values(by='Monto Total', ascending=False)
                    
                    st.markdown(f'<div style="background-color:#1a1c24; border:1px solid #4CAF50; border-radius:12px; padding:25px; text-align:center; margin-bottom:20px;"><div class="total-label" style="color:#4CAF50">Venta Total Consolidada</div><div style="font-size:45px; font-weight:bold; color:#4CAF50;">${resumen["Monto Total"].sum():,.2f}</div></div>', unsafe_allow_html=True)
                    
                    c1, c2 = st.columns([1, 1])
                    with c1: st.dataframe(resumen[['Vendedor', 'Monto Total']], use_container_width=True, hide_index=True)
                    with c2:
                        fig_pie = px.pie(resumen, values='Monto Total', names='Vendedor', hole=0.4, title="Participación")
                        fig_pie.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color="white")
                        st.plotly_chart(fig_pie, use_container_width=True)

            elif vendedor_sel:
                st.markdown(f'<div class="asesor-header">👤 Asesor: {vendedor_sel}</div>', unsafe_allow_html=True)
                df_ventas[col_vendedor] = df_ventas[col_vendedor].apply(lambda x: corregir_nombre(x, VENDEDORES_PERMITIDOS))
                fila_v = df_ventas[df_ventas[col_vendedor] == vendedor_sel]
                cols_excluir = [col_vendedor] + [c for c in df_ventas.columns if 'TOTAL' in str(c).upper()]
                cols_datos = [c for c in df_ventas.columns if c not in cols_excluir]
                datos_fila = fila_v[cols_datos].T.reset_index().iloc[:, :2]
                datos_fila.columns = ['Fecha', 'Venta']
                datos_fila['Venta'] = pd.to_numeric(datos_fila['Venta'], errors='coerce').fillna(0)
                datos_fila['Fecha_DT'] = pd.to_datetime(datos_fila['Fecha'], errors='coerce')
                ventas_ok = datos_fila[datos_fila['Venta'] > 0].copy().sort_values('Fecha_DT')
                ventas_ok['Fecha_Limpia'] = ventas_ok['Fecha_DT'].dt.strftime('%d-%m-%Y')

                cantidad_comprobada = 0
                if df_inst_conteo is not None:
                    df_inst_conteo['FECHA'] = pd.to_datetime(df_inst_conteo['FECHA'], errors='coerce')
                    df_inst_conteo['VENDEDOR'] = df_inst_conteo['VENDEDOR'].apply(lambda x: corregir_nombre(x, VENDEDORES_PERMITIDOS))
                    col_comp = 'PRECIO DEL PLAN CON IVA' if 'PRECIO DEL PLAN CON IVA' in df_inst_conteo.columns else df_inst_conteo.columns[-2]
                    f_dt = datos_fila['Fecha_DT'].dropna()
                    if not f_dt.empty:
                        mask = (df_inst_conteo['VENDEDOR'] == vendedor_sel) & (df_inst_conteo['FECHA'] >= f_dt.min()) & (df_inst_conteo['FECHA'] <= f_dt.max()) & (pd.to_numeric(df_inst_conteo[col_comp], errors='coerce') > 0)
                        cantidad_comprobada = len(df_inst_conteo[mask])

                st.markdown(f'''<div class="card-container">
                        <div class="card-total"><div class="total-label">Monto Total Vendido</div><div class="total-amount">${ventas_ok["Venta"].sum():,.2f}</div></div>
                        <div class="card-total" style="border-left-color: #2196F3;"><div class="total-label">Ventas Comprobadas</div><div class="total-amount">{cantidad_comprobada}</div></div>
                    </div>''', unsafe_allow_html=True)
                
                if not ventas_ok.empty:
                    fig_vend = px.line(ventas_ok, x='Fecha_Limpia', y='Venta', title="Tendencia de Rendimiento", markers=True, color_discrete_sequence=['#00D4FF'])
                    fig_vend.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white")
                    st.plotly_chart(fig_vend, use_container_width=True)
                    st.markdown('<div class="detalle-titulo">📅 Detalle de Transacciones</div>', unsafe_allow_html=True)
                    st.dataframe(ventas_ok[['Fecha_Limpia', 'Venta']], use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Error en ventas: {e}")

# ==========================================
# MÓDULO 2: INFORME DE INSTALACIONES
# ==========================================
elif seccion == "🛠️ Informe de Instalaciones":
    st.markdown('<div class="asesor-header">🛠️ Control de Instalaciones y Cumplimiento</div>', unsafe_allow_html=True)
    if mapa_meses:
        mes_sel = st.sidebar.selectbox("📅 Seleccionar Periodo (Mes):", list(mapa_meses.keys()))
        df_inst = cargar_datos(st.session_state.url_instalaciones, "Instalaciones")
        df_ref_mes = cargar_datos(st.session_state.url_ventas, mapa_meses[mes_sel])
        if df_inst is not None and df_ref_mes is not None:
            cols_fechas = [c for c in df_ref_mes.columns if any(char.isdigit() for char in str(c))]
            f_dt = pd.to_datetime(cols_fechas, errors='coerce').dropna()
            if not f_dt.empty:
                f_min, f_max = f_dt.min(), f_dt.max()
                df_inst['FECHA'] = pd.to_datetime(df_inst['FECHA'], errors='coerce')
                df_inst['VENDEDOR'] = df_inst['VENDEDOR'].apply(lambda x: corregir_nombre(x, VENDEDORES_PERMITIDOS))
                df_por_fecha = df_inst[(df_inst['FECHA'] >= f_min) & (df_inst['FECHA'] <= f_max)].copy()
                vendedores_con_datos = sorted([v for v in df_por_fecha['VENDEDOR'].unique() if v in VENDEDORES_PERMITIDOS])
                vendedor_sel = st.sidebar.selectbox("👤 Seleccionar Vendedor:", ["TODOS"] + vendedores_con_datos)
                estados_totales = df_por_fecha['ESTADO'].unique().tolist()
                estados_sel = st.sidebar.multiselect("📋 Filtrar por Estado:", estados_totales, default=estados_totales)
                df_f = df_por_fecha[df_por_fecha['ESTADO'].isin(estados_sel)].copy()
                if vendedor_sel != "TODOS": df_f = df_f[df_f['VENDEDOR'] == vendedor_sel]
                
                if not df_f.empty:
                    st.markdown("### 📊 Resumen por Estado")
                    df_graf = df_f['ESTADO'].value_counts().reset_index()
                    df_graf.columns = ['ESTADO_REAL', 'CANTIDAD']
                    cols_cards = st.columns(len(df_graf))
                    for i, row in df_graf.iterrows():
                        nombre_est, valor_est = str(row['ESTADO_REAL']).upper(), row['CANTIDAD']
                        color_c = "border-green" if i == 0 else "border-blue"
                        with cols_cards[i]:
                            st.markdown(f'<div class="card-image-style {color_c}"><div class="image-label">ESTADO</div><div class="image-value-number" style="font-size: 18px; min-height: 40px; display: flex; align-items: center; justify-content: center;">{nombre_est}</div><div class="image-value-number" style="font-size: 48px; margin-top: 10px;">{valor_est}</div></div>', unsafe_allow_html=True)
                    fig = px.bar(df_graf, x='CANTIDAD', y='ESTADO_REAL', orientation='h', text='CANTIDAD', color='ESTADO_REAL')
                    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
                    if 'PRODUCTO' in df_f.columns:
                        prod_estrella = df_f['PRODUCTO'].value_counts().idxmax()
                        st.markdown(f'<div class="card-image-style border-green" style="margin-top: 25px; margin-bottom: 25px;"><div class="image-label">PRODUCTO ESTRELLA</div><div class="image-value-number" style="font-size: 45px;">{str(prod_estrella).upper()}</div></div>', unsafe_allow_html=True)
                    st.markdown("### 📋 Registros en este rango")
                    df_f_display = df_f.copy()
                    df_f_display['FECHA'] = df_f_display['FECHA'].dt.strftime('%Y-%m-%d')
                    st.dataframe(df_f_display[['FECHA', 'VENDEDOR', 'CLIENTE', 'PRODUCTO', 'ESTADO']], use_container_width=True, hide_index=True)

# ==========================================
# MÓDULO 3: GESTIÓN DE ASESORES
# ==========================================
elif seccion == "📈 Gestión de Asesores":
    df_v = cargar_datos(st.session_state.url_gestion, "Ventas", limpiar_precios=True)
    if df_v is not None:
        COL_MES, COL_ASESOR, COL_VALOR = 'MES COMERCIAL', 'ASESOR', 'VALOR MENSUAL A PAGAR SIN IVA'
        meses = sorted(df_v[COL_MES].dropna().unique().tolist(), reverse=True)
        mes_sel_gest = st.sidebar.selectbox("📅 Mes Comercial (Gestión):", meses)
        df_mes = df_v[df_v[COL_MES] == mes_sel_gest].copy()
        v_sel_gest = st.sidebar.selectbox("👤 Seleccionar Vista:", ["TODOS LOS ASESORES"] + sorted(df_mes[COL_ASESOR].unique().tolist()))
        st.markdown(f'<div class="main-header">📈 Reporte: {v_sel_gest} ({mes_sel_gest})</div>', unsafe_allow_html=True)
        
        if v_sel_gest == "TODOS LOS ASESORES":
            c1, c2, c3 = st.columns(3)
            with c1: st.markdown(f'<div class="card-style border-green"><div class="image-label">Total Ventas</div><div class="image-value-number">{len(df_mes)}</div></div>', unsafe_allow_html=True)
            with c2: 
                total_recaudacion = df_mes[COL_VALOR].sum()
                st.markdown(f'<div class="card-style border-blue"><div class="image-label">Recaudación (Sin IVA)</div><div class="image-value-number">${total_recaudacion:,.2f}</div></div>', unsafe_allow_html=True)
            with c3: 
                promedio = df_mes[COL_VALOR].mean()
                st.markdown(f'<div class="card-style" style="border-left: 8px solid #f1c40f;"><div class="image-label">Promedio</div><div class="image-value-number">${promedio:,.2f}</div></div>', unsafe_allow_html=True)
            
            resumen = df_mes.groupby(COL_ASESOR).agg({COL_ASESOR: 'count', COL_VALOR: 'sum'}).rename(columns={COL_ASESOR: 'CANT.', COL_VALOR: 'MONTO'}).reset_index().sort_values('MONTO', ascending=False)
            st.markdown('<div class="section-title">🏆 Ranking de Recaudación por Asesor</div>', unsafe_allow_html=True)
            fig_bar_global = px.bar(resumen, x='MONTO', y=COL_ASESOR, orientation='h', text='MONTO', color='MONTO', color_continuous_scale='Greens', labels={'MONTO': 'Monto Recaudado', COL_ASESOR: 'Asesor'}, hover_data={'MONTO': ':.2f'})
            fig_bar_global.update_layout(yaxis={'categoryorder':'total ascending'}, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white")
            fig_bar_global.update_traces(texttemplate='$%{text:,.2f}', textposition='outside')
            st.plotly_chart(fig_bar_global, use_container_width=True)
            col_t, col_p = st.columns([1.2, 1])
            with col_t: st.dataframe(resumen.style.format({'MONTO': '${:,.2f}'}), use_container_width=True, hide_index=True)
            with col_p:
                fig_p = px.pie(resumen, values='MONTO', names=COL_ASESOR, hole=0.4, title="Participación en el Mercado")
                fig_p.update_traces(textinfo='percent+label', hovertemplate='Asesor: %{label}<br>Monto: $%{value:.2f}')
                fig_p.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color="white")
                st.plotly_chart(fig_p, use_container_width=True)
        else:
            df_ind = df_mes[df_mes[COL_ASESOR] == v_sel_gest].copy()
            c1, c2 = st.columns(2)
            with c1: st.markdown(f'<div class="card-style border-green"><div class="image-label">Ventas</div><div class="image-value-number">{len(df_ind)}</div></div>', unsafe_allow_html=True)
            with c2: 
                monto_ind = df_ind[COL_VALOR].sum()
                st.markdown(f'<div class="card-style border-blue"><div class="image-label">Monto (Sin IVA)</div><div class="image-value-number">${monto_ind:,.2f}</div></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="section-title">🔍 Análisis Detallado: {v_sel_gest}</div>', unsafe_allow_html=True)
            g1, g2 = st.columns(2)
            with g1:
                df_prod = df_ind['PRODUCTO'].value_counts().reset_index()
                df_prod.columns = ['PRODUCTO_NAME', 'VENTAS']
                fig_prod = px.bar(df_prod, x='VENTAS', y='PRODUCTO_NAME', orientation='h', title="Mix de Productos Vendidos", text='VENTAS', color='VENTAS', color_continuous_scale='Blues')
                fig_prod.update_layout(yaxis={'categoryorder':'total ascending'}, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white")
                st.plotly_chart(fig_prod, use_container_width=True)
            with g2:
                df_sector = df_ind['SECTOR'].value_counts().reset_index()
                df_sector.columns = ['SECTOR_NAME', 'VENTAS']
                fig_sector = px.bar(df_sector, x='VENTAS', y='SECTOR_NAME', orientation='h', title="Distribución Geográfica de Ventas", text='VENTAS', color='VENTAS', color_continuous_scale='Oranges')
                fig_sector.update_layout(yaxis={'categoryorder':'total ascending'}, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white")
                st.plotly_chart(fig_sector, use_container_width=True)
            st.markdown('<div class="section-title">📋 Listado de Clientes</div>', unsafe_allow_html=True)
            columnas_finales = [c for c in ['CLIENTE ', 'PRODUCTO', 'PAQUETE', COL_VALOR, 'SECTOR', 'FECHA DE INSTALACION'] if c in df_ind.columns]
            st.dataframe(df_ind[columnas_finales].style.format({COL_VALOR: '${:,.2f}'}), use_container_width=True, hide_index=True)
# ==========================================
# MÓDULO 4: GESTIÓN DIARIA (VERSIÓN FINAL PULIDA)
# ==========================================
elif seccion == "📈 Gestión Diaria":
    df_gd = cargar_datos(st.session_state.url_gestion, "Ventas", limpiar_precios=True)
    
    if df_gd is not None:
        # 1. Preparación de Fechas
        df_gd['FECHA DE INSTALACION'] = pd.to_datetime(df_gd['FECHA DE INSTALACION'], errors='coerce')
        
        def extraer_fecha_inicio(rango):
            try:
                fecha_str = str(rango).split('/')[0].strip()
                return pd.to_datetime(fecha_str, errors='coerce')
            except: return pd.NaT

        df_gd['FECHA_INICIO_MES'] = df_gd['MES COMERCIAL'].apply(extraer_fecha_inicio)
        
        # --- FILTROS EN BARRA LATERAL ---
        meses_gd = sorted(df_gd['MES COMERCIAL'].dropna().unique().tolist(), reverse=True)
        mes_sel_gd = st.sidebar.selectbox("📅 Seleccionar Mes Comercial:", meses_gd, key="gd_mes")
        
        vendedores_gd = sorted(df_gd['ASESOR'].dropna().unique().tolist())
        vendedor_sel_gd = st.sidebar.selectbox("👤 Seleccionar Asesor:", ["TODOS"] + vendedores_gd, key="gd_vendedor")

        # Lógica de Cascada para Sectores
        mask_preliminar = (df_gd['MES COMERCIAL'] == mes_sel_gd)
        if vendedor_sel_gd != "TODOS":
            mask_preliminar &= (df_gd['ASESOR'] == vendedor_sel_gd)
        
        df_temp = df_gd[mask_preliminar]
        sectores_disponibles = sorted(df_temp['SECTOR'].dropna().astype(str).unique().tolist())
        
        sector_sel_gd = st.sidebar.selectbox("📍 Filtrar por Sector:", ["TODOS LOS SECTORES"] + sectores_disponibles, key="gd_sector_filter")

        # Filtro Final
        mask_final = mask_preliminar.copy()
        if sector_sel_gd != "TODOS LOS SECTORES":
            mask_final &= (df_gd['SECTOR'].astype(str) == sector_sel_gd)
        
        df_filtered = df_gd[mask_final].copy()

        st.markdown(f'<div class="main-header">📋 Gestión Diaria: {vendedor_sel_gd}</div>', unsafe_allow_html=True)

        if not df_filtered.empty:
            # --- TARJETAS (KPIs) ---
            c1, c2, c3 = st.columns(3)
            
            # 1. Sector Principal
            sector_top = df_filtered['SECTOR'].mode()[0] if not df_filtered['SECTOR'].empty else "N/A"
            with c1:
                st.markdown(f'<div class="card-style border-green"><div class="image-label">Sector Principal</div><div class="image-value-number" style="font-size:20px;">{str(sector_top).upper()}</div></div>', unsafe_allow_html=True)

            # 2. Canal de Venta Top (Reemplazo del Promedio)
            col_canal = 'DE DONDE PROVIENE LA VENTA'
            canal_top = df_filtered[col_canal].mode()[0] if not df_filtered[col_canal].empty else "N/A"
            with c2:
                st.markdown(f'<div class="card-style border-blue"><div class="image-label">Canal de Venta Top</div><div class="image-value-number" style="font-size:18px;">{str(canal_top).upper()}</div></div>', unsafe_allow_html=True)

            # 3. Paquete Estrella
            paquete_top = df_filtered['PAQUETE'].mode()[0] if not df_filtered['PAQUETE'].empty else "N/A"
            with c3:
                st.markdown(f'<div class="card-style" style="border-left: 8px solid #f1c40f;"><div class="image-label">Paquete Estrella</div><div class="image-value-number" style="font-size:18px;">{paquete_top}</div></div>', unsafe_allow_html=True)

            # --- SECCIÓN DE GRÁFICOS ---
            st.markdown('<div class="section-title">📊 Análisis Detallado</div>', unsafe_allow_html=True)
            col_izq, col_der = st.columns(2)
            
            with col_izq:
                if sector_sel_gd == "TODOS LOS SECTORES":
                    fig_sector = px.pie(df_filtered, names='SECTOR', title="Ventas por Sector (%)", hole=0.4)
                else:
                    # GRÁFICA DE BARRAS SOLICITADA
                    ventas_dia = df_filtered.groupby(df_filtered['FECHA DE INSTALACION'].dt.date).size().reset_index(name='VENTAS')
                    ventas_dia.columns = ['FECHA', 'VENTAS']
                    fig_sector = px.bar(ventas_dia, x='FECHA', y='VENTAS', title=f"Instalaciones en {sector_sel_gd}", text_auto=True, color_discrete_sequence=['#2ecc71'])
                    fig_sector.update_traces(textposition="outside", cliponaxis=False)
                
                fig_sector.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white")
                st.plotly_chart(fig_sector, use_container_width=True)
                
            with col_der:
                fig_pie_origen = px.pie(df_filtered, names='DE DONDE PROVIENE LA VENTA', title="Canales de Venta")
                fig_pie_origen.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color="white")
                st.plotly_chart(fig_pie_origen, use_container_width=True)

            # --- TABLA CON FORMATO DE FECHA NATURAL (Día de Mes, Año) ---
            st.markdown('<div class="section-title">📝 Listado de Instalaciones</div>', unsafe_allow_html=True)
            
            df_display = df_filtered[['FECHA DE INSTALACION', 'ASESOR', 'SECTOR', 'PAQUETE', 'DE DONDE PROVIENE LA VENTA']].copy()
            
            # Diccionario para meses en español
            meses_es = {
                1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
                7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
            }

            # Aplicamos formato "15 de Octubre, 2025" y quitamos los ceros
            df_display['FECHA DE INSTALACION'] = df_display['FECHA DE INSTALACION'].apply(
                lambda x: f"{x.day} de {meses_es[x.month]}, {x.year}" if pd.notnull(x) else "Sin Fecha"
            )
            
            st.dataframe(df_display, use_container_width=True, hide_index=True)

        else:
            st.warning("Sin datos para esta selección.")
