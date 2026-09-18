import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# Configuración de página de Streamlit
st.set_page_config(
    page_title="Dashboard de Control de Despliegue FTTH",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---- ESTILO: OSCURO ATENUADO CON SIDEBAR CLARO ----
st.markdown("""
    <style>
    /* Fondo principal: Gris pizarra oscuro atenuado */
    .stApp {
        background-color: #1E293B;
        color: #F8FAFC;
    }
    
    /* Títulos principales */
    .main-title {
        font-size:32px !important;
        font-weight: bold;
        color: #38BDF8; 
        margin-bottom: 5px;
    }
    .subtitle {
        font-size:16px !important;
        color: #94A3B8;
        margin-bottom: 25px;
    }
    
    /* Tarjetas de métricas */
    .metric-card {
        background-color: #334155;
        border: 1px solid #475569;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.2);
    }
    .section-title {
        font-size: 24px !important;
        font-weight: bold;
        color: #F1F5F9;
        margin-top: 10px;
        margin-bottom: 15px;
    }
    .section-title-center {
        font-size: 24px !important;
        font-weight: bold;
        color: #4ADE80;
        text-align: center;
        margin-top: 10px;
        margin-bottom: 15px;
    }
    .section-title-right {
        font-size: 24px !important;
        font-weight: bold;
        color: #38BDF8;
        text-align: right;
        margin-top: 10px;
        margin-bottom: 15px;
    }
    
    /* CONTRASTE: Barra de filtros (Sidebar) en fondo blanco */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0;
    }
    [data-testid="stSidebar"] .css-17eq0hr, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] h2 {
        color: #0F172A !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="main-title">📈 Dashboard Control de Despliegue (UIPs)</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Análisis interactivo de indicadores totales y por contratista para el control de infraestructura de Fibra Óptica.</p>', unsafe_allow_html=True)

# 1. Cargar datos desde archivo EXCEL (.xlsx)
@st.cache_data
def cargar_datos():
    archivo_excel = "CONSOLIDADO_DESPLIEGUE_CONTROL.xlsx"
    nombre_pestana = "BASE_CONSOLIDADA"
    
    if not os.path.exists(archivo_excel):
        raise FileNotFoundError(f"No se encontró el archivo Excel '{archivo_excel}' en el directorio actual.")

    try:
        df = pd.read_excel(archivo_excel, sheet_name=nombre_pestana)
    except ValueError:
        df = pd.read_excel(archivo_excel, sheet_name=0)
    
    # Limpiar espacios en blanco en los nombres de las columnas
    df.columns = df.columns.astype(str).str.strip()
    
    # Búsqueda flexible de la columna UIP FIN
    uip_col = None
    for col in df.columns:
        if col.upper().replace('_', ' ').strip() in ['UIP FIN', 'UIP_FIN', 'UIPFIN']:
            uip_col = col
            break
            
    if uip_col:
        df['UIP_Métrica'] = pd.to_numeric(df[uip_col], errors='coerce').fillna(0)
    else:
        df['UIP_Métrica'] = 0
    
    if 'CONTRATISTA' in df.columns:
        df['CONTRATISTA'] = df['CONTRATISTA'].fillna('SIN CONTRATISTA').astype(str).str.upper().str.strip()
    else:
        df['CONTRATISTA'] = 'SIN CONTRATISTA'
        
    if 'META' in df.columns:
        meta_datetime = pd.to_datetime(df['META'], errors='coerce')
        df['META'] = meta_datetime.dt.strftime('%Y-%m').fillna('SIN META')
    else:
        df['META'] = 'SIN META'
        
    if 'TIPO SUB PROYECTO' in df.columns:
        df['TIPO SUB PROYECTO'] = df['TIPO SUB PROYECTO'].fillna('SIN TIPO').astype(str).str.strip()
    else:
        df['TIPO SUB PROYECTO'] = 'SIN TIPO'
    
    columnas_fecha = [
        'FECHA DE NORMALIZACION', 'FECHA DE APROBACION DE DISEÑO', 'FECHA FIN DE DISEÑO',
        'FECHA FIN DE REGISTRO 3 GIS', 'FECHA DE INICIO REGISTRO 3 GIS', 'FECHA DE RESPUESTA ASOCIACION DIREC', 
        'FECHA FIN INSTALACION', 'FECHA DE INICIO INSTALACION', 'FECHA DE PERMISO INTERNO'
    ]
    for col in columnas_fecha:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            
    return df

# 1.1 Cargar archivo de METAS (META.xlsx)
@st.cache_data
def cargar_metas():
    archivo_meta = "META.xlsx"
    if not os.path.exists(archivo_meta):
        return pd.DataFrame(columns=['CONTRATISTA', 'META_VALOR', 'MES_COSECHA_STR'])

    try:
        df_m = pd.read_excel(archivo_meta)
        if 'COTRATISTA' not in df_m.columns and 'CONTRATISTA' not in df_m.columns:
            df_m = pd.read_excel(archivo_meta, header=1)
        
        df_m.columns = df_m.columns.astype(str).str.strip()
        
        if 'COTRATISTA' in df_m.columns:
            df_m.rename(columns={'COTRATISTA': 'CONTRATISTA'}, inplace=True)
            
        df_m = df_m.dropna(how='all')
        
        if 'CONTRATISTA' in df_m.columns:
            df_m['CONTRATISTA'] = df_m['CONTRATISTA'].astype(str).str.upper().str.strip()
            
        if 'META' in df_m.columns:
            df_m['META_VALOR'] = pd.to_numeric(df_m['META'], errors='coerce').fillna(0)
            
        if 'MES COSECHA' in df_m.columns:
            df_m['MES_COSECHA_STR'] = pd.to_datetime(df_m['MES COSECHA'], errors='coerce').dt.strftime('%Y-%m').fillna('SIN META')
        else:
            df_m['MES_COSECHA_STR'] = 'SIN META'
            
        return df_m
    except Exception as e:
        st.warning(f"⚠️ No se pudo procesar 'META.xlsx': {e}")
        return pd.DataFrame(columns=['CONTRATISTA', 'META_VALOR', 'MES_COSECHA_STR'])

df_raw = pd.DataFrame()
df_metas_raw = pd.DataFrame()

try:
    df_raw = cargar_datos()
    df_metas_raw = cargar_metas()
except Exception as e:
    st.error(f"❌ Error al cargar el archivo de datos: {e}")
    st.stop()

# 2. Sidebar para Filtros Interactivos
st.sidebar.header("🔍 Filtros de Búsqueda")

municipios = sorted(df_raw['MUNICIPIO'].dropna().unique().tolist()) if 'MUNICIPIO' in df_raw.columns else []
municipio_sel = st.sidebar.multiselect("Municipio", options=municipios, default=[])

contratistas = sorted(df_raw['CONTRATISTA'].unique().tolist()) if 'CONTRATISTA' in df_raw.columns else []
contratista_sel = st.sidebar.multiselect("Contratista", options=contratistas, default=[])

metas = sorted(df_raw['META'].unique().tolist()) if 'META' in df_raw.columns else []
if 'SIN META' in metas:
    metas.remove('SIN META')
    metas.append('SIN META')
meta_sel = st.sidebar.multiselect("Meta (Año-Mes)", options=metas, default=[])

tipos_sub = sorted(df_raw['TIPO SUB PROYECTO'].unique().tolist()) if 'TIPO SUB PROYECTO' in df_raw.columns else []
tipo_sub_sel = st.sidebar.multiselect("Tipo Sub Proyecto", options=tipos_sub, default=[])

# Aplicación de Filtros
df_filtrado = df_raw.copy()
if municipio_sel:
    df_filtrado = df_filtrado[df_filtrado['MUNICIPIO'].isin(municipio_sel)]
if contratista_sel:
    df_filtrado = df_filtrado[df_filtrado['CONTRATISTA'].isin(contratista_sel)]
if meta_sel:
    df_filtrado = df_filtrado[df_filtrado['META'].isin(meta_sel)]
if tipo_sub_sel:
    df_filtrado = df_filtrado[df_filtrado['TIPO SUB PROYECTO'].isin(tipo_sub_sel)]

# Cálculo del valor de META según los filtros
df_metas_filtrado = df_metas_raw.copy()
if not df_metas_filtrado.empty:
    if contratista_sel:
        df_metas_filtrado = df_metas_filtrado[df_metas_filtrado['CONTRATISTA'].isin(contratista_sel)]
    if meta_sel:
        df_metas_filtrado = df_metas_filtrado[df_metas_filtrado['MES_COSECHA_STR'].isin(meta_sel)]
    
    meta_calculada = df_metas_filtrado['META_VALOR'].sum()
else:
    meta_calculada = 0

# 3. Cálculo de Indicadores
def calcular_metricas(df):
    if df.empty:
        return {
            'Normalizados': 0, 'Diseñados': 0, 'Dibujados 3GIS': 0, 'Asociados': 0, 
            'Construidos': 0, 'Registrados': 0, 'En Proceso Instalación': 0, 
            'Permisos ZC': 0, 'Sin Inicio Instalación': 0, 'UIPs_Meta_Suma': 0
        }
    
    m = {}
    m['Normalizados'] = df[df['FECHA DE NORMALIZACION'].notna()]['UIP_Métrica'].sum() if 'FECHA DE NORMALIZACION' in df.columns else 0
    m['Diseñados'] = df[df['FECHA FIN DE DISEÑO'].notna()]['UIP_Métrica'].sum() if 'FECHA FIN DE DISEÑO' in df.columns else 0
    
    if 'FECHA DE INICIO REGISTRO 3 GIS' in df.columns:
        condicion_dibujados = (df['FECHA DE INICIO REGISTRO 3 GIS'].notna() & (df['FECHA DE INICIO REGISTRO 3 GIS'] != '') & (df['FECHA DE INICIO REGISTRO 3 GIS'] != 0) & (df['FECHA DE INICIO REGISTRO 3 GIS'].astype(str).str.strip() != '0'))
        m['Dibujados 3GIS'] = df[condicion_dibujados]['UIP_Métrica'].sum()
    else:
        m['Dibujados 3GIS'] = 0
        
    # ---- AJUSTE: CÁLCULO DE ASOCIADOS CON 'FECHA DE RESPUESTA ASOCIACION DIREC' Y SUMA DE UIP FIN ----
    if 'FECHA DE RESPUESTA ASOCIACION DIREC' in df.columns:
        serie_resp_str = df['FECHA DE RESPUESTA ASOCIACION DIREC'].astype(str).str.strip()
        condicion_asociados = (
            df['FECHA DE RESPUESTA ASOCIACION DIREC'].notna() & 
            (serie_resp_str != '') & 
            (serie_resp_str.str.lower() != 'nan') & 
            (serie_resp_str.str.lower() != 'nat') & 
            (serie_resp_str != '0') & 
            (df['FECHA DE RESPUESTA ASOCIACION DIREC'] != 0)
        )
        m['Asociados'] = df[condicion_asociados]['UIP_Métrica'].sum()
    else:
        m['Asociados'] = 0
        
    m['Construidos'] = df[df['FECHA FIN INSTALACION'].notna()]['UIP_Métrica'].sum() if 'FECHA FIN INSTALACION' in df.columns else 0
    m['Registrados'] = df[df['FECHA FIN DE REGISTRO 3 GIS'].notna()]['UIP_Métrica'].sum() if 'FECHA FIN DE REGISTRO 3 GIS' in df.columns else 0
    m['En Proceso Instalación'] = df[df['FECHA DE INICIO INSTALACION'].notna() & df['FECHA FIN INSTALACION'].isna()]['UIP_Métrica'].sum() if ('FECHA DE INICIO INSTALACION' in df.columns and 'FECHA FIN INSTALACION' in df.columns) else 0
    
    if 'ESTADO GENERAL COMERCIAL' in df.columns:
        estado_comercial_clean = df['ESTADO GENERAL COMERCIAL'].fillna('').astype(str).str.strip().str.upper().str.replace('Ó', 'O', regex=False)
        condicion_permisos = (estado_comercial_clean == 'APROBADO CONSTRUCCION')
        m['Permisos ZC'] = df[condicion_permisos]['UIP_Métrica'].sum()
    else:
        m['Permisos ZC'] = 0
        
    m['Sin Inicio Instalación'] = df[df['FECHA DE INICIO INSTALACION'].isna()]['UIP_Métrica'].sum() if 'FECHA DE INICIO INSTALACION' in df.columns else 0
    m['UIPs_Meta_Suma'] = df[df['META'] != 'SIN META']['UIP_Métrica'].sum() if 'META' in df.columns else 0
    
    return m

totales = calcular_metricas(df_filtrado)

# 4. Encabezado e Indicadores
col_titulo, col_meta_empresa, col_bolsa = st.columns([1, 1, 1])

with col_titulo:
    st.markdown('<p class="section-title">📊 Resumen de UIPs Totales</p>', unsafe_allow_html=True)

with col_meta_empresa:
    st.markdown(f'<p class="section-title-center">🎯 META COSECHA: {meta_calculada:,.0f}</p>', unsafe_allow_html=True)

with col_bolsa:
    st.markdown(f'<p class="section-title-right">🎯 UIPs BOLSA: {totales["UIPs_Meta_Suma"]:,.0f}</p>', unsafe_allow_html=True)

# Grid de Tarjetas / KPIs
cols = st.columns(3)
kpis = [
    ("UIPs Normalizados", totales['Normalizados'], "🔵"),
    ("UIPs Diseñados", totales['Diseñados'], "🟢"),
    ("UIPs Dibujados 3GIS", totales['Dibujados 3GIS'], "🎨"),
    ("UIPs Asociados", totales['Asociados'], "🔗"),
    ("UIPs Construidos", totales['Construidos'], "🏗️"),
    ("UIPs Registrados", totales['Registrados'], "📝"),
    ("UIPs en Proceso de Instalación", totales['En Proceso Instalación'], "⏳"),
    ("UIPs con Permisos ZC", totales['Permisos ZC'], "🏢"),
    ("UIPs Sin Inicio Instalación", totales['Sin Inicio Instalación'], "❌")
]

for idx, (label, value, icon) in enumerate(kpis):
    col = cols[idx % 3]
    with col:
        st.markdown(f"""
        <div class="metric-card">
            <span style="font-size: 24px;">{icon}</span>
            <p style="margin: 0; color: #94A3B8; font-size: 14px; font-weight: 500;">{label}</p>
            <h3 style="margin: 5px 0 0 0; color: #FFFFFF; font-size: 28px;">{value:,.0f}</h3>
        </div>
        """, unsafe_allow_html=True)
        st.write("")

# ==============================================================================
# 5. TABLA DETALLADA DE PROYECTOS
# ==============================================================================
st.markdown("---")
st.subheader("📋 Detalle de Proyectos")

campos_solicitados = [
    'ESTADO',
    'N° DE PROYECTO',
    'NOMBRE DEL PROYECTO',
    'UIP FIN',
    'MUNICIPIO',
    'CENTRAL',
    'CABLE',
    'TIPO SUB PROYECTO',
    'CONTRATISTA',
    'ESTADO GENERAL COMERCIAL'
]

def normalizar_texto(txt):
    return txt.upper().replace('°', '').replace('N°', 'N').replace('_', ' ').replace(' ', '').replace('Ó', 'O')

mapa_columnas_existentes = {normalizar_texto(c): c for c in df_filtrado.columns}

cols_encontradas = []
renombres = {}

for campo_req in campos_solicitados:
    clave_norm = normalizar_texto(campo_req)
    if clave_norm in mapa_columnas_existentes:
        col_real = mapa_columnas_existentes[clave_norm]
        cols_encontradas.append(col_real)
        renombres[col_real] = campo_req

if not df_filtrado.empty and cols_encontradas:
    df_tabla_detalle = df_filtrado[cols_encontradas].copy()
    df_tabla_detalle.rename(columns=renombres, inplace=True)
    
    if 'UIP FIN' in df_tabla_detalle.columns:
        df_tabla_detalle['UIP FIN'] = pd.to_numeric(df_tabla_detalle['UIP FIN'], errors='coerce').fillna(0)
        total_uip_fin = df_tabla_detalle['UIP FIN'].sum()
    else:
        total_uip_fin = 0

    st.dataframe(
        df_tabla_detalle.style.format({'UIP FIN': "{:,.0f}"}, na_rep=""),
        use_container_width=True,
        hide_index=True
    )
    st.caption(f"📌 Total registros: **{len(df_tabla_detalle):,}** | Suma Total UIP FIN: **{total_uip_fin:,.0f}**")
else:
    st.info("ℹ️ No hay proyectos disponibles para los filtros seleccionados o no se encontraron las columnas en la base de datos.")

# ==============================================================================
# 6. DESGLOSE POR CONTRATISTA
# ==============================================================================
st.markdown("---")
st.subheader("👷 Comparativa Detallada por Contratista")

if not df_filtrado.empty and 'CONTRATISTA' in df_filtrado.columns:
    contratistas_list = df_filtrado['CONTRATISTA'].unique()
    data_contratistas = []

    for c in contratistas_list:
        df_c = df_filtrado[df_filtrado['CONTRATISTA'] == c]
        m_c = calcular_metricas(df_c)
        m_c['Contratista'] = c
        data_contratistas.append(m_c)

    df_contratistas = pd.DataFrame(data_contratistas)
else:
    df_contratistas = pd.DataFrame()

if not df_contratistas.empty:
    cols_orden = ['Contratista', 'Normalizados', 'Diseñados', 'Dibujados 3GIS', 'Asociados', 'Construidos', 'Registrados', 'En Proceso Instalación', 'Permisos ZC', 'Sin Inicio Instalación']
    df_contratistas = df_contratistas[cols_orden]
    
    st.markdown("#### Distribución de Avances Totales Simultáneos")
    df_grafico_largo = df_contratistas.melt(id_vars=['Contratista'], value_vars=cols_orden[1:], var_name='Indicador', value_name='UIPs')
    
    fig = px.bar(
        df_grafico_largo,
        x='Contratista',
        y='UIPs',
        color='Indicador',
        barmode='group',
        text_auto=',.0f',
        title="Comparativo Global de todos los Indicadores por Empresa Contratista",
        template="plotly_dark"
    )
    
    fig.update_layout(
        title=dict(
            font=dict(color='#F8FAFC', size=16)
        ),
        xaxis_tickangle=-45, 
        height=500,
        plot_bgcolor='#1E293B',
        paper_bgcolor='#1E293B',
        font=dict(color='#F1F5F9'),
        legend_title_text='Indicadores UIP',
        legend=dict(
            font=dict(color='#FFFFFF', size=12),
            title=dict(font=dict(color='#FFFFFF', size=13))
        ),
        xaxis=dict(showgrid=True, gridcolor='#334155'),
        yaxis=dict(showgrid=True, gridcolor='#334155')
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # Tabla Transpuesta por Contratista
    st.markdown("#### Tabla de Datos por Contratista")
    df_transpuesto = df_contratistas.set_index('Contratista').T
    
    fondo_oscuro = "#222E3F"
    linea_grilla = "1px solid #334155"
    
    styled_df = (df_transpuesto.style
                 .format("{:,.0f}")
                 .set_properties(**{
                     'background-color': fondo_oscuro, 
                     'color': '#FFFFFF', 
                     'text-align': 'center !important',
                     'font-size': '15px !important',
                     'font-weight': 'bold !important',
                     'border': f'{linea_grilla} !important'
                 }) 
                 .set_table_styles([
                     {'selector': 'th.col_heading', 'props': [
                         ('background-color', f'{fondo_oscuro} !important'),
                         ('color', '#38BDF8 !important'),
                         ('text-align', 'center !important'),
                         ('font-size', '15px !important'),
                         ('font-weight', 'bold !important'),
                         ('border', f'{linea_grilla} !important')
                     ]},
                     {'selector': 'th.row_heading', 'props': [
                         ('background-color', f'{fondo_oscuro} !important'),
                         ('color', '#E2E8F0 !important'),
                         ('text-align', 'left !important'),
                         ('font-size', '14px !important'),
                         ('font-weight', 'bold !important'),
                         ('border', f'{linea_grilla} !important')
                     ]},
                     {'selector': 'th.index_name', 'props': [
                         ('background-color', f'{fondo_oscuro} !important'),
                         ('border', f'{linea_grilla} !important')
                     ]}
                 ]))
    
    st.dataframe(
        styled_df,
        use_container_width=True
    )
else:
    st.info("No hay datos disponibles para los filtros aplicados.")