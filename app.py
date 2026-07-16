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

# Título y estilos de la cabecera
st.markdown("""
    <style>
    .main-title {
        font-size:32px !important;
        font-weight: bold;
        color: #1E3A8A;
        margin-bottom: 5px;
    }
    .subtitle {
        font-size:16px !important;
        color: #555555;
        margin-bottom: 25px;
    }
    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    /* Estilo para los títulos de sección alineados */
    .section-title {
        font-size: 24px !important;
        font-weight: bold;
        color: #1E293B;
        margin-top: 10px;
        margin-bottom: 15px;
    }
    .section-title-right {
        font-size: 24px !important;
        font-weight: bold;
        color: #0F766E; /* Color turquesa oscuro para diferenciar UIPs BOLSA */
        text-align: right;
        margin-top: 10px;
        margin-bottom: 15px;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="main-title">📈 Dashboard Control de Despliegue (UIPs)</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Análisis interactivo de indicadores totales y por contratista para el control de infraestructura de Fibra Óptica.</p>', unsafe_allow_html=True)

# 1. Cargar datos desde archivo EXCEL (.xlsx)
@st.cache_data
def cargar_datos():
    archivo_excel = "CONSOLIDADO_DESPLIEGUE_CONTROL.xlsx"
    nombre_pestana = "BASE_CONSOLIDADA"  # Cambiar aquí si la pestaña tiene otro nombre
    
    if not os.path.exists(archivo_excel):
        raise FileNotFoundError(f"No se encontró el archivo Excel '{archivo_excel}' en el directorio actual.")

    # Cargamos el archivo de Excel usando openpyxl como motor
    try:
        df = pd.read_excel(archivo_excel, sheet_name=nombre_pestana)
    except ValueError:
        # En caso de que no encuentre la pestaña especificada, carga la primera disponible
        df = pd.read_excel(archivo_excel, sheet_name=0)
    
    # Limpieza de nombres de columnas (eliminar espacios en blanco alrededor de los nombres)
    df.columns = df.columns.str.strip()
    
    # AJUSTE: El cálculo de UIPs se realiza ESTRICTAMENTE con el campo UIP FIN
    uip_fin_col = 'UIP FIN' if 'UIP FIN' in df.columns else 'UIP_FIN'
    df['UIP_Métrica'] = pd.to_numeric(df[uip_fin_col], errors='coerce').fillna(0)
    
    if 'CONTRATISTA' in df.columns:
        df['CONTRATISTA'] = df['CONTRATISTA'].fillna('SIN CONTRATISTA').astype(str).str.upper().str.strip()
    else:
        df['CONTRATISTA'] = 'SIN CONTRATISTA'
        
    # Ajuste y formateo de la columna META
    if 'META' in df.columns:
        meta_datetime = pd.to_datetime(df['META'], errors='coerce')
        df['META'] = meta_datetime.dt.strftime('%Y-%m').fillna('SIN META')
    else:
        df['META'] = 'SIN META'
        
    # Limpieza de valores en TIPO SUB PROYECTO
    if 'TIPO SUB PROYECTO' in df.columns:
        df['TIPO SUB PROYECTO'] = df['TIPO SUB PROYECTO'].fillna('SIN TIPO').astype(str).str.strip()
    else:
        df['TIPO SUB PROYECTO'] = 'SIN TIPO'
    
    # Conversión de fechas críticas para asegurar compatibilidad
    columnas_fecha = [
        'FECHA DE NORMALIZACION', 'FECHA DE APROBACION DE DISEÑO', 'FECHA FIN DE DISEÑO',
        'FECHA FIN DE REGISTRO 3 GIS', 'FECHA DE INICIO REGISTRO 3 GIS', 'FECHA DE RESPUESTA ASOCIACION DIREC', 
        'FECHA FIN INSTALACION', 'FECHA DE INICIO INSTALACION', 
        'FECHA DE PERMISO INTERNO'
    ]
    for col in columnas_fecha:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            
    return df

# Inicializamos df_raw fuera del bloque try para asegurar su existencia global
df_raw = pd.DataFrame()

try:
    df_raw = cargar_datos()
except Exception as e:
    st.error(f"❌ Error al cargar el archivo de datos: {e}")
    st.info("Por favor, asegúrate de que el archivo Excel **'CONSOLIDADO_DESPLIEGUE_CONTROL.xlsx'** esté en la misma carpeta que tu script **'app.py'**.")
    st.stop()

# 2. Sidebar para Filtros Interactivos
st.sidebar.header("🔍 Filtros de Búsqueda")

# Filtro de Municipio
municipios = sorted(df_raw['MUNICIPIO'].dropna().unique().tolist()) if 'MUNICIPIO' in df_raw.columns else []
municipio_sel = st.sidebar.multiselect("Municipio", options=municipios, default=[])

# Filtro de Contratista
contratistas = sorted(df_raw['CONTRATISTA'].unique().tolist()) if 'CONTRATISTA' in df_raw.columns else []
contratista_sel = st.sidebar.multiselect("Contratista", options=contratistas, default=[])

# Filtro por el campo META (Año-Mes)
metas = sorted(df_raw['META'].unique().tolist()) if 'META' in df_raw.columns else []
if 'SIN META' in metas:
    metas.remove('SIN META')
    metas.append('SIN META')
meta_sel = st.sidebar.multiselect("Meta (Año-Mes)", options=metas, default=[])

# ---- FILTRO POR TIPO SUB PROYECTO ----
tipos_sub = sorted(df_raw['TIPO SUB PROYECTO'].unique().tolist()) if 'TIPO SUB PROYECTO' in df_raw.columns else []
tipo_sub_sel = st.sidebar.multiselect("Tipo Sub Proyecto", options=tipos_sub, default=[])

# Aplicar filtros dinámicamente
df_filtrado = df_raw.copy()
if municipio_sel:
    df_filtrado = df_filtrado[df_filtrado['MUNICIPIO'].isin(municipio_sel)]
if contratista_sel:
    df_filtrado = df_filtrado[df_filtrado['CONTRATISTA'].isin(contratista_sel)]
if meta_sel:
    df_filtrado = df_filtrado[df_filtrado['META'].isin(meta_sel)]
if tipo_sub_sel:
    df_filtrado = df_filtrado[df_filtrado['TIPO SUB PROYECTO'].isin(tipo_sub_sel)]

# 3. Cálculo de Indicadores Solicitados
def calcular_metricas(df):
    m = {}
    if df.empty:
        return {
            'Normalizados': 0, 'Diseñados': 0, 'Dibujados 3GIS': 0, 'Asociados': 0, 
            'Construidos': 0, 'Registrados': 0, 'En Proceso Instalación': 0, 
            'Permisos ZC': 0, 'Sin Inicio Instalación': 0, 'UIPs_Meta_Suma': 0
        }
    
    # UIPs Normalizados
    m['Normalizados'] = df[df['FECHA DE NORMALIZACION'].notna()]['UIP_Métrica'].sum() if 'FECHA DE NORMALIZACION' in df.columns else 0
    
    # UIPs Diseñados calculado a partir de [FECHA FIN DE DISEÑO] no vacío
    m['Diseñados'] = df[df['FECHA FIN DE DISEÑO'].notna()]['UIP_Métrica'].sum() if 'FECHA FIN DE DISEÑO' in df.columns else 0
    
    # UIPs Dibujados 3GIS basado en FECHA DE INICIO REGISTRO 3 GIS
    if 'FECHA DE INICIO REGISTRO 3 GIS' in df.columns:
        condicion_dibujados = (
            df['FECHA DE INICIO REGISTRO 3 GIS'].notna() & 
            (df['FECHA DE INICIO REGISTRO 3 GIS'] != '') & 
            (df['FECHA DE INICIO REGISTRO 3 GIS'] != 0) & 
            (df['FECHA DE INICIO REGISTRO 3 GIS'].astype(str).str.strip() != '0')
        )
        m['Dibujados 3GIS'] = df[condicion_dibujados]['UIP_Métrica'].sum()
    else:
        m['Dibujados 3GIS'] = 0
        
    # UIPs Asociados basado en FECHA DE ENVIO ASOCIACION DIREC
    if 'FECHA DE ENVIO ASOCIACION DIREC' in df.columns:
        serie_envio_str = df['FECHA DE ENVIO ASOCIACION DIREC'].astype(str).str.strip()
        
        condicion_asociados = (
            df['FECHA DE ENVIO ASOCIACION DIREC'].notna() & 
            (serie_envio_str != '') & 
            (serie_envio_str.str.lower() != 'nan') & 
            (serie_envio_str.str.lower() != 'nat') & 
            (serie_envio_str != '0') &
            (df['FECHA DE ENVIO ASOCIACION DIREC'] != 0)
        )
        m['Asociados'] = df[condicion_asociados]['UIP_Métrica'].sum()
    else:
        m['Asociados'] = 0
        
    # UIPs Construidos
    m['Construidos'] = df[df['FECHA FIN INSTALACION'].notna()]['UIP_Métrica'].sum() if 'FECHA FIN INSTALACION' in df.columns else 0
    # UIPs Registrados
    m['Registrados'] = df[df['FECHA FIN DE REGISTRO 3 GIS'].notna()]['UIP_Métrica'].sum() if 'FECHA FIN DE REGISTRO 3 GIS' in df.columns else 0
    # UIPs en Proceso de Instalación
    m['En Proceso Instalación'] = df[df['FECHA DE INICIO INSTALACION'].notna() & df['FECHA FIN INSTALACION'].isna()]['UIP_Métrica'].sum() if ('FECHA DE INICIO INSTALACION' in df.columns and 'FECHA FIN INSTALACION' in df.columns) else 0
    
    # --------------------------------------------------------------------------------------------------------------------------------------
    # Criterio Corregido: Filtramos 'APROBADO CONSTRUCCION' o 'APROBADO CONSTRUCCIÓN' removiendo tildes y espacios extra
    # --------------------------------------------------------------------------------------------------------------------------------------
    if 'ESTADO GENERAL COMERCIAL' in df.columns:
        # Reemplaza tildes en Ó, remueve espacios y convierte a mayúsculas
        estado_comercial_clean = (
            df['ESTADO GENERAL COMERCIAL']
            .fillna('')
            .astype(str)
            .str.strip()
            .str.upper()
            .str.replace('Ó', 'O', regex=False)
        )
        
        # Filtramos por las dos variantes lógicas posibles en español para prevenir problemas del origen Excel
        condicion_permisos = (estado_comercial_clean == 'APROBADO CONSTRUCCION')
        m['Permisos ZC'] = df[condicion_permisos]['UIP_Métrica'].sum()
    else:
        m['Permisos ZC'] = 0
        
    # UIPs Sin Inicio Instalación
    m['Sin Inicio Instalación'] = df[df['FECHA DE INICIO INSTALACION'].isna()]['UIP_Métrica'].sum() if 'FECHA DE INICIO INSTALACION' in df.columns else 0
    
    # UIPs META
    m['UIPs_Meta_Suma'] = df[df['META'] != 'SIN META']['UIP_Métrica'].sum() if 'META' in df.columns else 0
    
    return m

totales = calcular_metricas(df_filtrado)

# 4. Mostrar Encabezado con Título a la Izquierda y UIPs BOLSA a la Derecha
col_titulo, col_meta = st.columns([1, 1])

with col_titulo:
    st.markdown('<p class="section-title">📊 Resumen de UIPs Totales</p>', unsafe_allow_html=True)

with col_meta:
    # REEMPLAZADO: UIPs META cambiado a UIPs BOLSA
    st.markdown(f'<p class="section-title-right">🎯 UIPs BOLSA: {totales["UIPs_Meta_Suma"]:,.0f}</p>', unsafe_allow_html=True)

# Grid de KPIs
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
            <p style="margin: 0; color: #64748B; font-size: 14px; font-weight: 500;">{label}</p>
            <h3 style="margin: 5px 0 0 0; color: #1E293B; font-size: 28px;">{value:,.0f}</h3>
        </div>
        """, unsafe_allow_html=True)
        st.write("")

# 5. Desglose por Contratista
st.markdown("---")
st.subheader("👷 Comparativa Detallada por Contratista")

# Agrupar cálculos por contratista
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
    # Asegurar orden de columnas
    cols_orden = [
        'Contratista', 'Normalizados', 'Diseñados', 'Dibujados 3GIS', 
        'Asociados', 'Construidos', 'Registrados', 'En Proceso Instalación', 
        'Permisos ZC', 'Sin Inicio Instalación'
    ]
    df_contratistas = df_contratistas[cols_orden]
    
    # Gráfico de barras interactivo
    st.markdown("#### Distribución de Avances")
    indicador_grafico = st.selectbox("Selecciona un indicador para visualizar:", cols_orden[1:])
    
    fig = px.bar(
        df_contratistas.sort_values(by=indicador_grafico, ascending=False),
        x='Contratista',
        y=indicador_grafico,
        text_auto=',.0f',
        title=f"{indicador_grafico} por Contratista",
        color=indicador_grafico,
        color_continuous_scale=px.colors.sequential.Viridis
    )
    fig.update_layout(xaxis_tickangle=-45, height=500)
    st.plotly_chart(fig, use_container_width=True)
    
    # Tabla Interactiva de Datos
    st.markdown("#### Tabla de Datos por Contratista")
    st.dataframe(
        df_contratistas.style.format({
            c: "{:,.0f}" for c in cols_orden[1:] if c != 'Contratista'
        }),
        use_container_width=True
    )
else:
    st.info("No hay datos disponibles para los filtros aplicados.")
    