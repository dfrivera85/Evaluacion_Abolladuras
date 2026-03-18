"""
dstrain_app.py
==============
Interfaz Streamlit para evaluación de strain (deformación) en abolladuras
de tuberías de hidrocarburos.

Uso:
    streamlit run dstrain_app.py
"""

import csv
import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import io
import traceback
from dstrain_module import process_dataframe, COL

# ─── Configuración de página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Evaluación de Strain – Abolladuras",
    page_icon="🔩",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS personalizado ────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Fondo claro profesional */
    .stApp {
        background-color: #f8fafc;
        min-height: 100vh;
    }

    /* Header principal */
    .main-header {
        background: linear-gradient(90deg, #1e40af 0%, #3b82f6 100%);
        border-radius: 16px;
        padding: 2rem 2.5rem;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px -5px rgba(59, 130, 246, 0.2);
        border: 1px solid rgba(255,255,255,0.1);
    }
    .main-header h1 {
        color: #ffffff;
        font-size: 2rem;
        font-weight: 700;
        margin: 0 0 0.25rem 0;
        letter-spacing: -0.5px;
    }
    .main-header p {
        color: rgba(255,255,255,0.9);
        font-size: 0.95rem;
        margin: 0;
    }

    /* Tarjetas de métricas */
    .metric-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.25rem 1.5rem;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .metric-card .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #1e293b;
    }
    .metric-card .metric-label {
        font-size: 0.8rem;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-top: 0.25rem;
    }
    .metric-red   { color: #ef4444 !important; }
    .metric-green { color: #10b981 !important; }
    .metric-gray  { color: #64748b !important; }
    .metric-yellow{ color: #f59e0b !important; }

    /* Sección */
    .section-title {
        color: #1e40af;
        font-size: 0.85rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #e2e8f0;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #e2e8f0;
    }

    /* Botón principal */
    .stButton > button {
        background: #2563eb;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        transition: all 0.2s;
    }
    .stButton > button:hover {
        background: #1d4ed8;
        border-color: #1d4ed8;
    }

    /* Info box */
    .info-box {
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        border-radius: 10px;
        padding: 1rem 1.25rem;
        color: #1e40af;
        font-size: 0.88rem;
        margin-bottom: 1rem;
    }

    /* Ajustes para textos generales */
    p, span, label {
        color: #334155;
    }

    div[data-testid="stMetric"] {
        background: #ffffff;
        border-radius: 10px;
        padding: 0.75rem 1rem;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    div[data-testid="stMetric"] label {
        color: #64748b !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #0f172a !important;
    }
</style>
""", unsafe_allow_html=True)


# ─── Nombres de columnas esperadas en EntradaDatos ────────────────────────────
COLUMN_NAMES = [
    "Dist. Registro (km)", "Latitud", "Longitud", "Altura (m)",
    "Espesor (mm)", "SMYS (psi)", "SMTS (psi)", "Diám. Externo (mm)",
    "Tipo Anomalía", "Comentario", "Pos. Pared", "Pos. Horaria",
    "Profundidad (%)", "Profundidad (mm)", "Longitud (mm)", "Ancho (mm)",
    "N° Soldadura", "D.Sol.Inf (km)", "D.Sol.Sup (km)", "Long. Sold.",
    "En Soldadura", "Presión Diseño (psi)", "Factor Diseño",
    "Descripción", "Norma", "Dictamen", "Intervención", "Recomendación",
    "Strain Original", "Col30", "Fecha Inicio",
    "Hr mm P95", "Hr % P95", "Dict P95",
    "Hr mm P50", "Hr % P50", "Dict P50", "Reparada",
    "MOP (psi)", "Alt. Arruga",
]


# ─── Función para leer Excel ──────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes, sheet_name: str = "EntradaDatos") -> pd.DataFrame:
    """
    Lee la hoja EntradaDatos del xlsm.
    Datos comienzan en fila 8 (índice 7 en openpyxl, base 1).
    El header está en la fila 7 del Excel.
    """
    wb = openpyxl.load_workbook(
        filename=io.BytesIO(file_bytes),
        read_only=True,
        data_only=True,
        keep_vba=False,
    )
    ws = wb[sheet_name]

    all_rows = list(ws.iter_rows(values_only=True))
    # Fila 7 = índice 6 (header), fila 8 en adelante = datos
    header_row = all_rows[6] if len(all_rows) > 6 else []
    data_rows  = all_rows[7:] if len(all_rows) > 7 else []

    # Filtrar filas completamente vacías
    data_rows = [r for r in data_rows if any(c is not None for c in r)]

    if not data_rows:
        return pd.DataFrame()

    df = pd.DataFrame(data_rows)

    # Asignar nombres de columna
    # Usa los nombres reales del Excel si están disponibles, o nombres predefinidos
    n_cols = df.shape[1]
    col_labels = []
    for idx in range(n_cols):
        if idx < len(header_row) and header_row[idx] is not None:
            col_labels.append(str(header_row[idx]).strip())
        elif idx < len(COLUMN_NAMES):
            col_labels.append(COLUMN_NAMES[idx])
        else:
            col_labels.append(f"Col_{idx + 1}")

    # Ensure unique column names by appending suffixes to duplicates
    final_cols = []
    counts = {}
    for col in col_labels:
        if col in counts:
            counts[col] += 1
            final_cols.append(f"{col}.{counts[col]}")
        else:
            counts[col] = 0
            final_cols.append(col)

    df.columns = final_cols
    wb.close()
    return df


def safe_read_csv(file_obj):
    """
    Lee un archivo CSV intentando UTF-8 y cayendo a Latin-1 si falla.
    """
    try:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
        return pd.read_csv(file_obj)
    except UnicodeDecodeError:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
        return pd.read_csv(file_obj, encoding="latin1")


# ─── Función para colorear dictamen ──────────────────────────────────────────
def color_dictamen(val):
    if isinstance(val, str):
        if "No cumple" in val:
            return "background-color: rgba(248,113,113,0.25); color: #f87171; font-weight:600;"
        elif "Cumple criterio" in val:
            return "background-color: rgba(74,222,128,0.15); color: #4ade80; font-weight:600;"
        elif "No evaluada" in val:
            return "background-color: rgba(148,163,184,0.1); color: #94a3b8;"
        elif "faltante" in val or "Error" in val:
            return "background-color: rgba(251,191,36,0.15); color: #fbbf24;"
    return ""


def color_strain(val):
    if pd.isna(val) or val is None or val == "":
        return "color: #64748b;"
    try:
        f = float(val)
        pct = abs(f) * 100
        if pct >= 6:
            return "color: #f87171; font-weight: 600;"
        elif pct >= 3:
            return "color: #fbbf24;"
        else:
            return "color: #4ade80;"
    except Exception:
        return ""


# ─── Layout ───────────────────────────────────────────────────────────────────

# Header
st.markdown("""
<div class="main-header">
    <h1>🔩 Evaluación de Strain – Abolladuras</h1>
    <p>Módulo de integridad para tuberías de hidrocarburos · API-1160 / ASME B31.4</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    app_mode = st.radio("Modo de Operación", ["Evaluación de Strain", "Análisis Rainflow"])
    st.markdown("---")
    
    st.markdown("### 📂 Archivo de datos")
    uploaded = st.file_uploader(
        "Seleccione el archivo Excel (.xlsm / .xlsx)",
        type=["xlsm", "xlsx"],
        help="Hoja requerida: EntradaDatos · Datos desde fila 8",
    )
    st.markdown("---")
    st.markdown("### ⚙️ Configuración General")
    sheet_name = st.text_input("Nombre de hoja", value="EntradaDatos")
    
    if app_mode == "Evaluación de Strain":
        show_all_cols = st.checkbox("Mostrar todas las columnas de entrada", value=False)
    else:
        st.markdown("### 🌊 Datos SCADA y Topología")
        juntas_file = st.file_uploader("Seleccione juntas_soldadura.csv", type=["csv"], key="juntas")
        scada_succion = st.file_uploader("SCADA Succión (CSV)", type=["csv"], key="succion")
        scada_descarga = st.file_uploader("SCADA Descarga (CSV)", type=["csv"], key="descarga")
        
        st.markdown("#### Propiedades del Fluido (Por defecto Crudo)")
        specific_gravity = st.number_input("Gravedad Específica", value=0.85, step=0.01)
        viscosity = st.number_input("Viscosidad (cSt)", value=15.0, step=0.1)
        show_all_cols = False
        
    st.markdown("---")
    st.markdown("""
    <div style='color:#64748b; font-size:0.78rem; line-height:1.6;'>
    <b style='color:#94a3b8;'>Metodología</b><br>
    Algoritmo de strain según perfil de abolladura (polinomio de 6° grado).
    Criterio de falla: |ε| ≥ 6%.<br><br>
    <b style='color:#94a3b8;'>Normas aplicadas</b><br>
    API-1160 · ASME B31.4 · API-1183
    </div>
    """, unsafe_allow_html=True)


if uploaded is None:
    st.markdown("""
    <div class="info-box">
    📌 Cargue el archivo <b>Modulo 2. Distorsiones.xlsm</b> en el panel lateral para comenzar.
    Los datos de entrada se leerán automáticamente desde la hoja <code>EntradaDatos</code> (filas desde la 8).
    </div>
    """, unsafe_allow_html=True)

    # Instrucciones de uso
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-value metric-gray">1</div>
            <div class="metric-label">Cargar Excel</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-value metric-gray">2</div>
            <div class="metric-label">Verificar datos y SCADA</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-value metric-gray">3</div>
            <div class="metric-label">Calcular &amp; Exportar</div>
        </div>
        """, unsafe_allow_html=True)
    st.stop()

# ─── Cargar datos ─────────────────────────────────────────────────────────────
with st.spinner("Leyendo hoja EntradaDatos…"):
    try:
        file_bytes = uploaded.read()
        df_input = load_excel(file_bytes, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"❌ Error al leer el archivo: {e}")
        st.stop()

if df_input.empty:
    st.warning("⚠️ No se encontraron datos en la hoja especificada.")
    st.stop()

n_rows = len(df_input)
n_cols_in = len(df_input.columns)

# ─── Contenido principal ──────────────────────────────────────────────────────
if app_mode == "Análisis Rainflow":
    import rainflow
    import tempfile
    import os
    st.markdown('<p class="section-title">Análisis Rainflow (Datos SCADA)</p>', unsafe_allow_html=True)
    st.markdown("Seleccione la abolladura de interés para estimar su espectro de presión y ejecutar el recuento.")
    
    if not (juntas_file and scada_succion and scada_descarga):
        st.warning("⚠️ Cargue los 3 archivos CSV laterales (juntas, succión y descarga) para continuar.")
        st.stop()
        
    try:
        df_succion = safe_read_csv(scada_succion)
        if 'timestamp' not in df_succion.columns and 'Fecha' in df_succion.columns:
             df_succion.rename(columns={'Fecha': 'timestamp'}, inplace=True)
        if 'presion_psi' not in df_succion.columns and 'Presion' in df_succion.columns:
             df_succion.rename(columns={'Presion': 'presion_psi'}, inplace=True)
             
        df_descarga = safe_read_csv(scada_descarga)
        if 'timestamp' not in df_descarga.columns and 'Fecha' in df_descarga.columns:
             df_descarga.rename(columns={'Fecha': 'timestamp'}, inplace=True)
        if 'presion_psi' not in df_descarga.columns and 'Presion' in df_descarga.columns:
             df_descarga.rename(columns={'Presion': 'presion_psi'}, inplace=True)
             
        # Guardar juntas.csv temporalmente
        with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp:
            tmp.write(juntas_file.getvalue())
            juntas_path = tmp.name
            
    except Exception as e:
        st.error(f"❌ Error al leer archivos CSV: {e}")
        st.stop()
        
    # Crear selectbox para abolladuras
    # Asumimos que la columna de distancia es "Dist. Registro (km)"
    dist_col = "Dist. Registro (km)" if "Dist. Registro (km)" in df_input.columns else df_input.columns[0]
    
    dent_options = []
    for idx, row in df_input.iterrows():
        dist = row[dist_col]
        dent_options.append(f"Índice {idx} - Abscisa {dist} km")
        
    selected_dent_str = st.selectbox("Seleccione la Abolladura:", dent_options)
    selected_idx = int(selected_dent_str.split(" - ")[0].replace("Índice ", ""))
    selected_Lx = df_input.loc[selected_idx, dist_col]
    
    if st.button("▶ Ejecutar Análisis Rainflow en Abolladura", use_container_width=True):
        with st.spinner("Extrayendo topología y calculando ciclos…"):
            try:
                dent_dict, station_dict = rainflow.extract_topological_data(juntas_path, selected_Lx)
                
                analyzer = rainflow.DentSpectrumAnalyzer(specific_gravity, viscosity)
                
                cycles = analyzer.interpolate_pressure_timeseries(
                    scada_discharge_df=df_descarga,
                    scada_suction_df=df_succion,
                    dent_dict=dent_dict,
                    station_dict=station_dict,
                    time_col='timestamp',
                    pressure_col='presion_psi'
                )
                
                if cycles:
                    df_cycles = pd.DataFrame(cycles, columns=["Rango de Presión (psi)", "Conteo de Ciclos"])
                    
                    st.markdown(f"#### Espectro Rainflow (Interpolación Temporal) - Abscisa {selected_Lx/1000.0:.3f} km")
                    st.bar_chart(df_cycles.set_index("Rango de Presión (psi)"), use_container_width=True)
                    
                    st.markdown("#### Datos de Agrupación (25 Bins)")
                    st.dataframe(df_cycles, use_container_width=True)
                else:
                    st.warning("No se encontraron ciclos o no hubo solapamiento de tiempo en los SCADA.")
            except Exception as e:
                st.error(f"❌ Error durante el análisis Rainflow: {e}")
                
                # Proporcionar detalles adicionales del error
                with st.expander("🔍 Detalles técnicos del error"):
                    st.exception(e)
                    st.markdown("### Traza completa (Traceback)")
                    st.code(traceback.format_exc())
                    
                    # Sugerencias basadas en el tipo de error
                    if isinstance(e, KeyError):
                        st.warning(f"⚠️ **Sugerencia:** Falta la columna `{e}` en uno de los archivos CSV cargados. "
                                   "Verifique que `juntas_soldadura.csv` y los archivos SCADA tengan las columnas requeridas.")
                    elif isinstance(e, ValueError) and "nbins" in str(e):
                        st.info("ℹ️ **Sugerencia:** Hubo un problema al agrupar los datos en bins. Verifique que los datos de presión sean consistentes.")
                    elif "timestamp" in str(e).lower():
                        st.warning("⚠️ **Sugerencia:** Verifique que la columna de tiempo en los archivos SCADA sea correcta (ej: 'Fecha' o 'timestamp').")
                
    st.stop()



# ─── Preview de datos de entrada ─────────────────────────────────────────────
st.markdown('<p class="section-title">Vista previa · Datos de entrada</p>', unsafe_allow_html=True)

# Columnas clave para mostrar por defecto
key_col_indices = [0, 7, 4, 5, 8, 9, 12, 14, 15]
key_cols = [df_input.columns[i] for i in key_col_indices if i < n_cols_in]

if show_all_cols:
    df_preview = df_input
else:
    df_preview = df_input[key_cols] if key_cols else df_input

st.dataframe(
    df_preview,
    use_container_width=True,
    height=250,
)
st.caption(f"📊 {n_rows} anomalías cargadas · {n_cols_in} columnas")

st.markdown("---")

# ─── Botón de cálculo ─────────────────────────────────────────────────────────
st.markdown('<p class="section-title">Cálculo de Strain</p>', unsafe_allow_html=True)

col_btn, col_info = st.columns([1, 4])
with col_btn:
    run_btn = st.button("▶ Procesar Strain", use_container_width=True)

with col_info:
    st.markdown("""
    <div style="color:#64748b; font-size:0.85rem; margin-top:0.5rem;">
    Calcula la deformación de strain para cada abolladura usando el algoritmo de perfil de
    desplazamiento (polinomio de 6° grado). Aplica criterio API-1160: falla si |ε| ≥ 6%.
    </div>
    """, unsafe_allow_html=True)
df_input.to_csv("df_input.csv", index=False)

if run_btn or "df_result" in st.session_state:
    if run_btn:
        with st.spinner("Calculando strain para todas las anomalías…"):
            try:
                df_result = process_dataframe(df_input.copy())
                # Ensure unique index to avoid Styler errors
                df_result = df_result.reset_index(drop=True)
                st.session_state["df_result"] = df_result
            except Exception as e:
                st.error(f"❌ Error durante el cálculo: {e}")
                st.stop()
    else:
        df_result = st.session_state["df_result"]

    st.markdown("---")
    st.markdown('<p class="section-title">Resultados</p>', unsafe_allow_html=True)

    # ─── Métricas de resumen ──────────────────────────────────────────────────
    col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)

    total     = len(df_result)
    n_cumple  = (df_result["Dictamen_Strain"] == "Cumple criterio (strain < 6%)").sum()
    n_no_cumple = df_result["Dictamen_Strain"].str.contains("No cumple", na=False).sum()
    n_no_eval   = df_result["Dictamen_Strain"].str.contains("No evaluada", na=False).sum()
    n_errores   = df_result["Dictamen_Strain"].str.contains("faltante|Error", na=False).sum()

    with col_m1:
        st.metric("Total anomalías", total)
    with col_m2:
        st.metric("✅ Cumple criterio", int(n_cumple))
    with col_m3:
        st.metric("❌ No cumple criterio", int(n_no_cumple))
    with col_m4:
        st.metric("⬜ No evaluadas", int(n_no_eval))
    with col_m5:
        st.metric("⚠️ Datos faltantes", int(n_errores))

    st.markdown("<br>", unsafe_allow_html=True)

    # ─── DataFrame de resultados ──────────────────────────────────────────────
    # Columnas a mostrar: datos clave + resultados
    res_cols_to_show = key_cols + ["Strain_calc", "Dictamen_Strain"]
    res_cols_to_show = [c for c in res_cols_to_show if c in df_result.columns]

    df_display = df_result[res_cols_to_show].copy()

    # Formateo de strain como porcentaje visual
    if "Strain_calc" in df_display.columns:
        df_display["Strain (%)"] = df_display["Strain_calc"].apply(
            lambda x: f"{x * 100:.2f}%" if pd.notna(x) and x is not None else "—"
        )
        df_display = df_display.drop(columns=["Strain_calc"])

    # Aplicar estilos
    styled = (
        df_display.style
        .applymap(color_dictamen, subset=["Dictamen_Strain"])
        .applymap(
            lambda v: "color: #4ade80;" if isinstance(v, str) and v != "—" and "%" in v and abs(float(v.replace("%", "").replace(",", "."))) < 6
            else ("color: #f87171; font-weight:600;" if isinstance(v, str) and v != "—" and "%" in v else ""),
            subset=["Strain (%)"]
        )
        .set_properties(**{
            "font-size": "0.82rem",
        })
    )

    st.dataframe(styled, use_container_width=True, height=420)

    # ─── Tabla completa con todos los campos ──────────────────────────────────
    with st.expander("📋 Ver DataFrame completo (todas las columnas)"):
        # reset_index is redundant if df_result has it, but safe
        st.dataframe(
            df_result.reset_index(drop=True).style.applymap(color_dictamen, subset=["Dictamen_Strain"]),
            use_container_width=True,
            height=350,
        )

    # ─── Descarga de resultados ───────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<p class="section-title">Exportar resultados</p>', unsafe_allow_html=True)

    @st.cache_data
    def to_excel_bytes(df: pd.DataFrame) -> bytes:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Resultados_Strain")
            # Hoja resumen
            summary_data = {
                "Indicador": ["Total anomalías", "Cumple criterio", "No cumple criterio", "No evaluadas", "Datos faltantes"],
                "Cantidad": [total, int(n_cumple), int(n_no_cumple), int(n_no_eval), int(n_errores)],
                "Porcentaje (%)": [
                    "100%",
                    f"{n_cumple/total*100:.1f}%" if total else "0%",
                    f"{n_no_cumple/total*100:.1f}%" if total else "0%",
                    f"{n_no_eval/total*100:.1f}%" if total else "0%",
                    f"{n_errores/total*100:.1f}%" if total else "0%",
                ]
            }
            pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name="Resumen")
        return buf.getvalue()

    excel_bytes = to_excel_bytes(df_result)

    col_dl1, col_dl2 = st.columns([1, 4])
    with col_dl1:
        st.download_button(
            label="⬇ Descargar Excel",
            data=excel_bytes,
            file_name="resultados_strain.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_dl2:
        st.markdown("""
        <div style="color:#64748b; font-size:0.82rem; margin-top:0.6rem;">
        El archivo Excel incluye la hoja <b>Resultados_Strain</b> con todos los campos
        y una hoja <b>Resumen</b> con las estadísticas del dictamen.
        </div>
        """, unsafe_allow_html=True)
