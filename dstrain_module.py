"""
dstrain_module.py
=================
Traducción del código VBA (procedures.txt / helper.txt / keys.txt) a Python.

Implementa:
  - strain_f()       → Function StrainF() de VBA
  - is_dent()        → Sub is_dent() de VBA
  - is_repaired()    → Sub is_repaired() de VBA
  - classify_dent()  → lógica central de Sub dStrain()
  - process_dataframe() → aplica classify_dent() a todo el DataFrame
"""

import math
import pandas as pd
import numpy as np


# ---------------------------------------------------------------------------
# Mapeo de columnas (índice base-0 de pandas) según helper.txt → Sub start()
# Las columnas del Excel están en base-1; aquí se restan 1 para iloc base-0.
# ---------------------------------------------------------------------------
COL = {
    "Dregistro":   0,   # col 1
    "Latitud":     1,   # col 2
    "Longitud":    2,   # col 3
    "Altura":      3,   # col 4
    "Espesor":     4,   # col 5  → idth  (thickness mm)
    "SMYS":        5,   # col 6  → idy
    "SMTS":        6,   # col 7  → idu
    "DE":          7,   # col 8  → idde  (external diameter mm)
    "TipoAnomalia":8,   # col 9  → idAnomalyType
    "Comentario":  9,   # col 10 → idComment
    "PosicionPared":10, # col 11 → idpos
    "PosicionHoraria":11,# col 12 → idHourPosition
    "ProfPct":    12,   # col 13 → idp   (depth %, 0–100)
    "Longitud_mm":14,   # col 15 → idl   (length mm)
    "Ancho_mm":   15,   # col 16 → idwth (width mm)
    "NumSoldadura":16,  # col 17
    "DSoldInf":   17,   # col 18 → idja (distancia soldadura circunferencial aguas arriba mm)
    "DSoldSup":   18,   # col 19 → idjp (distancia soldadura circunferencial aguas abajo mm)
}

# Columnas de resultado que se añaden al DataFrame
RESULT_COLS = ["Strain_calc", "Dictamen_Strain"]


# ---------------------------------------------------------------------------
# Funciones auxiliares de clasificación de anomalías
# ---------------------------------------------------------------------------

def is_dent(anomaly_type: str) -> bool:
    """
    Equivalente a Sub is_dent() en keys.txt.
    Devuelve True si el tipo de anomalía corresponde a una abolladura.
    """
    if pd.isna(anomaly_type):
        return False
    s = str(anomaly_type).upper()
    return any(kw in s for kw in ["ABOL", "LLADUR", "DENT", "DIÁM", "DIAM", "INTER"])


def is_repaired(comment: str) -> bool:
    """
    Equivalente a Sub is_repaired() en keys.txt.
    Devuelve True si el comentario indica que la anomalía fue reparada.
    """
    if pd.isna(comment):
        return False
    s = str(comment).upper()
    if "REPARADA" in s or "REPARADO" in s or "REPAIRED" in s:
        return True
    if "BAJO" in s and "REPAR" in s:
        return True
    if "UNDER" in s and "REP" in s:
        return True
    if "REPA" in s and "PREVIA" in s:
        return True
    return False


def is_kinked_dent(ro: float, li: float, w: float, t: float, dti: float) -> bool:
    """
    Determina si una abolladura es plegada (kinked dent) de acuerdo con los criterios
    de los radios de curvatura transversal y longitudinal.

    Parámetros
    ----------
    ro  : Radio interno en mm (normalmente de / 2)
    li  : Longitud de la abolladura en mm
    w   : Ancho de la abolladura en mm
    t   : Espesor de pared en mm
    dti : Profundidad de la abolladura en mm
    
    Retorna
    -------
    True si la abolladura es plegada, False en caso contrario.
    """
    if ro <= 0 or w <= 0 or dti <= 0 or t <= 0:
        return False
        
    x = (w / 2) / ro
    if abs(x) >= 1.0:
        return False

    acrsen = math.asin(x)
    teta_with = 2 * math.degrees(acrsen)
    
    h = math.cos(math.radians(teta_with / 2)) * ro
    val_b = (ro ** 2) - (h ** 2)
    if val_b < 0:
        val_b = 0
    B = val_b ** 0.5
    c = 2 * B
    hd = ro - h
    hab = dti - hd

    if hab == 0:
        return False

    R1 = abs((hab / 2) + ((c ** 2) / (8 * hab)))
    R2 = abs((dti / 2) + ((li ** 2) / (8 * dti)))

    if R1 <= (5 * t) or R2 <= (5 * t):
        return True

    return False


# ---------------------------------------------------------------------------
# Algoritmo de deformación: StrainF()
# ---------------------------------------------------------------------------

def strain_f(de: float, dti: float, li: float, W: float, t: float) -> float:
    """
    Calcula el strain en el ápice de la abolladura según ASME B31.8 Apéndice R.

    Parámetros
    ----------
    de  : Diámetro externo (mm)
    dti : Profundidad relativa de la abolladura (0.0–100)
    li  : Longitud de la abolladura (mm)
    W   : Ancho de la abolladura (mm)
    t   : Espesor de pared (mm)

    Retorna
    -------
    Deformación combinada máxima (fracción, ej. 0.04 = 4%)
    """
    if de <= 0 or t <= 0 or dti <= 0 or W <= 0 or li <= 0:
        return 0.0

    # ---------------------------------------------------------------------------------
    # INTEGRACIÓN DE API 1183 (Sección 7.2) / ASME B31.8 Apéndice R
    # ---------------------------------------------------------------------------------
    # Profundidad absoluta de la abolladura en mm
    d_mm = (dti / 100.0) * de
    
    # Radio un-deformed de la tubería
    R0 = de / 2.0
    
    # Radios de curvatura en el ápice
    # R1: Radio transversal (asumido positivo, se conserva curvatura original de superficie externa)
    R1 = +(W ** 2) / (8.0 * d_mm)
    
    # R2: Radio longitudinal
    R2 = (li ** 2) / (8.0 * d_mm)

    # 1. Strain por Flexión Circunferencial (e1) - superficie externa
    e1_ext = (t / 2.0) * ((1.0 / R0) - (1.0 / R1))
    e1_int = -e1_ext

    # 2. Strain por Flexión Longitudinal (e2) - superficie externa
    e2_ext = (t / 2.0) * (1.0 / R2)
    e2_int = -e2_ext
    
    # 3. Strain de Membrana Longitudinal (e3)
    e3 = 0.5 * ((d_mm / li) ** 2)
        
    # 4. Strain Combinado (Ecuación derivada de Von Mises combinada - API 1183)
    strain_ext = (2.0 / math.sqrt(3.0)) * math.sqrt(e1_ext**2 + e1_ext * (e2_ext + e3) + (e2_ext + e3)**2)
    strain_int = (2.0 / math.sqrt(3.0)) * math.sqrt(e1_int**2 + e1_int * (e2_int + e3) + (e2_int + e3)**2)
    
    # Retorna el máximo valor absoluto
    return max(strain_ext, strain_int)


# ---------------------------------------------------------------------------
# Clasificación de una fila (lógica de dStrain)
# ---------------------------------------------------------------------------

def classify_dent(row: pd.Series) -> dict:
    """
    Aplica la lógica de Sub dStrain() para una fila del DataFrame.

    Retorna un dict con:
      - 'Strain_calc' : valor numérico (float) o None
      - 'Dictamen_Strain' : texto de dictamen (str)
    """
    anomaly_type = row.iloc[COL["TipoAnomalia"]]
    comment      = row.iloc[COL["Comentario"]]
    de   = _to_float(row.iloc[COL["DE"]])
    dti  = _to_float(row.iloc[COL["ProfPct"]])
    li   = _to_float(row.iloc[COL["Longitud_mm"]])
    W    = _to_float(row.iloc[COL["Ancho_mm"]])
    t    = _to_float(row.iloc[COL["Espesor"]])

    # ¿Es una abolladura?
    if not is_dent(anomaly_type):
        return {"Strain_calc": None, "Dictamen_Strain": "No evaluada"}

    # ¿Fue reparada?
    if is_repaired(comment):
        return {"Strain_calc": None, "Dictamen_Strain": "No evaluada (Reparada)"}

    # Validación de datos requeridos
    if t == 0 or de == 0 or dti == 0 or li == 0:
        return {"Strain_calc": None, "Dictamen_Strain": "Valor faltante o incorrecto"}

    # Revisión de abolladura plegada
    dti_mm = (dti / 100.0) * de
    if is_kinked_dent(de / 2.0, li, W, t, dti_mm):
        return {"Strain_calc": None, "Dictamen_Strain": "Abolladura plegada - requiere FFS"}

    # Interacción con Soldaduras (API 1183 - Sección 6.5.1.1)
    clock_pos = _parse_clock_position(row.iloc[COL["PosicionHoraria"]])
    dist_girth = _get_girth_weld_dist(row.iloc[COL["DSoldInf"]], row.iloc[COL["DSoldSup"]])
    weld_interaction = check_weld_interaction(de, dist_girth, clock_pos)

    # Cálculo de strain
    try:
        strain_val = round(strain_f(de, dti, li, W, t), 4)
    except Exception as e:
        return {"Strain_calc": None, "Dictamen_Strain": f"Error de cálculo: {e}"}

    # Dictamen según criterio API-1183 (Sección 7.2) / ASME B31.8
    # NOTA: Como en todos los casos de evaluación se desconoce la elongación específica 
    # del material (MTRs no disponibles), se deberá usar el criterio del 6% para tubería base,
    # y estrictamente del 4% (strain <= 4%) si la abolladura interactúa con una soldadura.
    if weld_interaction["interacts_girth"]:
        # Límite máximo permitido de deformación es estrictamente del 4% si interactúa con soldadura
        if abs(strain_val) >= 0.04:
            dictamen = "No cumple criterio (strain ≥ 4%)"
        else:
            dictamen = "Cumple criterio (strain < 4%)"
        dictamen += " | Interactúa con Soldadura Girth"
    else:
        # Límite máximo permitido del 6% para cuerpo de tubo
        if abs(strain_val) >= 0.06:
            dictamen = "No cumple criterio (strain ≥ 6%)"
        else:
            dictamen = "Cumple criterio (strain < 6%)"

    return {"Strain_calc": strain_val, "Dictamen_Strain": dictamen}


# ---------------------------------------------------------------------------
# Procesamiento de DataFrame completo
# ---------------------------------------------------------------------------

def process_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe el DataFrame de la hoja EntradaDatos (datos desde fila 8 del Excel)
    y retorna el mismo DataFrame con columnas de resultado añadidas:
      - Strain_calc
      - Dictamen_Strain
    """
    results = df.apply(classify_dent, axis=1, result_type="expand")
    df_out = df.copy()
    df_out["Strain_calc"]     = results["Strain_calc"]
    df_out["Dictamen_Strain"] = results["Dictamen_Strain"]
    return df_out


# ---------------------------------------------------------------------------
# Helper interno
# ---------------------------------------------------------------------------

def _to_float(val) -> float:
    """Convierte un valor de celda a float de forma segura; retorna 0.0 si falla."""
    if pd.isna(val) or val == "" or val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def _parse_clock_position(val) -> float:
    """Convierte posición horaria a float (ej: '12:30' -> 12.5)."""
    if pd.isna(val) or val == "" or val is None:
        return 0.0
    try:
        val_str = str(val).strip()
        if ":" in val_str:
            parts = val_str.split(":")
            h = float(parts[0])
            m = float(parts[1]) if len(parts) > 1 else 0.0
            return h + (m / 60.0)
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def _get_girth_weld_dist(dsoldinf, dsoldsup) -> float:
    """Obtiene la distancia mínima absoluta a la soldadura circunferencial.
    Retorna float('inf') si ambos valores están vacíos o nulos.
    """
    def is_empty(v):
        return pd.isna(v) or str(v).strip() == "" or v is None
        
    inf_empty = is_empty(dsoldinf)
    sup_empty = is_empty(dsoldsup)
    
    if inf_empty and sup_empty:
        return float('inf')
        
    try:
        v_inf = float(dsoldinf) if not inf_empty else float('inf')
    except:
        v_inf = float('inf')
        
    try:
        v_sup = float(dsoldsup) if not sup_empty else float('inf')
    except:
        v_sup = float('inf')
        
    if v_inf == float('inf') and v_sup == float('inf'):
        return float('inf')
        
    return min(abs(v_inf), abs(v_sup))

def check_weld_interaction(de_mm: float, dist_girth_weld_mm: float, clock_pos: float) -> dict:
    """
    Evalúa la interacción de una abolladura con soldadura circunferencial (Girth Weld) 
    según API 1183 (2020) Sección 6.5.1.1. 
    Se omite la evaluación de soldadura longitudinal (Seam Weld) al desconocer su posición.
    """
    # Restringida: Fondo de la tubería (entre las 4 y las 8 en punto)
    # No restringida: Parte superior
    is_restrained = (4.0 <= clock_pos <= 8.0)
    
    if is_restrained:
        a, b = 0.418, 94.6
    else:
        a, b = 0.129, 109.6
    
    dc_threshold = a * de_mm + b
    interacts_girth = dist_girth_weld_mm <= dc_threshold
    
    return {
        "interacts_girth": interacts_girth,
        "dc_threshold": dc_threshold,
        "is_restrained": is_restrained
    }

def calcular_screening_fatiga(
    df_rainflow: pd.DataFrame,
    de_mm: float,
    t_mm: float,
    smys_psi: float,
    dti_pct: float,
    clock_pos: float,
    time_span_years: float = 1.0
) -> float | str:
    """
    Calcula la Vida a la Fatiga por Filtrado (Screening) según API 1183 (2020) Cap. 7.4.3.
    
    Parámetros:
    -----------
    df_rainflow     : pd.DataFrame con columnas 'Rango de Presión (psi)' y 'Conteo de Ciclos'
    de_mm           : Diámetro externo en mm
    t_mm            : Espesor de pared en mm
    smys_psi        : SMYS en psi
    dti_pct         : Profundidad relativa de abolladura (%)
    clock_pos       : Posición horaria general
    time_span_years : Duración temporal real del espectro de presión SCADA analizado (en años).
                      Ejemplo: 1.0 si se procesó un año completo, 0.25 si fue un trimestre.
                      Valor por defecto: 1.0

    Retorna:
    --------
    Vida remanente en años (float), o el string 'Requiere FFS' si no aplica.
    """
    if df_rainflow is None or df_rainflow.empty:
        return "Requiere FFS"

    if de_mm <= 0 or t_mm <= 0 or smys_psi <= 0:
        return "Requiere FFS"

    # Paso A: Validación de Aplicabilidad
    is_restrained = (4.0 <= clock_pos <= 8.0)
    de_in = de_mm / 25.4
    
    if de_in <= 12.75:
        is_shallow = dti_pct < 4.0
    else:
        is_shallow = dti_pct < 2.5
        
    if is_restrained and not is_shallow:
        return "Requiere FFS"

    # Paso B: Cálculo de K_M_Max
    od_t = de_mm / t_mm
    y = od_t
    
    delta_p = df_rainflow['Rango de Presión (psi)'].values
    ciclos = df_rainflow['Conteo de Ciclos'].values
    
    delta_p_smys = (2.0 * t_mm * smys_psi) / de_mm
    x = delta_p / delta_p_smys  # Fracción de rango de presión / presión de cedencia

    if is_restrained:
        k_m_max = np.full_like(delta_p, 0.1183 * od_t - 1.146)
        k_m_max = np.maximum(k_m_max, 1.0)
    else:
        # Eq 16 API 1183
        a00 = 6.61847
        a10 = -12.26386
        a01 = 0.06748
        a20 = 15.58507
        a11 = -0.12358
        a02 = -0.00032
        a30 = -8.58441
        a21 = 0.03803
        a12 = 0.00047
        
        k_m_max = (
            a00 + 
            a10 * x + 
            a01 * y + 
            a20 * (x**2) + 
            a11 * x * y + 
            a02 * (y**2) + 
            a30 * (x**3) + 
            a21 * (x**2) * y + 
            a12 * x * (y**2)
        )
        k_m_max = np.maximum(k_m_max, 1.0)
        
    # Paso C: Esfuerzos Críticos (Barlow)
    delta_sigma_hoop_psi = delta_p * de_mm / (2.0 * t_mm)
    delta_sigma_peak_psi = delta_sigma_hoop_psi * k_m_max
    
    # Esfuerzo a ksi
    delta_sigma_peak_ksi = delta_sigma_peak_psi / 1000.0
    delta_sigma_peak_ksi = np.where(delta_sigma_peak_ksi <= 0, 1e-9, delta_sigma_peak_ksi)
    
    # Paso D: Daño por Fatiga y Regla de Miner (Eq 12)
    m = 3.0
    log10_C = 10.08514
    
    log10_n_falla = log10_C - m * np.log10(delta_sigma_peak_ksi)
    n_falla = 10 ** log10_n_falla
    
    d_i = ciclos / n_falla
    d_total = np.sum(d_i)
    
    if d_total > 0:
        # Vida remanente = (1 / D_total) × Time Span del espectro analizado
        # D_total es el daño acumulado por unidad de espectro; al multiplicar por
        # time_span_years se obtiene directamente la vida remanente en años.
        vida_remanente = (1.0 / d_total) * time_span_years
        return float(vida_remanente)
    else:
        return float('inf')
