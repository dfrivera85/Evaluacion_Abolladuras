"""
Implements rainflow cycle counting algorythm for fatigue analysis
according to section 5.4.4 in ASTM E1049-85 (2011).
"""
from __future__ import division
from collections import deque, defaultdict
import math

__version__ = "3.2.0"

import pandas as pd
import numpy as np
def _get_round_function(ndigits=None):
    if ndigits is None:
        def func(x):
            return x
    else:
        def func(x):
            return round(x, ndigits)
    return func


def reversals(series):
    """Iterate reversal points in the series.

    A reversal point is a point in the series at which the first derivative
    changes sign. Reversal is undefined at the first (last) point because the
    derivative before (after) this point is undefined. The first and the last
    points are treated as reversals.

    Parameters
    ----------
    series : iterable sequence of numbers

    Yields
    ------
    Reversal points as tuples (index, value).
    """
    series = iter(series)

    x_last, x = next(series, None), next(series, None)
    if x_last is None or x is None:
        return

    d_last = (x - x_last)

    yield 0, x_last
    index = None
    for index, x_next in enumerate(series, start=1):
        if x_next == x:
            continue
        d_next = x_next - x
        if d_last * d_next < 0:
            yield index, x
        x_last, x = x, x_next
        d_last = d_next

    if index is not None:
        yield index + 1, x_next


def extract_cycles(series):
    """Iterate cycles in the series.

    Parameters
    ----------
    series : iterable sequence of numbers

    Yields
    ------
    cycle : tuple
        Each tuple contains (range, mean, count, start index, end index).
        Count equals to 1.0 for full cycles and 0.5 for half cycles.
    """
    points = deque()

    def format_output(point1, point2, count):
        i1, x1 = point1
        i2, x2 = point2
        rng = abs(x1 - x2)
        mean = 0.5 * (x1 + x2)
        return rng, mean, count, i1, i2

    for point in reversals(series):
        points.append(point)

        while len(points) >= 3:
            # Form ranges X and Y from the three most recent points
            x1, x2, x3 = points[-3][1], points[-2][1], points[-1][1]
            X = abs(x3 - x2)
            Y = abs(x2 - x1)

            if X < Y:
                # Read the next point
                break
            elif len(points) == 3:
                # Y contains the starting point
                # Count Y as one-half cycle and discard the first point
                yield format_output(points[0], points[1], 0.5)
                points.popleft()
            else:
                # Count Y as one cycle and discard the peak and the valley of Y
                yield format_output(points[-3], points[-2], 1.0)
                last = points.pop()
                points.pop()
                points.pop()
                points.append(last)
    else:
        # Count the remaining ranges as one-half cycles
        while len(points) > 1:
            yield format_output(points[0], points[1], 0.5)
            points.popleft()


def count_cycles(series, ndigits=None, nbins=None, binsize=None):
    """Count cycles in the series.

    Parameters
    ----------
    series : iterable sequence of numbers
    ndigits : int, optional
        Round cycle magnitudes to the given number of digits before counting.
        Use a negative value to round to tens, hundreds, etc.
    nbins : int, optional
        Specifies the number of cycle-counting bins.
    binsize : int, optional
        Specifies the width of each cycle-counting bin

    Arguments ndigits, nbins and binsize are mutually exclusive.

    Returns
    -------
    A sorted list containing pairs of range and cycle count.
    The counts may not be whole numbers because the rainflow counting
    algorithm may produce half-cycles. If binning is used then ranges
    correspond to the right (high) edge of a bin.
    """
    if sum(value is not None for value in (ndigits, nbins, binsize)) > 1:
        raise ValueError(
            "Arguments ndigits, nbins and binsize are mutually exclusive"
        )

    counts = defaultdict(float)
    cycles = (
        (rng, count)
        for rng, mean, count, i_start, i_end in extract_cycles(series)
    )

    if nbins is not None:
        binsize = (max(series) - min(series)) / nbins

    if binsize is not None:
        nmax = 0
        for rng, count in cycles:
            quotient = rng / binsize
            n = int(math.ceil(quotient))  # using int for Python 2 compatibility

            if nbins and n > nbins:
                # Due to floating point accuracy we may get n > nbins,
                # in which case we move rng to the preceeding bin.
                if (quotient % 1) > 1e-6:
                    raise Exception("Unexpected error")
                n = n - 1

            counts[n * binsize] += count
            nmax = max(n, nmax)

        for i in range(1, nmax):
            counts.setdefault(i * binsize, 0.0)

    elif ndigits is not None:
        round_ = _get_round_function(ndigits)
        for rng, count in cycles:
            counts[round_(rng)] += count

    else:
        for rng, count in cycles:
            counts[rng] += count

    return sorted(counts.items())


def extract_topological_data(juntas_csv_path, Lx):
    """
    Extrae la información topológica requerida a partir del archivo juntas.csv basándose
    en la abscisa de la abolladura (Lx).
    Devuelve los diccionarios:
      - dent_dict: Lx, hx, D1, D2
      - station_dict: L1, h1, L2, h2
    """
    try:
        df = pd.read_csv(juntas_csv_path)
    except UnicodeDecodeError:
        df = pd.read_csv(juntas_csv_path, encoding='latin1')
    
    # L1 y h1: estación de bombeo aguas arriba
    idx_min = df['distancia_inicio_m'].idxmin()
    L1 = df.loc[idx_min, 'distancia_inicio_m']
    h1 = df.loc[idx_min, 'altura_m']
    
    # L2 y h2: estación de recibo aguas abajo
    idx_max = df['distancia_inicio_m'].idxmax()
    L2 = df.loc[idx_max, 'distancia_inicio_m']
    h2 = df.loc[idx_max, 'altura_m']
    
    # Diámetros
    D1 = df.loc[df['distancia_inicio_m'] < Lx, 'diametro'].mean()
    D2 = df.loc[df['distancia_inicio_m'] > Lx, 'diametro'].mean()
    if pd.isna(D1): D1 = df['diametro'].mean()
    if pd.isna(D2): D2 = df['diametro'].mean()
    
    # hx
    dent_row = df[(df['distancia_inicio_m'] <= Lx) & (df['distancia_fin_m'] >= Lx)]
    if not dent_row.empty:
        hx = dent_row.iloc[0]['altura_m']
    else:
        closest_idx = (df['distancia_inicio_m'] - Lx).abs().idxmin()
        hx = df.loc[closest_idx, 'altura_m']
        
    dent_dict = {'Lx': Lx, 'hx': hx, 'D1': D1, 'D2': D2}
    station_dict = {'L1': L1, 'h1': h1, 'L2': L2, 'h2': h2}
    
    pd.DataFrame([dent_dict]).to_csv('dent_dict.csv', index=False)
    pd.DataFrame([station_dict]).to_csv('station_dict.csv', index=False)

    return dent_dict, station_dict



class DentSpectrumAnalyzer:
    """
    Analizador del espectro de presión y análisis Rainflow en una abolladura (API 1183).
    """
    def __init__(self, specific_gravity, viscosity):
        """
        Inicializa el analizador configurando propiedades del fluido.
        """
        self.specific_gravity = specific_gravity
        self.viscosity = viscosity
        
        self.K = self.specific_gravity * 0.433

    def _merge_scada(self, df_discharge, df_suction, time_col, pressure_col):
        """
        Une los dos DataFrames de SCADA basándose en el tiempo o el tiempo más cercano
        dentro de una ventana máxima de 5 minutos.
        """
        df1 = df_discharge.copy()
        df2 = df_suction.copy()
        
        # Asegurar tipos datetime
        df1[time_col] = pd.to_datetime(df1[time_col])
        df2[time_col] = pd.to_datetime(df2[time_col])
        
        # Ordenar (necesario para merge_asof)
        df1 = df1.sort_values(time_col)
        df2 = df2.sort_values(time_col)
        
        merged = pd.merge_asof(
            df1, df2, 
            on=time_col, 
            direction='nearest', 
            tolerance=pd.Timedelta('5min'),
            suffixes=('_discharge', '_suction')
        )
        # Limpiar filas donde no hubo coincidencia en tiempo
        merged_clean = merged.dropna(subset=[f'{pressure_col}_discharge', f'{pressure_col}_suction'])
        
        # Si las filas sin coincidencia superan el 50% de df1, retornar df1 (enfoque conservador)
        if (len(df1) - len(merged_clean)) > 0.5 * len(df1):
            merged_conservative = df1.copy()
            merged_conservative = merged_conservative.rename(columns={pressure_col: f'{pressure_col}_discharge'})
            merged_conservative[f'{pressure_col}_suction'] = merged_conservative[f'{pressure_col}_discharge']
            return merged_conservative

        return merged_clean

    def interpolate_pressure_timeseries(self, scada_discharge_df, scada_suction_df, dent_dict, station_dict, time_col='timestamp', pressure_col='pressure_psi'):
        """
        Aplica el Enfoque A (Ecuación 5 - Interpolación en el dominio del tiempo).
        Estima el Px (presión en la abolladura) y devuelve una tupla: (conteos de 25 bins, time_span_years).
        """
        merged = self._merge_scada(scada_discharge_df, scada_suction_df, time_col, pressure_col)
        if merged.empty:
            return [], 0.0
            
        P1 = merged[f'{pressure_col}_discharge'].values
        P2 = merged[f'{pressure_col}_suction'].values
        
        pd.DataFrame([P1]).to_csv('P1.csv', index=False)
        pd.DataFrame([P2]).to_csv('P2.csv', index=False)


        L1, h1 = station_dict['L1'], station_dict['h1']
        L2, h2 = station_dict['L2'], station_dict['h2']
        Lx, hx = dent_dict['Lx'], dent_dict['hx']
        D1, D2 = dent_dict['D1'], dent_dict['D2']
        
        # Convertir L1, L2, Lx, h1, h2 y hx a pies (1 m = 3.28084 ft)
        m_to_ft = 3.28084
        L1, h1 = L1 * m_to_ft, h1 * m_to_ft
        L2, h2 = L2 * m_to_ft, h2 * m_to_ft
        Lx, hx = Lx * m_to_ft, hx * m_to_ft
        
        # Ajustar D1 y D2 a pulgadas (1 mm = 0.0393701 in)
        mm_to_in = 0.0393701
        D1, D2 = D1 * mm_to_in, D2 * mm_to_in

        # (Lx - L1) * D2^5 / ((L2 - Lx) * D1^5)
        num = (Lx - L1) * (D2**5)
        den = (L2 - Lx) * (D1**5)
        
        if den == 0:
            factor = 0.0
        else:
            denom_full = (num / den) + 1.0
            factor = 1.0 / denom_full if denom_full != 0 else 0.0
            
        # Px = (P1 + K*h1 - P2 - K*h2)*(factor) - K*(hx - h2) + P2
        Px = (P1 + self.K * h1 - P2 - self.K * h2) * factor - self.K * (hx - h2) + P2
        
        pd.DataFrame([Px]).to_csv('Px.csv', index=False)
        
        # Calcular la duración (Time Span) en años
        time_diff = merged[time_col].max() - merged[time_col].min()
        time_span_years = time_diff.total_seconds() / (365.25 * 24 * 3600)
        
        if time_span_years <= 0:
            time_span_years = 1.0  # fallback si hay un solo registro o error

        # Aplicar el conteo Rainflow y retornar los 25 bins resultantes junto con el time_span
        return count_cycles(Px, nbins=25), float(time_span_years)

    def interpolate_rainflow_cycles(self, scada_discharge_df, scada_suction_df, dent_dict, station_dict, pressure_col='pressure_psi'):
        """
        Aplica el Enfoque B (Ecuación 6 - Interpolación de los ciclos Rainflow).
        Fuerza 25 bins constantes tanto en P1 como P2, y luego calcula los ciclos atenuados.
        """
        P1 = scada_discharge_df[pressure_col].values
        P2 = scada_suction_df[pressure_col].values
        
        if len(P1) == 0 or len(P2) == 0:
            return []
            
        # Para forzar exactamente los mismos ranges (bins), determinamos el rango máximo posible.
        max_rng_P1 = P1.max() - P1.min() if len(P1) > 0 else 0
        max_rng_P2 = P2.max() - P2.min() if len(P2) > 0 else 0
        global_max_rng = max(max_rng_P1, max_rng_P2)
        
        if global_max_rng <= 0:
            return []
            
        # Al forzar binsize = global_max_rng / 25, garantizamos 25 bins en ambos conteos.
        binsize = global_max_rng / 25.0
        
        res_D = dict(count_cycles(P1, binsize=binsize))
        res_S = dict(count_cycles(P2, binsize=binsize))
        
        # Parámetros topológicos
        L1, L2, Lx = station_dict['L1'], station_dict['L2'], dent_dict['Lx']
        di = Lx - L1
        ds = L2 - L1
        
        # Constantes de viscosidad
        if self.viscosity <= 100:
            a, b, c_const, d_const = 1.048, 0.858, 0.993, 0.81
        else:
            a, b, c_const, d_const = 1.150, 0.750, 1.200, 1.600
            
        term1 = (a * b)**(ds / di) if ds != 0 else 0.0
        
        NI_bins = {}
        # Siempre devolver exactamente 25 bins
        for n in range(1, 26):
            bin_val = n * binsize
            ND = res_D.get(bin_val, 0.0)
            NS = res_S.get(bin_val, 0.0)
            
            # Interpolación (Ecuación 6)
            NI = ND - term1 * (c_const * ND - d_const * NS)
            NI_bins[bin_val] = max(0.0, NI)
            
        return sorted(NI_bins.items())
