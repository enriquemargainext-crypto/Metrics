import math
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import os
import sys
import time
import cProfile
import pstats
import io

import numpy as np
import pandas as pd


try:
    import streamlit as st
except Exception:  # pragma: no cover
    # Allow running without Streamlit for pure function tests
    st = None  # type: ignore


############################################################
# Configuración inicial
############################################################

# Ruta configurable al archivo Excel. Puedes cambiar esta variable o usar la UI.
DEFAULT_EXCEL_PATH = Path("202506_Financials_by_Country.xlsx")

# Hoja por defecto (país). Puedes cambiar esta variable o usar la UI.
DEFAULT_COUNTRY_SHEET = "Mexico Consolidated"


############################################################
# Mapeo de etiquetas preferidas por métrica
############################################################

# Para cada métrica, lista de etiquetas candidatas en orden de preferencia.
# Buscamos primero en columna C (índice 2) y luego en D (índice 3).
PREFERRED_LABELS_BY_KEY: Dict[str, List[str]] = {
    # Unidades
    "sales": [
        "Total Cars Delivered",
        "Cars Sold - Delivered",
    ],
    "purchases": [
        "Cars Purchased",
    ],
    "inventory_bop": [
        "Inventory BoM",
    ],

    # Montos en millones
    "total_net_revenues_mn": [
        "Total Net Revenues",
    ],
    "metal_margin_mn": [
        "Metal Margin (mm)",
    ],
    "other_revenues_gp_mn": [
        "Gross Profit from Other Revenues (mm)",
    ],
    "pc1_mn": [
        "PC1 (mm)",
    ],
    "ecac_mn": [
        "Performance Marketing",
    ],
    "sga_mn": [
        "SG&A",
    ],
    "it_mn": [
        "IT",
    ],
    "warehousing_mn": [
        "Warehousing",
    ],
    "ebitda_mn": [
        "ebitda",        # usar solo EBITDA (sin Adj.)
    ],
    "burn_mn": [
        "cash flow before financing and working capital adjustments",
        "cash flow before financing and capital adjustments",
    ],
}

# Alias de claves internas a nombres mostrables en el dashboard/CSV
ALIAS_KEYS: Dict[str, str] = {
    "burn_mn": "burn",
    "ebitda_mn": "ebitda",
    "ecac_mn": "Performance Marketing",
    # Alias para métricas derivadas (nombres visibles en la tabla/CSV)
    "ecac_unit": "ecac",
    "pc1_ecac_unit": "PC1-ecac",
}


DERIVED_FORMULAS_ORDER = [
    "ecac_unit",
    "sales_efficiency_pct",
    "pc1_ecac_unit",
    "sga_unit",
    "it_unit",
    "warehousing_unit",
]

# Métricas base que deben ocultarse en la tabla/export (pero sí usarse para derivadas)
HIDE_BASE_FROM_DISPLAY = {
    "ecac_mn",
}


############################################################
# Utilidades
############################################################

class Profiler:
    """Perfilado ligero activable por variable de entorno PROFILE.
    - Timings agregados en milisegundos por etapa
    - cProfile del top 20 por tiempo acumulado
    """

    def __init__(self, enabled: bool) -> None:
        self.enabled = enabled
        self.timings_ms: Dict[str, float] = {}
        self._cprof: Optional[cProfile.Profile] = cProfile.Profile() if enabled else None

    def add_ms(self, stage: str, delta_ms: float) -> None:
        if not self.enabled:
            return
        self.timings_ms[stage] = self.timings_ms.get(stage, 0.0) + float(delta_ms)

    def start_cprofile(self) -> None:
        if self.enabled and self._cprof is not None:
            self._cprof.enable()

    def stop_cprofile(self) -> None:
        if self.enabled and self._cprof is not None:
            self._cprof.disable()

    def get_cprofile_report(self, limit: int = 20) -> str:
        if not self.enabled or self._cprof is None:
            return ""
        s = io.StringIO()
        ps = pstats.Stats(self._cprof, stream=s).sort_stats(pstats.SortKey.CUMULATIVE)
        ps.print_stats(limit)
        return s.getvalue()

    def format_timings(self) -> str:
        if not self.enabled:
            return ""
        # Orden sugerido de etapas
        order = [
            "lectura",
            "deteccion_header",
            "mapeo_etiquetas",
            "extraccion_valores",
            "calculo_derivados",
            "render",
        ]
        lines: List[str] = ["Timings por etapa (ms):"]
        for key in order:
            if key in self.timings_ms:
                lines.append(f"  - {key}: {self.timings_ms[key]:.2f} ms")
        # Cualquier otra clave adicional
        for key, val in self.timings_ms.items():
            if key not in order:
                lines.append(f"  - {key}: {val:.2f} ms")
        return "\n".join(lines)

def _normalize_label(text: str) -> str:
    """Normaliza guiones Unicode a ASCII y recorta espacios.
    Esto mejora la robustez frente a '(–)' vs '(-)'.
    """
    if not isinstance(text, str):
        return ""
    normalized = (
        text.replace("\u2013", "-")  # en dash
        .replace("\u2014", "-")  # em dash
        .replace("\u2212", "-")  # minus
        .strip()
    )
    # Comparación case-insensitive
    return normalized.lower()


def _safe_div(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    """Divide elemento a elemento devolviendo NaN cuando el denominador es 0 o NaN.
    Si el denominador es cero o nulo, devolver NaN (luego se convertirá a None en salida final).
    """
    result = numerator.astype(float) / denominator.astype(float)
    result = result.where((denominator != 0) & (~denominator.isna()))
    return result


def _compute_growth_pct(series: pd.Series, months_lag: int) -> pd.Series:
    """(valor actual / valor de hace N meses - 1) * 100.
    Devuelve NaN cuando no hay suficiente historia o el denominador es 0/NaN.
    """
    prev = series.shift(months_lag)
    with np.errstate(divide="ignore", invalid="ignore"):
        growth = (series / prev - 1.0) * 100.0
    growth = growth.where(prev.notna() & (prev != 0))
    return growth


############################################################
# Extracción de datos desde Excel
############################################################

def read_sheet_as_dataframe(
    excel_path: Path, sheet_name: str, profiler: Optional[Profiler] = None
) -> Tuple[pd.DataFrame, List[pd.Timestamp], List[int]]:
    """Lee una hoja del Excel con header sin procesar y devuelve:
    - df: DataFrame sin cabeceras (header=None) para permitir indexación fija
    - periods: lista de periodos (fechas fin de mes) extraídos de la FILA 10 (índice 9)

    Supuestos del archivo:
    - La fila de encabezados de periodos está en fila 10 (índice 9)
    - Las etiquetas de métricas están en la columna C (índice 2) y columna D (índice 3)
    - Los valores de cada métrica están desde la columna J (índice 9) en adelante
    """
    t0 = time.perf_counter()
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, engine="openpyxl")
    if profiler is not None:
        profiler.add_ms("lectura", (time.perf_counter() - t0) * 1000.0)

    # Fila 10 (índice 9) contiene los periodos con formato "Mon-YY" (e.g., Jun-25)
    # Solo tomar la primera sección consecutiva de periodos, no deltas u otras columnas
    header_row_idx = 9
    t1 = time.perf_counter()
    header_row = df.iloc[header_row_idx, :]
    valid_periods: List[pd.Timestamp] = []
    period_col_indices: List[int] = []
    
    # Detectar solo la primera sección consecutiva de periodos
    for col_idx, cell in enumerate(header_row):
        parsed: Optional[pd.Timestamp] = None
        if isinstance(cell, str):
            text = cell.strip()
            # Intentar formato "Mon-YY" primero
            try:
                parsed = pd.to_datetime(text, format="%b-%y", errors="raise")
            except Exception:
                # Si no es formato Mon-YY, parar aquí (probablemente llegamos a deltas)
                if len(valid_periods) > 0:  # Ya tenemos algunos periodos
                    break
                # Si no tenemos periodos aún, intentar parseo genérico
                try:
                    parsed = pd.to_datetime(text, errors="coerce")
                except Exception:
                    continue
        else:
            try:
                parsed = pd.to_datetime(cell, errors="coerce")
            except Exception:
                continue
                
        if pd.notna(parsed):
            parsed_eom = pd.DatetimeIndex([parsed]).to_period("M").to_timestamp("M")[0]
            valid_periods.append(parsed_eom)
            period_col_indices.append(col_idx)
        elif len(valid_periods) > 0:
            # Si ya tenemos periodos y encontramos algo que no es periodo, parar
            break
            
    if profiler is not None:
        profiler.add_ms("deteccion_header", (time.perf_counter() - t1) * 1000.0)

    return df, valid_periods, period_col_indices


def extract_metric_series(
    df: pd.DataFrame,
    periods: List[pd.Timestamp],
    period_col_indices: List[int],
    metric_label_exact: str,
    search_column: int = 2,
    profiler: Optional[Profiler] = None,
) -> Optional[pd.Series]:
    """Devuelve una serie temporal (indexada por `periods`) para una métrica dada.

    Documentación de celdas:
    - Etiquetas de métricas: columna C (índice 2) o columna D (índice 3)
    - Valores por periodo: misma fila de la etiqueta, desde columna J (índice 9) en adelante
    - Encabezados de periodos: fila 10 (índice 9)
    
    Args:
        search_column: Columna donde buscar la etiqueta (2 para C, 3 para D)
    """
    # Normalizamos para evitar diferencias de guiones
    t_map = time.perf_counter()
    target = _normalize_label(metric_label_exact)
    labels_col = df.iloc[:, search_column].apply(_normalize_label)
    matches = labels_col[labels_col == target]
    if profiler is not None:
        profiler.add_ms("mapeo_etiquetas", (time.perf_counter() - t_map) * 1000.0)
    if matches.empty:
        return None

    # Tomar la última ocurrencia por robustez
    row_idx = matches.index[-1]
    t_ext = time.perf_counter()
    # Extraer exactamente las columnas de periodos detectadas
    row_values = [df.iat[row_idx, c] for c in period_col_indices]
    values = pd.to_numeric(row_values, errors="coerce").astype(float)
    if profiler is not None:
        profiler.add_ms("extraccion_valores", (time.perf_counter() - t_ext) * 1000.0)
    series = pd.Series(values, index=pd.DatetimeIndex(periods), name=metric_label_exact)
    return series


def extract_all_base_metrics(
    df: pd.DataFrame,
    periods: List[pd.Timestamp],
    period_col_indices: List[int],
    profiler: Optional[Profiler] = None,
) -> Dict[str, pd.Series]:
    """Extrae todas las métricas base (tal como están en el Excel) y devuelve
    un dict con claves internas y series como valores.
    """
    metrics: Dict[str, pd.Series] = {}

    # Buscar cada métrica usando etiquetas preferidas en columnas C y D
    for key, candidate_labels in PREFERRED_LABELS_BY_KEY.items():
        series_found: Optional[pd.Series] = None
        for label in candidate_labels:
            for search_col in (2, 3):
                series = extract_metric_series(
                    df,
                    periods,
                    period_col_indices,
                    label,
                    search_column=search_col,
                    profiler=profiler,
                )
                if series is not None:
                    series_found = series
                    break
            if series_found is not None:
                break
        if series_found is not None:
            metrics[key] = series_found
            
    return metrics


def compute_derived_metrics(base: Dict[str, pd.Series], profiler: Optional[Profiler] = None) -> Dict[str, pd.Series]:
    """Calcula métricas derivadas a partir de las series base.

    Fórmulas (exactamente como solicitado):
    - sales_efficiency_pct = sales / inventory_bop
    - pc1_ecac_unit = ((pc1_mn + ecac_mn) * 1e6) / sales
    - sga_unit = (sga_mn * 1e6) / sales
    - it_unit = (it_mn * 1e6) / sales
    - warehousing_unit = (warehousing_mn * 1e6) / sales

    Si el denominador es 0 o NaN → NaN (lo convertiremos a None al final).
    """
    derived: Dict[str, pd.Series] = {}

    sales = base.get("sales")
    inventory_bop = base.get("inventory_bop")
    pc1_mn = base.get("pc1_mn")
    ecac_mn = base.get("ecac_mn")
    sga_mn = base.get("sga_mn")
    it_mn = base.get("it_mn")
    warehousing_mn = base.get("warehousing_mn")

    t_der = time.perf_counter()
    if sales is not None and inventory_bop is not None:
        derived["sales_efficiency_pct"] = _safe_div(sales, inventory_bop) * 100.0  # Mostrar como porcentaje

    # eCAC unit: (Performance Marketing (mm) * 1e6) / sales
    if sales is not None and ecac_mn is not None:
        derived["ecac_unit"] = _safe_div(ecac_mn * 1_000_000.0, sales)

    # PC1 - eCAC unit: ((PC1 (mm) + Performance Marketing (mm)) * 1e6) / sales
    if sales is not None and pc1_mn is not None and ecac_mn is not None:
        combined_mn = pc1_mn + ecac_mn
        derived["pc1_ecac_unit"] = _safe_div(combined_mn * 1_000_000.0, sales)

    if sales is not None and sga_mn is not None:
        derived["sga_unit"] = _safe_div(sga_mn * 1_000_000.0, sales)

    if sales is not None and it_mn is not None:
        derived["it_unit"] = _safe_div(it_mn * 1_000_000.0, sales)

    if sales is not None and warehousing_mn is not None:
        derived["warehousing_unit"] = _safe_div(warehousing_mn * 1_000_000.0, sales)

    if profiler is not None:
        profiler.add_ms("calculo_derivados", (time.perf_counter() - t_der) * 1000.0)

    return derived


def build_wide_table(
    metrics: Dict[str, pd.Series],
    periods: List[pd.Timestamp],
    delta_label: str = "MoM_Δ",
) -> pd.DataFrame:
    """Construye una tabla ancha:
    - Una fila por métrica
    - Una columna por periodo (orden cronológico ascendente)
    - Columna adicional MoM_Δ = (mes_reciente - mes_anterior) / mes_anterior, formateada con '%'
    """
    period_cols = [p.strftime("%Y-%m-%d") for p in periods]
    rows: List[Dict[str, object]] = []

    for metric_key, series in metrics.items():
        series = series.astype(float)
        display_metric = ALIAS_KEYS.get(metric_key, metric_key)

        aligned = series.reindex(pd.DatetimeIndex(periods))
        row: Dict[str, object] = {"metric": display_metric}
        for col_name, value in zip(period_cols, aligned.values):
            row[col_name] = None if pd.isna(value) else float(value)

        if len(periods) >= 2:
            last_val = aligned.iloc[-1]
            prev_val = aligned.iloc[-2]
            if pd.notna(last_val) and pd.notna(prev_val) and prev_val != 0:
                mom_delta = (float(last_val) - float(prev_val)) / float(prev_val) * 100.0
                row[delta_label] = f"{mom_delta:.2f}%"
            else:
                row[delta_label] = ""
        else:
            row[delta_label] = ""

        rows.append(row)

    columns = ["metric"] + period_cols + [delta_label]
    wide_df = pd.DataFrame(rows, columns=columns)
    return wide_df


def _aggregator_type_for_key(metric_key: str) -> str:
    key = metric_key.lower()
    # Métricas unitarias o porcentajes → promedio
    if key.endswith("_unit") or key.endswith("_pct") or key in {"ecac_unit", "pc1_ecac_unit"}:
        return "mean"
    # Montos → suma
    return "sum"


def build_wide_table_from_groups(
    metrics: Dict[str, pd.Series],
    periods: List[pd.Timestamp],
    groups: List[List[int]],
    group_labels: List[str],
    delta_label: str,
) -> pd.DataFrame:
    """Agrega valores mensuales por grupos (trimestres/años) con reglas por métrica y construye tabla ancha."""
    period_index = pd.DatetimeIndex(periods)
    rows: List[Dict[str, object]] = []

    for metric_key, series in metrics.items():
        series = series.astype(float)
        aligned = series.reindex(period_index)
        agg = _aggregator_type_for_key(metric_key)
        values: List[Optional[float]] = []
        for idxs in groups:
            slice_vals = aligned.iloc[idxs]
            if slice_vals.dropna().empty:
                values.append(None)
                continue
            if agg == "mean":
                values.append(float(slice_vals.mean()))
            else:
                values.append(float(slice_vals.sum()))

        display_metric = ALIAS_KEYS.get(metric_key, metric_key)
        row: Dict[str, object] = {"metric": display_metric}
        for label, val in zip(group_labels, values):
            row[label] = None if val is None else float(val)

        if len(values) >= 2 and values[-2] not in (None, 0) and values[-1] is not None:
            prev = float(values[-2])
            curr = float(values[-1])
            delta = (curr - prev) / prev * 100.0 if prev != 0 else None
            row[delta_label] = f"{delta:.2f}%" if delta is not None else ""
        else:
            row[delta_label] = ""

        rows.append(row)

    columns = ["metric"] + group_labels + [delta_label]
    return pd.DataFrame(rows, columns=columns)


############################################################
# App Streamlit (tabla interactiva)
############################################################

def run_app():  # pragma: no cover
    if st is None:
        raise RuntimeError("Streamlit no está instalado. Ejecuta con: pip install streamlit")

    st.set_page_config(page_title="Kavak Metrics", layout="wide")
    st.title("Kavak Metrics")

    # Autenticación simple por contraseña (solo UI)
    PASSWORD = "K4vakmetrics2025"
    if "auth_ok" not in st.session_state:
        st.session_state["auth_ok"] = False
    if not st.session_state["auth_ok"]:
        # Ocultar el hint "Press Enter to apply" del input de contraseña
        st.markdown(
            """
            <style>
            [data-testid="stTextInput"] div[role="alert"] { display: none; }
            </style>
            """,
            unsafe_allow_html=True,
        )
        with st.form("auth_form"):
            pwd = st.text_input("Contraseña", type="password")
            ok = st.form_submit_button("Entrar")
            if ok:
                if pwd == PASSWORD:
                    st.session_state["auth_ok"] = True
                else:
                    st.error("Contraseña incorrecta")
        if not st.session_state["auth_ok"]:
            st.stop()

    # Entrada de ruta configurable
    st.sidebar.header("Configuración")
    excel_path_str = st.sidebar.text_input(
        "Ruta del archivo Excel",
        value=str(DEFAULT_EXCEL_PATH.resolve()),
        help="Ejemplo: C:/Users/usuario/Documentos/202506_Financials_by_Country.xlsx",
    )
    excel_path = Path(excel_path_str)

    if not excel_path.exists():
        st.error(f"No se encontró el archivo: {excel_path}")
        st.stop()

    # Detectar hojas disponibles y seleccionar país
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception as e:
        st.exception(e)
        st.stop()

    sheet_names = xls.sheet_names
    # El usuario pidió que se pueda seleccionar también vía variable country_sheet
    # Aquí preseleccionamos con DEFAULT_COUNTRY_SHEET si existe; si no, la primera hoja.
    preselect = DEFAULT_COUNTRY_SHEET if DEFAULT_COUNTRY_SHEET in sheet_names else (sheet_names[0] if sheet_names else "")
    country_sheet = st.sidebar.selectbox("Hoja (país)", options=sheet_names, index=(sheet_names.index(preselect) if preselect in sheet_names else 0))

    if not country_sheet:
        st.warning("No hay hojas en el archivo.")
        st.stop()

    # Perfilado activable por env var PROFILE
    profile_enabled = os.getenv("PROFILE", "false").lower() in {"1", "true", "yes", "on"}
    profiler = Profiler(profile_enabled)
    profiler.start_cprofile()

    # Leer hoja seleccionada
    try:
        df, periods, period_col_indices = read_sheet_as_dataframe(excel_path, country_sheet, profiler=profiler)
    except Exception as e:
        st.exception(e)
        st.stop()

    st.caption(
        "Notas de extracción: etiquetas en columnas C y D (índices 2 y 3), periodos en fila 10 (índice 9), valores desde columna J (índice 9)."
    )

    # Extraer métricas base
    base_metrics = extract_all_base_metrics(df, periods, period_col_indices, profiler=profiler)

    # Métricas derivadas
    derived_metrics = compute_derived_metrics(base_metrics, profiler=profiler)

    # Consolidar todas las métricas (base + derivadas), ocultando algunas bases
    all_metrics: Dict[str, pd.Series] = {}
    # Orden: base según PREFERRED_LABELS_BY_KEY, luego derivadas
    for key in PREFERRED_LABELS_BY_KEY.keys():
        if key in base_metrics and key not in HIDE_BASE_FROM_DISPLAY:
            all_metrics[key] = base_metrics[key]
    for key in DERIVED_FORMULAS_ORDER:
        if key in derived_metrics:
            all_metrics[key] = derived_metrics[key]

    # Mostrar selección de métricas a visualizar/descargar
    all_metric_keys = list(all_metrics.keys())
    selected_metrics = st.multiselect(
        "Selecciona métricas a mostrar",
        options=all_metric_keys,
        default=all_metric_keys,
    )

    # Selector de formato de periodo
    period_format = st.sidebar.radio("Period Format", ["Monthly", "Quarterly", "Yearly"], index=0)

    # Construir opciones de periodos según formato
    if period_format == "Monthly":
        period_labels = [p.strftime("%b-%y") for p in periods]
        selected_labels = st.sidebar.multiselect("Periods", options=period_labels, default=period_labels)
        # Mapear a índices
        selected_idx = [i for i, lbl in enumerate(period_labels) if lbl in selected_labels]
        sel_periods = [periods[i] for i in selected_idx]
        filtered_metrics = {k: v for k, v in all_metrics.items() if k in selected_metrics}
        wide_table = build_wide_table(filtered_metrics, sel_periods, delta_label="MoM_Δ")

    elif period_format == "Quarterly":
        # Construir grupos trimestrales
        q_labels: List[str] = []
        q_groups: List[List[int]] = []
        for i, p in enumerate(periods):
            q = (p.month - 1) // 3 + 1
            label = f"Q{q}-{p.strftime('%y')}"
            if not q_labels or q_labels[-1] != label:
                q_labels.append(label)
                q_groups.append([i])
            else:
                q_groups[-1].append(i)
        selected_q = st.sidebar.multiselect("Quarters", options=q_labels, default=q_labels)
        # Filtrar grupos por selección
        sel_groups = [grp for lbl, grp in zip(q_labels, q_groups) if lbl in selected_q]
        filtered_metrics = {k: v for k, v in all_metrics.items() if k in selected_metrics}
        wide_table = build_wide_table_from_groups(filtered_metrics, periods, sel_groups, selected_q, delta_label="QoQ_Δ")

    else:  # Yearly
        y_labels: List[str] = []
        y_groups: List[List[int]] = []
        for i, p in enumerate(periods):
            label = p.strftime('%Y')
            if not y_labels or y_labels[-1] != label:
                y_labels.append(label)
                y_groups.append([i])
            else:
                y_groups[-1].append(i)
        selected_y = st.sidebar.multiselect("Years", options=y_labels, default=y_labels)
        sel_groups = [grp for lbl, grp in zip(y_labels, y_groups) if lbl in selected_y]
        filtered_metrics = {k: v for k, v in all_metrics.items() if k in selected_metrics}
        wide_table = build_wide_table_from_groups(filtered_metrics, periods, sel_groups, selected_y, delta_label="YoY_Δ")

    # Render interactivo
    t_render = time.perf_counter()
    st.subheader("Tabla por periodos (ancha) + MoM_Δ")
    st.dataframe(
        wide_table,
        use_container_width=True,
        hide_index=True,
    )

    # Descarga CSV
    csv_bytes = wide_table.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Descargar CSV",
        data=csv_bytes,
        file_name=f"financials_wide_{country_sheet.replace(' ', '_')}.csv",
        mime="text/csv",
    )
    profiler.add_ms("render", (time.perf_counter() - t_render) * 1000.0)

    # Mostrar perfilado (si está activo)
    profiler.stop_cprofile()
    if profile_enabled:
        with st.expander("Perfilado (PROFILE=true)"):
            st.text(profiler.format_timings())
            st.text("\nTop cProfile (20):\n" + profiler.get_cprofile_report(20))


if __name__ == "__main__":  # pragma: no cover
    # Permite ejecutar como script para validar extracción sin UI
    # Si deseas la UI interactiva: `streamlit run financials_dashboard.py`
    import argparse

    parser = argparse.ArgumentParser(description="Genera tabla larga de métricas financieras por país")
    parser.add_argument("--excel_path", type=str, default=str(DEFAULT_EXCEL_PATH), help="Ruta del archivo Excel")
    parser.add_argument("--country_sheet", type=str, default=DEFAULT_COUNTRY_SHEET, help="Nombre exacto de la hoja (país)")
    parser.add_argument("--no_ui", action="store_true", help="Ejecutar en modo no interactivo y volcar CSV por stdout")
    args = parser.parse_args()

    excel_path = Path(args.excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {excel_path}")

    if args.no_ui:
        # Perfilado por env var
        profile_enabled = os.getenv("PROFILE", "false").lower() in {"1", "true", "yes", "on"}
        profiler = Profiler(profile_enabled)
        profiler.start_cprofile()

        df, periods, period_col_indices = read_sheet_as_dataframe(excel_path, args.country_sheet, profiler=profiler)
        base_metrics = extract_all_base_metrics(df, periods, period_col_indices, profiler=profiler)
        derived_metrics = compute_derived_metrics(base_metrics, profiler=profiler)
        # Unir todo y construir tabla ancha
        all_metrics = {**base_metrics, **derived_metrics}
        wide_table = build_wide_table(all_metrics, periods)
        # Salida CSV a stdout
        print(wide_table.to_csv(index=False))

        # Imprimir perfilado al final por stderr para no contaminar CSV
        profiler.stop_cprofile()
        if profile_enabled:
            sys.stderr.write(profiler.format_timings() + "\n")
            sys.stderr.write("\nTop cProfile (20):\n")
            sys.stderr.write(profiler.get_cprofile_report(20) + "\n")
    else:
        if st is None:
            raise RuntimeError("Para la interfaz interactiva, instala Streamlit: pip install streamlit")
        run_app()


