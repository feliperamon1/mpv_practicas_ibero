from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# =============================================================================
# Configuración general
# =============================================================================

APP_TITLE = "Dashboard Ejecutivo SG-SST | IMA Company SAS"
DATA_DIR = Path(__file__).parent / "data"

FILES = {
    "presupuesto": "Formato_Asignacion_Recursos_Presupuesto_SST_2023_2026.xlsx",
    "formacion": "Plan_Formacion_y_Capacitaciones_2023_2026.xlsx",
    "autodiagnostico": "Autodiagnostico_Res_0312_2019_2023_2025.xlsx",
    "plan_anual": "Plan_Anual_Trabajo_PHVA_2023_2026.xlsx",
    "legal": "Matriz_Legal_SGRL_SST_Colombia.xlsx",
    "reporte_laboral": "Base Reporte Laboral Autogestión - IMA Company SAS (1).xlsx",
    "proveedores": "Gestion_Proveedores_Contratistas_Vigilancia_2023_2026.xlsx",
    "emo": "Matriz_Seguimiento_EMO_2023_2026.xlsx",
    "accidentalidad": "Matriz_Accidentalidad_2023_2026.xlsx",
    "enfermedad": "Matriz_Enfermedad_Laboral_2023_2026.xlsx",
    "indicadores": "Matriz_Indicadores_Gestion_SGSST_2023_2026.xlsx",
    "peligros": "Matriz_Peligros_Riesgos_GTC45_Vigilancia.xlsx",
    "mantenimiento": "Matriz_Mantenimiento_Preventivo_Vigilancia_2023_2026.xlsx",
    "apcm": "Matriz_APCM_Vigilancia_2023_2026.xlsx",
}

MODULE_LABELS = {
    "resumen": "Resumen ejecutivo",
    "calidad": "Carga y calidad de datos",
    "presupuesto": "1. Presupuesto SST",
    "formacion": "2. Formación y capacitaciones",
    "autodiagnostico": "3. Autodiagnóstico Res. 0312",
    "plan_anual": "4. Plan anual PHVA",
    "legal": "5. Matriz legal SGRL",
    "reporte_laboral": "6. Reporte laboral autogestión",
    "proveedores": "7. Proveedores y contratistas",
    "emo": "8. Exámenes médicos ocupacionales",
    "accidentalidad": "9. Accidentalidad",
    "enfermedad": "10. Enfermedad laboral",
    "indicadores": "11. Indicadores SG-SST",
    "peligros": "12. Peligros y riesgos GTC-45",
    "mantenimiento": "13. Mantenimiento preventivo",
    "apcm": "14. Acciones preventivas, correctivas y mejora",
}

SHEETS = {
    "presupuesto": {"detalle": "Detalle_Presupuesto", "anual": "Resumen_Anual"},
    "formacion": {"registro": "Registro", "anual": "Resumen_Anual", "sedes": "Resumen_Sedes", "mensual": "Resumen_Mensual"},
    "autodiagnostico": {"resumen": "Resumen_Comparativo", "detalle": ["Auto_2023", "Auto_2024", "Auto_2025"]},
    "plan_anual": {"anual": "Resumen_Anual", "mensual": "Resumen_Mensual", "detalle": ["Plan_2023", "Plan_2024", "Plan_2025", "Plan_2026"]},
    "legal": {"matriz": "Matriz_Legal_SGRL", "resumen": "Resumen"},
    "reporte_laboral": {"respuestas": "Respuestas del formulario", "gestion": "Gestión SST"},
    "proveedores": {"detalle": ["GPC_2023", "GPC_2024", "GPC_2025", "GPC_2026"], "anual": "Resumen_Anual"},
    "emo": {"detalle": "Consolidado", "anual": "Resumen_Anual", "sedes": "Resumen_Sedes"},
    "accidentalidad": {"detalle": "Consolidado", "anual": "Resumen_Anual", "sedes": "Resumen_Sedes"},
    "enfermedad": {"detalle": "Consolidado", "anual": "Resumen_Anual", "sedes": "Resumen_Sedes"},
    "indicadores": {"ficha": "Ficha_Tecnica", "medicion": "Consolidado", "anual": "Resumen_Anual"},
    "peligros": {"matriz": "Matriz_General", "resumen": "Resumen"},
    "mantenimiento": {"detalle": "Consolidado", "anual": "Resumen_Anual", "sedes": "Resumen_Sedes"},
    "apcm": {"detalle": "Consolidado", "anual": "Resumen_Anual", "sedes": "Resumen_Sedes"},
}

STATUS_COLORS = {
    "Ejecutado": "#16A34A",
    "Ejecutada": "#16A34A",
    "Cerrada": "#16A34A",
    "Cerrado": "#16A34A",
    "Vigente": "#16A34A",
    "En meta": "#16A34A",
    "Sí": "#16A34A",
    "Si": "#16A34A",
    "Apto": "#16A34A",
    "Aceptable": "#16A34A",
    "Pendiente": "#F59E0B",
    "En gestión": "#F59E0B",
    "En progreso": "#F59E0B",
    "Reprogramado": "#0EA5E9",
    "Reprogramada": "#0EA5E9",
    "Parcial": "#F59E0B",
    "Próximo a vencer": "#F59E0B",
    "Alerta": "#F59E0B",
    "No": "#DC2626",
    "No cumple": "#DC2626",
    "Cancelado": "#64748B",
    "Cancelada": "#64748B",
    "Vencido": "#DC2626",
    "Vencida": "#DC2626",
    "Crítico": "#DC2626",
    "Critico": "#DC2626",
}

PALETTE = ["#0F172A", "#1D4ED8", "#0EA5E9", "#10B981", "#F59E0B", "#DC2626", "#7C3AED", "#64748B"]
MONTH_ORDER = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    html, body, [class*="css"] {font-family: 'Inter', sans-serif;}
    .main .block-container {padding-top: 1.0rem; padding-bottom: 2.2rem; max-width: 1480px;}
    .hero {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 58%, #0EA5E9 100%);
        border-radius: 24px; padding: 28px 32px; color: #FFFFFF; margin-bottom: 18px;
        box-shadow: 0 14px 30px rgba(15, 23, 42, .20);
    }
    .hero h1 {font-size: 2.0rem; line-height: 1.1; margin: 0; font-weight: 850; letter-spacing: -0.03em;}
    .hero p {font-size: .98rem; margin: .55rem 0 0 0; color: #D8E7FF; max-width: 1100px;}
    .pill {display: inline-block; padding: 6px 12px; border-radius: 999px; background: rgba(255,255,255,.12); color: #E0F2FE; border: 1px solid rgba(255,255,255,.20); font-size: .78rem; margin-right: 8px; margin-bottom: 10px;}
    .section-title {font-size: 1.08rem; font-weight: 800; color: #0F172A; margin-top: 1rem; margin-bottom: .55rem;}
    .kpi-card {
        background: #FFFFFF; border: 1px solid #E2E8F0; border-left: 6px solid #1D4ED8;
        border-radius: 18px; padding: 16px 16px 14px 16px; min-height: 118px;
        box-shadow: 0 8px 18px rgba(15, 23, 42, .06);
    }
    .kpi-card.good {border-left-color: #16A34A;}
    .kpi-card.warn {border-left-color: #F59E0B;}
    .kpi-card.bad {border-left-color: #DC2626;}
    .kpi-card.info {border-left-color: #0EA5E9;}
    .kpi-label {font-size: .78rem; color: #64748B; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; margin-bottom: 8px;}
    .kpi-value {font-size: 1.55rem; font-weight: 850; color: #0F172A; letter-spacing: -0.03em;}
    .kpi-help {font-size: .78rem; color: #64748B; margin-top: 6px;}
    .insight-card {background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 16px; padding: 14px 16px; color:#0F172A;}
    .stTabs [data-baseweb="tab-list"] {gap: 8px;}
    .stTabs [data-baseweb="tab"] {height: 42px; border-radius: 999px; background: #F1F5F9; padding: 10px 16px;}
    .stTabs [aria-selected="true"] {background: #DBEAFE; color: #1D4ED8; font-weight: 800;}
    div[data-testid="stSidebar"] {background: #F8FAFC;}
    div[data-testid="stMetricValue"] {font-weight: 850;}
    </style>
    """,
    unsafe_allow_html=True,
)

# =============================================================================
# Funciones de lectura robusta y limpieza de datos
# =============================================================================

HEADER_KEYWORDS = {
    "ano", "año", "fecha", "mes", "sede", "estado", "indicador", "codigo", "código", "tipo",
    "proceso", "actividad", "responsable", "cumple", "cumplimiento", "valor", "costo",
    "resultado", "riesgo", "nivel", "accion", "acción", "caso", "activo", "presupuesto",
    "proveedor", "nombre", "cargo", "frecuencia", "clasificacion", "clasificación", "fuente",
    "programadas", "ejecutadas", "puntaje", "estandar", "estándar", "item", "ítem",
}


def strip_accents(value: Any) -> str:
    value = unicodedata.normalize("NFKD", str(value))
    return "".join(ch for ch in value if not unicodedata.combining(ch))


def norm(value: Any) -> str:
    value = strip_accents(str(value or "").strip().lower())
    value = re.sub(r"[^a-z0-9]+", "_", value)
    return re.sub(r"_+", "_", value).strip("_")


def clean_col(value: Any) -> str:
    cleaned = norm(value)
    return cleaned if cleaned else "columna"


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return str(value).strip() == ""


def detect_header_row(raw: pd.DataFrame, max_rows: int = 25) -> int:
    """Encuentra la fila de encabezados en hojas con títulos, celdas combinadas y filas vacías."""
    best_idx, best_score = 0, -999.0
    rows = min(max_rows, len(raw))
    for idx in range(rows):
        values = [v for v in raw.iloc[idx].tolist() if not is_blank(v)]
        if len(values) < 2:
            continue
        normalized = [norm(v) for v in values]
        hits = 0
        for v in normalized:
            if v in HEADER_KEYWORDS:
                hits += 2
            elif any(k in v for k in HEADER_KEYWORDS if len(k) > 3):
                hits += 1
        numeric_years = sum(1 for v in values if str(v).strip() in {"2023", "2024", "2025", "2026"})
        uniqueness = len(set(normalized)) / max(len(normalized), 1)
        score = hits * 4 + len(values) * 0.25 + uniqueness + numeric_years
        # Penalizar filas descriptivas largas con pocas columnas.
        if len(values) <= 3 and max(len(str(v)) for v in values) > 80:
            score -= 8
        if score > best_score:
            best_idx, best_score = idx, score
    return best_idx


def clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    seen: dict[str, int] = {}
    cols: list[str] = []
    for c in df.columns:
        base = clean_col(c)
        if base.startswith("unnamed"):
            base = "columna"
        seen[base] = seen.get(base, 0) + 1
        cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
    df.columns = cols
    # Eliminar columnas completamente vacías o separadores sin contenido.
    empty_cols = [c for c in df.columns if df[c].isna().all()]
    df = df.drop(columns=empty_cols, errors="ignore")
    # Eliminar filas vacías y filas que repiten encabezados.
    df = df.dropna(how="all")
    if not df.empty:
        repeated_header = pd.Series(False, index=df.index)
        for c in df.columns[: min(8, len(df.columns))]:
            repeated_header = repeated_header | (df[c].astype(str).str.lower().str.strip() == c.replace("_", " "))
        if repeated_header.any():
            df = df.loc[~repeated_header]
    return df.reset_index(drop=True)


def fix_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if any(token in col for token in ["fecha", "timestamp", "corte"]):
            parsed = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
            numeric = pd.to_numeric(df[col], errors="coerce")
            excel_dt = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
            df[col] = parsed.fillna(excel_dt)
    return df


def file_path(key: str) -> Path:
    expected = DATA_DIR / FILES[key]
    if expected.exists():
        return expected
    target = norm(FILES[key])
    for path in DATA_DIR.glob("*.xlsx"):
        if norm(path.name) == target or norm(path.stem) in target or target in norm(path.stem):
            return path
    return expected


def module_available(key: str) -> bool:
    return file_path(key).exists()


@st.cache_data(show_spinner=False)
def read_sheet_cached(path_str: str, sheet_name: str) -> pd.DataFrame:
    path = Path(path_str)
    try:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None, engine="openpyxl")
    except Exception:
        return pd.DataFrame()
    header_row = detect_header_row(raw)
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=header_row, engine="openpyxl")
    except Exception:
        return pd.DataFrame()
    df = clean_headers(df)
    df = fix_date_columns(df)
    return df


def read_sheet(key: str, sheet_name: str) -> pd.DataFrame:
    path = file_path(key)
    if not path.exists():
        return pd.DataFrame()
    return read_sheet_cached(str(path), sheet_name)


def read_many(key: str, sheet_names: Iterable[str]) -> pd.DataFrame:
    frames = []
    for sheet in sheet_names:
        df = read_sheet(key, sheet)
        if not df.empty:
            df["hoja_origen"] = sheet
            frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def load_module(key: str) -> dict[str, pd.DataFrame]:
    result: dict[str, pd.DataFrame] = {}
    for alias, spec in SHEETS[key].items():
        if isinstance(spec, list):
            result[alias] = read_many(key, spec)
        else:
            result[alias] = read_sheet(key, spec)
    return result


@st.cache_data(show_spinner=False)
def load_all_summary() -> pd.DataFrame:
    rows = []
    for key, filename in FILES.items():
        path = file_path(key)
        status = "Encontrado" if path.exists() else "No encontrado"
        if not path.exists():
            rows.append({"Módulo": MODULE_LABELS.get(key, key), "Archivo": filename, "Hoja": "—", "Filas": 0, "Columnas": 0, "Estado": status})
            continue
        for alias, spec in SHEETS[key].items():
            if isinstance(spec, list):
                for sh in spec:
                    df = read_sheet(key, sh)
                    rows.append({"Módulo": MODULE_LABELS.get(key, key), "Archivo": filename, "Hoja": sh, "Filas": len(df), "Columnas": len(df.columns), "Estado": status})
            else:
                df = read_sheet(key, spec)
                rows.append({"Módulo": MODULE_LABELS.get(key, key), "Archivo": filename, "Hoja": spec, "Filas": len(df), "Columnas": len(df.columns), "Estado": status})
    return pd.DataFrame(rows)

# =============================================================================
# Utilidades analíticas y visuales
# =============================================================================


def col(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    if df.empty:
        return None
    normalized = [clean_col(c) for c in candidates]
    for c in normalized:
        if c in df.columns:
            return c
    for candidate in normalized:
        for existing in df.columns:
            if candidate and (candidate in existing or existing in candidate):
                return existing
    return None


def num(s: Any) -> pd.Series | float:
    if isinstance(s, pd.Series):
        return pd.to_numeric(s, errors="coerce")
    try:
        return float(s)
    except Exception:
        return np.nan


def safe_sum(df: pd.DataFrame, candidates: Iterable[str]) -> float:
    c = col(df, candidates)
    if c is None:
        return 0.0
    return float(num(df[c]).sum())


def safe_mean(df: pd.DataFrame, candidates: Iterable[str]) -> float:
    c = col(df, candidates)
    if c is None:
        return np.nan
    return float(num(df[c]).mean())


def count_text(df: pd.DataFrame, candidates: Iterable[str], pattern: str) -> int:
    c = col(df, candidates)
    if c is None or df.empty:
        return 0
    return int(df[c].astype(str).str.lower().str.contains(pattern.lower(), na=False).sum())


def format_num(value: Any, percent: bool = False, money: bool = False, decimals: int = 1) -> str:
    try:
        v = float(value)
        if np.isnan(v):
            return "—"
        if percent:
            if abs(v) <= 1.5:
                v *= 100
            return f"{v:,.{decimals}f}%".replace(",", "X").replace(".", ",").replace("X", ".")
        if money:
            return "$" + f"{v:,.0f}".replace(",", ".")
        if v.is_integer():
            return f"{v:,.0f}".replace(",", ".")
        return f"{v:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        if value is None or pd.isna(value):
            return "—"
        return str(value)


def status_class(value: Any, percent: bool = False, inverse: bool = False) -> str:
    try:
        v = float(value)
        if percent and abs(v) <= 1.5:
            v *= 100
        if inverse:
            return "good" if v == 0 else "warn" if v <= 3 else "bad"
        return "good" if v >= 85 else "warn" if v >= 60 else "bad"
    except Exception:
        return "info"


def kpi(label: str, value: Any, help_text: str = "", percent: bool = False, money: bool = False, state: str = "info", inverse: bool = False) -> None:
    if state == "auto":
        state = status_class(value, percent=percent, inverse=inverse)
    st.markdown(
        f"""
        <div class='kpi-card {state}'>
            <div class='kpi-label'>{label}</div>
            <div class='kpi-value'>{format_num(value, percent=percent, money=money)}</div>
            <div class='kpi-help'>{help_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def hero(title: str, subtitle: str, chips: list[str] | None = None) -> None:
    chip_html = "".join([f"<span class='pill'>{x}</span>" for x in (chips or [])])
    st.markdown(f"<div class='hero'>{chip_html}<h1>{title}</h1><p>{subtitle}</p></div>", unsafe_allow_html=True)


def section(title: str) -> None:
    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)


def fig_style(fig: go.Figure, height: int = 410) -> go.Figure:
    fig.update_layout(
        template="plotly_white",
        height=height,
        font=dict(family="Inter, Arial", size=12, color="#0F172A"),
        title=dict(font=dict(size=18, color="#0F172A"), x=0.02, xanchor="left"),
        margin=dict(l=16, r=16, t=58, b=24),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
    )
    return fig


def show_fig(fig: go.Figure, height: int = 410) -> None:
    st.plotly_chart(fig_style(fig, height), use_container_width=True, config={"displayModeBar": True, "responsive": True})


def chart_empty(title: str) -> None:
    st.info(f"No hay datos suficientes para generar: {title}")


def count_chart(df: pd.DataFrame, colname: str | None, title: str, orientation: str = "v", top: int = 15, height: int = 410) -> None:
    if df.empty or colname is None or colname not in df.columns:
        chart_empty(title)
        return
    d = df[colname].fillna("Sin dato").astype(str).str.strip().replace({"": "Sin dato"}).value_counts().head(top).reset_index()
    d.columns = [colname, "cantidad"]
    if orientation == "h":
        fig = px.bar(d.sort_values("cantidad"), x="cantidad", y=colname, orientation="h", text="cantidad", color=colname, color_discrete_sequence=PALETTE)
    else:
        fig = px.bar(d, x=colname, y="cantidad", text="cantidad", color=colname, color_discrete_sequence=PALETTE)
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(title=title, showlegend=False)
    show_fig(fig, height)


def donut_chart(df: pd.DataFrame, colname: str | None, title: str, height: int = 410) -> None:
    if df.empty or colname is None or colname not in df.columns:
        chart_empty(title)
        return
    d = df[colname].fillna("Sin dato").astype(str).str.strip().replace({"": "Sin dato"}).value_counts().reset_index()
    d.columns = [colname, "cantidad"]
    fig = px.pie(d, names=colname, values="cantidad", hole=.55, color=colname, color_discrete_map=STATUS_COLORS, color_discrete_sequence=PALETTE)
    fig.update_traces(textinfo="percent+label", hovertemplate="%{label}<br>Cantidad: %{value}<br>%{percent}<extra></extra>")
    fig.update_layout(title=title)
    show_fig(fig, height)


def sum_chart(df: pd.DataFrame, group_col: str | None, value_col: str | None, title: str, top: int = 15, orientation: str = "v", money: bool = False) -> None:
    if df.empty or group_col is None or value_col is None or group_col not in df.columns or value_col not in df.columns:
        chart_empty(title)
        return
    d = df.groupby(group_col, dropna=False)[value_col].apply(lambda x: pd.to_numeric(x, errors="coerce").sum()).reset_index().sort_values(value_col, ascending=False).head(top)
    if orientation == "h":
        fig = px.bar(d.sort_values(value_col), x=value_col, y=group_col, orientation="h", color=group_col, color_discrete_sequence=PALETTE, text=value_col)
    else:
        fig = px.bar(d, x=group_col, y=value_col, color=group_col, color_discrete_sequence=PALETTE, text=value_col)
    fig.update_traces(texttemplate="$%{text:,.0f}" if money else "%{text:,.0f}", textposition="outside", cliponaxis=False)
    fig.update_layout(title=title, showlegend=False)
    show_fig(fig)


def mean_chart(df: pd.DataFrame, group_col: str | None, value_col: str | None, title: str, top: int = 15) -> None:
    if df.empty or group_col is None or value_col is None or group_col not in df.columns or value_col not in df.columns:
        chart_empty(title)
        return
    d = df.groupby(group_col, dropna=False)[value_col].apply(lambda x: pd.to_numeric(x, errors="coerce").mean()).reset_index().sort_values(value_col, ascending=False).head(top)
    fig = px.bar(d, x=group_col, y=value_col, color=group_col, color_discrete_sequence=PALETTE, text=value_col)
    fig.update_traces(texttemplate="%{text:.1%}", textposition="outside", cliponaxis=False)
    fig.update_layout(title=title, showlegend=False)
    show_fig(fig)


def line_chart(df: pd.DataFrame, x_col: str | None, y_col: str | None, title: str, color_col: str | None = None, agg: str = "sum") -> None:
    if df.empty or x_col is None or y_col is None or x_col not in df.columns or y_col not in df.columns:
        chart_empty(title)
        return
    d = df.copy()
    d[y_col] = pd.to_numeric(d[y_col], errors="coerce")
    if agg == "mean":
        d = d.groupby([x_col] + ([color_col] if color_col and color_col in d.columns else []), dropna=False)[y_col].mean().reset_index()
    else:
        d = d.groupby([x_col] + ([color_col] if color_col and color_col in d.columns else []), dropna=False)[y_col].sum().reset_index()
    if x_col == "mes" and set(d[x_col].dropna().astype(str)).intersection(MONTH_ORDER):
        d[x_col] = pd.Categorical(d[x_col], categories=MONTH_ORDER, ordered=True)
        d = d.sort_values(x_col)
    fig = px.line(d, x=x_col, y=y_col, color=color_col if color_col in d.columns else None, markers=True, color_discrete_sequence=PALETTE)
    fig.update_traces(line=dict(width=3), marker=dict(size=8))
    fig.update_layout(title=title)
    show_fig(fig)


def stacked_status_by_group(df: pd.DataFrame, group_col: str | None, status_col: str | None, title: str) -> None:
    if df.empty or group_col is None or status_col is None or group_col not in df.columns or status_col not in df.columns:
        chart_empty(title)
        return
    d = df.groupby([group_col, status_col], dropna=False).size().reset_index(name="cantidad")
    fig = px.bar(d, x=group_col, y="cantidad", color=status_col, text="cantidad", color_discrete_map=STATUS_COLORS, color_discrete_sequence=PALETTE)
    fig.update_traces(textposition="inside")
    fig.update_layout(title=title, barmode="stack")
    show_fig(fig)


def heatmap_month_status(df: pd.DataFrame, month_col: str | None, status_col: str | None, title: str) -> None:
    if df.empty or month_col is None or status_col is None or month_col not in df.columns or status_col not in df.columns:
        chart_empty(title)
        return
    d = df.groupby([month_col, status_col], dropna=False).size().reset_index(name="cantidad")
    pivot = d.pivot(index=status_col, columns=month_col, values="cantidad").fillna(0)
    ordered_cols = [m for m in MONTH_ORDER if m in pivot.columns]
    if ordered_cols:
        pivot = pivot[ordered_cols]
    fig = px.imshow(pivot, text_auto=True, aspect="auto", color_continuous_scale="Blues")
    fig.update_layout(title=title, xaxis_title="Mes", yaxis_title="Estado")
    show_fig(fig)


def table_with_download(df: pd.DataFrame, title: str, key: str, height: int = 420) -> None:
    section(title)
    if df.empty:
        st.info("No hay registros disponibles.")
        return
    c1, c2 = st.columns([4, 1])
    with c1:
        search = st.text_input("Buscar en la tabla", key=f"search_{key}", placeholder="Digite una palabra clave...")
    d = df.copy()
    if search:
        mask = d.astype(str).apply(lambda row: row.str.contains(search, case=False, na=False).any(), axis=1)
        d = d.loc[mask]
    with c2:
        st.download_button("Descargar CSV", d.to_csv(index=False).encode("utf-8-sig"), file_name=f"{key}.csv", mime="text/csv")
    st.dataframe(d, use_container_width=True, height=height, hide_index=True)


def unique_values(df: pd.DataFrame, column: str | None) -> list[Any]:
    if df.empty or column is None or column not in df.columns:
        return []
    vals = df[column].dropna().astype(str).str.strip()
    vals = [v for v in vals.unique().tolist() if v and v.lower() != "nan"]
    def sort_key(v: Any):
        try:
            return (0, int(float(v)))
        except Exception:
            return (1, str(v))
    return sorted(vals, key=sort_key)


def filter_data(df: pd.DataFrame, key: str, fields: list[tuple[str, list[str]]]) -> pd.DataFrame:
    if df.empty:
        return df
    d = df.copy()
    with st.sidebar.expander("Filtros interactivos", expanded=True):
        for label, candidates in fields:
            c = col(d, candidates)
            vals = unique_values(d, c)
            if not vals or len(vals) > 250:
                continue
            default = vals
            if label.lower() in {"año", "ano"} and vals:
                # Por defecto se muestra el año más reciente para que el dashboard sea ejecutivo.
                default = [vals[-1]]
            selected = st.multiselect(label, vals, default=default, key=f"{key}_{clean_col(label)}")
            if selected and c:
                d = d[d[c].astype(str).isin([str(x) for x in selected])]
    return d


def latest_year(df: pd.DataFrame) -> int | None:
    c = col(df, ["año", "ano"])
    if c is None:
        return None
    vals = pd.to_numeric(df[c], errors="coerce").dropna()
    if vals.empty:
        return None
    return int(vals.max())

# =============================================================================
# Datos de cada módulo
# =============================================================================


def get_presupuesto() -> dict[str, pd.DataFrame]: return load_module("presupuesto")
def get_formacion() -> dict[str, pd.DataFrame]: return load_module("formacion")
def get_autodiagnostico() -> dict[str, pd.DataFrame]: return load_module("autodiagnostico")
def get_plan_anual() -> dict[str, pd.DataFrame]: return load_module("plan_anual")
def get_legal() -> dict[str, pd.DataFrame]: return load_module("legal")
def get_reporte_laboral() -> dict[str, pd.DataFrame]: return load_module("reporte_laboral")
def get_proveedores() -> dict[str, pd.DataFrame]: return load_module("proveedores")
def get_emo() -> dict[str, pd.DataFrame]: return load_module("emo")
def get_accidentalidad() -> dict[str, pd.DataFrame]: return load_module("accidentalidad")
def get_enfermedad() -> dict[str, pd.DataFrame]: return load_module("enfermedad")
def get_indicadores() -> dict[str, pd.DataFrame]: return load_module("indicadores")
def get_peligros() -> dict[str, pd.DataFrame]: return load_module("peligros")
def get_mantenimiento() -> dict[str, pd.DataFrame]: return load_module("mantenimiento")
def get_apcm() -> dict[str, pd.DataFrame]: return load_module("apcm")

# =============================================================================
# Dashboards
# =============================================================================


def render_resumen() -> None:
    hero(
        "Dashboard Integral SG-SST",
        "Vista ejecutiva del ecosistema digital de 14 instrumentos para seguimiento normativo, operativo y gerencial del SG-SST.",
        ["Google Sheets / Excel", "Streamlit", "Plotly interactivo", "SG-SST Colombia"],
    )

    p = get_presupuesto()["detalle"]
    f = get_formacion()["registro"]
    a = get_autodiagnostico()["detalle"]
    pa = get_plan_anual()["detalle"]
    l = get_legal()["matriz"]
    r = get_reporte_laboral()["gestion"]
    pr = get_proveedores()["detalle"]
    emo = get_emo()["detalle"]
    at = get_accidentalidad()["detalle"]
    el = get_enfermedad()["detalle"]
    ind = get_indicadores()["medicion"]
    pel = get_peligros()["matriz"]
    mt = get_mantenimiento()["detalle"]
    ap = get_apcm()["detalle"]

    # Filtro global de año para los módulos que tienen año.
    years = sorted(set().union(*[
        set(pd.to_numeric(df[col(df, ["año", "ano"])] if col(df, ["año", "ano"]) else pd.Series(dtype=float), errors="coerce").dropna().astype(int).tolist())
        for df in [p, f, a, pa, pr, emo, at, el, ind, mt, ap]
    ]))
    default_year = max(years) if years else 2026
    selected_year = st.sidebar.selectbox("Año ejecutivo", years if years else [2026], index=(years.index(default_year) if years else 0), key="resumen_year")

    def fy(df: pd.DataFrame) -> pd.DataFrame:
        c = col(df, ["año", "ano"])
        return df[pd.to_numeric(df[c], errors="coerce") == selected_year] if c else df

    p_y, f_y, a_y, pa_y, pr_y, emo_y, at_y, el_y, ind_y, mt_y, ap_y = [fy(x) for x in [p, f, a, pa, pr, emo, at, el, ind, mt, ap]]

    presupuesto_planeado = safe_sum(p_y, ["valor_presupuestado"])
    presupuesto_ejecutado = safe_sum(p_y, ["valor_ejecutado"])
    ejec_pres = presupuesto_ejecutado / presupuesto_planeado if presupuesto_planeado else np.nan
    form_cov = safe_mean(f_y, ["cobertura", "cobertura_promedio"])
    auto_puntaje = safe_sum(a_y, ["puntaje_obtenido"])
    auto_ponderacion = safe_sum(a_y, ["ponderacion"])
    auto_cump = auto_puntaje / auto_ponderacion if auto_ponderacion else np.nan
    plan_prog = safe_sum(pa_y, ["programadas_a_corte"])
    plan_ejec = safe_sum(pa_y, ["ejecutadas_a_corte"])
    plan_cump = plan_ejec / plan_prog if plan_prog else safe_mean(pa_y, ["cumplimiento_a_corte", "cumplimiento_global"])

    cols = st.columns(5)
    with cols[0]: kpi("Ejecución presupuesto", ejec_pres, "Ejecutado / planeado", percent=True, state="auto")
    with cols[1]: kpi("Cobertura formación", form_cov, "Promedio registros", percent=True, state="auto")
    with cols[2]: kpi("Autodiagnóstico", auto_cump, "Cumplimiento estándares", percent=True, state="auto")
    with cols[3]: kpi("Plan PHVA", plan_cump, "Ejecución a corte", percent=True, state="auto")
    with cols[4]: kpi("Riesgos GTC-45", len(pel), "Peligros identificados", state="info")

    cols = st.columns(5)
    with cols[0]: kpi("EMO vencidos", count_text(emo_y, ["estado"], "venc"), "Alertas médicas", state="auto", inverse=True)
    with cols[1]: kpi("Accidentes", len(at_y), "Eventos registrados", state="info")
    with cols[2]: kpi("Enfermedad laboral", len(el_y), "Casos gestionados", state="info")
    with cols[3]: kpi("Reportes laborales", len(r), "Autogestión SST", state="info")
    with cols[4]: kpi("APCM vencidas", count_text(ap_y, ["estado"], "venc"), "Acciones críticas", state="auto", inverse=True)

    tabs = st.tabs(["📌 Semáforo ejecutivo", "📈 Tendencias", "🚨 Alertas", "📋 Datos fuente"])
    with tabs[0]:
        components = pd.DataFrame([
            {"Componente": "Presupuesto", "Resultado": ejec_pres * 100 if pd.notna(ejec_pres) else np.nan, "Meta": 85},
            {"Componente": "Formación", "Resultado": form_cov * 100 if pd.notna(form_cov) else np.nan, "Meta": 85},
            {"Componente": "Autodiagnóstico", "Resultado": auto_cump * 100 if pd.notna(auto_cump) else np.nan, "Meta": 85},
            {"Componente": "Plan PHVA", "Resultado": plan_cump * 100 if pd.notna(plan_cump) else np.nan, "Meta": 85},
            {"Componente": "Proveedores", "Resultado": safe_mean(pr_y, ["cumplimiento_total", "cumplimiento_promedio"]) * 100, "Meta": 85},
            {"Componente": "Indicadores", "Resultado": count_text(ind_y, ["estado"], "meta") / len(ind_y) * 100 if len(ind_y) else np.nan, "Meta": 85},
            {"Componente": "Mantenimiento", "Resultado": count_text(mt_y, ["estado"], "ejecut") / len(mt_y) * 100 if len(mt_y) else np.nan, "Meta": 85},
            {"Componente": "APCM", "Resultado": count_text(ap_y, ["estado"], "cerr") / len(ap_y) * 100 if len(ap_y) else np.nan, "Meta": 85},
        ]).dropna()
        components["Estado"] = np.where(components["Resultado"] >= 85, "En meta", np.where(components["Resultado"] >= 60, "Alerta", "Crítico"))
        c1, c2 = st.columns([1.8, 1])
        with c1:
            fig = px.bar(components, x="Componente", y="Resultado", color="Estado", text="Resultado", color_discrete_map=STATUS_COLORS, range_y=[0, 110])
            fig.add_hline(y=85, line_dash="dash", line_color="#16A34A", annotation_text="Meta 85%")
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig.update_layout(title=f"Semáforo de cumplimiento por componente — {selected_year}", showlegend=True)
            show_fig(fig, 460)
        with c2:
            donut_chart(components, "Estado", "Distribución del estado ejecutivo", 460)
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1:
            m = col(f, ["mes_planeado", "mes"])
            line_chart(f_y, m, col(f_y, ["horas_ejecutadas", "no_asistentes"]), "Tendencia formación por mes")
        with c2:
            m2 = col(mt_y, ["mes_programado", "mes"])
            line_chart(mt_y, m2, col(mt_y, ["costo_real"]), "Costo real de mantenimiento por mes")
        c1, c2 = st.columns(2)
        with c1:
            count_chart(at_y, col(at_y, ["mes"]), "Accidentalidad por mes")
        with c2:
            count_chart(ap_y, col(ap_y, ["mes"]), "Acciones APCM por mes")
    with tabs[2]:
        alerts = pd.DataFrame([
            {"Módulo": "EMO", "Alerta": "Exámenes vencidos", "Cantidad": count_text(emo_y, ["estado"], "venc"), "Severidad": "Alta"},
            {"Módulo": "APCM", "Alerta": "Acciones vencidas", "Cantidad": count_text(ap_y, ["estado"], "venc"), "Severidad": "Alta"},
            {"Módulo": "Mantenimiento", "Alerta": "Pendientes", "Cantidad": count_text(mt_y, ["estado"], "pend"), "Severidad": "Media"},
            {"Módulo": "Indicadores", "Alerta": "Indicadores críticos", "Cantidad": count_text(ind_y, ["estado"], "crit"), "Severidad": "Alta"},
            {"Módulo": "Reporte laboral", "Alerta": "Reportes abiertos / en gestión", "Cantidad": count_text(r, ["estado", "estado_gestion"], "abiert") + count_text(r, ["estado", "estado_gestion"], "gest"), "Severidad": "Media"},
            {"Módulo": "Matriz legal", "Alerta": "Requisitos no cumple / parcial", "Cantidad": count_text(l, ["cumple"], "no") + count_text(l, ["cumple"], "parcial"), "Severidad": "Alta"},
        ])
        c1, c2 = st.columns([1.5, 1])
        with c1:
            fig = px.bar(alerts, x="Módulo", y="Cantidad", color="Severidad", text="Cantidad", color_discrete_map={"Alta": "#DC2626", "Media": "#F59E0B", "Baja": "#16A34A"})
            fig.update_traces(textposition="outside")
            fig.update_layout(title="Alertas operativas del SG-SST")
            show_fig(fig, 430)
        with c2:
            st.dataframe(alerts, use_container_width=True, hide_index=True)
    with tabs[3]:
        st.dataframe(load_all_summary(), use_container_width=True, hide_index=True)


def render_presupuesto() -> None:
    hero("Presupuesto SST", "Control financiero del SG-SST: presupuesto planeado, ejecución real, variación y priorización por componente.", ["Financiero", "Ejecución", "Componentes SG-SST"])
    data = get_presupuesto(); d = data["detalle"]
    d = filter_data(d, "pres", [("Año", ["año", "ano"]), ("Mes", ["mes"]), ("Componente", ["componente_sg_sst"]), ("Estado", ["estado"]), ("Prioridad", ["prioridad"])])
    val_p = col(d, ["valor_presupuestado"]); val_e = col(d, ["valor_ejecutado"]); estado = col(d, ["estado"]); comp = col(d, ["componente_sg_sst"]); mes = col(d, ["mes"])
    total_p = safe_sum(d, ["valor_presupuestado"]); total_e = safe_sum(d, ["valor_ejecutado"]); pct = total_e / total_p if total_p else np.nan
    cols = st.columns(5)
    with cols[0]: kpi("Presupuesto planeado", total_p, money=True)
    with cols[1]: kpi("Ejecución registrada", total_e, money=True)
    with cols[2]: kpi("% ejecución", pct, percent=True, state="auto")
    with cols[3]: kpi("Variación", total_p - total_e, money=True)
    with cols[4]: kpi("Actividades", len(d))
    tabs = st.tabs(["Análisis", "Estado y prioridades", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: sum_chart(d, comp, val_p, "Presupuesto planeado por componente", money=True)
        with c2: sum_chart(d, comp, val_e, "Ejecución registrada por componente", money=True)
        c1, c2 = st.columns(2)
        with c1: line_chart(d, mes, val_p, "Presupuesto planeado por mes")
        with c2: line_chart(d, mes, val_e, "Ejecución real por mes")
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, estado, "Distribución por estado")
        with c2: stacked_status_by_group(d, comp, estado, "Estado por componente")
        heatmap_month_status(d, mes, estado, "Mapa de calor: estado por mes")
    with tabs[2]: table_with_download(d, "Detalle del presupuesto SST", "presupuesto")


def render_formacion() -> None:
    hero("Formación y capacitaciones", "Seguimiento del plan de formación: ejecución, cobertura, horas, efectividad y distribución por sedes y categorías.", ["Capacitación", "Cobertura", "Efectividad"])
    d = get_formacion()["registro"]
    d = filter_data(d, "form", [("Año", ["año", "ano"]), ("Sede", ["sede"]), ("Mes", ["mes_planeado", "mes"]), ("Categoría", ["categoria"]), ("Estado", ["estado"]), ("Público", ["publico_objeto"]), ("Efectividad", ["fue_efectiva"])] )
    estado = col(d, ["estado"]); sede = col(d, ["sede"]); categoria = col(d, ["categoria"]); mes = col(d, ["mes_planeado", "mes"]); cobertura = col(d, ["cobertura"]); horas = col(d, ["horas_ejecutadas"]); efectiva = col(d, ["fue_efectiva"])
    cols = st.columns(5)
    with cols[0]: kpi("Actividades", len(d))
    with cols[1]: kpi("Ejecutadas", count_text(d, ["estado"], "ejecut"))
    with cols[2]: kpi("Cobertura promedio", safe_mean(d, ["cobertura"]), percent=True, state="auto")
    with cols[3]: kpi("Horas ejecutadas", safe_sum(d, ["horas_ejecutadas"]))
    with cols[4]: kpi("Efectivas", count_text(d, ["fue_efectiva"], "sí") + count_text(d, ["fue_efectiva"], "si"))
    tabs = st.tabs(["Análisis", "Cobertura", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: count_chart(d, categoria, "Capacitaciones por categoría", orientation="h")
        with c2: donut_chart(d, estado, "Estado de ejecución")
        c1, c2 = st.columns(2)
        with c1: line_chart(d, mes, horas, "Horas ejecutadas por mes")
        with c2: count_chart(d, sede, "Actividades por sede", orientation="h")
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: mean_chart(d, sede, cobertura, "Cobertura promedio por sede")
        with c2: donut_chart(d, efectiva, "Resultado de efectividad")
        stacked_status_by_group(d, categoria, estado, "Estado por categoría")
    with tabs[2]: table_with_download(d, "Registro de formación", "formacion")


def render_autodiagnostico() -> None:
    hero("Autodiagnóstico Resolución 0312 de 2019", "Evaluación del cumplimiento de estándares mínimos del SG-SST por ciclo PHVA, puntaje y planes de mejora.", ["Resolución 0312", "PHVA", "Cumplimiento"])
    d = get_autodiagnostico()["detalle"]
    d = filter_data(d, "auto", [("Año", ["año", "ano"]), ("Sede", ["sede"]), ("Ciclo PHVA", ["ciclo_phva"]), ("Calificación", ["calificacion"]), ("Plan de mejora", ["plan_de_mejora_requerido"]), ("Estado mejora", ["estado_mejora"])] )
    ciclo = col(d, ["ciclo_phva"]); cal = col(d, ["calificacion"]); punt = col(d, ["puntaje_obtenido"]); pond = col(d, ["ponderacion"]); plan = col(d, ["plan_de_mejora_requerido"]); mejora = col(d, ["estado_mejora"])
    total_punt = safe_sum(d, ["puntaje_obtenido"]); total_pond = safe_sum(d, ["ponderacion"]); pct = total_punt / total_pond if total_pond else np.nan
    cols = st.columns(5)
    with cols[0]: kpi("Estándares", len(d))
    with cols[1]: kpi("Puntaje", total_punt)
    with cols[2]: kpi("Cumplimiento", pct, percent=True, state="auto")
    with cols[3]: kpi("Cumplen", count_text(d, ["calificacion"], "cumple"))
    with cols[4]: kpi("Planes de mejora", count_text(d, ["plan_de_mejora_requerido"], "sí") + count_text(d, ["plan_de_mejora_requerido"], "si"))
    tabs = st.tabs(["Análisis", "Planes de mejora", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: sum_chart(d, ciclo, punt, "Puntaje obtenido por ciclo PHVA")
        with c2: donut_chart(d, cal, "Cumplimiento de estándares")
        if ciclo and punt and pond:
            g = d.groupby(ciclo, dropna=False).agg(punt=(punt, lambda x: pd.to_numeric(x, errors="coerce").sum()), pond=(pond, lambda x: pd.to_numeric(x, errors="coerce").sum())).reset_index()
            g["cumplimiento"] = g["punt"] / g["pond"]
            fig = px.bar(g, x=ciclo, y="cumplimiento", text="cumplimiento", color=ciclo, color_discrete_sequence=PALETTE, range_y=[0, 1.1])
            fig.update_traces(texttemplate="%{text:.1%}", textposition="outside")
            fig.update_layout(title="% cumplimiento por ciclo PHVA", showlegend=False)
            show_fig(fig)
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, plan, "Estándares que requieren plan de mejora")
        with c2: donut_chart(d, mejora, "Estado de los planes de mejora")
        stacked_status_by_group(d, ciclo, mejora, "Estado de mejora por ciclo")
    with tabs[2]: table_with_download(d, "Detalle autodiagnóstico", "autodiagnostico")


def render_plan_anual() -> None:
    hero("Plan anual de trabajo PHVA", "Control de actividades del SG-SST por ciclo, componente, responsable, programación y ejecución a corte.", ["Planeación", "PHVA", "Seguimiento"])
    d = get_plan_anual()["detalle"]
    d = filter_data(d, "plan", [("Año", ["año", "ano"]), ("Sede", ["sede_cobertura", "sede"]), ("Ciclo PHVA", ["ciclo_phva"]), ("Componente", ["componente_sg_sst"]), ("Estado", ["estado_general"]), ("Responsable", ["responsable"])] )
    ciclo = col(d, ["ciclo_phva"]); comp = col(d, ["componente_sg_sst"]); estado = col(d, ["estado_general"]); resp = col(d, ["responsable"]); prog = col(d, ["programadas_a_corte"]); ejec = col(d, ["ejecutadas_a_corte"]); pct_c = col(d, ["cumplimiento_a_corte"])
    total_prog = safe_sum(d, ["programadas_a_corte"]); total_ejec = safe_sum(d, ["ejecutadas_a_corte"]); pct = total_ejec / total_prog if total_prog else safe_mean(d, ["cumplimiento_a_corte"])
    cols = st.columns(5)
    with cols[0]: kpi("Actividades", len(d))
    with cols[1]: kpi("Programadas a corte", total_prog)
    with cols[2]: kpi("Ejecutadas a corte", total_ejec)
    with cols[3]: kpi("Cumplimiento", pct, percent=True, state="auto")
    with cols[4]: kpi("Presupuesto estimado", safe_sum(d, ["rec_financieros_estimados_cop", "presupuesto_estimado_cop"]), money=True)
    tabs = st.tabs(["Análisis", "Responsables y estado", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: count_chart(d, ciclo, "Actividades por ciclo PHVA")
        with c2: count_chart(d, comp, "Actividades por componente", orientation="h")
        c1, c2 = st.columns(2)
        with c1: sum_chart(d, ciclo, ejec, "Ejecución a corte por ciclo")
        with c2: mean_chart(d, comp, pct_c, "% cumplimiento promedio por componente")
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, estado, "Estado general del plan")
        with c2: count_chart(d, resp, "Carga por responsable", orientation="h")
        stacked_status_by_group(d, ciclo, estado, "Estado por ciclo PHVA")
    with tabs[2]: table_with_download(d, "Detalle plan anual", "plan_anual")


def render_legal() -> None:
    hero("Matriz legal SGRL", "Verificación de requisitos legales aplicables al SG-SST, estado normativo, cumplimiento y evidencias sugeridas.", ["Cumplimiento legal", "SGRL", "Normatividad"])
    d = get_legal()["matriz"]
    d = filter_data(d, "legal", [("Tipo documento", ["tipo_documento"]), ("Año norma", ["año", "ano"]), ("Tema", ["tema_eje", "tema"]), ("Estado normativo", ["estado_normativo"]), ("Aplica", ["aplica_a_la_empresa"]), ("Cumple", ["cumple"])] )
    tipo = col(d, ["tipo_documento"]); tema = col(d, ["tema_eje", "tema"]); estado_norma = col(d, ["estado_normativo"]); aplica = col(d, ["aplica_a_la_empresa"]); cumple = col(d, ["cumple"])
    cols = st.columns(5)
    with cols[0]: kpi("Normas / requisitos", len(d))
    with cols[1]: kpi("Aplicables", count_text(d, ["aplica_a_la_empresa"], "sí") + count_text(d, ["aplica_a_la_empresa"], "si"))
    with cols[2]: kpi("Cumplen", count_text(d, ["cumple"], "sí") + count_text(d, ["cumple"], "si"), state="good")
    with cols[3]: kpi("Parcial / no cumple", count_text(d, ["cumple"], "parcial") + count_text(d, ["cumple"], "no"), state="auto", inverse=True)
    with cols[4]: kpi("Temas", len(unique_values(d, tema)))
    tabs = st.tabs(["Análisis", "Priorización", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, cumple, "Estado de cumplimiento")
        with c2: donut_chart(d, aplica, "Aplicabilidad")
        c1, c2 = st.columns(2)
        with c1: count_chart(d, tipo, "Normas por tipo de documento")
        with c2: count_chart(d, estado_norma, "Estado normativo")
    with tabs[1]:
        count_chart(d, tema, "Requisitos por tema / eje", orientation="h", top=15)
        stacked_status_by_group(d, tema, cumple, "Cumplimiento por tema")
    with tabs[2]: table_with_download(d, "Matriz legal completa", "legal")


def render_reporte_laboral() -> None:
    hero("Reporte laboral de autogestión", "Canal participativo para reportar condiciones de salud, actos inseguros, condiciones inseguras, incidentes y accidentes.", ["Autogestión", "Participación", "Reportes"])
    data = get_reporte_laboral(); d = data["gestion"] if not data["gestion"].empty else data["respuestas"]
    d = filter_data(d, "rep", [("Sede", ["sede"]), ("Área", ["area", "area_o_proceso"]), ("Tipo reporte", ["tipo_reporte"]), ("Estado", ["estado", "estado_gestion"]), ("Riesgo inmediato", ["riesgo_inmediato", "existe_riesgo_inmediato_para_usted_u_otras_personas"])] )
    tipo = col(d, ["tipo_reporte"]); estado = col(d, ["estado", "estado_gestion"]); sede = col(d, ["sede"]); area = col(d, ["area", "area_o_proceso"]); riesgo = col(d, ["riesgo_inmediato", "existe_riesgo_inmediato_para_usted_u_otras_personas"])
    cols = st.columns(5)
    with cols[0]: kpi("Reportes", len(d))
    with cols[1]: kpi("Riesgo inmediato", count_text(d, ["riesgo_inmediato", "existe_riesgo_inmediato_para_usted_u_otras_personas"], "sí") + count_text(d, ["riesgo_inmediato", "existe_riesgo_inmediato_para_usted_u_otras_personas"], "si"), state="warn")
    with cols[2]: kpi("Abiertos / gestión", count_text(d, ["estado", "estado_gestion"], "abiert") + count_text(d, ["estado", "estado_gestion"], "gest"), state="warn")
    with cols[3]: kpi("Cerrados", count_text(d, ["estado", "estado_gestion"], "cerr"), state="good")
    with cols[4]: kpi("Áreas", len(unique_values(d, area)))
    tabs = st.tabs(["Análisis", "Gestión", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, tipo, "Reportes por tipo")
        with c2: count_chart(d, sede, "Reportes por sede")
        c1, c2 = st.columns(2)
        with c1: count_chart(d, area, "Reportes por área / proceso", orientation="h")
        with c2: donut_chart(d, riesgo, "Riesgo inmediato")
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, estado, "Estado de gestión")
        with c2: stacked_status_by_group(d, tipo, estado, "Estado por tipo de reporte")
    with tabs[2]: table_with_download(d, "Detalle reportes laborales", "reporte_laboral")


def render_proveedores() -> None:
    hero("Proveedores y contratistas", "Evaluación de requisitos SST, nivel de riesgo, cumplimiento y estado de proveedores/contratistas.", ["Contratistas", "Cumplimiento SST", "Riesgo"])
    d = get_proveedores()["detalle"]
    d = filter_data(d, "prov", [("Año", ["año", "ano"]), ("Sede", ["sede"]), ("Nivel de riesgo", ["nivel_de_riesgo"]), ("Estado", ["estado", "estado_final"]), ("Servicio", ["servicio_contratado"])] )
    riesgo = col(d, ["nivel_de_riesgo"]); estado = col(d, ["estado", "estado_final"]); sede = col(d, ["sede"]); servicio = col(d, ["servicio_contratado"]); cumplimiento = col(d, ["cumplimiento_total", "cumplimiento_promedio"])
    cols = st.columns(5)
    with cols[0]: kpi("Proveedores", len(d))
    with cols[1]: kpi("Cumplimiento promedio", safe_mean(d, ["cumplimiento_total", "cumplimiento_promedio"]), percent=True, state="auto")
    with cols[2]: kpi("Riesgo alto", count_text(d, ["nivel_de_riesgo"], "alto"), state="warn")
    with cols[3]: kpi("Aprobados", count_text(d, ["estado", "estado_final"], "aprob"), state="good")
    with cols[4]: kpi("Condicionados / no", count_text(d, ["estado", "estado_final"], "cond") + count_text(d, ["estado", "estado_final"], "no"), state="auto", inverse=True)
    tabs = st.tabs(["Análisis", "Cumplimiento", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, riesgo, "Distribución por nivel de riesgo")
        with c2: donut_chart(d, estado, "Estado de proveedores")
        c1, c2 = st.columns(2)
        with c1: count_chart(d, sede, "Proveedores por sede")
        with c2: count_chart(d, servicio, "Servicios contratados", orientation="h", top=12)
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: mean_chart(d, sede, cumplimiento, "Cumplimiento promedio por sede")
        with c2: mean_chart(d, riesgo, cumplimiento, "Cumplimiento por nivel de riesgo")
        stacked_status_by_group(d, riesgo, estado, "Estado por nivel de riesgo")
    with tabs[2]: table_with_download(d, "Detalle proveedores", "proveedores")


def render_emo() -> None:
    hero("Exámenes médicos ocupacionales", "Control de exámenes ocupacionales, resultados, restricciones, seguimientos y vencimientos.", ["Medicina laboral", "Vencimientos", "Restricciones"])
    d = get_emo()["detalle"]
    d = filter_data(d, "emo", [("Año", ["año", "ano"]), ("Sede", ["sede"]), ("Tipo examen", ["tipo_de_examen"]), ("Resultado", ["resultado"]), ("Estado", ["estado"]), ("Área", ["area_proceso", "area"])] )
    sede = col(d, ["sede"]); tipo = col(d, ["tipo_de_examen"]); resultado = col(d, ["resultado"]); estado = col(d, ["estado"]); area = col(d, ["area_proceso", "area"]); seguimiento = col(d, ["seguimiento", "requiere_seguimiento"])
    cols = st.columns(5)
    with cols[0]: kpi("Exámenes", len(d))
    with cols[1]: kpi("Vigentes", count_text(d, ["estado"], "vigente"), state="good")
    with cols[2]: kpi("Próximos a vencer", count_text(d, ["estado"], "proximo") + count_text(d, ["estado"], "próximo"), state="warn")
    with cols[3]: kpi("Vencidos", count_text(d, ["estado"], "venc"), state="auto", inverse=True)
    with cols[4]: kpi("Con seguimiento", count_text(d, ["seguimiento", "requiere_seguimiento"], "sí") + count_text(d, ["seguimiento", "requiere_seguimiento"], "si"), state="info")
    tabs = st.tabs(["Análisis", "Vencimientos", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, resultado, "Resultado de exámenes")
        with c2: count_chart(d, tipo, "Tipo de examen")
        c1, c2 = st.columns(2)
        with c1: count_chart(d, sede, "Exámenes por sede")
        with c2: count_chart(d, area, "Exámenes por área", orientation="h")
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, estado, "Estado de vencimiento")
        with c2: stacked_status_by_group(d, sede, estado, "Vencimientos por sede")
        stacked_status_by_group(d, resultado, estado, "Estado por resultado médico")
    with tabs[2]: table_with_download(d, "Detalle EMO", "emo")


def render_accidentalidad() -> None:
    hero("Accidentalidad", "Análisis de accidentes de trabajo: frecuencia, severidad, mecanismos, causas, acciones y cierres.", ["AT", "IF / IS / ILI", "Causalidad"])
    d = get_accidentalidad()["detalle"]
    d = filter_data(d, "at", [("Año", ["año", "ano"]), ("Sede", ["sede"]), ("Mes", ["mes"]), ("Área", ["area_proceso", "area"]), ("Tipo evento", ["tipo_de_evento", "tipo_accidente"]), ("Estado evento", ["estado_evento"]), ("Estado acción", ["estado_accion"])] )
    sede = col(d, ["sede"]); mes = col(d, ["mes"]); area = col(d, ["area_proceso", "area"]); tipo = col(d, ["tipo_de_evento", "tipo_accidente"]); mecanismo = col(d, ["mecanismo_del_accidente", "mecanismo"]); causa = col(d, ["causa_raiz", "causas_basicas"]); dias = col(d, ["dias_de_incapacidad", "dias_perdidos"]); estado_acc = col(d, ["estado_accion"])
    total_dias = safe_sum(d, ["dias_de_incapacidad", "dias_perdidos"])
    cols = st.columns(5)
    with cols[0]: kpi("Accidentes", len(d), state="info")
    with cols[1]: kpi("Días perdidos", total_dias, state="warn")
    with cols[2]: kpi("Con incapacidad", count_text(d, ["tipo_de_evento", "tipo_accidente"], "incap"), state="warn")
    with cols[3]: kpi("Investigados", count_text(d, ["investigacion_realizada"], "sí") + count_text(d, ["investigacion_realizada"], "si"), state="good")
    with cols[4]: kpi("Acciones cerradas", count_text(d, ["estado_accion"], "cerr"), state="good")
    tabs = st.tabs(["Análisis", "Causas y acciones", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: count_chart(d, mes, "Accidentes por mes")
        with c2: count_chart(d, sede, "Accidentes por sede")
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, tipo, "Tipo de evento")
        with c2: count_chart(d, mecanismo, "Mecanismo del accidente", orientation="h", top=12)
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: count_chart(d, causa, "Causas raíz recurrentes", orientation="h", top=12)
        with c2: donut_chart(d, estado_acc, "Estado de acciones")
        stacked_status_by_group(d, area, estado_acc, "Estado de acciones por área")
    with tabs[2]: table_with_download(d, "Detalle accidentalidad", "accidentalidad")


def render_enfermedad() -> None:
    hero("Enfermedad laboral", "Gestión de casos de enfermedad laboral: diagnóstico, calificación de origen, seguimiento, PVE, reubicación y cierre.", ["EL", "Vigilancia epidemiológica", "Reubicación"])
    d = get_enfermedad()["detalle"]
    d = filter_data(d, "el", [("Año", ["año", "ano"]), ("Sede", ["sede"]), ("Área", ["area_proceso", "area"]), ("Factor riesgo", ["factor_riesgo_asociado", "factores_riesgo"]), ("Resultado calificación", ["resultado_calificacion"]), ("Estado caso", ["estado_caso"])] )
    sede = col(d, ["sede"]); area = col(d, ["area_proceso", "area"]); factor = col(d, ["factor_riesgo_asociado", "factores_riesgo"]); result = col(d, ["resultado_calificacion"]); estado = col(d, ["estado_caso"]); diag = col(d, ["diagnostico_confirmado", "diagnostico_presuntivo"]); dias = col(d, ["dias_gestion"])
    cols = st.columns(5)
    with cols[0]: kpi("Casos", len(d))
    with cols[1]: kpi("Laborales", count_text(d, ["resultado_calificacion"], "laboral"), state="warn")
    with cols[2]: kpi("Cerrados", count_text(d, ["estado_caso"], "cerr"), state="good")
    with cols[3]: kpi("Prom. días gestión", safe_mean(d, ["dias_gestion"]))
    with cols[4]: kpi("PVE activo", count_text(d, ["pve_activado", "plan_de_vigilancia"], "sí") + count_text(d, ["pve_activado", "plan_de_vigilancia"], "si"), state="info")
    tabs = st.tabs(["Análisis", "Seguimiento", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, factor, "Casos por factor de riesgo")
        with c2: donut_chart(d, result, "Resultado de calificación")
        c1, c2 = st.columns(2)
        with c1: count_chart(d, sede, "Casos por sede")
        with c2: count_chart(d, diag, "Diagnósticos registrados", orientation="h", top=12)
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, estado, "Estado del caso")
        with c2: mean_chart(d, factor, dias, "Promedio días de gestión por factor")
        stacked_status_by_group(d, area, estado, "Estado por área")
    with tabs[2]: table_with_download(d, "Detalle enfermedad laboral", "enfermedad_laboral")


def render_indicadores() -> None:
    hero("Indicadores de gestión del SG-SST", "Medición de indicadores de estructura, proceso, resultado e impacto con metas, tendencias y semáforos.", ["KPIs", "Semáforos", "Tendencias"])
    data = get_indicadores(); d = data["medicion"]
    d = filter_data(d, "ind", [("Año", ["año", "ano"]), ("Mes", ["mes"]), ("Tipo", ["tipo"]), ("Indicador", ["indicador"]), ("Estado", ["estado"]), ("Frecuencia", ["frecuencia"])] )
    tipo = col(d, ["tipo"]); indicador = col(d, ["indicador"]); estado = col(d, ["estado"]); mes = col(d, ["mes"]); resultado = col(d, ["resultado"]); meta = col(d, ["meta"])
    in_meta = count_text(d, ["estado"], "meta")
    cols = st.columns(5)
    with cols[0]: kpi("Mediciones", len(d))
    with cols[1]: kpi("En meta", in_meta, state="good")
    with cols[2]: kpi("% en meta", in_meta / len(d) if len(d) else np.nan, percent=True, state="auto")
    with cols[3]: kpi("Alerta", count_text(d, ["estado"], "alerta"), state="warn")
    with cols[4]: kpi("Críticos", count_text(d, ["estado"], "crit"), state="bad")
    tabs = st.tabs(["Análisis", "Tendencias", "Ficha técnica / detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, estado, "Estado de indicadores")
        with c2: donut_chart(d, tipo, "Indicadores por tipo")
        c1, c2 = st.columns(2)
        with c1: count_chart(d, indicador, "Mediciones por indicador", orientation="h", top=15)
        with c2: mean_chart(d, tipo, resultado, "Resultado promedio por tipo")
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: line_chart(d, mes, resultado, "Resultado promedio mensual", color_col=tipo, agg="mean")
        with c2: line_chart(d, mes, meta, "Meta promedio mensual", color_col=tipo, agg="mean")
        stacked_status_by_group(d, indicador, estado, "Semáforo por indicador")
    with tabs[2]:
        table_with_download(data["ficha"], "Ficha técnica", "ficha_indicadores", 320)
        table_with_download(d, "Mediciones del periodo", "indicadores", 420)


def render_peligros() -> None:
    hero("Matriz de peligros y riesgos GTC-45", "Identificación de peligros, valoración de riesgos, aceptabilidad y medidas de intervención conforme a GTC-45:2012.", ["GTC-45", "Riesgo", "Controles"])
    d = get_peligros()["matriz"]
    d = filter_data(d, "pel", [("Sede", ["sede"]), ("Proceso", ["proceso"]), ("Clasificación", ["clasificacion_del_peligro"]), ("Nivel riesgo", ["nivel_de_riesgo_e_intervencion", "interpretacion_nr"]), ("Aceptabilidad", ["aceptabilidad_del_riesgo"]), ("Responsable", ["responsable"])] )
    sede = col(d, ["sede"]); proceso = col(d, ["proceso"]); clas = col(d, ["clasificacion_del_peligro"]); nivel = col(d, ["nivel_de_riesgo_e_intervencion", "interpretacion_nr"]); acept = col(d, ["aceptabilidad_del_riesgo"]); estado = col(d, ["estado_de_implementacion", "estado"]); nr = col(d, ["nivel_de_riesgo_nr_np_nc"]); expuestos = col(d, ["n_o_expuestos"])
    cols = st.columns(5)
    with cols[0]: kpi("Peligros", len(d))
    with cols[1]: kpi("No aceptables", count_text(d, ["aceptabilidad_del_riesgo"], "no"), state="bad")
    with cols[2]: kpi("Nivel I/II", count_text(d, ["nivel_de_riesgo_e_intervencion", "interpretacion_nr"], "i") + count_text(d, ["nivel_de_riesgo_e_intervencion", "interpretacion_nr"], "ii"), state="warn")
    with cols[3]: kpi("Expuestos", safe_sum(d, ["n_o_expuestos"]))
    with cols[4]: kpi("Procesos", len(unique_values(d, proceso)))
    tabs = st.tabs(["Análisis", "Priorización", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, clas, "Clasificación del peligro")
        with c2: donut_chart(d, acept, "Aceptabilidad del riesgo")
        c1, c2 = st.columns(2)
        with c1: count_chart(d, proceso, "Peligros por proceso", orientation="h")
        with c2: count_chart(d, sede, "Peligros por sede")
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: count_chart(d, nivel, "Distribución por nivel de riesgo", orientation="h")
        with c2: sum_chart(d, clas, expuestos, "Expuestos por clasificación de peligro", orientation="h")
        if nr and expuestos and proceso:
            dd = d.copy(); dd[nr] = pd.to_numeric(dd[nr], errors="coerce"); dd[expuestos] = pd.to_numeric(dd[expuestos], errors="coerce")
            fig = px.scatter(dd, x=expuestos, y=nr, color=clas, size=nr, hover_name=proceso, color_discrete_sequence=PALETTE, title="Mapa de priorización: expuestos vs nivel de riesgo")
            show_fig(fig, 500)
    with tabs[2]: table_with_download(d, "Detalle matriz GTC-45", "peligros")


def render_mantenimiento() -> None:
    hero("Mantenimiento preventivo", "Control de mantenimientos programados, ejecución, oportunidad, costos, activos críticos y alertas por sede.", ["Infraestructura", "Costos", "Oportunidad"])
    d = get_mantenimiento()["detalle"]
    d = filter_data(d, "mant", [("Año", ["año", "ano"]), ("Sede", ["sede"]), ("Mes", ["mes_programado", "mes"]), ("Estado", ["estado"]), ("Tipo activo", ["tipo_de_activo"]), ("Criticidad", ["criticidad"]), ("Frecuencia", ["frecuencia"])] )
    estado = col(d, ["estado"]); sede = col(d, ["sede"]); mes = col(d, ["mes_programado", "mes"]); tipo_activo = col(d, ["tipo_de_activo"]); criticidad = col(d, ["criticidad"]); frecuencia = col(d, ["frecuencia"]); costo_est = col(d, ["costo_estimado"]); costo_real = col(d, ["costo_real"]); oportuno = col(d, ["cumplimiento_oportuno"])
    total = len(d); ejecutados = count_text(d, ["estado"], "ejecut"); pct = ejecutados / total if total else np.nan
    cols = st.columns(5)
    with cols[0]: kpi("Programados", total)
    with cols[1]: kpi("Ejecutados", ejecutados, state="good")
    with cols[2]: kpi("% ejecución", pct, percent=True, state="auto")
    with cols[3]: kpi("Costo real", safe_sum(d, ["costo_real"]), money=True)
    with cols[4]: kpi("Oportunos", count_text(d, ["cumplimiento_oportuno"], "sí") + count_text(d, ["cumplimiento_oportuno"], "si"), state="good")
    tabs = st.tabs(["Análisis", "Costos y oportunidad", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, estado, "Estado del mantenimiento")
        with c2: count_chart(d, tipo_activo, "Mantenimientos por tipo de activo", orientation="h")
        c1, c2 = st.columns(2)
        with c1: count_chart(d, sede, "Mantenimientos por sede")
        with c2: count_chart(d, criticidad, "Criticidad de activos")
        heatmap_month_status(d, mes, estado, "Mapa de calor: estado por mes")
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: sum_chart(d, sede, costo_real, "Costo real por sede", money=True)
        with c2: sum_chart(d, tipo_activo, costo_real, "Costo real por tipo de activo", orientation="h", money=True)
        c1, c2 = st.columns(2)
        with c1: line_chart(d, mes, costo_real, "Costo real por mes")
        with c2: donut_chart(d, oportuno, "Cumplimiento oportuno")
        stacked_status_by_group(d, frecuencia, estado, "Estado por frecuencia")
    with tabs[2]: table_with_download(d, "Detalle mantenimiento preventivo", "mantenimiento")


def render_apcm() -> None:
    hero("Acciones preventivas, correctivas y de mejora", "Gestión de acciones derivadas de auditorías, inspecciones, investigaciones, reportes y revisión gerencial.", ["Mejora continua", "Cierre", "Eficacia"])
    d = get_apcm()["detalle"]
    d = filter_data(d, "apcm", [("Año", ["año", "ano"]), ("Sede", ["sede"]), ("Proceso", ["proceso"]), ("Tipo", ["tipo"]), ("Fuente", ["fuente_de_origen"]), ("Estado", ["estado"]), ("Prioridad", ["prioridad"])] )
    estado = col(d, ["estado"]); tipo = col(d, ["tipo"]); fuente = col(d, ["fuente_de_origen"]); sede = col(d, ["sede"]); proceso = col(d, ["proceso"]); prioridad = col(d, ["prioridad"]); avance = col(d, ["avance"]); eficacia = col(d, ["eficacia_verificada"]); mes = col(d, ["mes"]); costo = col(d, ["costo_ejecutado"])
    total = len(d); cerradas = count_text(d, ["estado"], "cerr"); cierre = cerradas / total if total else np.nan
    cols = st.columns(5)
    with cols[0]: kpi("Acciones", total)
    with cols[1]: kpi("Cerradas", cerradas, state="good")
    with cols[2]: kpi("% cierre", cierre, percent=True, state="auto")
    with cols[3]: kpi("Vencidas", count_text(d, ["estado"], "venc"), state="auto", inverse=True)
    with cols[4]: kpi("Eficaces", count_text(d, ["eficacia_verificada"], "sí") + count_text(d, ["eficacia_verificada"], "si"), state="good")
    tabs = st.tabs(["Análisis", "Cierre y eficacia", "Detalle"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1: donut_chart(d, estado, "Estado de acciones")
        with c2: donut_chart(d, tipo, "Tipo de acción")
        c1, c2 = st.columns(2)
        with c1: count_chart(d, fuente, "Fuente de origen", orientation="h", top=12)
        with c2: count_chart(d, prioridad, "Prioridad")
    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1: stacked_status_by_group(d, proceso, estado, "Estado por proceso")
        with c2: donut_chart(d, eficacia, "Eficacia verificada")
        c1, c2 = st.columns(2)
        with c1: line_chart(d, mes, avance, "% avance promedio mensual", agg="mean")
        with c2: sum_chart(d, sede, costo, "Costo ejecutado por sede", money=True)
    with tabs[2]: table_with_download(d, "Detalle APCM", "apcm")


def render_calidad() -> None:
    hero("Carga y calidad de datos", "Control técnico de archivos, hojas cargadas, filas disponibles y trazabilidad de fuentes del tablero.", ["Diagnóstico", "Carga", "Validación"])
    summary = load_all_summary()
    cols = st.columns(4)
    with cols[0]: kpi("Archivos esperados", len(FILES))
    with cols[1]: kpi("Archivos encontrados", summary.drop_duplicates("Archivo")["Estado"].eq("Encontrado").sum(), state="good")
    with cols[2]: kpi("Hojas cargadas", len(summary[summary["Filas"] > 0]), state="info")
    with cols[3]: kpi("Registros disponibles", summary["Filas"].sum())
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(summary, x="Módulo", y="Filas", color="Estado", text="Filas", color_discrete_map={"Encontrado": "#16A34A", "No encontrado": "#DC2626"})
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_layout(title="Filas cargadas por módulo")
        show_fig(fig, 520)
    with c2:
        donut_chart(summary, "Estado", "Estado de archivos fuente", 520)
    table_with_download(summary, "Inventario de hojas y registros", "calidad")

# =============================================================================
# Navegación
# =============================================================================

st.sidebar.markdown("### IMA Company SAS")
st.sidebar.caption("Dashboard ejecutivo SG-SST")

page_keys = [
    "resumen", "calidad", "presupuesto", "formacion", "autodiagnostico", "plan_anual", "legal",
    "reporte_laboral", "proveedores", "emo", "accidentalidad", "enfermedad", "indicadores",
    "peligros", "mantenimiento", "apcm",
]
page = st.sidebar.radio("Seleccione dashboard", page_keys, format_func=lambda k: MODULE_LABELS[k])
st.sidebar.divider()
missing = [MODULE_LABELS[k] for k in FILES if not module_available(k)]
if missing:
    st.sidebar.error("Archivos faltantes: " + ", ".join(missing))
else:
    st.sidebar.success("Todos los archivos fuente están disponibles.")
st.sidebar.caption("Los gráficos son interactivos: puede hacer zoom, filtrar leyendas, descargar imagen y explorar datos con hover.")

ROUTERS = {
    "resumen": render_resumen,
    "calidad": render_calidad,
    "presupuesto": render_presupuesto,
    "formacion": render_formacion,
    "autodiagnostico": render_autodiagnostico,
    "plan_anual": render_plan_anual,
    "legal": render_legal,
    "reporte_laboral": render_reporte_laboral,
    "proveedores": render_proveedores,
    "emo": render_emo,
    "accidentalidad": render_accidentalidad,
    "enfermedad": render_enfermedad,
    "indicadores": render_indicadores,
    "peligros": render_peligros,
    "mantenimiento": render_mantenimiento,
    "apcm": render_apcm,
}

ROUTERS[page]()
