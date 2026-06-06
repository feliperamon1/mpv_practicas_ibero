from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

APP_TITLE = "Dashboard Integral SG-SST | IMA Company SAS"
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
    "apcm": "Matriz_APCM_Vigilancia_2023_2026.xlsx",
    "mantenimiento": "Matriz_Mantenimiento_Preventivo_Vigilancia_2023_2026.xlsx",
}

MODULE_LABELS = {
    "resumen": "Resumen ejecutivo",
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
    "apcm": "14. Acciones preventivas, correctivas y de mejora",
}

MONTHS = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    .main .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    .kpi-card {background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 14px; padding: 18px; min-height: 112px;}
    .kpi-label {font-size: 0.86rem; color: #475569; margin-bottom: 8px;}
    .kpi-value {font-size: 1.55rem; font-weight: 800; color: #0F172A;}
    .kpi-help {font-size: 0.75rem; color: #64748B; margin-top: 6px;}
    .section-title {font-size: 1.15rem; font-weight: 800; margin-top: 1.0rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Utilidades generales
# -----------------------------------------------------------------------------

def strip_accents(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value))
    return "".join(ch for ch in value if not unicodedata.combining(ch))


def norm_col(value: Any) -> str:
    value = strip_accents(str(value or "").strip().lower())
    value = re.sub(r"[^a-z0-9]+", "_", value)
    value = re.sub(r"_+", "_", value).strip("_")
    return value or "columna"


def clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    seen: dict[str, int] = {}
    cols = []
    for c in df.columns:
        base = norm_col(c)
        seen[base] = seen.get(base, 0) + 1
        cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
    df.columns = cols
    df = df.dropna(how="all")
    return df


def file_path(key: str) -> Path:
    filename = FILES[key]
    direct = DATA_DIR / filename
    if direct.exists():
        return direct
    # Búsqueda tolerante a tildes, espacios y variaciones del nombre.
    if DATA_DIR.exists():
        target = norm_col(filename)
        for p in DATA_DIR.glob("*.xlsx"):
            if norm_col(p.name) == target:
                return p
            if norm_col(filename.split(".")[0]) in norm_col(p.stem):
                return p
    return direct


def module_available(key: str) -> bool:
    return file_path(key).exists()


@st.cache_data(show_spinner=False)
def read_excel_cached(path: str, sheet_name: str, header: int = 0) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=header, engine="openpyxl")
    except Exception:
        return pd.DataFrame()
    return clean_headers(df)


def load(key: str, sheet: str, header: int = 0) -> pd.DataFrame:
    path = file_path(key)
    if not path.exists():
        return pd.DataFrame()
    return read_excel_cached(str(path), sheet, header)


def load_many(key: str, sheets: Iterable[str], header: int = 0) -> pd.DataFrame:
    frames = []
    for sheet in sheets:
        df = load(key, sheet, header)
        if not df.empty:
            df["_hoja_origen"] = sheet
            frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def to_num(series: pd.Series | Any) -> pd.Series | float:
    if isinstance(series, pd.Series):
        if series.empty:
            return series
        return pd.to_numeric(series, errors="coerce")
    try:
        return float(series)
    except Exception:
        return np.nan


def excel_date_to_datetime(s: pd.Series) -> pd.Series:
    # Convierte seriales Excel solo cuando la columna representa fecha/timestamp.
    if s.empty:
        return s
    numeric = pd.to_numeric(s, errors="coerce")
    converted = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
    parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return parsed.fillna(converted)


def prepare_dates(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if any(token in col for token in ["fecha", "timestamp"]):
            df[col] = excel_date_to_datetime(df[col])
    return df


def best_col(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    if df is None or df.empty:
        return None
    cols = set(df.columns)
    for cand in candidates:
        n = norm_col(cand)
        if n in cols:
            return n
    for cand in candidates:
        n = norm_col(cand)
        for c in df.columns:
            if n and n in c:
                return c
    return None


def filter_year(df: pd.DataFrame, year: int | None) -> pd.DataFrame:
    if df.empty or year is None:
        return df
    c = best_col(df, ["año", "ano"])
    if not c:
        return df
    return df[to_num(df[c]) == year]


def filter_sede(df: pd.DataFrame, sede: str | None) -> pd.DataFrame:
    if df.empty or not sede or sede == "Todas":
        return df
    c = best_col(df, ["sede", "sede_cobertura"])
    if not c:
        return df
    return df[df[c].astype(str).str.strip() == sede]


def available_years(*dfs: pd.DataFrame) -> list[int]:
    vals: set[int] = set()
    for df in dfs:
        if df is None or df.empty:
            continue
        c = best_col(df, ["año", "ano"])
        if c:
            vals.update(to_num(df[c]).dropna().astype(int).tolist())
    return sorted(vals) if vals else [2026]


def sede_options(df: pd.DataFrame) -> list[str]:
    c = best_col(df, ["sede", "sede_cobertura"])
    if not c or df.empty:
        return ["Todas"]
    vals = sorted([x for x in df[c].dropna().astype(str).str.strip().unique() if x and x.lower() != "nan"])
    return ["Todas"] + vals


def latest_row_by_year(df: pd.DataFrame, year: int | None = None) -> pd.Series:
    if df.empty:
        return pd.Series(dtype="object")
    d = filter_year(df, year) if year else df
    if d.empty:
        d = df
    return d.iloc[-1]


def fmt_value(value: Any, as_percent: bool = False, money: bool = False, decimals: int = 1) -> str:
    if isinstance(value, pd.Series):
        value = value.iloc[0] if not value.empty else np.nan
    try:
        num = float(value)
        if as_percent:
            if abs(num) <= 1.5:
                num *= 100
            return f"{num:,.{decimals}f}%".replace(",", "X").replace(".", ",").replace("X", ".")
        if money:
            return "$" + f"{num:,.0f}".replace(",", ".")
        if num.is_integer():
            return f"{num:,.0f}".replace(",", ".")
        return f"{num:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        if pd.isna(value):
            return "—"
        return str(value)


def kpi(label: str, value: Any, help_text: str = "", as_percent: bool = False, money: bool = False) -> None:
    st.markdown(
        f"""
        <div class='kpi-card'>
            <div class='kpi-label'>{label}</div>
            <div class='kpi-value'>{fmt_value(value, as_percent=as_percent, money=money)}</div>
            <div class='kpi-help'>{help_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section(title: str) -> None:
    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)


def show_table(df: pd.DataFrame, title: str = "Detalle", height: int = 380) -> None:
    with st.expander(title, expanded=False):
        if df.empty:
            st.info("No hay datos disponibles para esta tabla.")
        else:
            st.dataframe(df, use_container_width=True, height=height)


def plot_bar(df: pd.DataFrame, x: str, y: str, title: str, color: str | None = None, text: str | None = None) -> None:
    if df.empty or x not in df.columns or y not in df.columns:
        st.info(f"No hay datos suficientes para: {title}")
        return
    d = df.copy()
    d[y] = to_num(d[y])
    fig = px.bar(d, x=x, y=y, color=color if color in d.columns else None, text=text if text in d.columns else None, title=title)
    fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=420)
    st.plotly_chart(fig, use_container_width=True)


def plot_line(df: pd.DataFrame, x: str, y: str, title: str, color: str | None = None) -> None:
    if df.empty or x not in df.columns or y not in df.columns:
        st.info(f"No hay datos suficientes para: {title}")
        return
    d = df.copy()
    d[y] = to_num(d[y])
    fig = px.line(d, x=x, y=y, color=color if color in d.columns else None, markers=True, title=title)
    fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=420)
    st.plotly_chart(fig, use_container_width=True)


def plot_pie(df: pd.DataFrame, names: str, values: str | None, title: str) -> None:
    if df.empty or names not in df.columns:
        st.info(f"No hay datos suficientes para: {title}")
        return
    d = df.copy()
    if values and values in d.columns:
        d[values] = to_num(d[values])
        fig = px.pie(d, names=names, values=values, title=title)
    else:
        counts = d[names].value_counts(dropna=False).reset_index()
        counts.columns = [names, "cantidad"]
        fig = px.pie(counts, names=names, values="cantidad", title=title)
    fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=420)
    st.plotly_chart(fig, use_container_width=True)


def numeric_sum(df: pd.DataFrame, colname: str) -> float:
    c = best_col(df, [colname])
    if not c:
        return np.nan
    return float(to_num(df[c]).sum())


def count_contains(df: pd.DataFrame, col: str | None, text: str) -> int:
    if df.empty or not col:
        return 0
    return int(df[col].astype(str).str.lower().str.contains(text.lower(), na=False).sum())

# -----------------------------------------------------------------------------
# Carga de datos por módulo
# -----------------------------------------------------------------------------

def presupuesto_data() -> dict[str, pd.DataFrame]:
    return {
        "detalle": load("presupuesto", "Detalle_Presupuesto", 2),
        "anual": load("presupuesto", "Resumen_Anual", 2),
    }


def formacion_data() -> dict[str, pd.DataFrame]:
    return {
        "registro": prepare_dates(load("formacion", "Registro", 0)),
        "anual": load("formacion", "Resumen_Anual", 0),
        "sedes": load("formacion", "Resumen_Sedes", 0),
        "mensual": load("formacion", "Resumen_Mensual", 0),
    }


def autodiagnostico_data() -> dict[str, pd.DataFrame]:
    return {
        "resumen": load("autodiagnostico", "Resumen_Comparativo", 1),
        "detalle": load_many("autodiagnostico", ["Auto_2023", "Auto_2024", "Auto_2025"], 3),
    }


def plan_anual_data() -> dict[str, pd.DataFrame]:
    return {
        "anual": load("plan_anual", "Resumen_Anual", 1),
        "mensual": load("plan_anual", "Resumen_Mensual", 1),
        "detalle": load_many("plan_anual", ["Plan_2023", "Plan_2024", "Plan_2025", "Plan_2026"], 4),
    }


def legal_data() -> dict[str, pd.DataFrame]:
    return {
        "matriz": load("legal", "Matriz_Legal_SGRL", 2),
        "resumen": load("legal", "Resumen", 1),
    }


def reporte_laboral_data() -> dict[str, pd.DataFrame]:
    return {
        "respuestas": prepare_dates(load("reporte_laboral", "Respuestas del formulario", 0)),
        "gestion": prepare_dates(load("reporte_laboral", "Gestión SST", 0)),
    }


def proveedores_data() -> dict[str, pd.DataFrame]:
    return {
        "detalle": prepare_dates(load_many("proveedores", ["GPC_2023", "GPC_2024", "GPC_2025", "GPC_2026"], 1)),
        "anual": load("proveedores", "Resumen_Anual", 1),
    }


def emo_data() -> dict[str, pd.DataFrame]:
    return {
        "detalle": prepare_dates(load("emo", "Consolidado", 0)),
        "anual": load("emo", "Resumen_Anual", 2),
        "sedes": load("emo", "Resumen_Sedes", 2),
    }


def accidentalidad_data() -> dict[str, pd.DataFrame]:
    return {
        "detalle": prepare_dates(load("accidentalidad", "Consolidado", 0)),
        "anual": load("accidentalidad", "Resumen_Anual", 1),
        "sedes": load("accidentalidad", "Resumen_Sedes", 1),
    }


def enfermedad_data() -> dict[str, pd.DataFrame]:
    return {
        "detalle": prepare_dates(load("enfermedad", "Consolidado", 1)),
        "anual": load("enfermedad", "Resumen_Anual", 1),
        "sedes": load("enfermedad", "Resumen_Sedes", 1),
    }


def indicadores_data() -> dict[str, pd.DataFrame]:
    return {
        "ficha": load("indicadores", "Ficha_Tecnica", 1),
        "medicion": prepare_dates(load("indicadores", "Consolidado", 2)),
        "anual": load("indicadores", "Resumen_Anual", 2),
    }


def peligros_data() -> dict[str, pd.DataFrame]:
    return {
        "matriz": load("peligros", "Matriz_General", 0),
        "resumen": load("peligros", "Resumen", 1),
    }


def apcm_data() -> dict[str, pd.DataFrame]:
    return {
        "detalle": prepare_dates(load("apcm", "Consolidado", 2)),
        "anual": load("apcm", "Resumen_Anual", 2),
        "sedes": load("apcm", "Resumen_Sedes", 2),
    }


def mantenimiento_data() -> dict[str, pd.DataFrame]:
    if not module_available("mantenimiento"):
        return {"detalle": pd.DataFrame(), "anual": pd.DataFrame(), "sedes": pd.DataFrame()}
    return {
        # En este archivo, Consolidado tiene título en fila 1 y encabezados en fila 2.
        "detalle": prepare_dates(load("mantenimiento", "Consolidado", 1)),
        # Resumen_Anual y Resumen_Sedes tienen título + línea en blanco antes de encabezados.
        "anual": load("mantenimiento", "Resumen_Anual", 2),
        "sedes": load("mantenimiento", "Resumen_Sedes", 2),
    }

# -----------------------------------------------------------------------------
# Páginas del dashboard
# -----------------------------------------------------------------------------

def render_resumen() -> None:
    st.title("📊 Dashboard Integral SG-SST")
    st.caption("Modelo de transformación digital y analítica de datos aplicado al SG-SST de IMA Company SAS.")

    p = presupuesto_data(); f = formacion_data(); a = autodiagnostico_data(); pa = plan_anual_data()
    l = legal_data(); r = reporte_laboral_data(); pr = proveedores_data(); e = emo_data()
    at = accidentalidad_data(); el = enfermedad_data(); ind = indicadores_data(); pg = peligros_data(); mt = mantenimiento_data(); ap = apcm_data()

    selected_year = st.sidebar.selectbox("Año de análisis", available_years(p["anual"], f["anual"], pa["anual"], pr["anual"], e["anual"], at["anual"], el["anual"], ind["anual"], mt["anual"]), index=None, placeholder="Seleccione año", key="resumen_year")
    if selected_year is None:
        selected_year = 2026

    section(f"Indicadores clave consolidados — {selected_year}")
    cols = st.columns(4)
    pres_row = latest_row_by_year(p["anual"], selected_year)
    form_row = latest_row_by_year(f["anual"], selected_year)
    plan_row = latest_row_by_year(pa["anual"], selected_year)
    auto_row = latest_row_by_year(a["resumen"], 2025 if selected_year >= 2025 else selected_year)
    with cols[0]:
        kpi("Ejecución presupuestal", pres_row.get("ejecucion_registrada", pres_row.get("ejecucion", np.nan)), "Presupuesto ejecutado", money=True)
    with cols[1]:
        kpi("Cumplimiento plan PHVA", plan_row.get("cumplimiento_global", plan_row.get("_cumplimiento_global", np.nan)), "Plan anual a corte", as_percent=True)
    with cols[2]:
        kpi("Cobertura formación", form_row.get("cobertura_promedio", form_row.get("asistencia_acumulada", np.nan)), "Promedio del año", as_percent=True)
    with cols[3]:
        kpi("Autodiagnóstico Res. 0312", auto_row.get("cumplimiento", auto_row.get("_cumplimiento", np.nan)), "Última evaluación disponible", as_percent=True)

    cols = st.columns(4)
    prov_row = latest_row_by_year(pr["anual"], selected_year)
    emo_row = latest_row_by_year(e["anual"], selected_year)
    at_row = latest_row_by_year(at["anual"], selected_year)
    el_row = latest_row_by_year(el["anual"], selected_year)
    with cols[0]:
        kpi("Cumplimiento proveedores", prov_row.get("cumplimiento_promedio", np.nan), "Promedio contratistas", as_percent=True)
    with cols[1]:
        kpi("EMO vencidos", emo_row.get("vencido", emo_row.get("vencidos", np.nan)), "Exámenes vencidos")
    with cols[2]:
        kpi("Accidentes de trabajo", at_row.get("n_at", at_row.get("at", np.nan)), "Casos del periodo")
    with cols[3]:
        kpi("Casos enfermedad laboral", el_row.get("n_casos", el_row.get("nº_casos", np.nan)), "Casos registrados")

    section("Mapa visual de avance por instrumento")
    cards = []
    cards.append({"Instrumento": "Presupuesto SST", "Resultado": pres_row.get("_ejecucion", pres_row.get("ejecucion", pres_row.get("ejecucion_registrada", np.nan))), "Tipo": "Valor"})
    cards.append({"Instrumento": "Plan Formación", "Resultado": form_row.get("efectividad", form_row.get("_efectividad", np.nan)), "Tipo": "%"})
    cards.append({"Instrumento": "Autodiagnóstico", "Resultado": auto_row.get("cumplimiento", auto_row.get("_cumplimiento", np.nan)), "Tipo": "%"})
    cards.append({"Instrumento": "Plan PHVA", "Resultado": plan_row.get("cumplimiento_global", np.nan), "Tipo": "%"})
    cards.append({"Instrumento": "Proveedores", "Resultado": prov_row.get("cumplimiento_promedio", np.nan), "Tipo": "%"})
    cards.append({"Instrumento": "Mantenimiento", "Resultado": latest_row_by_year(mt["anual"], selected_year).get("ejecucion", np.nan), "Tipo": "%"})
    cards.append({"Instrumento": "Indicadores", "Resultado": latest_row_by_year(ind["anual"], selected_year).get("resultado_anual", np.nan), "Tipo": "%"})
    cards_df = pd.DataFrame(cards)
    if not cards_df.empty:
        cards_df["Resultado_num"] = to_num(cards_df["Resultado"])
        # Si parece porcentaje decimal, escalar.
        cards_df.loc[(cards_df["Tipo"] == "%") & (cards_df["Resultado_num"] <= 1.5), "Resultado_num"] *= 100
        plot_bar(cards_df, "Instrumento", "Resultado_num", "Resultado consolidado por instrumento")

    section("Alertas operativas rápidas")
    alertas = []
    if not e["anual"].empty:
        alertas.append(["EMO", "Vencidos", fmt_value(emo_row.get("vencido", emo_row.get("vencidos", 0)))])
    if not at["anual"].empty:
        alertas.append(["Accidentalidad", "Días perdidos", fmt_value(at_row.get("dias_perdidos", 0))])
    if not ap["detalle"].empty:
        estado = best_col(ap["detalle"], ["estado"])
        alertas.append(["APCM", "Acciones vencidas", count_contains(filter_year(ap["detalle"], selected_year), estado, "venc")])
    if not r["gestion"].empty:
        estado = best_col(r["gestion"], ["estado_gestion", "estado"])
        alertas.append(["Reportes laborales", "Abiertos / en gestión", count_contains(r["gestion"], estado, "abiert") + count_contains(r["gestion"], estado, "gest")])
    if not mt["detalle"].empty:
        estado = best_col(mt["detalle"], ["estado"])
        detalle_mt = filter_year(mt["detalle"], selected_year)
        alertas.append(["Mantenimiento", "Pendientes / vencidos", count_contains(detalle_mt, estado, "pend") + count_contains(detalle_mt, estado, "venc")])
    st.dataframe(pd.DataFrame(alertas, columns=["Módulo", "Alerta", "Valor"]), use_container_width=True, hide_index=True)


def render_presupuesto() -> None:
    st.title("💰 Presupuesto SST")
    data = presupuesto_data(); detalle, anual = data["detalle"], data["anual"]
    year = st.sidebar.selectbox("Año", available_years(detalle, anual), key="pres_year")
    d = filter_year(detalle, year); a = latest_row_by_year(anual, year)
    cols = st.columns(4)
    with cols[0]: kpi("Presupuesto planeado", a.get("presupuesto_planeado", numeric_sum(d, "valor_presupuestado")), money=True)
    with cols[1]: kpi("Ejecución registrada", a.get("ejecucion_registrada", numeric_sum(d, "valor_ejecutado")), money=True)
    with cols[2]: kpi("% ejecución", a.get("ejecucion", a.get("_ejecucion", np.nan)), as_percent=True)
    with cols[3]: kpi("Variación", a.get("variacion", np.nan), money=True)
    c1, c2 = st.columns(2)
    with c1: plot_bar(anual, "ano", "presupuesto_planeado", "Presupuesto planeado por año")
    with c2:
        comp = best_col(d, ["componente_sg_sst", "programa"])
        eje = best_col(d, ["ejecucion_registrada", "valor_ejecutado", "total_ejecutado"])
        if comp and eje:
            grp = d.groupby(comp, dropna=False)[eje].apply(lambda x: to_num(x).sum()).reset_index()
            plot_bar(grp, comp, eje, "Ejecución por componente")
    show_table(d, "Detalle presupuesto")


def render_formacion() -> None:
    st.title("🎓 Plan de formación y capacitaciones")
    data = formacion_data(); registro, anual, sedes, mensual = data["registro"], data["anual"], data["sedes"], data["mensual"]
    year = st.sidebar.selectbox("Año", available_years(registro, anual), key="form_year")
    sede = st.sidebar.selectbox("Sede", sede_options(registro), key="form_sede")
    r = filter_sede(filter_year(registro, year), sede)
    a = latest_row_by_year(anual, year)
    cols = st.columns(4)
    with cols[0]: kpi("Actividades planeadas", a.get("actividades_planeadas", len(r)))
    with cols[1]: kpi("Ejecutadas", a.get("ejecutadas", np.nan))
    with cols[2]: kpi("Cobertura promedio", a.get("cobertura_promedio", np.nan), as_percent=True)
    with cols[3]: kpi("% efectividad", a.get("efectividad", a.get("_efectividad", np.nan)), as_percent=True)
    c1, c2 = st.columns(2)
    with c1:
        m = filter_year(mensual, year)
        plot_line(m, "mes", "ejecutadas", "Ejecución mensual de capacitaciones")
    with c2:
        s = filter_year(sedes, year)
        plot_bar(s, "sede", "cobertura_promedio", "Cobertura por sede")
    c1, c2 = st.columns(2)
    with c1:
        plot_pie(r, best_col(r, ["categoria"]) or "categoria", None, "Distribución por categoría")
    with c2:
        plot_pie(r, best_col(r, ["fue_efectiva", "efectividad"]) or "fue_efectiva", None, "Efectividad registrada")
    show_table(r, "Registro de formación")


def render_autodiagnostico() -> None:
    st.title("✅ Autodiagnóstico Resolución 0312 de 2019")
    data = autodiagnostico_data(); resumen, detalle = data["resumen"], data["detalle"]
    year = st.sidebar.selectbox("Año", available_years(resumen, detalle), key="auto_year")
    d = filter_year(detalle, year); a = latest_row_by_year(resumen, year)
    cols = st.columns(4)
    with cols[0]: kpi("Puntaje obtenido", a.get("puntaje_obtenido", np.nan))
    with cols[1]: kpi("% cumplimiento", a.get("cumplimiento", a.get("_cumplimiento", np.nan)), as_percent=True)
    with cols[2]: kpi("Estándares cumplen", a.get("cumple", np.nan))
    with cols[3]: kpi("Planes de mejora", a.get("planes_de_mejora_requeridos", np.nan))
    c1, c2 = st.columns(2)
    with c1: plot_line(resumen, "ano", "cumplimiento", "% cumplimiento por año")
    with c2:
        ciclo = best_col(d, ["ciclo_phva"]); puntaje = best_col(d, ["puntaje_obtenido"])
        if ciclo and puntaje:
            grp = d.groupby(ciclo, dropna=False)[puntaje].apply(lambda x: to_num(x).sum()).reset_index()
            plot_bar(grp, ciclo, puntaje, "Puntaje por ciclo PHVA")
    show_table(d, "Detalle estándares evaluados")


def render_plan_anual() -> None:
    st.title("🗓️ Plan anual de trabajo SG-SST | Ciclo PHVA")
    data = plan_anual_data(); anual, mensual, detalle = data["anual"], data["mensual"], data["detalle"]
    year = st.sidebar.selectbox("Año", available_years(anual, mensual, detalle), key="plan_year")
    sede = st.sidebar.selectbox("Sede/cobertura", sede_options(detalle), key="plan_sede")
    d = filter_sede(filter_year(detalle, year), sede); a = latest_row_by_year(anual, year)
    cols = st.columns(4)
    with cols[0]: kpi("Programadas a corte", a.get("programadas_a_corte", np.nan))
    with cols[1]: kpi("Ejecutadas a corte", a.get("ejecutadas_a_corte", np.nan))
    with cols[2]: kpi("% cumplimiento global", a.get("cumplimiento_global", np.nan), as_percent=True)
    with cols[3]: kpi("Presupuesto estimado", a.get("presupuesto_estimado_cop", np.nan), money=True)
    c1, c2 = st.columns(2)
    with c1:
        m = filter_year(mensual, year)
        plot_line(m, "mes", "cumplimiento", "Cumplimiento mensual")
    with c2:
        ciclo = best_col(d, ["ciclo_phva"])
        if ciclo:
            plot_pie(d, ciclo, None, "Actividades por ciclo PHVA")
    show_table(d, "Detalle plan anual")


def render_legal() -> None:
    st.title("⚖️ Matriz legal SGRL - SST Colombia")
    data = legal_data(); matriz = data["matriz"]
    if matriz.empty:
        st.warning("No se encontró información de la matriz legal.")
        return
    aplica = best_col(matriz, ["aplica_a_la_empresa", "aplica"])
    cumple = best_col(matriz, ["cumple"])
    tipo = best_col(matriz, ["tipo_documento"])
    tema = best_col(matriz, ["tema_eje", "tema"])
    cols = st.columns(4)
    with cols[0]: kpi("Requisitos registrados", len(matriz))
    with cols[1]: kpi("Aplicables", count_contains(matriz, aplica, "sí") + count_contains(matriz, aplica, "si"))
    with cols[2]: kpi("Cumplen", count_contains(matriz, cumple, "sí") + count_contains(matriz, cumple, "si"))
    with cols[3]: kpi("Parciales / no cumplen", count_contains(matriz, cumple, "parcial") + count_contains(matriz, cumple, "no"))
    c1, c2 = st.columns(2)
    with c1:
        if cumple: plot_pie(matriz, cumple, None, "Estado de cumplimiento legal")
    with c2:
        if tipo: plot_pie(matriz, tipo, None, "Distribución por tipo de norma")
    if tema:
        top = matriz[tema].value_counts().head(12).reset_index()
        top.columns = [tema, "cantidad"]
        plot_bar(top, tema, "cantidad", "Principales temas normativos")
    show_table(matriz, "Matriz legal completa")


def render_reporte_laboral() -> None:
    st.title("📣 Reporte laboral de autogestión")
    data = reporte_laboral_data(); respuestas, gestion = data["respuestas"], data["gestion"]
    base = gestion if not gestion.empty else respuestas
    sede = st.sidebar.selectbox("Sede", sede_options(base), key="rep_sede")
    base = filter_sede(base, sede)
    tipo = best_col(base, ["tipo_reporte", "10_tipo_de_reporte"])
    riesgo = best_col(base, ["riesgo_inmediato", "existe_riesgo_inmediato_para_usted_u_otras_personas"])
    estado = best_col(base, ["estado", "estado_gestion"])
    cols = st.columns(4)
    with cols[0]: kpi("Total reportes", len(base))
    with cols[1]: kpi("Riesgo inmediato", count_contains(base, riesgo, "sí") + count_contains(base, riesgo, "si"))
    with cols[2]: kpi("Abiertos/en gestión", count_contains(base, estado, "abiert") + count_contains(base, estado, "gest"))
    with cols[3]: kpi("Cerrados", count_contains(base, estado, "cerr"))
    c1, c2 = st.columns(2)
    with c1:
        if tipo: plot_pie(base, tipo, None, "Reportes por tipo")
    with c2:
        if estado: plot_pie(base, estado, None, "Estado de gestión")
    if not respuestas.empty:
        show_table(respuestas, "Respuestas recibidas")
    show_table(gestion, "Gestión SST")


def render_proveedores() -> None:
    st.title("🤝 Gestión de proveedores y contratistas")
    data = proveedores_data(); detalle, anual = data["detalle"], data["anual"]
    year = st.sidebar.selectbox("Año", available_years(detalle, anual), key="prov_year")
    sede = st.sidebar.selectbox("Sede", sede_options(detalle), key="prov_sede")
    d = filter_sede(filter_year(detalle, year), sede); a = latest_row_by_year(anual, year)
    cols = st.columns(4)
    with cols[0]: kpi("N° proveedores", a.get("n_proveedores", len(d)))
    with cols[1]: kpi("Cumplimiento promedio", a.get("cumplimiento_promedio", np.nan), as_percent=True)
    with cols[2]: kpi("Aprobados", a.get("aprobados", count_contains(d, best_col(d, ["estado", "cumple"]), "aprob")))
    with cols[3]: kpi("Riesgo alto", a.get("riesgo_alto", count_contains(d, best_col(d, ["nivel_de_riesgo"]), "alto")))
    c1, c2 = st.columns(2)
    with c1: plot_line(anual, "ano", "cumplimiento_promedio", "Cumplimiento promedio anual")
    with c2:
        riesgo = best_col(d, ["nivel_de_riesgo"])
        if riesgo: plot_pie(d, riesgo, None, "Proveedores por nivel de riesgo")
    show_table(d, "Detalle proveedores y contratistas")


def render_emo() -> None:
    st.title("🩺 Exámenes médicos ocupacionales")
    data = emo_data(); detalle, anual, sedes = data["detalle"], data["anual"], data["sedes"]
    year = st.sidebar.selectbox("Año", available_years(detalle, anual), key="emo_year")
    sede = st.sidebar.selectbox("Sede", sede_options(detalle), key="emo_sede")
    d = filter_sede(filter_year(detalle, year), sede); a = latest_row_by_year(anual, year)
    cols = st.columns(4)
    with cols[0]: kpi("Total exámenes", a.get("total_examenes", len(d)))
    with cols[1]: kpi("Vigentes", a.get("vigente", a.get("vigentes", np.nan)))
    with cols[2]: kpi("Próximos a vencer", a.get("proximo_a_vencer", a.get("prox_a_vencer", np.nan)))
    with cols[3]: kpi("Vencidos", a.get("vencido", a.get("vencidos", np.nan)))
    c1, c2 = st.columns(2)
    with c1:
        res = best_col(d, ["resultado"])
        if res: plot_pie(d, res, None, "Distribución de resultados EMO")
    with c2:
        s = filter_year(sedes, year) if best_col(sedes, ["ano"]) else sedes
        plot_bar(s, "sede", "vencidos", "Exámenes vencidos por sede")
    show_table(d, "Detalle EMO")


def render_accidentalidad() -> None:
    st.title("🚨 Accidentalidad")
    data = accidentalidad_data(); detalle, anual, sedes = data["detalle"], data["anual"], data["sedes"]
    year = st.sidebar.selectbox("Año", available_years(detalle, anual), key="at_year")
    sede = st.sidebar.selectbox("Sede", sede_options(detalle), key="at_sede")
    d = filter_sede(filter_year(detalle, year), sede); a = latest_row_by_year(anual, year)
    cols = st.columns(5)
    with cols[0]: kpi("N° AT", a.get("n_at", a.get("at", len(d))))
    with cols[1]: kpi("Días perdidos", a.get("dias_perdidos", np.nan))
    with cols[2]: kpi("IF", a.get("if", np.nan))
    with cols[3]: kpi("IS", a.get("is", np.nan))
    with cols[4]: kpi("ILI", a.get("ili", np.nan))
    c1, c2 = st.columns(2)
    with c1: plot_line(anual, "ano", "n_at", "Accidentes por año")
    with c2:
        mecanismo = best_col(d, ["mecanismo_del_accidente"])
        if mecanismo: plot_pie(d, mecanismo, None, "Mecanismo del accidente")
    plot_bar(sedes, "sede", "n_at", "AT por sede")
    show_table(d, "Detalle accidentalidad")


def render_enfermedad() -> None:
    st.title("🧬 Enfermedad laboral")
    data = enfermedad_data(); detalle, anual, sedes = data["detalle"], data["anual"], data["sedes"]
    year = st.sidebar.selectbox("Año", available_years(detalle, anual), key="el_year")
    sede = st.sidebar.selectbox("Sede", sede_options(detalle), key="el_sede")
    d = filter_sede(filter_year(detalle, year), sede); a = latest_row_by_year(anual, year)
    cols = st.columns(4)
    with cols[0]: kpi("N° casos", a.get("n_casos", len(d)))
    with cols[1]: kpi("Casos laborales", a.get("casos_laborales", np.nan))
    with cols[2]: kpi("Casos cerrados", a.get("casos_cerrados", np.nan))
    with cols[3]: kpi("Prom. días gestión", a.get("promedio_dias_gestion", np.nan))
    c1, c2 = st.columns(2)
    with c1:
        factor = best_col(d, ["factor_riesgo_asociado", "factores_riesgo"])
        if factor: plot_pie(d, factor, None, "Casos por factor de riesgo")
    with c2:
        res = best_col(d, ["resultado_calificacion", "resultado_calificación"])
        if res: plot_pie(d, res, None, "Resultado de calificación")
    plot_bar(sedes, "sede", "n_casos", "Casos por sede")
    show_table(d, "Detalle enfermedad laboral")


def render_indicadores() -> None:
    st.title("📈 Indicadores de gestión del SG-SST")
    data = indicadores_data(); ficha, medicion, anual = data["ficha"], data["medicion"], data["anual"]
    year = st.sidebar.selectbox("Año", available_years(medicion, anual), key="ind_year")
    d = filter_year(medicion, year); a = filter_year(anual, year)
    estado = best_col(d, ["estado"])
    tipo = best_col(d, ["tipo"])
    resultado = best_col(d, ["resultado"])
    cols = st.columns(4)
    with cols[0]: kpi("Mediciones", len(d))
    with cols[1]: kpi("En meta", count_contains(d, estado, "meta"))
    with cols[2]: kpi("Alerta", count_contains(d, estado, "alerta"))
    with cols[3]: kpi("Crítico", count_contains(d, estado, "crit"))
    c1, c2 = st.columns(2)
    with c1:
        if estado: plot_pie(d, estado, None, "Estado de indicadores")
    with c2:
        if tipo: plot_pie(d, tipo, None, "Indicadores por tipo")
    if resultado and best_col(d, ["mes"]):
        line = d.groupby(["mes"], dropna=False)[resultado].apply(lambda x: to_num(x).mean()).reset_index()
        plot_line(line, "mes", resultado, "Resultado promedio mensual")
    show_table(a, "Resumen anual de indicadores")
    show_table(ficha, "Ficha técnica de indicadores")
    show_table(d, "Mediciones del periodo")


def render_peligros() -> None:
    st.title("⚠️ Matriz de peligros y riesgos GTC-45")
    data = peligros_data(); matriz = data["matriz"]
    sede = st.sidebar.selectbox("Sede", sede_options(matriz), key="pel_sede")
    d = filter_sede(matriz, sede)
    nivel = best_col(d, ["interpretacion_nr", "nivel_de_riesgo", "nivel"])
    clasif = best_col(d, ["clasificacion_del_peligro"])
    acept = best_col(d, ["aceptabilidad_del_riesgo", "aceptabilidad"])
    estado = best_col(d, ["estado_de_implementacion", "estado"])
    cols = st.columns(4)
    with cols[0]: kpi("Riesgos identificados", len(d))
    with cols[1]: kpi("No aceptables", count_contains(d, acept, "no"))
    with cols[2]: kpi("Nivel I/II", count_contains(d, nivel, "i") + count_contains(d, nivel, "ii"))
    with cols[3]: kpi("Implementados", count_contains(d, estado, "implement"))
    c1, c2 = st.columns(2)
    with c1:
        if nivel: plot_pie(d, nivel, None, "Distribución por nivel de riesgo")
    with c2:
        if clasif: plot_pie(d, clasif, None, "Clasificación del peligro")
    show_table(d, "Detalle matriz GTC-45")


def render_mantenimiento() -> None:
    st.title("🛠️ Mantenimiento preventivo")
    if not module_available("mantenimiento"):
        st.warning("No se encontró el archivo `Matriz_Mantenimiento_Preventivo_Vigilancia_2023_2026.xlsx` en la carpeta data. Cuando lo agregues, esta página quedará lista para conectarse.")
        st.code("data/Matriz_Mantenimiento_Preventivo_Vigilancia_2023_2026.xlsx")
        return

    data = mantenimiento_data(); detalle, anual, sedes = data["detalle"], data["anual"], data["sedes"]
    year = st.sidebar.selectbox("Año", available_years(detalle, anual), key="mant_year")
    sede = st.sidebar.selectbox("Sede", sede_options(detalle), key="mant_sede")
    d = filter_sede(filter_year(detalle, year), sede)
    a = latest_row_by_year(anual, year)
    s = sedes if year == 2026 else pd.DataFrame()

    estado = best_col(d, ["estado"])
    cumplimiento = best_col(d, ["cumplimiento_oportuno"])
    tipo_activo = best_col(d, ["tipo_de_activo", "tipo_activo"])
    sede_col = best_col(d, ["sede"])
    costo_estimado = best_col(d, ["costo_estimado"])
    costo_real = best_col(d, ["costo_real"])

    cols = st.columns(5)
    with cols[0]: kpi("Total programado", a.get("total_programado", len(d)))
    with cols[1]: kpi("Ejecutados", a.get("ejecutados", count_contains(d, estado, "ejecut")))
    with cols[2]: kpi("Pendientes", a.get("pendientes", count_contains(d, estado, "pend")))
    with cols[3]: kpi("% ejecución", a.get("ejecucion", np.nan), as_percent=True)
    with cols[4]: kpi("Costo real", a.get("costo_real", numeric_sum(d, "costo_real")), money=True)

    cols = st.columns(4)
    with cols[0]: kpi("Reprogramados", a.get("reprogramados", count_contains(d, estado, "reprogram")))
    with cols[1]: kpi("Cancelados", a.get("cancelados", count_contains(d, estado, "cancel")))
    with cols[2]: kpi("Cumplimiento oportuno", a.get("cumplimiento_oportuno", np.nan), as_percent=True)
    with cols[3]: kpi("Costo estimado", a.get("costo_estimado", numeric_sum(d, "costo_estimado")), money=True)

    c1, c2 = st.columns(2)
    with c1:
        if estado: plot_pie(d, estado, None, "Estado de mantenimientos")
    with c2:
        if tipo_activo: plot_pie(d, tipo_activo, None, "Mantenimientos por tipo de activo")

    c1, c2 = st.columns(2)
    with c1:
        if not anual.empty:
            plot_line(anual, "ano", "ejecucion", "% ejecución anual")
    with c2:
        if not anual.empty:
            plot_bar(anual, "ano", "costo_real", "Costo real anual")

    if sede_col and costo_real:
        section("Costo real por sede")
        grp = d.groupby(sede_col, dropna=False)[costo_real].apply(lambda x: to_num(x).sum()).reset_index()
        plot_bar(grp, sede_col, costo_real, "Costo real de mantenimiento por sede")

    if not s.empty:
        show_table(s, "Resumen por sede 2026")
    show_table(d, "Detalle mantenimiento")


def render_apcm() -> None:
    st.title("🔁 Acciones preventivas, correctivas y de mejora")
    data = apcm_data(); detalle, anual, sedes = data["detalle"], data["anual"], data["sedes"]
    year = st.sidebar.selectbox("Año", available_years(detalle), key="apcm_year")
    sede = st.sidebar.selectbox("Sede", sede_options(detalle), key="apcm_sede")
    d = filter_sede(filter_year(detalle, year), sede)
    estado = best_col(d, ["estado"])
    tipo = best_col(d, ["tipo"])
    fuente = best_col(d, ["fuente_de_origen"])
    eficacia = best_col(d, ["eficacia_verificada"])
    cols = st.columns(5)
    with cols[0]: kpi("Acciones", len(d))
    with cols[1]: kpi("Cerradas", count_contains(d, estado, "cerr"))
    with cols[2]: kpi("Vencidas", count_contains(d, estado, "venc"))
    with cols[3]: kpi("En progreso", count_contains(d, estado, "progreso"))
    with cols[4]: kpi("Eficaces", count_contains(d, eficacia, "sí") + count_contains(d, eficacia, "si"))
    c1, c2 = st.columns(2)
    with c1:
        if tipo: plot_pie(d, tipo, None, "Acciones por tipo")
    with c2:
        if fuente: plot_pie(d, fuente, None, "Acciones por fuente de origen")
    show_table(d, "Detalle APCM")

# -----------------------------------------------------------------------------
# Navegación
# -----------------------------------------------------------------------------

st.sidebar.title("IMA Company SAS")
st.sidebar.caption("Ecosistema digital SG-SST")

missing = [MODULE_LABELS[k] for k in FILES if not module_available(k)]
if missing:
    st.sidebar.warning("Archivos no encontrados: " + ", ".join(missing))

page_keys = [
    "resumen", "presupuesto", "formacion", "autodiagnostico", "plan_anual", "legal",
    "reporte_laboral", "proveedores", "emo", "accidentalidad", "enfermedad", "indicadores",
    "peligros", "mantenimiento", "apcm",
]
page = st.sidebar.radio("Seleccione dashboard", page_keys, format_func=lambda k: MODULE_LABELS[k])

st.sidebar.divider()
st.sidebar.caption("Fuente de datos: archivos Excel ubicados en /data. Los filtros se aplican en memoria con pandas.")

ROUTERS = {
    "resumen": render_resumen,
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
