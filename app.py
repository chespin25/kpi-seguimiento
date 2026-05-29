import os
import sys
import tempfile

import io
import zipfile
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

sys.path.insert(0, os.path.dirname(__file__))

from config import PERIODO_ACTIVO, FICHAS_DIR
from modules.data_loader import upsert_workers_to_db, load_workers_from_db
from modules.ficha_parser import scan_fichas
from modules.state_manager import bulk_set
from modules.report_builder import build_full_table, build_summary, build_global_summary
from modules.export_xlsx import export_gerencia_xlsx, export_all_xlsx
from modules.export_pdf  import export_gerencia_pdf

st.set_page_config(page_title="KPI Distribución", layout="wide")

# Invalida sesión si el código cambió (evita datos cacheados con versión vieja)
_CODE_VER = "b4c3ec6"
if st.session_state.get("_ver") != _CODE_VER:
    st.session_state.clear()
    st.session_state["_ver"] = _CODE_VER

if "_planta_ok" in st.session_state:
    st.success(f"{st.session_state.pop('_planta_ok')} trabajadores actualizados correctamente.")

_LOGO = os.path.join(os.path.dirname(__file__), "assets", "logo_bn.png")
col_logo, col_title = st.columns([1, 6])
with col_logo:
    if os.path.exists(_LOGO):
        st.image(_LOGO, width=140)
with col_title:
    st.title("Seguimiento KPI — Fichas de Objetivos Individuales")

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    if os.path.exists(_LOGO):
        st.image(_LOGO, use_container_width=True)
    st.header("Configuración")
    periodo = st.selectbox("Periodo", ["2026", "2025", "2024"], index=0)

    st.divider()
    st.subheader("1. Planta de Personal")
    planta_file = st.file_uploader("Subir Planta de Personal.xlsx", type=["xlsx"])
    if planta_file and st.button("Actualizar Planta"):
        with st.spinner("Procesando..."):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(planta_file.read())
                tmp_path = tmp.name
            n = upsert_workers_to_db(tmp_path)
            os.unlink(tmp_path)
            for k in list(st.session_state.keys()):
                if k != "_ver":
                    st.session_state.pop(k, None)
            st.session_state["_planta_ok"] = n
        st.rerun()

    st.divider()
    st.subheader("2. Fichas desde OneDrive")

    # ── Subir ZIP ───────────────────────────────────────────────────────────────
    st.markdown("**Subir ZIP de OneDrive**")
    st.caption("En OneDrive: selecciona la carpeta → clic derecho → Descargar → sube el .zip aquí.")
    zip_file = st.file_uploader("Carpeta de fichas (.zip)", type=["zip"], key="fichas_zip")
    if zip_file and st.button("Procesar ZIP", type="primary"):
        with st.spinner("Extrayendo fichas..."):
            dest = os.path.join(tempfile.gettempdir(), "fichas", periodo)
            os.makedirs(dest, exist_ok=True)
            with zipfile.ZipFile(io.BytesIO(zip_file.read())) as zf:
                extraidos = 0
                for member in zf.namelist():
                    ext = os.path.splitext(member)[1].lower()
                    if ext in (".xlsx", ".xls", ".pdf"):
                        fname = os.path.basename(member)
                        if fname:
                            with zf.open(member) as src, \
                                 open(os.path.join(dest, fname), "wb") as dst:
                                dst.write(src.read())
                            extraidos += 1
            results = scan_fichas(periodo, base_dir=dest,
                                  workers_df=st.session_state.get("workers_df"))
            bulk_set(results, periodo)
            st.session_state.pop("full_df", None)
        st.success(f"{extraidos} archivos extraídos · {len(results)} procesados")

    st.divider()
    st.subheader("3. Fichas KPI (subida manual)")
    fichas_files = st.file_uploader(
        "Subir fichas (.xlsx / .pdf)", type=["xlsx", "xls", "pdf"],
        accept_multiple_files=True)
    if fichas_files and st.button("Procesar fichas"):
        with st.spinner("Procesando fichas..."):
            dest = os.path.join(tempfile.gettempdir(), "fichas", periodo)
            os.makedirs(dest, exist_ok=True)
            for f in fichas_files:
                with open(os.path.join(dest, f.name), "wb") as out:
                    out.write(f.read())
            results = scan_fichas(periodo, base_dir=dest, workers_df=st.session_state.get("workers_df"))
            bulk_set(results, periodo)
            st.session_state.pop("full_df", None)
        st.success(f"{len(results)} fichas procesadas")

# ── Cargar workers ─────────────────────────────────────────────────────────────
if ("workers_df" not in st.session_state
        or "full_df" not in st.session_state
        or st.session_state.get("periodo_loaded") != periodo):
    with st.spinner("Cargando datos..."):
        df_w = load_workers_from_db()
        if df_w.empty:
            st.info("Sube la Planta de Personal para comenzar.")
            st.stop()
        st.session_state["workers_df"] = df_w
        st.session_state["full_df"]     = build_full_table(df_w, periodo)
        st.session_state["periodo_loaded"] = periodo

workers_df = st.session_state["workers_df"]
full_df    = st.session_state["full_df"]
# Compatibilidad: columna area puede no existir en Supabase aún
if "area" not in full_df.columns:
    full_df["area"] = ""

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab_tabla, tab_reporte, tab_export = st.tabs(["Tabla por gerencia", "Reporte", "Exportar"])

# ── TAB 1 ──────────────────────────────────────────────────────────────────────
with tab_tabla:
    gerencias    = sorted(full_df["gerencia"].dropna().unique())
    gerencia_sel = st.selectbox("Gerencia", gerencias)
    df_g = full_df[full_df["gerencia"] == gerencia_sel].copy().reset_index(drop=True)

    si_count = (df_g["subio_ficha"] == "SI").sum()
    total    = len(df_g)
    c1, c2, c3 = st.columns(3)
    c1.metric("Total", total)
    c2.metric("Subieron ficha", si_count)
    c3.metric("% Avance", f"{round(si_count/total*100,1)}%" if total else "0%")

    _display_cols = ["codigo", "nombre", "gerencia", "subgerencia", "cargo",
                     "subio_ficha", "comentario", "nro_objetivos", "formato_firmado"]
    # Garantizar que todas las columnas existen
    for _c in _display_cols:
        if _c not in df_g.columns:
            df_g[_c] = "" if _c != "nro_objetivos" else None

    edited = st.data_editor(
        df_g[_display_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            "codigo":          st.column_config.TextColumn("CÓDIGO",                        disabled=True, width="small"),
            "nombre":          st.column_config.TextColumn("APELLIDOS Y NOMBRES",           disabled=True, width="large"),
            "gerencia":        st.column_config.TextColumn("GERENCIA",                      disabled=True, width="medium"),
            "subgerencia":     st.column_config.TextColumn("GERENCIA / SUBGERENCIA / JEFATURA", disabled=True, width="large"),
            "cargo":           st.column_config.TextColumn("CARGO",                         disabled=True, width="medium"),
            "subio_ficha":     st.column_config.SelectboxColumn("FICHA",                    options=["SI","NO"], width="small"),
            "comentario":      st.column_config.TextColumn("COMENTARIO",                    width="large"),
            "nro_objetivos":   st.column_config.NumberColumn("NRO OBJETIVOS ASIGNADOS",     min_value=0, step=1, width="small"),
            "formato_firmado": st.column_config.SelectboxColumn("FORMATO FIRMADO",          options=["SI","NO"], width="small"),
        },
        num_rows="fixed",
        key=f"editor_{gerencia_sel}_{periodo}",
    )

    if st.button("Guardar cambios", type="primary"):
        updates = (edited.drop(columns=["gerencia", "subgerencia", "cargo"], errors="ignore")
                        .rename(columns={"comentario": "comentario_extra"})
                        .to_dict("records"))
        bulk_set(updates, periodo)
        st.session_state["full_df"] = build_full_table(workers_df, periodo)
        st.success("Guardado")
        st.rerun()

# ── TAB 2 ──────────────────────────────────────────────────────────────────────
with tab_reporte:
    st.subheader(f"REPORTE AL {date.today().strftime('%d/%m/%Y')}")

    summary_raw = build_summary(full_df)   # % AVANCE como float

    # ── Gráfico ─────────────────────────────────────────────────────────────────
    gerencias_chart = summary_raw["GERENCIA"].tolist()
    requeridas      = summary_raw["FICHAS REQUERIDAS"].tolist()
    registradas     = summary_raw["FICHAS REGISTRADAS"].tolist()
    pct             = summary_raw["% AVANCE"].tolist()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Requeridas",
        y=gerencias_chart, x=requeridas,
        orientation="h",
        marker_color="#D0D5DD",
        text=requeridas, textposition="outside",
        textfont=dict(size=10, color="#555"),
    ))
    fig.add_trace(go.Bar(
        name="Registradas",
        y=gerencias_chart, x=registradas,
        orientation="h",
        marker_color="#003F7F",
        text=[f"{p}%" for p in pct],
        textposition="inside",
        textfont=dict(size=10, color="white"),
        insidetextanchor="middle",
    ))

    # Totales en anotación al pie
    tot_req = sum(requeridas)
    tot_reg = sum(registradas)
    tot_pct = round(tot_reg / tot_req * 100, 1) if tot_req else 0
    fig.add_annotation(
        text=f"<b>TOTAL — Requeridas: {tot_req} | Registradas: {tot_reg} | Avance: {tot_pct}%</b>",
        xref="paper", yref="paper", x=0, y=-0.08,
        showarrow=False, font=dict(size=12, color="#003F7F"), align="left",
    )

    fig.update_layout(
        barmode="overlay",
        height=max(420, len(gerencias_chart) * 38 + 120),
        margin=dict(l=10, r=90, t=20, b=60),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0),
        xaxis_title="Número de Fichas",
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(showgrid=True, gridcolor="#F0F0F0"),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig, use_container_width=True)

    # ── Tabla y resumen general ─────────────────────────────────────────────────
    st.divider()
    summary_display = summary_raw.copy()
    summary_display["% AVANCE"] = summary_display["% AVANCE"].apply(lambda x: f"{x}%")
    st.dataframe(summary_display, use_container_width=True, hide_index=True)
    st.divider()
    st.subheader("REGISTRO GENERAL")
    st.dataframe(build_global_summary(full_df), use_container_width=True, hide_index=True)

# ── TAB EXPORTAR ───────────────────────────────────────────────────────────────
with tab_export:
    st.subheader("Exportar por gerencia")
    gerencia_exp = st.selectbox("Gerencia", sorted(full_df["gerencia"].dropna().unique()),
                                key="exp_gerencia")
    c1, c2, c3 = st.columns(3)

    with c1:
        if st.button("Descargar XLSX (esta gerencia)"):
            data = export_gerencia_xlsx(full_df, gerencia_exp, periodo)
            st.download_button("Guardar XLSX", data,
                               file_name=f"KPI_{gerencia_exp}_{periodo}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        if st.button("Descargar PDF (esta gerencia)"):
            data = export_gerencia_pdf(full_df, gerencia_exp, periodo)
            st.download_button("Guardar PDF", data,
                               file_name=f"KPI_{gerencia_exp}_{periodo}.pdf",
                               mime="application/pdf")
    with c3:
        if st.button("Descargar XLSX (todas las gerencias)"):
            data = export_all_xlsx(full_df, periodo)
            st.download_button("Guardar XLSX completo", data,
                               file_name=f"KPI_COMPLETO_{periodo}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
