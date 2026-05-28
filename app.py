import os
import sys
import tempfile

import pandas as pd
import streamlit as st

sys.path.insert(0, os.path.dirname(__file__))

from config import PERIODO_ACTIVO, FICHAS_DIR
from modules.data_loader import upsert_workers_to_db, load_workers_from_db
from modules.ficha_parser import scan_fichas
from modules.state_manager import bulk_set
from modules.report_builder import build_full_table, build_summary, build_global_summary

st.set_page_config(page_title="KPI Distribución", layout="wide")
st.title("Seguimiento KPI — Fichas de Objetivos Individuales")

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
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
            st.session_state.pop("workers_df", None)
            st.session_state.pop("full_df", None)
        st.success(f"{n} trabajadores actualizados")

    st.divider()
    st.subheader("2. Fichas KPI")
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
            results = scan_fichas(periodo, base_dir=dest)
            bulk_set(results, periodo)
            st.session_state.pop("full_df", None)
        st.success(f"{len(results)} fichas procesadas")

# ── Cargar workers ─────────────────────────────────────────────────────────────
if "workers_df" not in st.session_state or st.session_state.get("periodo_loaded") != periodo:
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

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab_tabla, tab_reporte = st.tabs(["Tabla por gerencia", "Reporte"])

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

    edited = st.data_editor(
        df_g[["codigo", "nombre", "subgerencia", "cargo",
              "subio_ficha", "comentario", "nro_objetivos", "formato_firmado"]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "codigo":          st.column_config.TextColumn("Código",       disabled=True, width="small"),
            "nombre":          st.column_config.TextColumn("Nombres",      disabled=True, width="large"),
            "subgerencia":     st.column_config.TextColumn("Subgerencia",  disabled=True, width="medium"),
            "cargo":           st.column_config.TextColumn("Cargo",        disabled=True, width="medium"),
            "subio_ficha":     st.column_config.SelectboxColumn("Ficha",   options=["SI","NO"], width="small"),
            "comentario":      st.column_config.TextColumn("Comentario",   width="large"),
            "nro_objetivos":   st.column_config.NumberColumn("Nro. Obj",   min_value=0, step=1, width="small"),
            "formato_firmado": st.column_config.SelectboxColumn("Firmado", options=["SI","NO"], width="small"),
        },
        num_rows="fixed",
        key=f"editor_{gerencia_sel}_{periodo}",
    )

    if st.button("Guardar cambios", type="primary"):
        updates = edited.rename(columns={"comentario": "comentario_extra"}).to_dict("records")
        bulk_set(updates, periodo)
        st.session_state["full_df"] = build_full_table(workers_df, periodo)
        st.success("Guardado")
        st.rerun()

# ── TAB 2 ──────────────────────────────────────────────────────────────────────
with tab_reporte:
    from datetime import date
    st.subheader(f"REPORTE AL {date.today().strftime('%d/%m/%Y')}")
    summary = build_summary(full_df)
    summary["% AVANCE"] = summary["% AVANCE"].apply(lambda x: f"{x}%")
    st.dataframe(summary, use_container_width=True, hide_index=True)
    st.divider()
    st.subheader("REGISTRO GENERAL")
    st.dataframe(build_global_summary(full_df), use_container_width=True, hide_index=True)
