import pandas as pd
from config import PERIODO_ACTIVO
from modules.state_manager import get_all_for_periodo


OUTPUT_COLS = [
    "codigo",
    "nombre",
    "gerencia",
    "subgerencia",
    "area",
    "cargo",
    "subio_ficha",
    "comentario",
    "nro_objetivos",
    "formato_firmado",
]

DISPLAY_COLS = {
    "codigo":          "CÓDIGO",
    "nombre":          "APELLIDOS Y NOMBRES",
    "gerencia":        "GERENCIA",
    "subgerencia":     "GERENCIA / SUBGERENCIA / JEFATURA",
    "area":            "ÁREA",
    "cargo":           "CARGO",
    "subio_ficha":     "SUBIÓ FICHA\nSI / NO",
    "comentario":      "COMENTARIO",
    "nro_objetivos":   "NRO DE OBJETIVOS\nASIGNADOS",
    "formato_firmado": "FORMATO FIRMADO\nSI / NO",
}


def build_full_table(workers_df: pd.DataFrame, periodo: str = PERIODO_ACTIVO) -> pd.DataFrame:
    """
    Combina workers_df con el estado guardado (subio_ficha, nro_objetivos, formato_firmado).
    Retorna DataFrame con OUTPUT_COLS listo para mostrar/exportar.
    """
    state = get_all_for_periodo(periodo)

    df = workers_df.copy()
    df["subio_ficha"]     = df["codigo"].map(lambda c: state.get(c, {}).get("subio_ficha") or "NO")
    df["nro_objetivos"]   = df["codigo"].map(lambda c: state.get(c, {}).get("nro_objetivos"))
    df["formato_firmado"] = df["codigo"].map(lambda c: state.get(c, {}).get("formato_firmado") or "NO")
    if "area" not in df.columns:
        df["area"] = ""

    # comentario: combina el de encargo (data_loader) + comentario_extra del estado
    extra = df["codigo"].map(lambda c: state.get(c, {}).get("comentario_extra") or "")
    df["comentario"] = df["comentario"].fillna("").str.strip()
    df["comentario"] = df.apply(
        lambda r: "; ".join(filter(None, [r["comentario"], extra[r.name]])), axis=1
    )

    return (df[OUTPUT_COLS]
            .sort_values(["gerencia", "subgerencia", "nombre"])
            .reset_index(drop=True))


def build_summary(full_df: pd.DataFrame) -> pd.DataFrame:
    """
    Resumen por gerencia: FICHAS REQUERIDAS, FICHAS REGISTRADAS, % AVANCE.
    """
    rows = []
    for gerencia, gdf in full_df.groupby("gerencia"):
        total    = len(gdf)
        registradas = (gdf["subio_ficha"] == "SI").sum()
        pct      = registradas / total if total else 0
        rows.append({
            "GERENCIA":           gerencia,
            "FICHAS REQUERIDAS":  total,
            "FICHAS REGISTRADAS": int(registradas),
            "% AVANCE":           round(pct * 100, 1),
        })
    return pd.DataFrame(rows).sort_values("GERENCIA").reset_index(drop=True)


def build_global_summary(full_df: pd.DataFrame) -> pd.DataFrame:
    """
    Resumen general: Total, Subieron ficha, No subieron, % avance.
    """
    total       = len(full_df)
    si          = (full_df["subio_ficha"] == "SI").sum()
    no          = total - si
    firmados    = (full_df["formato_firmado"] == "SI").sum()
    return pd.DataFrame([
        {"CRITERIO": "FICHAS REQUERIDAS",  "FICHAS REGISTRADAS": total,      "% AVANCE": "—"},
        {"CRITERIO": "FICHAS REGISTRADAS", "FICHAS REGISTRADAS": int(si),    "% AVANCE": f"{round(si/total*100,1)}%" if total else "0%"},
        {"CRITERIO": "FICHAS PENDIENTES",  "FICHAS REGISTRADAS": int(no),    "% AVANCE": f"{round(no/total*100,1)}%" if total else "0%"},
        {"CRITERIO": "FORMATOS FIRMADOS",  "FICHAS REGISTRADAS": int(firmados), "% AVANCE": f"{round(firmados/total*100,1)}%" if total else "0%"},
    ])
