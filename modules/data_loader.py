import unicodedata
import pandas as pd
from config import PLANTA_SHEET, PLANTA_ENCARGOS_SHEET, PLANTA_HEADER_ROW, PLANTA_IDX, ENC_IDX

_CORPORATE_TIPOS = {"GERENCIA", "SUBGERENCIA", "SECCION", "JEFATURA", "OFICINA", "DIVISION", "DEPARTAMENTO"}


def _format_subgerencia(tipo, oficina):
    tipo = (tipo or "").strip().upper()
    oficina = (oficina or "").strip()
    if not oficina:
        return ""
    mapping = {
        "GERENCIA":     f"GERENCIA DE {oficina}",
        "SUBGERENCIA":  f"SUBGERENCIA {oficina}",
        "SECCION":      f"SECCIÓN {oficina}",
        "JEFATURA":     f"JEFATURA {oficina}",
        "OFICINA":      f"OFICINA {oficina}",
        "DIVISION":     f"DIVISIÓN {oficina}",
        "DEPARTAMENTO": f"DEPARTAMENTO {oficina}",
    }
    return mapping.get(tipo, f"{tipo} {oficina}".strip())


def _normalize(s: str) -> str:
    return unicodedata.normalize("NFD", s.upper()).encode("ascii", "ignore").decode()


def _canonical_gerencias(series: pd.Series) -> dict:
    counts = series.value_counts().to_dict()
    names = list(counts.keys())
    canonical = {n: n for n in names}
    norm_map = {_normalize(n): n for n in names}
    for n in names:
        for prefix in ("GERENCIA DE ", "GERENCIA "):
            if n.upper().startswith(prefix):
                stripped = n[len(prefix):]
                match = norm_map.get(_normalize(stripped))
                if match:
                    canonical[n] = match
                break
    norm_to_best = {}
    for n in sorted(names, key=lambda x: -counts.get(x, 0)):
        key = _normalize(n)
        if key not in norm_to_best:
            norm_to_best[key] = n
    for n in names:
        best = norm_to_best[_normalize(n)]
        if canonical[n] == n and best != n:
            canonical[n] = best
    return canonical


def load_planta(filepath: str) -> pd.DataFrame:
    df_raw = pd.read_excel(filepath, sheet_name=PLANTA_SHEET,
                           header=PLANTA_HEADER_ROW - 1, dtype=str)
    df_raw = df_raw.dropna(how="all")
    idx = PLANTA_IDX
    df = df_raw.iloc[:, list(idx.values())].copy()
    df.columns = ["codigo", "nombre", "cargo", "gerencia", "tipo_oficina", "oficina", "area"]
    for col in df.columns:
        df[col] = df[col].fillna("").str.strip()
    df["subgerencia"] = df.apply(
        lambda r: _format_subgerencia(r["tipo_oficina"], r["oficina"]), axis=1)
    def _resolve(row):
        if row["gerencia"] == "OFICINA PRINCIPAL":
            return "OFICINA PRINCIPAL"
        if row["tipo_oficina"].upper() not in _CORPORATE_TIPOS:
            return "RED DE AGENCIAS"
        return row["gerencia"]
    df["gerencia"] = df.apply(_resolve, axis=1)
    df = df[df["codigo"].str.match(r"^\d{7}$", na=False)].reset_index(drop=True)
    return df[["codigo", "nombre", "cargo", "gerencia", "subgerencia", "area", "tipo_oficina", "oficina"]]


def load_encargos(filepath: str) -> dict:
    df_raw = pd.read_excel(filepath, sheet_name=PLANTA_ENCARGOS_SHEET,
                           header=0, dtype=str)
    df_raw = df_raw.dropna(how="all")
    idx = ENC_IDX
    result = {}
    for _, row in df_raw.iterrows():
        codigo = str(row.iloc[idx["codigo"]]).strip()
        if not codigo or codigo == "nan":
            continue
        repart_orig  = str(row.iloc[idx["reparticion_orig"]]).strip()
        oficina_orig = str(row.iloc[idx["oficina_orig"]]).strip()
        repart_enc   = str(row.iloc[idx["repart_enc"]]).strip()
        tipo_ofi_enc = str(row.iloc[idx["tipo_ofi_enc"]]).strip()
        oficina_enc  = str(row.iloc[idx["oficina_enc"]]).strip()
        subgerencia_enc = _format_subgerencia(tipo_ofi_enc, oficina_enc)
        comentario = (f"Oficina de origen: {repart_orig} - {oficina_orig}"
                      if oficina_orig else f"Gerencia de origen: {repart_orig}")
        if repart_enc == "OFICINA PRINCIPAL":
            gerencia_final = "OFICINA PRINCIPAL"
        elif _normalize(tipo_ofi_enc) not in _CORPORATE_TIPOS:
            # tipo_ofi_enc no reconocido: usar repart_enc directamente
            # salvo que sea claramente una agencia/oficina especial
            if any(repart_enc.upper().startswith(x)
                   for x in ("AGENCIA", "OFIC ESPECIAL", "OFICINA ESPECIAL")):
                gerencia_final = "RED DE AGENCIAS"
            else:
                gerencia_final = repart_enc
        else:
            gerencia_final = repart_enc
        result[codigo] = {"gerencia": gerencia_final, "subgerencia": subgerencia_enc,
                          "comentario": comentario}
    return result


def build_workers_table(filepath: str) -> pd.DataFrame:
    df = load_planta(filepath)
    encargos = load_encargos(filepath)

    def apply_encargo(row):
        enc = encargos.get(row["codigo"])
        if enc:
            return pd.Series(enc)
        return pd.Series({"gerencia": row["gerencia"],
                          "subgerencia": row["subgerencia"], "comentario": ""})

    override = df.apply(apply_encargo, axis=1)
    df["gerencia"]    = override["gerencia"]
    df["subgerencia"] = override["subgerencia"]
    df["comentario"]  = override["comentario"]
    df = df.drop(columns=["tipo_oficina", "oficina"])
    if "area" not in df.columns:
        df["area"] = ""
    canonical = _canonical_gerencias(df["gerencia"])
    df["gerencia"] = df["gerencia"].map(canonical)
    return df.sort_values(["gerencia", "subgerencia", "nombre"]).reset_index(drop=True)


def upsert_workers_to_db(filepath: str):
    """Carga Planta y sube workers a Supabase."""
    from modules.db import get_client
    df = build_workers_table(filepath)
    records = df.fillna("").to_dict("records")
    db = get_client()
    # upsert en lotes de 500
    for i in range(0, len(records), 500):
        db.table("workers").upsert(records[i:i+500], on_conflict="codigo").execute()
    return len(records)


def load_workers_from_db() -> pd.DataFrame:
    """Lee workers desde Supabase."""
    from modules.db import get_client
    db = get_client()
    rows = db.table("workers").select("*").execute().data
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df = df.drop(columns=["updated_at"], errors="ignore")
    if "area" not in df.columns:
        df["area"] = ""
    return df.sort_values(["gerencia", "subgerencia", "nombre"]).reset_index(drop=True)
