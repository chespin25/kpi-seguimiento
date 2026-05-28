import os
import re
import openpyxl
from config import FICHAS_DIR, FICHA_SHEET, FICHA_COL_N


_CODE_RE = re.compile(r"^(\d{7})")


def _extract_code(filename: str) -> str | None:
    m = _CODE_RE.match(os.path.basename(filename))
    return m.group(1) if m else None


def _count_objectives(filepath: str) -> int | None:
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        if FICHA_SHEET not in wb.sheetnames:
            return None
        ws = wb[FICHA_SHEET]

        # Encontrar columna "Nª" en las primeras 20 filas
        col_idx = None
        header_row = None
        for row in ws.iter_rows(max_row=20, values_only=True):
            for i, cell in enumerate(row):
                if str(cell).strip() == FICHA_COL_N:
                    col_idx = i
                    header_row = row
                    break
            if col_idx is not None:
                break

        if col_idx is None:
            return None

        count = 0
        for row in ws.iter_rows(min_row=(ws.min_row or 1), values_only=True):
            if row is header_row:
                continue
            val = row[col_idx] if col_idx < len(row) else None
            if val is not None:
                try:
                    int(float(str(val)))
                    count += 1
                except (ValueError, TypeError):
                    pass
        return count if count > 0 else None
    except Exception:
        return None


def scan_fichas(periodo: str, base_dir: str = None) -> list[dict]:
    """
    Escanea base_dir (o FICHAS_DIR/<periodo>/) y retorna lista de dicts:
    {"codigo": ..., "subio_ficha": "SI", "nro_objetivos": N o None}
    """
    folder = base_dir or os.path.join(FICHAS_DIR, periodo)
    if not os.path.isdir(folder):
        return []

    results = []
    seen = set()
    for fname in os.listdir(folder):
        ext = os.path.splitext(fname)[1].lower()
        if ext not in (".xlsx", ".xls", ".pdf"):
            continue
        codigo = _extract_code(fname)
        if not codigo or codigo in seen:
            continue
        seen.add(codigo)

        entry = {"codigo": codigo, "subio_ficha": "SI", "nro_objetivos": None}
        if ext in (".xlsx", ".xls"):
            entry["nro_objetivos"] = _count_objectives(os.path.join(folder, fname))
        results.append(entry)

    return results
