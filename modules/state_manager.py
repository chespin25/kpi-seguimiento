from modules.db import get_client


def get_all_for_periodo(periodo: str) -> dict:
    db = get_client()
    all_rows, offset = [], 0
    while True:
        page = (db.table("kpi_state").select("*")
                  .eq("periodo", periodo)
                  .range(offset, offset + 999)
                  .execute().data)
        if not page:
            break
        all_rows.extend(page)
        if len(page) < 1000:
            break
        offset += 1000
    return {r["codigo"]: r for r in all_rows}


def bulk_set(updates: list[dict], periodo: str):
    if not updates:
        return
    db = get_client()
    rows = []
    for u in updates:
        row = {"codigo": u["codigo"], "periodo": periodo}
        for field in ("subio_ficha", "nro_objetivos", "formato_firmado", "comentario_extra"):
            if field in u and u[field] is not None:
                row[field] = u[field]
        rows.append(row)
    db.table("kpi_state").upsert(rows, on_conflict="codigo,periodo").execute()


def set_worker_state(codigo: str, periodo: str, **kwargs):
    bulk_set([{"codigo": codigo, **kwargs}], periodo)
