import io
from datetime import date
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer


_DARK_BLUE  = colors.HexColor("#1F3864")
_MID_BLUE   = colors.HexColor("#2E75B6")
_GREEN      = colors.HexColor("#C6EFCE")
_RED        = colors.HexColor("#FFCCCC")
_LIGHT_GRAY = colors.HexColor("#F2F2F2")


def export_gerencia_pdf(full_df: pd.DataFrame, gerencia: str, periodo: str) -> bytes:
    df = full_df[full_df["gerencia"] == gerencia].copy().reset_index(drop=True)
    si  = (df["subio_ficha"] == "SI").sum()
    pct = round(si / len(df) * 100, 1) if len(df) else 0

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=1*cm, rightMargin=1*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    small  = ParagraphStyle("small", fontSize=7, leading=9)
    header_style = ParagraphStyle("hdr", fontSize=7, leading=9,
                                  textColor=colors.white, fontName="Helvetica-Bold")

    elements = []

    # Título
    elements.append(Paragraph(
        f"<b>FICHAS DE OBJETIVOS INDIVIDUALES — {gerencia.upper()} — PERIODO {periodo}</b>",
        ParagraphStyle("title", fontSize=11, textColor=_DARK_BLUE, spaceAfter=4)))
    elements.append(Paragraph(
        f"Fecha: {date.today().strftime('%d/%m/%Y')}  |  Total: {len(df)}  |  Subieron: {int(si)}  |  Avance: {pct}%",
        ParagraphStyle("sub", fontSize=8, textColor=colors.gray, spaceAfter=8)))

    # Tabla
    col_headers = ["CÓDIGO","NOMBRES","SUBGERENCIA / JEFATURA","CARGO",
                   "FICHA","COMENTARIO","NRO\nOBJ","FIRMADO"]
    col_fields  = ["codigo","nombre","subgerencia","cargo",
                   "subio_ficha","comentario","nro_objetivos","formato_firmado"]
    col_widths  = [1.8*cm, 5.5*cm, 5.5*cm, 4*cm, 1.4*cm, 6.5*cm, 1.2*cm, 1.8*cm]

    data = [[Paragraph(h, header_style) for h in col_headers]]
    for _, row in df.iterrows():
        data_row = []
        for f in col_fields:
            val = row.get(f, "")
            val = "" if pd.isna(val) else str(val)
            data_row.append(Paragraph(val, small))
        data.append(data_row)

    table = Table(data, colWidths=col_widths, repeatRows=1)

    style_cmds = [
        ("BACKGROUND", (0,0), (-1,0), _DARK_BLUE),
        ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 7),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, _LIGHT_GRAY]),
        ("GRID",       (0,0), (-1,-1), 0.3, colors.gray),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN",      (4,0), (4,-1), "CENTER"),
        ("ALIGN",      (6,0), (7,-1), "CENTER"),
    ]
    # Color SI/NO
    for r_idx, (_, row) in enumerate(df.iterrows(), 1):
        if row.get("subio_ficha") == "SI":
            style_cmds.append(("BACKGROUND", (4,r_idx), (4,r_idx), _GREEN))
        else:
            style_cmds.append(("BACKGROUND", (4,r_idx), (4,r_idx), _RED))
        if row.get("formato_firmado") == "SI":
            style_cmds.append(("BACKGROUND", (7,r_idx), (7,r_idx), _GREEN))
        else:
            style_cmds.append(("BACKGROUND", (7,r_idx), (7,r_idx), _RED))

    table.setStyle(TableStyle(style_cmds))
    elements.append(table)

    doc.build(elements)
    return buf.getvalue()
