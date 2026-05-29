PERIODO_ACTIVO = "2026"

# Planta de Personal - hoja PLANTA
PLANTA_HEADER_ROW = 4          # fila real donde están los encabezados (1-based)
PLANTA_SHEET = "PLANTA"
PLANTA_ENCARGOS_SHEET = "ENCARGOS"

# Índices de columnas en hoja PLANTA (0-based, a partir de la fila de header)
PLANTA_IDX = {
    "codigo":       0,   # CODIGO
    "nombre":       1,   # APELLIDOS Y NOMBRES
    "cargo":        7,   # CARGO
    "reparticion":  8,   # REPARTICIÓN  ← nivel gerencia (top-level)
    "tipo_oficina": 10,  # TIPO OFICINA
    "oficina":      11,  # OFICINA      ← subgerencia/jefatura/sección
    "area":         12,  # ÁREA         ← área geográfica (OFI. PRINC, etc.)
}

# Índices de columnas en hoja ENCARGOS (0-based, fila 1 = header)
ENC_IDX = {
    "codigo":           0,
    "reparticion_orig": 5,   # origen gerencia
    "oficina_orig":     7,   # origen oficina
    "repart_enc":       12,  # destino gerencia (encargo)
    "tipo_ofi_enc":     14,  # destino tipo oficina
    "oficina_enc":      15,  # destino oficina
}

# Ficha KPI (xlsx subidas por gerentes)
FICHA_SHEET = "Ficha de Objetivos"
FICHA_COL_N = "Nª"             # columna de numeración de objetivos

STATE_PATH = "data/state/state.json"
FICHAS_DIR = "data/fichas"
UPLOADS_DIR = "data/uploads"
EXPORTS_DIR = "data/exports"
