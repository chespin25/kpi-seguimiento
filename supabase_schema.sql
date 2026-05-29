CREATE TABLE IF NOT EXISTS workers (
  codigo       TEXT PRIMARY KEY,
  nombre       TEXT NOT NULL,
  gerencia     TEXT NOT NULL,
  subgerencia  TEXT,
  area         TEXT DEFAULT '',
  cargo        TEXT,
  comentario   TEXT DEFAULT '',
  updated_at   TIMESTAMPTZ DEFAULT NOW()
);

-- Migración: agregar columna area si la tabla ya existe
ALTER TABLE workers ADD COLUMN IF NOT EXISTS area TEXT DEFAULT '';

CREATE TABLE IF NOT EXISTS kpi_state (
  codigo           TEXT NOT NULL,
  periodo          TEXT NOT NULL,
  subio_ficha      TEXT DEFAULT 'NO',
  nro_objetivos    INT,
  formato_firmado  TEXT DEFAULT 'NO',
  comentario_extra TEXT DEFAULT '',
  updated_at       TIMESTAMPTZ DEFAULT NOW(),
  PRIMARY KEY (codigo, periodo)
);

CREATE INDEX IF NOT EXISTS idx_workers_gerencia ON workers(gerencia);
CREATE INDEX IF NOT EXISTS idx_kpi_periodo      ON kpi_state(periodo);

ALTER TABLE workers   DISABLE ROW LEVEL SECURITY;
ALTER TABLE kpi_state DISABLE ROW LEVEL SECURITY;
