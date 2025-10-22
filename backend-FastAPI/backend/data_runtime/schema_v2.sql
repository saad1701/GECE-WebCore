-- ===============================
-- GECE Schema v2 (SQLite version)
-- ===============================

PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

-- ========== Core Tables ==========

CREATE TABLE IF NOT EXISTS sheets (
    id INTEGER PRIMARY KEY,
    name TEXT UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS named_ranges (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    sheet_id INTEGER NULL REFERENCES sheets(id),
    cell_range TEXT NOT NULL,
    category TEXT NULL,
    visible_state TEXT DEFAULT 'visible',
    protected INTEGER DEFAULT 0,
    used_cells INTEGER DEFAULT 0,
    notes TEXT NULL,
    UNIQUE(name, sheet_id, cell_range)
);
CREATE INDEX IF NOT EXISTS ix_named_ranges_name_sheet ON named_ranges(name, sheet_id);

CREATE TABLE IF NOT EXISTS named_ranges_agg (
    id INTEGER PRIMARY KEY,
    named_range_id INTEGER NOT NULL REFERENCES named_ranges(id),
    n_rows INTEGER,
    n_cols INTEGER,
    n_cells INTEGER,
    dup_count INTEGER,
    sample_value TEXT
);

CREATE TABLE IF NOT EXISTS cells (
    id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL REFERENCES sheets(id),
    address TEXT NOT NULL,          -- e.g. B12
    row_idx INTEGER NOT NULL,       -- 1-based
    col_idx INTEGER NOT NULL,       -- 1-based
    value_text TEXT NULL,
    value_num REAL NULL,
    value_type TEXT NULL,           -- text | number | bool | error | blank
    UNIQUE(sheet_id, address)
);
CREATE INDEX IF NOT EXISTS ix_cells_sheet_row_col ON cells(sheet_id, row_idx, col_idx);

CREATE TABLE IF NOT EXISTS range_cells (
    named_range_id INTEGER NOT NULL REFERENCES named_ranges(id),
    cell_id INTEGER NOT NULL REFERENCES cells(id),
    PRIMARY KEY(named_range_id, cell_id)
);
CREATE INDEX IF NOT EXISTS ix_range_cells_cell ON range_cells(cell_id);

CREATE TABLE IF NOT EXISTS formulas (
    id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL REFERENCES sheets(id),
    cell_address TEXT NOT NULL,
    formula TEXT NOT NULL,
    parsed_ok INTEGER DEFAULT 1,
    error TEXT NULL,
    UNIQUE(sheet_id, cell_address)
);
CREATE INDEX IF NOT EXISTS ix_formulas_sheet_cell ON formulas(sheet_id, cell_address);

CREATE TABLE IF NOT EXISTS functions (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    kind TEXT NULL,             -- Sub | Function
    module_name TEXT NULL,
    signature TEXT NULL,
    body TEXT NULL,
    UNIQUE(name, module_name)
);

CREATE TABLE IF NOT EXISTS comments (
    id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL REFERENCES sheets(id),
    cell_address TEXT NOT NULL,
    comment TEXT NOT NULL,
    UNIQUE(sheet_id, cell_address)
);
CREATE INDEX IF NOT EXISTS ix_comments_sheet_cell ON comments(sheet_id, cell_address);

CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY,
    event TEXT NOT NULL,
    actor TEXT NOT NULL,
    details TEXT NULL,
    status TEXT DEFAULT 'ok',
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);

-- ========== Staging Tables ==========

CREATE TABLE IF NOT EXISTS stg_CellProvenance (
    id INTEGER PRIMARY KEY,
    source_file TEXT NOT NULL,
    source_row INTEGER,
    imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
    checksum TEXT,
    repo_owner TEXT NULL,
    repo_name TEXT NULL,
    ref_or_sha TEXT NULL
);
CREATE INDEX IF NOT EXISTS ix_stg_CellProvenance_checksum ON stg_CellProvenance(checksum);

CREATE TABLE IF NOT EXISTS stg_Formulas (
    id INTEGER PRIMARY KEY,
    source_file TEXT NOT NULL,
    source_row INTEGER,
    imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
    checksum TEXT,
    repo_owner TEXT NULL,
    repo_name TEXT NULL,
    ref_or_sha TEXT NULL
);
CREATE INDEX IF NOT EXISTS ix_stg_Formulas_checksum ON stg_Formulas(checksum);

CREATE TABLE IF NOT EXISTS stg_Functions (
    id INTEGER PRIMARY KEY,
    source_file TEXT NOT NULL,
    source_row INTEGER,
    imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
    checksum TEXT,
    repo_owner TEXT NULL,
    repo_name TEXT NULL,
    ref_or_sha TEXT NULL
);
CREATE INDEX IF NOT EXISTS ix_stg_Functions_checksum ON stg_Functions(checksum);

CREATE TABLE IF NOT EXISTS stg_Comments (
    id INTEGER PRIMARY KEY,
    source_file TEXT NOT NULL,
    source_row INTEGER,
    imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
    checksum TEXT,
    repo_owner TEXT NULL,
    repo_name TEXT NULL,
    ref_or_sha TEXT NULL
);
CREATE INDEX IF NOT EXISTS ix_stg_Comments_checksum ON stg_Comments(checksum);

-- ========== Ledger for Imports ==========

CREATE TABLE IF NOT EXISTS import_ledger (
    id INTEGER PRIMARY KEY,
    file_name TEXT NOT NULL,
    file_mtime TEXT,
    row_count INTEGER,
    last_checksum TEXT,
    imported_at TEXT DEFAULT CURRENT_TIMESTAMP
);

-- ========== Index Summary ==========

CREATE INDEX IF NOT EXISTS ix_named_ranges_name_sheet_range ON named_ranges(name, sheet_id, cell_range);
CREATE INDEX IF NOT EXISTS ix_cells_sheet_rowcol ON cells(sheet_id, row_idx, col_idx);
CREATE INDEX IF NOT EXISTS ix_formulas_sheet_addr ON formulas(sheet_id, cell_address);
CREATE INDEX IF NOT EXISTS ix_comments_sheet_addr ON comments(sheet_id, cell_address);
CREATE INDEX IF NOT EXISTS ix_range_cells_named ON range_cells(named_range_id);
CREATE INDEX IF NOT EXISTS ix_range_cells_cellid ON range_cells(cell_id);
