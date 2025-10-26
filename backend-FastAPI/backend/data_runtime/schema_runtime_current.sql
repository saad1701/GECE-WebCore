CREATE INDEX idx_cells_sheet_addr ON cells(sheet_id, address);

CREATE INDEX idx_named_ranges_sheet ON named_ranges(sheet_id);

CREATE TABLE cells (id INTEGER PRIMARY KEY, sheet_id INTEGER, address TEXT, row_idx INTEGER, col_idx INTEGER, value_text TEXT, value_num REAL, value_type TEXT, UNIQUE(sheet_id,address));

CREATE TABLE comments (id INTEGER PRIMARY KEY, sheet_id INTEGER, cell_address TEXT, comment TEXT);

CREATE TABLE formulas (id INTEGER PRIMARY KEY, sheet_id INTEGER, cell_address TEXT, formula_a1 TEXT, formula_r1c1 TEXT);

CREATE TABLE functions (id INTEGER PRIMARY KEY, name TEXT, kind TEXT, module_name TEXT, signature TEXT, body TEXT);

CREATE TABLE named_ranges (id INTEGER PRIMARY KEY, name TEXT, sheet_id INTEGER, ref_a1 TEXT, first_cell TEXT, last_cell TEXT, n_rows INTEGER, n_cols INTEGER, n_cells INTEGER, sample_value TEXT, category TEXT);

CREATE TABLE range_cells (named_range_id INTEGER, sheet_id INTEGER, cell_address TEXT);

CREATE TABLE sheets (id INTEGER PRIMARY KEY, name TEXT UNIQUE);
