GECE — Context Engineering Brief (DDL + Staged ETL)

Objective
Apply the provided SQLite DDL (schema_v2.sql) and perform a staged, idempotent import of four Excel sources into stg_* tables with row-level checksums, then normalize into the core schema. Keep strictly to the plan and recommendations previously outlined (staging, normalizing, indexes, idempotency, and SQLite-friendly performance).

Environment / Paths
- Runtime DB (target): backend/data_runtime/GECE_Runtime.db
- DDL file: backend/data_runtime/schema_v2.sql  (included in ZIP)
- Source Excel files (included in ZIP, same folder as DB/DDL):
  1) GECE_Workbook_CellProvenance.xlsx
  2) GECE_ALL_Formulas.xlsx
  3) GECE_Logic_and_Functions.xlsx
  4) GECE_workbook_CellComments.xlsx

Scope (do exactly this)
1) DDL
   - Execute schema_v2.sql against GECE_Runtime.db.
   - PRAGMA settings as per DDL (WAL, foreign_keys ON).

2) Staged Import (stg_* with provenance + checksums)
   - Create/load the following staging tables (defined in DDL):
     * stg_CellProvenance
     * stg_Formulas
     * stg_Functions
     * stg_Comments
   - Each staging row must include:
     source_file, source_row, imported_at (CURRENT_TIMESTAMP), checksum
     (plus optional: repo_owner, repo_name, ref_or_sha — leave NULL)
   - Compute checksum = SHA256 of canonical concatenated fields (trimmed), per table:
     * CellProvenance: [Sheet, Address, ValueText, DataType, HasFormula, FormulaA1, InNamedRange, NameList, HasValidation, IsUnlocked, Role, Notes, CellFormat]
     * Formulas:      [Class, Sheet, AddressOrName, FormulaA1, FormulaR1C1, IsArray]
     * Functions:     [Subroutine, Sub_Name]  (add derived "kind" = Sub/Function if obvious)
     * Comments:      [Sheet, Cell, Comment]
   - Use chunked pandas loads (e.g., chunksize ≈ 10k) and single transaction per table.
   - Use SQLite speed-ups during load:
     PRAGMA journal_mode=WAL, synchronous=OFF, temp_store=MEMORY, cache_size=-80000
     (Restore safe defaults after loads if you need to.)

   - Maintain import_ledger (file_name, file_mtime, row_count, last_checksum, imported_at).
   - Idempotency rule: upsert/skip rows that have unchanged checksum per file.

3) Normalization into Core Schema (from staging)
   - Populate core tables exactly as in schema_v2.sql:
     sheets, named_ranges, cells, range_cells, formulas, functions, comments, named_ranges_agg, audit_log.
   - Parse A1 addresses to (row_idx, col_idx). Persist both raw address and numeric coordinates.
   - Insert/merge rules & suggested mappings:

     3.1 Sheets
         - Derive unique sheet names from all inputs (CellProvenance, Formulas, Comments).
         - INSERT IGNORE/UPSERT into sheets(name).

     3.2 Cells (from CellProvenance)
         - For each (Sheet, Address), compute row_idx/col_idx (1-based).
         - value_text = ValueText; value_num = numeric cast if applicable; value_type from DataType normalized to one of: text|number|bool|error|blank.
         - UPSERT by UNIQUE(sheet_id, address).

     3.3 Named Ranges (+ linkage to cells)
         - If available from inputs (NameList / InNamedRange / AddressOrName when Class indicates a name):
           * named_ranges(name, sheet_id?, cell_range, category, visible_state, protected, used_cells, notes)
           * Normalize “visible” fields and booleans (0/1).
         - Expand each range once (A1 notation → bounds) and populate range_cells (named_range_id, cell_id).
         - Populate named_ranges_agg (n_rows, n_cols, n_cells, dup_count, sample_value).

     3.4 Formulas
         - From GECE_ALL_Formulas.xlsx → formulas (sheet_id, cell_address, formula = FormulaA1, parsed_ok default 1; keep FormulaR1C1 in error/aux fields if helpful).
         - UPSERT by UNIQUE(sheet_id, cell_address).

     3.5 Functions
         - From GECE_Logic_and_Functions.xlsx:
           * functions(name, kind, module_name, signature, body) — keep what is available.
           * Use “Subroutine/Sub_Name” columns; infer kind if possible.
           * UPSERT by UNIQUE(name, module_name).

     3.6 Comments
         - From GECE_workbook_CellComments.xlsx → comments(sheet_id, cell_address, comment).
         - UPSERT by UNIQUE(sheet_id, cell_address).

4) Indexes & Hot Paths
   - Ensure all indexes in DDL exist, especially:
     named_ranges(name, sheet_id[, cell_range])
     cells(sheet_id, row_idx, col_idx)
     formulas(sheet_id, cell_address)
     comments(sheet_id, cell_address)
     range_cells(named_range_id) and range_cells(cell_id)

5) Validation & Audit
   - Compare row counts between staging and core (where applicable); record in audit_log (event='validation', status, details JSON/text).
   - Validate range expansions: assert n_rows * n_cols == n_cells for each named range; log anomalies.
   - Sample a few random cells/ranges; log any parse errors to audit_log (event='parse_error').

6) Deliverables (return all of the following)
   - GECE_Runtime.db (populated with core + staging + indexes).
   - ETL scripts/notebooks used to perform the above (Python scripts preferred, pandas/openpyxl).
   - Short README.txt describing how to re-run ETL locally (one command per table + full refresh).
   - (Optional) Updated FastAPI routers (see below).

7) Optional — FastAPI Routers (SQLite-friendly)
   Implement lightweight endpoints:
     - GET /ranges/{name}                → range metadata + summary (agg) + counts
     - GET /ranges/{name}/cells          → paged list of cells (value/formula/comment union)
     - GET /sheets/{sheet}/cells/{addr}  → single cell composite view
     - GET /formulas/search?q=           → search in formula text by token
     - GET /functions                    → list functions metadata
     - GET /meta/stats                   → counts for stg/core + checksum/import status

Constraints
- Keep everything SQLite-friendly (no vendor-specific SQL).
- Loads must be idempotent via checksums and/or UPSERT logic.
- Do not drop bad rows silently; put failures in audit_log with context.
- Preserve raw addresses (A1) AND numeric coords for fast joins.
- Stay within the schema_v2.sql; add only helper views if necessary.

Acceptance Criteria
- GECE_Runtime.db opens with all tables per DDL.
- stg_* tables loaded with provenance + checksum columns.
- Core tables populated; range_cells and named_ranges_agg filled.
- Indexes present; typical lookups (by named range, cell, formula) are fast.
- README + ETL scripts provided and executable locally.

Thank you — please execute exactly as above and return the populated DB + scripts.
