"""
etl_staging_import.py
ETL script to import Excel files into staging tables with row-level checksums and provenance.
"""
import sqlite3
import pandas as pd
import hashlib
from datetime import datetime

DB_PATH = '../data_runtime/GECE_Runtime.db'

def compute_row_hash(row_dict):
    """Compute SHA256 hash of row data for deduplication."""
    row_str = '|'.join(str(v) for v in sorted(row_dict.items()))
    return hashlib.sha256(row_str.encode('utf-8')).hexdigest()

def import_cell_provenance(file_path, encoding='utf-16'):
    """Import GECE_Workbook_CellProvenance.xlsx into stg_CellProvenance."""
    conn = sqlite3.connect(DB_PATH)
    
    # Read Excel file
    df = pd.read_excel(file_path, sheet_name=0)
    
    # Add provenance columns
    df['source_file'] = file_path
    df['import_ts'] = datetime.now().isoformat()
    df['row_hash'] = df.apply(lambda r: compute_row_hash(r.to_dict()), axis=1)
    
    # Insert into staging
    df.to_sql('stg_CellProvenance', conn, if_exists='replace', index=False)
    conn.commit()
    
    count = len(df)
    conn.close()
    print('Imported ' + str(count) + ' rows into stg_CellProvenance')
    return count

def import_formulas(file_path):
    """Import GECE_ALL_Formulas.xlsx into stg_Formulas."""
    conn = sqlite3.connect(DB_PATH)
    
    df = pd.read_excel(file_path, sheet_name=0)
    df['source_file'] = file_path
    df['import_ts'] = datetime.now().isoformat()
    df['row_hash'] = df.apply(lambda r: compute_row_hash(r.to_dict()), axis=1)
    
    df.to_sql('stg_Formulas', conn, if_exists='replace', index=False)
    conn.commit()
    
    count = len(df)
    conn.close()
    print('Imported ' + str(count) + ' rows into stg_Formulas')
    return count

def import_functions(file_path):
    """Import GECE_Logic_and_Functions.xlsx into stg_Functions."""
    conn = sqlite3.connect(DB_PATH)
    
    df = pd.read_excel(file_path, sheet_name=0)
    df['source_file'] = file_path
    df['import_ts'] = datetime.now().isoformat()
    df['row_hash'] = df.apply(lambda r: compute_row_hash(r.to_dict()), axis=1)
    
    df.to_sql('stg_Functions', conn, if_exists='replace', index=False)
    conn.commit()
    
    count = len(df)
    conn.close()
    print('Imported ' + str(count) + ' rows into stg_Functions')
    return count

def import_comments(file_path):
    """Import GECE_workbook_CellComments.xlsx into stg_Comments."""
    conn = sqlite3.connect(DB_PATH)
    
    df = pd.read_excel(file_path, sheet_name=0)
    df['source_file'] = file_path
    df['import_ts'] = datetime.now().isoformat()
    df['row_hash'] = df.apply(lambda r: compute_row_hash(r.to_dict()), axis=1)
    
    df.to_sql('stg_Comments', conn, if_exists='replace', index=False)
    conn.commit()
    
    count = len(df)
    conn.close()
    print('Imported ' + str(count) + ' rows into stg_Comments')
    return count

if __name__ == '__main__':
    import_cell_provenance('../../GECE_Workbook_CellProvenance.xlsx')
    import_formulas('../../GECE_ALL_Formulas.xlsx')
    import_functions('../../GECE_Logic_and_Functions.xlsx')
    import_comments('../../GECE_workbook_CellComments.xlsx')
    print('All staging imports complete')
