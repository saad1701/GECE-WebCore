
# Database module: creates a fresh SQLite database with schema and seed data
import sqlite3
from pathlib import Path
from datetime import datetime

DB_PATH = Path(__file__).resolve().parents[1] / 'data' / 'GECE_Master.db'

def get_conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.executescript(
        """
        PRAGMA journal_mode=WAL;
        CREATE TABLE IF NOT EXISTS sheets(
            id INTEGER PRIMARY KEY,
            name TEXT UNIQUE,
            description TEXT
        );
        CREATE TABLE IF NOT EXISTS named_ranges(
            id INTEGER PRIMARY KEY,
            name TEXT UNIQUE,
            sheet TEXT,
            range_address TEXT,
            category TEXT,
            formula TEXT,
            visible_state TEXT,
            protected INTEGER,
            used_cells INTEGER,
            notes TEXT
        );
        CREATE TABLE IF NOT EXISTS currencies(
            code TEXT PRIMARY KEY,
            name TEXT,
            symbol TEXT
        );
        CREATE TABLE IF NOT EXISTS exchange_rates(
            base TEXT,
            quote TEXT,
            rate REAL,
            as_of TEXT,
            PRIMARY KEY(base, quote, as_of)
        );
        CREATE TABLE IF NOT EXISTS cost_elements(
            id INTEGER PRIMARY KEY,
            code TEXT UNIQUE,
            name TEXT,
            category TEXT,
            unit TEXT,
            default_rate REAL
        );
        CREATE TABLE IF NOT EXISTS projects(
            id INTEGER PRIMARY KEY,
            name TEXT,
            region TEXT,
            currency_code TEXT,
            created_at TEXT
        );
        CREATE TABLE IF NOT EXISTS work_packages(
            id INTEGER PRIMARY KEY,
            project_id INTEGER,
            name TEXT,
            description TEXT,
            FOREIGN KEY(project_id) REFERENCES projects(id)
        );
        CREATE TABLE IF NOT EXISTS line_items(
            id INTEGER PRIMARY KEY,
            work_package_id INTEGER,
            cost_element_id INTEGER,
            quantity REAL,
            unit_rate REAL,
            currency_code TEXT,
            total REAL,
            FOREIGN KEY(work_package_id) REFERENCES work_packages(id),
            FOREIGN KEY(cost_element_id) REFERENCES cost_elements(id)
        );
        CREATE TABLE IF NOT EXISTS audit_log(
            id INTEGER PRIMARY KEY,
            event TEXT,
            actor TEXT,
            created_at TEXT,
            details TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_named_ranges_name ON named_ranges(name);
        CREATE INDEX IF NOT EXISTS idx_work_packages_project ON work_packages(project_id);
        CREATE INDEX IF NOT EXISTS idx_line_items_wp ON line_items(work_package_id);
        CREATE INDEX IF NOT EXISTS idx_line_items_ce ON line_items(cost_element_id);
        """
    )
    # Seed baseline data
    cur.executemany('INSERT OR IGNORE INTO currencies(code,name,symbol) VALUES(?,?,?)', [
        ('USD','US Dollar','$'),
        ('EUR','Euro','€'),
        ('GBP','British Pound','£')
    ])
    cur.executemany('INSERT OR IGNORE INTO exchange_rates(base,quote,rate,as_of) VALUES(?,?,?,?)', [
        ('USD','EUR',0.92,'2025-10-01'),
        ('USD','GBP',0.78,'2025-10-01'),
        ('EUR','USD',1.09,'2025-10-01')
    ])
    cur.executemany('INSERT OR IGNORE INTO cost_elements(code,name,category,unit,default_rate) VALUES(?,?,?,?,?)', [
        ('CE001','Senior Engineer','Labor','hour',150.0),
        ('CE002','Junior Engineer','Labor','hour',90.0),
        ('CE100','Steel Beam','Material','meter',45.0)
    ])
    cur.executemany('INSERT OR IGNORE INTO sheets(name,description) VALUES(?,?)', [
        ('Summary','Project summary'),
        ('Costing','Detailed cost rollup')
    ])
    cur.executemany('INSERT OR IGNORE INTO named_ranges(name,sheet,range_address,category,formula,visible_state,protected,used_cells,notes) VALUES(?,?,?,?,?,?,?,?,?)', [
        ('NR_ProjectName','Summary','A2','meta',None,'visible',1,1,'Project title cell'),
        ('NR_TotalCost','Summary','B10','calc','SUM(B2:B9)','visible',1,1,'Auto-calculated total'),
        ('NR_LaborRates','Costing','D2:D10','config',None,'hidden',0,9,'Rate table')
    ])
    now = datetime.utcnow().isoformat()
    cur.execute('INSERT INTO projects(name,region,currency_code,created_at) VALUES(?,?,?,?)', ('Sample Project','NA','USD',now))
    pid = cur.lastrowid
    cur.execute('INSERT INTO work_packages(project_id,name,description) VALUES(?,?,?)', (pid,'WP-001','Foundation works'))
    wpid = cur.lastrowid
    cur.execute('SELECT id FROM cost_elements WHERE code = ?',( 'CE001', ))
    ceid = cur.fetchone()['id']
    cur.execute('INSERT INTO line_items(work_package_id,cost_element_id,quantity,unit_rate,currency_code,total) VALUES(?,?,?,?,?,?)', (wpid,ceid,100,150.0,'USD',15000.0))
    conn.commit()
    conn.close()

def log_event(event, actor, details):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('INSERT INTO audit_log(event, actor, created_at, details) VALUES(?,?,?,?)', (event, actor, datetime.utcnow().isoformat(), details))
    conn.commit()
    conn.close()
