
# Meta endpoints
from fastapi import APIRouter
from ..db import get_conn

router = APIRouter()

@router.get('/health')
async def health():
    return {'ok': True}

@router.get('/stats')
async def stats():
    conn = get_conn()
    cur = conn.cursor()
    tables = ['sheets','named_ranges','currencies','exchange_rates','cost_elements','projects','work_packages','line_items','audit_log']
    out = {}
    for t in tables:
        cur.execute('SELECT COUNT(1) AS c FROM ' + t)
        out[t] = cur.fetchone()['c']
    conn.close()
    return out
