
# Named ranges CRUD
from fastapi import APIRouter, HTTPException
from typing import List
from ..db import get_conn, log_event
from ..models.entities import NamedRange, NamedRangeCreate

router = APIRouter()

@router.get('/', response_model=List[NamedRange])
async def list_ranges(q: str | None = None):
    conn = get_conn()
    cur = conn.cursor()
    if q:
        cur.execute('SELECT * FROM named_ranges WHERE name LIKE ? OR sheet LIKE ? ORDER BY name', ('%' + q + '%','%' + q + '%'))
    else:
        cur.execute('SELECT * FROM named_ranges ORDER BY name')
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

@router.get('/{name}', response_model=NamedRange)
async def get_range(name: str):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('SELECT * FROM named_ranges WHERE name = ?', (name,))
    r = cur.fetchone()
    conn.close()
    if not r:
        raise HTTPException(status_code=404, detail='not found')
    return dict(r)

@router.post('/', response_model=NamedRange, status_code=201)
async def create_range(payload: NamedRangeCreate):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('INSERT INTO named_ranges(name,sheet,range_address,category,formula,visible_state,protected,used_cells,notes) VALUES(?,?,?,?,?,?,?,?,?)', (
        payload.name, payload.sheet, payload.range_address, payload.category, payload.formula, payload.visible_state, payload.protected, payload.used_cells, payload.notes
    ))
    rid = cur.lastrowid
    cur.execute('SELECT * FROM named_ranges WHERE id = ?', (rid,))
    row = dict(cur.fetchone())
    conn.commit()
    conn.close()
    log_event('create_named_range','api','name=' + payload.name)
    return row

@router.delete('/{name}', status_code=204)
async def delete_range(name: str):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('DELETE FROM named_ranges WHERE name = ?', (name,))
    conn.commit()
    conn.close()
    log_event('delete_named_range','api','name=' + name)
    return None
