
# Project and work breakdown endpoints
from fastapi import APIRouter, HTTPException
from typing import List
from ..db import get_conn
from ..models.entities import Project, ProjectCreate, WorkPackage, WorkPackageCreate, LineItem, LineItemCreate

router = APIRouter()

@router.get('/', response_model=List[Project])
async def list_projects():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('SELECT * FROM projects ORDER BY created_at DESC')
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

@router.post('/', response_model=Project, status_code=201)
async def create_project(payload: ProjectCreate):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('INSERT INTO projects(name,region,currency_code,created_at) VALUES(?,?,?,datetime("now"))', (payload.name, payload.region, payload.currency_code))
    pid = cur.lastrowid
    cur.execute('SELECT * FROM projects WHERE id = ?', (pid,))
    row = dict(cur.fetchone())
    conn.commit()
    conn.close()
    return row

@router.get('/{project_id}/packages', response_model=List[WorkPackage])
async def list_packages(project_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('SELECT * FROM work_packages WHERE project_id = ? ORDER BY id', (project_id,))
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

@router.post('/{project_id}/packages', response_model=WorkPackage, status_code=201)
async def create_package(project_id: int, payload: WorkPackageCreate):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('INSERT INTO work_packages(project_id,name,description) VALUES(?,?,?)', (project_id, payload.name, payload.description))
    wpid = cur.lastrowid
    cur.execute('SELECT * FROM work_packages WHERE id = ?', (wpid,))
    row = dict(cur.fetchone())
    conn.commit()
    conn.close()
    return row

@router.get('/packages/{work_package_id}/items', response_model=List[LineItem])
async def list_items(work_package_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('SELECT * FROM line_items WHERE work_package_id = ? ORDER BY id', (work_package_id,))
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

@router.post('/packages/{work_package_id}/items', response_model=LineItem, status_code=201)
async def add_item(work_package_id: int, payload: LineItemCreate):
    total = payload.quantity * payload.unit_rate
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('INSERT INTO line_items(work_package_id,cost_element_id,quantity,unit_rate,currency_code,total) VALUES(?,?,?,?,?,?)', (
        work_package_id, payload.cost_element_id, payload.quantity, payload.unit_rate, payload.currency_code, total
    ))
    iid = cur.lastrowid
    cur.execute('SELECT * FROM line_items WHERE id = ?', (iid,))
    row = dict(cur.fetchone())
    conn.commit()
    conn.close()
    return row
