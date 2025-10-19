
# Costing rollups and summaries
from fastapi import APIRouter
from ..db import get_conn

router = APIRouter()

@router.get('/rollup/{project_id}')
async def rollup_project(project_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('''
        SELECT p.id as project_id, p.name as project_name, p.currency_code,
               SUM(li.total) as total
        FROM projects p
        JOIN work_packages wp ON wp.project_id = p.id
        JOIN line_items li ON li.work_package_id = wp.id
        WHERE p.id = ?
    ''', (project_id,))
    r = cur.fetchone()
    conn.close()
    return dict(r) if r else {'project_id': project_id, 'total': 0}
