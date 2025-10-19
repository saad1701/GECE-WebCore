
# Exchange rates endpoints
from fastapi import APIRouter
from typing import List
from ..db import get_conn
from ..models.entities import ExchangeRate

router = APIRouter()

@router.get('/', response_model=List[ExchangeRate])
async def list_rates():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('SELECT * FROM exchange_rates ORDER BY as_of DESC, base, quote')
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows
