
# Export helpers
from fastapi import APIRouter
from fastapi.responses import PlainTextResponse
from ..db import get_conn

router = APIRouter()

@router.get('/ranges.csv')
async def export_ranges_csv():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute('SELECT name,sheet,range_address,category,visible_state,protected,used_cells FROM named_ranges ORDER BY name')
    rows = cur.fetchall()
    conn.close()
    header = 'name,sheet,range_address,category,visible_state,protected,used_cells'
    lines = [header]
    for r in rows:
        vals = [r['name'] or '', r['sheet'] or '', r['range_address'] or '', r['category'] or '', r['visible_state'] or '', str(r['protected'] or 0), str(r['used_cells'] or 0)]
        lines.append(','.join([v.replace(',', ';') for v in vals]))
    csv_text = '\n'.join(lines)
    return PlainTextResponse(csv_text, media_type='text/csv')