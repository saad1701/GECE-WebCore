
# FastAPI application entrypoint
import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from .db import init_db

app = FastAPI(title='GECE Backend', version='1.0.0')

app.add_middleware(
    CORSMiddleware,
    allow_origins=['*'],
    allow_credentials=True,
    allow_methods=['*'],
    allow_headers=['*'],
)

init_db()

from .routers import meta, ranges, projects, costing, exchange, exports
app.include_router(meta.router, prefix='/meta', tags=['meta'])
app.include_router(ranges.router, prefix='/ranges', tags=['named_ranges'])
app.include_router(projects.router, prefix='/projects', tags=['projects'])
app.include_router(costing.router, prefix='/costing', tags=['costing'])
app.include_router(exchange.router, prefix='/exchange', tags=['exchange'])
app.include_router(exports.router, prefix='/exports', tags=['exports'])

@app.get('/')
async def root():
    return {'status':'ok','service':'gece-backend'}
