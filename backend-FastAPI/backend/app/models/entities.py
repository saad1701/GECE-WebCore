
# Pydantic models
from pydantic import BaseModel
from typing import Optional, List

class NamedRange(BaseModel):
    id: int
    name: str
    sheet: Optional[str] = None
    range_address: Optional[str] = None
    category: Optional[str] = None
    formula: Optional[str] = None
    visible_state: Optional[str] = None
    protected: Optional[int] = 0
    used_cells: Optional[int] = 0
    notes: Optional[str] = None

class NamedRangeCreate(BaseModel):
    name: str
    sheet: Optional[str] = None
    range_address: Optional[str] = None
    category: Optional[str] = None
    formula: Optional[str] = None
    visible_state: Optional[str] = 'visible'
    protected: Optional[int] = 0
    used_cells: Optional[int] = 0
    notes: Optional[str] = None

class Project(BaseModel):
    id: int
    name: str
    region: Optional[str] = None
    currency_code: Optional[str] = None
    created_at: str

class ProjectCreate(BaseModel):
    name: str
    region: Optional[str] = None
    currency_code: Optional[str] = 'USD'

class WorkPackage(BaseModel):
    id: int
    project_id: int
    name: str
    description: Optional[str] = None

class WorkPackageCreate(BaseModel):
    project_id: int
    name: str
    description: Optional[str] = None

class LineItem(BaseModel):
    id: int
    work_package_id: int
    cost_element_id: int
    quantity: float
    unit_rate: float
    currency_code: str
    total: float

class LineItemCreate(BaseModel):
    work_package_id: int
    cost_element_id: int
    quantity: float
    unit_rate: float
    currency_code: str

class ExchangeRate(BaseModel):
    base: str
    quote: str
    rate: float
    as_of: str
