### **GECE Backend Development – Master Documentation**

#### **1. Current Status**  
- Backend FastAPI scaffold working ✅  
- Local SQLite DB connected (`GECE_Master.db`)  
- Docs & run scripts verified  
- Environment: Python 3.11 (.venv active)  

---

#### **2. Next Major Objective**  
**ETL + Backend Refactor using extracted GECE XLSX data**  
Target: make backend logic replicate XLSM behavior.

---

#### **3. ETL Inputs (Excel Files)**
| File | Purpose | Size | Notes |
|------|----------|------|-------|
| GECE_Workbook_CellProvenance.xlsx | Cell-level metadata | 10.6 MB | 13 columns × 268 K rows |
| GECE_Logic_and_Functions.xlsx | Logic/subroutine map | 109 KB | 566 rows |
| GECE_ALL_Formulas.xlsx | Formula repository | 2.9 MB | 6 cols × 65 K rows |
| GECE_workbook_CellComments.xlsx | Comments & tooltips | Small | 280 rows |

---

#### **4. Planned ETL → Integration Steps**  
1. **Data Profiling & Normalization** – clean & align columns.  
2. **Database Schema Expansion** – new tables for provenance, formulas, logic, comments.  
3. **Data Injection** – import cleaned data to SQLite via FastAPI or direct script.  
4. **Backend Refactor** – update routers/models to use enriched data.  
5. **Validation & Testing** – verify outputs match XLSM logic.  

---

#### **5. Tools & Collaboration**  
- Julius.ai → automated ETL + schema inference.  
- ChatGPT → backend design & FastAPI logic alignment.  
- GitHub → versioning & documentation updates.  

---

#### **6. Next Action**  
→ Upload ETL planning prompt for Julius.  
→ Generate intermediate `.csv` / `.db` outputs for review.  
→ Begin backend refactor (ranges / logic / costing endpoints).
