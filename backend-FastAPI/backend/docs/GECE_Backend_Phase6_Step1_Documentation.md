# GECE Backend – Phase 6 Step 1 Documentation

## 🎯 الهدف
تحويل منطق GECE (VBA/Excel) إلى **Backend FastAPI** متصل بقاعدة تشغيلية (Runtime DB) مع تشغيل محلي ثابت ومستقر.

---

## ⚙️ المكونات الأساسية
- **FastAPI + Uvicorn + Pydantic**
- **SQLite Runtime DB**
- **بيئة افتراضية:** `.venv`
- **إعدادات:** `.env`
- **تشغيل سريع:** `run_backend.bat`

---

## 📁 هيكل المشروع الحالي
GECE_WebCore/
│
├── backend-FastAPI/
│   ├── .env
│   ├── run_backend.bat
│   ├── .venv/
│   └── backend/
│       ├── app/
│       │   ├── main.py
│       │   ├── routers/
│       │   └── models/
│       ├── data/              ← Runtime DB
│       └── requirements.txt
│
└── data/
    └── GECE_Master.db         ← الأصلية (مرجع فقط)

---

## 🧩 الإعداد المبدئي
### إنشاء البيئة الافتراضية
```powershell
cd "C:\GECE Rev.0\GECE_WebCore\backend-FastAPI\backend"
py -3.11 -m venv ..\.venv
..\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### ملف البيئة `.env`
يوجد في `backend-FastAPI/`
```env
DB_PATH=.\backend\data\GECE_Master.db
PYTHONPATH=.
PORT=8000
```

---

## ▶️ التشغيل
```powershell
# من داخل backend/
..\.venv\Scripts\python.exe -m uvicorn app.main:app --reload --port 8000
```

### أو باستخدام السكربت:
📄 `run_backend.bat`
```bat
@echo off
cd /d "%~dp0"
echo Starting GECE Backend Server...
start "" http://127.0.0.1:8000/docs
".\.venv\Scripts\python.exe" -m uvicorn backend.app.main:app --reload --port 8000
pause
```

---

## 🔍 أوامر اختبار سريعة (PowerShell)
```powershell
curl.exe http://127.0.0.1:8000/meta/health
curl.exe http://127.0.0.1:8000/meta/stats
curl.exe "http://127.0.0.1:8000/ranges?page=1&page_size=10"
```

---

## 🧱 النقاط المهمة
- قاعدة البيانات التشغيلية (Runtime) داخل:
  backend-FastAPI/backend/data/
- القاعدة الأصلية (GECE_Master.db) في الجذر للمرجعية فقط.
- السيرفر يشتغل على **Python 3.11** لدعم Type Unions (`str | None`).

---

## 🧰 المشاكل التي تم حلها
| المشكلة | السبب | الحل |
|----------|--------|------|
| Activate.ps1 Blocked | PowerShell ExecutionPolicy | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` |
| TypeError للـ`| None` | Python قديم (3.9) | تحديث لـ3.11 |
| SyntaxError في exports.py | String غير مغلقة | تصحيح السطر 21–22 |
| ModuleNotFoundError: app | مسار خاطئ | ضبط `PYTHONPATH="."` |
| curl errors | استخدام PowerShell curl | استخدم `curl.exe` بدلًا منه |

---

## 🧾 Phase Log
- إعداد FastAPI بنجاح وتشغيله محليًا ✅  
- ربط السيرفر بقاعدة Runtime ✅  
- فصل الأصلية عن التشغيلية ✅  
- اختبار جميع المسارات الأساسية ✅  
- تشغيل صفحة Swagger ✅  
- إضافة `.gitignore` و`.env` ✅  
- إعداد `run_backend.bat` ✅  

---

## 🧠 قرارات تصميم
- **الفصل التام** بين قاعدة البيانات الأصلية (Reference) وقاعدة التشغيل (Runtime).
- اعتماد **Python 3.11** كأساس ثابت للبيئة.
- الاحتفاظ بالـ`.env` خارج Git.
- جعل كل التحديثات على Runtime DB فقط.
- توثيق كامل المسارات داخل هذا الملف الواحد.

---

## 🧩 هيكل المشروع JSON (للرجوع السريع)
```json
{
  "backend-FastAPI": {
    ".env": "Runtime config",
    ".venv/": "Local virtual environment",
    "run_backend.bat": "Double-click to start server and open Swagger",
    "backend": {
      "app": {
        "main.py": "FastAPI entry point",
        "routers/": "Endpoints (meta, ranges, projects, costing, exchange, exports)",
        "models/": "Pydantic schemas"
      },
      "data/": "Runtime SQLite DB",
      "requirements.txt": "Python dependencies"
    }
  },
  "data": {
    "GECE_Master.db": "Original reference DB (not used by server)"
  }
}
```

---

## 📦 Git Commit & Tag
```bash
git add backend-FastAPI backend/docs .gitignore
git commit -m "docs: Phase6 Step1 full backend setup and run"
git tag v1.1-backend-docs
git push origin main --tags
```

---

## ✅ الحالة النهائية
كل شيء في Phase 6 Step 1 تم إنجازه بنجاح.
السيرفر يعمل بثبات محليًا، جاهز للتوسعة في:
> **Phase 6 Step 2 – API Expansion & Data Binding**

---

© GECE Backend Development – maintained by Ahmed (Founder & Dev Lead)
