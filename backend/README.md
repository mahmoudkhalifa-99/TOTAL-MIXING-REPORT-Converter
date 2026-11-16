# وحدة استخراج الجداول عبر Python FastAPI (Camelot + pdfplumber)

توفر هذه الخدمة نقطة نهاية محلية (`POST /extract`) لمعالجة ملفات PDF المعقدة باستخدام **Camelot** و**pdfplumber** بدون أي خدمات مدفوعة أو ذكاء اصطناعي.

## التثبيت والتشغيل

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn pdf_extractor_api:app --host 0.0.0.0 --port 8000
```

> ملاحظة: يرتبط تطبيق Node/Express تلقائياً بهذه الخدمة عبر المسار `/api/python-pdf/*` (أو المتغير البيئي `PYTHON_PDF_API_URL`، والافتراضي `http://127.0.0.1:8000`).
