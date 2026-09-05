# Electrical Business Management Website

Python Flask + SQLite + Excel. Includes customer site, admin login, company editing, services, products/stock, project gallery upload, tenders, quotations, CRM, enquiries, reviews, password change and Excel reports.

## Run (Windows PowerShell)
```powershell
cd "C:\path\to\electrical_business_pro"
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```
Open http://127.0.0.1:5000
Admin http://127.0.0.1:5000/admin/login
Default: admin / admin123 (change immediately).

Automatic WhatsApp/SMS sending requires provider credentials/API; this version stores enquiries locally and supports a WhatsApp link.

update karaych asel tr 

git add .
git commit -m "Update admin credentials"
git push origin main