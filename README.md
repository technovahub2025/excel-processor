# Universal Excel Processor

A local-first Excel/CSV processing application. Upload multiple files, inspect sheets and columns, map fields, match records, merge selected columns, filter results, and export Excel/CSV.

## Stack
- Frontend: React + Vite
- Backend: Python + FastAPI
- Processing: pandas, openpyxl, RapidFuzz
- No database required.

## Requirements
- Python 3.11+
- Node.js 20+ (LTS recommended)

## Run on Windows

### 1. Start the backend
Open PowerShell in the project folder:

```powershell
cd backend
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

If PowerShell blocks activation, run:
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

### 2. Start the frontend
Open a second terminal:

```powershell
cd frontend
npm install
npm run dev
```

Open the URL printed by Vite (normally http://localhost:5173).

## Typical workflow
1. Upload a primary file and one or more additional files.
2. Select the sheet for each file.
3. Choose the matching column in the primary file and each additional file.
4. Choose exact or fuzzy matching. Fuzzy matching is a suggestion mechanism; review ambiguous matches.
5. Select columns to bring into the primary dataset.
6. Process and inspect the merged result.
7. Filter by any column and export the current result to XLSX or CSV.

## Important matching notes
- The app creates temporary IDs such as `SYS-000001` for rows in the primary dataset during a processing session.
- These IDs are internal row identities, not magic identifiers shared across separate workbooks. Other files are linked to primary rows through the matching rules selected by the user.
- Exact matching is preferred. Fuzzy matching proposes likely matches and should be reviewed before relying on the result.
- When a key is duplicated on either side, the app avoids silently treating it as a unique one-to-one match. Check the match report and your source data.
- Original uploads are not modified. Files are stored temporarily by the backend and are deleted when the server process stops. For production use, add authentication, quotas, cleanup jobs, and secure persistent storage.

## Supported formats
- `.xlsx`, `.xlsm`, `.xls`, `.csv`
- Excel workbooks with multiple sheets (select a sheet during upload).

## API docs
With the backend running, visit https://excel-processor-tiht.onrender.com/docs.

## Troubleshooting
- `.xls` support requires `xlrd`, included in requirements.
- If a package install fails, confirm Python and Node versions and retry in a fresh virtual environment.
- Keep uploaded files reasonably sized for your machine's available RAM. This initial release processes data in memory and is intended as a local MVP, not a hardened multi-user production deployment.
