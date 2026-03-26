# EmberAppd

A web-based underwriting and reporting platform for residential land development. Built as a full replacement for the Excel pro forma model — all calculations run server-side in Python, and the results are accessible to the whole team through a shared URL with no Excel required.

Deployed on [Railway](https://railway.app) with a PostgreSQL database.

---

## Purpose

Ember's land acquisition team underwrites residential tract deals using a detailed 10-sheet Excel model. This application replicates that model exactly in Python, adds multi-user project management, and pulls in live portfolio reporting from a separately maintained Excel dashboard file. The goal is a single source of truth for all active and prospective deals, accessible from any browser.

---

## Application Pages

### MPC Underwriting (`/`)
The core underwriting tool. Users enter deal inputs (land cost, lot mix, front footage, section pacing, infrastructure costs, debt terms, MUD/WCID structure, etc.) and the calculation engine produces a full monthly pro forma — revenues, costs, cashflows, IRR, and equity multiple. Projects are saved per-user to the database and can be revisited or updated at any time.

### Active Project Returns (`/returns`)
Displays consolidated LP-level return metrics for all active projects. Populated by uploading the Ember Dashboard Excel file. Shows per-project tables (LP distributions, contributions, profit, IRR, equity multiple, promote) broken out by year, plus a portfolio-level cashflow summary. Exportable as PDF or Excel.

### Loan Capacities & Debt Schedules (`/loans`)
Displays the current loan book — terms, drawn amounts, interest reserves, utilization, and capacity health for each active loan facility. Populated from the same uploaded dashboard file. Exportable as PDF.

### Ember Operating Revenues (`/operations`)
Tracks fee revenue across all active projects — development fees, project personnel, bookkeeping, receivables & bond fees, and brokerage. Shows KPI cards, an annual forecast, monthly history (with date filter), next 12 months, and next 12 quarters. Populated from the uploaded dashboard file. Exportable as PDF or Excel.

### Home (`/home`)
Landing page with navigation cards to all five pages above, plus admin controls for user management.

---

## Repository Structure

```
EmberAcquisitions/
│
├── app.py                  # Flask application — routes, auth, DB access, Excel export
├── calc.py                 # Pure-Python calculation engine (ports the 10-sheet Excel model)
├── report_parser.py        # Parses the uploaded Ember Dashboard Excel file into JSON
├── excel_export.py         # Excel export helper for MPC underwriting output
├── requirements.txt        # Python dependencies
├── Procfile                # Process entry point for Railway
├── railway.toml            # Railway deployment configuration
│
├── templates/
│   ├── login.html          # Login page
│   ├── home.html           # Home / navigation page
│   ├── app.html            # MPC Underwriting (inputs + pro forma output)
│   ├── returns.html        # Active Project Returns
│   ├── loans.html          # Loan Capacities & Debt Schedules
│   └── operations.html     # Ember Operating Revenues
│
└── static/
    └── img/
        └── ember_logo.png
```

---

## Key Files

### `calc.py`
The calculation engine. A faithful Python port of the Excel underwriting model — every revenue line, cost line, and cashflow formula is implemented here with explicit comments referencing the original Excel sheet and row. Takes a flat dict of user inputs and returns a dict of monthly arrays and summary outputs. No Excel dependency at runtime.

### `report_parser.py`
Reads the Ember Dashboard `.xlsx` file (uploaded by an admin) using openpyxl and extracts three datasets — project returns, loan schedules, and operating revenues — into structured JSON, which is stored in the database and served to the reporting pages.

### `app.py`
Flask backend. Handles:
- Session-based authentication (username/password, bcrypt hashing)
- Per-user page access controls (JSONB column on the `users` table)
- CRUD API for underwriting projects
- Dashboard file upload and report storage
- Excel export routes for Returns (`/api/export-returns-excel`) and Operating Revenues (`/api/export-operations-excel`)
- Admin endpoints for user management

### `templates/app.html`
The MPC underwriting frontend. A single-page vanilla JS application (~4000 lines) — all input handling, chart rendering (Chart.js), and PDF export logic live here. Communicates with the backend via `fetch` JSON calls.

---

## Stack

| Layer | Technology |
|---|---|
| Backend | Python 3, Flask |
| Database | PostgreSQL (psycopg2) |
| Frontend | Vanilla JS, Chart.js, jsPDF, html2canvas |
| Auth | Session-based, werkzeug password hashing |
| Hosting | Railway |
| Excel I/O | openpyxl |

---

## Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Set environment variables
export DATABASE_URL="postgresql://user:password@localhost/ember"
export SECRET_KEY="dev-secret-key"

# Run
python app.py
# Open http://localhost:5001
```

---

## Deployment

See [`DEPLOY.md`](DEPLOY.md) for full Railway deployment instructions, environment variable setup, and first-login credentials.

---

## Data Flow

```
User inputs (browser)
        │
        ▼
    app.py API
        │
        ▼
    calc.py  ──────────────────► Pro forma outputs (JSON)
                                          │
                                          ▼
                                  PostgreSQL (projects table)


Admin uploads Dashboard .xlsx
        │
        ▼
  report_parser.py
        │
        ▼
  PostgreSQL (reports table)
        │
        ├──► /returns   (project returns data)
        ├──► /loans     (loan schedule data)
        └──► /operations (fee revenue data)
```
