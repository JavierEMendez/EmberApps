"""
Ember Tract Underwriting Web App
Flask + PostgreSQL + Flask-Login — no Excel required
"""
import os, json, datetime
from functools import wraps
from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
import psycopg2
import psycopg2.extras
from werkzeug.security import generate_password_hash, check_password_hash
from calc import calculate
from report_parser import parse_dashboard
import io

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "ember-dev-secret-change-in-production")

# Auto-initialize DB on first request
_db_initialized = False

@app.before_request
def auto_init():
    global _db_initialized
    if not _db_initialized:
        try:
            init_db()
            _db_initialized = True
        except Exception as e:
            print(f"DB init error: {e}")

# ─── DATABASE ────────────────────────────────────────────────────────────────
def get_db():
    conn = psycopg2.connect(os.environ["DATABASE_URL"], cursor_factory=psycopg2.extras.RealDictCursor)
    return conn

def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            is_admin BOOLEAN DEFAULT FALSE,
            page_access JSONB NOT NULL DEFAULT '{"mpc_underwriting":true,"returns":true,"loans":true,"operations":true}'::jsonb,
            created_at TIMESTAMP DEFAULT NOW()
        );
        -- Add page_access column if upgrading from older schema
        ALTER TABLE users ADD COLUMN IF NOT EXISTS page_access JSONB NOT NULL DEFAULT '{"mpc_underwriting":true,"returns":true,"loans":true,"operations":true}'::jsonb;
        CREATE TABLE IF NOT EXISTS projects (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            address TEXT,
            created_by INTEGER REFERENCES users(id),
            created_at TIMESTAMP DEFAULT NOW(),
            updated_at TIMESTAMP DEFAULT NOW(),
            inputs JSONB NOT NULL DEFAULT '{}'::jsonb,
            outputs JSONB NOT NULL DEFAULT '{}'::jsonb,
            archived BOOLEAN DEFAULT FALSE
        );
        CREATE TABLE IF NOT EXISTS reports (
            id SERIAL PRIMARY KEY,
            report_type TEXT NOT NULL,
            data JSONB NOT NULL DEFAULT '{}'::jsonb,
            uploaded_by INTEGER REFERENCES users(id),
            uploaded_at TIMESTAMP DEFAULT NOW()
        );
    """)
    # Create default admin if no users exist
    cur.execute("SELECT COUNT(*) as cnt FROM users")
    row = cur.fetchone()
    if row["cnt"] == 0:
        cur.execute(
            "INSERT INTO users (username, password_hash, is_admin) VALUES (%s, %s, TRUE)",
            ("admin", generate_password_hash("ember2024"))
        )
    conn.commit()
    cur.close()
    conn.close()

# ─── AUTH HELPERS ─────────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            if request.is_json:
                return jsonify({"error": "Unauthorized"}), 401
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("is_admin"):
            return jsonify({"error": "Admin required"}), 403
        return f(*args, **kwargs)
    return decorated

# ─── AUTH ROUTES ─────────────────────────────────────────────────────────────
@app.route("/health")
def health():
    return "ok", 200

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE username = %s", (username,))
        user = cur.fetchone()
        cur.close(); conn.close()
        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["is_admin"] = user["is_admin"]
            session["page_access"] = user.get("page_access") or {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True}
            return redirect(url_for("home"))
        error = "Invalid username or password."
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ─── MAIN APP ─────────────────────────────────────────────────────────────────
@app.route("/home")
@login_required
def home():
    pa = session.get("page_access") or {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True}
    # Admins always have full access
    if session.get("is_admin"):
        pa = {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True}
    return render_template("home.html", username=session.get("username"), is_admin=session.get("is_admin"), page_access=pa)

@app.route("/")
@login_required
def index():
    pa = session.get("page_access") or {}
    if not session.get("is_admin") and not pa.get("mpc_underwriting", True):
        return redirect(url_for("home"))
    pa = session.get("page_access") or {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True}
    if session.get("is_admin"):
        pa = {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True}
    return render_template("app.html", username=session.get("username"), is_admin=session.get("is_admin"), page_access=pa)

# ─── PROJECT API ─────────────────────────────────────────────────────────────
@app.route("/api/projects", methods=["GET"])
@login_required
def list_projects():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT p.id, p.name, p.address, p.updated_at,
               u.username as created_by,
               p.outputs->>'total_revenue' as total_revenue,
               p.outputs->>'gross_margin_pct' as gross_margin_pct,
               p.outputs->>'total_lots' as total_lots,
               p.outputs->>'unlevered_irr' as unlevered_irr,
               p.outputs->>'project_length_years' as project_length_years,
               p.archived
        FROM projects p
        LEFT JOIN users u ON p.created_by = u.id
        WHERE p.archived = FALSE
        ORDER BY p.updated_at DESC
    """)
    rows = cur.fetchall()
    cur.close(); conn.close()
    return jsonify([dict(r) for r in rows])

@app.route("/api/projects", methods=["POST"])
@login_required
def create_project():
    data = request.json or {}
    name = data.get("name", "New Project")
    inputs = default_inputs(name)
    try:
        outputs = calculate(inputs)
    except Exception:
        outputs = {}
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO projects (name, address, created_by, inputs, outputs) VALUES (%s, %s, %s, %s, %s) RETURNING id",
        (name, data.get("address", ""), session["user_id"], json.dumps(inputs), json.dumps(outputs))
    )
    pid = cur.fetchone()["id"]
    conn.commit(); cur.close(); conn.close()
    return jsonify({"id": pid, "name": name})

@app.route("/api/projects/<int:pid>", methods=["GET"])
@login_required
def get_project(pid):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM projects WHERE id = %s", (pid,))
    row = cur.fetchone()
    cur.close(); conn.close()
    if not row:
        return jsonify({"error": "Not found"}), 404
    return jsonify(dict(row))

@app.route("/api/projects/<int:pid>", methods=["PUT"])
@login_required
def save_project(pid):
    data = request.json or {}
    inputs = data.get("inputs", {})
    try:
        outputs = calculate(inputs)
    except Exception as e:
        return jsonify({"error": f"Calculation error: {e}"}), 500
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE projects
        SET inputs = %s, outputs = %s, name = %s, address = %s, updated_at = NOW()
        WHERE id = %s
    """, (
        json.dumps(inputs),
        json.dumps(outputs),
        inputs.get("project_name", "Unnamed"),
        inputs.get("address", ""),
        pid
    ))
    conn.commit(); cur.close(); conn.close()
    return jsonify({"ok": True, "outputs": outputs})

@app.route("/api/projects/<int:pid>", methods=["DELETE"])
@login_required
def delete_project(pid):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE projects SET archived = TRUE WHERE id = %s", (pid,))
    conn.commit(); cur.close(); conn.close()
    return jsonify({"ok": True})

@app.route("/api/projects/<int:pid>/calculate", methods=["POST"])
@login_required
def recalculate(pid):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT inputs FROM projects WHERE id = %s", (pid,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return jsonify({"error": "Not found"}), 404
    inputs = row["inputs"]
    try:
        outputs = calculate(inputs)
    except Exception as e:
        cur.close(); conn.close()
        return jsonify({"error": str(e)}), 500
    cur.execute("UPDATE projects SET outputs = %s, updated_at = NOW() WHERE id = %s",
                (json.dumps(outputs), pid))
    conn.commit(); cur.close(); conn.close()
    return jsonify({"ok": True, "outputs": outputs})

# ─── ADMIN API ────────────────────────────────────────────────────────────────
@app.route("/api/admin/users", methods=["GET"])
@login_required
@admin_required
def list_users():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id, username, is_admin, page_access, created_at FROM users ORDER BY id")
    rows = cur.fetchall()
    cur.close(); conn.close()
    return jsonify([dict(r) for r in rows])

@app.route("/api/admin/users", methods=["POST"])
@login_required
@admin_required
def create_user():
    data = request.json or {}
    username = data.get("username", "").strip()
    password = data.get("password", "")
    is_admin = data.get("is_admin", False)
    page_access = data.get("page_access", {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True})
    if not username or not password:
        return jsonify({"error": "Username and password required"}), 400
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO users (username, password_hash, is_admin, page_access) VALUES (%s, %s, %s, %s) RETURNING id",
            (username, generate_password_hash(password), is_admin, json.dumps(page_access))
        )
        uid = cur.fetchone()["id"]
        conn.commit()
    except psycopg2.errors.UniqueViolation:
        conn.rollback()
        cur.close(); conn.close()
        return jsonify({"error": "Username already exists"}), 409
    cur.close(); conn.close()
    return jsonify({"id": uid, "username": username})

@app.route("/api/admin/users/<int:uid>", methods=["DELETE"])
@login_required
@admin_required
def delete_user(uid):
    if uid == session["user_id"]:
        return jsonify({"error": "Cannot delete yourself"}), 400
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM users WHERE id = %s", (uid,))
    conn.commit(); cur.close(); conn.close()
    return jsonify({"ok": True})

@app.route("/api/admin/users/<int:uid>/password", methods=["PUT"])
@login_required
@admin_required
def reset_password(uid):
    data = request.json or {}
    password = data.get("password", "")
    if not password:
        return jsonify({"error": "Password required"}), 400
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE users SET password_hash = %s WHERE id = %s",
                (generate_password_hash(password), uid))
    conn.commit(); cur.close(); conn.close()
    return jsonify({"ok": True})

@app.route("/api/admin/users/<int:uid>/access", methods=["PUT"])
@login_required
@admin_required
def update_page_access(uid):
    data = request.json or {}
    page_access = data.get("page_access", {})
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE users SET page_access = %s WHERE id = %s",
                (json.dumps(page_access), uid))
    conn.commit(); cur.close(); conn.close()
    return jsonify({"ok": True})

# ─── DEFAULT INPUTS TEMPLATE ──────────────────────────────────────────────────
def default_inputs(name="New Project"):
    # Lot size defaults match Excel Cost Inputs rows 72-87 exactly
    # Columns: FF, on, yield/ac, pace lots/mo, home_price, wsd/ff, paving/ff, landscaping/lot, urd/lot, lots_per_streetlight, fence/ff
    lot_size_defaults = [
        {"front_footage":25,  "on":0, "yield_per_ac":8.25, "pace":5,    "home_price":200000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":4, "fence_cost_per_ff":94},
        {"front_footage":30,  "on":0, "yield_per_ac":5.54, "pace":5,    "home_price":360000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":4, "fence_cost_per_ff":94},
        {"front_footage":35,  "on":0, "yield_per_ac":8.25, "pace":6,    "home_price":275000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":4, "fence_cost_per_ff":94},
        {"front_footage":40,  "on":1, "yield_per_ac":5.5,  "pace":7,    "home_price":330168,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":4, "fence_cost_per_ff":94},
        {"front_footage":45,  "on":1, "yield_per_ac":5.0,  "pace":6,    "home_price":380000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":4, "fence_cost_per_ff":94},
        {"front_footage":50,  "on":1, "yield_per_ac":4.5,  "pace":5,    "home_price":430000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":4, "fence_cost_per_ff":94},
        {"front_footage":55,  "on":0, "yield_per_ac":4.0,  "pace":5,    "home_price":500000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":4, "fence_cost_per_ff":94},
        {"front_footage":60,  "on":1, "yield_per_ac":3.5,  "pace":2,    "home_price":580000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":4, "fence_cost_per_ff":94},
        {"front_footage":65,  "on":0, "yield_per_ac":3.0,  "pace":2,    "home_price":615000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":3, "fence_cost_per_ff":94},
        {"front_footage":70,  "on":0, "yield_per_ac":2.5,  "pace":1,    "home_price":675000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":3, "fence_cost_per_ff":94},
        {"front_footage":75,  "on":0, "yield_per_ac":2.0,  "pace":1,    "home_price":720000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":3, "fence_cost_per_ff":94},
        {"front_footage":80,  "on":1, "yield_per_ac":1.5,  "pace":0.75, "home_price":750000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":3, "fence_cost_per_ff":94},
        {"front_footage":85,  "on":0, "yield_per_ac":5.5,  "pace":0.75, "home_price":325000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":2, "fence_cost_per_ff":94},
        {"front_footage":90,  "on":0, "yield_per_ac":5.5,  "pace":0.75, "home_price":360000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":2, "fence_cost_per_ff":94},
        {"front_footage":95,  "on":0, "yield_per_ac":1.15, "pace":0.75, "home_price":385000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":2, "fence_cost_per_ff":94},
        {"front_footage":100, "on":0, "yield_per_ac":1.0,  "pace":0.75, "home_price":410000,    "wsd_per_ff":290, "paving_per_ff":220, "dev_start_month":1, "landscaping_per_lot":2000, "urd_per_lot":35, "lots_per_streetlight":2, "fence_cost_per_ff":94},
    ]
    return {
        "project_name": name,
        "address": "",
        "gross_acreage": 0,
        "land_escalator": 0.05,
        "purchase_price_per_acre": 0,
        "closing_costs_pct": 0.045,
        "closing_date": "",
        "default_other_pct": 0.17,
        "sectional_other_pct": 0.17,       # Excel B6 = 0.17
        "landscaping_other_pct": 0.12,
        "contingency": 0.05,
        "site_work_pct": 0.01,
        "fenced_pct": 0.25,
        "cost_per_mailbox": 200,
        "cost_per_streetlight": 1700,
        "default_start_month": 1,
        "det_storage_rate": 1.1,            # Excel B31 = 1.1
        "det_depth": 9,                     # Excel B33 = 9
        "det_num_projects": 6,              # Excel B34 = 6
        "parks_pct": 0.03,                  # Excel B51 = 3%
        "drill_site_acres": 0,
        "commercial_pod_acres": 0,
        "residential_pod_acres": 0,
        "plants": [{"type":"None","notes":""} for _ in range(8)],
        "amenities": [{"type":"None","acres":0} for _ in range(6)],
        "other_netouts": [{"desc":"","acres":0} for _ in range(6)],
        "roads": [{"type":"","lf":0,"width":0,"road_setback":0,"landscaping_setback":0} for _ in range(6)],
        "takedowns": [{"period":0,"pct":0.5},{"period":36,"pct":0.5},{"period":0,"pct":0.0}],
        "plant_costs": [{"base_cost":0,"other_pct":0.17,"start_month":1,"ph2_base_cost":0,"ph2_other_pct":0.17,"ph2_start_month":37} for _ in range(8)],
        "amenity_costs": [{"base_cost":0,"other_pct":0.17,"start_month":1} for _ in range(6)],
        "det_costs": [{"other_pct":0.17,"landscaping_per_foot":2} for _ in range(6)],
        "other_costs": [{"base_cost":0,"other_pct":0.17,"start_month":1,"duration":1} for _ in range(6)],
        "road_costs": [{"other_pct":0.17,"start_month":1,"landscaping_per_sf":2,"light_spacing":0} for _ in range(6)],
        "lot_sizes": lot_size_defaults,
        "timing_method": "50/25/25",        # Excel B2 = 50/25/25
        "bem_period": 9,                    # Excel B3 = 9
        "bem_pct": 0.18,                    # Excel B4 = 18%
        "brokerage_fees": 0.03,             # Excel B5 = 3%
        "lot_closing_costs": 0.015,         # Excel B6 = 1.5%
        "take1_pct": 0.50,
        "take2_pct": 0.25,
        "take3_pct": 0.25,
        "price_per_ff": {str(yr): 1800 for yr in range(11)},
        "res_pod_acreage": 0,
        "res_pod_count": 1,
        "res_pods": [{"price_per_acre":120000,"closing_costs_pct":0.045,"implied_lots_per_acre":3.5,"impact_fee_per_lot":10000,"sale_period":12} for _ in range(6)],
        "comm_pod_acreage": 0,
        "comm_pod_count": 6,
        "comm_pods": [{"price_per_sf":8,"closing_costs_pct":0.045,"sale_period":12+i*24,"av_per_acre":1200000,"av_delay_months":18} for i in range(6)],
        "mud_bond": {"toggle":1,"amount":0,"reimbursement_pct":0.85,"first_bond_period":48,"bond_interval":12,"pct_to_dev":0.85,"receivables_fee":0.025,"debt_ratio":0.12},
        "wcid_bond": {"toggle":1,"amount":0,"reimbursement_pct":0.85,"first_bond_period":48,"bond_interval":12,"pct_to_dev":0.85,"receivables_fee":0.025,"debt_ratio":0.042},
        "marketing_pct": 0.02,
        "prof_svc_pct": 0.015,              # Excel B95 = 1.5%
        "dmf_pct": 0.025,                   # Excel B99 = 2.5%
        "personnel_monthly": 50000,         # Excel C103 = 50,000
        "marketing_personnel_monthly": 15000, # Excel C104 = 15,000
        "legal_monthly": 10000,             # Excel C108 = 10,000
        "mud_monthly": 35000,               # Excel C112 = 35,000
        "mud_pct": 0.2,                     # Excel D112 = 20% (what % of project MUD runs)
        "insurance_monthly": 10000,         # Excel C116 = 10,000
        "bookkeeping_monthly": 10000,       # Excel C120 = 10,000
    }

@app.route("/api/portfolio", methods=["GET"])
@login_required
def portfolio():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT p.id, p.name, p.address, p.outputs
        FROM projects p
        WHERE p.archived = FALSE
        ORDER BY p.name
    """)
    rows = cur.fetchall()
    cur.close(); conn.close()
    result = []
    for r in rows:
        o = r["outputs"] or {}
        result.append({"id": r["id"], "name": r["name"], "address": r["address"], "outputs": o})
    return jsonify(result)

@app.route("/api/projects/<int:pid>/export_excel", methods=["GET"])
@login_required
def export_excel(pid):
    try:
        from excel_export import export_excel as _export
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM projects WHERE id=%s", (pid,))
        proj = cur.fetchone()
        cur.close(); conn.close()
        if not proj:
            return jsonify({"error": "not found"}), 404
        inputs = proj["inputs"] or {}
        excel_bytes = _export(inputs)
        name = (proj.get("name") or "project").replace(" ", "_")
        return send_file(
            io.BytesIO(excel_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"{name}_Underwriting.xlsx"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/projects/<int:pid>/backup", methods=["GET"])
@login_required
def backup_project(pid):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM projects WHERE id=%s", (pid,))
    proj = dict(cur.fetchone() or {})
    cur.close(); conn.close()
    if not proj:
        return jsonify({"error": "not found"}), 404
    for k, v in proj.items():
        if hasattr(v, 'isoformat'):
            proj[k] = v.isoformat()
    name = (proj.get("name") or "project").replace(" ", "_")
    backup_data = json.dumps(proj, indent=2)
    return send_file(
        io.BytesIO(backup_data.encode()),
        mimetype="application/json",
        as_attachment=True,
        download_name=f"{name}_backup.json"
    )

@app.route("/api/projects/restore", methods=["POST"])
@login_required
def restore_project():
    data = request.json or {}
    inputs  = data.get("inputs", {})
    outputs = data.get("outputs", {})
    name    = data.get("name", "Restored Project")
    address = data.get("address", "")
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO projects (name, address, created_by, inputs, outputs) VALUES (%s,%s,%s,%s,%s) RETURNING id",
        (name, address, session["user_id"], json.dumps(inputs), json.dumps(outputs))
    )
    new_id = cur.fetchone()["id"]
    conn.commit(); cur.close(); conn.close()
    return jsonify({"ok": True, "id": new_id})

@app.route("/api/projects/import_excel", methods=["POST"])
@login_required
def import_excel_project():
    """Upload an Ember underwriting Excel and create a new project from it."""
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file provided"}), 400
    try:
        from excel_import import import_excel
        file_bytes = f.read()
        inputs = import_excel(file_bytes)
    except Exception as e:
        return jsonify({"error": f"Failed to parse Excel: {e}"}), 400
    try:
        outputs = calculate(inputs)
    except Exception:
        outputs = {}
    name = inputs.get("project_name", "Imported Project")
    address = inputs.get("address", "")
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO projects (name, address, created_by, inputs, outputs) VALUES (%s, %s, %s, %s, %s) RETURNING id",
        (name, address, session["user_id"], json.dumps(inputs), json.dumps(outputs))
    )
    pid = cur.fetchone()["id"]
    conn.commit(); cur.close(); conn.close()
    return jsonify({"ok": True, "id": pid, "name": name})

@app.route("/api/parse_excel", methods=["POST"])
@login_required
def parse_excel():
    """Parse an Ember underwriting Excel and return the inputs dict (no project created)."""
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file provided"}), 400
    try:
        from excel_import import import_excel
        file_bytes = f.read()
        inputs = import_excel(file_bytes)
    except Exception as e:
        return jsonify({"error": f"Failed to parse Excel: {e}"}), 400
    return jsonify({"ok": True, "inputs": inputs})

# ─── DASHBOARD REPORTS ────────────────────────────────────────────────────────
@app.route("/api/upload-dashboard", methods=["POST"])
@login_required
@admin_required
def upload_dashboard():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file provided"}), 400
    try:
        file_bytes = f.read()
        data = parse_dashboard(file_bytes)
    except Exception as e:
        return jsonify({"error": f"Failed to parse file: {e}"}), 400
    conn = get_db()
    cur = conn.cursor()
    # Upsert returns
    cur.execute("DELETE FROM reports WHERE report_type = 'returns'")
    cur.execute(
        "INSERT INTO reports (report_type, data, uploaded_by) VALUES (%s, %s, %s)",
        ("returns", json.dumps(data.get("returns", {})), session["user_id"])
    )
    # Upsert loans
    cur.execute("DELETE FROM reports WHERE report_type = 'loans'")
    cur.execute(
        "INSERT INTO reports (report_type, data, uploaded_by) VALUES (%s, %s, %s)",
        ("loans", json.dumps(data.get("loans", {})), session["user_id"])
    )
    # Upsert operations
    cur.execute("DELETE FROM reports WHERE report_type = 'operations'")
    if data.get("operations"):
        cur.execute(
            "INSERT INTO reports (report_type, data, uploaded_by) VALUES (%s, %s, %s)",
            ("operations", json.dumps(data["operations"]), session["user_id"])
        )
    conn.commit(); cur.close(); conn.close()
    return jsonify({"ok": True})

@app.route("/api/export-returns-excel")
@login_required
def export_returns_excel():
    pa = session.get("page_access") or {}
    if not session.get("is_admin") and not pa.get("returns", True):
        return jsonify({"error": "Access denied"}), 403
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT data, uploaded_at FROM reports WHERE report_type = 'returns' ORDER BY uploaded_at DESC LIMIT 1")
    row = cur.fetchone()
    cur.close(); conn.close()
    if not row or not row["data"]:
        return jsonify({"error": "No data available"}), 404

    data = row["data"]
    uploaded_at = row["uploaded_at"].strftime("%B %d, %Y") if row["uploaded_at"] else ""

    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from io import BytesIO

    LABEL_MAP = {"LP IRR": "Net Cashflow", "LP Equity Multiple": "Cumulative Net Cashflow"}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Project Returns"

    GOLD = "C8A96E"
    TEAL = "5E9E8C"
    HEADER_FILL = PatternFill("solid", fgColor="191E28")
    PROJ_FILL   = PatternFill("solid", fgColor="191E28")
    SUMM_FILL   = PatternFill("solid", fgColor="151921")
    thin = Side(style="thin", color="1E2535")
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _vf(bold=False, color="E8EAF0", size=9):
        return Font(name="Calibri", size=size, bold=bold, color=color)

    def _num_fmt(ws_cell, val, label):
        """Apply number format and value based on metric label."""
        if label == "LP IRR":
            ws_cell.value = val if val else None
            ws_cell.number_format = "0.0%"
        elif label == "LP Equity Multiple":
            ws_cell.value = val if val else None
            ws_cell.number_format = '0.00"x"'
        elif isinstance(val, (int, float)) and val != 0:
            ws_cell.value = val
            ws_cell.number_format = "#,##0"
        else:
            ws_cell.value = None

    years = data.get("years", [])

    # --- Determine active year columns (have any non-zero data across all projects) ---
    active_idxs = []
    for i in range(len(years)):
        for proj in data.get("projects", []):
            if any(m["yearly"][i] != 0 for m in proj.get("metrics", []) if i < len(m["yearly"])):
                active_idxs.append(i)
                break
        else:
            for s in data.get("summary", []):
                if i < len(s["yearly"]) and s["yearly"][i] != 0:
                    active_idxs.append(i)
                    break

    active_years = [years[i] for i in active_idxs]

    r = 1
    # Title
    ws.cell(row=r, column=1, value="Consolidated Ember Project Returns").font = Font(name="Calibri", bold=True, size=14, color=GOLD)
    r += 1
    ws.cell(row=r, column=1, value=f"Last updated: {uploaded_at}  |  ($ in 000s)").font = Font(name="Calibri", size=9, color="8B95A8")
    r += 2

    num_cols = 2 + len(active_years)  # label + total + years

    def write_project(r, proj, color=GOLD):
        name = proj["name"]
        metrics = proj.get("metrics", [])

        # Per-project first active year index
        pfa = len(years)
        for i in range(len(years)):
            for m in metrics:
                if i < len(m["yearly"]) and m["yearly"][i] != 0 and i < pfa:
                    pfa = i
                    break

        proj_active = [i for i in active_idxs if i >= pfa]
        proj_years  = [years[i] for i in proj_active]

        # Project header spanning all cols
        hdr = ws.cell(row=r, column=1, value=name)
        hdr.font = Font(name="Calibri", bold=True, size=10, color=color)
        hdr.fill = PROJ_FILL
        hdr.border = cell_border
        for ci in range(2, num_cols + 1):
            c = ws.cell(row=r, column=ci)
            c.fill = PROJ_FILL
            c.border = cell_border
        r += 1

        # Column headers
        ws.cell(row=r, column=1, value="Metric").font = _vf(bold=True, color="8B95A8")
        ws.cell(row=r, column=1).fill = HEADER_FILL
        ws.cell(row=r, column=1).border = cell_border
        ws.cell(row=r, column=2, value="Total").font = _vf(bold=True, color="8B95A8")
        ws.cell(row=r, column=2).fill = HEADER_FILL
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2).border = cell_border
        for ci, yr in enumerate(proj_years, 3):
            c = ws.cell(row=r, column=ci, value=yr)
            c.font = _vf(bold=True, color="8B95A8")
            c.fill = HEADER_FILL
            c.alignment = Alignment(horizontal="center")
            c.border = cell_border
        r += 1

        # Metric rows
        for m in metrics:
            label = m["label"]
            display = LABEL_MAP.get(label, label)
            # Compute total for renamed rows
            if label == "LP IRR":
                total = m["total"]
            elif label == "LP Equity Multiple":
                total = m["total"]
            else:
                total = m["total"]

            lc = ws.cell(row=r, column=1, value=display)
            lc.font = _vf(bold=(label in ("LP IRR", "LP Equity Multiple")), color=color if label in ("LP IRR", "LP Equity Multiple") else "E8EAF0")
            lc.border = cell_border

            tc = ws.cell(row=r, column=2)
            tc.font = _vf(bold=(label in ("LP IRR", "LP Equity Multiple")), color=color if label in ("LP IRR", "LP Equity Multiple") else "E8EAF0")
            tc.alignment = Alignment(horizontal="right")
            tc.border = cell_border
            _num_fmt(tc, total, label)

            for ci, i in enumerate(proj_active, 3):
                yc = ws.cell(row=r, column=ci)
                yc.font = _vf()
                yc.alignment = Alignment(horizontal="right")
                yc.border = cell_border
                val = m["yearly"][i] if i < len(m["yearly"]) else 0
                _num_fmt(yc, val, label)
            r += 1
        return r + 1

    # Projects
    for proj in data.get("projects", []):
        r = write_project(r, proj, color=GOLD)

    # Portfolio Summary
    summary = data.get("summary", [])
    if summary:
        hdr = ws.cell(row=r, column=1, value="Portfolio Summary")
        hdr.font = Font(name="Calibri", bold=True, size=10, color=TEAL)
        hdr.fill = SUMM_FILL
        hdr.border = cell_border
        for ci in range(2, num_cols + 1):
            c = ws.cell(row=r, column=ci)
            c.fill = SUMM_FILL
            c.border = cell_border
        r += 1

        ws.cell(row=r, column=1, value="Category").font = _vf(bold=True, color="8B95A8")
        ws.cell(row=r, column=1).fill = HEADER_FILL
        ws.cell(row=r, column=1).border = cell_border
        ws.cell(row=r, column=2, value="Total").font = _vf(bold=True, color="8B95A8")
        ws.cell(row=r, column=2).fill = HEADER_FILL
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2).border = cell_border
        for ci, yr in enumerate(active_years, 3):
            c = ws.cell(row=r, column=ci, value=yr)
            c.font = _vf(bold=True, color="8B95A8")
            c.fill = HEADER_FILL
            c.alignment = Alignment(horizontal="center")
            c.border = cell_border
        r += 1

        for s in summary:
            lc = ws.cell(row=r, column=1, value=s["label"])
            lc.font = _vf()
            lc.border = cell_border
            tc = ws.cell(row=r, column=2, value=s["total"] if s["total"] else None)
            tc.font = _vf()
            tc.alignment = Alignment(horizontal="right")
            tc.border = cell_border
            if isinstance(s["total"], (int, float)):
                tc.number_format = "#,##0"
            for ci, i in enumerate(active_idxs, 3):
                yc = ws.cell(row=r, column=ci)
                val = s["yearly"][i] if i < len(s["yearly"]) else 0
                yc.value = val if val else None
                yc.font = _vf()
                yc.alignment = Alignment(horizontal="right")
                yc.border = cell_border
                if isinstance(val, (int, float)) and val:
                    yc.number_format = "#,##0"
            r += 1

    # Column widths
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 14
    for ci in range(3, 3 + len(active_years)):
        ws.column_dimensions[get_column_letter(ci)].width = 11

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    from flask import send_file
    return send_file(output, as_attachment=True,
                     download_name="Ember_Project_Returns.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/returns")
@login_required
def returns_report():
    pa = session.get("page_access") or {}
    if not session.get("is_admin") and not pa.get("returns", True):
        return redirect(url_for("home"))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT data, uploaded_at FROM reports WHERE report_type = 'returns' ORDER BY uploaded_at DESC LIMIT 1")
    row = cur.fetchone()
    cur.close(); conn.close()
    data = row["data"] if row else None
    uploaded_at = row["uploaded_at"].strftime("%B %d, %Y") if row else None
    pa = session.get("page_access") or {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True}
    if session.get("is_admin"):
        pa = {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True}
    return render_template("returns.html", data=data, uploaded_at=uploaded_at, is_admin=session.get("is_admin"), page_access=pa)

@app.route("/loans")
@login_required
def loans_report():
    pa = session.get("page_access") or {}
    if not session.get("is_admin") and not pa.get("loans", True):
        return redirect(url_for("home"))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT data, uploaded_at FROM reports WHERE report_type = 'loans' ORDER BY uploaded_at DESC LIMIT 1")
    row = cur.fetchone()
    cur.close(); conn.close()
    data = row["data"] if row else None
    uploaded_at = row["uploaded_at"].strftime("%B %d, %Y") if row else None
    pa = session.get("page_access") or {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True}
    if session.get("is_admin"):
        pa = {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True}
    return render_template("loans.html", data=data, uploaded_at=uploaded_at, is_admin=session.get("is_admin"), page_access=pa)

@app.route("/api/export-operations-excel")
@login_required
def export_operations_excel():
    pa = session.get("page_access") or {}
    if not session.get("is_admin") and not pa.get("operations", True):
        return jsonify({"error": "Access denied"}), 403
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT data, uploaded_at FROM reports WHERE report_type = 'operations' ORDER BY uploaded_at DESC LIMIT 1")
    row = cur.fetchone()
    cur.close(); conn.close()
    if not row or not row["data"]:
        return jsonify({"error": "No data available"}), 404

    data = row["data"]
    uploaded_at = row["uploaded_at"].strftime("%B %d, %Y") if row["uploaded_at"] else ""

    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from io import BytesIO

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Operating Revenues"

    GOLD = "C8A96E"
    HEADER_FILL = PatternFill("solid", fgColor="1E2535")
    TOTALS_FILL = PatternFill("solid", fgColor="161B24")
    thin = Side(style="thin", color="2E3750")
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _hdr_font(bold=False):
        return Font(name="Calibri", size=9, bold=bold, color="8B95A8")

    def _val_font(bold=False):
        return Font(name="Calibri", size=9, bold=bold)

    def _gold_font(bold=True, size=10):
        return Font(name="Calibri", size=size, bold=bold, color=GOLD)

    def write_section(r, title):
        c = ws.cell(row=r, column=1, value=title)
        c.font = _gold_font(size=11)
        return r + 1

    def write_table(r, col_headers, data_rows, totals):
        # Header row
        for ci, h in enumerate(col_headers, 1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font = _hdr_font(bold=True)
            c.fill = HEADER_FILL
            c.alignment = Alignment(horizontal="center" if ci > 1 else "left")
            c.border = cell_border
        r += 1
        # Data rows
        for dr in data_rows:
            for ci, v in enumerate(dr, 1):
                cell = ws.cell(row=r, column=ci, value=v if v != 0 else None)
                cell.font = _val_font()
                cell.border = cell_border
                cell.alignment = Alignment(horizontal="left" if ci == 1 else "right")
                if ci > 1 and isinstance(v, (int, float)) and v:
                    cell.number_format = "#,##0"
            r += 1
        # Totals row
        ws.cell(row=r, column=1, value="Total").font = _val_font(bold=True)
        ws.cell(row=r, column=1).border = cell_border
        ws.cell(row=r, column=1).fill = TOTALS_FILL
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="left")
        for ci, v in enumerate(totals, 2):
            cell = ws.cell(row=r, column=ci, value=v if v else None)
            cell.font = _val_font(bold=True)
            cell.fill = TOTALS_FILL
            cell.border = cell_border
            cell.alignment = Alignment(horizontal="right")
            if isinstance(v, (int, float)):
                cell.number_format = "#,##0"
        return r + 2

    r = 1
    # Title
    ws.cell(row=r, column=1, value="Ember Operating Revenues").font = Font(name="Calibri", bold=True, size=14, color=GOLD)
    r += 1
    ws.cell(row=r, column=1, value=f"Last updated: {uploaded_at}").font = Font(name="Calibri", size=9, color="8B95A8")
    r += 2

    # KPIs
    r = write_section(r, "KPI Summary")
    for kpi in data.get("kpis", []):
        ws.cell(row=r, column=1, value=kpi["label"]).font = _val_font()
        vc = ws.cell(row=r, column=2, value=kpi["value"])
        vc.font = _val_font(bold=True)
        vc.number_format = "#,##0"
        vc.alignment = Alignment(horizontal="right")
        r += 1
    r += 1

    # Annual Forecast
    yr = data.get("yearly_rollup", {})
    if yr.get("years"):
        r = write_section(r, "Annual Revenue Forecast (Next 5 Years)")
        headers = ["Revenue Source"] + [str(y) for y in yr["years"]]
        rows = [[row["label"]] + row["values"] for row in yr.get("rows", [])]
        r = write_table(r, headers, rows, yr.get("totals", []))

    # Monthly Revenue
    mo = data.get("monthly", {})
    if mo.get("dates"):
        r = write_section(r, "Monthly Fee Revenue")
        dates = mo["dates"]
        headers = ["Project / Category"] + [f"{d[5:7]}/{d[2:4]}" for d in dates]
        rows = [[f"{row['project']} — {row['category']}"] + row["values"] for row in mo.get("rows", [])]
        r = write_table(r, headers, rows, mo.get("totals", []))

    # Next 12 Months
    n12 = data.get("next_12_months", {})
    if n12.get("dates"):
        r = write_section(r, "Next 12 Months")
        dates = n12["dates"]
        headers = ["Revenue Source"] + [f"{d[5:7]}/{d[2:4]}" for d in dates]
        rows = [[row["label"]] + row["values"] for row in n12.get("rows", [])]
        r = write_table(r, headers, rows, n12.get("totals", []))

    # Next 12 Quarters
    qr = data.get("quarterly_rollup", {})
    if qr.get("quarters"):
        r = write_section(r, "Next 12 Quarters")
        headers = ["Revenue Source"] + qr["quarters"]
        rows = [[row["label"]] + row["values"] for row in qr.get("rows", [])]
        r = write_table(r, headers, rows, qr.get("totals", []))

    ws.column_dimensions["A"].width = 36
    for ci in range(2, 50):
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(ci)].width = 11

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    from flask import send_file
    return send_file(output, as_attachment=True,
                     download_name="Ember_Operating_Revenues.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/operations")
@login_required
def operations_report():
    pa = session.get("page_access") or {}
    if not session.get("is_admin") and not pa.get("operations", True):
        return redirect(url_for("home"))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT data, uploaded_at FROM reports WHERE report_type = 'operations' ORDER BY uploaded_at DESC LIMIT 1")
    row = cur.fetchone()
    cur.close(); conn.close()
    data = row["data"] if row else None
    uploaded_at = row["uploaded_at"].strftime("%B %d, %Y") if row else None
    pa = session.get("page_access") or {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True}
    if session.get("is_admin"):
        pa = {"mpc_underwriting": True, "returns": True, "loans": True, "operations": True}
    return render_template("operations.html", data=data, uploaded_at=uploaded_at, is_admin=session.get("is_admin"), page_access=pa)

if __name__ == "__main__":
    init_db()
    app.run(debug=True, port=5001)
