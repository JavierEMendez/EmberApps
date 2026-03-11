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
            created_at TIMESTAMP DEFAULT NOW()
        );
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
            return redirect(url_for("index"))
        error = "Invalid username or password."
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ─── MAIN APP ─────────────────────────────────────────────────────────────────
@app.route("/")
@login_required
def index():
    return render_template("app.html", username=session.get("username"), is_admin=session.get("is_admin"))

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
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO projects (name, address, created_by, inputs) VALUES (%s, %s, %s, %s) RETURNING id",
        (name, data.get("address", ""), session["user_id"], json.dumps(default_inputs(name)))
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
    # Run calculation
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
    cur.execute("SELECT id, username, is_admin, created_at FROM users ORDER BY id")
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
    if not username or not password:
        return jsonify({"error": "Username and password required"}), 400
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO users (username, password_hash, is_admin) VALUES (%s, %s, %s) RETURNING id",
            (username, generate_password_hash(password), is_admin)
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

# ─── DEFAULT INPUTS TEMPLATE ──────────────────────────────────────────────────
def default_inputs(name="New Project"):
    return {
        "project_name": name,
        "address": "",
        "gross_acreage": 0,
        "land_escalator": 0.03,
        "purchase_price_per_acre": 0,
        "closing_costs_pct": 0.015,
        "closing_date": "",
        "default_other_pct": 0.17,
        "sectional_other_pct": 0.20,
        "landscaping_other_pct": 0.10,
        "contingency": 0.05,
        "site_work_pct": 0.10,
        "fenced_pct": 0.50,
        "cost_per_mailbox": 300,
        "cost_per_streetlight": 5000,
        "default_start_month": 1,
        "det_storage_rate": 0.5,
        "det_depth": 3,
        "det_num_projects": 1,
        "parks_pct": 0.02,
        "drill_site_acres": 0,
        "commercial_pod_acres": 0,
        "residential_pod_acres": 0,
        "plants": [{"type":"None","notes":""} for _ in range(8)],
        "amenities": [{"type":"None","acres":0} for _ in range(6)],
        "other_netouts": [{"desc":"","acres":0} for _ in range(6)],
        "roads": [{"type":"","lf":0,"width":0,"road_setback":0,"landscaping_setback":0} for _ in range(6)],
        "takedowns": [{"period":1,"pct":1.0}],
        "plant_costs": [{"base_cost":0,"other_pct":0.17,"start_month":1,"ph2_base_cost":0,"ph2_other_pct":0.17,"ph2_start_month":37} for _ in range(8)],
        "amenity_costs": [{"base_cost":0,"other_pct":0.17,"start_month":1} for _ in range(6)],
        "det_costs": [{"base_cost_per_cyd":0,"other_pct":0.17,"landscaping_per_foot":0} for _ in range(6)],
        "other_costs": [{"base_cost":0,"other_pct":0.17,"start_month":1,"duration":1} for _ in range(6)],
        "road_costs": [{"other_pct":0.17,"start_month":1,"landscaping_per_sf":0,"light_spacing":0} for _ in range(6)],
        "lot_sizes": [
            {"on":0,"lot_sf":5000,"depth":120,"yield_per_ac":0,"pace":0,"home_price":0,
             "wsd_per_ff":0,"paving_per_ff":0,"dev_start_month":1,"landscaping_per_lot":0,
             "urd_per_lot":0,"lots_per_streetlight":0,"fence_cost_per_ff":0} for _ in range(16)
        ],
        "timing_method": "1 Takedown",
        "bem_period": 0,
        "bem_pct": 0,
        "brokerage_fees": 0.03,
        "lot_closing_costs": 0.01,
        "take1_pct": 1.0,
        "take2_pct": 0.0,
        "take3_pct": 0.0,
        "ff_year_0":0,"ff_year_1":0,"ff_year_2":0,"ff_year_3":0,"ff_year_4":0,
        "ff_year_5":0,"ff_year_6":0,"ff_year_7":0,"ff_year_8":0,"ff_year_9":0,"ff_year_10":0,
        "res_pods": [{"acres":0,"price_per_acre":0,"closing_costs_pct":0.01,"implied_lots_per_acre":0,"impact_fee_per_lot":0,"sale_period":0} for _ in range(6)],
        "comm_pods": [{"acres":0,"price_per_sf":0,"closing_costs_pct":0.01,"sale_period":0,"av_per_acre":0,"av_delay_months":0} for _ in range(6)],
        "mud_bond": {"amount":0,"reimbursement_pct":0.8,"period":0,"rate":0,"term":0,"annual_payment":0},
        "wcid_bond": {"amount":0,"reimbursement_pct":0.8,"period":0,"rate":0,"term":0,"annual_payment":0},
        "marketing_pct": 0.02,
        "prof_svc_pct": 0.02,
        "dmf_pct": 0.005,
        "personnel_monthly": 0,
        "legal_monthly": 0,
        "mud_monthly": 0,
        "mud_pct": 0,
        "insurance_monthly": 0,
        "bookkeeping_monthly": 0,
    }

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
    """Download all project data as JSON — safe across redeployments."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM projects WHERE id=%s", (pid,))
    proj = dict(cur.fetchone() or {})
    cur.close(); conn.close()
    if not proj:
        return jsonify({"error": "not found"}), 404
    # Convert datetime fields to strings
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
    """Restore a project from a JSON backup file."""
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


if __name__ == "__main__":
    init_db()
    app.run(debug=True, port=5001)
