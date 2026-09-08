"""Acquisitions GIS routes.

A blueprint rather than handlers inline in app.py. The subsystem is ~3,400
lines of engine (acq_gis) plus these handlers, most of them ported from the
standalone app, and they call engine helpers by bare name - `cached_layer_query`,
`_cbas_get_data`, `_overpass_query`. The star import below is what lets those
bodies port across unchanged rather than being rewritten with a prefix on every
call, which is where a port this size would otherwise spend its bugs.

Auth, permissions and persistence are EmberApps'. `init_app` is handed the
callables it needs so this module never imports app.py, which would be circular.
"""

from flask import (Blueprint, request, jsonify, session, redirect,
                   render_template, url_for, send_file, current_app)
import datetime
import re as _re            # explicit: not inherited from the star import
import time                    # explicit: `from acq_gis import *` below
                               # must not be what supplies this name
import io as _io

from acq_gis import *            # noqa: F401,F403 - see the docstring
import acq_gis

# `import *` skips names beginning with an underscore, and most of the engine's
# helpers are private by convention - _overpass_query, _acs_year, _cbas_get_data
# and so on. The ported handlers call them by bare name, so pull them in too.
# Dunders stay out; they would clobber this module's own __name__ and __doc__.
globals().update({k: v for k, v in vars(acq_gis).items()
                  if k.startswith("_") and not k.startswith("__")})
import acq_parcels as parcel_cache
import acq_store

acq_bp = Blueprint("acq", __name__)

# Filled in by init_app so this module never imports app.py.
get_db = None
login_required = None
admin_required = None
_refresh_page_access_from_db = None
_log_activity = None


def init_app(app):
    """Wire the blueprint to the host app's auth, DB and logging."""
    # Any unhandled exception under /api/acq/ must come back as JSON.
    #
    # Flask's default is an HTML error page, so a failure here reached the
    # front end as: Unexpected token '<', "<!doctype "... is not valid JSON.
    # That message says only "something broke" — it hides the traceback, the
    # status code, and which call failed. I spent several rounds guessing at a
    # cause I could have simply read.
    # Registered on the BLUEPRINT, not the app.
    #
    # This was an @app.errorhandler(Exception), which intercepted every
    # exception on every page of the portal — financials, loans, operations —
    # and re-raised from inside the handler for anything outside /api/acq/.
    # An acquisitions concern has no business sitting in front of the rest of
    # the portal's error handling, and re-raising inside a handler is not a
    # safe way to say "not mine".
    #
    # Scoped here it only sees exceptions raised by this blueprint's own views,
    # which is all it ever needed.
    @acq_bp.errorhandler(Exception)
    def _acq_json_errors(e):
        from werkzeug.exceptions import HTTPException
        import traceback
        path = request.path or ""
        if isinstance(e, HTTPException):
            if not path.startswith("/api/acq/"):
                return e            # let pages render the normal error page
            return jsonify({"error": f"{e.code} {e.name}", "detail": e.description,
                            "path": path}), e.code
        traceback.print_exc()
        if not path.startswith("/api/acq/"):
            raise e
        return jsonify({"error": f"{type(e).__name__}: {e}", "path": path}), 500

    global get_db, login_required, admin_required
    global _refresh_page_access_from_db, _log_activity
    get_db = app.config["ACQ_GET_DB"]
    login_required = app.config["ACQ_LOGIN_REQUIRED"]
    admin_required = app.config["ACQ_ADMIN_REQUIRED"]
    _refresh_page_access_from_db = app.config["ACQ_REFRESH_PAGE_ACCESS"]
    _log_activity = app.config["ACQ_LOG_ACTIVITY"]
    app.register_blueprint(acq_bp)



def _admin_required(f):
    """Defer to the host app's admin_required, resolved per request.

    Needed because bootstrapping the parcel cache downloads millions of parcels
    from TxGIO; it was admin-only in the standalone app and stays that way.
    """
    from functools import wraps

    @wraps(f)
    def wrapper(*a, **kw):
        return admin_required(f)(*a, **kw)
    return wrapper


def _login_required(f):
    """Defer to the host app's login_required.

    The real decorator is not available at import time - init_app has not run
    yet - so this wrapper resolves it per request instead of at decoration.
    """
    from functools import wraps

    @wraps(f)
    def wrapper(*a, **kw):
        return login_required(f)(*a, **kw)
    return wrapper

# ══════════════════════════════════════════════════════════════════════════
# ACQUISITIONS GIS
#
# Land screening: search parcels, assemble them into a project, run the
# constraint/yield analysis, pull market and competition context.
#
# The engine lives in acq_gis (live GIS layer queries, geometry, spatial
# enrichment), the statewide parcel cache in acq_parcels, and persistence in
# acq_store. These routes are the thin layer between them and the browser.
#
# Everything namespaces under /api/acq/ deliberately. The standalone app used
# bare /api/projects, which is already MPC Underwriting's in this app - four
# other routes collided the same way (/, /login, /logout, /api/projects/<id>).
# ══════════════════════════════════════════════════════════════════════════

ACQ_PAGE_KEY = "acquisitions"


def _can_view_acquisitions() -> bool:
    """Per-user gate for /acquisitions. Admin overrides; otherwise the user
    must hold page_access.acquisitions.

    Defaults to False. Land screening exposes owner names, mailing addresses
    and appraisal values for parcels the company may be about to approach, so
    it is closer to the financials tier than to macro. Re-reads from the DB so
    an admin grant applies without forcing a re-login."""
    if session.get("is_admin"):
        return True
    pa = _refresh_page_access_from_db()
    return bool(pa.get(ACQ_PAGE_KEY, False))


def _acq_guard():
    """403 JSON for API routes. Returns None when the user may proceed."""
    if not _can_view_acquisitions():
        return jsonify({"error": "forbidden"}), 403
    return None



def _utcnow():
    """UTC now, immune to the star import above.

    acq_gis does `from datetime import datetime`, so `from acq_gis import *`
    rebinds the name `datetime` in this module from the MODULE to the CLASS —
    and every `datetime.datetime.utcnow()` here raised AttributeError. That is
    what surfaced in the browser as "Unexpected token '<'": Flask returned its
    HTML error page, and the front end reported only that the JSON would not
    parse. Creating a project, saving a note, logging outreach — all of it.

    Importing the module inside the function sidesteps the shadowed global
    entirely, so it cannot break again if the imports are reordered.
    """
    import datetime as _dt
    return _dt.datetime.utcnow().isoformat()


def _acq_activity(action, detail=None):
    """Log through the portal's activity log, if it wired one in."""
    try:
        if _log_activity:
            _log_activity(action, detail or {})
    except Exception as e:
        print(f"[acq] activity log failed: {e}", flush=True)


def _acq_owner():
    return session.get("user_id")


def _acq_is_admin():
    return bool(session.get("is_admin"))


@acq_bp.route("/acquisitions")
@_login_required
def acquisitions_page():
    if not _can_view_acquisitions():
        return redirect(url_for("home"))
    # The map page's script needs the tax-rate table and the current user's
    # name. In the standalone app those were Jinja expressions inside the
    # script itself; here the script is a static file, so they are handed over
    # on window.ACQ_* and have to be passed in.
    return render_template(
        "acquisitions.html",
        username=session.get("username"),
        display_name=session.get("display_name", session.get("username")),
        is_admin=session.get("is_admin", False),
        page_access=session.get("page_access") or {},
        tax_rates=_all_tax_rates(),
        tax_year=acq_gis.TAX_YEAR,
        user={"name": session.get("display_name") or session.get("username") or ""},
    )


# ── Geocoding ─────────────────────────────────────────────────────────────

@acq_bp.route("/api/acq/geocode/suggest")
@_login_required
def api_acq_geocode_suggest():
    guard = _acq_guard()
    if guard:
        return guard
    q = (request.args.get("q") or "").strip()
    if len(q) < 3:
        return jsonify({"suggestions": []})
    try:
        return jsonify({"suggestions": geocode_suggest(q)})
    except Exception as e:
        print(f"[acq] geocode suggest failed: {e}", flush=True)
        return jsonify({"suggestions": []})


@acq_bp.route("/api/acq/geocode")
@_login_required
def api_acq_geocode():
    guard = _acq_guard()
    if guard:
        return guard
    q = (request.args.get("q") or "").strip()
    if not q:
        return jsonify({"error": "q required"}), 400
    try:
        hit = geocode(q)
        if not hit:
            return jsonify({"error": "no match"}), 404
        return jsonify(hit)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── Map layers ────────────────────────────────────────────────────────────

@acq_bp.route("/api/acq/load-layer/<key>")
@_login_required
def api_acq_load_layer(key):
    """One ancillary layer for the visible bbox.

    Bbox is required and capped: a whole-state query against FEMA or NWI would
    run for minutes and tie up a worker.
    """
    guard = _acq_guard()
    if guard:
        return guard
    from shapely.geometry import box as shp_box
    try:
        minx = float(request.args["minx"]); miny = float(request.args["miny"])
        maxx = float(request.args["maxx"]); maxy = float(request.args["maxy"])
    except (KeyError, TypeError, ValueError) as e:
        return jsonify({"error": f"bad bbox: {e}"}), 400

    width, height = maxx - minx, maxy - miny
    if width <= 0 or height <= 0:
        return jsonify({"error": "bbox is empty"}), 400
    if width > 5.0 or height > 5.0:
        return jsonify({"error": "bbox too large; zoom in further"}), 400

    bbox_poly = shp_box(minx, miny, maxx, maxy)
    G, E = acq_gis, ENDPOINTS
    fetchers = {
        "counties":     lambda: G.cached_layer_query("counties", bbox_poly),
        "etj":          lambda: G.arcgis_query(E["etj"], bbox=bbox_poly.bounds),
        "schools":      lambda: G.cached_layer_query("tea_schools", bbox_poly),
        "water_dist":   lambda: G.cached_layer_query(
                            "tceq_water", bbox_poly,
                            where="ACCURACY='P' AND STATUS='A'",
                            out_fields="NAME,TYPE,TYPE_DESCRIPTION,COUNTY,DISTRICT_ID,Area_Acres"),
        "muds":         lambda: G.cached_layer_query(
                            "tceq_water", bbox_poly,
                            where="ACCURACY='P' AND STATUS='A' AND TYPE='MUD'",
                            out_fields="NAME,TYPE,TYPE_DESCRIPTION,COUNTY,Area_Acres"),
        "ccn":          lambda: G.cached_layer_query("puc_ccn", bbox_poly,
                            out_fields="UTILITY,CCN_NO,COUNTY,CCN_TYPE"),
        "electric":     lambda: G.cached_layer_query("electric", bbox_poly,
                            out_fields="Utility_Name"),
        "flood":        lambda: G.cached_layer_query(
                            "fema_flood", bbox_poly,
                            where="FLD_ZONE IN ('A','AE','AH','AO','AR','A99','V','VE')",
                            out_fields="FLD_ZONE,ZONE_SUBTY"),
        "wetlands":     lambda: G.cached_wetlands(bbox_poly, (miny + maxy) / 2,
                            (minx + maxx) / 2, max(width, height) * 69),
        "streams":      lambda: G.cached_layer_query("streams", bbox_poly),
        "pipelines":    lambda: G.arcgis_query(E["rrc_pipelines"], geometry_polygon=bbox_poly),
        "transmission": lambda: G.cached_layer_query("transmission", bbox_poly),
        "wells":        lambda: G.arcgis_query(E["rrc_wells"], geometry_polygon=bbox_poly),
        "txdot_projects": lambda: G.cached_layer_query("txdot_projects", bbox_poly,
                            page_size=500, max_pages=4),
    }
    if key not in fetchers:
        return jsonify({"error": f"unknown layer key: {key!r}",
                        "known": sorted(fetchers)}), 400
    try:
        fc = fetchers[key]()
        return jsonify({"key": key, "fc": fc,
                        "count": len(fc.get("features") or [])})
    except Exception as e:
        return jsonify({"error": f"layer fetch failed: {type(e).__name__}: {e}"}), 500


@acq_bp.route("/api/acq/parcels")
@_login_required
def api_acq_parcels():
    """Parcels in the visible map area, for browse mode."""
    guard = _acq_guard()
    if guard:
        return guard
    try:
        minx = float(request.args["minx"]); miny = float(request.args["miny"])
        maxx = float(request.args["maxx"]); maxy = float(request.args["maxy"])
    except (KeyError, TypeError, ValueError) as e:
        return jsonify({"error": f"bad bbox: {e}"}), 400
    if (maxx - minx) > 0.5 or (maxy - miny) > 0.5:
        return jsonify({"type": "FeatureCollection", "features": [],
                        "note": "Zoom in further to load parcels."})
    try:
        return jsonify(arcgis_query(
            ENDPOINTS["parcels"], bbox=(minx, miny, maxx, maxy),
            out_fields="Prop_ID,OWNER_NAME,SITUS_ADDR,MAIL_ADDR,LEGAL_AREA,GIS_AREA,LEGAL_DESC",
            page_size=2000, max_pages=5))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── Owner search ──────────────────────────────────────────────────────────

@acq_bp.route("/api/acq/me")
def api_acq_me():
    """The signed-in user, for the folder-sharing dialog.

    The standalone app had its own /api/me and /api/users. The portal has
    neither — it identifies users through the session and its own users table —
    so the two the ported page needs are provided here rather than reaching
    into app.py, which this module deliberately never imports.
    """
    guard = _acq_guard()
    if guard:
        return guard
    return jsonify({"id": _acq_owner(), "is_admin": bool(_acq_is_admin())})


@acq_bp.route("/api/acq/users")
def api_acq_users():
    """Teammates a folder can be shared with.

    Username only. This list exists to populate a share dialog, so it carries
    nothing about access levels or contact details.
    """
    guard = _acq_guard()
    if guard:
        return guard
    conn = get_db()
    try:
        cur = conn.cursor()
        cur.execute("SELECT id, username FROM users ORDER BY username")
        rows = cur.fetchall()
        users = [{"id": r["id"], "name": r["username"]} if isinstance(r, dict)
                 else {"id": r[0], "name": r[1]} for r in rows]
        return jsonify({"users": users})
    finally:
        try:
            conn.close()
        except Exception:
            pass


@acq_bp.route("/api/acq/owner-related")
@_login_required
def api_acq_owner_related():
    """Owners sharing a mailing address with this one.

    Stands in for a registered-agent search, which Texas does not publish:
    the SOS keeps agent data behind SOSDirect and the Comptroller's franchise
    file has no agent field. A shared mailbox is evidence of a relationship,
    not proof of one, so the response carries the addresses it used and
    anything it skipped alongside the owners it found.
    """
    guard = _acq_guard()
    if guard:
        return guard
    name = (request.args.get("name") or "").strip()
    if not name:
        return jsonify({"error": "name required"}), 400
    try:
        return jsonify(parcel_cache.find_related_owners(name))
    except Exception as e:
        return jsonify({"error": str(e), "owners": []}), 200


@acq_bp.route("/api/acq/owner-dossier")
@_login_required
def api_acq_owner_dossier():
    """Every cached parcel belonging to one owner.

    Fuzzy matching only for names carrying two or more distinctive tokens.
    The caller always passes a full OWNER_NAME off a parcel popup, never a
    typed fragment, so a bare one-token query staying exact-only costs nothing
    and keeps "EMBER" from dragging in every EMBERG in the state.
    """
    guard = _acq_guard()
    if guard:
        return guard
    name = (request.args.get("name") or "").strip()
    if not name:
        return jsonify({"error": "name required"}), 400
    exact_only = request.args.get("exact", "0") == "1"
    # The owner panel asks for geometry=1 because it draws the holding on the
    # map and analyses it. This endpoint ignored the flag, so every parcel came
    # back without a boundary and "Analyse all" reported "no tracts with
    # boundaries to analyse" for an owner whose tracts are all perfectly well
    # defined. Boundaries stay opt-in: a 500-parcel holding is megabytes of
    # rings and the list itself only needs centroids.
    want_geom = request.args.get("geometry", "0") == "1"
    tokens = parcel_cache._owner_tokens(name)
    used_fuzzy = False
    if exact_only or len(tokens) < 2:
        result = parcel_cache.find_parcels_by_owner(
            name, exact=True, limit=500, include_geometry=want_geom)
    else:
        result = parcel_cache.find_parcels_by_owner(
            name, exact=False, limit=500, include_geometry=want_geom)
        used_fuzzy = True
    passes = {p.get("match_pass") for p in result.get("parcels") or []}
    result["match_mode"] = ("fuzzy" if used_fuzzy and not (passes and passes <= {"exact"})
                            else "exact")
    result["query"] = name
    return jsonify(result)


@acq_bp.route("/api/acq/parcel-detail/<prop_id>")
@_login_required
def api_acq_parcel_detail(prop_id):
    guard = _acq_guard()
    if guard:
        return guard
    try:
        return jsonify(parcel_detail(prop_id))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── Saved objects (projects, pins, folders, notes, favorites) ─────────────

def _acq_list(kind):
    conn = get_db()
    try:
        return acq_store.list_objects(conn, kind, _acq_owner(), _acq_is_admin())
    finally:
        conn.close()


def _acq_save(kind, obj):
    conn = get_db()
    try:
        return acq_store.put_object(conn, kind, obj, _acq_owner())
    finally:
        conn.close()


@acq_bp.route("/api/acq/projects", methods=["GET"])
@_login_required
def api_acq_projects_list():
    guard = _acq_guard()
    if guard:
        return guard
    projects = _acq_list("project")
    # A one-off "quick analysis" is not a project you meant to keep. Both come
    # back, split, so the page can list intentional work first and park the
    # rest in history.
    return jsonify({
        "projects": [p for p in projects if not p.get("is_quick_analysis")],
        "history":  [p for p in projects if p.get("is_quick_analysis")],
    })


@acq_bp.route("/api/acq/projects", methods=["POST"])
@_login_required
def api_acq_projects_create():
    guard = _acq_guard()
    if guard:
        return guard
    body = request.get_json(silent=True) or {}
    name = (body.get("name") or "").strip()
    if not name:
        return jsonify({"error": "name required"}), 400
    proj = {
        "name": name,
        "tracts": body.get("tracts") or [],
        # The page sends project_kind:'quick_analysis' for a one-off run and
        # is_user_project:false alongside it; accept either rather than
        # silently filing every quick analysis as a real project.
        "is_quick_analysis": bool(
            body.get("is_quick_analysis")
            or body.get("project_kind") == "quick_analysis"
            or body.get("is_user_project") is False),
        "created_at": _utcnow(),
    }
    proj["total_acres"] = round(
        sum(float(t.get("acres") or 0) for t in proj["tracts"]), 2)
    saved = _acq_save("project", proj)
    # Return the project both nested and flat: quickAnalyzeTract reads d.id
    # while analyzeOwnerHolding reads (j.project || j).id, and a project that
    # saves fine but opens /acquisitions/project/undefined is indistinguishable
    # from one that failed.
    out = dict(saved)
    out["project"] = saved
    return jsonify(out), 201


@acq_bp.route("/api/acq/projects/<pid>", methods=["GET"])
@_login_required
def api_acq_project_get(pid):
    guard = _acq_guard()
    if guard:
        return guard
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, _acq_owner(), _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404
    # Overrides live in their own table, not on the project, so a tract added
    # before the acreage was corrected still carries the old figure. Apply
    # them on the way out and the card, the analysis and the report agree.
    proj["tracts"], _ = parcel_cache.hydrate_tract_acres(proj.get("tracts") or [])
    proj["total_acres"] = round(
        sum(float(t.get("acres") or 0) for t in proj["tracts"]), 2)
    return jsonify({"project": proj})


@acq_bp.route("/api/acq/projects/<pid>", methods=["PATCH"])
@_login_required
def api_acq_project_patch(pid):
    guard = _acq_guard()
    if guard:
        return guard
    body = request.get_json(silent=True) or {}
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, _acq_owner(), _acq_is_admin())
        if not proj:
            return jsonify({"error": "project not found"}), 404
        for k in ("name", "tracts", "yield_assumptions", "netout_assumptions",
                  "netout_overrides", "notes", "is_quick_analysis"):
            if k in body:
                proj[k] = body[k]
        if "tracts" in body:
            proj["total_acres"] = round(
                sum(float(t.get("acres") or 0) for t in proj.get("tracts") or []), 2)
        saved = acq_store.put_object(conn, "project", proj, proj.get("_owner_id") or _acq_owner())
    finally:
        conn.close()
    return jsonify({"project": saved})


@acq_bp.route("/api/acq/projects/<pid>", methods=["DELETE"])
@_login_required
def api_acq_project_delete(pid):
    guard = _acq_guard()
    if guard:
        return guard
    conn = get_db()
    try:
        ok = acq_store.delete_object(conn, "project", pid, _acq_owner(), _acq_is_admin())
    finally:
        conn.close()
    return jsonify({"deleted": ok}), (200 if ok else 404)


@acq_bp.route("/api/acq/tract-pins", methods=["GET", "POST"])
@_login_required
def api_acq_tract_pins():
    guard = _acq_guard()
    if guard:
        return guard
    if request.method == "GET":
        return jsonify({"pins": _acq_list("tract_pin")})
    body = request.get_json(silent=True) or {}
    if not (body.get("prop_id") or "").strip():
        return jsonify({"error": "prop_id required"}), 400
    return jsonify({"pin": _acq_save("tract_pin", body)}), 201


@acq_bp.route("/api/acq/tract-pins/<pin_id>", methods=["DELETE"])
@_login_required
def api_acq_tract_pin_delete(pin_id):
    guard = _acq_guard()
    if guard:
        return guard
    conn = get_db()
    try:
        ok = acq_store.delete_object(conn, "tract_pin", pin_id, _acq_owner(), _acq_is_admin())
    finally:
        conn.close()
    return jsonify({"deleted": ok}), (200 if ok else 404)


@acq_bp.route("/api/acq/folders", methods=["GET", "POST"])
@_login_required
def api_acq_folders():
    guard = _acq_guard()
    if guard:
        return guard
    if request.method == "GET":
        return jsonify({"folders": _acq_list("folder")})
    body = request.get_json(silent=True) or {}
    name = (body.get("name") or "").strip()
    if not name:
        return jsonify({"error": "name required"}), 400
    return jsonify({"folder": _acq_save("folder", {"name": name,
                                                    "color": body.get("color")})}), 201


@acq_bp.route("/api/acq/notes/by-props", methods=["POST"])
@_login_required
def api_acq_notes_by_props():
    """Note counts for a screenful of parcels, so the map can flag which ones
    already carry a note without one request per parcel."""
    guard = _acq_guard()
    if guard:
        return guard
    prop_ids = (request.get_json(silent=True) or {}).get("prop_ids") or []
    conn = get_db()
    try:
        found = acq_store.find_by_prop(conn, "note", prop_ids,
                                        _acq_owner(), _acq_is_admin())
    finally:
        conn.close()
    return jsonify({"counts": {k: len(v) for k, v in found.items()}})


@acq_bp.route("/api/acq/notes/<prop_id>", methods=["GET", "POST"])
@_login_required
def api_acq_notes(prop_id):
    guard = _acq_guard()
    if guard:
        return guard
    conn = get_db()
    try:
        if request.method == "GET":
            found = acq_store.find_by_prop(conn, "note", [prop_id],
                                            _acq_owner(), _acq_is_admin())
            return jsonify({"notes": found.get(str(prop_id), [])})
        body = request.get_json(silent=True) or {}
        text = (body.get("text") or "").strip()
        if not text:
            return jsonify({"error": "text required"}), 400
        note = acq_store.put_object(conn, "note", {
            "prop_id": str(prop_id),
            "text": text,
            "author": session.get("username"),
            "created_at": _utcnow(),
        }, _acq_owner())
    finally:
        conn.close()
    return jsonify({"note": note}), 201


# Distinct path rather than /api/acq/notes/<note_id>: that collides with the
# GET/POST rule above, which takes a prop_id. Werkzeug does disambiguate them
# by method, but two rules sharing a path with different meanings for the same
# segment is a trap for whoever adds a method next.
@acq_bp.route("/api/acq/notes/id/<note_id>", methods=["DELETE"])
@_login_required
def api_acq_note_delete(note_id):
    guard = _acq_guard()
    if guard:
        return guard
    conn = get_db()
    try:
        ok = acq_store.delete_object(conn, "note", note_id, _acq_owner(), _acq_is_admin())
    finally:
        conn.close()
    return jsonify({"deleted": ok}), (200 if ok else 404)


@acq_bp.route("/api/acq/search", methods=["POST"])
@_login_required
def api_acq_search():
    """Run a parcel search. Thin wrapper over the engine, as in the standalone.

    This used to reimplement the search — its own buffering, coverage check and
    parcel query — and returned a flat {count, total_acres, tracts} instead of
    the engine's {layers, summary}. The ported front end reads
    data.summary.tracts, so every search failed with "Cannot read properties of
    undefined". Reimplementing an engine the module already imports is how the
    contract drifted; calling run_search is what keeps it from drifting again.
    """
    guard = _acq_guard()
    if guard:
        return guard
    p = request.get_json(force=True) or {}
    polygon_geojson = p.get("polygon")
    try:
        # With a polygon, lat/lon are optional — the centroid is derived.
        lat = float(p.get("lat") or 0) if polygon_geojson else float(p["lat"])
        lon = float(p.get("lon") or 0) if polygon_geojson else float(p["lon"])
        radius_mi = float(p.get("radius_mi", 10))
        min_acres = float(p.get("min_acres", 300))
        max_acres = float(p.get("max_acres", 100000))
    except (KeyError, TypeError, ValueError) as e:
        return jsonify({"error": f"Bad input: {e}"}), 400

    try:
        result = run_search(lat, lon, radius_mi, min_acres, max_acres,
                            polygon_geojson=polygon_geojson)
        annotate_tracts_with_enrichment(result["layers"])
        _last_search_cache["data"] = result["layers"]
        _last_search_cache["saved_at"] = time.time()
        _last_search_cache["meta"] = {"center": [lat, lon], "radius_mi": radius_mi,
                                      "min_acres": min_acres, "max_acres": max_acres,
                                      "polygon": polygon_geojson,
                                      "label": p.get("label") or "search"}
        _acq_activity("acq_search", {"center": [lat, lon], "radius_mi": radius_mi,
                                     "min_acres": min_acres, "max_acres": max_acres,
                                     "tracts": result["summary"]["tracts"],
                                     "polygon": bool(polygon_geojson)})
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@acq_bp.route("/api/acq/projects/<pid>/analyze", methods=["POST"])
@_login_required
def api_acq_project_analyze(pid):
    """Run the constraint and yield analysis over the project's tracts.

    Slow by nature - it queries FEMA, USFWS, USGS and the RRC live, and a large
    assemblage against a busy wetlands service can take the better part of a
    minute. gunicorn runs with threads for this reason; see gunicorn.conf.py.
    """
    guard = _acq_guard()
    if guard:
        return guard
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, _acq_owner(), _acq_is_admin())
        if not proj:
            return jsonify({"error": "project not found"}), 404
        try:
            analysis = run_analysis(proj)
        except ValueError as e:
            return jsonify({"error": str(e)}), 400
        except Exception as e:
            print(f"[acq-analyze] {pid} failed: {type(e).__name__}: {e}", flush=True)
            return jsonify({"error": f"analysis failed: {type(e).__name__}: {e}"}), 500

        # Cache on the project so reopening the page is instant.
        proj["analysis_cache"] = analysis
        proj["updated_at"] = analysis.get("computed_at")
        acq_store.put_object(conn, "project", proj,
                             proj.get("_owner_id") or _acq_owner())
    finally:
        conn.close()

    # activity_log's second column is `path`, so pass one - the acreage detail
    # rides along after it rather than pretending to be a path of its own.
    _log_activity("acq_project_analyze",
                  f"/acquisitions/project/{pid} "
                  f"({analysis.get('gross_acres')} gross ac -> "
                  f"{analysis.get('net_saleable_acres')} saleable ac)")
    return jsonify({"ok": True, "analysis": analysis})


@acq_bp.route("/acquisitions/project/<pid>")
@_login_required
def acquisitions_project_page(pid):
    if not _can_view_acquisitions():
        return redirect(url_for("home"))
    return render_template(
        "acquisitions_project.html",
        project_id=pid,
        username=session.get("username"),
        display_name=session.get("display_name", session.get("username")),
        is_admin=session.get("is_admin", False),
        page_access=session.get("page_access") or {},
    )


# Status is polled every 5s by the admin page and costs a multi-second scan of
# a 3M-row table, so overlapping polls share one result. Bootstraps run for
# hours; a figure a few seconds old is not misleading. `?fresh=1` bypasses it.
_STATUS_TTL = 15.0
_CACHE_STATUS_MEMO = {"at": 0.0, "payload": None}


@acq_bp.route("/api/acq/parcel/<prop_id>/acres", methods=["POST", "DELETE"])
@_login_required
def api_acq_parcel_acres(prop_id):
    """Set or clear a hand-entered acreage for one parcel.

    Beats both StratMap and the appraisal district everywhere the app shows
    acreage. Kept in its own table so a county re-bootstrap cannot erase it.
    """
    guard = _acq_guard()
    if guard:
        return guard
    body = request.get_json(silent=True) or {}
    fips = str(body.get("county_fips") or request.args.get("county_fips") or "").strip()
    if not fips:
        return jsonify({"error": "county_fips required"}), 400
    acres = None if request.method == "DELETE" else body.get("acres")
    who = (session.get("display_name") or session.get("username") or "")[:60]
    res = parcel_cache.set_acreage_override(prop_id, fips, acres,
                                            note=(body.get("note") or None),
                                            set_by=who)
    if res.get("error"):
        return jsonify(res), 400
    _acq_log("acreage_override", {"prop_id": prop_id, "acres": res.get("acres")})
    return jsonify(res)


@acq_bp.route("/api/acq/parcel-overrides")
@_login_required
def api_acq_parcel_overrides():
    """Every hand-entered acreage, with what the cache would otherwise say."""
    guard = _acq_guard()
    if guard:
        return guard
    return jsonify({"overrides": parcel_cache.list_acreage_overrides()})


@acq_bp.route("/api/acq/cache/reconcile-cad", methods=["POST"])
@_login_required
def api_acq_cache_reconcile_cad():
    """Store the appraisal district's own acreage for every sizeable parcel.

    StratMap draws some accounts as the whole abstract rather than the tract --
    3.5% of sampled Harris parcels over 90 acres, always oversized, up to 8x.
    This pulls HCAD's figure for anything over five acres so that every
    acreage the app shows is the district's, not a polygon that may include
    land the parcel does not.
    """
    guard = _acq_guard()
    if guard:
        return guard
    if not _acq_is_admin():
        return jsonify({"error": "admin only"}), 403
    fips = (request.args.get("county_fips") or "48201").strip()
    try:
        min_ac = float(request.args.get("min_acres") or 5)
    except ValueError:
        min_ac = 5.0
    try:
        return jsonify(parcel_cache.reconcile_cad_acres(fips, min_ac))
    except Exception as e:
        print(f"[cad] reconcile failed: {e}", flush=True)
        return jsonify({"error": str(e)}), 500


@acq_bp.route("/api/acq/cache/status")
@_login_required
def api_acq_cache_status():
    """Parcel-cache coverage. Surfaces on the page so an empty or stale cache
    reads as a known state rather than as the map being broken."""
    guard = _acq_guard()
    if guard:
        return guard
    # cache_status() returns a bare list. The admin panel — and the standalone
    # app it came from — expect it under `counties`, alongside the orphan count
    # and any bootstrap in flight. Returning the raw list made the panel render
    # an empty table against a fully populated cache.
    # The page polls this every 5s while a bootstrap runs, so it has to be
    # cheap. It was not: cache_status() already runs a per-county anti-join for
    # its `unindexed` figure, and this handler then called rtree_missing_count()
    # which runs the SAME anti-join across the whole table -- 47s on its own
    # against 3M parcels, on top of cache_status()'s own 8s. At ~58s a response
    # and a poll every 5s, a dozen of these were in flight at once against two
    # workers of four threads, which starved every other request in the app and
    # eventually took longer than gunicorn's timeout: the worker was killed and
    # Railway answered "upstream error", which is not JSON and so surfaced as a
    # parse error rather than as a timeout.
    #
    # The total is just the sum of the per-county numbers already computed, so
    # the second scan bought nothing. The memo below then keeps overlapping
    # polls off the database entirely; a bootstrap takes hours, so a status
    # figure up to STATUS_TTL seconds old is not misleading.
    now = time.time()
    fresh = request.args.get("fresh") in ("1", "true", "yes")
    memo = _CACHE_STATUS_MEMO
    if not fresh and memo["payload"] is not None and now - memo["at"] < _STATUS_TTL:
        out = dict(memo["payload"])
        out["in_progress"] = dict(_cache_bootstrap_status)   # always live
        out["cached_for"] = round(now - memo["at"], 1)
        return jsonify(out)

    try:
        counties = parcel_cache.cache_status()
    except Exception as e:
        return jsonify({"error": str(e), "counties": []}), 200

    try:
        orphans = parcel_cache.rtree_orphan_count()
    except Exception:
        orphans = 0
    unindexed = sum(int(c.get("unindexed") or 0) for c in counties)

    payload = {
        "counties": counties,
        "rtree_orphans": orphans,
        "rtree_unindexed": unindexed,
    }
    memo["at"], memo["payload"] = now, payload
    out = dict(payload)
    out["in_progress"] = dict(_cache_bootstrap_status)
    out["cached_for"] = 0
    return jsonify(out)


# ══════════════════════════════════════════════════════════════════════════
# JSON-store compatibility shim
#
# Eight ported handlers - corridor search and import, the tract page, the
# outreach campaign PDF, rebuffer - talk to the standalone app's storage
# directly: `with _storage_lock: s = _load_storage()`, mutate `s["searches"]`,
# `_save_storage(s)`. Rewriting each against acq_store would mean rewriting
# their bodies, which is where a port this size goes wrong.
#
# So the old interface is presented over Postgres instead. `_load_storage`
# builds the dict-of-lists those handlers expect; `_save_storage` writes back
# only what actually changed, compared against a snapshot taken at load, so a
# handler touching one saved search does not rewrite every row.
#
# What it deliberately does NOT support is deletion by omission - removing an
# item from one of the lists and saving. No ported handler does that (they use
# the DELETE routes), and silently honouring it would make an accidental list
# rebuild destructive.
# ══════════════════════════════════════════════════════════════════════════

import json as _json
import threading as _threading

_STORE_KINDS = {
    "projects": "project", "searches": "search", "folders": "folder",
    "tract_pins": "tract_pin", "notes": "note", "polygons": "polygon",
    "favorites": "favorite", "outreach": "outreach",
}

# Request-scoped snapshot of what _load_storage handed out, so _save_storage
# can tell what moved. Thread-local because gunicorn runs threaded.
_shim_state = _threading.local()


class _NullLock:
    """The standalone app serialised writes to one JSON file. Postgres does
    its own concurrency control, so this is a no-op that keeps `with
    _storage_lock:` reading naturally in the ported bodies."""

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


_storage_lock = _NullLock()


def _current_user():
    return {"id": session.get("user_id"),
            "name": session.get("display_name") or session.get("username"),
            "username": session.get("username"),
            "is_admin": bool(session.get("is_admin"))}


def _load_storage():
    uid = session.get("user_id")
    admin = bool(session.get("is_admin"))
    out, snap = {}, {}
    conn = get_db()
    try:
        for plural, kind in _STORE_KINDS.items():
            rows = acq_store.list_objects(conn, kind, uid, admin)
            for r in rows:
                # Ported bodies read owner_id, not the store's _owner_id.
                r["owner_id"] = r.get("_owner_id")
            out[plural] = rows
            snap[plural] = {r["id"]: _json.dumps(r, sort_keys=True, default=str)
                            for r in rows if r.get("id")}
        # Users come from EmberApps, not from this tab. The outreach campaign
        # PDF looks up author names against this.
        users = []
        try:
            cur = conn.cursor()
            cur.execute("SELECT id, username, display_name FROM users")
            users = [{"id": r["id"], "username": r["username"],
                      "name": r.get("display_name") or r["username"]}
                     for r in cur.fetchall()]
            cur.close()
        except Exception as e:
            print(f"[acq-shim] user lookup failed: {e}", flush=True)
        out["users"] = users
    finally:
        conn.close()
    _shim_state.snapshot = snap
    return out


def _save_storage(s):
    snap = getattr(_shim_state, "snapshot", None) or {}
    uid = session.get("user_id")
    written = 0
    conn = get_db()
    try:
        for plural, kind in _STORE_KINDS.items():
            for obj in s.get(plural) or []:
                oid = obj.get("id")
                now = _json.dumps(obj, sort_keys=True, default=str)
                if oid and snap.get(plural, {}).get(oid) == now:
                    continue                      # untouched
                acq_store.put_object(conn, kind, obj, obj.get("owner_id") or uid)
                written += 1
    finally:
        conn.close()
    return written



def _acq_log(event, detail=None):
    """The standalone app's _log_action, mapped onto EmberApps' activity_log."""
    try:
        _log_activity(f"acq_{event}", str(detail)[:400] if detail else None)
    except Exception as e:
        print(f"[acq] activity log failed ({event}): {e}", flush=True)



# ══════════════════════════════════════════════════════════════════════════
# Market, competition and reports
#
# Ported from the standalone app. The bodies are unchanged; only the preamble
# differs - the project comes out of Postgres via acq_store rather than out of
# one JSON file under a file lock.
#
# These are the slow endpoints. CBAS, Overpass, the Census and FRED are all
# remote, and several fan out across mirrors. gunicorn runs with threads for
# this reason; see gunicorn.conf.py.
# ══════════════════════════════════════════════════════════════════════════


@acq_bp.route("/api/acq/tract-sheet/<prop_id>")
@_login_required
def acq_api_tract_sheet(prop_id):
    """Build a one-page PDF for one tract. Tries the current search cache first
    (has all the spatial enrichment); falls back to a live StratMap + HCAD/MCAD
    pull so the report works even after a server restart or for a tract the user
    hasn't run a search for yet."""
    pid = (prop_id or "").strip()
    if not pid:
        return jsonify({"error": "prop_id required"}), 400

    tract = None
    # 1. Try the last-search cache — already has flood %, wells, MUD, etc.
    cache = _last_search_cache.get("data") or {}
    features = (cache.get("tracts") or {}).get("features") or []
    tract = next((f for f in features
                  if str((f.get("properties") or {}).get("Prop_ID") or "").strip() == pid), None)

    # 2. Fallback — look up in the local parcel cache (instant by Prop_ID).
    if not tract:
        try:
            import acq_parcels as parcel_cache
            tract = parcel_cache.find_parcel_by_pid(pid)
        except Exception as e:
            print(f"[tract-sheet] cache lookup failed: {e}", flush=True)

    if not tract:
        return jsonify({"error": f"No parcel found for Prop_ID {pid} in StratMap or local cache."}), 404

    # Refresh CAD values for the sheet even when the tract came from the
    # last-search cache. That cache may have been built before appraised/taxable
    # values were part of the HCAD overlay.
    tmp_fc = {"type": "FeatureCollection", "features": [tract]}
    try: hcad_live_overlay(tmp_fc)
    except Exception as e: print(f"[tract-sheet] hcad overlay failed: {e}", flush=True)
    try: mcad_live_overlay(tmp_fc)
    except Exception as e: print(f"[tract-sheet] mcad overlay failed: {e}", flush=True)
    _enrich_tract_sheet_feature(tract)

    meta = _last_search_cache.get("meta") or {"label": f"tract_{pid}"}
    buf = build_tract_sheet_pdf(tract, meta)
    safe_pid = "".join(c if c.isalnum() else "_" for c in pid)
    owner_slug = (tract.get("properties", {}).get("OWNER_NAME") or "")
    owner_slug = "".join(c if c.isalnum() else "_" for c in owner_slug)[:30].strip("_")
    fname = f"ember_tract_{safe_pid}{('_' + owner_slug) if owner_slug else ''}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    _acq_log("export_tract_sheet", {"prop_id": pid})
    return send_file(buf, mimetype="application/pdf",
                      as_attachment=True, download_name=fname)


@acq_bp.route("/api/acq/projects/<pid>/cbas", methods=["GET"])
@_login_required
def acq_api_projects_cbas(pid):
    """CBAS competitor survey for everything within `radius_mi` of the project.

    Communities are filtered by true great-circle distance from the project
    centroid using CBAS's own coordinates, so the ring is centred on the deal
    rather than approximated by ZIP boundaries, and builders come back by name.

    Returns, for the ring:
      communities   — each with distance/direction, builder names, lot widths
                      (FF), price range, absorption, VDL, futures, % built out
      builders      — lots under control, communities, lot widths, avg price,
                      avg sqft, $/sf, estimated annual starts
      lot_bands     — supply and pricing by lot width (the product read)
      quarter_series— starts/closings by quarter, summed across the ring
    """
    import math as _m
    from shapely.geometry import shape as shp_shape
    from shapely.ops import unary_union

    token = os.environ.get("CBAS_TOKEN", "").strip()
    if not token:
        return jsonify({"error": "CBAS_TOKEN not configured on the server. "
                                 "Add it in Railway Variables to enable competitor data."}), 503

    guard = _acq_guard()
    if guard:
        return guard
    uid = _acq_owner()
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, uid, _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404

    geoms = []
    for t in proj.get("tracts") or []:
        g = t.get("geometry")
        if g:
            try:
                geoms.append(shp_shape(g))
            except Exception:
                pass
    if not geoms:
        return jsonify({"error": "no tract geometries"}), 400
    proj_union = unary_union(geoms)
    centroid = proj_union.centroid
    c_lat, c_lon = centroid.y, centroid.x

    try:
        radius_mi = float(request.args.get("radius_mi", "8"))
    except (TypeError, ValueError):
        radius_mi = 8.0
    radius_mi = max(1.0, min(radius_mi, 25.0))

    entries, latest, err = _cbas_get_data(token)
    if err:
        return jsonify({"error": err}), 502
    bmap = _cbas_builder_map(entries)

    _COMPASS = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
                "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]

    def _dist_mi(la, lo):
        R = 3958.8
        p1, p2 = _m.radians(c_lat), _m.radians(la)
        dp, dl = _m.radians(la - c_lat), _m.radians(lo - c_lon)
        x = _m.sin(dp / 2) ** 2 + _m.cos(p1) * _m.cos(p2) * _m.sin(dl / 2) ** 2
        return 2 * R * _m.asin(_m.sqrt(x))

    def _dir(la, lo):
        dy = la - c_lat
        dx = (lo - c_lon) * _m.cos(_m.radians(c_lat))
        ang = (_m.degrees(_m.atan2(dx, dy)) + 360) % 360
        return _COMPASS[int((ang + 11.25) // 22.5) % 16]

    # --- 1) The ring -------------------------------------------------------
    near = []
    for e in entries:
        la, lo = e.get("lat"), e.get("lng")
        if not la or not lo:
            continue
        try:
            d_mi = _dist_mi(float(la), float(lo))
        except Exception:
            continue
        if d_mi <= radius_mi:
            near.append((d_mi, e))
    near.sort(key=lambda t: t[0])

    # Every school district the project POLYGON touches, so in-district
    # competition sorts first. A point-in-polygon test on the centroid was wrong
    # for any tract straddling a boundary — it picked one district and flagged
    # the other half of the project's own market as out-of-district. Query the
    # bbox across all three district layers (Unified / Secondary / Elementary)
    # and keep every one that actually intersects the tract.
    dnames = []
    try:
        import requests as _rq
        b = proj_union.bounds
        for layer_id in (0, 1, 2):
            try:
                sd = _rq.get("https://tigerweb.geo.census.gov/arcgis/rest/services/"
                             f"TIGERweb/School/MapServer/{layer_id}/query",
                             params={"geometry": f"{b[0]},{b[1]},{b[2]},{b[3]}",
                                     "geometryType": "esriGeometryEnvelope", "inSR": "4326",
                                     "spatialRel": "esriSpatialRelIntersects",
                                     "outFields": "NAME", "returnGeometry": "true",
                                     "outSR": "4326", "f": "geojson"},
                             timeout=20)
                if sd.status_code != 200:
                    continue
                for feat in (sd.json().get("features") or []):
                    g = feat.get("geometry")
                    nm = (feat.get("properties") or {}).get("NAME")
                    if not g or not nm or nm in dnames:
                        continue
                    try:
                        if shp_shape(g).intersects(proj_union):
                            dnames.append(nm)
                    except Exception:
                        continue
            except Exception:
                continue
    except Exception:
        pass
    dname = " / ".join(dnames) if dnames else None

    def _norm_d(x):
        return "".join(ch for ch in (x or "").upper() if ch.isalnum()) \
            .replace("INDEPENDENTSCHOOLDISTRICT", "ISD").replace("SCHOOLDISTRICT", "ISD")

    proj_ds = {_norm_d(n) for n in dnames if n}

    def _cbas_num(v):
        try:
            if v in (None, ""):
                return None
            return float(v)
        except (TypeError, ValueError):
            return None

    def _price_stats(prices, sqfts, ppsfs=None):
        pr = [p for p in (_cbas_num(x) for x in prices) if p and p > 0]
        sf = [s for s in (_cbas_num(x) for x in sqfts) if s and s > 0]
        ppsf = [p for p in (_cbas_num(x) for x in (ppsfs or [])) if p and p > 0]
        if not ppsf:
            ppsf = [p / s for p, s in zip(pr, sf) if s]
        return {
            "avg_price": round(sum(pr) / len(pr)) if pr else None,
            "min_price": min(pr) if pr else None,
            "max_price": max(pr) if pr else None,
            "avg_sqft": round(sum(sf) / len(sf)) if sf else None,
            "avg_ppsf": round(sum(ppsf) / len(ppsf), 2) if ppsf else None,
            "plans": len(pr),
        }

    def _community_product_detail(e, ann_starts, ann_closings, vdls, futures):
        """Per-community lot-width and builder detail for the UI expand row.

        CBAS gives lot counts by builder/width, and plan pricing by builder/width.
        VDL/futures are community-level, so those are apportioned by lot share.
        """
        lot_rows = {}
        builder_rows = {}
        builder_lot_rows = {}
        section_lots = 0

        def _builder_name(bid, fallback=None):
            nm = str(fallback or bmap.get(bid) or f"Builder #{bid}").strip()
            return None if nm == "Builder TBD" else nm

        def _lot(ff):
            if ff not in lot_rows:
                lot_rows[ff] = {"lot_width_ff": ff, "lots": 0, "builders": {},
                                "prices": [], "sqfts": [], "ppsfs": []}
            return lot_rows[ff]

        def _builder(bid, name):
            key = bid if bid is not None else name
            if key not in builder_rows:
                builder_rows[key] = {"builder_id": bid, "name": name, "lots": 0,
                                     "lot_types": set(), "prices": [], "sqfts": [],
                                     "ppsfs": []}
            return builder_rows[key]

        def _builder_lot(bid, name, ff):
            key = (bid if bid is not None else name, ff)
            if key not in builder_lot_rows:
                builder_lot_rows[key] = {"builder_id": bid, "name": name,
                                         "lot_width_ff": ff, "lots": 0,
                                         "prices": [], "sqfts": [], "ppsfs": []}
            return builder_lot_rows[key]

        for sec in (e.get("sections") or []):
            for lb in (sec.get("lot_types_builders") or []):
                bid = lb.get("builder")
                if bid is None or bid in _CBAS_PLACEHOLDER_BUILDERS:
                    continue
                ff = _cbas_ff(lb.get("lot_type"))
                if not ff:
                    continue
                lots = _cbas_num(lb.get("num_lots")) or 0
                if lots <= 0:
                    continue
                name = _builder_name(bid)
                if not name:
                    continue
                section_lots += lots
                lrow = _lot(ff)
                lrow["lots"] += lots
                lrow["builders"][name] = lrow["builders"].get(name, 0) + lots
                brow = _builder(bid, name)
                brow["lots"] += lots
                brow["lot_types"].add(ff)
                blrow = _builder_lot(bid, name, ff)
                blrow["lots"] += lots

        for fp in ((e.get("latestFloorplanPricing") or {}).get("entries") or []):
            ff = _cbas_ff(fp.get("lot_type"))
            if not ff:
                continue
            bid = fp.get("builderID")
            if bid in _CBAS_PLACEHOLDER_BUILDERS:
                continue
            name = _builder_name(bid, fp.get("builderName"))
            if not name:
                continue
            price = _cbas_num(fp.get("price"))
            sqft = _cbas_num(fp.get("sqft"))
            if price:
                _lot(ff)["prices"].append(price)
                _builder(bid, name)["prices"].append(price)
                _builder_lot(bid, name, ff)["prices"].append(price)
            if sqft:
                _lot(ff)["sqfts"].append(sqft)
                _builder(bid, name)["sqfts"].append(sqft)
                _builder_lot(bid, name, ff)["sqfts"].append(sqft)
            if price and sqft:
                ppsf = price / sqft
                _lot(ff)["ppsfs"].append(ppsf)
                _builder(bid, name)["ppsfs"].append(ppsf)
                _builder_lot(bid, name, ff)["ppsfs"].append(ppsf)
            _builder(bid, name)["lot_types"].add(ff)

        def _finish(row, share_basis):
            share = (row.get("lots") or 0) / share_basis if share_basis and row.get("lots") else None
            row.update(_price_stats(row.pop("prices", []), row.pop("sqfts", []),
                                    row.pop("ppsfs", [])))
            row["share_pct"] = round(share * 100, 1) if share is not None else None
            row["est_vdls"] = round((vdls or 0) * share) if share is not None else None
            row["est_futures"] = round((futures or 0) * share) if share is not None else None
            row["est_annual_starts"] = round((ann_starts or 0) * share, 1) if share is not None else None
            row["est_annual_closings"] = round((ann_closings or 0) * share, 1) if share is not None else None
            return row

        lots_out = []
        for row in lot_rows.values():
            builders = [{"name": nm, "lots": round(lots)}
                        for nm, lots in sorted(row.pop("builders", {}).items(),
                                               key=lambda kv: -kv[1])]
            row = _finish(row, section_lots)
            row["builders"] = builders[:8]
            row["builder_count"] = len(builders)
            lots_out.append(row)
        lots_out.sort(key=lambda r: r["lot_width_ff"])

        builder_out = []
        for row in builder_rows.values():
            row["lot_types_ff"] = sorted(row.pop("lot_types", set()))
            builder_out.append(_finish(row, section_lots))
        builder_out.sort(key=lambda r: (-(r.get("lots") or 0), r["name"]))

        builder_lot_out = []
        for row in builder_lot_rows.values():
            builder_lot_out.append(_finish(row, section_lots))
        builder_lot_out.sort(key=lambda r: (r["name"], r["lot_width_ff"]))

        return {
            "section_lots": round(section_lots) if section_lots else None,
            "lot_widths": lots_out,
            "builders": builder_out,
            "builder_lot_widths": builder_lot_out,
        }

    comms = []
    for d_mi, e in near:
        total = e.get("total") or 0
        occ = e.get("occupied") or 0
        vdls = e.get("vdls") or 0
        futures = e.get("futures") or 0
        ann_cl = e.get("annual_closings") or 0
        ann_st = e.get("annual_starts") or 0
        lot_types = sorted({f for f in (_cbas_ff(x) for x in (e.get("lot_types") or []))
                            if f is not None})
        prices = e.get("prices") or {}
        cdist = e.get("districtsNames")
        detail = _community_product_detail(e, ann_st, ann_cl, vdls, futures)
        comms.append({
            "id": e.get("id"),
            "name": e.get("public_name") or e.get("name"),
            "community": e.get("communityName"),
            "developer": e.get("developerName"),
            "city": e.get("city"),
            "zip": e.get("zip"),
            "county": e.get("countyName"),
            "submarket": e.get("submarketName"),
            "status": e.get("status"),
            "lat": e.get("lat"), "lon": e.get("lng"),
            "distance_mi": round(d_mi, 2),
            "direction": _dir(float(e["lat"]), float(e["lng"])),
            "builders": [b for b in (e.get("builders_names") or [])
                         if b and b not in ("Builder TBD",)],
            "builder_count": len([b for b in (e.get("builders_names") or [])
                                  if b and b != "Builder TBD"]),
            "lot_types_ff": lot_types,
            "lot_type_range": e.get("lot_type_range"),
            "price_min": prices.get("min"), "price_max": prices.get("max"),
            "starts_qtr": e.get("starts"), "closings_qtr": e.get("closings"),
            "annual_starts": ann_st, "annual_closings": ann_cl,
            "occupied": occ, "models": e.get("models"),
            "complete_vacant": e.get("complete_vacant"),
            "under_construction": e.get("under_construction"),
            "vdls": vdls, "futures": futures, "total_lots": total,
            "pct_built_out": e.get("buildoutPercent"),
            "months_lot_supply": round(vdls / (ann_cl / 12.0), 1) if ann_cl > 0 else None,
            "years_to_sellout": round((vdls + futures) / ann_cl, 1) if ann_cl > 0 else None,
            "school_district": cdist,
            "schools": e.get("schoolNames") or {},
            "in_district": bool(proj_ds and _norm_d(cdist) in proj_ds),
            "link": e.get("link"),
            "detail": detail,
        })
    comms.sort(key=lambda c: (0 if c["in_district"] else 1, -(c["annual_closings"] or 0)))

    # --- 2) Builders in the ring, by name ----------------------------------
    # `sections[].lot_types_builders` is lots controlled per builder per lot
    # width — the truest read of who is actually buying here. Floorplan pricing
    # supplies price/sqft. Annual starts are attributed to a builder by its lot
    # share of the community (flagged as an estimate in the payload).
    bagg = {}

    def _b(bid):
        if bid not in bagg:
            bagg[bid] = {"builder_id": bid, "name": bmap.get(bid) or f"Builder #{bid}",
                         "named": bid in bmap, "lots": 0, "communities": set(),
                         "lot_types": set(), "prices": [], "sqfts": [],
                         "est_annual_starts": 0.0, "est_annual_closings": 0.0}
        return bagg[bid]

    for d_mi, e in near:
        cname = e.get("public_name") or e.get("name")
        per_builder = {}
        for sec in (e.get("sections") or []):
            for lb in (sec.get("lot_types_builders") or []):
                bid, n = lb.get("builder"), lb.get("num_lots") or 0
                if bid is None or bid in _CBAS_PLACEHOLDER_BUILDERS:
                    continue
                per_builder[bid] = per_builder.get(bid, 0) + n
                b = _b(bid)
                b["lots"] += n
                b["communities"].add(cname)
                ff = _cbas_ff(lb.get("lot_type"))
                if ff:
                    b["lot_types"].add(ff)
        tot = sum(per_builder.values()) or 0
        if tot:
            for bid, n in per_builder.items():
                share = n / tot
                _b(bid)["est_annual_starts"] += (e.get("annual_starts") or 0) * share
                _b(bid)["est_annual_closings"] += (e.get("annual_closings") or 0) * share
        # Builders present with no section detail still count as active here
        for bid in (e.get("builders") or []):
            if (bid is not None and bid not in _CBAS_PLACEHOLDER_BUILDERS
                    and bmap.get(bid) != "Builder TBD"):
                _b(bid)["communities"].add(cname)
        for fp in ((e.get("latestFloorplanPricing") or {}).get("entries") or []):
            bid = fp.get("builderID")
            if bid is None or bid in _CBAS_PLACEHOLDER_BUILDERS:
                continue
            b = _b(bid)
            if fp.get("builderName"):
                b["name"] = fp["builderName"]
                b["named"] = True
            if fp.get("price"):
                b["prices"].append(fp["price"])
            if fp.get("sqft"):
                b["sqfts"].append(fp["sqft"])
            ff = _cbas_ff(fp.get("lot_type"))
            if ff:
                b["lot_types"].add(ff)

    builders = []
    for b in bagg.values():
        if b["name"] == "Builder TBD":
            continue
        pr, sf = b["prices"], b["sqfts"]
        builders.append({
            "builder_id": b["builder_id"], "name": b["name"], "named": b["named"],
            "lots": b["lots"], "communities": len(b["communities"]),
            "in": sorted(b["communities"])[:6],
            "lot_types_ff": sorted(b["lot_types"]),
            "avg_price": round(sum(pr) / len(pr)) if pr else None,
            "min_price": min(pr) if pr else None, "max_price": max(pr) if pr else None,
            "avg_sqft": round(sum(sf) / len(sf)) if sf else None,
            "avg_ppsf": round((sum(pr) / len(pr)) / (sum(sf) / len(sf)), 2)
                        if pr and sf and sum(sf) else None,
            "plans": len(pr),
            "est_annual_starts": round(b["est_annual_starts"]),
            "est_annual_closings": round(b["est_annual_closings"]),
        })
    builders.sort(key=lambda b: (-(b["est_annual_starts"] or 0), -(b["lots"] or 0)))

    # --- 3) Product mix by lot width ---------------------------------------
    bands = {label: {"label": label, "min_ff": lo, "max_ff": hi, "lots": 0,
                     "communities": set(), "builders": set(), "prices": [], "sqfts": []}
             for lo, hi, label in _CBAS_LOT_BANDS}
    for d_mi, e in near:
        cname = e.get("public_name") or e.get("name")
        for sec in (e.get("sections") or []):
            for lb in (sec.get("lot_types_builders") or []):
                lbl = _cbas_band(lb.get("lot_type"))
                if not lbl:
                    continue
                bands[lbl]["lots"] += lb.get("num_lots") or 0
                bands[lbl]["communities"].add(cname)
                lbid = lb.get("builder")
                if (lbid is not None and lbid not in _CBAS_PLACEHOLDER_BUILDERS
                        and bmap.get(lbid) and bmap[lbid] != "Builder TBD"):
                    bands[lbl]["builders"].add(bmap[lbid])
        for fp in ((e.get("latestFloorplanPricing") or {}).get("entries") or []):
            lbl = _cbas_band(fp.get("lot_type"))
            if not lbl:
                continue
            if fp.get("price"):
                bands[lbl]["prices"].append(fp["price"])
            if fp.get("sqft"):
                bands[lbl]["sqfts"].append(fp["sqft"])
    lot_bands = []
    for lo, hi, label in _CBAS_LOT_BANDS:
        b = bands[label]
        pr, sf = b["prices"], b["sqfts"]
        lot_bands.append({
            "label": label, "min_ff": lo, "max_ff": hi, "lots": b["lots"],
            "communities": len(b["communities"]), "builders": len(b["builders"]),
            "avg_price": round(sum(pr) / len(pr)) if pr else None,
            "min_price": min(pr) if pr else None, "max_price": max(pr) if pr else None,
            "avg_sqft": round(sum(sf) / len(sf)) if sf else None,
            "avg_ppsf": round((sum(pr) / len(pr)) / (sum(sf) / len(sf)), 2)
                        if pr and sf and sum(sf) else None,
            "plans": len(pr),
        })

    # --- 4) Absorption by quarter across the ring --------------------------
    qs, qc = {}, {}
    for d_mi, e in near:
        for k, v in (e.get("startsByQuarter") or {}).items():
            qs[k] = qs.get(k, 0) + (v or 0)
        for k, v in (e.get("closingsByQuarter") or {}).items():
            qc[k] = qc.get(k, 0) + (v or 0)

    def _qkey(lbl):
        try:
            y, q = lbl.split(" Q")
            return (int(y), int(q))
        except Exception:
            return (0, 0)
    quarter_series = [{"label": k, "starts": qs.get(k, 0), "closings": qc.get(k, 0)}
                      for k in sorted(set(qs) | set(qc), key=_qkey)]

    # --- 4b) Sales volume the ring currently supports -----------------------
    # Dollar absorption, not just unit absorption: annual closings priced at the
    # community's own midpoint. Priced off each community rather than one blended
    # average so a high-volume entry-level community doesn't get valued at a
    # move-up community's price. Communities with no CBAS pricing fall back to
    # the ring median so they aren't silently counted as $0.
    _mids = [(c["price_min"] + c["price_max"]) / 2.0
             for c in comms if c.get("price_min") and c.get("price_max")]
    _mids.sort()
    ring_median_price = _mids[len(_mids) // 2] if _mids else None
    volume_total = 0.0
    volume_priced = 0
    for c in comms:
        mid = ((c["price_min"] + c["price_max"]) / 2.0
               if c.get("price_min") and c.get("price_max") else ring_median_price)
        c["avg_price"] = round(mid) if mid else None
        c["annual_volume"] = round((c["annual_closings"] or 0) * mid) if mid else None
        if c["annual_volume"]:
            volume_total += c["annual_volume"]
            if c.get("price_min"):
                volume_priced += 1
    for b in builders:
        if b.get("avg_price") and b.get("est_annual_closings"):
            b["est_annual_volume"] = round(b["avg_price"] * b["est_annual_closings"])
        else:
            b["est_annual_volume"] = None

    # --- 5) Ring rollup ----------------------------------------------------
    active = [c for c in comms if (c["annual_closings"] or 0) > 0]
    tot_cl = sum(c["annual_closings"] or 0 for c in comms)
    tot_st = sum(c["annual_starts"] or 0 for c in comms)
    tot_vdl = sum(c["vdls"] or 0 for c in comms)
    tot_fut = sum(c["futures"] or 0 for c in comms)
    all_prices = [c["price_min"] for c in comms if c.get("price_min")] + \
                 [c["price_max"] for c in comms if c.get("price_max")]

    # --- 6) Market-entry read ----------------------------------------------
    # The numbers an underwriter would otherwise derive by hand: what pace one
    # community can expect, what product the market is actually absorbing, and
    # how contested the lot supply is. Every figure traces to the ring above —
    # this summarises, it does not forecast.
    dominant = max(lot_bands, key=lambda b: b["lots"] or 0) if lot_bands else None
    selling = [b for b in lot_bands if (b["plans"] or 0) >= 5]
    best_ppsf = max(selling, key=lambda b: b["avg_ppsf"] or 0) if selling else None
    mos = (round(tot_vdl / (tot_cl / 12.0), 1) if tot_cl else None)
    pace = round(tot_cl / len(active), 1) if active else None

    # Lot absorption is a STARTS question, not a closings question. A start is a
    # builder pulling a lot and breaking ground — that is when the lot sells. A
    # closing is the homebuyer taking delivery months later, which is the
    # builder's metric, not the land developer's.
    #
    # Rather than assume a capture rate, measure what communities in this ring
    # actually take: each active community's share of ring-wide annual starts.
    # That gives a median (a typical community), a p75 (a strong one) and a max
    # (what the best performer here proves is achievable).
    starting = [c for c in comms if (c["annual_starts"] or 0) > 0]
    shares = sorted(((c["annual_starts"] or 0) / tot_st * 100) for c in starting) if tot_st else []

    def _pct(lst, q):
        if not lst:
            return None
        i = min(len(lst) - 1, max(0, int(round(q * (len(lst) - 1)))))
        return round(lst[i], 2)

    top_starter = max(starting, key=lambda c: c["annual_starts"] or 0) if starting else None

    # You don't compete for every start — only for starts in the lot widths you
    # intend to build. Apportion ring starts by the share of ring lots sitting in
    # the project's own target bands to get an addressable figure. Capture of
    # that is the number worth arguing about; capture of all starts understates
    # it whenever the project targets part of the product range.
    # Same default mix /analyze applies when a project hasn't set one, so the
    # addressable figure works before anyone touches the yield inputs.
    _lot_types = ((proj.get("yield_assumptions") or {}).get("lot_types")
                  or [{"label": "40 FF"}, {"label": "50 FF"},
                      {"label": "60 FF"}, {"label": "70 FF"}])
    target_ff = []
    for lt in _lot_types:
        m = _sub_parse_re.search(r"\d+", str(lt.get("label") or ""))
        if m:
            ff = _cbas_ff(m.group())
            if ff:
                target_ff.append(ff)
    target_bands = {_cbas_band(f) for f in target_ff}
    target_bands.discard(None)
    band_lots_total = sum(b["lots"] or 0 for b in lot_bands) or 0
    band_lots_target = sum(b["lots"] or 0 for b in lot_bands if b["label"] in target_bands)
    seg_share = (band_lots_target / band_lots_total) if band_lots_total else None
    addressable = round(tot_st * seg_share) if (tot_st and seg_share) else None

    capture = {
        "ring_annual_starts": tot_st,
        "active_communities": len(starting),
        "share_median_pct": _pct(shares, 0.50),
        "share_p75_pct": _pct(shares, 0.75),
        "share_max_pct": _pct(shares, 1.0),
        "top_community": (top_starter or {}).get("name"),
        "top_community_starts": (top_starter or {}).get("annual_starts"),
        "target_bands": sorted(target_bands),
        "target_ff": sorted(set(target_ff)),
        "segment_share_pct": round(seg_share * 100, 1) if seg_share else None,
        "addressable_starts": addressable,
    }
    # Anchor the ladder to observed percentiles rather than round numbers, then
    # append 20% and 25% because those get asked for — so the answer sits next to
    # the ceiling this ring has actually demonstrated instead of floating free.
    _pts = []
    for pct, lbl in ((capture["share_median_pct"], "median community here"),
                     (capture["share_p75_pct"], "strong community here"),
                     (capture["share_max_pct"], f"best here — {capture['top_community'] or 'top performer'}"),
                     (20.0, "commonly assumed"),
                     (25.0, "commonly assumed")):
        if pct and not any(abs(pct - p["capture_pct"]) < 0.25 for p in _pts):
            _pts.append({"capture_pct": round(pct, 1), "label": lbl})
    _pts.sort(key=lambda p: p["capture_pct"])
    ceiling = capture["share_max_pct"]
    for p in _pts:
        p["lots_per_year"] = round(tot_st * p["capture_pct"] / 100.0) if tot_st else None
        p["addressable_lots_per_year"] = (round(addressable * p["capture_pct"] / 100.0)
                                          if addressable else None)
        p["above_ceiling"] = bool(ceiling and p["capture_pct"] > ceiling)
    capture["scenarios"] = _pts
    capture["base_capture_pct"] = capture["share_p75_pct"]
    top3 = builders[:3]
    top3_share = (round(sum(b["est_annual_starts"] or 0 for b in top3) / tot_st * 100, 1)
                  if tot_st else None)
    market_entry = {
        "capture": capture,
        "expected_pace_per_community": pace,
        "expected_pace_note": (
            f"{pace} closings/yr is the average across {len(active)} actively "
            f"closing communities in the submarket" if pace else None),
        "dominant_product": dominant["label"] if dominant else None,
        "dominant_product_lots": dominant["lots"] if dominant else None,
        "dominant_product_price": dominant["avg_price"] if dominant else None,
        "highest_ppsf_band": best_ppsf["label"] if best_ppsf else None,
        "highest_ppsf": best_ppsf["avg_ppsf"] if best_ppsf else None,
        "months_lot_supply": mos,
        "lot_market": ("tight" if mos is not None and mos < 6
                       else "balanced" if mos is not None and mos < 18
                       else "long" if mos is not None else None),
        "years_of_pipeline": round((tot_vdl + tot_fut) / tot_cl, 1) if tot_cl else None,
        "annual_sales_volume": round(volume_total) if volume_total else None,
        "top3_builders": [b["name"] for b in top3],
        "top3_start_share_pct": top3_share,
        "builder_concentration": ("concentrated" if top3_share and top3_share >= 50
                                  else "fragmented" if top3_share else None),
        "price_band": {"min": min(all_prices) if all_prices else None,
                       "max": max(all_prices) if all_prices else None,
                       "median": round(ring_median_price) if ring_median_price else None},
        "priced_communities": volume_priced,
        "unpriced_communities": len([c for c in comms if not c.get("price_min")]),
    }

    return jsonify({
        "market_entry":      market_entry,
        "annual_sales_volume": round(volume_total) if volume_total else None,
        "ring_median_price": round(ring_median_price) if ring_median_price else None,
        "quarter": latest,
        "quarter_label": quarter_series[-1]["label"] if quarter_series else None,
        "radius_mi": radius_mi,
        "center": {"lat": c_lat, "lon": c_lon},
        "district_name": dname,
        "district_names": dnames,
        "district_community_count": sum(1 for c in comms if c["in_district"]),
        "community_count": len(comms),
        "active_count": len(active),
        "builder_count": len(builders),
        "tracked_universe": len(entries),
        "aggregate": {
            "annual_starts": tot_st, "annual_closings": tot_cl,
            "vdls": tot_vdl, "futures": tot_fut,
            "total_lots": sum(c["total_lots"] or 0 for c in comms),
            "under_construction": sum(c["under_construction"] or 0 for c in comms),
            "occupied": sum(c["occupied"] or 0 for c in comms),
            "complete_vacant": sum(c["complete_vacant"] or 0 for c in comms),
            "models": sum(c["models"] or 0 for c in comms),
            "months_lot_supply": round(tot_vdl / (tot_cl / 12.0), 1) if tot_cl else None,
            "years_of_pipeline": round((tot_vdl + tot_fut) / tot_cl, 1) if tot_cl else None,
            "price_min": min(all_prices) if all_prices else None,
            "price_max": max(all_prices) if all_prices else None,
        },
        "submarket_annual_closings": tot_cl,
        "submarket_annual_starts": tot_st,
        "avg_annual_closings_per_community": round(tot_cl / len(active), 1) if active else None,
        "quarter_series": quarter_series,
        "lot_bands": lot_bands,
        "builders": builders[:30],
        "communities": comms[:80],
        "source": "CBAS new-home survey (cv-server.cbashome.com)",
    })


@acq_bp.route("/api/acq/projects/<pid>/cbas/search", methods=["GET"])
@_login_required
def acq_api_projects_cbas_search(pid):
    """Look up any CBAS community or builder, anywhere — not just inside the ring.

    The competition card only shows what falls within radius_mi, but the cached
    /getData payload holds every community CBAS tracks. This searches that whole
    universe by community name, builder, developer, city or county, and still
    reports each hit's distance from the project so an out-of-ring comparable
    stays in context.

    ?q=      search text (min 2 chars)
    ?kind=   community | builder | all   (default all)
    """
    import math as _m
    from shapely.geometry import shape as shp_shape
    from shapely.ops import unary_union

    token = os.environ.get("CBAS_TOKEN", "").strip()
    if not token:
        return jsonify({"error": "CBAS_TOKEN not configured on the server."}), 503

    q = (request.args.get("q") or "").strip()
    if len(q) < 2:
        return jsonify({"error": "Enter at least 2 characters to search."}), 400
    kind = (request.args.get("kind") or "all").lower()
    ql = q.lower()

    guard = _acq_guard()
    if guard:
        return guard
    uid = _acq_owner()
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, uid, _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404

    geoms = []
    for t in proj.get("tracts") or []:
        g = t.get("geometry")
        if g:
            try:
                geoms.append(shp_shape(g))
            except Exception:
                pass
    c_lat = c_lon = None
    if geoms:
        cen = unary_union(geoms).centroid
        c_lat, c_lon = cen.y, cen.x

    entries, latest, err = _cbas_get_data(token)
    if err:
        return jsonify({"error": err}), 502
    bmap = _cbas_builder_map(entries)

    def _dist(la, lo):
        if c_lat is None or la is None or lo is None:
            return None
        R = 3958.8
        p1, p2 = _m.radians(c_lat), _m.radians(la)
        dp, dl = _m.radians(la - c_lat), _m.radians(lo - c_lon)
        x = _m.sin(dp / 2) ** 2 + _m.cos(p1) * _m.cos(p2) * _m.sin(dl / 2) ** 2
        return round(2 * R * _m.asin(_m.sqrt(x)), 2)

    communities = []
    if kind in ("all", "community"):
        for e in entries:
            hay = " ".join(str(x or "") for x in (
                e.get("name"), e.get("public_name"), e.get("communityName"),
                e.get("developerName"), e.get("city"), e.get("countyName"),
                e.get("submarketName"), e.get("zip"))).lower()
            builders_blob = " ".join(str(b or "") for b in (e.get("builders_names") or [])).lower()
            if ql not in hay and ql not in builders_blob:
                continue
            la, lo = e.get("lat"), e.get("lng")
            prices = e.get("prices") or {}
            communities.append({
                "id": e.get("id"),
                "name": e.get("public_name") or e.get("name"),
                "developer": e.get("developerName"), "city": e.get("city"),
                "county": e.get("countyName"), "submarket": e.get("submarketName"),
                "status": e.get("status"), "lat": la, "lon": lo,
                "distance_mi": _dist(la, lo),
                "builders": [b for b in (e.get("builders_names") or []) if b and b != "Builder TBD"],
                "lot_types_ff": sorted({f for f in (_cbas_ff(x) for x in (e.get("lot_types") or []))
                                        if f is not None}),
                "price_min": prices.get("min"), "price_max": prices.get("max"),
                "annual_starts": e.get("annual_starts"), "annual_closings": e.get("annual_closings"),
                "starts_qtr": e.get("starts"), "closings_qtr": e.get("closings"),
                "vdls": e.get("vdls"), "futures": e.get("futures"),
                "total_lots": e.get("total"), "pct_built_out": e.get("buildoutPercent"),
                "school_district": e.get("districtsNames"),
                "link": e.get("link"),
            })
        communities.sort(key=lambda c: (c["distance_mi"] is None, c["distance_mi"] or 0))

    builders = []
    if kind in ("all", "builder"):
        matched_ids = {bid for bid, nm in bmap.items() if ql in (nm or "").lower()}
        agg = {}
        for e in entries:
            for bid in (e.get("builders") or []):
                if bid not in matched_ids:
                    continue
                a = agg.setdefault(bid, {"builder_id": bid, "name": bmap.get(bid),
                                         "communities": [], "lots": 0, "prices": [],
                                         "sqfts": [], "annual_starts": 0.0})
                d_mi = _dist(e.get("lat"), e.get("lng"))
                a["communities"].append({
                    "name": e.get("public_name") or e.get("name"),
                    "city": e.get("city"), "distance_mi": d_mi,
                    "status": e.get("status"),
                    "annual_closings": e.get("annual_closings"),
                    "lat": e.get("lat"), "lon": e.get("lng"),
                })
            for sec in (e.get("sections") or []):
                for lb in (sec.get("lot_types_builders") or []):
                    if lb.get("builder") in matched_ids:
                        agg.setdefault(lb["builder"], {"builder_id": lb["builder"],
                                                       "name": bmap.get(lb["builder"]),
                                                       "communities": [], "lots": 0,
                                                       "prices": [], "sqfts": [],
                                                       "annual_starts": 0.0})["lots"] += lb.get("num_lots") or 0
            for fp in ((e.get("latestFloorplanPricing") or {}).get("entries") or []):
                if fp.get("builderID") in matched_ids:
                    a = agg.setdefault(fp["builderID"], {"builder_id": fp["builderID"],
                                                         "name": fp.get("builderName"),
                                                         "communities": [], "lots": 0,
                                                         "prices": [], "sqfts": [],
                                                         "annual_starts": 0.0})
                    if fp.get("price"):
                        a["prices"].append(fp["price"])
                    if fp.get("sqft"):
                        a["sqfts"].append(fp["sqft"])
        for a in agg.values():
            pr, sf = a["prices"], a["sqfts"]
            a["communities"].sort(key=lambda c: (c["distance_mi"] is None, c["distance_mi"] or 0))
            builders.append({
                "builder_id": a["builder_id"], "name": a["name"],
                "community_count": len(a["communities"]), "lots": a["lots"],
                "avg_price": round(sum(pr) / len(pr)) if pr else None,
                "avg_sqft": round(sum(sf) / len(sf)) if sf else None,
                "avg_ppsf": round((sum(pr) / len(pr)) / (sum(sf) / len(sf)), 2)
                            if pr and sf and sum(sf) else None,
                "nearest_mi": a["communities"][0]["distance_mi"] if a["communities"] else None,
                "communities": a["communities"][:25],
            })
        builders.sort(key=lambda b: (b["nearest_mi"] is None, b["nearest_mi"] or 0))

    return jsonify({
        "query": q, "kind": kind,
        "center": {"lat": c_lat, "lon": c_lon} if c_lat is not None else None,
        "universe": len(entries),
        "community_count": len(communities),
        "builder_count": len(builders),
        "communities": communities[:60],
        "builders": builders[:25],
        "source": "CBAS new-home survey (cv-server.cbashome.com)",
    })


@acq_bp.route("/api/acq/projects/<pid>/roads", methods=["GET"])
@_login_required
def acq_api_projects_roads(pid):
    """Planned and programmed road projects near the project.

    Highway capacity is the thing that turns an outer-ring tract into a viable
    community, so knowing what TxDOT intends to build — and when — belongs in
    the acquisition package. Three TxDOT ArcGIS layers, verified live:

      Planned_Programmed_Projects_1/20  the planning layer. Carries PT_PHASE
          ("Construction underway or begins soon" / "begins in 5 to 10 years" /
          "Corridor Studies, construction in 10+ years") and LET_YEAR, which is
          exactly the horizon question.
      TxDOT_DCIS_All_Projects/0         the programmed project list, with
          EST_CONST_COST and let dates.
      TxDOT_Future_Texas_Toll_Roads/0   future toll corridors with YearToOpen.

    Query is by bounding box around the project (?radius_mi=, default 15) —
    these are line features, so an envelope is the right primitive.
    """
    import requests as _rq
    from shapely.geometry import shape as shp_shape
    from shapely.ops import unary_union
    from concurrent.futures import ThreadPoolExecutor as _RTP

    guard = _acq_guard()
    if guard:
        return guard
    uid = _acq_owner()
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, uid, _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404

    geoms = []
    for t in proj.get("tracts") or []:
        g = t.get("geometry")
        if g:
            try:
                geoms.append(shp_shape(g))
            except Exception:
                pass
    if not geoms:
        return jsonify({"error": "no tract geometries"}), 400
    centroid = unary_union(geoms).centroid
    c_lat, c_lon = centroid.y, centroid.x

    try:
        radius_mi = float(request.args.get("radius_mi", "15"))
    except (TypeError, ValueError):
        radius_mi = 15.0
    radius_mi = max(2.0, min(radius_mi, 40.0))
    deg = radius_mi / 69.0
    bbox = f"{c_lon-deg},{c_lat-deg},{c_lon+deg},{c_lat+deg}"

    ROOT = "https://services.arcgis.com/KTcxiTD9dsQw4r7Z/arcgis/rest/services"

    # 1500, not 300: the DCIS layer returns 433 features in a 15-mi box around
    # the test tract, so the old cap silently dropped a third of them — and the
    # dropped set is not random, it's whatever the service happened to order
    # last. A truncated list reads as "this is everything" when it isn't.
    def _q(svc, layer, fields, want_geom=True):
        try:
            r = _rq.get(f"{ROOT}/{svc}/FeatureServer/{layer}/query",
                        params={"geometry": bbox, "geometryType": "esriGeometryEnvelope",
                                "inSR": "4326", "spatialRel": "esriSpatialRelIntersects",
                                "outFields": fields, "returnGeometry": str(want_geom).lower(),
                                "outSR": "4326", "resultRecordCount": 1500, "f": "geojson"},
                        timeout=60,
                        headers={"User-Agent": "EmberAcquisitionsGIS/1.0"})
            d = r.json()
            return d.get("features") or []
        except Exception:
            return []

    with _RTP(max_workers=3) as pool:
        f_plan = pool.submit(_q, "Planned_Programmed_Projects_1", 20,
                             "ROADWAY,TYPE_OF_WORK,LIMITS_FROM,LIMITS_TO,DESCRIPTION,"
                             "PROJ_LENGTH,PT_PHASE,LET_YEAR,CSJ,Source")
        f_dcis = pool.submit(_q, "TxDOT_DCIS_All_Projects", 0,
                             "HIGHWAY_NUMBER,PROJ_CLASS,TYPE_OF_WORK,LIMITS_FROM,LIMITS_TO,"
                             "PROJ_LENGTH,EST_CONST_COST,DIST_LET_DATE,ACTUAL_LET_DATE,"
                             "COUNTY_NAME,CONTROL_SECT_JOB")
        f_toll = pool.submit(_q, "TxDOT_Future_Texas_Toll_Roads", 0,
                             "TOLL_NM,OPERATOR,TOLL_STAT,YearToOpen,YearOpen,CMNT,URL")
        plan_f, dcis_f, toll_f = f_plan.result(), f_dcis.result(), f_toll.result()

    # Rank by LET_YEAR first, PT_PHASE only as a fallback. PT_PHASE is stale in
    # this layer — projects with a 2018 let date still read "Construction
    # underway or begins soon", which would put finished work at the top of a
    # list whose whole purpose is what is still coming.
    import datetime as _dt2
    _now_yr = _dt2.date.today().year

    def _horizon(phase, let_year):
        if let_year:
            yrs = let_year - _now_yr
            if yrs < -1:
                return (5, f"Let {let_year} — likely complete")
            if yrs <= 1:
                return (0, "Underway / imminent")
            if yrs <= 4:
                return (1, "1–4 years")
            if yrs <= 10:
                return (2, "5–10 years")
            return (3, "10+ years")
        p = (phase or "").lower()
        if "underway" in p or "begins soon" in p:
            return (0, "Underway / imminent")
        if "5 to 10" in p:
            return (2, "5–10 years")
        if "10+" in p or "corridor stud" in p:
            return (3, "10+ years / study")
        if "1 to 4" in p or "1 to 5" in p or "4 years" in p:
            return (1, "1–4 years")
        return (4, "Unscheduled")

    # Capacity work moves land value; overlays, seal coats, striping, signal
    # upgrades and safety lighting do not. Of the 48 scheduled projects around
    # the test tract only a handful add lanes or build an interchange, so this
    # flag is what lets the card lead with the ones that matter instead of
    # burying them in resurfacing.
    CAPACITY = ("widen", "interchange", "new location", "constr mainlanes",
                "overpass", "grade separation", "expand", "add lanes",
                "added capacity", "freeway", "high occupancy", "hov",
                "managed lane", "toll lane")

    planned = []
    for ft in plan_f:
        a = ft.get("properties") or {}
        rank, label = _horizon(a.get("PT_PHASE"), a.get("LET_YEAR"))
        _blob = f"{a.get('TYPE_OF_WORK','')} {a.get('DESCRIPTION','')}".lower()
        planned.append({
            "capacity": any(k in _blob for k in CAPACITY),
            "roadway": a.get("ROADWAY"), "work": a.get("TYPE_OF_WORK"),
            "description": a.get("DESCRIPTION"),
            "from": a.get("LIMITS_FROM"), "to": a.get("LIMITS_TO"),
            "length_mi": a.get("PROJ_LENGTH"), "phase": a.get("PT_PHASE"),
            "let_year": a.get("LET_YEAR"), "csj": a.get("CSJ"),
            "horizon": label, "horizon_rank": rank,
            "geometry": ft.get("geometry"),
        })
    planned.sort(key=lambda p: (p["horizon_rank"], p["let_year"] or 9999))

    def _ms_to_year(v):
        if not v:
            return None
        try:
            import datetime as _d
            return _d.datetime.utcfromtimestamp(v / 1000.0).year
        except Exception:
            return None

    programmed = []
    for ft in dcis_f:
        a = ft.get("properties") or {}
        cost = a.get("EST_CONST_COST")
        programmed.append({
            "highway": a.get("HIGHWAY_NUMBER"), "klass": a.get("PROJ_CLASS"),
            "work": a.get("TYPE_OF_WORK"), "from": a.get("LIMITS_FROM"),
            "to": a.get("LIMITS_TO"), "length_mi": a.get("PROJ_LENGTH"),
            "cost": cost, "county": a.get("COUNTY_NAME"),
            "let_year": _ms_to_year(a.get("DIST_LET_DATE")) or _ms_to_year(a.get("ACTUAL_LET_DATE")),
            "csj": a.get("CONTROL_SECT_JOB"),
            "geometry": ft.get("geometry"),
        })
    for p in programmed:
        blob = f"{p['klass']} {p['work']}".lower()
        p["capacity"] = any(k in blob for k in CAPACITY)
    programmed.sort(key=lambda p: (not p["capacity"], p["let_year"] or 9999))

    tolls = []
    for ft in toll_f:
        a = ft.get("properties") or {}
        tolls.append({
            "name": a.get("TOLL_NM"), "operator": a.get("OPERATOR"),
            "status": a.get("TOLL_STAT"),
            "year_open": a.get("YearOpen") or a.get("YearToOpen"),
            "note": a.get("CMNT"), "url": a.get("URL"),
            "geometry": ft.get("geometry"),
        })
    tolls = [t for t in tolls if (t["status"] or "").lower() != "existing"]

    near_term = [p for p in planned if p["horizon_rank"] <= 1]
    upcoming = [p for p in planned if p["horizon_rank"] <= 3]
    cap = [p for p in programmed if p["capacity"]]

    # The two layers are complementary rather than redundant, and it's worth
    # measuring by how much. Matching on CSJ (TxDOT's unique project key) over a
    # 15-mi box at the test tract: 35 of 39 scheduled projects also appear in
    # DCIS, but DCIS holds 395 the scheduled layer never lists — including 110
    # capacity projects, among them the US 290 mainlane widening and the
    # Becker / Bauer / Mueschke interchanges. Conversely DCIS publishes no let
    # date or cost at all (0 of 433 populated), so it can say what is programmed
    # but never when. Timing comes from one, coverage from the other.
    def _csj(v):
        return "".join(ch for ch in str(v or "") if ch.isdigit())
    sched_csj = {_csj(p.get("csj")) for p in planned if _csj(p.get("csj"))}
    cap_scheduled = sum(1 for p in cap if _csj(p.get("csj")) in sched_csj)

    # Roll the programmed-capacity list up by roadway server-side. The response
    # truncates `programmed`, so counting it client-side reported 60 against a
    # headline of 121 — the summary has to be built over the full set.
    _by_road = {}
    for p in cap:
        _by_road[p["highway"] or "—"] = _by_road.get(p["highway"] or "—", 0) + 1

    return jsonify({
        "capacity_by_roadway": sorted(_by_road.items(), key=lambda kv: -kv[1]),
        "upcoming_count": len(upcoming),
        "center": {"lat": c_lat, "lon": c_lon},
        "radius_mi": radius_mi,
        "planned": planned[:60],
        "programmed": programmed[:60],
        "future_tolls": tolls[:20],
        "counts": {"planned": len(planned), "programmed": len(programmed),
                   "future_tolls": len(tolls), "near_term": len(near_term),
                   "capacity": len(cap), "capacity_scheduled": cap_scheduled,
                   "capacity_unscheduled": len(cap) - cap_scheduled},
        "capacity_spend": sum(p["cost"] or 0 for p in cap) or None,
        "source": "TxDOT (Planned & Programmed Projects, DCIS, Future Toll Roads)",
    })


@acq_bp.route("/api/acq/school-district/<geoid>", methods=["GET"])
@_login_required
def acq_api_school_district_profile(geoid):
    """District profile: enrollment, schools roster (with distance from the
    project), staffing, growth trend, and rating links per school.

    Source: Urban Institute Education Data API (NCES Common Core of Data) —
    free, no key, verified live. LEAID == the Census GEOID for the district.

    Optional ?lat=&lon= — when supplied, each school gets distance_mi +
    compass direction from that point, and the roster sorts by distance.
    """
    import requests as _rq
    import math as _math
    from concurrent.futures import ThreadPoolExecutor as _TPE

    leaid = (geoid or "").strip()
    if not leaid.isdigit():
        return jsonify({"error": "bad district id"}), 400
    district_name_hint = (request.args.get("name") or "").strip()

    # Optional reference point (the project centroid) for distances
    try:
        ref_lat = float(request.args.get("lat")) if request.args.get("lat") else None
        ref_lon = float(request.args.get("lon")) if request.args.get("lon") else None
    except (TypeError, ValueError):
        ref_lat = ref_lon = None

    B = "https://educationdata.urban.org/api/v1"

    def _get(url, params=None):
        try:
            r = _rq.get(url, params=params or {}, timeout=20)
            if r.status_code != 200: return None
            return r.json()
        except Exception:
            return None

    LATEST = 2024                  # NCES CCD directory has data through 2024
    SCHOOL_ROSTER_YEAR = 2024      # 2024 roster: verified 12 campuses, all with lat/lon
    HISTORY_START = 1990           # verified: CCD directory goes back to 1990

    def _directory(year):
        d = _get(f"{B}/school-districts/ccd/directory/{year}/", {"leaid": leaid})
        res = (d or {}).get("results") or []
        return res[0] if res else None

    def _schools(year):
        d = _get(f"{B}/schools/ccd/directory/{year}/", {"leaid": leaid})
        return (d or {}).get("results") or []

    # Fetch the two REQUIRED calls first, on their own. Bundling them into a
    # 36-request burst got them throttled by the Urban Institute API, which is
    # why the campus roster came back empty. History is best-effort after.
    directory = _directory(LATEST)
    schools   = _schools(SCHOOL_ROSTER_YEAR)
    if not schools:                       # roster fallback if that year is thin
        for yr in (SCHOOL_ROSTER_YEAR - 1, SCHOOL_ROSTER_YEAR - 2):
            schools = _schools(yr)
            if schools: break

    # Fall back to an earlier year if the latest directory row is missing.
    # If NCES does not know the LEAID at all, do not fan out into 34 history
    # calls; return a partial TEA-only profile quickly.
    trend_raw = {}
    if not directory:
        for fallback_yr in (LATEST - 1, LATEST - 2):
            rec = _directory(fallback_yr)
            trend_raw[fallback_yr] = rec
            if rec:
                directory = rec
                break

    # Now the long enrollment history — modest concurrency so we don't trip
    # rate limits. 34 small requests at 6 workers ≈ 5s.
    if directory:
        with _TPE(max_workers=6) as pool:
            f_trend = {yr: pool.submit(_directory, yr)
                        for yr in range(HISTORY_START, LATEST)}
            trend_raw.update({yr: fut.result() for yr, fut in f_trend.items()})

    # ---- Official TEA A–F accountability ratings -------------------------
    # data.texas.gov Socrata dataset nui6-x374 = "School Year 2024-2025
    # Statewide Accountability Ratings" — 10,292 rows covering every Texas
    # district AND campus. Free, no key. Verified live.
    # Join key: TEA district name (NCES lea_name matches closely).
    TEA_URL = "https://data.texas.gov/resource/nui6-x374.json"
    tea_district = None
    tea_campuses = []
    try:
        nces_name = ((directory or {}).get("lea_name") or district_name_hint).strip().upper()
        # Try exact name match first, then a LIKE on the leading word(s).
        rows = _get(TEA_URL, {"district": nces_name, "$limit": 60}) or []
        if not rows and nces_name:
            stem = nces_name.replace(" ISD", "").strip()
            if stem:
                rows = _get(TEA_URL, {
                    "$where": f"upper(district) like '{stem[:24]}%'",
                    "$limit": 60,
                }) or []
        for row in rows:
            if row.get("school_type") == "District":
                tea_district = row
            else:
                tea_campuses.append(row)
    except Exception:
        pass

    def _rating_block(row):
        if not row: return None
        def _num(v):
            try: return float(v)
            except (TypeError, ValueError): return None
        def _pct(v):
            n = _num(v)
            return round(n * 100, 1) if n is not None else None
        return {
            "overall_rating":  row.get("overall_rating"),
            "overall_score":   _num(row.get("overall_score")),
            "student_achievement": {"rating": row.get("student_achievement_rating"),
                                     "score": _num(row.get("student_achievement_score"))},
            "school_progress":     {"rating": row.get("school_progress_rating"),
                                     "score": _num(row.get("school_progress_score"))},
            "closing_the_gaps":    {"rating": row.get("closing_the_gaps_rating"),
                                     "score": _num(row.get("closing_the_gaps_score"))},
            "students":            _num(row.get("number_of_students")),
            "econ_disadvantaged_pct": _pct(row.get("economically_disadvantaged")),
            "el_students_pct":     _pct(row.get("eb_el_students")),
            "region":              row.get("county"),
        }

    # Index TEA campus ratings by a normalized name for joining to NCES schools
    def _norm(n):
        return "".join(ch for ch in (n or "").upper() if ch.isalnum())
    tea_by_name = {}
    for row in tea_campuses:
        key = _norm(row.get("campus"))
        if key:
            tea_by_name[key] = row

    # ---- Current enrolment from TEA, which runs ~1 year ahead of NCES --------
    # NCES CCD stops at 2024 (the Urban Institute API returns count=0 for 2025),
    # so the national feed can only ever be that current. TEA's AskTED directory
    # is refreshed continuously — May 2026 at time of writing — and carries both
    # district and campus enrolment plus an nces_district_id to join on. For
    # Waller ISD that is 10,869 against NCES's 9,905, a year of growth the
    # national dataset has not published yet.
    #
    # NCES still supplies the 1990-2024 history; AskTED cannot, being a snapshot.
    # So: long trend from NCES, current figure from TEA.
    askted = []
    tea_current = None
    tea_as_of = None
    try:
        _at = _get("https://data.texas.gov/resource/hzek-udky.json",
                   {"nces_district_id": f"'{leaid}'", "$limit": 200})
        if not _at:      # the id column is stored with a leading apostrophe
            _at = _get("https://data.texas.gov/resource/hzek-udky.json",
                       {"$where": f"nces_district_id like '%{leaid}'", "$limit": 200})
        askted = _at or []
        if askted:
            try:
                tea_current = int(float(askted[0].get("district_enrollment_as_of")))
            except (TypeError, ValueError):
                tea_current = None
            tea_as_of = (askted[0].get("update_date") or "")[:10]
    except Exception:
        askted = []

    # Campuses TEA lists as not yet open — forward capacity the district is
    # building, which is a demand signal NCES has no field for.
    pipeline_campuses = [
        {"name": (c.get("school_name") or "").title(),
         "status": c.get("school_status"),
         "grades": (c.get("grade_range") or "").lstrip("'")}
        for c in askted
        if (c.get("school_status") or "").upper() not in ("ACTIVE", "")
    ]

    # Full enrollment history (1990 → latest) + growth stats over multiple windows
    trend = []
    for yr in sorted(trend_raw):
        rec = trend_raw[yr]
        if rec and rec.get("enrollment") not in (None, -1, -2):
            trend.append({"year": yr, "enrollment": rec["enrollment"]})
    cur_enroll = directory.get("enrollment") if directory else None
    if cur_enroll not in (None, -1, -2):
        trend.append({"year": LATEST, "enrollment": cur_enroll})
    # Extend the series with TEA's current count so the chart ends on the newest
    # real number rather than a two-year-old one. Flagged as a different source.
    if tea_current and tea_as_of:
        # AskTED is a live directory snapshot, so update_date is a file
        # timestamp, not the year the count describes. A May 2026 refresh is
        # the 2025-26 school year. NCES labels its series by starting year --
        # its 2024 figure is the 2024-25 count, matching TEA's own 2024-25
        # report -- so stamping this point with the calendar year put it a year
        # ahead of where it belonged and left a visible hole in the chart at
        # the year it should have filled.
        try:
            _y, _m = int(tea_as_of[:4]), int(tea_as_of[5:7])
        except (TypeError, ValueError):
            _y, _m = int(tea_as_of[:4]), 1
        _tea_yr = _y if _m >= 8 else _y - 1      # school year starts in August
        if not trend or _tea_yr > trend[-1]["year"]:
            trend.append({"year": _tea_yr, "enrollment": tea_current, "source": "TEA"})
        cur_enroll = tea_current

    def _window(years_back):
        """Growth over the last N years of the series."""
        if len(trend) < 2: return None
        last = trend[-1]
        target_year = last["year"] - years_back
        # nearest available year at/after the target
        start = next((t for t in trend if t["year"] >= target_year), trend[0])
        span = last["year"] - start["year"]
        if span <= 0 or not start["enrollment"]: return None
        ratio = last["enrollment"] / start["enrollment"]
        return {
            "from_year":      start["year"],
            "to_year":        last["year"],
            "span_years":     span,
            "from_enrollment": start["enrollment"],
            "to_enrollment":   last["enrollment"],
            "total_pct":      round((ratio - 1) * 100, 1),
            "cagr_pct":       round((ratio ** (1/span) - 1) * 100, 2),
            "added_students": last["enrollment"] - start["enrollment"],
        }

    growth = {}
    full = _window(999)          # entire available history
    if full:
        growth.update(full)      # keep flat keys for back-compat with the UI
        growth["windows"] = {
            "all_time": full,
            "20_year":  _window(20),
            "10_year":  _window(10),
            "5_year":   _window(5),
        }

    # School roster — level, enrollment, distance from the project
    def _level(s):
        hi = s.get("highest_grade_offered")
        try: hi_i = int(hi)
        except (TypeError, ValueError): hi_i = None
        if hi_i is None: return "Other"
        if hi_i <= 5:  return "Elementary"
        if hi_i <= 8:  return "Middle / Jr High"
        return "High School"

    def _grade_label(g):
        """NCES grade codes: -1 = PK, 0 = K, 1..12 = grade number."""
        try: gi = int(g)
        except (TypeError, ValueError): return "?"
        if gi == -1: return "PK"
        if gi == 0:  return "K"
        return str(gi)

    def _dist_dir(lat, lon):
        if ref_lat is None or ref_lon is None or lat is None or lon is None:
            return None, None
        R = 3958.7613
        p1, p2 = _math.radians(ref_lat), _math.radians(lat)
        dl = _math.radians(lon - ref_lon)
        a = (_math.sin((p2-p1)/2)**2
             + _math.cos(p1)*_math.cos(p2)*_math.sin(dl/2)**2)
        dist = round(2 * R * _math.asin(_math.sqrt(a)), 2)
        deg = (_math.degrees(_math.atan2(lon - ref_lon, lat - ref_lat)) + 360) % 360
        dirs = ["N","NE","E","SE","S","SW","W","NW"]
        return dist, dirs[int((deg + 22.5) / 45) % 8]

    roster = []
    for s in schools:
        enr = s.get("enrollment")
        lat = s.get("latitude"); lon = s.get("longitude")
        # NCES uses -1/-2 as missing-value sentinels
        if lat in (-1, -2): lat = None
        if lon in (-1, -2): lon = None
        dist, direction = _dist_dir(lat, lon)
        nces_id = s.get("ncessch") or ""
        nm = (s.get("school_name") or "").title()
        # Join to the official TEA A–F rating by normalized campus name
        tea_row = tea_by_name.get(_norm(s.get("school_name")))
        tea = _rating_block(tea_row)
        roster.append({
            "name":        nm,
            "level":       _level(s),
            "enrollment":  enr if enr not in (None, -1, -2) else None,
            "grades":      f"{_grade_label(s.get('lowest_grade_offered'))}–{_grade_label(s.get('highest_grade_offered'))}",
            "charter":     bool(s.get("charter")),
            "lat": lat, "lon": lon,
            "distance_mi": dist,
            "direction":   direction,
            # Official TEA A–F rating (inline, not a link)
            "tea_rating":  (tea or {}).get("overall_rating"),
            "tea_score":   (tea or {}).get("overall_score"),
            "tea_detail":  tea,
            # No external rating link: txschools.gov 404s every
            # /schools/<id>/overview form — NCES id and TEA campus_number alike.
            # The A–F rating below is the real payload anyway.
        })
    # Distance-first when we have a reference point, else biggest schools first
    if ref_lat is not None:
        roster.sort(key=lambda x: (x["distance_mi"] if x["distance_mi"] is not None else 9999))
    else:
        roster.sort(key=lambda x: -(x["enrollment"] or 0))

    teachers = (directory or {}).get("teachers_total_fte")
    ratio = None
    if teachers and cur_enroll and teachers > 0:
        ratio = round(cur_enroll / teachers, 1)
    nces_missing = not bool(directory)
    warnings = []
    if nces_missing:
        warnings.append(f"No NCES CCD directory record was found for district {leaid}; showing TEA data where available.")

    return jsonify({
        "leaid":            leaid,
        "name":             (((directory or {}).get("lea_name") or district_name_hint or f"District {leaid}")).title(),
        "year":             LATEST,
        "nces_available":   not nces_missing,
        "partial":          nces_missing,
        "warnings":         warnings,
        "enrollment":       cur_enroll if cur_enroll not in (-1, -2) else None,
        "enrollment_source": "TEA AskTED" if tea_current else (f"NCES CCD {LATEST}" if directory else "Unavailable"),
        "enrollment_as_of": tea_as_of,
        "nces_enrollment":  (directory or {}).get("enrollment"),
        "nces_year":        LATEST if directory else None,
        "pipeline_campuses": pipeline_campuses,
        "schools_count":    (directory or {}).get("number_of_schools") if directory else (len(schools) or None),
        "teachers_fte":     teachers if teachers not in (-1, -2) else None,
        "student_teacher_ratio": ratio,
        "county":           (directory or {}).get("county_name"),
        "city":             ((directory or {}).get("city_mailing") or "").title(),
        "phone":            (directory or {}).get("phone"),
        "enrollment_trend": trend,
        "growth":           growth,
        "schools":          roster,
        # Official TEA A–F district rating (inline)
        "tea":              _rating_block(tea_district),
        "tea_year":         "2024-2025",
        "source":           ("NCES Common Core of Data (Urban Institute API) for the 1990-"
                             f"{LATEST} history + TEA AskTED for current enrolment + "
                             "TEA 2024-25 Accountability Ratings (data.texas.gov)"),
    })


@acq_bp.route("/api/acq/projects/<pid>/builders", methods=["GET"])
@_login_required
def acq_api_projects_builders(pid):
    """Detect active homebuilders in the submarket by scanning cached parcel
    OWNER_NAME fields for known builder company patterns. When a builder owns
    N+ parcels in an area, that's a strong signal they're actively developing.

    Also groups by community/subdivision so you can see who's building where.
    Returns top builders by parcel count + a per-subdivision cross-reference.
    """
    import re as _re
    from shapely.geometry import shape as shp_shape
    from shapely.ops import unary_union, transform

    guard = _acq_guard()
    if guard:
        return guard
    uid = _acq_owner()
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, uid, _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404

    geoms = []
    for t in proj.get("tracts") or []:
        g = t.get("geometry")
        if g:
            try: geoms.append(shp_shape(g))
            except Exception as e: print(f"[{request.endpoint}] tract geometry skipped: {e}", flush=True)
    if not geoms:
        return jsonify({"error": "no tract geometries"}), 400
    proj_union = unary_union(geoms)
    centroid = proj_union.centroid

    try:
        radius_mi = float(request.args.get("radius_mi", "5"))
    except (TypeError, ValueError):
        radius_mi = 5.0
    radius_mi = max(0.5, min(radius_mi, 15.0))
    centroid_utm = transform(to_utm, centroid)
    circle_utm = centroid_utm.buffer(radius_mi * 1609.34)
    circle_wgs = transform(to_wgs, circle_utm)

    import acq_parcels as parcel_cache
    try:
        fc = parcel_cache.query_parcels_in_polygon(circle_wgs, min_acres=0.05, max_acres=50)
        feats = fc.get("features") or []
    except Exception as e:
        return jsonify({"error": f"parcel cache query failed: {e}"}), 500

    builders_all = {}    # builder_name -> {parcels, acres, subdivisions:set, sample_owners:set}
    builders_by_sub = {} # subdivision -> {builder_name: parcel_count}

    for f in feats:
        p = f.get("properties") or {}
        owner = (p.get("OWNER_NAME") or "").strip()
        builder = _classify_builder(owner)   # module-level shared classifier
        if not builder: continue

        # Shared subdivision parser — rejects rural abstracts, strips CAD codes
        sub = _parse_subdivision_name(p.get("LEGAL_DESC") or "")

        # Distance from project centroid
        try:
            pcent = shp_shape(f["geometry"]).centroid
            dist_mi = transform(to_utm, pcent).distance(centroid_utm) / 1609.34
        except Exception:
            dist_mi = None
        acres = float(p.get("Acres") or 0)

        entry = builders_all.setdefault(builder, {
            "name":          builder,
            "parcels":       0,
            "acres":         0.0,
            "subdivisions":  set(),
            "sample_owners": set(),
            "closest_mi":    None,
        })
        entry["parcels"] += 1
        entry["acres"] += acres
        if sub: entry["subdivisions"].add(sub)
        if len(entry["sample_owners"]) < 5:
            entry["sample_owners"].add(owner)
        if dist_mi is not None and (entry["closest_mi"] is None or dist_mi < entry["closest_mi"]):
            entry["closest_mi"] = dist_mi

        if sub:
            builders_by_sub.setdefault(sub, {})
            builders_by_sub[sub][builder] = builders_by_sub[sub].get(builder, 0) + 1

    builders_out = []
    for v in builders_all.values():
        builders_out.append({
            "name":          v["name"],
            "parcels":       v["parcels"],
            "acres":         round(v["acres"], 1),
            "subdivisions":  sorted(v["subdivisions"])[:15],
            "sub_count":     len(v["subdivisions"]),
            "sample_owners": sorted(v["sample_owners"]),
            "closest_mi":    round(v["closest_mi"], 2) if v["closest_mi"] is not None else None,
        })
    builders_out.sort(key=lambda x: x["parcels"], reverse=True)

    # Also produce a per-subdivision → dominant builder table
    sub_dominant = []
    for sub, blist in builders_by_sub.items():
        if not blist: continue
        top_builder = max(blist.items(), key=lambda kv: kv[1])
        sub_dominant.append({
            "subdivision":    sub,
            "top_builder":    top_builder[0],
            "top_count":      top_builder[1],
            "other_builders": [b for b, _ in sorted(blist.items(), key=lambda kv: kv[1], reverse=True)[1:5]],
        })
    sub_dominant.sort(key=lambda x: x["top_count"], reverse=True)

    return jsonify({
        "builders":         builders_out[:40],
        "by_subdivision":   sub_dominant[:40],
        "radius_mi":        radius_mi,
        "parcels_scanned":  len(feats),
        "notes":            [
            "Builders identified by matching OWNER_NAME against known Texas/Houston homebuilder patterns.",
            "A builder holding N+ parcels in a subdivision usually means they're actively developing / selling.",
            "Some parcels are held by the buyer, not the builder — so this UNDERSTATES builder activity in mature communities.",
            "For live builder pricing / absorption by community, need Zonda subscription.",
        ],
    })


@acq_bp.route("/api/acq/projects/<pid>/news", methods=["GET"])
@_login_required
def acq_api_projects_news(pid):
    """Pull recent news stories relevant to the project's submarket — economic
    development, major employer moves, infrastructure, jobs announcements.

    Uses Google News RSS (free, no key). Builds queries from the project's
    city / school district / county / MSA.
    """
    import requests as _rq
    import re as _re
    import xml.etree.ElementTree as _ET
    from shapely.geometry import shape as shp_shape
    from shapely.ops import unary_union

    guard = _acq_guard()
    if guard:
        return guard
    uid = _acq_owner()
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, uid, _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404

    # Build SUBMARKET-SPECIFIC search terms. Key insight from testing: the City
    # of Houston's boundary snakes up utility corridors 30+ miles from downtown,
    # so a rural Hockley project "contains" Houston — which floods the news with
    # metro-wide stories. So: EXCLUDE Houston, and instead pull the actual small
    # towns near the project from TIGERweb Places within a ~10-mi bbox.
    # TWO search terms only, per acquisitions feedback: the CITY the project is
    # in and its SCHOOL DISTRICT. Pulling every nearby town (Pine Island, Prairie
    # View, ...) produced irrelevant noise. Houston is excluded — its boundary
    # snakes up utility corridors so rural sites falsely "contain" it.
    _TERM_BLOCKLIST = {"houston", "houston city", "houston texas"}
    terms = []            # ordered: city first, then district

    geoms = []
    for t in proj.get("tracts") or []:
        g = t.get("geometry")
        if g:
            try: geoms.append(shp_shape(g))
            except Exception as e: print(f"[{request.endpoint}] tract geometry skipped: {e}", flush=True)
    if geoms:
        proj_union = unary_union(geoms)
        centroid = proj_union.centroid
        c_lon, c_lat = centroid.x, centroid.y

        # TIGERweb point-in-polygon. The Census *geocoder* silently omits school
        # districts and places for many rural coordinates (verified), so query
        # the TIGERweb layers directly — those return correctly.
        TIGER = "https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb"
        def _tiger(service, layer, geometry, geom_type):
            try:
                r = _rq.get(f"{TIGER}/{service}/MapServer/{layer}/query",
                            params={"geometry": geometry,
                                    "geometryType": geom_type,
                                    "inSR": "4326",
                                    "spatialRel": "esriSpatialRelIntersects",
                                    "outFields": "NAME,GEOID",
                                    "returnGeometry": "false",
                                    "f": "json"},
                            timeout=12)
                d = r.json()
                if d.get("error"): return []
                return [f.get("attributes") or {} for f in (d.get("features") or [])]
            except Exception:
                return []

        pt = f"{c_lon},{c_lat}"
        PLACES = "Places_CouSub_ConCity_SubMCD"

        # 1) City — point-in-polygon first (incorporated, then CDP). Rural sites
        #    are usually unincorporated, so fall back to the NEAREST town within
        #    ~8 mi, skipping Houston (its limits snake far up the corridors).
        city = None
        for layer in (4, 5):        # 4 = Incorporated Places, 5 = CDPs
            for a in _tiger(PLACES, layer, pt, "esriGeometryPoint"):
                nm = _re.sub(r"\s+(city|town|village|CDP)$", "",
                             (a.get("NAME") or "").strip(), flags=_re.I).strip()
                if nm and nm.lower() not in _TERM_BLOCKLIST:
                    city = nm
                    break
            if city: break
        if not city:
            deg = 8 / 69.0
            bbox = f"{c_lon-deg},{c_lat-deg},{c_lon+deg},{c_lat+deg}"
            candidates = []
            for layer in (4, 5):
                for a in _tiger(PLACES, layer, bbox, "esriGeometryEnvelope"):
                    nm = _re.sub(r"\s+(city|town|village|CDP)$", "",
                                 (a.get("NAME") or "").strip(), flags=_re.I).strip()
                    if nm and nm.lower() not in _TERM_BLOCKLIST:
                        candidates.append(nm)
            # Prefer the shortest name — proxies for the small local town over
            # sprawling multi-word entities, and Houston is already excluded.
            if candidates:
                city = sorted(set(candidates), key=len)[0]
        if city:
            terms.append(city + " Texas")

        # 2) School districts — ALL of them, not just the centroid's. A tract
        #    straddling a boundary sits in two markets and both matter; query the
        #    project bbox across Unified / Secondary / Elementary layers.
        _dg = 0.02
        _dbox = f"{c_lon-_dg},{c_lat-_dg},{c_lon+_dg},{c_lat+_dg}"
        try:
            _b = proj_union.bounds
            _dbox = f"{_b[0]},{_b[1]},{_b[2]},{_b[3]}"
        except Exception:
            pass
        _seen_d = set()
        for _layer in (0, 1, 2):
            for a in _tiger("School", _layer, _dbox, "esriGeometryEnvelope"):
                nm = (a.get("NAME") or "").strip()
                if nm and nm not in _seen_d:
                    _seen_d.add(nm)
                    terms.append(nm.replace(" Independent School District", " ISD"))

    # Fallback — county, only if the geocoder gave us nothing usable
    if not terms:
        for t in proj.get("tracts") or []:
            if t.get("county"):
                terms.append(f"{t['county']} County Texas")
                break

    # Economic-development keywords — universal across any Texas submarket.
    # Grouped with OR so each query casts a wide net without going off-topic.
    QUERY_TEMPLATES = [
        '"{term}" (economic development OR new jobs OR expansion)',
        '"{term}" (manufacturing OR distribution center OR data center OR industrial park)',
        '"{term}" (new development OR master planned community OR homebuilder OR subdivision)',
        '"{term}" (infrastructure OR highway OR road expansion OR utility OR water district)',
        '"{term}" (population growth OR fastest growing OR relocation)',
    ]
    # A free-text ?q= overrides the derived terms entirely — it's used by the
    # search box on the card so you can chase a specific story (a named employer,
    # a corridor, a competitor) without being limited to city + district.
    custom_q = (request.args.get("q") or "").strip()
    if custom_q:
        terms = [custom_q]
        queries = [custom_q]
    else:
        queries = []
        # City plus every school district the tract touches. Capped so a tract
        # spanning three districts doesn't fan out into 20 RSS fetches.
        for term in terms[:4]:
            for tpl in QUERY_TEMPLATES:
                queries.append(tpl.format(term=term))

    def fetch_google_news(q):
        try:
            url = "https://news.google.com/rss/search"
            r = _rq.get(url, params={"q": q, "hl": "en-US", "gl": "US", "ceid": "US:en"},
                        timeout=10, headers={"User-Agent": "Mozilla/5.0"})
            if r.status_code != 200: return []
            root = _ET.fromstring(r.text)
            items = []
            for item in root.iter("item"):
                title = item.findtext("title") or ""
                link = item.findtext("link") or ""
                pub = item.findtext("pubDate") or ""
                # source is inside a <source> element in Google News
                src_el = item.find("source")
                src = src_el.text if src_el is not None else ""
                items.append({"title": title, "link": link, "published": pub, "source": src,
                              "query": q})
            return items[:6]   # cap per query
        except Exception:
            return []

    from concurrent.futures import ThreadPoolExecutor as _TPE
    all_items = []
    seen_links = set()
    with _TPE(max_workers=8) as pool:
        results = list(pool.map(fetch_google_news, queries[:16]))
    # Google News matches loosely: a search for "Waller Texas" returned an INNIO
    # story about a Waukesha expansion in Wisconsin, which mentions neither the
    # town nor anything in the county. Require the place itself to appear in the
    # headline, and reject headlines naming a different state - the term alone
    # being in the query proves nothing about the article.
    place_words = set()
    for t in terms:
        for w in _re.split(r"[^A-Za-z0-9]+", str(t)):
            w = w.strip()
            if len(w) > 2 and w.upper() not in ("TEXAS", "ISD", "THE", "AND", "COUNTY"):
                place_words.add(w.lower())

    OTHER_STATES = (
        "alabama", "alaska", "arizona", "arkansas", "california", "colorado",
        "connecticut", "delaware", "florida", "georgia", "hawaii", "idaho",
        "illinois", "indiana", "iowa", "kansas", "kentucky", "louisiana", "maine",
        "maryland", "massachusetts", "michigan", "minnesota", "mississippi",
        "missouri", "montana", "nebraska", "nevada", "hampshire", "jersey",
        "mexico", "york", "carolina", "dakota", "ohio", "oklahoma", "oregon",
        "pennsylvania", "rhode island", "tennessee", "utah", "vermont",
        "virginia", "washington", "wisconsin", "wyoming",
    )

    filtered_out = 0
    for items in results:
        for it in items:
            if it["link"] in seen_links: continue
            hay = (it.get("title") or "").lower()
            src = (it.get("source") or "").lower()
            if place_words and not any(w in hay for w in place_words):
                filtered_out += 1
                continue
            # A headline or outlet naming another state is not local news, even
            # if the place word happens to appear.
            if any(st in hay or st in src for st in OTHER_STATES):
                filtered_out += 1
                continue
            seen_links.add(it["link"])
            all_items.append(it)

    # Very rough recency sort — parse pubDate (RFC 822 format)
    from email.utils import parsedate_to_datetime
    def _pub_to_dt(s):
        try: return parsedate_to_datetime(s)
        except Exception: return None
    all_items.sort(key=lambda x: _pub_to_dt(x["published"]) or 0, reverse=True)

    return jsonify({
        "stories":   all_items[:40],
        "filtered_out": filtered_out,
        "queries":   queries[:20],
        "terms":     sorted(terms),
        "custom_query": custom_q or None,
        "source":    "Google News RSS (aggregated across queries)",
    })


@acq_bp.route("/api/acq/projects/<pid>/amenities", methods=["GET"])
@_login_required
def acq_api_projects_amenities(pid):
    """Nearby grocery stores, schools, and other draws for residential demand.

    Uses OpenStreetMap Overpass API (free, no key). Returns for each POI:
      • name / brand
      • kind (supermarket, school, hospital, etc.)
      • lat/lon
      • distance from project centroid (miles)
      • driving direction (basic bearing)

    Sorted by distance ascending; major-brand grocery highlighted separately.
    """
    import requests as _rq
    import math as _math
    from shapely.geometry import shape as shp_shape
    from shapely.ops import unary_union, transform

    guard = _acq_guard()
    if guard:
        return guard
    uid = _acq_owner()
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, uid, _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404

    geoms = []
    for t in proj.get("tracts") or []:
        g = t.get("geometry")
        if g:
            try: geoms.append(shp_shape(g))
            except Exception as e: print(f"[{request.endpoint}] tract geometry skipped: {e}", flush=True)
    if not geoms:
        return jsonify({"error": "no tract geometries"}), 400
    proj_union = unary_union(geoms)
    centroid = proj_union.centroid
    c_lat, c_lon = centroid.y, centroid.x

    try:
        radius_mi = float(request.args.get("radius_mi", "5"))
    except (TypeError, ValueError):
        radius_mi = 5.0
    radius_mi = max(0.5, min(radius_mi, 15.0))
    radius_m = int(radius_mi * 1609.34)

    # Overpass — CORE query (grocery / hospital / school) kept lean so it returns
    # in seconds; parks fetched separately and degrade silently if slow. `nwr`
    # matches nodes+ways+relations in one clause; `out center N;` caps results
    # and returns way/relation centroids. UA header required (406 without it).
    # _overpass_query races every planet mirror — see its notes for why.
    # Schools live in the School District Profile card now, so they're out of
    # this query — fewer clauses means a faster, less throttle-prone request.
    core_q = f"""
[out:json][timeout:25];
(
  nwr["shop"="supermarket"](around:{radius_m},{c_lat},{c_lon});
  nwr["shop"="department_store"](around:{radius_m},{c_lat},{c_lon});
  nwr["shop"="wholesale"](around:{radius_m},{c_lat},{c_lon});
  nwr["amenity"="hospital"](around:{radius_m},{c_lat},{c_lon});
);
out center 200;
"""
    parks_q = f"""
[out:json][timeout:20];
nwr["leisure"="park"](around:{radius_m},{c_lat},{c_lon});
out center 100;
"""
    # Commute & access. For outer-ring Houston tracts this is the single biggest
    # value driver — "how far to 290?" is the first question asked about a deal.
    #
    # Query the motorway/trunk WAYS, not the junction nodes: motorway_junction
    # nodes here carry neither `ref` nor `name` (95 of them around the test
    # tract, all anonymous), so they can say "a ramp is 1.9 mi away" but never
    # which highway it belongs to. The ways carry ref + name — "US 290;TX 6,
    # Northwest Freeway" — which is the answer people actually want. Junction
    # nodes are still worth one clause as a bare nearest-ramp distance, since
    # being near a freeway you can't get onto is a real and different problem.
    #
    # These radii are deliberately generous and fixed, not tied to radius_mi:
    # the point is always to surface the CLOSEST one, then let the payload flag
    # whether it fell inside the user's ring.
    ACCESS_M = int(15 * 1609.34)
    PR_M = int(25 * 1609.34)
    AIR_M = int(50 * 1609.34)
    access_q = f"""
[out:json][timeout:60];
(
  way["highway"~"^(motorway|trunk)$"](around:{ACCESS_M},{c_lat},{c_lon});
  node["highway"="motorway_junction"](around:{ACCESS_M},{c_lat},{c_lon});
  nwr["park_ride"]["park_ride"!="no"](around:{PR_M},{c_lat},{c_lon});
  nwr["aeroway"="aerodrome"]["iata"](around:{AIR_M},{c_lat},{c_lon});
);
out center 400;
"""
    # Retail anchors — a rooftop-density and retail-maturity signal, not just
    # convenience. department_store/wholesale overlap the core grocery query on
    # purpose; _ingest routes by brand so Walmart lands in both reads correctly.
    # department_store / wholesale are repeated here at the wider retail radius
    # on purpose: the core query only reaches radius_mi, so on a rural tract the
    # Target and Costco that define the retail picture sit well outside it and
    # would never be seen. Duplicates are collapsed after ingest.
    RETAIL_M = int(15 * 1609.34)
    FUEL_M = int(8 * 1609.34)
    retail_q = f"""
[out:json][timeout:60];
(
  nwr["shop"="supermarket"](around:{RETAIL_M},{c_lat},{c_lon});
  nwr["shop"="department_store"](around:{RETAIL_M},{c_lat},{c_lon});
  nwr["shop"="wholesale"](around:{RETAIL_M},{c_lat},{c_lon});
  nwr["shop"="doityourself"](around:{RETAIL_M},{c_lat},{c_lon});
  nwr["amenity"="pharmacy"](around:{int(10 * 1609.34)},{c_lat},{c_lon});
  nwr["amenity"="fuel"](around:{FUEL_M},{c_lat},{c_lon});
);
out center 300;
"""
    _overpass = _overpass_query   # shared module-level helper

    # Run the categories concurrently, but only two in flight: each one already
    # races both planet mirrors, and firing all four at once means eight
    # simultaneous requests to two hosts that throttle aggressively.
    from concurrent.futures import ThreadPoolExecutor as _ATP
    with _ATP(max_workers=2) as _apool:
        _f_core = _apool.submit(_overpass, core_q, 45)
        _f_parks = _apool.submit(_overpass, parks_q, 30)
        _f_access = _apool.submit(_overpass, access_q, 60)
        _f_retail = _apool.submit(_overpass, retail_q, 60)
        data, core_err = _f_core.result()
        parks_data, _parks_err = _f_parks.result()
        access_data, _access_err = _f_access.result()
        retail_data, _retail_err = _f_retail.result()

    if data is None:
        return jsonify({"error": f"Overpass query failed on all mirrors: {core_err}",
                        "grocery_stores": [], "schools": [], "hospitals": [], "parks": [],
                        "highways": [], "retail_anchors": []}), 502
    # Everything past the core query is best-effort — a slow parks or access
    # lookup must never take grocery and hospitals down with it.
    for _extra in (parks_data, retail_data):
        if _extra and _extra.get("elements"):
            data.setdefault("elements", []).extend(_extra["elements"])

    def _distance_mi(lat, lon):
        # Haversine — good enough at these scales
        R = 3958.7613   # earth radius in miles
        lat1 = _math.radians(c_lat); lon1 = _math.radians(c_lon)
        lat2 = _math.radians(lat);   lon2 = _math.radians(lon)
        a = _math.sin((lat2-lat1)/2)**2 + _math.cos(lat1)*_math.cos(lat2)*_math.sin((lon2-lon1)/2)**2
        return round(2 * R * _math.asin(_math.sqrt(a)), 2)
    def _bearing(lat, lon):
        # Simple 8-direction compass bearing from centroid
        dy = lat - c_lat
        dx = lon - c_lon
        if abs(dy) < 1e-5 and abs(dx) < 1e-5: return ""
        deg = (_math.degrees(_math.atan2(dx, dy)) + 360) % 360
        dirs = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"]
        return dirs[int((deg + 22.5) / 45) % 8]

    grocery = []; schools = []; hospitals = []; parks = []
    retail = []; pharmacies = []; fuel = []
    MAJOR_GROCERY_BRANDS = {"heb", "h-e-b", "kroger", "walmart", "target", "whole foods",
                              "sprouts", "aldi", "trader joe", "randalls", "sam's club",
                              "costco", "food town", "fiesta", "central market"}
    # National anchors worth calling out by name — their presence says more about
    # a submarket's retail maturity than a count of unnamed storefronts does.
    ANCHOR_BRANDS = {"walmart", "target", "costco", "sam's club", "home depot",
                     "lowe's", "kohl's", "jcpenney", "best buy", "academy",
                     "tractor supply", "marshalls", "tj maxx", "ross", "burlington"}

    # The same OSM feature arrives more than once: `nwr` matches a POI tagged on
    # both a node and its building way, and the retail sweep repeats the
    # supermarket / department_store / wholesale clauses at a wider radius than
    # the core query. Skip anything already ingested by OSM identity.
    _seen_osm = set()

    def _clean_poi_name(tags):
        for key in ("name", "brand", "operator"):
            val = (tags.get(key) or "").strip()
            if val and val.lower() not in ("(unnamed)", "unnamed", "unknown"):
                return val
        return None

    def _ingest(elements):
        for el in elements or []:
            oid = (el.get("type"), el.get("id"))
            if oid[1] is not None:
                if oid in _seen_osm:
                    continue
                _seen_osm.add(oid)
            tags = el.get("tags") or {}
            name = _clean_poi_name(tags)
            if not name:
                continue
            brand = (tags.get("brand") or tags.get("operator") or "").strip()
            # Location — for ways use center, for nodes use lat/lon
            lat = el.get("lat") or (el.get("center") or {}).get("lat")
            lon = el.get("lon") or (el.get("center") or {}).get("lon")
            if lat is None or lon is None: continue
            d = _distance_mi(lat, lon)
            b = _bearing(lat, lon)
            entry = {
                "name":         name,
                "brand":        brand,
                "lat":          lat, "lon": lon,
                "distance_mi":  d,
                "direction":    b,
                "tags":         {k: v for k, v in tags.items()
                                  if k in ("name","brand","operator","addr:city","addr:street","addr:full")},
            }
            shop = tags.get("shop") or ""
            amenity = tags.get("amenity") or ""
            leisure = tags.get("leisure") or ""
            blob = (name + " " + brand).lower()
            if shop in ("supermarket", "department_store", "wholesale"):
                is_grocer = any(m in blob for m in MAJOR_GROCERY_BRANDS)
                # Only a store you can actually buy groceries at counts as
                # grocery. shop=department_store covers Walmart and Target but
                # equally Marshalls, Ross and JCPenney — before this check the
                # wider retail sweep pushed Marshalls to the top of "closest
                # grocery store", which is the one thing that list must get right.
                if shop == "supermarket" or is_grocer:
                    grocery.append(dict(entry, is_major=is_grocer))
                # A Walmart or Costco is both a grocery run and a retail anchor;
                # it belongs in both reads rather than being forced into one.
                if shop in ("department_store", "wholesale"):
                    retail.append(dict(entry, kind=shop,
                                       is_anchor=any(m in blob for m in ANCHOR_BRANDS)))
            elif shop in ("doityourself", "hardware"):
                retail.append(dict(entry, kind="home_improvement",
                                   is_anchor=any(m in blob for m in ANCHOR_BRANDS)))
            elif amenity == "pharmacy" or tags.get("healthcare") == "pharmacy":
                pharmacies.append(entry)
            elif amenity == "fuel":
                fuel.append(entry)
            elif amenity == "school":
                schools.append(entry)
            elif amenity == "hospital":
                hospitals.append(entry)
            elif leisure == "park":
                parks.append(entry)

    _ingest(data.get("elements"))

    # --- Commute & access -------------------------------------------------
    # Highways are deduped by `ref` and reduced to the nearest point on each,
    # so "US 290" appears once at its closest approach rather than as 40 way
    # segments. Junction nodes are anonymous, so they only yield a single
    # nearest-ramp distance.
    highways = {}
    ramp_mi = None
    park_ride = []
    airports = []
    for el in (access_data or {}).get("elements") or []:
        tags = el.get("tags") or {}
        lat = el.get("lat") or (el.get("center") or {}).get("lat")
        lon = el.get("lon") or (el.get("center") or {}).get("lon")
        if lat is None or lon is None:
            continue
        d = _distance_mi(lat, lon)
        if tags.get("highway") == "motorway_junction":
            if ramp_mi is None or d < ramp_mi:
                ramp_mi = d
            continue
        if tags.get("highway") in ("motorway", "trunk"):
            ref = (tags.get("ref") or tags.get("name") or "").strip()
            if not ref:
                continue
            cur = highways.get(ref)
            if not cur or d < cur["distance_mi"]:
                highways[ref] = {"ref": ref, "name": tags.get("name") or "",
                                 "kind": tags.get("highway"),
                                 "toll": tags.get("toll") == "yes",
                                 "lat": lat, "lon": lon,
                                 "distance_mi": d, "direction": _bearing(lat, lon)}
            continue
        if tags.get("park_ride"):
            park_ride.append({"name": tags.get("name") or "Park & Ride",
                              "lat": lat, "lon": lon, "distance_mi": d,
                              "direction": _bearing(lat, lon)})
            continue
        if tags.get("aeroway") == "aerodrome":
            iata = (tags.get("iata") or "").strip()
            name = _clean_poi_name(tags) or iata
            if not name:
                continue
            airports.append({"name": name,
                             "iata": iata,
                             "type": tags.get("aerodrome:type") or "",
                             "lat": lat, "lon": lon, "distance_mi": d,
                             "direction": _bearing(lat, lon)})
    highways = sorted(highways.values(), key=lambda h: h["distance_mi"])[:6]
    park_ride.sort(key=lambda p: p["distance_mi"])
    # A 16-mi general-aviation strip is not what "nearest airport" means to a
    # homebuyer, but OSM tags cannot tell the two apart here: of the Houston
    # aerodromes only IAH carries aerodrome=international, while Hobby — a major
    # commercial airport — has nothing beyond iata/icao, exactly like the private
    # fields. So match against the FAA commercial-service list instead. Anything
    # unmatched is still returned, flagged, so a genuinely remote site isn't left
    # with a blank row.
    airports.sort(key=lambda a: a["distance_mi"])
    for a in airports:
        a["commercial"] = a["iata"].upper() in _TX_COMMERCIAL_AIRPORTS
    _comm = [a for a in airports if a["commercial"]]
    airports = (_comm or airports)[:3]

    # "Closest grocery store" must mean CLOSEST, not "within the base radius" —
    # rural sites routinely have no supermarket within 5 mi. Groceries no longer
    # need a widening pass: the retail sweep already queries supermarkets out to
    # 15 mi, so the nearest one is in hand on the first round.
    #
    # That sweep is also why the widening below tests `hospitals` alone. It used
    # to bail as soon as `grocery` was non-empty, and once the 15-mi retail
    # results began landing in grocery, a Costco 8 mi out satisfied that test and
    # suppressed the pass that had been surfacing the H-E-B at 6.3 mi — the
    # nearest real grocery store silently disappeared from the card.
    grocery_radius_used = 15
    hospital_radius_used = radius_mi
    for wider_mi in (10, 20, 30):
        if hospitals:
            break
        wider_m = int(wider_mi * 1609.34)
        widen_q = ("[out:json][timeout:25];\n"
                   f'nwr["amenity"="hospital"](around:{wider_m},{c_lat},{c_lon});\n'
                   "out center 100;\n")
        widen_data, _werr = _overpass(widen_q, 40)
        if widen_data:
            _ingest(widen_data.get("elements"))
            if hospitals:
                hospital_radius_used = wider_mi

    # Sort by distance
    for lst in (grocery, schools, hospitals, parks, retail, pharmacies, fuel):
        lst.sort(key=lambda x: x["distance_mi"])

    def _dedupe(items, tol_mi=0.2):
        """Collapse the same real-world place appearing more than once.

        OSM identity alone isn't enough: a store mapped as both a node and a
        building outline has two different ids, and the node sits a few metres
        off the way's centroid, so a rounded-coordinate key misses it too. What
        the eye catches is the giveaway — same name, effectively same distance —
        so match on that: identical name within `tol_mi`. Unnamed features fall
        back to position, since every unnamed park would otherwise collapse into
        one entry.
        """
        out = []
        for x in items:
            nm = (x.get("name") or "").strip().lower()
            dup = False
            if nm and nm != "(unnamed)":
                for y in out:
                    if ((y.get("name") or "").strip().lower() == nm
                            and abs(y["distance_mi"] - x["distance_mi"]) <= tol_mi):
                        dup = True
                        break
            else:
                for y in out:
                    if (abs(y["lat"] - x["lat"]) < 1e-4
                            and abs(y["lon"] - x["lon"]) < 1e-4):
                        dup = True
                        break
            if not dup:
                out.append(x)
        return out

    grocery = _dedupe(grocery)
    schools = _dedupe(schools)
    hospitals = _dedupe(hospitals)
    parks = _dedupe(parks)
    pharmacies = _dedupe(pharmacies)
    fuel = _dedupe(fuel)
    _dedup = _dedupe(retail)
    # shop=doityourself is applied to national big-box stores and to one-room
    # tool shops alike, so an unfiltered 15-mi sweep buries Target and Costco
    # under "Katy Tools" and "1 Top Tools". Independents still count when they're
    # genuinely close (they serve the rooftops), but out at the edge of the ring
    # only the national anchors say anything about the submarket.
    retail = [r for r in _dedup if r.get("is_anchor") or r["distance_mi"] <= radius_mi]
    retail.sort(key=lambda r: (not r.get("is_anchor"), r["distance_mi"]))

    # Every list is searched on a generous fixed radius so the CLOSEST one always
    # surfaces; flag which entries actually fell inside the user's ring so the UI
    # can say "nearest is 9.3 mi — outside the 5 mi ring" instead of showing
    # nothing. Anything already widened above keeps its own radius note.
    for lst in (grocery, hospitals, parks, retail, pharmacies, fuel):
        for x in lst:
            x["within_radius"] = x["distance_mi"] <= radius_mi

    # Cap
    grocery = grocery[:30]
    schools = schools[:30]
    hospitals = hospitals[:20]
    parks = parks[:20]
    retail = retail[:30]
    pharmacies = pharmacies[:12]
    fuel = fuel[:12]

    # ------------------------------------------------------------------
    # Full-buildout demand + commercial gap
    # ------------------------------------------------------------------
    # What the trade area looks like when every platted lot is built, and which
    # commercial uses that population would support that aren't here yet.
    #
    # Population base: ACS for what exists today (it counts all housing, not just
    # the new-home communities CBAS tracks), plus the remaining CBAS lot pipeline
    # times household size for the growth. When Census is unavailable, fall back
    # to CBAS occupied homes, which undercounts existing rural housing — the
    # payload says which basis was used so the number is never read as more
    # precise than it is.
    TX_AVG_HH = 2.85          # Texas average household size, ACS
    demo = None
    try:
        demo = _acs_radius_demographics(c_lat, c_lon, radius_mi)
    except Exception:
        demo = None
    hh_size = (demo or {}).get("avg_household_size") or TX_AVG_HH

    occupied = remaining = total_lots = 0
    annual_closings = 0
    cbas_ok = False
    try:
        _tok = os.environ.get("CBAS_TOKEN", "").strip()
        if _tok:
            _ents, _latest, _err = _cbas_get_data(_tok)
            if _ents:
                for e in _ents:
                    la, lo = e.get("lat"), e.get("lng")
                    if not la or not lo:
                        continue
                    if _distance_mi(float(la), float(lo)) > radius_mi:
                        continue
                    cbas_ok = True
                    occupied += e.get("occupied") or 0
                    total_lots += e.get("total") or 0
                    remaining += ((e.get("vdls") or 0) + (e.get("futures") or 0)
                                  + (e.get("under_construction") or 0))
                    annual_closings += e.get("annual_closings") or 0
    except Exception:
        cbas_ok = False

    # Rooftops today come from OSM building footprints, NOT CBAS occupied homes.
    # CBAS only tracks actively-selling new-home communities, so it misses every
    # subdivision that finished selling — in mature Cinco Ranch that is the
    # difference between 405 and 44,444. Using it here would have understated
    # existing rooftops on any site with established housing nearby, and it is
    # the same denominator the commercial benchmarks are calibrated against, so
    # the two now agree by construction.
    roofs_now = None
    try:
        bq = (f'[out:json][timeout:120];\n(\n'
              f'  way["building"~"^(house|detached|residential|semidetached_house|terrace)$"]'
              f'(around:{radius_m},{c_lat},{c_lon});\n'
              f'  way["building"="yes"](around:{radius_m},{c_lat},{c_lon});\n'
              f');\nout count;\n')
        bd, _berr = _overpass(bq, 90)
        if bd:
            _els = bd.get("elements") or []
            if _els:
                roofs_now = int((_els[0].get("tags") or {}).get("total") or 0) or None
    except Exception:
        roofs_now = None
    if roofs_now is None:
        roofs_now = occupied or None       # last resort; known to undercount

    pop_now = ((demo or {}).get("population")
               or (round(roofs_now * hh_size) if roofs_now else 0))
    pop_basis = "Census ACS" if (demo or {}).get("population") else (
        "OSM building count x household size" if roofs_now else None)
    pop_buildout = pop_now + round(remaining * hh_size) if pop_now else None

    buildout = {
        "population_now": pop_now or None,
        "population_at_buildout": pop_buildout,
        "population_added": round(remaining * hh_size) if remaining else None,
        "growth_pct": (round((pop_buildout - pop_now) / pop_now * 100, 1)
                       if pop_now and pop_buildout else None),
        "rooftops_now": roofs_now,
        "rooftops_source": ("OSM building footprints" if roofs_now and roofs_now != occupied
                            else "CBAS occupied homes"),
        "cbas_occupied": occupied or None,
        "rooftops_at_buildout": ((roofs_now or 0) + remaining) or None,
        "lots_remaining": remaining or None,
        "total_lots_tracked": total_lots or None,
        "avg_household_size": hh_size,
        "household_size_source": ("Census ACS" if (demo or {}).get("avg_household_size")
                                  else "Texas average"),
        "population_basis": pop_basis,
        "years_to_buildout": (round(remaining / annual_closings, 1)
                              if remaining and annual_closings else None),
        "annual_closings": annual_closings or None,
        "demographics": demo,
        "cbas_available": cbas_ok,
    }

    # Benchmarks MEASURED from mature Houston new-home suburbs, not assumed.
    #
    # Method: six built-out reference areas (Cinco Ranch, Bridgeland, Riverstone,
    # Spring/Klein, Shadow Creek, Aliana). For each, count the same OSM
    # categories within 5 mi and divide by the OSM building footprint count in
    # the same ring. Median across the six is the benchmark; the observed range
    # is carried through so the UI can show how consistent each one actually is.
    #
    # The denominator is OSM buildings, not CBAS occupied homes. Calibrating on
    # CBAS was tried first and failed badly: it tracks only actively-selling
    # new-home communities, so Cinco Ranch came back as 405 rooftops and implied
    # one grocery store per 16 homes, with a 13x spread across references. OSM
    # carries the Microsoft US building import, which returns 44,444 structures
    # in the same ring — a real count, and it needs no Census key.
    #
    # `spread` is max/min across the six references. Low spread means the ratio
    # is stable and the gap is worth acting on; hospital at 15x is essentially
    # noise and is reported without an opportunity flag.
    BENCHMARKS = [
        ("Grocery / supermarket", "grocery", 1900, 3.1,
         "anchors a neighbourhood centre; first retail a new community needs"),
        ("Pharmacy", "pharmacy", 1800, 3.7, "usually follows the grocery anchor"),
        ("Fuel / convenience", "fuel", 550, 2.8, "earliest commercial to arrive"),
        ("Parks / open space", "park", 300, 1.9, "public or HOA; the most consistent ratio measured"),
        ("General merchandise", "general_merch", 3700, 5.7, "Target / Walmart scale"),
        ("Urgent care / clinic", "clinic", 2800, 7.1, "OSM coverage patchy; verify locally"),
        ("Home improvement", "home_improvement", 8000, 5.3, "big-box, serves several communities"),
        ("Hospital", "hospital", 7400, 15.4, "regional draw; ratio too variable to act on"),
    ]
    have = {
        "grocery": len([x for x in grocery if x.get("within_radius")]),
        "pharmacy": len([x for x in pharmacies if x.get("within_radius")]),
        "fuel": len([x for x in fuel if x.get("within_radius")]),
        "clinic": 0,          # not queried; reported as unknown rather than zero
        "hospital": len([x for x in hospitals if x.get("within_radius")]),
        "home_improvement": len([x for x in retail
                                 if x.get("kind") == "home_improvement" and x.get("within_radius")]),
        "general_merch": len([x for x in retail
                              if x.get("kind") in ("department_store", "wholesale")
                              and x.get("within_radius")]),
        "park": len([x for x in parks if x.get("within_radius")]),
    }
    opportunities = []
    roofs_bo = buildout["rooftops_at_buildout"]
    if roofs_now and roofs_bo:
        for label, key, per, spread, note in BENCHMARKS:
            sup_now = roofs_now / per
            sup_bo = roofs_bo / per
            cur = have.get(key, 0)
            gap_bo = sup_bo - cur
            # Only flag a gap worth acting on where the benchmark is stable.
            # A 15x spread across reference areas means the ratio isn't
            # describing a real requirement, so it gets reported without a flag.
            confidence = ("high" if spread <= 3.5 else
                          "medium" if spread <= 6.0 else "low")
            opportunities.append({
                "use": label, "key": key,
                "rooftops_per_facility": per,
                "spread": spread, "confidence": confidence,
                "existing": cur,
                "supported_now": round(sup_now, 1),
                "supported_at_buildout": round(sup_bo, 1),
                "gap_now": round(sup_now - cur, 1),
                "gap_at_buildout": round(gap_bo, 1),
                "opportunity": gap_bo >= 1 and confidence in ("high", "medium"),
                "unknown": key == "clinic",
                "note": note,
            })
        # Biggest shortfall first, but keep a stable order for equal gaps.
        opportunities.sort(key=lambda o: (-o["gap_at_buildout"], o["use"]))

    return jsonify({
        "centroid":            {"lat": c_lat, "lon": c_lon},
        "radius_mi":           radius_mi,
        "buildout":            buildout,
        "opportunities":       opportunities,
        "grocery_radius_mi":   grocery_radius_used,
        "hospital_radius_mi":  hospital_radius_used,
        "grocery_stores":      grocery,
        "schools":             schools,
        "hospitals":           hospitals,
        "parks":               parks,
        # Commute & access
        "highways":            highways,
        "nearest_ramp_mi":     ramp_mi,
        "park_ride":           park_ride[:4],
        "airports":            airports,
        # Retail anchors
        "retail_anchors":      retail,
        "pharmacies":          pharmacies,
        "fuel":                fuel,
        "source":              "OpenStreetMap Overpass API",
    })


@acq_bp.route("/api/acq/projects/<pid>/communities", methods=["GET"])
@_login_required
def acq_api_projects_communities(pid):
    """List residential communities / subdivisions in the project's submarket.

    Approach: pull parcels from the local StratMap cache (Houston-metro counties
    pre-cached) within the same school district as the project. Group by the
    subdivision name parsed from LEGAL_DESC. Returns the top N by parcel count.

    Useful for acquisitions: "what existing communities define this submarket,
    and which one is closest in product to what we'd build?"
    """
    import re as _re
    from shapely.geometry import shape as shp_shape, Point as _Pt
    from shapely.ops import unary_union, transform

    guard = _acq_guard()
    if guard:
        return guard
    uid = _acq_owner()
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, uid, _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404

    geoms = []
    for t in proj.get("tracts") or []:
        g = t.get("geometry")
        if g:
            try: geoms.append(shp_shape(g))
            except Exception as e: print(f"[{request.endpoint}] tract geometry skipped: {e}", flush=True)
    if not geoms:
        return jsonify({"error": "no tract geometries"}), 400
    project_union = unary_union(geoms)
    centroid = project_union.centroid
    c_lat, c_lon = centroid.y, centroid.x

    # Circle radius centered on the project centroid (default 5 mi).
    try:
        radius_mi = float(request.args.get("radius_mi", "5"))
    except (TypeError, ValueError):
        radius_mi = 5.0
    radius_mi = max(0.5, min(radius_mi, 15.0))

    # Buffer the CENTROID by the radius — proper circular submarket, project
    # included (centroid is inside the project). Not a polygon-outward buffer.
    centroid_utm = transform(to_utm, centroid)
    circle_utm = centroid_utm.buffer(radius_mi * 1609.34)
    circle_wgs = transform(to_wgs, circle_utm)

    import acq_parcels as parcel_cache
    try:
        fc = parcel_cache.query_parcels_in_polygon(circle_wgs, min_acres=0.05, max_acres=50)
        feats = fc.get("features") or []
    except Exception as e:
        return jsonify({"error": f"parcel cache query failed: {e}", "communities": []}), 500

    # Parse subdivision name from LEGAL_DESC using the shared parser (rejects
    # rural ABS abstracts, strips CAD plat codes like Waller's "S588000").
    # Compute distance from project centroid to each parcel for proximity sort.
    by_sub = {}
    for f in feats:
        p = f.get("properties") or {}
        sub = _parse_subdivision_name(p.get("LEGAL_DESC") or "")
        if not sub: continue
        acres = float(p.get("Acres") or 0)
        # Compute distance from project centroid to this parcel (miles)
        try:
            parcel_geom = shp_shape(f["geometry"])
            parcel_pt = parcel_geom.centroid
            # UTM distance for accuracy
            dist_m = transform(to_utm, parcel_pt).distance(centroid_utm)
            dist_mi = dist_m / 1609.34
        except Exception:
            dist_mi = None
        entry = by_sub.setdefault(sub, {
            "name": sub, "parcels": 0, "acres": 0.0,
            "owners": set(), "counties": set(),
            "min_dist_mi": None,
            "sum_dist_mi": 0.0,
            "builder_held": 0,        # parcels still owned by a known builder
            "builder_counts": {},     # builder name -> parcel count (for dominant)
        })
        entry["parcels"] += 1
        entry["acres"] += acres
        owner = (p.get("OWNER_NAME") or "").strip()
        if owner: entry["owners"].add(owner)
        # Absorption proxy: is this lot still held by a homebuilder?
        b = _classify_builder(owner)
        if b:
            entry["builder_held"] += 1
            entry["builder_counts"][b] = entry["builder_counts"].get(b, 0) + 1
        county = (p.get("_county") or "").strip()
        if county: entry["counties"].add(county)
        if dist_mi is not None:
            entry["sum_dist_mi"] += dist_mi
            if entry["min_dist_mi"] is None or dist_mi < entry["min_dist_mi"]:
                entry["min_dist_mi"] = dist_mi

    # Acquisitions care about REAL communities, not 6-lot splits: require 50+
    # platted lots. Tier badge: MPC (1000+), Large (300+), Community (50+).
    MIN_PARCELS = 50
    def _size_tier(n):
        if n >= 1000: return "MPC"
        if n >= 300:  return "Large"
        return "Community"
    communities = []
    for v in by_sub.values():
        if v["parcels"] < MIN_PARCELS: continue
        avg_dist = (v["sum_dist_mi"] / v["parcels"]) if v["parcels"] > 0 else None
        builder_pct = round(v["builder_held"] / v["parcels"] * 100, 1) if v["parcels"] else 0
        # Dominant builder (most parcels held) if any
        dom = max(v["builder_counts"].items(), key=lambda kv: kv[1])[0] if v["builder_counts"] else None
        # Lifecycle heuristic from builder-held %: active-selling communities
        # still have builder-owned inventory; sold-out ones are ~all individuals.
        if builder_pct >= 30:   status = "Actively selling"
        elif builder_pct >= 5:  status = "Late phase"
        else:                   status = "Sold out / resale"
        avg_lot_ac = (v["acres"] / v["parcels"]) if v["parcels"] else 0
        avg_lot_sqft = avg_lot_ac * 43560
        # Homebuilder convention: express lot size as FRONTAGE FEET assuming a
        # standard 120-ft depth (FF = sqft / 120). 6,000 sf -> 50 FF, 7,200 -> 60 FF.
        est_ff = avg_lot_sqft / 120.0 if avg_lot_sqft > 0 else 0
        communities.append({
            "name":              v["name"],
            "tier":              _size_tier(v["parcels"]),
            "parcels":           v["parcels"],
            "acres":             round(v["acres"], 1),
            "avg_lot_acres":     round(avg_lot_ac, 3),
            "avg_lot_sqft":      int(round(avg_lot_sqft)),
            "est_ff":            int(round(est_ff)),
            "counties":          sorted(v["counties"]),
            "owner_sample":      sorted(v["owners"])[:5],
            "owner_diversity":   len(v["owners"]),
            "builder_held":      v["builder_held"],
            "builder_held_pct":  builder_pct,
            "dominant_builder":  dom,
            "status":            status,
            "closest_mi":        round(v["min_dist_mi"], 2) if v["min_dist_mi"] is not None else None,
            "avg_dist_mi":       round(avg_dist, 2) if avg_dist is not None else None,
        })
    # ---- OSM named communities ------------------------------------------
    # CRITICAL for Harris County: StratMap leaves LEGAL_DESC EMPTY on Harris
    # home lots (verified against the local cache), so parcel-based detection
    # finds nothing there. OSM has named place nodes + named residential
    # landuse polygons for real communities (Bridgeland, Towne Lake, etc.) —
    # verified live: 300 named elements within 5 mi of Cypress.
    import math as _math
    radius_m = int(radius_mi * 1609.34)
    osm_q = f"""
[out:json][timeout:25];
(
  node["place"~"^(neighbourhood|suburb|quarter|village|hamlet)$"](around:{radius_m},{c_lat},{c_lon});
  way["landuse"="residential"]["name"](around:{radius_m},{c_lat},{c_lon});
  relation["landuse"="residential"]["name"](around:{radius_m},{c_lat},{c_lon});
);
out center 400;
"""
    osm_by_name = {}
    osm_data, _osm_err = _overpass_query(osm_q, 40)
    if osm_data:
        for el in osm_data.get("elements", []):
            tags = el.get("tags") or {}
            nm = (tags.get("name") or "").strip()
            if not nm or len(nm) < 3: continue
            lat = el.get("lat") or (el.get("center") or {}).get("lat")
            lon = el.get("lon") or (el.get("center") or {}).get("lon")
            if lat is None or lon is None: continue
            d_mi = _math.hypot((lon - c_lon) * 69.0 * _math.cos(_math.radians(c_lat)),
                                (lat - c_lat) * 69.0)
            key = nm.upper()
            e = osm_by_name.setdefault(key, {"name": nm, "closest_mi": d_mi, "sections": 0})
            e["sections"] += 1
            if d_mi < e["closest_mi"]:
                e["closest_mi"] = d_mi

    # ---- Merge: OSM names (authoritative for existence/distance) enriched
    # with parcel metrics (lots, FF, builder-held) where the names match.
    parcel_by_name = {c["name"].upper(): c for c in communities}
    merged = []
    for key, osm in osm_by_name.items():
        pc = parcel_by_name.pop(key, None)
        row = {
            "name":              osm["name"],
            "closest_mi":        round(osm["closest_mi"], 2),
            "source":            "OSM+parcels" if pc else "OSM",
            "tier":              (pc or {}).get("tier"),
            "parcels":           (pc or {}).get("parcels"),
            "est_ff":            (pc or {}).get("est_ff"),
            "avg_lot_sqft":      (pc or {}).get("avg_lot_sqft"),
            "builder_held_pct":  (pc or {}).get("builder_held_pct"),
            "dominant_builder":  (pc or {}).get("dominant_builder"),
            "status":            (pc or {}).get("status"),
            "counties":          (pc or {}).get("counties") or [],
        }
        merged.append(row)
    # Parcel-only communities (counties where LEGAL_DESC works but OSM lacks the name)
    for pc in parcel_by_name.values():
        merged.append({**pc, "source": "parcels"})
    merged.sort(key=lambda x: (x["closest_mi"] if x["closest_mi"] is not None else 9999))
    top = merged[:40]

    return jsonify({
        "communities":     top,
        "total_distinct":  len(merged),
        "radius_mi":       radius_mi,
        "parcels_scanned": len(feats),
        "min_parcels":     MIN_PARCELS,
        "centroid":        {"lat": c_lat, "lon": c_lon},
        "osm_found":       len(osm_by_name),
        "note":            (f"Community names from OpenStreetMap (place nodes + named residential areas) "
                             f"merged with parcel-cache metrics where legal descriptions carry subdivision names. "
                             f"Harris County home lots have blank legal descriptions in StratMap, so Harris rows "
                             f"show name + distance only. Lot metrics require {MIN_PARCELS}+ platted parcels."),
    })


@acq_bp.route("/api/acq/projects/<pid>/submarket", methods=["GET"])
@_login_required
def acq_api_projects_submarket(pid):
    """Submarket-level data for the project — built for acquisitions thinking
    at the school-district / ZIP / Census tract level (county is too broad —
    Harris County alone is 1,700 sq mi covering radically different markets).

    Returns:
      • Containing-tract list (the Census tracts the project sits in)
      • Neighboring tracts within a configurable radius (?radius_mi=3 default)
      • Population-weighted aggregates across containing tracts + same-school-district neighbors
      • School districts (Unified/Secondary/Elementary) from Census Geocoder
      • ZIP codes the project overlaps
      • Census Designated Place (CDP) / incorporated city the project is in
      • Tract polygons so the client can render them on the map

    Query string:
      ?radius_mi=3   — neighborhood radius (default 3, max 15)
    """
    import requests as _rq
    from shapely.geometry import shape as shp_shape, Point as _Pt
    from shapely.ops import unary_union, transform

    guard = _acq_guard()
    if guard:
        return guard
    uid = _acq_owner()
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, uid, _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404

    try:
        radius_mi = float(request.args.get("radius_mi", "3"))
    except (TypeError, ValueError):
        radius_mi = 3.0
    radius_mi = max(0.5, min(radius_mi, 15.0))

    geoms = []
    for t in proj.get("tracts") or []:
        g = t.get("geometry")
        if g:
            try: geoms.append(shp_shape(g))
            except Exception as e: print(f"[{request.endpoint}] tract geometry skipped: {e}", flush=True)
    if not geoms:
        return jsonify({"error": "no tract geometries"}), 400
    proj_union = unary_union(geoms)
    centroid = proj_union.centroid
    c_lon, c_lat = centroid.x, centroid.y

    # Build a 3-mi (or user-specified) buffer around the project for "neighboring" tracts.
    # Buffer in UTM so radius is in true meters.
    buffer_utm = transform(to_utm, proj_union).buffer(radius_mi * 1609.34)
    buffer_wgs = transform(to_wgs, buffer_utm)

    out = {
        "centroid":         {"lat": c_lat, "lon": c_lon},
        "radius_mi":        radius_mi,
        "tracts":           [],        # containing tracts
        "neighbor_tracts":  [],        # tracts within radius_mi (excluding containing)
        "tract_summary":    {},        # weighted aggregates over containing tracts
        "wider_summary":    {},        # weighted aggregates over containing + neighbor
        "school_districts": [],
        "zip_codes":        [],
        "place":            None,      # incorporated city or CDP
        "sources":          [],
        "caveats":          [],
    }

    # --- 1) Identify Census TRACT via the Census Geocoder (free, no key) ---
    census_key = os.environ.get("CENSUS_API_KEY", "").strip()
    tract_fips_list = []   # list of (state, county, tract) tuples
    try:
        geocoder_url = "https://geocoding.geo.census.gov/geocoder/geographies/coordinates"
        gr = _rq.get(geocoder_url, params={
            "x": c_lon, "y": c_lat,
            "benchmark": "Public_AR_Current",
            "vintage":   "Current_Current",
            "format":    "json",
        }, timeout=12)
        gr.raise_for_status()
        geo = gr.json().get("result", {}).get("geographies", {})
        # Census Tracts
        for t in (geo.get("Census Tracts") or []):
            tract_fips_list.append((str(t.get("STATE")).zfill(2),
                                     str(t.get("COUNTY")).zfill(3),
                                     str(t.get("TRACT")).zfill(6)))
        # Unified School Districts
        for sd in (geo.get("Unified School Districts") or []):
            out["school_districts"].append({
                "name":  sd.get("NAME"),
                "geoid": sd.get("GEOID"),
                "type":  "Unified",
            })
        # Secondary (high school) districts — some Texas areas have separate ones
        for sd in (geo.get("Secondary School Districts") or []):
            out["school_districts"].append({
                "name":  sd.get("NAME"),
                "geoid": sd.get("GEOID"),
                "type":  "Secondary",
            })
        # Elementary
        for sd in (geo.get("Elementary School Districts") or []):
            out["school_districts"].append({
                "name":  sd.get("NAME"),
                "geoid": sd.get("GEOID"),
                "type":  "Elementary",
            })
        # ZIP code(s) — Census ZCTAs (Zip Code Tabulation Areas)
        for zc in (geo.get("Zip Code Tabulation Areas") or []):
            out["zip_codes"].append({
                "zip":   zc.get("BASENAME") or zc.get("ZCTA5") or zc.get("NAME"),
                "geoid": zc.get("GEOID"),
            })
        # Census Place (incorporated city OR CDP)
        for pl in (geo.get("Incorporated Places") or []):
            out["place"] = {"name": pl.get("NAME"), "type": "Incorporated City",
                             "geoid": pl.get("GEOID")}
            break
        if not out["place"]:
            for pl in (geo.get("Census Designated Places") or []):
                out["place"] = {"name": pl.get("NAME"), "type": "CDP (unincorporated)",
                                 "geoid": pl.get("GEOID")}
                break
        out["sources"].append("Census Bureau Geocoder (tract/school-district/ZIP/place lookup)")
    except Exception as e:
        out["caveats"].append(f"Census geocoder unavailable: {e}")

    if not tract_fips_list:
        out["caveats"].append("No census tracts found for project centroid.")

    # --- 2) Pull tract-level ACS data for each containing tract ---
    YEAR = _acs_year()
    TRACT_VARS = {
        "B01003_001E": "population",
        "B01002_001E": "median_age",
        "B11001_001E": "households",
        "B25010_001E": "avg_household_size",
        "B19013_001E": "median_household_income",
        "B19301_001E": "per_capita_income",
        "B25001_001E": "housing_units",
        "B25002_002E": "occupied_units",
        "B25002_003E": "vacant_units",
        "B25003_002E": "owner_occupied",
        "B25003_003E": "renter_occupied",
        "B25077_001E": "median_home_value",
        "B25064_001E": "median_rent",
        "B23025_004E": "employed",
        "B15003_022E": "bachelors_only",
        "B15003_023E": "masters",
        "B15003_024E": "professional",
        "B15003_025E": "doctorate",
        "B15003_001E": "education_universe",
        "B25034_001E": "year_built_universe",
        "B25035_001E": "median_year_built",
    }

    def _int_or_none(v):
        try:
            if v is None or v == "" or str(v).startswith("-666666"): return None
            return int(v)
        except (ValueError, TypeError): return None

    for state_fips, county_fips, tract_fips in tract_fips_list:
        try:
            var_list = ",".join(["NAME"] + list(TRACT_VARS.keys()))
            params = {
                "get": var_list,
                "for": f"tract:{tract_fips}",
                "in":  f"state:{state_fips}+county:{county_fips}",
            }
            if census_key:
                params["key"] = census_key
            r = _rq.get(f"https://api.census.gov/data/{YEAR}/acs/acs5", params=params, timeout=15)
            if r.status_code != 200:
                out["caveats"].append(f"Tract {tract_fips}: Census HTTP {r.status_code}")
                continue
            rows = r.json()
            if not isinstance(rows, list) or len(rows) < 2:
                continue
            headers, data = rows[0], rows[1]
            rec = dict(zip(headers, data))
            tract_data = {label: _int_or_none(rec.get(var)) for var, label in TRACT_VARS.items()}
            # Derived ratios per tract
            edu_tot = sum(filter(None, [tract_data.get(k) for k in
                                          ("bachelors_only", "masters", "professional", "doctorate")]))
            edu_uni = tract_data.get("education_universe") or 0
            tract_data["pct_bachelors_or_higher"] = round(edu_tot / edu_uni * 100, 1) if edu_uni else None
            oo, occ = tract_data.get("owner_occupied") or 0, tract_data.get("occupied_units") or 0
            tract_data["owner_occupancy_pct"] = round(oo / occ * 100, 1) if occ else None
            vu, hu = tract_data.get("vacant_units") or 0, tract_data.get("housing_units") or 0
            tract_data["vacancy_pct"] = round(vu / hu * 100, 1) if hu else None
            out["tracts"].append({
                "fips": f"{state_fips}{county_fips}{tract_fips}",
                "name": rec.get("NAME"),
                "data": tract_data,
            })
        except Exception as e:
            out["caveats"].append(f"Tract {tract_fips} fetch failed: {e}")

    if out["tracts"]:
        out["sources"].append(f"Census ACS 5-year tract-level ({YEAR})")

    # --- 2.5) Find ALL school districts the project POLYGON intersects, not just
    # the centroid's district. Critical fix — a project that straddles a district
    # boundary needs both districts counted, otherwise the submarket excludes the
    # other half of the project's own market.
    sd_polygons = []
    try:
        b = proj_union.bounds
        # Query school districts spatially over the project bbox.
        # NOTE: the TIGERweb service is named "School" (verified live) — layers
        # 0/1/2 = current-vintage Unified / Secondary / Elementary districts.
        for layer_id in [0, 1, 2]:   # Unified / Secondary / Elementary
            try:
                tu = f"https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb/School/MapServer/{layer_id}/query"
                tr = _rq.get(tu, params={
                    "geometry":          f"{b[0]},{b[1]},{b[2]},{b[3]}",
                    "geometryType":      "esriGeometryEnvelope",
                    "inSR":              "4326",
                    "spatialRel":        "esriSpatialRelIntersects",
                    "outFields":         "GEOID,NAME",
                    "returnGeometry":    "true",
                    "outSR":             "4326",
                    "f":                 "geojson",
                }, timeout=15)
                if tr.status_code != 200: continue
                fc = tr.json()
                for feat in (fc.get("features") or []):
                    geom = feat.get("geometry")
                    if not geom: continue
                    try:
                        shp = shp_shape(geom)
                        if not shp.intersects(proj_union): continue
                        gid = (feat.get("properties") or {}).get("GEOID")
                        name = (feat.get("properties") or {}).get("NAME")
                        if any(p["geoid"] == gid for p in sd_polygons): continue
                        sd_polygons.append({
                            "name":     name,
                            "geoid":    gid,
                            "geometry": geom,
                            "shape":    shp,
                        })
                    except Exception: continue
            except Exception: continue
    except Exception as e:
        out["caveats"].append(f"School district spatial lookup failed: {e}")

    # Replace the centroid-based school_districts list with the spatial-intersect
    # result so the UI shows ALL districts the project actually overlaps.
    if sd_polygons:
        out["school_districts"] = [
            {"name": p["name"], "geoid": p["geoid"], "type": "Unified"}
            for p in sd_polygons
        ]

    # Return the school district polygons so client can render them on the map
    out["school_district_polygons"] = [
        {"name": p["name"], "geoid": p["geoid"], "geometry": p["geometry"]}
        for p in sd_polygons
    ]

    # --- 2.6) Discover NEIGHBORING census tracts within `radius_mi` of the project ---
    # Use Census TIGER REST API to find all tract polygons intersecting the buffer.
    # Then filter to ONLY those whose centroid falls inside the project's school
    # district(s) — that's the actual residential submarket.
    containing_keys = {(t["fips"][:2], t["fips"][2:5], t["fips"][5:])
                        for t in out["tracts"]}
    try:
        tigerweb = ("https://tigerweb.geo.census.gov/arcgis/rest/services/"
                     "TIGERweb/Tracts_Blocks/MapServer/0/query")
        b = buffer_wgs.bounds
        tr = _rq.get(tigerweb, params={
            "where":             "1=1",
            "geometry":          f"{b[0]},{b[1]},{b[2]},{b[3]}",
            "geometryType":      "esriGeometryEnvelope",
            "inSR":              "4326",
            "spatialRel":        "esriSpatialRelIntersects",
            "outFields":         "STATE,COUNTY,TRACT,GEOID,BASENAME",
            "returnGeometry":    "true",
            "outSR":             "4326",
            "f":                 "geojson",
        }, timeout=15)
        tr.raise_for_status()
        nearby_fc = tr.json()
        neighbor_features = []
        for feat in nearby_fc.get("features", []):
            props = feat.get("properties") or {}
            st_f  = str(props.get("STATE") or "").zfill(2)
            co_f  = str(props.get("COUNTY") or "").zfill(3)
            tr_f  = str(props.get("TRACT") or "").zfill(6)
            key   = (st_f, co_f, tr_f)
            # Verify the tract polygon ACTUALLY intersects our buffer (Esri's bbox
            # filter returns anything inside the bbox even if not the polygon).
            try:
                tract_geom = shp_shape(feat["geometry"])
                if not tract_geom.intersects(buffer_wgs): continue
            except Exception:
                continue
            if key in containing_keys:
                fips = f"{st_f}{co_f}{tr_f}"
                for containing in out["tracts"]:
                    if containing.get("fips") == fips and not containing.get("geometry"):
                        containing["geometry"] = feat.get("geometry")
                        break
                continue   # already in containing set
            # School-district filter: only count if the tract's centroid falls
            # inside any of the project's school district polygons. If we have
            # NO school district polygons (geocoder failed), keep all neighbors.
            in_district = True
            if sd_polygons:
                tract_centroid = tract_geom.centroid
                in_district = any(p["shape"].contains(tract_centroid) for p in sd_polygons)
            neighbor_features.append((key, props, feat.get("geometry"), tract_geom, in_district))
        # Pull ACS for each neighbor (cap at 50 to avoid blowing up the Census API)
        for (st_f, co_f, tr_f), props, geom, _shp_geom, in_district in neighbor_features[:50]:
            try:
                var_list = ",".join(["NAME"] + list(TRACT_VARS.keys()))
                params = {
                    "get": var_list,
                    "for": f"tract:{tr_f}",
                    "in":  f"state:{st_f}+county:{co_f}",
                }
                if census_key:
                    params["key"] = census_key
                r = _rq.get(f"https://api.census.gov/data/{YEAR}/acs/acs5", params=params, timeout=15)
                if r.status_code != 200: continue
                rows = r.json()
                if not isinstance(rows, list) or len(rows) < 2: continue
                headers, data = rows[0], rows[1]
                rec = dict(zip(headers, data))
                td = {label: _int_or_none(rec.get(var)) for var, label in TRACT_VARS.items()}
                edu_tot = sum(filter(None, [td.get(k) for k in
                                              ("bachelors_only", "masters", "professional", "doctorate")]))
                edu_uni = td.get("education_universe") or 0
                td["pct_bachelors_or_higher"] = round(edu_tot / edu_uni * 100, 1) if edu_uni else None
                oo, occ = td.get("owner_occupied") or 0, td.get("occupied_units") or 0
                td["owner_occupancy_pct"] = round(oo / occ * 100, 1) if occ else None
                vu, hu = td.get("vacant_units") or 0, td.get("housing_units") or 0
                td["vacancy_pct"] = round(vu / hu * 100, 1) if hu else None
                out["neighbor_tracts"].append({
                    "fips":             f"{st_f}{co_f}{tr_f}",
                    "name":             rec.get("NAME"),
                    "data":             td,
                    "geometry":         geom,
                    "in_school_district": in_district,
                })
            except Exception: continue
        if out["neighbor_tracts"]:
            out["sources"].append(f"TIGERweb tract polygons + ACS for {len(out['neighbor_tracts'])} neighbor tracts")
    except Exception as e:
        out["caveats"].append(f"Neighbor tract search failed: {e}")

    # --- 3) Aggregated tract summaries (population-weighted means) ---
    def _summarize(tract_list):
        if not tract_list: return {}
        sums = {}
        pop_weighted = {}
        total_pop = 0
        for t in tract_list:
            d = t["data"]
            pop = d.get("population") or 0
            total_pop += pop
            for k, v in d.items():
                if v is None: continue
                if k in ("population", "households", "housing_units", "occupied_units",
                         "vacant_units", "owner_occupied", "renter_occupied", "employed"):
                    sums[k] = sums.get(k, 0) + v
                elif pop > 0:
                    pop_weighted.setdefault(k, []).append((pop, v))
        summary = dict(sums)
        for k, pairs in pop_weighted.items():
            total_w = sum(p for p, v in pairs)
            if total_w:
                summary[k] = round(sum(p * v for p, v in pairs) / total_w, 2)
        summary["tract_count"] = len(tract_list)
        summary["total_population"] = total_pop
        return summary

    out["tract_summary"] = _summarize(out["tracts"])
    # Wider summary now respects school-district filter: containing tracts
    # (always counted) + neighbors INSIDE the same district. If we have no
    # district polygons, fall back to all neighbors so we still show something.
    in_district_neighbors = [t for t in out["neighbor_tracts"]
                              if t.get("in_school_district", True)]
    out["wider_summary"] = _summarize(out["tracts"] + in_district_neighbors)
    out["filter_applied"] = ("same school district" if sd_polygons else
                              f"within {radius_mi} mi (no school district detected)")
    out["wider_tract_count_in_district"] = len(in_district_neighbors) + len(out["tracts"])
    out["wider_tract_count_total"] = len(out["neighbor_tracts"]) + len(out["tracts"])

    # --- 4) Caveats for what we DON'T have ---
    out["caveats"].extend([
        "Submarket new-home metrics (absorption, avg pricing/SF by community) require a Zonda or similar subscription.",
        "Census ACS tract data refreshes annually (5-year estimates); HAR MLS data would be more current at submarket level.",
    ])
    return jsonify(out)


@acq_bp.route("/api/acq/projects/<pid>/market", methods=["GET"])
@_login_required
def acq_api_projects_market(pid):
    """Pull free public market data around the project: population, employment,
    median income, age distribution from the Census Bureau ACS 5-year API.

    No API key required for Census. Identifies the primary county the project
    sits in (via the cached counties layer) and fetches county-level stats.
    """
    import requests as _rq
    from shapely.geometry import shape as shp_shape
    from shapely.ops import unary_union

    guard = _acq_guard()
    if guard:
        return guard
    uid = _acq_owner()
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, uid, _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404

    # Use project union centroid to identify the county
    geoms = []
    for t in proj.get("tracts") or []:
        g = t.get("geometry")
        if g:
            try: geoms.append(shp_shape(g))
            except Exception as e: print(f"[{request.endpoint}] tract geometry skipped: {e}", flush=True)
    if not geoms:
        return jsonify({"error": "no tract geometries"}), 400
    proj_union = unary_union(geoms)
    centroid = proj_union.centroid

    # Find county FIPS by spatial lookup against the counties cached layer
    try:
        counties_fc = cached_layer_query("counties", proj_union)
        county_name = None
        county_fips = None
        for f in counties_fc.get("features", []):
            try:
                if shp_shape(f["geometry"]).contains(centroid):
                    props = f.get("properties") or {}
                    county_name = props.get("BASENAME") or props.get("NAME")
                    county_fips = props.get("GEOID") or props.get("COUNTYFP")
                    break
            except Exception: continue
    except Exception as e:
        return jsonify({"error": f"county lookup failed: {e}"}), 500

    if not county_fips:
        return jsonify({"error": "Could not identify county for this project"}), 404

    # Normalize FIPS — Census ACS needs state + county codes separately
    # county_fips might be 5-digit (state+county) or 3-digit (county only)
    fips_str = str(county_fips).zfill(5) if len(str(county_fips)) <= 5 else str(county_fips)
    if len(fips_str) == 5:
        state_fips = fips_str[:2]
        county_fips_only = fips_str[2:]
    else:
        # Default to Texas (48) if we only got county
        state_fips = "48"
        county_fips_only = fips_str.zfill(3)

    # Census ACS 5-year — most recent available. Expanded to 16 variables
    # covering population/income/housing/education/age structure.
    YEAR = _acs_year()
    VARS = {
        # --- Population / households ---
        "B01003_001E": "population",
        "B01002_001E": "median_age",
        "B11001_001E": "households",
        "B11016_001E": "family_households",
        "B25010_001E": "avg_household_size",
        # --- Income ---
        "B19013_001E": "median_household_income",
        "B19301_001E": "per_capita_income",
        # --- Housing stock & tenure ---
        "B25001_001E": "housing_units",
        "B25002_002E": "occupied_units",
        "B25002_003E": "vacant_units",
        "B25003_002E": "owner_occupied",
        "B25003_003E": "renter_occupied",
        "B25077_001E": "median_home_value",
        "B25064_001E": "median_rent",
        # --- Labor ---
        "B23025_004E": "employed",
        # --- Education (% bachelor's+ derived from age 25+ population) ---
        "B15003_022E": "bachelors_only",
        "B15003_023E": "masters",
        "B15003_024E": "professional",
        "B15003_025E": "doctorate",
        "B15003_001E": "education_universe",
    }
    url = f"https://api.census.gov/data/{YEAR}/acs/acs5"
    var_list = ",".join(["NAME"] + list(VARS.keys()))
    try:
        census_params = {
            "get": var_list,
            "for": f"county:{county_fips_only}",
            "in":  f"state:{state_fips}",
        }
        # Census now enforces keys above ~500 requests/day per IP.
        # CENSUS_API_KEY is free at https://api.census.gov/data/key_signup.html
        census_key = os.environ.get("CENSUS_API_KEY", "").strip()
        if census_key:
            census_params["key"] = census_key
        r = _rq.get(url, params=census_params, timeout=15)
        if r.status_code != 200:
            # Census returns 204 No Content for missing-data lookups, or HTML
            # error pages for malformed queries. Surface a useful message
            # instead of letting r.json() blow up with "Expecting value".
            return jsonify({"error": f"Census ACS HTTP {r.status_code} for state={state_fips} county={county_fips_only}. "
                                       f"Body: {r.text[:200]}"}), 502
        try:
            rows = r.json()
        except ValueError:
            return jsonify({"error": f"Census ACS returned non-JSON for state={state_fips} county={county_fips_only}. "
                                       f"Body: {r.text[:200]}"}), 502
        if not isinstance(rows, list) or len(rows) < 2:
            return jsonify({"error": f"Census ACS returned no data for state={state_fips} county={county_fips_only}. "
                                       f"This is normal for some smaller counties.",
                            "year": YEAR}), 502
        headers, data = rows[0], rows[1]
        record = dict(zip(headers, data))
        def _to_int(v):
            try:
                # Handle negative numbers (e.g. "-666666666" Census null sentinels)
                if v is None or v == "" or str(v).startswith("-666666"): return None
                return int(v)
            except (ValueError, TypeError): return None
        current = {label: _to_int(record.get(var)) for var, label in VARS.items()}
        # Derived: bachelor's degree or higher (%) — sum of bachelor's/master's/
        # professional/doctorate, divided by the universe (age 25+ population).
        edu_total = sum(filter(None, [current.get(k) for k in
                                       ("bachelors_only", "masters", "professional", "doctorate")]))
        edu_uni = current.get("education_universe") or 0
        current["pct_bachelors_or_higher"] = round(edu_total / edu_uni * 100, 1) if edu_uni else None
        # Derived: owner-occupancy rate
        oo = current.get("owner_occupied") or 0
        occ = current.get("occupied_units") or 0
        current["owner_occupancy_pct"] = round(oo / occ * 100, 1) if occ else None
        # Derived: vacancy rate
        vu = current.get("vacant_units") or 0
        hu = current.get("housing_units") or 0
        current["vacancy_pct"] = round(vu / hu * 100, 1) if hu else None
        current["county_name"] = record.get("NAME")
        current["year"] = YEAR
    except Exception as e:
        return jsonify({"error": f"Census ACS fetch failed: {type(e).__name__}: {e}"}), 502

    # 5-year history for growth trends
    history = []
    for yr in range(YEAR - 4, YEAR + 1):
        try:
            hist_params = {
                "get": "B01003_001E,B23025_004E,B25077_001E",   # pop, employed, home value
                "for": f"county:{county_fips_only}",
                "in":  f"state:{state_fips}",
            }
            if census_key:
                hist_params["key"] = census_key
            r = _rq.get(f"https://api.census.gov/data/{yr}/acs/acs5", params=hist_params, timeout=10)
            if r.status_code == 200:
                rows = r.json()
                if len(rows) >= 2:
                    d = dict(zip(rows[0], rows[1]))
                    history.append({
                        "year":       yr,
                        "population": int(d["B01003_001E"]) if d.get("B01003_001E", "").lstrip("-").isdigit() else None,
                        "employed":   int(d["B23025_004E"]) if d.get("B23025_004E", "").lstrip("-").isdigit() else None,
                        "home_value": int(d["B25077_001E"]) if d.get("B25077_001E", "").lstrip("-").isdigit() else None,
                    })
        except Exception: continue

    # Compute YoY growth rates from history
    growth = {}
    if len(history) >= 2:
        first, last = history[0], history[-1]
        years_span = last["year"] - first["year"]
        if years_span > 0:
            for k in ("population", "employed", "home_value"):
                if first.get(k) and last.get(k) and first[k] > 0:
                    cagr = ((last[k] / first[k]) ** (1 / years_span) - 1) * 100
                    growth[k + "_cagr_pct"] = round(cagr, 2)
                    growth[k + "_total_pct"] = round((last[k] / first[k] - 1) * 100, 1)

    return jsonify({
        "county_name": county_name,
        "county_fips": fips_str,
        "current": current,
        "history": history,
        "growth": growth,
        "sources": [
            f"Census Bureau ACS 5-year ({YEAR}, with {YEAR-4}-{YEAR} history)",
        ],
    })


@acq_bp.route("/api/acq/projects/<pid>/fred", methods=["GET"])
@_login_required
def acq_api_projects_fred(pid):
    """Pull Houston-MSA + Texas + national economic indicators from FRED.

    Requires FRED_API_KEY env var. Returns each series with current value,
    prior-year value, 5-year CAGR, and a short time series for sparkline rendering.
    """
    import requests as _rq

    api_key = os.environ.get("FRED_API_KEY", "").strip()
    if not api_key:
        return jsonify({"error": "FRED_API_KEY not configured on the server. "
                                  "Add it in Railway Variables (free at fred.stlouisfed.org)."
                        }), 503

    guard = _acq_guard()
    if guard:
        return guard
    uid = _acq_owner()
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, uid, _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404

    def fetch_series(series_id):
        """Return list of recent observations + computed stats."""
        try:
            r = _rq.get("https://api.stlouisfed.org/fred/series/observations",
                        params={
                            "series_id":            series_id,
                            "api_key":              api_key,
                            "file_type":            "json",
                            "sort_order":           "desc",
                            "limit":                60,    # ~5 years of monthly data
                        }, timeout=12)
            r.raise_for_status()
            obs = r.json().get("observations", [])
            obs = [{"date": o["date"], "value": (float(o["value"]) if o["value"] not in (".", None) else None)}
                    for o in obs]
            obs = [o for o in obs if o["value"] is not None]
            if not obs:
                return {"error": "no observations"}
            current = obs[0]
            # Find a value ~1 year ago and ~5 years ago for trend
            from datetime import datetime as _dt
            current_dt = _dt.strptime(current["date"], "%Y-%m-%d")
            year_ago = next((o for o in obs
                              if (current_dt - _dt.strptime(o["date"], "%Y-%m-%d")).days >= 360), None)
            five_year_ago = next((o for o in obs
                                   if (current_dt - _dt.strptime(o["date"], "%Y-%m-%d")).days >= 1800), None)

            stats = {"current": current["value"], "current_date": current["date"]}
            if year_ago:
                stats["yoy_pct"] = round((current["value"] / year_ago["value"] - 1) * 100, 2) \
                                    if year_ago["value"] != 0 else None
                stats["year_ago_value"] = year_ago["value"]
                stats["year_ago_date"] = year_ago["date"]
            if five_year_ago and five_year_ago["value"] > 0:
                years = (current_dt - _dt.strptime(five_year_ago["date"], "%Y-%m-%d")).days / 365.25
                if years > 0:
                    stats["five_year_cagr_pct"] = round(
                        ((current["value"] / five_year_ago["value"]) ** (1/years) - 1) * 100, 2)
                    stats["five_year_total_pct"] = round(
                        (current["value"] / five_year_ago["value"] - 1) * 100, 1)

            # Sparkline: last 24 observations in chronological order (oldest first)
            sparkline = list(reversed(obs[:24]))
            stats["sparkline"] = [{"date": o["date"], "value": o["value"]} for o in sparkline]
            return stats
        except Exception as e:
            return {"error": str(e)[:200]}

    out = []
    # Run fetches in parallel — FRED limits to 120/min, we're doing 8.
    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = {pool.submit(fetch_series, sid): (sid, label, fmt)
                    for sid, label, fmt in _FRED_HOUSTON_SERIES}
        for fut in futures:
            sid, label, fmt = futures[fut]
            try:
                data = fut.result()
            except Exception as e:
                data = {"error": str(e)[:200]}
            out.append({
                "series_id": sid,
                "label":     label,
                "format":    fmt,
                "data":      data,
            })
    return jsonify({
        "indicators": out,
        "source":     "Federal Reserve Bank of St. Louis (FRED)",
        "msa":        "Houston-The Woodlands-Sugar Land MSA (26420)",
        "notes":      "All FRED series IDs included for verification — query directly on fred.stlouisfed.org.",
    })


# ══════════════════════════════════════════════════════════════════════════
# Executive acquisition summary (HTML -> WeasyPrint)
#
# The layout lives in templates/acq_report.html as ordinary HTML and CSS.
# Nothing here recomputes the analysis: the payloads come from the very view
# functions the project page already calls, so a number in the report and the
# same number on screen cannot drift apart.
# ══════════════════════════════════════════════════════════════════════════

def _report_last_frame(tb):
    """The last line of our own code in a traceback, for the error message.

    A user staring at an alert box needs "acq_report.py:812 in _map_comps",
    not the exception text on its own.
    """
    ours = [l.strip() for l in (tb or "").splitlines()
            if l.strip().startswith("File ") and ("acq_report" in l or "acq_routes" in l)]
    return ours[-1][:180] if ours else None


def _report_payloads(pid, want=("cbas", "roads", "amenities", "news", "market", "fred")):
    """Call the existing section endpoints and keep whatever answers.

    Each one is optional. A section whose service is unconfigured or down is
    left out of the context entirely and the template drops it, rather than
    printing an empty card -- which in a document going to an investment
    committee would read as "there is nothing here" rather than "we could not
    reach the source".
    """
    fetchers = {
        "cbas": acq_api_projects_cbas, "roads": acq_api_projects_roads,
        "amenities": acq_api_projects_amenities, "news": acq_api_projects_news,
        "market": acq_api_projects_market, "fred": acq_api_projects_fred,
    }
    out, failed = {}, []
    for key in want:
        fn = fetchers.get(key)
        if fn is None:
            continue
        try:
            resp = fn(pid)
            payload = resp[0] if isinstance(resp, tuple) else resp
            data = payload.get_json() if hasattr(payload, "get_json") else payload
            if isinstance(data, dict) and data.get("error"):
                failed.append(f"{key}: {str(data['error'])[:90]}")
                continue
            out[key] = data
        except Exception as e:
            failed.append(f"{key}: {type(e).__name__} {str(e)[:70]}")
    out["_failed"] = failed
    return out


def _build_report_context(pid):
    import acq_report
    from shapely.geometry import shape as shp_shape
    from shapely.ops import unary_union

    conn = get_db()
    try:
        # The admin/include-all flag matters: without it this lookup is
        # stricter than the one behind the project page itself, so a project an
        # admin can open and analyse reported "project not found" the moment it
        # was asked for as a report.
        proj = acq_store.get_object(conn, "project", pid, _acq_owner(),
                                    _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return None, ("Project not found, or not visible to this account. "
                      "Open it from the Analyses list and use the ID in that "
                      "page's address.")
    analysis = proj.get("analysis_cache")
    if not analysis:
        return None, ("Run the acquisition analysis first -- the report is built "
                      "from it.")
    # Same hydration the project page gets, so the tract-composition table and
    # the gross acreage above it cannot quote different figures.
    proj["tracts"], _ = parcel_cache.hydrate_tract_acres(proj.get("tracts") or [])

    data = _report_payloads(pid)

    # Schools take two hops, the same two the project page takes: /submarket
    # resolves which district(s) the site sits in and the project centroid,
    # then the district profile is fetched for the first of them.
    centroid = None
    try:
        sm = acq_api_projects_submarket(pid)
        sm = sm[0] if isinstance(sm, tuple) else sm
        sm = sm.get_json() if hasattr(sm, "get_json") else sm
        centroid = (sm or {}).get("centroid")
        districts = ((sm or {}).get("school_districts") or []) if centroid else []
        # A tract can straddle a district line, and the app's own schools card
        # shows every district it touches. Taking districts[0] silently threw
        # the others away, so the export disagreed with the screen.
        got = []
        for d0 in districts[:3]:
            # test_request_context starts with an EMPTY session, so
            # @_login_required on the district endpoint rejected the call and
            # the schools section vanished from every report without ever
            # raising. Carry the caller's session into the inner context.
            outer_session = dict(session)
            with current_app.test_request_context(
                    f"/api/acq/school-district/{d0.get('geoid')}",
                    query_string={"lat": centroid.get("lat"), "lon": centroid.get("lon"),
                                  "name": d0.get("name") or ""}):
                session.update(outer_session)
                sd = acq_api_school_district_profile(str(d0.get("geoid")))
                sd = sd[0] if isinstance(sd, tuple) else sd
                sd = sd.get_json() if hasattr(sd, "get_json") else sd
            if isinstance(sd, dict) and not sd.get("error"):
                got.append(sd)
            else:
                print(f"[report] school profile refused for "
                      f"{d0.get('name')}: {str(sd)[:120] if sd else 'empty'}",
                      flush=True)
        if got:
            data["schools"] = got
        elif not districts:
            print(f"[report] no school district resolved "
                  f"(centroid={bool(centroid)})", flush=True)
    except Exception as e:
        print(f"[report] schools lookup failed: {e}", flush=True)

    # Competitor map, centred on the site with the CBAS ring.
    try:
        cb = data.get("cbas") or {}
        comms = cb.get("communities") or []
        centre = centroid or cb.get("center") or (data.get("amenities") or {}).get("centroid")
        if comms and centre:
            data["comp_map"] = acq_report.render_competitor_map(
                centre, comms, cb.get("radius_mi"))
    except Exception as e:
        print(f"[report] competitor map failed: {e}", flush=True)

    # Regional roadway map. The table names the projects; the map is what
    # shows whether the investment surrounds this site or sits across the
    # county.
    try:
        rd = data.get("roads") or {}
        projs = (rd.get("planned") or []) + (rd.get("programmed") or [])
        if projs and rd.get("center"):
            data["roads_map"] = acq_report.render_roads_map(
                rd["center"], projs, radius_mi=rd.get("radius_mi"))
    except Exception as e:
        print(f"[report] roads map failed: {e}", flush=True)

    ctx = acq_report.build_context(proj, analysis, data,
                                   (analysis.get("elevation") or None))

    # Subject map, drawn from the project's own geometry.
    try:
        geoms = [shp_shape(t["geometry"]) for t in (proj.get("tracts") or [])
                 if t.get("geometry")]
        if geoms:
            ctx["site_map"] = acq_report.render_site_map(
                unary_union(geoms), analysis.get("constraint_geoms"),
                proj.get("tracts"))
    except Exception as e:
        print(f"[report] site map failed: {e}", flush=True)

    ctx["_failed"] = data.get("_failed") or []
    return ctx, None


@acq_bp.route("/acquisitions/project/<pid>/report")
@_login_required
def acq_project_report_html(pid):
    """HTML view of the report. Not linked from anywhere -- it exists so the
    layout can be inspected in a browser without a WeasyPrint round trip."""
    guard = _acq_guard()
    if guard:
        return guard
    ctx, err = _build_report_context(pid)
    if err:
        return jsonify({"error": err}), 400
    # Debug view only: say which sections were skipped and why. Chasing a
    # missing page through the Railway log is slower than reading it here, and
    # this banner must never reach the PDF.
    skipped = (ctx.get("missing") or []) + (ctx.get("_failed") or [])
    present = [k for k in ("market", "comps", "schools", "demographics", "access",
                           "amenities", "news", "macro") if ctx.get(k)]
    banner = (
        '<div style="font:12px/1.5 monospace;background:#13344E;color:#fff;'
        'padding:10px 14px">'
        '<b>DEBUG — not part of the PDF.</b><br>sections present: '
        + (", ".join(present) or "none")
        + "<br>skipped: " + ("<br>&nbsp;&nbsp;" + "<br>&nbsp;&nbsp;".join(skipped)
                             if skipped else "none")
        + "</div>")
    return banner + render_template("acq_report.html", r=ctx)


@acq_bp.route("/api/acq/projects/<pid>/executive-pdf")
@_login_required
def acq_project_executive_pdf(pid):
    """The executive acquisition summary as a PDF."""
    guard = _acq_guard()
    if guard:
        return guard
    # Build and render are wrapped separately so a failure says WHICH stage
    # broke and on which line. A bare str(e) -- "'builtin_function_or_method'
    # object is not iterable" -- names neither, and the section mappers touch
    # payloads that cannot be exercised without the live CBAS/FRED/Census keys.
    import traceback
    try:
        ctx, err = _build_report_context(pid)
    except Exception as e:
        tb = traceback.format_exc()
        print(f"[report] context build failed:{chr(10)}{tb}", flush=True)
        frame = _report_last_frame(tb)
        return jsonify({"error": f"Report data build failed: {e}",
                        "where": frame, "stage": "context"}), 500
    if err:
        return jsonify({"error": err}), 400

    try:
        html = render_template("acq_report.html", r=ctx)
    except Exception as e:
        tb = traceback.format_exc()
        print(f"[report] template render failed:{chr(10)}{tb}", flush=True)
        return jsonify({"error": f"Report template failed: {e}",
                        "where": _report_last_frame(tb), "stage": "template"}), 500
    try:
        from weasyprint import HTML as _WHTML
    except Exception as e:
        # Named plainly rather than falling back to a different-looking
        # document: a silent substitution is worse than a clear failure when
        # the file is about to be emailed to a committee.
        print(f"[report] WeasyPrint unavailable: {e}", flush=True)
        return jsonify({"error": "PDF engine unavailable on this server "
                                 "(WeasyPrint could not load). The HTML view at "
                                 f"/acquisitions/project/{pid}/report still works."}), 500
    try:
        pdf = _WHTML(string=html, base_url=request.host_url).write_pdf()
    except Exception as e:
        print(f"[report] render failed: {e}", flush=True)
        return jsonify({"error": f"Report render failed: {e}"}), 500

    name = _re.sub(r"[^A-Za-z0-9]+", "_",
                  (ctx.get("project") or {}).get("name") or "Project").strip("_")[:48]
    fname = f"EMBER_{name}_Acquisition_Summary_{_utcnow()[:10]}.pdf"
    _acq_log("export_executive_pdf", {"project_id": pid})
    return send_file(_io.BytesIO(pdf), mimetype="application/pdf",
                     as_attachment=True, download_name=fname)


@acq_bp.route("/api/acq/projects/<pid>/pdf")
@_login_required
def acq_api_projects_pdf(pid):
    """Generate a presentation-grade Acquisition Analysis Summary PDF.

    Multi-page report with:
      • Cover page
      • Executive summary + project map image
      • Constraint bar chart + table
      • Yield analysis
      • Elevation + topography
      • Market context (Census + FRED)
      • Tract roster
      • Sources / methodology

    All charts rendered via matplotlib (Agg backend), then embedded in
    reportlab. Uses cached analysis (must run /analyze first) and refetches
    elevation/market/FRED data live.
    """
    import matplotlib
    import io
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    from matplotlib.patches import Polygon as MplPoly
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                       TableStyle, PageBreak, Image as RLImage,
                                       KeepTogether, CondPageBreak)
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from shapely.geometry import shape as shp_shape, mapping as shp_mapping
    from shapely.ops import unary_union
    import requests as _rq

    guard = _acq_guard()
    if guard:
        return guard
    uid = _acq_owner()
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, uid, _acq_is_admin())
    finally:
        conn.close()
    if not proj:
        return jsonify({"error": "project not found"}), 404

    analysis = proj.get("analysis_cache")
    if not analysis:
        return jsonify({"error": "Run the analysis first (click 'Run Acquisition Analysis' on the project page)."}), 400

    # Ember brand
    EMBER_ORANGE = "#F25929"
    EMBER_BLUE   = "#13344E"
    GRAY_700     = "#6B7B8B"
    GRAY_300     = "#DDE3E8"
    GRAY_100     = "#F6F8FA"
    OK_GREEN     = "#2E7D32"
    WARN_RED     = "#C62828"

    def _save_fig_to_buf(fig):
        b = io.BytesIO()
        fig.savefig(b, format="png", dpi=150, bbox_inches="tight",
                    facecolor="white", edgecolor="none")
        plt.close(fig)
        b.seek(0)
        return b

    # ===== Chart 1: Project map (with satellite tiles + constraints overlay) =====
    def render_project_map(proj_union, constraint_geoms=None, tracts_list=None):
        from shapely.geometry import MultiPolygon, Polygon as SP, LineString as LS
        from shapely.geometry import MultiLineString as MLS
        import math
        fig, ax = plt.subplots(figsize=(7.5, 5.5), dpi=160)
        ax.set_facecolor("#F0F4F8")

        # Bounds + padding
        minx, miny, maxx, maxy = proj_union.bounds
        pad_x = (maxx - minx) * 0.20 or 0.01
        pad_y = (maxy - miny) * 0.20 or 0.01
        west, east   = minx - pad_x, maxx + pad_x
        south, north = miny - pad_y, maxy + pad_y

        # Try to fetch a satellite tile basemap via Esri World Imagery — a simple
        # XYZ tile composite. Skips silently on failure (no internet, throttling).
        try:
            from PIL import Image as _PILImage
            import math as _math
            def lonlat_to_tile(lon, lat, z):
                lat_r = _math.radians(lat)
                n = 2 ** z
                xt = int((lon + 180) / 360 * n)
                yt = int((1 - _math.log(_math.tan(lat_r) + 1/_math.cos(lat_r)) / _math.pi) / 2 * n)
                return xt, yt
            def tile_to_lonlat(x, y, z):
                n = 2 ** z
                lon = x / n * 360 - 180
                lat = _math.degrees(_math.atan(_math.sinh(_math.pi * (1 - 2 * y / n))))
                return lon, lat
            # Choose zoom that gives ~6-12 tiles across — readable but not too many fetches
            zoom = 12
            while zoom > 6:
                x0, y0 = lonlat_to_tile(west, north, zoom)
                x1, y1 = lonlat_to_tile(east, south, zoom)
                if (x1 - x0 + 1) * (y1 - y0 + 1) <= 30: break
                zoom -= 1
            x0, y0 = lonlat_to_tile(west, north, zoom)
            x1, y1 = lonlat_to_tile(east, south, zoom)
            cols = x1 - x0 + 1
            rows = y1 - y0 + 1
            composite = _PILImage.new("RGB", (cols * 256, rows * 256))
            for tx in range(x0, x1 + 1):
                for ty in range(y0, y1 + 1):
                    try:
                        url = f"https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{zoom}/{ty}/{tx}"
                        tr = requests.get(url, timeout=5)
                        if tr.status_code == 200:
                            tile_img = _PILImage.open(io.BytesIO(tr.content))
                            composite.paste(tile_img, ((tx - x0) * 256, (ty - y0) * 256))
                    except Exception: continue
            # Compute the geographic bounds of the composite
            lon_w, lat_n = tile_to_lonlat(x0, y0, zoom)
            lon_e, lat_s = tile_to_lonlat(x1 + 1, y1 + 1, zoom)
            ax.imshow(composite, extent=(lon_w, lon_e, lat_s, lat_n), aspect="auto", zorder=0)
        except Exception:
            pass

        # Constraint overlays (drawn UNDER tract boundary so the tract pops)
        if constraint_geoms:
            from shapely.geometry import shape as _shp
            cstyles = {
                "floodplain":   ("#3380ff", "#0040aa", 0.35, 1.0, "polygon"),
                "wetlands":     ("#33a06f", "#005c2e", 0.40, 1.0, "polygon"),
                "transmission": ("#9c27b0", "#9c27b0", 0,    2.5, "line"),
                "streams":      ("#1976d2", "#1976d2", 0,    1.5, "line"),
                "pipelines":    ("#F25929", "#F25929", 0,    2.0, "line"),
            }
            for key, fc in (constraint_geoms or {}).items():
                if not fc: continue
                fillc, edgec, alpha, lw, kind = cstyles.get(key, ("#888", "#444", 0.3, 1, "polygon"))
                feats = fc.get("features") if "features" in fc else [fc]
                for feat in feats:
                    g = feat.get("geometry") if isinstance(feat, dict) else None
                    if not g: continue
                    try:
                        s = _shp(g)
                        if s.geom_type == "Polygon":
                            xs, ys = s.exterior.xy
                            ax.fill(xs, ys, color=fillc, alpha=alpha, edgecolor=edgec, linewidth=0.8, zorder=2)
                        elif s.geom_type == "MultiPolygon":
                            for sp in s.geoms:
                                xs, ys = sp.exterior.xy
                                ax.fill(xs, ys, color=fillc, alpha=alpha, edgecolor=edgec, linewidth=0.8, zorder=2)
                        elif s.geom_type == "LineString":
                            xs, ys = s.xy
                            ax.plot(xs, ys, color=edgec, linewidth=lw, alpha=0.85, zorder=3)
                        elif s.geom_type == "MultiLineString":
                            for sl in s.geoms:
                                xs, ys = sl.xy
                                ax.plot(xs, ys, color=edgec, linewidth=lw, alpha=0.85, zorder=3)
                    except Exception: continue

        # Individual tract boundaries (so user can see assemblage)
        if tracts_list:
            from shapely.geometry import shape as _shp
            for t in tracts_list:
                if not t.get("geometry"): continue
                try:
                    s = _shp(t["geometry"])
                    if s.geom_type == "Polygon":
                        xs, ys = s.exterior.xy
                        ax.plot(xs, ys, color="#FFFF00", linewidth=1.0, alpha=0.7, zorder=4)
                    elif s.geom_type == "MultiPolygon":
                        for sp in s.geoms:
                            xs, ys = sp.exterior.xy
                            ax.plot(xs, ys, color="#FFFF00", linewidth=1.0, alpha=0.7, zorder=4)
                except Exception: continue

        # PROJECT UNION boundary (bold orange)
        def _add_poly_outline(p):
            xs, ys = p.exterior.xy
            ax.fill(xs, ys, color=EMBER_ORANGE, alpha=0.18, zorder=5)
            ax.plot(xs, ys, color=EMBER_ORANGE, linewidth=3.0, zorder=6)
        if isinstance(proj_union, MultiPolygon):
            for p in proj_union.geoms: _add_poly_outline(p)
        elif isinstance(proj_union, SP):
            _add_poly_outline(proj_union)

        ax.set_xlim(west, east)
        ax.set_ylim(south, north)
        ax.set_aspect("equal")

        # N arrow + scale-ish text
        ax.text(0.97, 0.97, "N\n▲", transform=ax.transAxes, ha="center", va="top",
                fontsize=12, color="#FFFFFF", fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.3", facecolor=EMBER_BLUE, edgecolor="white"))

        # No degree ticks. Raw decimal latitude and longitude down the side of
        # an aerial reads as a plotting artefact in a partner deck, and nobody
        # navigates by them -- a scale bar is what the page actually needs.
        ax.set_xticks([]); ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_edgecolor(GRAY_300)

        # Scale bar, sized from the map's own extent. A degree of longitude
        # shortens with latitude, so it is measured at the map's centre.
        try:
            import math
            x0, x1 = ax.get_xlim()
            y0, y1 = ax.get_ylim()
            mid_lat = (y0 + y1) / 2.0
            span_ft = abs(x1 - x0) * 364000.0 * math.cos(math.radians(mid_lat))
            target = span_ft * 0.25
            nice = min([264, 528, 1320, 2640, 5280, 10560, 26400],
                       key=lambda v: abs(v - target))
            frac = (nice / span_ft) if span_ft else 0.25
            bx = x0 + (x1 - x0) * 0.04
            by = y0 + (y1 - y0) * 0.045
            ax.plot([bx, bx + (x1 - x0) * frac], [by, by],
                    color="#FFFFFF", lw=3.2, solid_capstyle="butt", zorder=60)
            ax.plot([bx, bx + (x1 - x0) * frac], [by, by],
                    color=EMBER_BLUE, lw=1.6, solid_capstyle="butt", zorder=61)
            ax.text(bx + (x1 - x0) * frac / 2, by + (y1 - y0) * 0.018,
                    (f"{nice/5280:g} mi" if nice >= 5280 else f"{nice:,} ft"),
                    ha="center", va="bottom", fontsize=7, color="#FFFFFF",
                    fontweight="bold", zorder=62,
                    bbox=dict(boxstyle="round,pad=0.18", facecolor=EMBER_BLUE,
                              edgecolor="none", alpha=0.85))
        except Exception:
            pass                      # a missing scale bar must not lose the map
        ax.set_title(f"Project · {analysis['tract_count']} tract{'' if analysis['tract_count'] == 1 else 's'} · {analysis['gross_acres']:,.0f} ac gross  ·  Floodplain/Wetlands/Transmission/Streams/Pipelines overlaid",
                      fontsize=9, color=EMBER_BLUE, pad=10, fontweight="bold")
        ax.set_xlabel("")
        ax.set_ylabel("")
        return _save_fig_to_buf(fig)

    # ===== Chart 2: Constraints bar chart =====
    def render_constraints_chart(constraints, gross):
        fig, ax = plt.subplots(figsize=(7.0, 3.0), dpi=150)
        labels = ["Floodplain", "Wetlands", "Trans. ROW", "Streams\n(50ft)", "Pipelines\n(50ft)"]
        keys = ["floodplain", "wetlands", "transmission_row", "stream_buffers", "pipeline_easements"]
        acres = []
        for k in keys:
            c = constraints.get(k) or {}
            acres.append(0 if c.get("error") else c.get("acres", 0))
        pcts = [(a / gross * 100) if gross > 0 else 0 for a in acres]

        bars = ax.barh(labels, acres, color=EMBER_BLUE, edgecolor=EMBER_ORANGE, linewidth=1.0, height=0.6)
        for bar, a, p in zip(bars, acres, pcts):
            width = bar.get_width()
            ax.text(width + max(acres) * 0.01, bar.get_y() + bar.get_height()/2,
                    f"  {a:,.1f} ac ({p:.1f}%)", va="center", fontsize=9, color=EMBER_BLUE)
        ax.invert_yaxis()
        ax.set_xlabel("Acres affected", fontsize=9, color=GRAY_700)
        ax.tick_params(labelsize=9, colors=GRAY_700)
        ax.set_xlim(0, max(max(acres) * 1.40, 1))
        for spine in ["top", "right"]:
            ax.spines[spine].set_visible(False)
        for spine in ["bottom", "left"]:
            ax.spines[spine].set_edgecolor(GRAY_300)
        ax.set_title("Net-out constraints (overlapping geometries counted once in net developable)",
                      fontsize=10, color=EMBER_BLUE, pad=8, fontweight="bold")
        return _save_fig_to_buf(fig)

    # ===== Chart 3: Gross → Net developable donut =====
    def render_developable_donut(gross, net):
        fig, ax = plt.subplots(figsize=(3.0, 3.0), dpi=150)
        unavail = max(gross - net, 0)
        if gross > 0:
            sizes = [net, unavail]
            labels = [f"Developable\n{net:,.0f} ac", f"Constrained\n{unavail:,.0f} ac"]
            colors_pie = [OK_GREEN, GRAY_300]
            wedges, _ = ax.pie(sizes, colors=colors_pie, startangle=90, counterclock=False,
                                wedgeprops={"width": 0.35, "edgecolor": "white", "linewidth": 2})
            pct = (net / gross * 100)
            ax.text(0, 0.05, f"{pct:.0f}%", ha="center", va="center",
                    fontsize=22, color=EMBER_ORANGE, fontweight="bold")
            ax.text(0, -0.18, "developable", ha="center", va="center",
                    fontsize=8, color=GRAY_700)
            # legend at bottom
            ax.legend(wedges, labels, loc="upper center", bbox_to_anchor=(0.5, -0.05),
                       ncol=2, frameon=False, fontsize=8)
        ax.set_aspect("equal")
        return _save_fig_to_buf(fig)

    # ===== Chart 4: FRED trend chart (Houston HPI + Texas HPI + Case-Shiller) =====
    def render_hpi_trend(fred_data):
        if not fred_data: return None
        fig, ax = plt.subplots(figsize=(7.0, 3.0), dpi=150)
        plotted_any = False
        for indicator, color, label in [
            ("ATNHPIUS26420Q", EMBER_ORANGE, "Houston HPI"),
            ("TXSTHPI",         EMBER_BLUE,   "Texas HPI"),
            ("CSUSHPINSA",      GRAY_700,     "U.S. 20-city HPI"),
        ]:
            ind = next((x for x in fred_data.get("indicators", [])
                        if x.get("series_id") == indicator), None)
            if not ind or ind.get("data", {}).get("error"): continue
            sparkline = ind["data"].get("sparkline") or []
            if len(sparkline) < 2: continue
            xs = list(range(len(sparkline)))
            ys = [p["value"] for p in sparkline]
            ax.plot(xs, ys, color=color, linewidth=2.0, label=label, marker="o", markersize=2)
            plotted_any = True
        if not plotted_any: return None
        ax.set_title("Home Price Index — 24-period trend (FRED)",
                      fontsize=10, color=EMBER_BLUE, fontweight="bold", pad=8)
        ax.tick_params(labelsize=8, colors=GRAY_700)
        ax.legend(loc="upper left", fontsize=8, frameon=False)
        ax.grid(True, linestyle=":", color=GRAY_300, linewidth=0.5)
        for spine in ["top", "right"]:
            ax.spines[spine].set_visible(False)
        for spine in ["bottom", "left"]:
            ax.spines[spine].set_edgecolor(GRAY_300)
        ax.set_xlabel("Periods (most-recent first → newest right)", fontsize=8, color=GRAY_700)
        ax.set_ylabel("Index value", fontsize=8, color=GRAY_700)
        return _save_fig_to_buf(fig)

    # ===== Gather all data (analysis is cached; elevation/market/FRED fetched fresh) =====
    tract_geoms = []
    for t in proj.get("tracts") or []:
        g = t.get("geometry")
        if g:
            try: tract_geoms.append(shp_shape(g))
            except Exception as e: print(f"[{request.endpoint}] tract geometry skipped: {e}", flush=True)
    proj_union = unary_union(tract_geoms) if tract_geoms else None

    # These four cards are their own endpoints. The standalone app called them
    # by building synthetic WSGI environments and pushing request contexts,
    # with a dead no-op block above it - all of that is unnecessary here. We
    # are already inside a request with a valid session, so they can simply be
    # called, and each is wrapped so one failing card cannot lose the report.
    def _card(fn, label, *a):
        try:
            resp = fn(*a)
            if isinstance(resp, tuple):
                resp = resp[0]
            return resp.get_json() if hasattr(resp, "get_json") else None
        except Exception as e:
            print(f"[acq-pdf] {label} card skipped: {type(e).__name__}: {e}", flush=True)
            return {"error": f"{label} skipped: {e}"}

    elev_data = (_card(acq_api_elevation_profile, "elevation",
                       shp_mapping(proj_union))
                 if proj_union else None)
    market_data = _card(acq_api_projects_market, "market", pid)
    fred_data = _card(acq_api_projects_fred, "FRED", pid)

    submarket_data = _card(acq_api_projects_submarket, "submarket", pid)

    # ===== Build PDF =====
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                              leftMargin=0.6*inch, rightMargin=0.6*inch,
                              topMargin=0.6*inch, bottomMargin=0.6*inch,
                              title=f"{proj.get('name','Project')} — Acquisition Analysis Summary",
                              author="Ember Acquisitions")

    styles = getSampleStyleSheet()
    body = ParagraphStyle('body', parent=styles['Normal'], fontSize=10, leading=14)
    body_center = ParagraphStyle('body_c', parent=body, alignment=TA_CENTER)
    title_xl = ParagraphStyle('title_xl', parent=styles['Heading1'], fontSize=32,
                                 textColor=colors.HexColor(EMBER_ORANGE),
                                 alignment=TA_CENTER, spaceAfter=8, leading=38)
    cover_sub = ParagraphStyle('cover_sub', parent=styles['Heading2'], fontSize=14,
                                  textColor=colors.HexColor(EMBER_BLUE), alignment=TA_CENTER,
                                  spaceAfter=4)
    h1 = ParagraphStyle('h1', parent=styles['Heading1'], fontSize=18,
                          textColor=colors.HexColor(EMBER_ORANGE),
                          spaceBefore=4, spaceAfter=8)
    h2 = ParagraphStyle('h2', parent=styles['Heading2'], fontSize=13,
                          textColor=colors.HexColor(EMBER_BLUE),
                          spaceBefore=12, spaceAfter=6)
    label = ParagraphStyle('label', parent=styles['Normal'], fontSize=8,
                              textColor=colors.HexColor(GRAY_700), leading=11)
    callout = ParagraphStyle('callout', parent=body, fontSize=11,
                                textColor=colors.HexColor(EMBER_BLUE), leftIndent=10,
                                spaceBefore=6, spaceAfter=6,
                                borderColor=colors.HexColor(EMBER_ORANGE), borderWidth=0,
                                leftBorderColor=colors.HexColor(EMBER_ORANGE),
                                leftBorderWidth=3, leftBorderPadding=6)

    def section_break(s):
        return [PageBreak()]

    def kpi_box(label_str, value_str, color=EMBER_BLUE):
        return [
            [Paragraph(f'<font size=8 color="{GRAY_700}">{label_str}</font>', label)],
            [Paragraph(f'<font size=22 color="{color}"><b>{value_str}</b></font>', body)],
        ]

    story = []

    # ============ PAGE 1: COVER ============
    story.append(Spacer(1, 1.4*inch))
    story.append(Paragraph("EMBER ACQUISITIONS", cover_sub))
    story.append(Spacer(1, 0.05*inch))
    story.append(Paragraph(proj.get("name", "(unnamed project)"), title_xl))
    # The map is the single most useful thing on the page, so it is rendered
    # once here and used on both the cover and the executive summary. The
    # cover was previously a title over 40% empty sheet.
    _map_png, _map_err = None, None
    if proj_union:
        try:
            _mb = render_project_map(
                proj_union,
                constraint_geoms=analysis.get("constraint_geoms"),
                tracts_list=proj.get("tracts") or [],
            )
            _map_png = _mb.getvalue()
        except Exception as e:
            _map_err = e

    _counties = sorted({(t.get("county") or "").strip()
                        for t in (proj.get("tracts") or []) if t.get("county")})
    _where = (", ".join(_counties) + (" County" if len(_counties) == 1 else " Counties")
              ) if _counties else "Texas"

    story.append(Paragraph("Acquisition Analysis Summary", cover_sub))
    story.append(Paragraph(
        f'<para alignment="center"><font size=11 color="{GRAY_700}">'
        f'{_where} &nbsp;·&nbsp; {analysis["gross_acres"]:,.0f} gross acres &nbsp;·&nbsp; '
        f'{analysis["tract_count"]} tract{"" if analysis["tract_count"] == 1 else "s"}'
        f'</font></para>', body))
    if _map_png:
        story.append(Spacer(1, 0.22*inch))
        story.append(RLImage(io.BytesIO(_map_png), width=6.4*inch, height=3.3*inch))
    story.append(Spacer(1, 0.26*inch))

    # Headline KPIs (centered)
    #
    # Three stacked paragraphs per cell, not one with <br/>: a 22pt number set
    # on the body style's 9pt leading printed the label straight through the
    # figure, so the cover read "GROSS" struck across "1,065".
    _kpi_lbl = ParagraphStyle("kpiLbl", parent=body, fontSize=8, leading=11,
                              alignment=1, textColor=colors.HexColor(GRAY_700))
    _kpi_val = ParagraphStyle("kpiVal", parent=body, fontSize=22, leading=26,
                              alignment=1, fontName="Helvetica-Bold")
    _kpi_sub = ParagraphStyle("kpiSub", parent=body, fontSize=8, leading=11,
                              alignment=1, textColor=colors.HexColor(GRAY_700))

    def _kpi(label, value, sub, colour):
        return [Paragraph(label.upper(), _kpi_lbl),
                Paragraph(f'<font color="{colour}">{value}</font>', _kpi_val),
                Paragraph(sub, _kpi_sub)]

    _y = analysis.get("yield_estimates", {}) or {}
    _upa = (_y.get("assumptions", {}) or {}).get("units_per_acre", 3.5)
    _tc = analysis["tract_count"]
    kpi_data = [
        [
            _kpi("Gross", f'{analysis["gross_acres"]:,.0f}', "acres", EMBER_BLUE),
            _kpi("Developable", f'{analysis["net_developable_acres"]:,.0f}',
                 f'{analysis["net_developable_pct"]:.0f}% of gross', EMBER_ORANGE),
            _kpi("Est. lots", f'{_y.get("total_lots", 0):,}',
                 f'@ {_upa:.1f} units/ac', EMBER_BLUE),
            _kpi("Tracts", f'{_tc}',
                 "combined" if _tc != 1 else "single tract", EMBER_BLUE),
        ]
    ]
    kt = Table(kpi_data, colWidths=[1.6*inch]*4, rowHeights=[1.2*inch])
    kt.setStyle(TableStyle([
        ('BOX',         (0,0), (-1,-1), 1.0, colors.HexColor(GRAY_300)),
        ('INNERGRID',   (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
        ('BACKGROUND',  (0,0), (-1,-1), colors.HexColor(GRAY_100)),
        ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(kt)
    story.append(Spacer(1, 0.42*inch))
    story.append(Paragraph(f'<para alignment="center"><font size=10 color="{GRAY_700}">Prepared {datetime.now().strftime("%B %d, %Y")} &nbsp;·&nbsp; Ember Group Acquisitions</font><br/><font size=8 color="{GRAY_700}">Confidential — for internal and partner review. Figures are preliminary and derived from public data; verify before underwriting.</font></para>', body))
    story.append(PageBreak())

    # ============ PAGE 2: EXECUTIVE SUMMARY (Map + Donut) ============
    story.append(Paragraph("Executive Summary", h1))

    if _map_png:
        story.append(RLImage(io.BytesIO(_map_png), width=7.2*inch, height=5.3*inch))
    elif _map_err:
        story.append(Paragraph(f"<i>Map render failed: {_map_err}</i>", label))

    # Side-by-side: donut + key callouts
    try:
        donut_buf = render_developable_donut(analysis["gross_acres"], analysis["net_developable_acres"])
        donut_img = RLImage(donut_buf, width=2.6*inch, height=2.6*inch)
    except Exception:
        donut_img = Paragraph("", body)

    callouts_html = []
    # Build narrative callouts based on the data
    if analysis["net_developable_pct"] >= 75:
        callouts_html.append(f"<b>Clean site.</b> {analysis['net_developable_pct']:.0f}% of gross acres are unconstrained — minimal floodplain/wetland deductions.")
    elif analysis["net_developable_pct"] >= 50:
        callouts_html.append(f"<b>Moderate constraints.</b> {(100-analysis['net_developable_pct']):.0f}% of acreage is encumbered. Confirm constraint geometry before underwriting.")
    else:
        callouts_html.append(f"<b>Heavy constraints.</b> Only {analysis['net_developable_pct']:.0f}% of gross is developable. Validate density assumptions against zoning + drainage requirements.")

    c = analysis.get("constraints", {})
    if (c.get("floodplain") or {}).get("pct", 0) > 10:
        callouts_html.append(f"<b>{c['floodplain']['pct']:.0f}% in 100-yr floodplain.</b> Significant fill/grading or out-conveyance needed.")
    if (c.get("wetlands") or {}).get("pct", 0) > 5:
        callouts_html.append(f"<b>{c['wetlands']['pct']:.0f}% wetlands (NWI).</b> Confirm USACE jurisdictional status before disturbance.")
    if (c.get("transmission_row") or {}).get("line_count", 0) > 0:
        callouts_html.append(f"<b>{c['transmission_row']['line_count']} transmission line(s)</b> cross the project — 150ft ROW assumed.")
    if (c.get("pipeline_easements") or {}).get("pipeline_count", 0) > 0:
        callouts_html.append(f"<b>{c['pipeline_easements']['pipeline_count']} pipeline(s)</b> cross the project — verify easement widths against RRC records.")

    callout_para = "".join([f"• {x}<br/><br/>" for x in callouts_html])
    callout_cell = Paragraph(callout_para, callout)

    exec_table = Table([[donut_img, callout_cell]], colWidths=[2.8*inch, 4.4*inch])
    exec_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING',  (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(exec_table)
    story.append(PageBreak())

    # ============ PAGE 3: CONSTRAINT ANALYSIS ============
    story.append(Paragraph("Constraints & Net Developable", h1))
    story.append(Paragraph(
        f"<b>Gross:</b> {analysis['gross_acres']:,.1f} ac  ·  "
        f"<b>Net developable:</b> <font color=\"{EMBER_ORANGE}\"><b>{analysis['net_developable_acres']:,.1f} ac</b></font>  ·  "
        f"<b>Yield:</b> {analysis['net_developable_pct']:.1f}%",
        body))
    story.append(Spacer(1, 8))

    try:
        chart_buf = render_constraints_chart(analysis.get("constraints", {}), analysis["gross_acres"])
        story.append(RLImage(chart_buf, width=7.0*inch, height=3.0*inch))
    except Exception as e:
        story.append(Paragraph(f"<i>Chart failed: {e}</i>", label))

    constraint_rows = [["Constraint", "Acres", "% of gross", "Detail / assumption"]]
    constraint_labels = {
        "floodplain":         ("FEMA 100-yr floodplain",  "feature_count", "polygons (Zone A/AE/VE)"),
        "wetlands":           ("Wetlands (USFWS NWI)",    "feature_count", "polygons"),
        "transmission_row":   ("Transmission ROW",        "line_count",    "lines × 150ft ROW assumed"),
        "stream_buffers":     ("Stream buffers",          "stream_count",  "NHD streams × 50ft each side"),
        "pipeline_easements": ("Pipeline easements",      "pipeline_count","RRC pipelines × 50ft each side"),
    }
    for key, (label_str, count_key, count_unit) in constraint_labels.items():
        cc = analysis.get("constraints", {}).get(key) or {}
        if cc.get("error"):
            constraint_rows.append([label_str, "—", "—", f"err: {cc['error'][:50]}"])
        else:
            cnt = cc.get(count_key, 0)
            detail = f"{cnt} {count_unit}" if cnt else "none detected"
            constraint_rows.append([label_str, f"{cc.get('acres', 0):,.1f}",
                                     f"{cc.get('pct', 0):.1f}%", detail])
    constraint_rows.append([
        Paragraph(f"<b>Net developable (union of constraints subtracted)</b>", body),
        Paragraph(f"<b>{analysis['net_developable_acres']:,.1f}</b>", body),
        Paragraph(f"<b>{analysis['net_developable_pct']:.1f}%</b>", body),
        "Overlapping constraints counted once",
    ])
    t = Table(constraint_rows, colWidths=[1.9*inch, 1.0*inch, 1.0*inch, 3.0*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor(EMBER_BLUE)),
        ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 9),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING',    (0,0), (-1,-1), 5),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#FFF4ED")),
    ]))
    story.append(Spacer(1, 8))
    story.append(t)
    story.append(PageBreak())

    # ============ PAGE 4: YIELD ESTIMATES ============
    story.append(Paragraph("Yield (pro-forma)", h1))
    ye = analysis.get("yield_estimates", {})
    breakdown = ye.get("breakdown", [])
    weighted_density = ye.get("weighted_density", 0)
    total_lots = ye.get("total_lots", 0)

    story.append(Paragraph(
        f"Pro-forma yield with multi-product mix · weighted density <b>{weighted_density} u/ac</b> · "
        f"applied to <b>{analysis['net_developable_acres']:,.0f} net developable ac</b>.",
        body))
    story.append(Spacer(1, 12))

    yield_kpi = Table([
        [
            [Paragraph("TOTAL LOTS", _kpi_lbl),
             Paragraph(f'<font size=42 color="{EMBER_ORANGE}"><b>{total_lots:,}</b></font>',
                       ParagraphStyle("yieldBig", parent=_kpi_val, fontSize=42, leading=48)),
             Paragraph(f'across {len(breakdown)} product type{"s" if len(breakdown)!=1 else ""}'
                       f' · {weighted_density} u/ac weighted', _kpi_sub)],
        ]
    ], colWidths=[6.8*inch], rowHeights=[1.8*inch])
    yield_kpi.setStyle(TableStyle([
        ('BOX',         (0,0), (-1,-1), 1.0, colors.HexColor(GRAY_300)),
        ('INNERGRID',   (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
        ('BACKGROUND',  (0,0), (-1,-1), colors.HexColor(GRAY_100)),
        ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(yield_kpi)
    story.append(Spacer(1, 16))

    # Product mix breakdown
    story.append(Paragraph("Product mix", h2))
    mix_rows = [["Product", "Density (u/ac)", "Allocation", "Acres", "Lots"]]
    for b in breakdown:
        mix_rows.append([
            b.get("label", "—"),
            f"{b.get('units_per_acre', 0):.1f}",
            f"{b.get('allocation_pct', 0):.0f}%",
            f"{b.get('acres', 0):,.1f}",
            f"{b.get('lots', 0):,}",
        ])
    mix_rows.append([
        Paragraph("<b>Total</b>", body), "",
        Paragraph(f"<b>{ye.get('total_allocation_pct', 0):.0f}%</b>", body),
        Paragraph(f"<b>{analysis['net_developable_acres']:,.1f}</b>", body),
        Paragraph(f"<b>{total_lots:,}</b>", body),
    ])
    mt = Table(mix_rows, colWidths=[2.4*inch, 1.3*inch, 1.1*inch, 1.0*inch, 1.0*inch])
    mt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0),  colors.HexColor(EMBER_BLUE)),
        ('TEXTCOLOR',  (0,0), (-1,0),  colors.white),
        ('FONTNAME',   (0,0), (-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 10),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
        ('ALIGN',      (1,1), (-1,-1), 'RIGHT'),
        ('PADDING',    (0,0), (-1,-1), 6),
        ('BACKGROUND', (0,-1),(-1,-1), colors.HexColor("#FFF4ED")),
    ]))
    story.append(mt)
    story.append(Spacer(1, 16))

    # Sensitivity table — show what happens at different densities
    story.append(Paragraph("Density sensitivity", h2))
    sens_rows = [["Units per acre", "Estimated lots", "Implied lot size (sqft)"]]
    for try_upa in [2.5, 3.0, 3.5, 4.0, 5.0, 6.0]:
        try_lots = int(round(analysis["net_developable_acres"] * try_upa))
        try_sqft = 43560 / try_upa if try_upa > 0 else 0
        sens_rows.append([f"{try_upa:.1f}", f"{try_lots:,}", f"{try_sqft:,.0f}"])
    st = Table(sens_rows, colWidths=[2.2*inch, 2.2*inch, 2.4*inch])
    st.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor(EMBER_BLUE)),
        ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 10),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING',    (0,0), (-1,-1), 6),
        ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
    ]))
    story.append(st)
    story.append(CondPageBreak(3.2*inch))
    story.append(Spacer(1, 0.18*inch))

    # ============ PAGE 5: ELEVATION / TOPOGRAPHY ============
    if elev_data and not elev_data.get("error"):
        story.append(Paragraph("Topography & Elevation", h1))
        slope_mean = elev_data.get("slope_mean_pct", 0)
        slope_max  = elev_data.get("slope_max_pct", 0)
        rng_ft     = elev_data.get("range_ft", 0)
        drainage   = elev_data.get("drainage_dir", "—")

        if slope_mean < 0.5:
            slope_label = "very flat — engineered grading required for drainage"
            slope_color = WARN_RED
        elif slope_mean < 2.0:
            slope_label = "gentle — standard grading"
            slope_color = OK_GREEN
        elif slope_mean < 5.0:
            slope_label = "moderate — engineered grading"
            slope_color = "#F59E0B"
        else:
            slope_label = "significant relief — natural drainage; verify floodplain on low side"
            slope_color = WARN_RED

        # Same overlap the cover had: an 18pt figure inside a paragraph
        # whose leading is set for 9pt text, so the label printed through it.
        elev_kpi = Table([
            [
                _kpi("Elevation range", f"{rng_ft:.0f} ft",
                     f'{elev_data.get("min_ft",0):.0f} - {elev_data.get("max_ft",0):.0f} ft',
                     EMBER_BLUE),
                _kpi("Mean slope", f"{slope_mean:.2f}%",
                     f"max {slope_max:.2f}%", slope_color),
                _kpi("Drainage", drainage, "fall direction", EMBER_BLUE),
            ]
        ], colWidths=[2.3*inch]*3, rowHeights=[1.0*inch])
        elev_kpi.setStyle(TableStyle([
            ('BOX',         (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
            ('INNERGRID',   (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
            ('BACKGROUND',  (0,0), (-1,-1), colors.HexColor(GRAY_100)),
            ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
        ]))
        story.append(elev_kpi)
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<b>Site character:</b> {slope_label}", callout))

        # ----- Render an elevation heatmap + contours image -----
        try:
            import numpy as _np
            grid = elev_data.get("grid")
            bbox = elev_data.get("bbox")
            if grid and bbox:
                Z = _np.array([[(v if v is not None else _np.nan) for v in row] for row in grid])
                minx, miny, maxx, maxy = bbox
                figmap, axm = plt.subplots(figsize=(7.0, 4.5), dpi=150)
                im = axm.imshow(Z, extent=(minx, maxx, miny, maxy), origin='lower',
                                 cmap='terrain', aspect='auto', alpha=0.85)
                # Contour lines on top
                contours = elev_data.get("contours") or []
                for c in contours:
                    for line in c.get("lines") or []:
                        xs = [p[0] for p in line]
                        ys = [p[1] for p in line]
                        axm.plot(xs, ys, color='black', linewidth=0.5, alpha=0.45)
                # Project boundary outline
                if proj_union:
                    if proj_union.geom_type == 'MultiPolygon':
                        for p in proj_union.geoms:
                            xs, ys = p.exterior.xy
                            axm.plot(xs, ys, color=EMBER_ORANGE, linewidth=2.5)
                    else:
                        xs, ys = proj_union.exterior.xy
                        axm.plot(xs, ys, color=EMBER_ORANGE, linewidth=2.5)
                # High/low markers
                hl = elev_data.get("highest_latlon")
                ll = elev_data.get("lowest_latlon")
                if hl: axm.plot(hl[1], hl[0], 'r^', markersize=10, markeredgecolor='white', markeredgewidth=1)
                if ll: axm.plot(ll[1], ll[0], 'bv', markersize=10, markeredgecolor='white', markeredgewidth=1)
                # Degree axes on an aerial read as a plotting artefact; the
                # colour bar already carries the only scale that matters here.
                axm.set_xticks([]); axm.set_yticks([])
                axm.set_xlabel("", fontsize=8, color=GRAY_700)
                axm.set_ylabel("", fontsize=8, color=GRAY_700)
                axm.tick_params(labelsize=7, colors=GRAY_700)
                cbar = figmap.colorbar(im, ax=axm, shrink=0.85)
                cbar.set_label("Elevation (ft)", fontsize=8, color=GRAY_700)
                cbar.ax.tick_params(labelsize=7, colors=GRAY_700)
                axm.set_title(f"Elevation surface · contour interval {elev_data.get('contour_interval_ft', '?')} ft · ▲ high {elev_data.get('highest_ft',0):.0f}ft  ▼ low {elev_data.get('lowest_ft',0):.0f}ft",
                                fontsize=9, color=EMBER_BLUE, pad=8, fontweight="bold")
                story.append(Spacer(1, 8))
                story.append(RLImage(_save_fig_to_buf(figmap), width=7.0*inch, height=4.5*inch))
        except Exception as e:
            story.append(Paragraph(f"<i>Elevation map render failed: {e}</i>", label))

        # ----- N-S and E-W cross-section profiles -----
        try:
            ns = elev_data.get("ns_profile") or []
            ew = elev_data.get("ew_profile") or []
            if ns and ew:
                figcross, (ax_ns, ax_ew) = plt.subplots(1, 2, figsize=(7.0, 2.3), dpi=150)
                # `frac` is the position along the transect; distance_ft
                # was never in the payload, so reading it put every point at
                # zero and drew a single vertical line.
                _ns_span = float(elev_data.get("ns_span_ft") or 0)
                _ew_span = float(elev_data.get("ew_span_ft") or 0)
                ns_x = [float(p.get("frac", 0)) * (_ns_span or 1.0) for p in ns]
                ns_y = [p.get("elev_ft") for p in ns]
                ew_x = [float(p.get("frac", 0)) * (_ew_span or 1.0) for p in ew]
                ew_y = [p.get("elev_ft") for p in ew]
                _unit = "ft" if _ns_span else "position (0-1)"
                ax_ns.fill_between(ns_x, ns_y, color=EMBER_ORANGE, alpha=0.20)
                ax_ns.plot(ns_x, ns_y, color=EMBER_ORANGE, linewidth=1.5)
                ax_ns.set_title("North–South cross-section", fontsize=9, color=EMBER_BLUE, fontweight="bold")
                ax_ew.fill_between(ew_x, ew_y, color=EMBER_BLUE, alpha=0.20)
                ax_ew.plot(ew_x, ew_y, color=EMBER_BLUE, linewidth=1.5)
                ax_ew.set_title("East–West cross-section", fontsize=9, color=EMBER_BLUE, fontweight="bold")
                for ax in (ax_ns, ax_ew):
                    ax.set_xlabel(f"Distance ({_unit})", fontsize=7, color=GRAY_700)
                    ax.set_ylabel("Elevation (ft)", fontsize=7, color=GRAY_700)
                    # Elevation range is tens of feet on a site thousands wide;
                    # a zero-based y-axis flattens every profile into a line.
                    _yy = [v for v in (ns_y + ew_y) if v is not None]
                    if _yy and max(_yy) > min(_yy):
                        _pad = (max(_yy) - min(_yy)) * 0.15
                        ax.set_ylim(min(_yy) - _pad, max(_yy) + _pad)
                    ax.tick_params(labelsize=7, colors=GRAY_700)
                    ax.grid(True, linestyle=":", color=GRAY_300, linewidth=0.5)
                    for spine in ["top", "right"]: ax.spines[spine].set_visible(False)
                figcross.tight_layout()
                story.append(Spacer(1, 6))
                story.append(RLImage(_save_fig_to_buf(figcross), width=7.0*inch, height=2.3*inch))
        except Exception as e:
            pass

        # ----- Hypsometric histogram (acres by elevation band) -----
        try:
            hist = elev_data.get("histogram") or []
            if hist:
                fighist, axh = plt.subplots(figsize=(7.0, 2.2), dpi=150)
                centers = [h.get("center_ft", 0) for h in hist]
                widths = [(h.get("high_ft", 0) - h.get("low_ft", 0)) for h in hist]
                acres = [h.get("acres", 0) for h in hist]
                axh.bar(centers, acres, width=widths, color=EMBER_BLUE, edgecolor="white", linewidth=0.5)
                axh.set_title("Acres by elevation band (hypsometric)", fontsize=9, color=EMBER_BLUE, fontweight="bold")
                axh.set_xlabel("Elevation (ft)", fontsize=7, color=GRAY_700)
                axh.set_ylabel("Acres", fontsize=7, color=GRAY_700)
                axh.tick_params(labelsize=7, colors=GRAY_700)
                axh.grid(True, linestyle=":", color=GRAY_300, linewidth=0.5, axis='y')
                for spine in ["top", "right"]: axh.spines[spine].set_visible(False)
                fighist.tight_layout()
                story.append(Spacer(1, 6))
                story.append(RLImage(_save_fig_to_buf(fighist), width=7.0*inch, height=2.2*inch))
        except Exception:
            pass

        story.append(Spacer(1, 6))
        story.append(Paragraph(
            f"Sampled via USGS 3DEP at {elev_data.get('sample_count', '?')} grid points across the project. "
            f"Median elevation {elev_data.get('median_ft', 0):.0f} ft.",
            label))
        story.append(PageBreak())

    # ============ SUBMARKET PAGE (tract + school district) ============
    if submarket_data and not submarket_data.get("error"):
        story.append(Paragraph("Submarket Profile", h1))
        # School districts
        sds = submarket_data.get("school_districts") or []
        if sds:
            sd_rows = [["School district", "Type", "GEOID"]]
            for sd in sds:
                sd_rows.append([sd.get("name") or "—", sd.get("type") or "—",
                                  sd.get("geoid") or "—"])
            t_sd = Table(sd_rows, colWidths=[3.5*inch, 1.5*inch, 2.0*inch])
            t_sd.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor(EMBER_BLUE)),
                ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
                ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE',   (0,0), (-1,-1), 10),
                ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
                ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
                ('PADDING',    (0,0), (-1,-1), 6),
            ]))
            story.append(t_sd)
            story.append(Spacer(1, 10))

        # Tract-level aggregates
        sumd = submarket_data.get("tract_summary") or {}
        if sumd and sumd.get("tract_count"):
            story.append(Paragraph(
                f"Census tract aggregates · <b>{sumd.get('tract_count')}</b> census tract{'' if sumd.get('tract_count') == 1 else 's'} containing the project · "
                f"population-weighted averages where applicable", body))
            story.append(Spacer(1, 6))
            sm_rows = [["Indicator", "Submarket"]]
            def _fmt(v, kind="int", prefix=""):
                if v is None: return "—"
                if kind == "float":  return f"{prefix}{v:,.1f}"
                if kind == "pct":    return f"{v:.1f}%"
                if kind == "dollar": return f"${v:,.0f}"
                return f"{prefix}{v:,.0f}"
            sm_rows.append(["Population",                  _fmt(sumd.get("total_population") or sumd.get("population"))])
            sm_rows.append(["Median age",                  _fmt(sumd.get("median_age"), "float")])
            sm_rows.append(["Households",                  _fmt(sumd.get("households"))])
            sm_rows.append(["Avg household size",          _fmt(sumd.get("avg_household_size"), "float")])
            sm_rows.append(["Median household income",     _fmt(sumd.get("median_household_income"), "dollar")])
            sm_rows.append(["Per capita income",           _fmt(sumd.get("per_capita_income"), "dollar")])
            sm_rows.append(["Median home value",           _fmt(sumd.get("median_home_value"), "dollar")])
            sm_rows.append(["Median rent",                 _fmt(sumd.get("median_rent"), "dollar")])
            sm_rows.append(["Housing units",               _fmt(sumd.get("housing_units"))])
            sm_rows.append(["Owner-occupancy %",           _fmt(sumd.get("owner_occupancy_pct"), "pct")])
            sm_rows.append(["Vacancy %",                   _fmt(sumd.get("vacancy_pct"), "pct")])
            sm_rows.append(["% Bachelor's+ (25+)",         _fmt(sumd.get("pct_bachelors_or_higher"), "pct")])
            sm_rows.append(["Employed",                    _fmt(sumd.get("employed"))])
            sm_rows.append(["Median year built",           _fmt(sumd.get("median_year_built"))])
            t_sm = Table(sm_rows, colWidths=[3.8*inch, 3.2*inch])
            t_sm.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor(EMBER_BLUE)),
                ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
                ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE',   (0,0), (-1,-1), 10),
                ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
                ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN',      (1,1), (-1,-1), 'RIGHT'),
                ('PADDING',    (0,0), (-1,-1), 5),
            ]))
            story.append(t_sm)

        # Caveats
        caveats = submarket_data.get("caveats") or []
        if caveats:
            story.append(Spacer(1, 10))
            for c in caveats:
                story.append(Paragraph(f"<i>• {c}</i>", label))
        story.append(CondPageBreak(3.4*inch))
        story.append(Spacer(1, 0.2*inch))

    # ============ PAGE 6: MARKET CONTEXT ============
    if market_data and not market_data.get("error"):
        story.append(Paragraph("Market Context", h1))
        cur = market_data.get("current", {})
        gro = market_data.get("growth", {})
        story.append(Paragraph(
            f"<b>{market_data.get('county_name','?')} County</b> · Census ACS 5-year ({cur.get('year','—')})",
            body))
        story.append(Spacer(1, 8))

        demo_rows = [["Indicator", "Current", "5-yr change", "5-yr CAGR"]]
        def _row(label_str, key, prefix="", fmt_kind="int"):
            v = cur.get(key)
            tot = gro.get(f"{key}_total_pct")
            cagr = gro.get(f"{key}_cagr_pct")
            if v is None: v_str = "—"
            elif fmt_kind == "float": v_str = f"{prefix}{v:,.1f}"
            elif fmt_kind == "pct":   v_str = f"{v:.1f}%"
            else: v_str = f"{prefix}{v:,.0f}"
            tot_str = f"{tot:+.1f}%" if tot is not None else "—"
            cagr_str = f"{cagr:+.1f}%" if cagr is not None else "—"
            demo_rows.append([label_str, v_str, tot_str, cagr_str])
        # Population
        _row("Population", "population")
        _row("Median age", "median_age", "", "float")
        _row("Households", "households")
        _row("Avg household size", "avg_household_size", "", "float")
        # Income
        _row("Median household income", "median_household_income", "$")
        _row("Per capita income", "per_capita_income", "$")
        # Housing
        _row("Housing units", "housing_units")
        _row("Median home value", "median_home_value", "$")
        _row("Median rent", "median_rent", "$")
        _row("Owner-occupancy", "owner_occupancy_pct", "", "pct")
        _row("Vacancy rate", "vacancy_pct", "", "pct")
        # Labor / education
        _row("Employed", "employed")
        _row("% Bachelor's+ (25+)", "pct_bachelors_or_higher", "", "pct")
        t = Table(demo_rows, colWidths=[2.2*inch, 1.8*inch, 1.5*inch, 1.5*inch])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor(EMBER_BLUE)),
            ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
            ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',   (0,0), (-1,-1), 9),
            ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN',      (1,1), (-1,-1), 'RIGHT'),
            ('PADDING',    (0,0), (-1,-1), 5),
        ]))
        story.append(t)
        story.append(Spacer(1, 14))

    # FRED indicators + HPI trend chart
    if fred_data and not fred_data.get("error"):
        story.append(Paragraph("Economic Indicators (FRED)", h2))
        try:
            hpi_buf = render_hpi_trend(fred_data)
            if hpi_buf:
                story.append(RLImage(hpi_buf, width=7.0*inch, height=3.0*inch))
                story.append(Spacer(1, 8))
        except Exception:
            pass

        fred_rows = [["Indicator", "Current", "YoY", "5-yr total", "5-yr CAGR"]]
        for ind in fred_data.get("indicators", []):
            d = ind.get("data") or {}
            if d.get("error"):
                fred_rows.append([ind["label"], "err", "—", "—", "—"])
                continue
            cur_v = d.get("current")
            fmt = ind.get("format", "")
            if fmt == "percent":   cur_str = f"{cur_v:.2f}%" if cur_v is not None else "—"
            elif fmt == "index":    cur_str = f"{cur_v:.1f}"   if cur_v is not None else "—"
            elif fmt == "thousands":cur_str = f"{cur_v:,.1f}K" if cur_v is not None else "—"
            elif fmt == "count":    cur_str = f"{cur_v:,.0f}"  if cur_v is not None else "—"
            else:                   cur_str = f"{cur_v:.2f}"   if cur_v is not None else "—"
            yoy = d.get("yoy_pct")
            five_total = d.get("five_year_total_pct")
            five_cagr = d.get("five_year_cagr_pct")
            fred_rows.append([
                ind["label"],
                cur_str,
                f"{yoy:+.2f}%" if yoy is not None else "—",
                f"{five_total:+.1f}%" if five_total is not None else "—",
                f"{five_cagr:+.2f}%" if five_cagr is not None else "—",
            ])
        t = Table(fred_rows, colWidths=[2.5*inch, 1.2*inch, 1.0*inch, 1.1*inch, 1.2*inch])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor(EMBER_BLUE)),
            ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
            ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',   (0,0), (-1,-1), 9),
            ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN',      (1,1), (-1,-1), 'RIGHT'),
            ('PADDING',    (0,0), (-1,-1), 5),
        ]))
        story.append(t)
        story.append(CondPageBreak(3.4*inch))
        story.append(Spacer(1, 0.2*inch))

    # ============ PAGE 7: TRACT ROSTER ============
    story.append(Paragraph("Tracts in this Project", h1))
    tract_rows = [["#", "Owner", "Prop ID", "County", "Acres"]]
    for i, tt in enumerate(proj.get("tracts") or [], 1):
        tract_rows.append([
            str(i),
            (tt.get("owner_name") or "—")[:48],
            tt.get("prop_id") or "—",
            tt.get("county") or "—",
            f"{(tt.get('acres') or 0):,.1f}",
        ])
    # Append total row
    tract_rows.append([
        Paragraph(f"<b>Total ({len(proj.get('tracts') or [])} tracts)</b>", body),
        "", "", "",
        Paragraph(f"<b>{analysis['gross_acres']:,.1f}</b>", body),
    ])
    tt_table = Table(tract_rows, colWidths=[0.4*inch, 3.0*inch, 1.4*inch, 1.4*inch, 0.8*inch])
    tt_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor(EMBER_BLUE)),
        ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 9),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor(GRAY_300)),
        ('ALIGN',      (0,0), (0,-1),  'CENTER'),
        ('ALIGN',      (-1,1),(-1,-1), 'RIGHT'),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING',    (0,0), (-1,-1), 4),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#FFF4ED")),
    ]))
    story.append(tt_table)
    story.append(Spacer(1, 18))

    # ---- Per-tract detail -------------------------------------------------
    # A project stores only prop_id / county / owner_name / acres / geometry.
    # The appraisal-district attributes the tract drawer shows are fetched live
    # per parcel and never persisted, which is why they were absent here.
    # Resolve each tract the same way the tract sheet does: last-search cache
    # first (it carries the spatial enrichment too), then the parcel cache with
    # one batched HCAD/MCAD overlay across all of them.
    DETAIL_CAP = 25
    _ptracts = proj.get("tracts") or []
    _detail_for = _ptracts[:DETAIL_CAP]

    _resolved = {}
    try:
        _cache_feats = ((_last_search_cache.get("data") or {}).get("tracts") or {}).get("features") or []
        _by_pid = {str((f.get("properties") or {}).get("Prop_ID") or "").strip(): f
                   for f in _cache_feats}
        _need = []
        for _tt in _detail_for:
            _pid = str(_tt.get("prop_id") or "").strip()
            if _pid and _pid in _by_pid:
                _resolved[_pid] = _by_pid[_pid]
            elif _pid:
                _need.append(_pid)

        if _need:
            import acq_parcels as _pc
            _found = []
            for _pid in _need:
                try:
                    _f = _pc.find_parcel_by_pid(_pid)
                except Exception as e:
                    print(f"[project-pdf] parcel lookup failed for {_pid}: {e}", flush=True)
                    _f = None
                if _f:
                    _resolved[_pid] = _f
                    _found.append(_f)
            # One overlay call for all of them - per-tract calls would make a
            # 20-tract assemblage unbearably slow.
            if _found:
                _fc = {"type": "FeatureCollection", "features": _found}
                try: hcad_live_overlay(_fc)
                except Exception as e: print(f"[project-pdf] hcad overlay failed: {e}", flush=True)
                try: mcad_live_overlay(_fc)
                except Exception as e: print(f"[project-pdf] mcad overlay failed: {e}", flush=True)
    except Exception as e:
        print(f"[project-pdf] tract detail resolution failed: {e}", flush=True)

    if _detail_for:
        story.append(Paragraph("Tract Detail", h2))
        story.append(Paragraph(
            "Appraisal-district record for each tract, refreshed live at export. "
            "Fields shown as \u2014 are not carried by the source for that parcel.",
            label))
        story.append(Spacer(1, 8))

    def _kv2(lbl, val):
        return [Paragraph(f"<b>{lbl}</b>", body),
                Paragraph(str(val) if val not in (None, "", 0) else "\u2014", body)]

    for _i, _tt in enumerate(_detail_for, 1):
        _pid = str(_tt.get("prop_id") or "").strip()
        _feat = _resolved.get(_pid) or {}
        _pr = _feat.get("properties") or {}

        _mkt  = _pr.get("MARKET_VAL") or _pr.get("total_market_val")
        _appr = _pr.get("total_appraised_val")
        _taxv = _pr.get("tax_value")
        try:
            _src = _tract_source_label(_pr) if _pr else "project record only"
        except Exception:
            _src = "\u2014"

        _left = [
            _kv2("Owner",        _pr.get("OWNER_NAME") or _tt.get("owner_name") or ""),
            _kv2("Mailing",      _pr.get("MAIL_ADDR") or ""),
            _kv2("Site address", _pr.get("SITUS_ADDR") or ""),
            _kv2("Legal desc",   (_pr.get("LEGAL_DESC") or "")[:120]),
            _kv2("Acres",        f"{(_tt.get('acres') or 0):,.1f}"),
            _kv2("County",       _tt.get("county") or _pr.get("_county") or ""),
            _kv2("ETJ / city",   _pr.get("_city_etj") or ""),
            _kv2("School dist",  _pr.get("_school_dist") or ""),
        ]
        _right = [
            _kv2("Prop ID",       _pid),
            _kv2("Source",        _src),
            _kv2("Owner since",   _pr.get("_owner_since") or ""),
            # "Tax year", not "HCAD tax year": the label column is an inch wide
            # and the longer form wraps onto two lines, and it is MCAD anyway
            # for Montgomery parcels.
            _kv2("Tax year",      _pr.get("_hcad_tax_year") or ""),
            _kv2("Market value",  f"${_mkt:,.0f}" if _mkt else ""),
            _kv2("Appraised",     f"${_appr:,.0f}" if _appr else ""),
            _kv2("Taxable",       f"${_taxv:,.0f}" if _taxv else ""),
            _kv2("Flood %",       _pr.get("_flood_pct") or ""),
        ]

        def _sub(rows):
            _t = Table(rows, colWidths=[1.0*inch, 2.4*inch])
            _t.setStyle(TableStyle([
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("LEFTPADDING", (0,0), (-1,-1), 0),
                ("RIGHTPADDING", (0,0), (-1,-1), 4),
                ("TOPPADDING", (0,0), (-1,-1), 1),
                ("BOTTOMPADDING", (0,0), (-1,-1), 1),
            ]))
            return _t

        _hdr = Paragraph(
            f'<font color="{EMBER_ORANGE}"><b>{_i}.</b></font> '
            f'<b>{(_pr.get("OWNER_NAME") or _tt.get("owner_name") or "\u2014")[:60]}</b>'
            f' <font color="{GRAY_700}" size=8>&nbsp;&nbsp;{_pid or ""}</font>', body)

        _outer = Table([[_sub(_left), _sub(_right)]], colWidths=[3.5*inch, 3.5*inch])
        _outer.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ]))
        story.append(KeepTogether([_hdr, Spacer(1, 3), _outer, Spacer(1, 12)]))

    if len(_ptracts) > DETAIL_CAP:
        # Say so rather than letting the report look complete when it is not.
        story.append(Paragraph(
            f"Detail shown for the first {DETAIL_CAP} of {len(_ptracts)} tracts. "
            "The roster above covers all of them.", label))

    story.append(Spacer(1, 18))

    # ============ FINAL: SOURCES + METHODOLOGY ============
    story.append(Paragraph("Sources & Methodology", h2))
    story.append(Paragraph(
        "<b>Parcel data:</b> TxGIO StratMap (statewide), refreshed weekly. "
        "Owner names verified live via HCAD (Harris) / MCAD (Montgomery) for those counties.<br/>"
        "<b>Floodplain:</b> FEMA National Flood Hazard Layer, Zone A/AE/AH/AO/AR/A99/V/VE (100-yr).<br/>"
        "<b>Wetlands:</b> USFWS National Wetlands Inventory (NWI). All wetland classes included.<br/>"
        "<b>Transmission:</b> HIFLD Electric Power Transmission Lines. 150ft ROW assumed (75ft each side).<br/>"
        "<b>Streams:</b> USGS National Hydrography Dataset (NHD). 50ft riparian buffer each side.<br/>"
        "<b>Pipelines:</b> Texas Railroad Commission. 50ft easement each side.<br/>"
        "<b>Elevation:</b> USGS 3DEP, sampled at a 25×25-grid across the project polygon.<br/>"
        "<b>Demographics:</b> U.S. Census Bureau ACS 5-year estimates (county-level).<br/>"
        "<b>Economic indicators:</b> Federal Reserve Bank of St. Louis (FRED).",
        label))
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        "<i>Constraint acreages reflect the overlapping geometries detected within the project boundary. "
        "Net developable subtracts the UNION of constraints to avoid double-counting. Yield estimates assume "
        "100% of net developable area yields lots — actual yields depend on roads, drainage, amenities, open "
        "space, and entitlement. Treat all numbers as planning-grade until verified by a survey and "
        "civil-engineering yield study. Computed " + analysis.get("computed_at", "—") + ".</i>",
        label))
    story.append(Spacer(1, 8))
    story.append(Paragraph(
        f'<para alignment="center"><font color="{GRAY_700}" size=8>Prepared by Ember Acquisitions · {datetime.now().strftime("%B %Y")}</font></para>',
        body))

    doc.build(story)
    buf.seek(0)
    _acq_log("project_pdf", {"id": pid, "name": proj.get("name"),
                                 "had_elevation": bool(elev_data and not elev_data.get("error")),
                                 "had_market":    bool(market_data and not market_data.get("error")),
                                 "had_fred":      bool(fred_data and not fred_data.get("error"))})

    safe_name = "".join(c for c in (proj.get("name") or "project") if c.isalnum() or c in " -_")[:40].strip() or "project"
    return send_file(buf, mimetype="application/pdf",
                     as_attachment=True,
                     download_name=f"{safe_name} — Acquisitions Summary.pdf")


# Elevation. The project PDF calls the profile endpoint directly, and the map's
# draw-a-line tool calls the line endpoint.

@acq_bp.route("/api/acq/elevation-line", methods=["POST"])
@_login_required
def acq_api_elevation_line():
    """Sample USGS 3DEP elevation along a polyline. Returns ordered samples
    with cumulative distance, ready to plot as a profile chart.

    Request: {"line": [[lat, lon], [lat, lon], ...], "samples": 60 (optional)}
    """
    body = request.get_json(force=True) or {}
    line = body.get("line") or []
    if len(line) < 2:
        return jsonify({"error": "need at least 2 points"}), 400
    n_samples = max(10, min(int(body.get("samples", 60)), 150))

    # Walk the line, build interpolated sample points spaced evenly by cumulative distance
    import math
    def haversine_ft(lat1, lon1, lat2, lon2):
        R = 20902231.0   # earth radius in feet
        p1, p2 = math.radians(lat1), math.radians(lat2)
        dp = math.radians(lat2 - lat1); dl = math.radians(lon2 - lon1)
        a = math.sin(dp/2)**2 + math.cos(p1) * math.cos(p2) * math.sin(dl/2)**2
        return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))

    # Total length + per-segment cumulative distance
    seg_dists = [0.0]
    for i in range(1, len(line)):
        lat1, lon1 = line[i-1]
        lat2, lon2 = line[i]
        seg_dists.append(seg_dists[-1] + haversine_ft(lat1, lon1, lat2, lon2))
    total_ft = seg_dists[-1]
    if total_ft < 1:
        return jsonify({"error": "line too short"}), 400

    # Generate sample points evenly along the line
    samples = []   # list of (distance_ft, lat, lon)
    for i in range(n_samples):
        d = total_ft * i / (n_samples - 1)
        # Find which segment d falls in
        for j in range(1, len(seg_dists)):
            if seg_dists[j] >= d:
                seg_start = seg_dists[j-1]
                seg_len = seg_dists[j] - seg_start
                t = (d - seg_start) / seg_len if seg_len > 0 else 0
                lat1, lon1 = line[j-1]; lat2, lon2 = line[j]
                lat = lat1 + (lat2 - lat1) * t
                lon = lon1 + (lon2 - lon1) * t
                samples.append((d, lat, lon))
                break

    # Query USGS 3DEP in parallel
    def fetch_one(s):
        d, lat, lon = s
        try:
            r = requests.get(ENDPOINTS["usgs_elevation"], params={
                "x": lon, "y": lat, "units": "Feet", "wkid": 4326,
            }, timeout=8)
            r.raise_for_status()
            v = r.json().get("value")
            return (d, lat, lon, float(v) if v is not None else None)
        except Exception:
            return (d, lat, lon, None)

    with ThreadPoolExecutor(max_workers=12) as pool:
        results = list(pool.map(fetch_one, samples))

    valid = [(d, e) for d, lat, lon, e in results if e is not None]
    if not valid:
        return jsonify({"error": "no elevation data (USGS 3DEP may be down)"}), 502
    elevs = [e for _, e in valid]
    return jsonify({
        "samples": [{"distance_ft": round(d, 1), "lat": lat, "lon": lon,
                      "elev_ft": round(e, 1) if e is not None else None}
                     for d, lat, lon, e in results],
        "total_length_ft": round(total_ft, 1),
        "total_length_mi": round(total_ft / 5280, 3),
        "min_ft": round(min(elevs), 1),
        "max_ft": round(max(elevs), 1),
        "range_ft": round(max(elevs) - min(elevs), 1),
        "drop_ft": round(elevs[0] - elevs[-1], 1),   # positive = downhill end-to-end
    })


@acq_bp.route("/api/acq/elevation-profile", methods=["POST"])
@_login_required
def acq_api_elevation_profile(geometry=None):
    """Sample USGS 3DEP elevation at a DENSE grid inside the supplied polygon,
    then return a rich analysis package:

      • summary stats: min / max / range / median / mean / slope%
      • drainage direction (compass) + magnitude
      • highest- and lowest-point locations on the tract
      • CONTOUR LINES at auto-chosen interval (2/5/10/20 ft), clipped to polygon
      • N-S and E-W cross-section profiles through tract centroid
      • hypsometric histogram (acres per elevation band)

    Grid density auto-scales with tract size: 15x15 for tiny tracts (<10 ac),
    25x25 default, 30x30 for large tracts (>1000 ac).

    Request: {"geometry": <GeoJSON Polygon or MultiPolygon>}

    `geometry` may also be passed directly, which is how the project PDF calls
    this: it is already inside a GET request that carries no body, so reading
    the body would 400.
    """
    if geometry is not None:
        body = {"geometry": geometry}
    else:
        body = request.get_json(force=True) or {}
    geom_json = body.get("geometry")
    if not geom_json:
        return jsonify({"error": "geometry required"}), 400

    from shapely.geometry import shape as shp_shape, Point, LineString
    import numpy as np
    try:
        geom = shp_shape(geom_json)
    except Exception as e:
        return jsonify({"error": f"bad geometry: {e}"}), 400
    if geom.is_empty:
        return jsonify({"error": "empty geometry"}), 400

    minx, miny, maxx, maxy = geom.bounds

    # ---- Grid density auto-scales with tract size ----
    # Cheap acreage estimate from shapely's planar area (rough, but fine for sizing)
    lat_mid = (miny + maxy) / 2
    lat_to_ft = 364320.0
    lon_to_ft = lat_to_ft * np.cos(np.radians(lat_mid))
    approx_acres = geom.area * lat_to_ft * lon_to_ft / 43560.0
    if approx_acres < 10:    GRID = 15
    elif approx_acres > 1000: GRID = 30
    else:                    GRID = 25

    # ---- Build FULL bbox grid (we'll keep "inside polygon" as a separate mask).
    # Sampling outside the bbox edges helps contour lines come out cleanly along
    # the polygon boundary — we clip them afterward via shapely.
    grid_points = []
    for i in range(GRID):
        for j in range(GRID):
            lat = miny + (maxy - miny) * (i + 0.5) / GRID
            lon = minx + (maxx - minx) * (j + 0.5) / GRID
            grid_points.append((i, j, lat, lon))

    # ---- Bulk sample USGS 3DEP via getSamples — but SPLIT the grid into
    # parallel batches. The endpoint scales linearly with point count, so 4
    # parallel batches of ~150 points each finish in ~1/4 the time of a single
    # 625-point call. Falls back to per-point endpoint if all batches fail.
    BULK_URL = "https://elevation.nationalmap.gov/arcgis/rest/services/3DEPElevation/ImageServer/getSamples"
    results = [None] * len(grid_points)
    M_TO_FT = 3.28084
    BATCH_SIZE = 160
    batches = [(start, grid_points[start:start + BATCH_SIZE])
               for start in range(0, len(grid_points), BATCH_SIZE)]

    def _fetch_batch(batch_start, batch_pts):
        try:
            geom_param = json.dumps({
                "points": [[gp[3], gp[2]] for gp in batch_pts],   # [lon, lat]
                "spatialReference": {"wkid": 4326},
            })
            r = requests.post(BULK_URL, data={
                "geometry": geom_param,
                "geometryType": "esriGeometryMultipoint",
                "returnFirstValueOnly": "true",
                "interpolation": "RSP_BilinearInterpolation",
                "f": "json",
            }, timeout=30)
            r.raise_for_status()
            j_data = r.json()
            if "error" in j_data:
                return False
            samples = j_data.get("samples") or []
            for k, s in enumerate(samples):
                v = s.get("value")
                try:
                    fv = float(v) * M_TO_FT if v not in (None, "", "NoData") else None
                except (ValueError, TypeError):
                    fv = None
                i, j, lat, lon = batch_pts[k]
                results[batch_start + k] = (i, j, lat, lon, fv)
            return True
        except Exception as e:
            print(f"[elevation] bulk batch failed at {batch_start}: {e}", flush=True)
            return False

    with ThreadPoolExecutor(max_workers=len(batches)) as pool:
        ok_flags = list(pool.map(lambda b: _fetch_batch(b[0], b[1]), batches))

    if not all(ok_flags):
        # Fallback: per-point single endpoint, parallelized
        session = requests.Session()
        def fetch_one(p):
            i, j, lat, lon = p
            try:
                r = session.get(ENDPOINTS["usgs_elevation"], params={
                    "x": lon, "y": lat, "units": "Feet", "wkid": 4326,
                }, timeout=10)
                r.raise_for_status()
                v = r.json().get("value")
                return (i, j, lat, lon, float(v) if v is not None else None)
            except Exception:
                return (i, j, lat, lon, None)
        with ThreadPoolExecutor(max_workers=30) as pool:
            results = list(pool.map(fetch_one, grid_points))

    # Fill in any None slots that the bulk call didn't return (rare edge case
    # — e.g. service returns fewer samples than requested) so downstream code
    # can safely unpack every entry.
    for k in range(len(results)):
        if results[k] is None:
            i, j, lat, lon = grid_points[k]
            results[k] = (i, j, lat, lon, None)

    # ---- Build numpy arrays ----
    Z   = np.full((GRID, GRID), np.nan, dtype=float)
    LAT = np.zeros((GRID, GRID), dtype=float)
    LON = np.zeros((GRID, GRID), dtype=float)
    inside_mask = np.zeros((GRID, GRID), dtype=bool)
    for i, j, lat, lon, e in results:
        LAT[i, j] = lat
        LON[i, j] = lon
        if e is not None: Z[i, j] = e
        if geom.contains(Point(lon, lat)): inside_mask[i, j] = True

    valid_inside_mask = inside_mask & ~np.isnan(Z)
    inside_z = Z[valid_inside_mask]
    if inside_z.size == 0:
        return jsonify({"error": "no elevation data returned (USGS 3DEP may be down or polygon too small)"}), 502

    min_e   = float(np.min(inside_z))
    max_e   = float(np.max(inside_z))
    median  = float(np.median(inside_z))
    mean    = float(np.mean(inside_z))
    range_e = max_e - min_e

    # ---- Slope (in %) via numpy gradient on the filled grid ----
    cell_dx_ft = (maxx - minx) / GRID * lon_to_ft
    cell_dy_ft = (maxy - miny) / GRID * lat_to_ft
    Z_filled = np.where(np.isnan(Z), mean, Z)
    dy_arr, dx_arr = np.gradient(Z_filled, cell_dy_ft, cell_dx_ft)
    slope_mag = np.sqrt(dx_arr**2 + dy_arr**2)   # ft/ft (rise/run)
    slope_mean_pct = float(np.mean(slope_mag[valid_inside_mask])) * 100
    slope_max_pct  = float(np.max(slope_mag[valid_inside_mask])) * 100

    # ---- Drainage direction = compass bearing of mean -gradient (downhill) ----
    flow_x_mean = -float(np.mean(dx_arr[valid_inside_mask]))   # E component
    flow_y_mean = -float(np.mean(dy_arr[valid_inside_mask]))   # N component
    flow_mag    = float(np.sqrt(flow_x_mean**2 + flow_y_mean**2))
    drainage_dir = "flat (no clear gradient)"
    if flow_mag > 0.001:   # >= ~0.1% slope
        angle = float(np.degrees(np.arctan2(flow_x_mean, flow_y_mean)))
        if angle < 0: angle += 360
        compass = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
                   "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]
        drainage_dir = compass[int((angle + 11.25) % 360 / 22.5)]

    # ---- Highest / lowest point locations ----
    # Mask out cells outside polygon by setting them to NaN for argmax/argmin
    inside_or_nan = np.where(valid_inside_mask, Z, np.nan)
    i_hi, j_hi = np.unravel_index(np.nanargmax(inside_or_nan), Z.shape)
    i_lo, j_lo = np.unravel_index(np.nanargmin(inside_or_nan), Z.shape)

    def _quadrant_label(i, j):
        """Human label for where in the tract this cell sits."""
        ns_thr = GRID * 0.30
        ew_thr = GRID * 0.30
        ns = "north" if (i - GRID/2) > ns_thr else ("south" if (GRID/2 - i) > ns_thr else "")
        ew = "east"  if (j - GRID/2) > ew_thr else ("west"  if (GRID/2 - j) > ew_thr else "")
        if ns and ew:    return f"{ns}{ew} corner"
        if ns:           return f"{ns} side"
        if ew:           return f"{ew} side"
        return "center"

    highest_loc = _quadrant_label(i_hi, j_hi)
    lowest_loc  = _quadrant_label(i_lo, j_lo)
    highest_latlon = [float(LAT[i_hi, j_hi]), float(LON[i_hi, j_hi])]
    lowest_latlon  = [float(LAT[i_lo, j_lo]), float(LON[i_lo, j_lo])]

    # ---- Choose contour interval based on total relief ----
    if   range_e < 8:   interval = 1
    elif range_e < 20:  interval = 2
    elif range_e < 80:  interval = 5
    elif range_e < 200: interval = 10
    else:               interval = 20

    # ---- Generate contour lines via matplotlib, then clip to polygon ----
    contours = []
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        levels = list(np.arange(np.floor(min_e/interval)*interval,
                                  np.ceil(max_e/interval + 1)*interval,
                                  interval))
        fig, ax = plt.subplots()
        cs = ax.contour(LON, LAT, Z, levels=levels)
        plt.close(fig)

        for level, allsegs in zip(cs.levels, cs.allsegs):
            for seg in allsegs:
                if len(seg) < 2: continue
                line = LineString([(float(x), float(y)) for x, y in seg])
                try:
                    clipped = line.intersection(geom)
                except Exception:
                    continue
                if clipped.is_empty: continue

                if clipped.geom_type == "LineString":
                    lines_to_add = [clipped]
                elif clipped.geom_type == "MultiLineString":
                    lines_to_add = list(clipped.geoms)
                else:
                    lines_to_add = [g for g in getattr(clipped, "geoms", [])
                                    if g.geom_type == "LineString"]

                for ln in lines_to_add:
                    if len(ln.coords) < 2: continue
                    contours.append({
                        "elev_ft": round(float(level), 1),
                        "coords": [[round(x, 6), round(y, 6)] for x, y in ln.coords],
                    })
    except Exception as e:
        print(f"[elevation] contour generation failed: {e}", flush=True)

    # ---- Cross-section profiles: middle row (E-W) and middle column (N-S) ----
    mid_row = GRID // 2
    mid_col = GRID // 2
    ew_profile = []
    ns_profile = []
    for j in range(GRID):
        if not np.isnan(Z[mid_row, j]):
            ew_profile.append({
                "frac": round(j / (GRID - 1), 3),
                "elev_ft": round(float(Z[mid_row, j]), 1),
                "inside": bool(inside_mask[mid_row, j]),
            })
    for i in range(GRID):
        if not np.isnan(Z[i, mid_col]):
            ns_profile.append({
                "frac": round(i / (GRID - 1), 3),
                "elev_ft": round(float(Z[i, mid_col]), 1),
                "inside": bool(inside_mask[i, mid_col]),
            })

    # ---- Hypsometric histogram: acres per elevation band ----
    poly_area_sqft = float(geom.area * lat_to_ft * lon_to_ft)
    poly_acres_calc = poly_area_sqft / 43560.0
    cells_inside = int(np.sum(valid_inside_mask))
    acres_per_cell = poly_acres_calc / cells_inside if cells_inside else 0.0

    bands = {}
    for z in inside_z:
        band_low = int(np.floor(z / interval) * interval)
        bands[band_low] = bands.get(band_low, 0.0) + acres_per_cell
    histogram = [
        {
            "low":   k,
            "high":  k + interval,
            "acres": round(v, 2),
            "pct":   round(v / poly_acres_calc * 100, 1) if poly_acres_calc else 0,
        }
        for k, v in sorted(bands.items())
    ]

    return jsonify({
        # core stats
        "min_ft":     round(min_e, 1),
        "max_ft":     round(max_e, 1),
        "median_ft":  round(median, 1),
        "mean_ft":    round(mean, 1),
        "range_ft":   round(range_e, 1),
        "sample_count": cells_inside,
        "polygon_acres": round(poly_acres_calc, 1),
        "grid_size": GRID,
        "bbox": [minx, miny, maxx, maxy],
        # slope & drainage
        "slope_mean_pct": round(slope_mean_pct, 2),
        "slope_max_pct":  round(slope_max_pct,  2),
        "drainage_dir":   drainage_dir,
        # high/low point
        "highest_ft":      round(max_e, 1),
        "lowest_ft":       round(min_e, 1),
        "highest_loc":     highest_loc,
        "lowest_loc":      lowest_loc,
        "highest_latlon":  highest_latlon,
        "lowest_latlon":   lowest_latlon,
        # rich payload
        "contour_interval_ft": interval,
        "contours":            contours,
        "ns_profile":          ns_profile,
        "ew_profile":          ew_profile,
        # Transect lengths, so a profile point's `frac` can be drawn against
        # real distance. Without these the PDF's cross-sections had nothing to
        # scale by and silently collapsed to a spike at zero.
        "ns_span_ft":          round(float((maxy - miny) * lat_to_ft), 1),
        "ew_span_ft":          round(float((maxx - minx) * lon_to_ft), 1),
        "histogram":           histogram,
        # legacy (back-compat — drop later)
        "grid": [[None if np.isnan(Z[i, j]) else round(float(Z[i, j]), 1)
                  for j in range(GRID)] for i in range(GRID)],
    })


# ══════════════════════════════════════════════════════════════════════════
# Saved objects: searches, folders, outreach, favourites, polygons
#
# Written against acq_store rather than ported. In the standalone app each of
# these loaded the entire JSON store, mutated a list and wrote the whole file
# back; against a document store they are a few lines each, and carrying that
# read-modify-write-the-world shape into Postgres would have been pointless.
# ══════════════════════════════════════════════════════════════════════════


def _crud_list(kind):
    conn = get_db()
    try:
        return acq_store.list_objects(conn, kind, _acq_owner(), _acq_is_admin())
    finally:
        conn.close()


def _crud_create(kind, body, required=()):
    for f in required:
        if not str(body.get(f) or "").strip():
            return None, (jsonify({"error": f"{f} required"}), 400)
    body = dict(body)
    body.setdefault("created_at", _utcnow())
    body.setdefault("created_by", session.get("username"))
    conn = get_db()
    try:
        return acq_store.put_object(conn, kind, body, _acq_owner()), None
    finally:
        conn.close()


def _crud_patch(kind, oid, body, fields):
    conn = get_db()
    try:
        obj = acq_store.get_object(conn, kind, oid, _acq_owner(), _acq_is_admin())
        if not obj:
            return None, (jsonify({"error": "not found"}), 404)
        for f in fields:
            if f in body:
                obj[f] = body[f]
        obj["updated_at"] = _utcnow()
        return acq_store.put_object(conn, kind, obj,
                                    obj.get("_owner_id") or _acq_owner()), None
    finally:
        conn.close()


def _crud_delete(kind, oid):
    conn = get_db()
    try:
        return acq_store.delete_object(conn, kind, oid, _acq_owner(), _acq_is_admin())
    finally:
        conn.close()


# ── Saved searches ────────────────────────────────────────────────────────

@acq_bp.route("/api/acq/searches", methods=["GET", "POST"])
@_login_required
def acq_searches():
    guard = _acq_guard()
    if guard:
        return guard
    if request.method == "GET":
        return jsonify({"searches": _crud_list("search")})
    obj, err = _crud_create("search", request.get_json(silent=True) or {}, ("name",))
    return err or (jsonify({"search": obj}), 201)


@acq_bp.route("/api/acq/searches/<sid>", methods=["PATCH", "DELETE"])
@_login_required
def acq_search_one(sid):
    guard = _acq_guard()
    if guard:
        return guard
    if request.method == "DELETE":
        ok = _crud_delete("search", sid)
        return jsonify({"deleted": ok}), (200 if ok else 404)
    obj, err = _crud_patch("search", sid, request.get_json(silent=True) or {},
                           ("name", "starred", "folder_id", "criteria", "notes"))
    return err or jsonify({"search": obj})


@acq_bp.route("/api/acq/searches/<sid>/folder", methods=["POST"])
@_login_required
def acq_search_folder(sid):
    guard = _acq_guard()
    if guard:
        return guard
    body = request.get_json(silent=True) or {}
    obj, err = _crud_patch("search", sid, {"folder_id": body.get("folder_id")},
                           ("folder_id",))
    return err or jsonify({"search": obj})


# ── Folders ───────────────────────────────────────────────────────────────

@acq_bp.route("/api/acq/folders/<fid>", methods=["PATCH", "DELETE"])
@_login_required
def acq_folder_one(fid):
    guard = _acq_guard()
    if guard:
        return guard
    if request.method == "DELETE":
        ok = _crud_delete("folder", fid)
        return jsonify({"deleted": ok}), (200 if ok else 404)
    obj, err = _crud_patch("folder", fid, request.get_json(silent=True) or {},
                           ("name", "color", "notes"))
    return err or jsonify({"folder": obj})


@acq_bp.route("/api/acq/folders/<fid>/share", methods=["POST"])
@_login_required
def acq_folder_share(fid):
    """Share a folder with other users by id.

    EmberApps has no per-object ACL, so this records the list on the folder and
    the read path honours it. Anything more would mean inventing a permissions
    model the rest of the portal does not have.
    """
    guard = _acq_guard()
    if guard:
        return guard
    body = request.get_json(silent=True) or {}
    ids = body.get("user_ids")
    if not isinstance(ids, list):
        return jsonify({"error": "user_ids must be a list"}), 400
    obj, err = _crud_patch("folder", fid, {"shared_with": ids}, ("shared_with",))
    return err or jsonify({"folder": obj})


# ── Outreach ──────────────────────────────────────────────────────────────

@acq_bp.route("/api/acq/outreach/<prop_id>", methods=["GET", "POST", "DELETE"])
@_login_required
def acq_outreach(prop_id):
    guard = _acq_guard()
    if guard:
        return guard
    conn = get_db()
    try:
        found = acq_store.find_by_prop(conn, "outreach", [prop_id],
                                       _acq_owner(), _acq_is_admin())
        rec = (found.get(str(prop_id)) or [None])[0]
        if request.method == "GET":
            return jsonify({"outreach": rec})
        if request.method == "DELETE":
            if not rec:
                return jsonify({"deleted": False}), 404
            ok = acq_store.delete_object(conn, "outreach", rec["id"],
                                         _acq_owner(), _acq_is_admin())
            return jsonify({"deleted": ok})
        body = request.get_json(silent=True) or {}
        rec = rec or {"prop_id": str(prop_id), "log": [],
                      "created_at": _utcnow()}
        for f in ("status", "broker_name", "broker_phone", "broker_email",
                  "next_action", "next_action_date", "asking_price", "notes",
                  "archived"):
            if f in body:
                rec[f] = body[f]
        rec["updated_at"] = _utcnow()
        saved = acq_store.put_object(conn, "outreach", rec,
                                     rec.get("_owner_id") or _acq_owner())
    finally:
        conn.close()
    return jsonify({"outreach": saved})


@acq_bp.route("/api/acq/outreach/<prop_id>/log", methods=["POST"])
@_login_required
def acq_outreach_log(prop_id):
    guard = _acq_guard()
    if guard:
        return guard
    body = request.get_json(silent=True) or {}
    conn = get_db()
    try:
        found = acq_store.find_by_prop(conn, "outreach", [prop_id],
                                       _acq_owner(), _acq_is_admin())
        rec = (found.get(str(prop_id)) or [None])[0] or {
            "prop_id": str(prop_id), "log": [],
            "created_at": _utcnow()}
        entry = {
            "id": acq_store.new_id("e"),
            "at": body.get("at") or _utcnow(),
            "method": body.get("method") or "other",
            "notes": body.get("notes") or "",
            "by": session.get("username"),
        }
        rec.setdefault("log", []).insert(0, entry)
        rec["updated_at"] = entry["at"]
        saved = acq_store.put_object(conn, "outreach", rec,
                                     rec.get("_owner_id") or _acq_owner())
    finally:
        conn.close()
    return jsonify({"outreach": saved, "entry": entry}), 201


@acq_bp.route("/api/acq/outreach/<prop_id>/log/<entry_id>",
              methods=["PATCH", "DELETE"])
@_login_required
def acq_outreach_log_entry(prop_id, entry_id):
    guard = _acq_guard()
    if guard:
        return guard
    conn = get_db()
    try:
        found = acq_store.find_by_prop(conn, "outreach", [prop_id],
                                       _acq_owner(), _acq_is_admin())
        rec = (found.get(str(prop_id)) or [None])[0]
        if not rec:
            return jsonify({"error": "not found"}), 404
        log = rec.get("log") or []
        if request.method == "DELETE":
            before = len(log)
            rec["log"] = [e for e in log if e.get("id") != entry_id]
            if len(rec["log"]) == before:
                return jsonify({"error": "entry not found"}), 404
        else:
            body = request.get_json(silent=True) or {}
            hit = next((e for e in log if e.get("id") == entry_id), None)
            if not hit:
                return jsonify({"error": "entry not found"}), 404
            for f in ("method", "notes", "at"):
                if f in body:
                    hit[f] = body[f]
        rec["updated_at"] = _utcnow()
        saved = acq_store.put_object(conn, "outreach", rec,
                                     rec.get("_owner_id") or _acq_owner())
    finally:
        conn.close()
    return jsonify({"outreach": saved})


@acq_bp.route("/api/acq/outreach/<prop_id>/archive", methods=["POST"])
@_login_required
def acq_outreach_archive(prop_id):
    guard = _acq_guard()
    if guard:
        return guard
    conn = get_db()
    try:
        found = acq_store.find_by_prop(conn, "outreach", [prop_id],
                                       _acq_owner(), _acq_is_admin())
        rec = (found.get(str(prop_id)) or [None])[0]
        if not rec:
            return jsonify({"error": "not found"}), 404
        rec["archived"] = bool((request.get_json(silent=True) or {}).get("archived", True))
        rec["updated_at"] = _utcnow()
        saved = acq_store.put_object(conn, "outreach", rec,
                                     rec.get("_owner_id") or _acq_owner())
    finally:
        conn.close()
    return jsonify({"outreach": saved})


@acq_bp.route("/api/acq/outreach/by-props", methods=["POST"])
@_login_required
def acq_outreach_by_props():
    guard = _acq_guard()
    if guard:
        return guard
    ids = (request.get_json(silent=True) or {}).get("prop_ids") or []
    conn = get_db()
    try:
        found = acq_store.find_by_prop(conn, "outreach", ids,
                                       _acq_owner(), _acq_is_admin())
    finally:
        conn.close()
    return jsonify({"outreach": {k: v[0] for k, v in found.items() if v}})


@acq_bp.route("/api/acq/outreach/pipeline")
@_login_required
def acq_outreach_pipeline():
    """Everything currently in the outreach pipeline, newest movement first.

    Terminal statuses are hidden unless asked for - the point of the list is
    what still needs doing.
    """
    guard = _acq_guard()
    if guard:
        return guard
    show_closed = request.args.get("include_closed", "0") == "1"
    rows = _crud_list("outreach")
    if not show_closed:
        rows = [r for r in rows
                if r.get("status") not in OUTREACH_OUT_OF_SCOPE and not r.get("archived")]
    rows.sort(key=lambda r: r.get("updated_at") or r.get("created_at") or "", reverse=True)

    # The pipeline board renders columns per status, so it reads by_status,
    # statuses and total — not a flat list. Returning {pipeline: [...]} left
    # every column empty with no error to show for it. statuses is expanded
    # from the (value, label, colour) tuples the board needs to draw headings.
    by_status = {}
    for r in rows:
        by_status.setdefault(r.get("status") or "lead", []).append(r)
    return jsonify({
        "by_status": by_status,
        "statuses": [{"value": v, "label": l, "color": c}
                     for v, l, c in OUTREACH_STATUSES],
        "total": len(rows),
        "methods": OUTREACH_METHODS,
        "pipeline": rows,          # kept: some callers still read the flat list
    })


# ── Favourites ────────────────────────────────────────────────────────────

@acq_bp.route("/api/acq/favorites", methods=["GET", "POST"])
@_login_required
def acq_favorites():
    guard = _acq_guard()
    if guard:
        return guard
    if request.method == "GET":
        return jsonify({"favorites": _crud_list("favorite")})
    obj, err = _crud_create("favorite", request.get_json(silent=True) or {}, ("prop_id",))
    return err or (jsonify({"favorite": obj}), 201)


@acq_bp.route("/api/acq/favorites/<fid>", methods=["DELETE"])
@_login_required
def acq_favorite_delete(fid):
    guard = _acq_guard()
    if guard:
        return guard
    ok = _crud_delete("favorite", fid)
    return jsonify({"deleted": ok}), (200 if ok else 404)


@acq_bp.route("/api/acq/favorites/by-props", methods=["POST"])
@_login_required
def acq_favorites_by_props():
    guard = _acq_guard()
    if guard:
        return guard
    ids = (request.get_json(silent=True) or {}).get("prop_ids") or []
    conn = get_db()
    try:
        found = acq_store.find_by_prop(conn, "favorite", ids,
                                       _acq_owner(), _acq_is_admin())
    finally:
        conn.close()
    return jsonify({"favorites": {k: v[0] for k, v in found.items() if v}})


# ── Drawn polygons ────────────────────────────────────────────────────────

@acq_bp.route("/api/acq/polygons", methods=["GET", "POST"])
@_login_required
def acq_polygons():
    guard = _acq_guard()
    if guard:
        return guard
    if request.method == "GET":
        return jsonify({"polygons": _crud_list("polygon")})
    obj, err = _crud_create("polygon", request.get_json(silent=True) or {},
                            ("name", "geometry"))
    return err or (jsonify({"polygon": obj}), 201)


@acq_bp.route("/api/acq/polygons/<poly_id>", methods=["PATCH", "DELETE"])
@_login_required
def acq_polygon_one(poly_id):
    guard = _acq_guard()
    if guard:
        return guard
    if request.method == "DELETE":
        ok = _crud_delete("polygon", poly_id)
        return jsonify({"deleted": ok}), (200 if ok else 404)
    obj, err = _crud_patch("polygon", poly_id, request.get_json(silent=True) or {},
                           ("name", "geometry", "notes", "color"))
    return err or jsonify({"polygon": obj})


# ── Tract pins and project assembly ───────────────────────────────────────

@acq_bp.route("/api/acq/tract-pins/<pin_id>/folder", methods=["POST"])
@_login_required
def acq_tract_pin_folder(pin_id):
    guard = _acq_guard()
    if guard:
        return guard
    body = request.get_json(silent=True) or {}
    obj, err = _crud_patch("tract_pin", pin_id, {"folder_id": body.get("folder_id")},
                           ("folder_id",))
    return err or jsonify({"pin": obj})


@acq_bp.route("/api/acq/projects/<pid>/add-tracts", methods=["POST"])
@_login_required
def acq_project_add_tracts(pid):
    """Add tracts to an existing project, skipping ones already on it."""
    guard = _acq_guard()
    if guard:
        return guard
    incoming = (request.get_json(silent=True) or {}).get("tracts") or []
    conn = get_db()
    try:
        proj = acq_store.get_object(conn, "project", pid, _acq_owner(), _acq_is_admin())
        if not proj:
            return jsonify({"error": "project not found"}), 404
        have = {str(t.get("prop_id") or "").strip() for t in proj.get("tracts") or []}
        added = 0
        for t in incoming:
            pid_str = str(t.get("prop_id") or "").strip()
            if not pid_str or pid_str in have:
                continue
            try:
                ac = float(t.get("acres") or 0)
            except (TypeError, ValueError):
                ac = 0.0
            proj.setdefault("tracts", []).append({
                "prop_id": pid_str,
                "county": (t.get("county") or "").strip(),
                "owner_name": (t.get("owner_name") or "").strip(),
                "acres": round(ac, 2),
                "geometry": t.get("geometry") or None,
            })
            have.add(pid_str)
            added += 1
        proj["total_acres"] = round(
            sum(float(x.get("acres") or 0) for x in proj.get("tracts") or []), 2)
        # The cached analysis described a different set of tracts; leaving it
        # would show yesterday's acreage against today's assemblage.
        if added:
            proj.pop("analysis_cache", None)
        saved = acq_store.put_object(conn, "project", proj,
                                     proj.get("_owner_id") or _acq_owner())
    finally:
        conn.close()
    return jsonify({"project": saved, "added": added})


@acq_bp.route("/api/acq/tracts/bulk-folder", methods=["POST"])
@_login_required
def acq_bulk_folder():
    """Pin a batch of tracts into one folder in a single call."""
    guard = _acq_guard()
    if guard:
        return guard
    body = request.get_json(silent=True) or {}
    folder_id = body.get("folder_id")
    tracts = body.get("tracts") or []
    conn = get_db()
    made = 0
    try:
        for t in tracts:
            pid_str = str(t.get("prop_id") or "").strip()
            if not pid_str:
                continue
            acq_store.put_object(conn, "tract_pin", {
                "prop_id": pid_str, "folder_id": folder_id,
                "owner_name": t.get("owner_name"), "acres": t.get("acres"),
                "county": t.get("county"), "geometry": t.get("geometry"),
                "created_at": _utcnow(),
            }, _acq_owner())
            made += 1
    finally:
        conn.close()
    return jsonify({"pinned": made})


@acq_bp.route("/api/acq/tracts/bulk-pipeline", methods=["POST"])
@_login_required
def acq_bulk_pipeline():
    """Open outreach records for a batch of tracts at one status."""
    guard = _acq_guard()
    if guard:
        return guard
    body = request.get_json(silent=True) or {}
    status = body.get("status") or "lead"
    valid = {v for v, _l, _c in OUTREACH_STATUSES}
    if status not in valid:
        return jsonify({"error": f"unknown status {status!r}",
                        "valid": sorted(valid)}), 400
    tracts = body.get("tracts") or []
    conn = get_db()
    made = skipped = 0
    try:
        existing = acq_store.find_by_prop(
            conn, "outreach",
            [str(t.get("prop_id") or "") for t in tracts],
            _acq_owner(), _acq_is_admin())
        for t in tracts:
            pid_str = str(t.get("prop_id") or "").strip()
            if not pid_str:
                continue
            if existing.get(pid_str):
                skipped += 1          # already in the pipeline; don't reset it
                continue
            acq_store.put_object(conn, "outreach", {
                "prop_id": pid_str, "status": status, "log": [],
                "owner_name": t.get("owner_name"), "acres": t.get("acres"),
                "county": t.get("county"),
                "created_at": _utcnow(),
            }, _acq_owner())
            made += 1
    finally:
        conn.close()
    return jsonify({"created": made, "already_in_pipeline": skipped})


# ══════════════════════════════════════════════════════════════════════════
# Corridors, imports and exports
#
# Ported from the standalone app. Corridor searches run over KMZ/KML boundary
# imports; the export routes build the KML, Excel and PDF deliverables.
# ══════════════════════════════════════════════════════════════════════════

@acq_bp.route("/api/acq/last-search")
def acq_api_last_search():
    """Returns the last search the server processed (its tracts + meta) so the
    main map page can restore results when the user comes back to it after
    navigating away. The map page calls this on load when no ?focus_* param
    is present. Single-user — fine for now since the app is per-user."""
    from flask import jsonify
    if not _last_search_cache.get("data"):
        return jsonify({"ok": False, "reason": "no_search"})
    return jsonify({
        "ok":         True,
        "layers":     _last_search_cache["data"],
        "meta":       _last_search_cache["meta"],
        "saved_at":   _last_search_cache.get("saved_at"),
    })


@acq_bp.route("/api/acq/tract-page-data/<prop_id>")
@_login_required
def acq_api_tract_page_data(prop_id):
    """Single endpoint that gathers everything the tract page needs in one
    parallel fetch: parcel geom from cache, HCAD live, outreach record, notes
    thread. Elevation loads separately (it's slow).

    Optional `?county=<name>` query param disambiguates Prop_ID collisions
    across counties (188287 is a 2,362-ac Fort Bend ranch AND a 0.1-ac
    Galveston home — caller passes county to get the right one)."""
    pid = (prop_id or "").strip()
    if not pid:
        return jsonify({"error": "prop_id required"}), 400
    county_hint = (request.args.get("county") or "").strip() or None

    import acq_parcels as parcel_cache
    uid = _current_user()["id"]

    # Parallel fetches
    def fetch_parcel():
        # strict=True when caller passed a county hint: never silently return a
        # different-county parcel (the root of the "wrong tract" bug — e.g.
        # pid=14966 is Grimes 7-D Investments AND Fort Bend Medina Bello).
        try: return parcel_cache.find_parcel_by_pid(pid, county=county_hint, strict=bool(county_hint))
        except Exception: return None

    def fetch_hcad():
        try:
            fc = arcgis_query(ENDPOINTS["hcad"], where=f"HCAD_NUM='{pid}'",
                              out_fields="*", page_size=1, max_pages=1,
                              return_geometry=False, parallel_pagination=False)
            return fc["features"][0].get("properties") if fc.get("features") else None
        except Exception:
            return None

    def fetch_mcad():
        if not pid.isdigit():
            return None
        try:
            fc = arcgis_query(ENDPOINTS["mcad"], where=f"PIN={int(pid)}",
                              out_fields="*", page_size=1, max_pages=1,
                              return_geometry=False, parallel_pagination=False)
            return fc["features"][0].get("properties") if fc.get("features") else None
        except Exception:
            return None

    with ThreadPoolExecutor(max_workers=3) as pool:
        f_parcel = pool.submit(fetch_parcel)
        f_hcad   = pool.submit(fetch_hcad)
        f_mcad   = pool.submit(fetch_mcad)
        parcel, hcad, mcad = f_parcel.result(), f_hcad.result(), f_mcad.result()

    # NOTE: don't 404 when there's no parcel/hcad/mcad. The drawer still needs
    # to render with outreach + notes from storage (e.g. a Grimes record where
    # we don't cache parcels). It will show "No parcel data available" in body
    # while the header still has owner/county from the outreach record.

    # Outreach + notes from storage (in-memory)
    with _storage_lock:
        s = _load_storage()
        user_lookup = {u["id"]: u for u in s["users"]}

    outreach = next((r for r in s.get("outreach", []) if r.get("prop_id") == pid), None)
    if outreach:
        out_obj = {k: v for k, v in outreach.items() if k != "log"}
        log = []
        for e in outreach.get("log", []) or []:
            u = user_lookup.get(e.get("by"))
            log.append({**e, "by_name": u.get("name") if u else "Unknown",
                        "can_edit": e.get("by") == uid, "can_delete": e.get("by") == uid})
        out_obj["log"] = log
    else:
        out_obj = None

    notes = []
    for n in s.get("notes", []):
        if n.get("prop_id") == pid and (n.get("text") or "").strip():
            u = user_lookup.get(n.get("owner_id"))
            notes.append({"id": n.get("id"), "text": n.get("text", ""),
                          "by": n.get("owner_id"),
                          "by_name": u.get("name") if u else "Unknown",
                          "at": n.get("created_at") or n.get("updated_at") or "",
                          "can_delete": n.get("owner_id") == uid})
    notes.sort(key=lambda e: e["at"] or "", reverse=True)

    # Sources marker
    sources = []
    if hcad: sources.append("HCAD live (Harris)")
    if mcad: sources.append("MCAD live (Montgomery)")
    if parcel and not (hcad or mcad): sources.append("StratMap (cache)")

    return jsonify({
        "prop_id": pid,
        "parcel": parcel,
        "hcad": hcad,
        "mcad": mcad,
        "sources": sources,
        "outreach": out_obj,
        "notes": notes,
        "outreach_statuses": [{"value": v, "label": l, "color": c} for v, l, c in OUTREACH_STATUSES],
        "outreach_methods":  OUTREACH_METHODS,
    })


@acq_bp.route("/api/acq/search-corridors", methods=["POST"])
@_login_required
def acq_api_search_corridors():
    """Run a single search across the UNION of multiple corridor saved searches.

    Body:
      • search_ids  — list of corridor search ids to combine (default: all of
                      the current user's KMZ corridor searches)
      • min_acres   — same as /api/search (default 300)
      • max_acres   — same as /api/search (default 100000)
      • group_by    — 'owner' or 'plat' (optional)
      • label       — display label for the cached run (optional)

    Returns the standard /api/search response (tracts + ancillary layers)
    plus a `corridor_meta` field describing which corridors were combined.
    """
    from shapely.geometry import shape as shp_shape, mapping as shp_mapping
    from shapely.ops import unary_union

    body = request.get_json(force=True) or {}
    uid = _current_user()["id"]
    requested_ids = body.get("search_ids") or []
    try:
        min_acres = float(body.get("min_acres", 300))
        max_acres = float(body.get("max_acres", 100000))
    except (TypeError, ValueError) as e:
        return jsonify({"error": f"bad acres: {e}"}), 400

    # `buffer_miles` (optional) — when provided, every selected corridor is
    # re-buffered AT RUNTIME from its stored centerline at this half-width
    # before unioning. This lets the user drive the corridor width with the
    # sidebar's "Radius" input — one knob to widen / narrow every corridor at
    # once. No persistence; the stored polygon stays at its original width.
    runtime_buffer_miles = None
    if "buffer_miles" in body and body["buffer_miles"] not in (None, ""):
        try:
            runtime_buffer_miles = float(body["buffer_miles"])
        except (TypeError, ValueError):
            return jsonify({"error": "buffer_miles must be a number"}), 400
        if not (0 < runtime_buffer_miles <= 50):
            return jsonify({"error": "buffer_miles must be between 0 and 50"}), 400

    with _storage_lock:
        s = _load_storage()

    all_corridors = [
        x for x in s.get("searches", [])
        if x.get("owner_id") == uid
        and ((x.get("polygon") or {}).get("properties") or {}).get("source") == "kmz_corridor"
    ]
    if requested_ids:
        wanted = set(requested_ids)
        targets = [x for x in all_corridors if x["id"] in wanted]
    else:
        targets = all_corridors
    if not targets:
        return jsonify({"error": "no corridors selected"}), 400

    geoms, used = [], []
    for t in targets:
        props = ((t.get("polygon") or {}).get("properties") or {})
        centerline = props.get("centerline_coords")
        effective_buffer = props.get("buffer_miles")
        fresh_geom = None

        # When a runtime buffer is requested AND we have the original centerline,
        # re-buffer at the requested width — same UTM/metric math as import.
        if runtime_buffer_miles is not None and centerline and len(centerline) >= 2:
            try:
                outer = _buffer_line_to_polygon(centerline,
                                                 runtime_buffer_miles * 1609.34)
                if outer:
                    fresh_geom = {"type": "Polygon", "coordinates": [outer]}
                    effective_buffer = runtime_buffer_miles
            except Exception as e:
                print(f"[corridor-search] runtime rebuffer failed for {t['id']}: {e}", flush=True)

        g = fresh_geom or (t.get("polygon") or {}).get("geometry")
        if not g:
            continue
        try:
            geoms.append(shp_shape(g))
            used.append({
                "id":           t["id"],
                "label":        t.get("label"),
                "filename":     props.get("filename"),
                "buffer_miles": effective_buffer,
                "rebuffered":   fresh_geom is not None,
            })
        except Exception as e:
            print(f"[corridor-search] bad geometry for {t['id']}: {e}", flush=True)
    if not geoms:
        return jsonify({"error": "no valid corridor geometries"}), 400

    union = unary_union(geoms)
    union_feature = {
        "type": "Feature",
        "geometry": shp_mapping(union),
        "properties": {
            "source": "corridor_union",
            "corridor_count": len(used),
            "corridor_ids":   [u["id"] for u in used],
        },
    }
    centroid = union.centroid

    if body.get("label"):
        label = body["label"].strip()
    elif runtime_buffer_miles is not None:
        label = f"All {len(used)} corridors @ ±{runtime_buffer_miles:g}mi"
    else:
        label = f"All {len(used)} corridors"
    group_by = (body.get("group_by") or "").strip().lower()
    if group_by not in ("owner", "plat"):
        group_by = None

    try:
        # tracts_only=True skips the 13 ancillary layer queries (flood, wetlands,
        # pipelines, etc.) and use_hcad_overlay=False skips the per-tract HCAD
        # live calls. Both are crushingly slow for a multi-corridor union polygon
        # that spans hundreds of miles. Layers can be loaded per-tract via the
        # focus-tract flow when the user clicks an individual tract. 5-10x speedup.
        result = run_search(float(centroid.y), float(centroid.x), 0.0,
                            min_acres, max_acres,
                            polygon_geojson=union_feature,
                            tracts_only=True,
                            use_hcad_overlay=False)
        annotate_tracts_with_enrichment(result["layers"])
        _last_search_cache["data"] = result["layers"]
        _last_search_cache["saved_at"] = time.time()
        _last_search_cache["meta"] = {
            "center": [float(centroid.y), float(centroid.x)], "radius_mi": 0,
            "min_acres": min_acres, "max_acres": max_acres,
            "polygon": union_feature,
            "label": label,
        }
        _acq_log("search_corridors_union",
                    {"corridor_count": len(used),
                     "min_acres": min_acres, "max_acres": max_acres,
                     "tracts": result["summary"]["tracts"]})
        result["corridor_meta"] = {
            "count":         len(used),
            "corridors":     used,
            "label":         label,
            "buffer_miles":  runtime_buffer_miles,
            "rebuffered":    runtime_buffer_miles is not None,
        }
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@acq_bp.route("/api/acq/corridors-list", methods=["GET"])
@_login_required
def acq_api_corridors_list():
    """Lightweight metadata for the Corridors page tab strip — no search runs.
    Returns one entry per saved corridor for the current user."""
    uid = _current_user()["id"]
    with _storage_lock:
        s = _load_storage()
    out = []
    for x in s.get("searches", []):
        if x.get("owner_id") != uid:
            continue
        props = ((x.get("polygon") or {}).get("properties") or {})
        if props.get("source") != "kmz_corridor":
            continue
        out.append({
            "id":               x["id"],
            "label":            x.get("label"),
            "filename":         props.get("filename"),
            "buffer_miles":     props.get("buffer_miles"),
            "centerline_pts":   len(props.get("centerline_coords") or []),
            "has_centerline":   bool(props.get("centerline_coords")),
            "saved_at":         x.get("saved_at"),
        })
    return jsonify({"corridors": out, "count": len(out)})


@acq_bp.route("/api/acq/corridor/<sid>/tracts", methods=["POST"])
@_login_required
def acq_api_corridor_tracts(sid):
    """Run a search inside ONE corridor and return enriched tracts.

    Each tract is annotated with `_outreach_status`, `_outreach_next_action`,
    `_outreach_next_action_date`, `_folder_names`, `_is_favorite`, `_has_notes`
    so the Corridors page can display indicators and offer actions without a
    second round-trip.

    Body: { buffer_miles?, min_acres?, max_acres?, force? }
    Pass `force=true` to bypass the per-user / per-args cache.
    """
    from shapely.geometry import shape as shp_shape, mapping as shp_mapping
    body = request.get_json(force=True) or {}
    uid = _current_user()["id"]

    try:
        min_acres = float(body.get("min_acres", 50))
        max_acres = float(body.get("max_acres", 3000))
    except (TypeError, ValueError):
        return jsonify({"error": "bad acres"}), 400

    runtime_buffer_miles = None
    if "buffer_miles" in body and body["buffer_miles"] not in (None, ""):
        try:
            runtime_buffer_miles = float(body["buffer_miles"])
        except (TypeError, ValueError):
            return jsonify({"error": "bad buffer_miles"}), 400
        if not (0 < runtime_buffer_miles <= 50):
            return jsonify({"error": "buffer_miles must be 0–50"}), 400

    force = bool(body.get("force"))
    cache_key = (uid, sid, runtime_buffer_miles, min_acres, max_acres)
    if not force:
        cached = _CORRIDOR_TRACTS_CACHE.get(cache_key)
        if cached and (time.time() - cached["t"]) < _CORRIDOR_CACHE_TTL:
            cached["data"]["from_cache"] = True
            cached["data"]["cached_age_seconds"] = int(time.time() - cached["t"])
            return jsonify(cached["data"])

    # Load the corridor
    with _storage_lock:
        s = _load_storage()
    corridor = next((x for x in s.get("searches", [])
                     if x["id"] == sid and x.get("owner_id") == uid), None)
    if not corridor:
        return jsonify({"error": "corridor not found"}), 404
    props = ((corridor.get("polygon") or {}).get("properties") or {})
    if props.get("source") != "kmz_corridor":
        return jsonify({"error": "not a corridor search"}), 400

    # Re-buffer if requested
    centerline = props.get("centerline_coords")
    polygon_feature = corridor["polygon"]
    effective_buffer = props.get("buffer_miles")
    if runtime_buffer_miles is not None and centerline and len(centerline) >= 2:
        outer = _buffer_line_to_polygon(centerline, runtime_buffer_miles * 1609.34)
        if outer:
            polygon_feature = {
                "type": "Feature",
                "geometry": {"type": "Polygon", "coordinates": [outer]},
                "properties": {**props, "buffer_miles": runtime_buffer_miles},
            }
            effective_buffer = runtime_buffer_miles

    # Run search
    geom = shp_shape(polygon_feature["geometry"])
    c = geom.centroid
    try:
        # For the corridors page we skip BOTH the ancillary layer queries
        # AND the HCAD/MCAD live overlays. The overlays make per-tract HTTP
        # calls (hundreds-to-thousands per corridor) and the corridors view
        # only needs owner+acres for triage — StratMap's annual data is fine.
        # Users who need fresh owner data can click "Show on map" or "Full
        # tract page" which do the live HCAD lookup on demand.
        result = run_search(float(c.y), float(c.x), 0.0,
                            min_acres, max_acres, polygon_geojson=polygon_feature,
                            tracts_only=True, use_hcad_overlay=False)
        annotate_tracts_with_enrichment(result["layers"])
    except Exception as e:
        return jsonify({"error": f"search failed: {e}"}), 500

    tracts_fc = result["layers"]["tracts"]
    tract_features = tracts_fc.get("features") or []

    # ---- Enrichment: outreach + folder + favorite + notes lookups ----
    user_lookup = {u["id"]: u for u in s.get("users", [])}
    outreach_by_pid = {}
    for o in s.get("outreach", []):
        pid = o.get("prop_id")
        if pid:
            outreach_by_pid[pid] = o

    notes_pids = set()
    for n in s.get("notes", []):
        if (n.get("text") or "").strip() and n.get("prop_id"):
            notes_pids.add(n["prop_id"])

    fav_pids = set()
    for f in s.get("favorites", []):
        if f.get("owner_id") == uid and f.get("prop_id"):
            fav_pids.add(f["prop_id"])

    # Folders this user owns OR is shared with — map prop_id -> [folder names]
    folder_lookup = {fld["id"]: fld for fld in s.get("folders", [])}
    pin_folders_by_pid = {}
    for pin in s.get("tract_pins", []):
        pid = pin.get("prop_id")
        if not pid: continue
        fid = pin.get("folder_id")
        fld = folder_lookup.get(fid) if fid else None
        if fld:
            pin_folders_by_pid.setdefault(pid, []).append(fld.get("name") or "(folder)")
        else:
            pin_folders_by_pid.setdefault(pid, []).append("(loose pin)")

    # Stamp each tract
    for f in tract_features:
        pid = str((f.get("properties") or {}).get("Prop_ID") or "")
        p = f["properties"]
        out = outreach_by_pid.get(pid)
        if out:
            p["_outreach_status"]            = out.get("status") or "lead"
            p["_outreach_next_action"]       = out.get("next_action")
            p["_outreach_next_action_date"]  = out.get("next_action_date")
            p["_outreach_broker"]            = out.get("broker_name")
            p["_outreach_log_count"]         = len(out.get("log") or [])
        else:
            p["_outreach_status"] = None
        p["_folder_names"] = pin_folders_by_pid.get(pid, [])
        p["_is_favorite"]  = pid in fav_pids
        p["_has_notes"]    = pid in notes_pids

    # ---- Summary KPIs ----
    total_acres = sum((f["properties"].get("Acres") or 0) for f in tract_features)
    counties = {}
    statuses = {}
    in_pipeline = 0
    closed = 0
    untouched = 0
    for f in tract_features:
        p = f["properties"]
        c2 = p.get("_county") or "Unknown"
        counties[c2] = counties.get(c2, 0) + 1
        st = p.get("_outreach_status")
        if st:
            in_pipeline += 1
            statuses[st] = statuses.get(st, 0) + 1
            if st == "closed":
                closed += 1
        if (not p.get("_outreach_status")) and (not p.get("_has_notes")) \
                and (not p.get("_is_favorite")) and not p.get("_folder_names"):
            untouched += 1
    top_counties = sorted(counties.items(), key=lambda kv: -kv[1])[:3]

    response = {
        "corridor": {
            "id":               corridor["id"],
            "label":            corridor.get("label"),
            "filename":         props.get("filename"),
            "buffer_miles":     effective_buffer,
            "rebuffered":       runtime_buffer_miles is not None,
            "polygon":          polygon_feature,
        },
        "tracts":  tract_features,
        "summary": {
            "tracts":       len(tract_features),
            "total_acres":  round(total_acres, 1),
            "avg_acres":    round(total_acres / len(tract_features), 1) if tract_features else 0,
            "in_pipeline":  in_pipeline,
            "closed":       closed,
            "untouched":    untouched,
            "counties":     [{"name": n, "count": c2} for n, c2 in top_counties],
            "statuses":     statuses,
            "min_acres":    min_acres,
            "max_acres":    max_acres,
            "truncated":    result.get("summary", {}).get("truncated", False),
        },
        "from_cache":         False,
        "cached_age_seconds": 0,
    }
    _CORRIDOR_TRACTS_CACHE[cache_key] = {"t": time.time(), "data": response}
    # Trim cache if it gets too big
    if len(_CORRIDOR_TRACTS_CACHE) > 100:
        oldest = sorted(_CORRIDOR_TRACTS_CACHE.items(), key=lambda kv: kv[1]["t"])[:20]
        for k, _ in oldest:
            _CORRIDOR_TRACTS_CACHE.pop(k, None)
    return jsonify(response)


@acq_bp.route("/api/acq/cache/reindex", methods=["POST"])
@_admin_required
def api_acq_cache_reindex():
    """Rebuild spatial-index entries for parcels that have none.

    A parcel with no R-Tree row is in the table, counts toward the county
    total, and can never be returned by a search — so the cache reads healthy
    while every result is quietly short. This repairs from the geometry already
    stored; nothing is re-downloaded.
    """
    guard = _acq_guard()
    if guard:
        return guard
    try:
        return jsonify(parcel_cache.reindex_missing_rtree())
    except Exception as e:
        return jsonify({"error": f"{type(e).__name__}: {e}"}), 500


@acq_bp.route("/api/acq/cache/vacuum", methods=["POST"])
@_admin_required
def api_acq_cache_vacuum():
    """Drop R-Tree entries with no surviving parcel.

    Re-bootstrapping a county leaves the old rows' index entries behind, and
    they accumulate: the standalone's cache reached 4.6M entries against 1.9M
    parcels, 2.4x oversized, and every spatial query walked the dead ones
    first. Results stay correct either way — the query joins parcels to the
    index — so this is a speed fix, not a correctness one.
    """
    guard = _acq_guard()
    if guard:
        return guard
    try:
        return jsonify(parcel_cache.vacuum_rtree_orphans())
    except Exception as e:
        return jsonify({"error": f"{type(e).__name__}: {e}"}), 500


@acq_bp.route("/api/acq/cache/bootstrap", methods=["POST"])
@_login_required
@_admin_required
def acq_api_cache_bootstrap():
    """Trigger a (re)bootstrap of one or more counties in the background.
    Body: {"counties": ["48201", ...] or "all"}.  Returns immediately; the
    actual work runs on a thread — poll /api/cache/status for progress.
    """
    import acq_parcels as parcel_cache
    body = request.get_json(force=True) or {}
    targets = body.get("counties")
    if targets == "all" or not targets:
        target_pairs = parcel_cache.HOUSTON_METRO_COUNTIES
    else:
        wanted = set(targets)
        target_pairs = [(f, n) for f, n in parcel_cache.HOUSTON_METRO_COUNTIES if f in wanted]
    if not target_pairs:
        return jsonify({"error": "No valid county FIPS in request"}), 400

    if _cache_bootstrap_status["running"]:
        return jsonify({"error": "A cache bootstrap is already running",
                        "in_progress": dict(_cache_bootstrap_status)}), 409

    def _run():
        with _cache_bootstrap_lock:
            _cache_bootstrap_status.update({"running": True, "pct": 0, "msg": "Starting…"})
            try:
                for fips, name in target_pairs:
                    _cache_bootstrap_status.update({"county": name, "pct": 0,
                                                     "msg": f"Bootstrapping {name}…"})
                    def _prog(pct, msg):
                        _cache_bootstrap_status.update({"pct": pct, "msg": msg})
                    try:
                        parcel_cache.bootstrap_county(fips, name, _prog)
                    except Exception as e:
                        _cache_bootstrap_status.update({"msg": f"{name}: ERROR {e}"})
                        print(f"[cache] bootstrap of {name} failed: {e}", flush=True)
                _cache_bootstrap_status.update({"running": False, "county": None,
                                                  "pct": 100, "msg": "All requested counties done."})
            except Exception as e:
                _cache_bootstrap_status.update({"running": False, "msg": f"FATAL: {e}"})

    threading.Thread(target=_run, name="cache-bootstrap", daemon=True).start()
    _acq_log("cache_bootstrap", {"counties": [n for _, n in target_pairs]})
    return jsonify({"started": True,
                    "counties": [{"fips": f, "name": n} for f, n in target_pairs]})


@acq_bp.route("/api/acq/export/<fmt>")
@_login_required
def acq_api_export(fmt):
    if not _last_search_cache["data"]:
        return jsonify({"error": "No search run yet."}), 400
    tracts = _last_search_cache["data"]["tracts"]
    meta = _last_search_cache["meta"]
    label = meta["label"]
    # Optional: filter to a subset of Prop_IDs (bulk-export from the tract list)
    prop_ids_param = (request.args.get("prop_ids") or "").strip()
    if prop_ids_param:
        wanted = set(p for p in prop_ids_param.split(",") if p)
        if wanted:
            tracts = {
                "type": "FeatureCollection",
                "features": [f for f in tracts.get("features", [])
                             if str((f.get("properties") or {}).get("Prop_ID") or "") in wanted],
            }
            label = (label or "search") + f"_selected{len(tracts['features'])}"
    if fmt == "kml":
        buf = build_kml(tracts)
        return send_file(buf, mimetype="application/vnd.google-earth.kml+xml",
                          as_attachment=True, download_name=_export_filename(label, "kml"))
    if fmt == "kmz":
        buf = build_kmz(tracts)
        return send_file(buf, mimetype="application/vnd.google-earth.kmz",
                          as_attachment=True, download_name=_export_filename(label, "kmz"))
    if fmt == "xlsx":
        buf = build_excel(tracts)
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                          as_attachment=True, download_name=_export_filename(label, "xlsx"))
    if fmt == "pdf":
        buf = build_pdf(tracts, meta)
        return send_file(buf, mimetype="application/pdf",
                          as_attachment=True, download_name=_export_filename(label, "pdf"))
    return jsonify({"error": "Unknown format"}), 400


@acq_bp.route("/api/acq/outreach-campaign/<folder_id>")
@_login_required
def acq_api_outreach_campaign(folder_id):
    """Build a multi-page PDF that contains a detailed summary for every
    pinned tract in the folder. One page per tract with map + property block
    + outreach log + notes thread, preceded by a cover page with overall
    folder stats and a table of contents.

    Enriches each tract from the parcel cache + a live HCAD/MCAD overlay so
    ownership data is current; rendering is parallelized for speed."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph,
                                       Spacer, Image as RLImage, PageBreak)
    import acq_parcels as parcel_cache

    with _storage_lock:
        s = _load_storage()
    folder = next((f for f in s.get("folders", []) if f.get("id") == folder_id), None)
    if not folder:
        return jsonify({"error": "folder not found"}), 404
    pins = [p for p in s.get("tract_pins", []) if p.get("folder_id") == folder_id]
    if not pins:
        return jsonify({"error": "no tracts in this folder"}), 404

    outreach_by_pid = {r.get("prop_id"): r for r in s.get("outreach", [])}
    # All notes per prop_id (newest first) so each tract section can show its full thread.
    notes_by_pid = {}
    for n in sorted(s.get("notes", []),
                    key=lambda n: n.get("created_at") or n.get("updated_at") or "",
                    reverse=True):
        if (n.get("text") or "").strip() and n.get("prop_id"):
            notes_by_pid.setdefault(n["prop_id"], []).append(n)
    user_lookup = {u["id"]: u for u in s.get("users", [])}

    # ---- Enrich each pin in parallel: parcel-cache lookup, then HCAD/MCAD live overlay
    def enrich(p):
        pid = str(p.get("prop_id") or "")
        if not pid:
            return None
        try:
            feat = parcel_cache.find_parcel_by_pid(pid)
        except Exception as e:
            print(f"[campaign] cache lookup failed for {pid}: {e}", flush=True)
            feat = None
        if feat:
            fc = {"type": "FeatureCollection", "features": [feat]}
            try: hcad_live_overlay(fc)
            except Exception as e: print(f"[campaign] hcad overlay failed for {pid}: {e}", flush=True)
            try: mcad_live_overlay(fc)
            except Exception as e: print(f"[campaign] mcad overlay failed for {pid}: {e}", flush=True)
        else:
            # No cache hit — fall back to whatever's stored on the pin so the
            # tract still shows up in the PDF (just without a map).
            feat = {"type": "Feature", "geometry": None, "properties": {}}
        props = feat.setdefault("properties", {})
        # Backfill from pin for any fields the cache didn't have
        for pin_key, prop_key in [("owner_name", "OWNER_NAME"), ("mail_addr", "MAIL_ADDR"),
                                   ("site_addr", "SITUS_ADDR"), ("county", "_county"),
                                   ("acres", "Acres")]:
            if not props.get(prop_key) and p.get(pin_key):
                props[prop_key] = p[pin_key]
        if not props.get("Prop_ID"):
            props["Prop_ID"] = pid
        return feat

    with ThreadPoolExecutor(max_workers=8) as pool:
        features = [f for f in pool.map(enrich, pins) if f]

    # ---- Pre-render every tract's map snippet in parallel (network-bound tile fetches)
    def render_map(f):
        if not f.get("geometry"):
            return None
        try:
            tracts_fc = {"type": "FeatureCollection", "features": [f]}
            return _build_search_map_png(tracts_fc, {"label": folder["name"]},
                                          width=900, height=340)
        except Exception as e:
            print(f"[campaign] map render failed: {e}", flush=True)
            return None

    with ThreadPoolExecutor(max_workers=4) as pool:
        map_pngs = list(pool.map(render_map, features))

    # ---- Build the PDF
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=0.45*inch,
                            rightMargin=0.45*inch, topMargin=0.45*inch, bottomMargin=0.45*inch)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("title", parent=styles["Heading1"], fontName="Helvetica-Bold",
                                  fontSize=22, textColor=colors.HexColor("#13344E"), spaceAfter=2)
    folder_style = ParagraphStyle("folder", parent=styles["Normal"], fontName="Helvetica-Bold",
                                   fontSize=14, textColor=colors.HexColor("#F25929"), spaceAfter=14)
    body_style = ParagraphStyle("body", parent=styles["Normal"], fontName="Helvetica",
                                 fontSize=9, textColor=colors.HexColor("#13344E"), leading=11)
    body_muted = ParagraphStyle("muted", parent=body_style, textColor=colors.HexColor("#7A828D"))
    cover_h_style = ParagraphStyle("ch", parent=styles["Normal"], fontName="Helvetica-Bold",
                                    fontSize=11, textColor=colors.HexColor("#F25929"), spaceBefore=12,
                                    spaceAfter=8, leading=14)
    tract_title_style = ParagraphStyle("ttitle", parent=styles["Heading1"], fontName="Helvetica-Bold",
                                        fontSize=15, textColor=colors.HexColor("#13344E"), spaceAfter=2)
    tract_sub_style = ParagraphStyle("tsub", parent=styles["Normal"], fontName="Helvetica",
                                      fontSize=9, textColor=colors.HexColor("#58595B"), spaceAfter=8)
    section_h_style = ParagraphStyle("sh", parent=styles["Normal"], fontName="Helvetica-Bold",
                                      fontSize=10, textColor=colors.HexColor("#F25929"), spaceBefore=8,
                                      spaceAfter=4, leading=12)

    story = []

    # ===== Cover page =====
    story.append(Paragraph("Outreach Campaign", title_style))
    story.append(Paragraph(folder["name"], folder_style))

    total_acres = sum((f.get("properties") or {}).get("Acres", 0) or 0 for f in features)
    counties = sorted({(f.get("properties") or {}).get("_county") for f in features
                       if (f.get("properties") or {}).get("_county")})
    story.append(Paragraph(
        f"<b>{len(features)}</b> tracts &nbsp;·&nbsp; "
        f"<b>{total_acres:,.0f}</b> acres total &nbsp;·&nbsp; "
        f"{len(counties)} {'county' if len(counties) == 1 else 'counties'}"
        f"{(' (' + ', '.join(counties[:5]) + ')') if counties else ''}"
        f" &nbsp;·&nbsp; "
        f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}", body_style))

    # Status breakdown
    status_counts = {}
    for f in features:
        pid = (f.get("properties") or {}).get("Prop_ID", "")
        o = outreach_by_pid.get(pid)
        st_v = (o.get("status") if o else None) or "lead"
        status_counts[st_v] = status_counts.get(st_v, 0) + 1
    if status_counts:
        story.append(Paragraph("STATUS BREAKDOWN", cover_h_style))
        sb_rows = [["Status", "Tracts"]]
        for v, l, _ in OUTREACH_STATUSES:
            ct = status_counts.get(v, 0)
            if ct:
                sb_rows.append([l, str(ct)])
        sb = Table(sb_rows, colWidths=[3.0*inch, 1.0*inch])
        sb.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#13344E")),
            ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
            ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",   (0, 0), (-1, -1), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F1F2F3")]),
            ("GRID",       (0, 0), (-1, -1), 0.25, colors.HexColor("#D1D3D4")),
            ("LEFTPADDING",(0, 0), (-1, -1), 6),
            ("ALIGN",      (1, 0), (1, -1), "RIGHT"),
        ]))
        story.append(sb)

    # Table of contents
    story.append(Paragraph("TRACTS IN THIS CAMPAIGN", cover_h_style))
    toc_rows = [["#", "Owner / property", "Acres", "County", "Status"]]
    for i, f in enumerate(features, start=1):
        props = f.get("properties") or {}
        pid = props.get("Prop_ID", "")
        o = outreach_by_pid.get(pid)
        st_label = next((l for v, l, _ in OUTREACH_STATUSES if v == (o.get("status") if o else None)),
                        "—")
        owner = (props.get("OWNER_NAME") or "").strip()
        if not owner or owner.upper() in ("?", "UNKNOWN OWNER", "(NO OWNER)"):
            # Fall back to site/mailing address so the row says something useful
            owner = (props.get("SITUS_ADDR") or props.get("MAIL_ADDR") or f"Prop ID {pid}")
        owner = owner[:48]
        acres = props.get("Acres", 0) or 0
        toc_rows.append([str(i), owner, f"{acres:,.1f}",
                         (props.get("_county") or "")[:14], st_label])
    toc = Table(toc_rows, colWidths=[0.4*inch, 4.0*inch, 0.8*inch, 1.0*inch, 1.1*inch])
    toc.setStyle(TableStyle([
        ("BACKGROUND",     (0, 0), (-1, 0), colors.HexColor("#13344E")),
        ("TEXTCOLOR",      (0, 0), (-1, 0), colors.white),
        ("FONTNAME",       (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",       (0, 0), (-1, -1), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F1F2F3")]),
        ("GRID",           (0, 0), (-1, -1), 0.25, colors.HexColor("#D1D3D4")),
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",    (0, 0), (-1, -1), 4),
        ("ALIGN",          (2, 1), (2, -1), "RIGHT"),
        ("ALIGN",          (0, 0), (0, -1), "CENTER"),
    ]))
    story.append(toc)

    # ===== Per-tract pages =====
    def kv(label, value):
        return [Paragraph(f"<b>{label}</b>", body_style),
                Paragraph(str(value) if value not in (None, "", 0) else "—", body_style)]

    def col_table(rows):
        t = Table(rows, colWidths=[1.2*inch, 2.4*inch])
        t.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
        ]))
        return t

    for i, f in enumerate(features, start=1):
        story.append(PageBreak())
        props = f.get("properties") or {}
        pid = props.get("Prop_ID", "")
        owner = (props.get("OWNER_NAME") or "").strip()
        owner_display = owner if (owner and owner.upper() not in ("?", "UNKNOWN OWNER", "(NO OWNER)")) \
            else (props.get("SITUS_ADDR") or props.get("MAIL_ADDR") or f"Prop ID {pid}")
        acres = props.get("Acres", 0) or 0
        county = props.get("_county") or ""

        story.append(Paragraph(f"<font color='#7A828D'>#{i}</font>&nbsp;&nbsp;{owner_display}",
                               tract_title_style))
        story.append(Paragraph(
            f"<b>{acres:,.1f} ac</b> &nbsp;·&nbsp; "
            f"{county + ' County' if county else 'County —'} &nbsp;·&nbsp; "
            f"Prop ID {pid}",
            tract_sub_style))

        # Map snippet (skip if no geometry — falls through with no error)
        png = map_pngs[i - 1] if i - 1 < len(map_pngs) else None
        if png:
            img = RLImage(io.BytesIO(png), width=7.4*inch, height=2.8*inch)
            img.hAlign = "CENTER"
            story.append(img)
            story.append(Spacer(1, 4))
        else:
            story.append(Paragraph(
                "<i>No parcel geometry on file — map omitted.</i>", body_muted))
            story.append(Spacer(1, 4))

        # Two-column property block
        mkt   = props.get("MARKET_VAL") or props.get("total_market_val")
        appr  = props.get("total_appraised_val")
        taxv  = props.get("tax_value")
        since = props.get("_owner_since") or "—"
        src   = _tract_source_label(props)

        left = [
            kv("Owner",         owner or "(no owner on file)"),
            kv("Mailing",       props.get("MAIL_ADDR") or ""),
            kv("Site address",  props.get("SITUS_ADDR") or "(no street address)"),
            kv("Legal desc",    (props.get("LEGAL_DESC") or "")[:120]),
            kv("County",        county),
            kv("Acres",         f"{acres:,.1f}"),
        ]
        right = [
            kv("Source",        src),
            kv("Owner since",   since),
            kv("Market value",  f"${mkt:,.0f}" if mkt else "—"),
            kv("Appraised",     f"${appr:,.0f}" if appr else "—"),
            kv("Taxable",       f"${taxv:,.0f}" if taxv else "—"),
            kv("Flood %",       props.get("_flood_pct", "—")),
        ]
        two_col = Table([[col_table(left), col_table(right)]],
                        colWidths=[3.9*inch, 3.9*inch])
        two_col.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
        story.append(two_col)

        # Outreach block
        o = outreach_by_pid.get(pid)
        if o:
            status = o.get("status", "lead")
            status_label = next((l for v, l, _ in OUTREACH_STATUSES if v == status), status)
            story.append(Paragraph(f"OUTREACH — {status_label.upper()}", section_h_style))
            meta_bits = []
            if o.get("broker_name"): meta_bits.append(f"Broker: <b>{o['broker_name']}</b>")
            if o.get("next_action"):
                bit = f"Next: <b>{o['next_action']}</b>"
                if o.get("next_action_date"): bit += f" by {o['next_action_date']}"
                meta_bits.append(bit)
            if meta_bits:
                story.append(Paragraph(" &nbsp;·&nbsp; ".join(meta_bits), body_style))
            log = (o.get("log") or [])[-5:]   # last 5 entries
            for e in reversed(log):
                who = user_lookup.get(e.get("by"))
                who_name = who["name"] if who else "Someone"
                when = (e.get("at") or "")[:16].replace("T", " ")
                story.append(Paragraph(
                    f"<font color='#9a9a9a'>[{when} · {who_name} · {e.get('method','')}]</font> "
                    f"{(e.get('notes','') or '')[:280]}",
                    body_style))
        else:
            story.append(Paragraph("OUTREACH — NOT STARTED", section_h_style))
            story.append(Paragraph("<i>No outreach record yet for this tract.</i>", body_muted))

        # Notes thread
        tract_notes = notes_by_pid.get(pid, [])
        if tract_notes:
            story.append(Paragraph(f"NOTES &nbsp; <font color='#7A828D'>({len(tract_notes)} entries — newest first)</font>",
                                   section_h_style))
            for n in tract_notes[:6]:   # cap at 6 entries to keep PDF compact
                u = user_lookup.get(n.get("owner_id"))
                who = u.get("name") if u else "?"
                when = (n.get("created_at") or n.get("updated_at") or "")[:16].replace("T", " ")
                story.append(Paragraph(
                    f"<font color='#9a9a9a'>[{when} · {who}]</font> "
                    f"{(n.get('text') or '')[:500]}",
                    body_style))
            if len(tract_notes) > 6:
                story.append(Paragraph(
                    f"<i>… {len(tract_notes) - 6} older note(s) omitted. Open the tract in-app for the full thread.</i>",
                    body_muted))

    doc.build(story)
    buf.seek(0)
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in folder["name"])
    fname = f"campaign_{safe}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    _acq_log("export_outreach_campaign",
                {"folder_id": folder_id, "tracts": len(features)})
    return send_file(buf, mimetype="application/pdf",
                      as_attachment=True, download_name=fname)


@acq_bp.route("/api/acq/import-boundary", methods=["POST"])
@_login_required
def acq_api_import_boundary():
    """Accept a .kmz or .kml upload. Returns a GeoJSON Feature containing
    a Polygon or MultiPolygon built from the file's polygons / closed paths."""
    f = request.files.get("file")
    if f is None:
        return jsonify({"error": "no file uploaded (form field: 'file')"}), 400
    raw = f.read()
    if not raw:
        return jsonify({"error": "uploaded file is empty"}), 400

    # KMZ files start with the ZIP magic bytes 'PK'
    if raw[:2] == b"PK":
        import zipfile
        try:
            zf = zipfile.ZipFile(io.BytesIO(raw))
        except zipfile.BadZipFile:
            return jsonify({"error": "file looks like a KMZ but is not a valid zip"}), 400
        kml_names = [n for n in zf.namelist() if n.lower().endswith(".kml")]
        if not kml_names:
            return jsonify({"error": "KMZ has no .kml file inside"}), 400
        # Prefer doc.kml if present; otherwise first .kml
        target = "doc.kml" if "doc.kml" in kml_names else kml_names[0]
        kml_bytes = zf.read(target)
    else:
        kml_bytes = raw

    try:
        polygons = _kml_to_polygons(kml_bytes)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    if not polygons:
        return jsonify({"error": "No polygons (or closed paths) found in this file. "
                                  "Make sure your KMZ contains at least one Polygon, "
                                  "or a Path/LineString whose first and last points coincide."}), 400

    if len(polygons) == 1:
        geom = {"type": "Polygon", "coordinates": polygons[0]}
    else:
        # MultiPolygon's coordinates is a list of polygon-coordinate-arrays
        geom = {"type": "MultiPolygon", "coordinates": polygons}

    _acq_log("import_boundary", {"filename": f.filename, "polygons": len(polygons)})
    return jsonify({
        "type": "Feature",
        "geometry": geom,
        "properties": {"source": "kmz_import", "filename": f.filename, "polygon_count": len(polygons)},
    })


@acq_bp.route("/api/acq/import-corridors", methods=["POST"])
@_login_required
def acq_api_import_corridors():
    """Bulk import multiple KMZ/KML files. Each polygon in a file becomes a
    saved search as-is; each OPEN LineString gets buffered by `buffer_miles`
    (default 1.0 mi on each side) and saved as a corridor search. Returns a
    list of created searches plus any per-file errors.

    Form fields:
      • files         — one or more file uploads (KMZ or KML)
      • buffer_miles  — half-width of the corridor in miles (default 1.0)
      • folder_id     — optional folder to drop the new searches into
    """
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "no files uploaded (form field: 'files')"}), 400
    try:
        buffer_miles = float(request.form.get("buffer_miles", "1.0"))
    except ValueError:
        buffer_miles = 1.0
    if not (0 < buffer_miles <= 50):
        return jsonify({"error": "buffer_miles must be between 0 and 50"}), 400
    buffer_m = buffer_miles * 1609.34
    folder_id = request.form.get("folder_id") or None

    # `replace=1` form field overrides the default dedupe behavior — when set,
    # any existing corridor whose source filename matches an incoming file is
    # deleted first, then the new one is created with the user's current buffer.
    replace_existing = (request.form.get("replace", "0") == "1")

    uid = _current_user()["id"]
    imported = []
    skipped  = []
    errors   = []

    # Index existing corridors by source filename (per user) so we can dedupe.
    from shapely.geometry import shape as shp_shape
    with _storage_lock:
        s = _load_storage()
        existing_by_file = {}
        for ex in s.get("searches", []):
            if ex.get("owner_id") != uid: continue
            ep = (ex.get("polygon") or {}).get("properties") or {}
            if ep.get("source") == "kmz_corridor" and ep.get("filename"):
                existing_by_file.setdefault(ep["filename"], []).append(ex)

        for f in files:
            raw = f.read()
            if not raw:
                errors.append({"file": f.filename, "error": "empty file"}); continue

            # ---- Dedupe check: if this filename already has a corridor, either
            # skip it (default) or replace it (when replace=1 was passed).
            already = existing_by_file.get(f.filename, [])
            if already and not replace_existing:
                skipped.append({
                    "file": f.filename,
                    "reason": f"already imported as {len(already)} corridor(s); pass replace=1 to overwrite, or use the buffer pill to change width in place",
                    "existing_ids": [ex["id"] for ex in already],
                })
                continue
            if already and replace_existing:
                ids = {ex["id"] for ex in already}
                s["searches"] = [x for x in s["searches"] if x["id"] not in ids]
                # also drop them from our local index so subsequent files in this
                # request don't re-trigger the same path
                existing_by_file[f.filename] = []
            if raw[:2] == b"PK":
                import zipfile
                try:
                    zf = zipfile.ZipFile(io.BytesIO(raw))
                except zipfile.BadZipFile:
                    errors.append({"file": f.filename, "error": "not a valid KMZ"}); continue
                kml_names = [n for n in zf.namelist() if n.lower().endswith(".kml")]
                if not kml_names:
                    errors.append({"file": f.filename, "error": "no .kml inside KMZ"}); continue
                target = "doc.kml" if "doc.kml" in kml_names else kml_names[0]
                kml_bytes = zf.read(target)
            else:
                kml_bytes = raw

            try:
                polygons = _kml_to_polygons(kml_bytes)
                lines    = _kml_to_linestrings(kml_bytes)
            except ValueError as e:
                errors.append({"file": f.filename, "error": str(e)}); continue

            if not polygons and not lines:
                errors.append({"file": f.filename, "error": "no polygons or lines found"})
                continue

            base_label = (f.filename or "imported").rsplit(".", 1)[0]

            # Polygon(s) → save each as-is
            for i, rings in enumerate(polygons):
                geom = {"type": "Polygon", "coordinates": rings}
                try:
                    c = shp_shape(geom).centroid
                    lat0, lon0 = float(c.y), float(c.x)
                except Exception:
                    lat0, lon0 = 0.0, 0.0
                label = base_label + (f" ({i+1})" if len(polygons) > 1 else "")
                new = {
                    "id": uuid.uuid4().hex[:10],
                    "owner_id": uid,
                    "folder_id": folder_id,
                    "label": label,
                    "lat": lat0, "lon": lon0, "radius_mi": 0,
                    "min_acres": 1.0, "max_acres": 100000.0,
                    "group_by": None,
                    "polygon": {"type": "Feature", "geometry": geom,
                                 "properties": {"source": "kmz_import",
                                                "filename": f.filename}},
                    "saved_at": datetime.now().isoformat(timespec="seconds"),
                }
                s["searches"].append(new)
                imported.append({"file": f.filename, "label": label, "kind": "polygon",
                                  "search_id": new["id"]})

            # LineString(s) → buffer into corridor polygon, save
            for i, coords in enumerate(lines):
                outer = _buffer_line_to_polygon(coords, buffer_m)
                if not outer:
                    errors.append({"file": f.filename,
                                    "error": f"line #{i+1} buffer failed"}); continue
                geom = {"type": "Polygon", "coordinates": [outer]}
                try:
                    c = shp_shape(geom).centroid
                    lat0, lon0 = float(c.y), float(c.x)
                except Exception:
                    lat0, lon0 = 0.0, 0.0
                # Label includes buffer width so the user can tell them apart
                label = base_label + (f" ({i+1})" if len(lines) > 1 else "") \
                         + f"  ±{buffer_miles:g}mi"
                new = {
                    "id": uuid.uuid4().hex[:10],
                    "owner_id": uid,
                    "folder_id": folder_id,
                    "label": label,
                    "lat": lat0, "lon": lon0, "radius_mi": 0,
                    "min_acres": 1.0, "max_acres": 100000.0,
                    "group_by": None,
                    "polygon": {"type": "Feature", "geometry": geom,
                                 "properties": {"source":              "kmz_corridor",
                                                "filename":            f.filename,
                                                "buffer_miles":        buffer_miles,
                                                "centerline_vertices": len(coords),
                                                # Store the centerline so we can RE-BUFFER
                                                # at a different width later without re-importing.
                                                "centerline_coords":   [[round(c[0], 6),
                                                                          round(c[1], 6)]
                                                                         for c in coords],
                                                "base_label":          base_label + (f" ({i+1})" if len(lines) > 1 else "")}},
                    "saved_at": datetime.now().isoformat(timespec="seconds"),
                }
                s["searches"].append(new)
                imported.append({"file": f.filename, "label": label, "kind": "corridor",
                                  "search_id": new["id"], "buffer_miles": buffer_miles})
        _save_storage(s)

    _acq_log("import_corridors",
                {"files": len(files), "imported": len(imported),
                 "skipped": len(skipped), "errors": len(errors),
                 "buffer_miles": buffer_miles})
    return jsonify({"imported": imported, "skipped": skipped, "errors": errors,
                    "buffer_miles": buffer_miles, "folder_id": folder_id})


@acq_bp.route("/api/acq/draw-corridor", methods=["POST"])
@_login_required
def acq_api_draw_corridor():
    """Create a corridor saved-search from a centerline drawn in the browser
    (no KMZ upload required).

    Body: {
      "label":        string,                  // user-supplied name
      "buffer_miles": float (0 < x <= 50),     // half-width
      "centerline":   [[lng, lat], ...]        // ordered polyline vertices, EPSG:4326
    }

    Returns: { id, label, buffer_miles, polygon }
    """
    from shapely.geometry import shape as shp_shape

    body = request.get_json(force=True) or {}
    label = (body.get("label") or "").strip()
    if not label:
        return jsonify({"error": "label is required"}), 400
    if len(label) > 120:
        return jsonify({"error": "label too long (max 120 chars)"}), 400

    try:
        buffer_miles = float(body.get("buffer_miles", 1.0))
    except (TypeError, ValueError):
        return jsonify({"error": "buffer_miles must be a number"}), 400
    if not (0 < buffer_miles <= 50):
        return jsonify({"error": "buffer_miles must be between 0 and 50"}), 400

    centerline = body.get("centerline") or []
    if not isinstance(centerline, list) or len(centerline) < 2:
        return jsonify({"error": "centerline must be a list of at least 2 [lng, lat] points"}), 400
    coords = []
    for pt in centerline:
        if not isinstance(pt, (list, tuple)) or len(pt) < 2:
            return jsonify({"error": "each centerline point must be [lng, lat]"}), 400
        try:
            lng, lat = float(pt[0]), float(pt[1])
        except (TypeError, ValueError):
            return jsonify({"error": "centerline coords must be numeric"}), 400
        coords.append((lng, lat))

    buffer_m = buffer_miles * 1609.34
    outer = _buffer_line_to_polygon(coords, buffer_m)
    if not outer:
        return jsonify({"error": "could not buffer this line — check that the points span a real path"}), 500

    geom = {"type": "Polygon", "coordinates": [outer]}
    try:
        c = shp_shape(geom).centroid
        lat0, lon0 = float(c.y), float(c.x)
    except Exception:
        lat0, lon0 = 0.0, 0.0

    full_label = f"{label}  ±{buffer_miles:g}mi"
    uid = _current_user()["id"]
    new = {
        "id": uuid.uuid4().hex[:10],
        "owner_id": uid,
        "folder_id": None,
        "label": full_label,
        "lat": lat0, "lon": lon0, "radius_mi": 0,
        "min_acres": 1.0, "max_acres": 100000.0,
        "group_by": None,
        "polygon": {"type": "Feature", "geometry": geom,
                    "properties": {"source":              "kmz_corridor",   # same shape as KMZ-imported corridors
                                    "filename":            "(drawn in-app)",
                                    "buffer_miles":        buffer_miles,
                                    "centerline_vertices": len(coords),
                                    "centerline_coords":   [[round(c[0], 6), round(c[1], 6)] for c in coords],
                                    "base_label":          label,
                                    "drawn":               True}},
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    with _storage_lock:
        s = _load_storage()
        s["searches"].append(new)
        _save_storage(s)

    _acq_log("draw_corridor", {"label": label, "buffer_miles": buffer_miles,
                                    "vertices": len(coords)})
    return jsonify({"id": new["id"], "label": full_label,
                    "buffer_miles": buffer_miles, "polygon": new["polygon"]})


@acq_bp.route("/api/acq/searches/<sid>/rebuffer", methods=["POST"])
@_login_required
def acq_api_search_rebuffer(sid):
    """Re-buffer a corridor search at a new half-width (miles), in place.

    Requires the search to have been imported by api_import_corridors so that
    its `polygon.properties.centerline_coords` is populated. Returns the new
    polygon + updated label.

    Body: {"buffer_miles": float}
    """
    body = request.get_json(force=True) or {}
    try:
        buffer_miles = float(body.get("buffer_miles", 1.0))
    except (TypeError, ValueError):
        return jsonify({"error": "buffer_miles must be a number"}), 400
    if not (0 < buffer_miles <= 50):
        return jsonify({"error": "buffer_miles must be between 0 and 50"}), 400

    uid = _current_user()["id"]
    with _storage_lock:
        s = _load_storage()
        search = next((x for x in s.get("searches", []) if x.get("id") == sid), None)
        if not search:
            return jsonify({"error": "search not found"}), 404
        if search.get("owner_id") != uid:
            return jsonify({"error": "not your search"}), 403

        poly = search.get("polygon") or {}
        props = (poly.get("properties") or {}) if isinstance(poly, dict) else {}
        centerline = props.get("centerline_coords")
        if not centerline or len(centerline) < 2:
            return jsonify({
                "error": ("This corridor was imported before we started storing the "
                          "centerline. Delete it and re-import the KMZ to enable "
                          "buffer editing.")
            }), 400

        outer = _buffer_line_to_polygon(centerline, buffer_miles * 1609.34)
        if not outer:
            return jsonify({"error": "buffer computation failed"}), 500

        new_geom = {"type": "Polygon", "coordinates": [outer]}
        # Refresh centroid
        try:
            from shapely.geometry import shape as shp_shape
            c = shp_shape(new_geom).centroid
            search["lat"], search["lon"] = float(c.y), float(c.x)
        except Exception:
            pass

        # Update props + label
        props["buffer_miles"] = buffer_miles
        poly["geometry"] = new_geom
        poly["properties"] = props
        search["polygon"] = poly

        # Rebuild the label so the buffer width in the name matches the geometry
        base = props.get("base_label") or _strip_buffer_suffix(search.get("label", ""))
        search["label"] = f"{base}  ±{buffer_miles:g}mi"
        search["saved_at"] = datetime.now().isoformat(timespec="seconds")
        _save_storage(s)

    _acq_log("rebuffer_corridor", {"search_id": sid, "buffer_miles": buffer_miles})
    return jsonify({
        "ok": True,
        "id": sid,
        "label": search["label"],
        "buffer_miles": buffer_miles,
        "polygon": search["polygon"],
    })


# ══════════════════════════════════════════════════════════════════════════
# Remaining pages
#
# Corridors, folders, pipeline and the single-tract view. Each is the
# standalone app's page body and JS inside the portal shell, so the sidebar,
# theme and account modal come from EmberApps rather than from the old app.
# ══════════════════════════════════════════════════════════════════════════

def _acq_render(template, **extra):
    if not _can_view_acquisitions():
        return redirect(url_for("home"))
    return render_template(
        template,
        username=session.get("username"),
        display_name=session.get("display_name", session.get("username")),
        is_admin=session.get("is_admin", False),
        page_access=session.get("page_access") or {},
        **extra)


@acq_bp.route("/acquisitions/corridors")
@_login_required
def acq_corridors_page():
    return _acq_render("acquisitions_corridors.html")


@acq_bp.route("/acquisitions/admin")
@_login_required
def acquisitions_admin_page():
    """Parcel-cache admin: bootstrap counties, watch progress, vacuum the index.

    Admin-only, and separate from the portal's own /admin — this manages the
    acquisitions parcel cache, nothing else. Without it the cache API existed
    with no way to reach it, so there was nowhere to bootstrap a county from.
    """
    if not session.get("is_admin"):
        return redirect(url_for("acq.acquisitions_page"))
    return render_template(
        "acquisitions_admin.html",
        username=session.get("username"),
        display_name=session.get("display_name", session.get("username")),
        is_admin=True,
        page_access=session.get("page_access") or {},
    )


@acq_bp.route("/acquisitions/folders")
@_login_required
def acq_folders_page():
    return _acq_render("acquisitions_folders.html")


@acq_bp.route("/acquisitions/analyses")
@_login_required
def acquisitions_analyses_page():
    """Index of every project and quick analysis.

    The map page keeps a Projects list in its sidebar, but once you navigate
    away there was no route back to a saved analysis except browser history.
    """
    if not _can_view_acquisitions():
        return redirect(url_for("home"))
    return render_template(
        "acquisitions_analyses.html",
        username=session.get("username"),
        display_name=session.get("display_name", session.get("username")),
        is_admin=session.get("is_admin", False),
        page_access=session.get("page_access") or {},
    )


@acq_bp.route("/acquisitions/pipeline")
@_login_required
def acq_pipeline_page():
    return _acq_render("acquisitions_pipeline.html")


@acq_bp.route("/acquisitions/tract/<prop_id>")
@_login_required
def acq_tract_page(prop_id):
    county = request.args.get("county", "")
    return _acq_render("acquisitions_tract.html", prop_id=prop_id, county=county)
