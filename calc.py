"""
Ember Tract Underwriting - Pure Python Calculation Engine
Faithfully ports all Excel formulas from the 10-sheet model.
No Excel or openpyxl required at runtime.
"""
import math
import datetime
import calendar
from typing import Any

# ─── LOOKUP TABLES (from Calc_Lookups sheet) ─────────────────────────────────
PLANT_LOOKUPS = {
    "WWTP":        {"acres": 10.0,  "duration": 8},
    "Water Plant": {"acres": 3.5,   "duration": 8},
    "Lift Station":{"acres": 0.75,  "duration": 3},
    "None":        {"acres": 0.0,   "duration": 0},
}
AMENITY_LOOKUPS = {
    "Pocket Park":          {"duration": 2,  "acres": 0.5},
    "Small Amenity Center": {"duration": 8,  "acres": 3.0},
    "Large Amenity Center": {"duration": 12, "acres": 6.0},
    "None":                 {"duration": 0,  "acres": 0.0},
}
ROAD_LOOKUPS = {
    "2 Lane": {"wsd": 450, "paving": 343},
    "4 Lane": {"wsd": 460, "paving": 663},
}

# Lot size front footage by row index (rows 0-15 = 25, 30, 35, ..., 100 FF)
LOT_FF_BY_INDEX = [25 + 5 * i for i in range(16)]

def safe(v, default=0):
    """Return v if numeric, else default."""
    if v is None or v == "" or (isinstance(v, float) and math.isnan(v)):
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default

def getd(inp, key, default=0):
    """Get input value treating None/missing/empty string as default (not 0)."""
    v = inp.get(key)
    if v is None or v == "":
        return default
    return safe(v, default)

def mround(v, multiple):
    """Excel MROUND equivalent."""
    if multiple == 0:
        return 0
    return round(v / multiple) * multiple

def iferr(v, default=0):
    """Return v if not None/nan else default."""
    try:
        f = float(v)
        if math.isnan(f) or math.isinf(f):
            return default
        return f
    except Exception:
        return default

def npv_irr(cashflows, guess=0.1, max_iter=1000):
    """Newton-Raphson IRR on monthly cashflows, returns annual rate."""
    if not any(c < 0 for c in cashflows) or not any(c > 0 for c in cashflows):
        return None
    r = guess / 12  # monthly
    for _ in range(max_iter):
        npv = sum(c / (1 + r) ** t for t, c in enumerate(cashflows))
        dnpv = sum(-t * c / (1 + r) ** (t + 1) for t, c in enumerate(cashflows))
        if dnpv == 0:
            break
        r_new = r - npv / dnpv
        if abs(r_new - r) < 1e-8:
            r = r_new
            break
        r = r_new
    annual = (1 + r) ** 12 - 1
    return annual if -0.5 < annual < 10 else None


def xirr(cashflows, dates, guess=0.1, max_iter=1000):
    """Newton-Raphson XIRR matching Excel's 365-day convention.
    cashflows: list of floats, dates: list of date objects.
    Tries multiple initial guesses if Newton-Raphson fails to converge."""
    if not any(c < 0 for c in cashflows) or not any(c > 0 for c in cashflows):
        return None
    d0 = dates[0]
    # Year fractions from first date (Excel: actual/365)
    yf = [(d - d0).days / 365.0 for d in dates]
    # Filter to non-zero cashflows for efficiency
    pairs = [(c, t) for c, t in zip(cashflows, yf) if c != 0]

    def _try_solve(r):
        for _ in range(max_iter):
            npv = sum(c / (1 + r) ** t for c, t in pairs)
            dnpv = sum(-t * c / (1 + r) ** (t + 1) for c, t in pairs)
            if abs(dnpv) < 1e-14:
                return None
            r_new = r - npv / dnpv
            if r_new <= -1:
                return None  # prevent divergence
            if abs(r_new - r) < 1e-9:
                return r_new
            r = r_new
        return r if abs(sum(c / (1 + r) ** t for c, t in pairs)) < 1.0 else None

    for g in [guess, 0.05, 0.15, 0.25, 0.01, -0.05]:
        result = _try_solve(g)
        if result is not None and -0.5 < result < 10:
            return result
    return None


def _end_of_month(year, month):
    """Return the last day of the given month."""
    return datetime.date(year, month, calendar.monthrange(year, month)[1])


def calculate(inp: dict) -> dict:
    """
    Master calculation function.
    inp: flat dict of all user inputs (see schema in app.py)
    Returns: dict of all computed outputs for display.
    """
    out = {}

    # ── 1. TRACT INPUTS derived values ───────────────────────────────────────
    gross_ac      = getd(inp, "gross_acreage",          0)
    ppa           = getd(inp, "purchase_price_per_acre", 0)
    escalator     = getd(inp, "land_escalator",          0.05)   # Excel default 5%
    closing_pct   = getd(inp, "closing_costs_pct",       0.045)  # Excel default 4.5%

    purchase_price = ppa * gross_ac                        # B10
    closing_costs  = closing_pct * purchase_price          # B12
    total_land     = purchase_price + closing_costs         # B13

    # Build effective lookups — user-editable lk_ inputs override hardcoded constants
    lk_plants_inp = inp.get("lk_plants", [])
    eff_plant_lk = dict(PLANT_LOOKUPS)
    for row in lk_plants_inp:
        pt = row.get("type","").strip()
        if pt:
            eff_plant_lk[pt] = {"acres": safe(row.get("acres",0)), "duration": int(safe(row.get("duration",0)))}

    lk_amenities_inp = inp.get("lk_amenities", [])
    eff_amenity_lk = dict(AMENITY_LOOKUPS)
    for row in lk_amenities_inp:
        pt = row.get("type","").strip()
        if pt:
            acres_val = safe(row.get("acres", AMENITY_LOOKUPS.get(pt, {}).get("acres", 0)))
            eff_amenity_lk[pt] = {"duration": int(safe(row.get("duration",0))), "acres": acres_val}

    lk_roads_inp = inp.get("lk_roads", [])
    eff_road_lk = dict(ROAD_LOOKUPS)
    for row in lk_roads_inp:
        pt = row.get("type","").strip()
        if pt:
            eff_road_lk[pt] = {"wsd": safe(row.get("wsd",0)), "paving": safe(row.get("paving",0))}

    # Plant acres lookup
    plants = inp.get("plants", [])  # list of {type, notes}
    plant_acres_list = [eff_plant_lk.get(p.get("type","None"), {"acres":0})["acres"] for p in plants]
    total_plant_acres = sum(plant_acres_list)

    # Detention
    det_rate        = getd(inp, "det_storage_rate", 0)
    det_depth       = getd(inp, "det_depth",         0)
    det_num         = getd(inp, "det_num_projects",  0)
    parks_pct       = getd(inp, "parks_pct",         0.03)  # Excel: 3%
    drill_site_ac   = getd(inp, "drill_site_acres",  0)

    # Net-out area (B72): plants + detention + amenities + other + roads
    amenities = inp.get("amenities", [])  # list of {type, acres}
    amenity_acres_total = sum(safe(a.get("acres")) for a in amenities)

    other_netouts = inp.get("other_netouts", [])  # list of {desc, acres}
    other_acres_total = sum(safe(o.get("acres")) for o in other_netouts)

    roads = inp.get("roads", [])  # list of {type, lf, width, road_setback, landscaping_setback}
    road_acres_list = []
    for r in roads:
        lf = safe(r.get("linear_feet") or r.get("lf"))
        w  = safe(r.get("width"))
        rs = safe(r.get("road_setback"))
        ls = safe(r.get("landscaping_setback"))
        if lf:
            road_acres_list.append(lf * (w + (ls + rs) * 2) / 43560)
        else:
            road_acres_list.append(0)
    road_acres_total = sum(road_acres_list)

    # Detention acres: Excel B32 = (storage_rate * gross_ac / depth) * 1.3 = total footprint
    # Excel B45 (each project) = B32 / num_projects
    det_total_footprint = (det_rate * gross_ac / det_depth * 1.3) if det_depth else 0
    det_acres_each = det_total_footprint / det_num if det_num else 0
    det_total_acres = det_total_footprint  # Total detention acres for net-outs = B32

    parks_acres = parks_pct * gross_ac
    net_out_total = total_plant_acres + det_total_acres + amenity_acres_total + parks_acres + drill_site_ac + other_acres_total + road_acres_total  # B72
    dev_acres = gross_ac - net_out_total  # B73

    comm_pod_acres = safe(inp.get("commercial_pod_acres"))
    res_pod_acres  = safe(inp.get("residential_pod_acres"))
    residential_dev_acres = dev_acres - comm_pod_acres - res_pod_acres  # B78

    out["purchase_price"] = purchase_price
    out["closing_costs"]  = closing_costs
    out["total_land_cost"]= total_land
    out["gross_acreage"]  = gross_ac
    out["plant_net_out_acres"]   = round(total_plant_acres, 2)
    out["det_net_out_acres"]     = round(det_total_acres, 2)
    out["amenity_net_out_acres"] = round(amenity_acres_total, 2)
    out["parks_net_out_acres"]   = round(parks_acres, 2)
    out["other_net_out_acres"]   = round(other_acres_total, 2)
    out["road_net_out_acres"]    = round(road_acres_total, 2)
    out["net_out_total"]  = net_out_total
    out["dev_acres"]      = dev_acres
    out["comm_pod_acres_net"] = round(comm_pod_acres, 2)
    out["res_pod_acres_net"]  = round(res_pod_acres, 2)
    out["residential_dev_acres"] = residential_dev_acres
    out["det_acres_each"] = det_acres_each
    out["det_total_acres"] = det_total_acres
    out["det_total_footprint"] = det_total_footprint
    out["plant_acres_list"] = plant_acres_list
    out["road_acres_list"]  = road_acres_list

    # ── 2. COST INPUTS derived values ─────────────────────────────────────────
    default_other_pct      = getd(inp, "default_other_pct",     0.17)
    contingency            = getd(inp, "contingency",            0.05)
    cost_per_mailbox       = getd(inp, "cost_per_mailbox",       200)    # Excel: $200
    cost_per_streetlight   = getd(inp, "cost_per_streetlight",   1700)   # Excel: $1,700
    default_start_month    = getd(inp, "default_start_month",    1)
    site_work_pct          = getd(inp, "site_work_pct",          0.01)   # Excel: 1%
    fenced_pct             = getd(inp, "fenced_pct",             0.25)   # Excel: 25%
    sectional_other_pct    = getd(inp, "sectional_other_pct",    0.17)
    landscaping_other_pct  = getd(inp, "landscaping_other_pct",  0.12)  # Excel: 12%

    # Land cost takedowns — matching Excel Cost Inputs rows 17-21
    # Take 1 (period 0): no escalation on purchase; closings also no escalation
    # Take 2+: purchase escalated by (1+escalator)^(period/12)
    # Closing costs never escalate (Excel B20=closing_total*pct, no escalator)
    takedowns = inp.get("takedowns", [])
    # Filter out zero-pct takedowns but keep at least one
    valid_tds = [td for td in takedowns if safe(td.get("pct", 0)) > 0]
    if not valid_tds:
        valid_tds = [{"period": 0, "pct": 1.0}]
    td_rows = []
    total_td_purchase = 0
    total_td_closing = 0
    for i, td in enumerate(valid_tds):
        period = safe(td.get("period"))
        pct    = safe(td.get("pct"))
        if period == 0 or i == 0:
            purchase = purchase_price * pct
        else:
            purchase = purchase_price * pct * (1 + escalator) ** (period / 12)
        closing = closing_costs * pct  # closing costs never escalated
        total_td_purchase += purchase
        total_td_closing  += closing
        td_rows.append({"period": period, "pct": pct, "purchase": purchase, "closing": closing, "total": purchase + closing})
    # Rename fields to match app.html expectations (purchase_price, closing_costs, total_land_cost)
    land_takedowns = [{"period": r["period"], "pct": r["pct"],
                       "purchase_price": r["purchase"], "closing_costs": r["closing"],
                       "total_land_cost": r["total"]} for r in td_rows]
    out["land_takedowns"]      = land_takedowns
    out["land_check"]          = round(sum(td["pct"] for td in land_takedowns), 6)
    out["land_total_purchase"] = total_td_purchase
    out["land_total_closing"]  = total_td_closing
    out["land_total_cost"]     = total_td_purchase + total_td_closing

    # Plant cost rows
    plant_costs = inp.get("plant_costs", [])  # {base_cost, other_pct, start_month, ph2_base_cost, ph2_other_pct, ph2_start_month}
    plant_cost_rows = []
    for i in range(max(len(plants), len(plant_costs), 8)):
        pc = plant_costs[i] if i < len(plant_costs) else {}
        ptype = plants[i].get("type", "None") if i < len(plants) else "None"
        dur = eff_plant_lk.get(ptype, {"duration": 0})["duration"]
        base = safe(pc.get("base_cost"))
        other_p = safe(pc.get("other_pct", default_other_pct))
        total = base * (1 + other_p) if base else 0
        sm = safe(pc.get("start_month", default_start_month))
        ph2_base = safe(pc.get("ph2_base_cost"))
        ph2_other = safe(pc.get("ph2_other_pct", default_other_pct))
        ph2_total = ph2_base * (1 + ph2_other) if ph2_base else 0
        ph2_sm = safe(pc.get("ph2_start_month", sm + 36))
        plant_cost_rows.append({
            "type": ptype, "acres": plant_acres_list[i] if i < len(plant_acres_list) else 0,
            "base_cost": base, "other_pct": other_p, "total_cost": total,
            "start_month": sm, "duration": dur,
            "ph2_base_cost": ph2_base, "ph2_other_pct": ph2_other, "ph2_total_cost": ph2_total,
            "ph2_start_month": ph2_sm, "ph2_duration": dur,
        })
    out["plant_cost_rows"] = plant_cost_rows
    total_plant_cost = sum(r["total_cost"] + r["ph2_total_cost"] for r in plant_cost_rows)

    # Amenity cost rows
    amenity_costs = inp.get("amenity_costs", [])
    amenity_cost_rows = []
    for i in range(max(len(amenities), len(amenity_costs), 6)):
        ac = amenity_costs[i] if i < len(amenity_costs) else {}
        atype = amenities[i].get("type", "None") if i < len(amenities) else "None"
        dur = eff_amenity_lk.get(atype, {"duration": 0})["duration"]
        base = safe(ac.get("base_cost"))
        other_p = safe(ac.get("other_pct", default_other_pct))
        total = base * (1 + other_p) if base else 0
        sm = safe(ac.get("start_month", default_start_month))
        user_acres = safe(amenities[i].get("acres")) if i < len(amenities) else 0
        lookup_acres = eff_amenity_lk.get(atype, {}).get("acres", 0)
        amenity_acres = user_acres if user_acres else lookup_acres
        amenity_cost_rows.append({
            "type": atype,
            "acres": amenity_acres,
            "base_cost": base, "other_pct": other_p, "total_cost": total,
            "start_month": sm, "duration": dur,
        })
    out["amenity_cost_rows"] = amenity_cost_rows
    total_amenity_cost = sum(r["total_cost"] for r in amenity_cost_rows)

    # Detention cost rows — replicating Excel Cost Inputs formulas exactly
    # Excel: Footprint = storage_rate * gross_ac / depth * 1.3  (B32)
    # Excel: Volume CY = storage_rate * gross_ac * 43560 / 27   (B36)
    # Excel: Total base cost = volume_cy * $10/CY               (B37, A37=10 hardcoded)
    # Excel: Per-project acres = footprint / num_projects       (B45=B32/B34)
    # Excel: Per-project base cost = B37 / num_projects
    # Excel: Landscaping = lpf * 43560 * acres * (1+landscaping_other_pct) * 0.30
    DET_COST_PER_CY = 10.0  # hardcoded in Excel cell A37
    det_volume_cy = safe(inp.get("det_storage_rate", 0)) * gross_ac * 43560 / 27 if gross_ac else 0
    det_total_base = det_volume_cy * DET_COST_PER_CY
    det_base_per_proj = det_total_base / det_num if det_num else 0
    det_costs = inp.get("det_costs", [])
    det_cost_rows = []
    sm_base = safe(inp.get("default_start_month", 1))
    for idx in range(int(det_num)):
        dc = det_costs[idx] if idx < len(det_costs) else {}
        other_p = safe(dc.get("other_pct", default_other_pct))
        base_cost = det_base_per_proj
        total_cost = base_cost * (1 + other_p)
        sm = sm_base + idx * 15  # each project offset +15 months (Excel F46=F45+15)
        dur = 9  # fixed 9-month duration (Excel G45=9)
        delivery = dur + sm - 1
        lsf = safe(dc.get("landscaping_per_foot", 2))  # default $2/LF from Excel I45
        total_landscaping = lsf * 43560 * det_acres_each * (1 + landscaping_other_pct) * 0.30
        det_cost_rows.append({
            "acres": det_acres_each, "base_cost": base_cost, "other_pct": other_p,
            "total_cost": total_cost, "start_month": sm, "duration": dur,
            "delivery_period": delivery, "landscaping_per_foot": lsf,
            "total_landscaping": total_landscaping,
        })
    out["det_cost_rows"] = det_cost_rows
    total_det_cost = sum(r["total_cost"] for r in det_cost_rows)

    # Other items cost rows
    other_costs = inp.get("other_costs", [])
    other_cost_rows = []
    for i, oc in enumerate(other_costs):
        no = other_netouts[i] if i < len(other_netouts) else {}
        desc = no.get("desc") or no.get("description", "")
        acres = safe(no.get("acres")) if i < len(other_netouts) else 0
        base = safe(oc.get("base_cost"))
        other_p = safe(oc.get("other_pct", default_other_pct))
        total = base * (1 + other_p) if base else 0
        sm = safe(oc.get("start_month", default_start_month))
        dur = safe(oc.get("duration", 1))
        other_cost_rows.append({
            "desc": desc, "acres": acres, "base_cost": base,
            "other_pct": other_p, "total_cost": total,
            "start_month": sm, "duration": dur,
        })
    out["other_cost_rows"] = other_cost_rows
    total_other_cost = sum(r["total_cost"] for r in other_cost_rows)

    # Road cost rows — always 6 rows driven by tract road inputs
    road_costs = inp.get("road_costs", [])
    road_cost_rows = []
    for i in range(max(len(roads), len(road_costs), 6)):
        rc = road_costs[i] if i < len(road_costs) else {}
        rtype = roads[i].get("type", "") if i < len(roads) else ""
        lf = safe(roads[i].get("linear_feet") or roads[i].get("lf")) if i < len(roads) else 0
        ls_setback = safe(roads[i].get("landscaping_setback")) if i < len(roads) else 0
        rl = eff_road_lk.get(rtype, {"wsd": 0, "paving": 0})
        wsd_per_lf  = rl["wsd"]
        pave_per_lf = rl["paving"]
        base_cost = lf * (wsd_per_lf + pave_per_lf) if lf else 0
        other_p = safe(rc.get("other_pct", default_other_pct))
        total_cost = base_cost * (1 + other_p) if base_cost else 0
        # Excel default start months for roads: 1, 12, 48, 72, 96, 1
        rd_default_sms = [default_start_month, 12, 48, 72, 96, default_start_month]
        sm_default = rd_default_sms[i] if i < len(rd_default_sms) else default_start_month
        sm = safe(rc.get("start_month")) or sm_default
        dur = int(mround(lf / 300 + 6, 1)) if lf else 6
        delivery = dur + sm - 1
        lsf = safe(rc.get("landscaping_per_sf", 0))
        total_landscaping = lsf * ls_setback * lf * 2 * (1 + landscaping_other_pct) if lf else 0
        light_spacing = safe(rc.get("light_spacing", 0))
        total_lights = int(lf / light_spacing * 2) if light_spacing and lf else 0
        road_cost_rows.append({
            "type": rtype, "lf": lf,
            "wsd_per_lf": wsd_per_lf, "paving_per_lf": pave_per_lf,
            "base_cost": base_cost, "other_pct": other_p, "total_cost": total_cost,
            "start_month": sm, "duration": dur, "delivery_period": delivery,
            "landscaping_per_sf": lsf, "total_landscaping": total_landscaping,
            "light_spacing": light_spacing, "total_lights": total_lights,
        })
    out["road_cost_rows"] = road_cost_rows
    total_road_cost = sum(r["total_cost"] for r in road_cost_rows)
    total_road_landscaping = sum(r["total_landscaping"] for r in road_cost_rows)
    total_streetlights = sum(r["total_lights"] for r in road_cost_rows)

    # ── LOT SIZE SECTION SCHEDULE SETUP (matches Excel Calc_Costs rows 32-53) ─────
    # Excel allocates residential_dev_acres proportionally by each lot type's acres_18mo.
    # acres_18mo = pace*18 / yield_per_ac  (the acres consumed in one 18-month section)
    # share_i    = acres_18mo_i / sum(acres_18mo_j for all active j)
    # alloc_acres = residential_dev_acres * share
    # full_secs  = INT(alloc_acres / acres_18mo)     ← Excel INT() = truncate
    # last_acres = alloc_acres - full_secs * acres_18mo
    # total_lots = ROUNDDOWN(full_secs * lots_18mo + last_acres * yield, 0)
    lot_sizes = inp.get("lot_sizes", [])

    sum_18mo_acres = sum(
        (safe(ls.get("pace", 0)) * 18 / safe(ls.get("yield_per_ac", 1)))
        for ls in lot_sizes
        if safe(ls.get("on", 0)) and safe(ls.get("pace", 0)) > 0 and safe(ls.get("yield_per_ac", 0)) > 0
    )

    lot_rows = []
    for i, ls in enumerate(lot_sizes):
        ff = safe(ls.get("front_footage", LOT_FF_BY_INDEX[i] if i < len(LOT_FF_BY_INDEX) else 25 + 5 * i))
        if ff <= 0:
            ff = LOT_FF_BY_INDEX[i] if i < len(LOT_FF_BY_INDEX) else 25 + 5 * i

        if not safe(ls.get("on", 0)):
            lot_rows.append({**ls, "lots_18mo": 0, "total_lots": 0, "acres_18mo": 0,
                             "dev_cost_per_lot": 0, "ff": ff,
                             "full_sections": 0, "last_lots": 0, "last_section_acres": 0})
            continue

        yield_ac  = safe(ls.get("yield_per_ac", 0))
        pace      = safe(ls.get("pace", 0))
        lots_18mo = pace * 18          # float — e.g. 13.5 for 80FF (Excel E col)
        acres_18mo = lots_18mo / yield_ac if yield_ac else 0   # Excel G col

        # Acreage allocation (Excel Calc_Costs rows 38-53)
        if sum_18mo_acres > 0 and acres_18mo > 0:
            alloc_acres = residential_dev_acres * (acres_18mo / sum_18mo_acres)
        else:
            alloc_acres = 0

        full_secs  = int(alloc_acres / acres_18mo) if acres_18mo > 0 else 0  # Excel INT()
        last_acres = alloc_acres - full_secs * acres_18mo if acres_18mo > 0 else 0
        last_lots  = math.floor(last_acres * yield_ac) if yield_ac else 0    # Excel ROUNDDOWN
        # Excel total_lots = ROUNDDOWN(full_secs*lots_18mo + last_acres*yield, 0)
        total_lots_this = math.floor(full_secs * lots_18mo + (last_acres * yield_ac if yield_ac else 0))

        # Dev cost per lot — Excel Cost Inputs K col = FF*(WSD+Paving) ONLY
        # URD, Fencing, and Streetlights are separate delivery-timed costs (not phase-scheduled)
        # Contingency (sectional_other_pct) is applied per-section in the schedule, matching Excel.
        wsd_ff   = safe(ls.get("wsd_per_ff", 0))
        pave_ff  = safe(ls.get("paving_per_ff", 0))
        dev_cost = (wsd_ff + pave_ff) * ff

        lot_rows.append({
            **ls,
            "total_lots":         total_lots_this,
            "lots_18mo":          lots_18mo,
            "acres_18mo":         acres_18mo,
            "dev_cost_per_lot":   dev_cost,
            "ff":                 ff,
            "full_sections":      full_secs,
            "last_lots":          last_lots,
            "last_section_acres": last_acres,
        })
    out["lot_rows"] = lot_rows

    total_lots = sum(r.get("total_lots", 0) for r in lot_rows)
    # total_dev_cost and total_lot_landscaping are accumulated in the section loop below

    out["total_lots"] = total_lots

    # Operating costs
    project_length_months = safe(inp.get("project_length_months", 60))  # estimated, refined below
    marketing_pct          = safe(inp.get("marketing_pct", 0.02))
    prof_svc_pct           = safe(inp.get("prof_svc_pct", 0.015))
    dmf_pct                = safe(inp.get("dmf_pct", 0.025))
    personnel_mo           = safe(inp.get("personnel_monthly", 0))
    marketing_personnel_mo = safe(inp.get("marketing_personnel_monthly", 0))
    legal_mo               = safe(inp.get("legal_monthly", 0))
    mud_mo                 = safe(inp.get("mud_monthly", 0))
    mud_pct_duration       = safe(inp.get("mud_pct", 0.2))   # what % of project MUD runs
    insurance_mo           = safe(inp.get("insurance_monthly", 0))
    bookkeeping_mo         = safe(inp.get("bookkeeping_monthly", 0))

    # ── 3. REVENUE INPUTS derived values ──────────────────────────────────────
    timing_method = inp.get("timing_method", "1 Takedown")
    take1_pct = 1.0 if timing_method == "1 Takedown" else 0.5
    take2_pct = 0.0 if timing_method == "1 Takedown" else (0.5 if timing_method == "50/50" else 0.25)
    take3_pct = 0.25 if timing_method == "50/25/25" else 0.0

    brokerage_fees    = safe(inp.get("brokerage_fees", 0.03))
    lot_closing_costs = safe(inp.get("lot_closing_costs", 0.015))
    bem_period        = int(safe(inp.get("bem_period", 9)))    # months before delivery BEM is received
    bem_pct           = safe(inp.get("bem_pct", 0.18))         # % of lot revenue received as BEM
    cost_per_mailbox  = safe(inp.get("cost_per_mailbox", 200)) # Cost Inputs B11

    # $/FF by year
    ppff_dict = inp.get("price_per_ff", {})
    ff_by_year = [safe(ppff_dict.get(str(y), ppff_dict.get(y, 1800))) for y in range(11)]

    # Residential pods — use res_pod_acreage + res_pod_count for per-pod area
    res_pods = inp.get("res_pods", [])
    res_pod_total_acres = safe(inp.get("res_pod_acreage") or inp.get("residential_pod_acres", 0))
    res_pod_count = max(int(safe(inp.get("res_pod_count", len([rp for rp in res_pods if safe(rp.get("price_per_acre"))]) or 1))), 1)
    if res_pod_total_acres == 0:
        res_pod_total_acres = res_pod_acres
    acres_per_res_pod = res_pod_total_acres / res_pod_count if res_pod_count else 0

    # Commercial pods
    comm_pods = inp.get("comm_pods", [])
    comm_pod_total_acres = safe(inp.get("comm_pod_acreage") or inp.get("commercial_pod_acres", 0))
    comm_pod_count = max(int(safe(inp.get("comm_pod_count", len([cp for cp in comm_pods if safe(cp.get("price_per_sf"))]) or 1))), 1)
    if comm_pod_total_acres == 0:
        comm_pod_total_acres = comm_pod_acres
    acres_per_comm_pod = comm_pod_total_acres / comm_pod_count if comm_pod_count else 0

    # MUD / WCID bonds
    mud_row   = inp.get("mud_bond", {})
    wcid_row  = inp.get("wcid_bond", {})

    # lot_av / comm_av / bond revenues are computed after cashflow engine builds
    # lot_count_by_month — see "AV BUILDUP" section below

    # ── 4. CASHFLOW ENGINE ────────────────────────────────────────────────────
    MAX_MONTHS = 360

    # Determine project timeline
    all_deliveries = []
    for r in plant_cost_rows:
        if r["type"] not in ("None",""):
            all_deliveries.append(int(r["start_month"]) + int(r["duration"]) - 1)
    for r in road_cost_rows:
        all_deliveries.append(int(r["delivery_period"]))
    if lot_rows:
        max_lot_months = 0
        for lr in lot_rows:
            if lr.get("total_lots", 0) > 0:
                pace       = safe(lr.get("pace", 0))
                sm         = int(safe(lr.get("dev_start_month", default_start_month))) + 1
                build_time = max(0, int(safe(lr.get("build_time", 12))))
                full_secs  = lr.get("full_sections", 0)
                last_lots  = lr.get("last_lots", 0)
                if pace > 0:
                    sec_count = full_secs + (1 if last_lots > 0 else 0)
                    last_t3   = sm + sec_count * 18 + 9
                    # Home sales continue after last delivery + build_time
                    last_home_sale = sm + sec_count * 18 + build_time + max(1, int(round(pace * 18 / pace)))
                    max_lot_months = max(max_lot_months, last_t3, last_home_sale)
        all_deliveries.append(max_lot_months)

    proj_months = min(max(all_deliveries + [60]), MAX_MONTHS) if all_deliveries else 60
    project_length_months = proj_months
    out["project_length_months"] = proj_months

    # Monthly revenue/cost arrays
    rev_monthly = [0.0] * (MAX_MONTHS + 1)
    cost_monthly = [0.0] * (MAX_MONTHS + 1)

    # Land cost: takedowns (input period is 0-indexed; app month = period + 1)
    for td in td_rows:
        m = int(td["period"]) + 1
        if 1 <= m <= MAX_MONTHS:
            cost_monthly[m] += td["total"]

    # Infrastructure costs: spread evenly over duration
    def spread_cost(cost_array, amount, start_m, duration):
        if duration <= 0 or amount <= 0:
            return
        per_month = amount / duration
        for m in range(int(start_m), int(start_m) + int(duration)):
            if 1 <= m <= MAX_MONTHS:
                cost_array[m] += per_month

    for r in plant_cost_rows:
        spread_cost(cost_monthly, r["total_cost"], r["start_month"], r["duration"])
        spread_cost(cost_monthly, r["ph2_total_cost"], r["ph2_start_month"], r["ph2_duration"])
    for r in amenity_cost_rows:
        spread_cost(cost_monthly, r["total_cost"], r["start_month"], r["duration"])
    for r in det_cost_rows:
        spread_cost(cost_monthly, r["total_cost"], r["start_month"], r["duration"])
        # Detention landscaping: lump sum at delivery period (Excel Calc_Costs row 717)
        dp = int(r["delivery_period"])
        if 1 <= dp <= MAX_MONTHS and r.get("total_landscaping", 0) > 0:
            cost_monthly[dp] += r["total_landscaping"]
    for r in other_cost_rows:
        spread_cost(cost_monthly, r["total_cost"], r["start_month"], max(r["duration"], 1))
    for r in road_cost_rows:
        spread_cost(cost_monthly, r["total_cost"], r["start_month"], r["duration"])
        # Road landscaping: lump sum at delivery period (Excel Calc_Costs row 718)
        dp = int(r["delivery_period"])
        if 1 <= dp <= MAX_MONTHS and r["total_landscaping"] > 0:
            cost_monthly[dp] += r["total_landscaping"]

    # ── LOT SECTION SCHEDULE (matches Excel Calc_Costs rows 57-679) ─────────────
    # Section k of lot type i:
    #   start_month  = dev_start + (k-1)*18
    #   delivery_m   = start_month + 18
    #   lots         = lots_18mo (full sections) or last_lots (partial)
    #   dev_cost     = lots * dev_cost_per_lot * (1 + sectional_other_pct)  [Excel F col]
    #   Phase 1 cost = dev_cost * 0.10 / 12  per month, months 0-11 of section
    #   Phase 2 cost = dev_cost * 0.90 / 6   per month, months 12-17 of section
    #   landscaping  = lots * ls_per_lot * (1 + landscaping_other_pct)  [Excel I col]
    #   lot delivery tracked at delivery_m  [lot_count_by_month]
    lot_rev_by_month       = [0.0] * (MAX_MONTHS + 1)
    lot_cost_by_month      = [0.0] * (MAX_MONTHS + 1)
    lot_count_by_month     = [0.0] * (MAX_MONTHS + 1)
    lot_landscaping_by_month = [0.0] * (MAX_MONTHS + 1)
    fencing_by_month       = [0.0] * (MAX_MONTHS + 1)   # Fencing cost, delivery-timed
    urd_by_month           = [0.0] * (MAX_MONTHS + 1)   # Dry utilities (URD), delivery-timed
    lot_streetlight_by_month = [0.0] * (MAX_MONTHS + 1) # Lot-level streetlights, delivery-timed
    total_dev_cost         = 0.0
    total_lot_landscaping  = 0.0
    total_fencing_cost     = 0.0
    total_urd_cost         = 0.0
    total_lot_streetlight_cost = 0.0

    for lr in lot_rows:
        if not safe(lr.get("on", 0)) or lr.get("total_lots", 0) == 0:
            continue
        pace = safe(lr.get("pace", 0))
        if pace <= 0:
            continue
        sm              = int(safe(lr.get("dev_start_month", default_start_month))) + 1
        lots_18mo       = pace * 18                   # float (e.g. 13.5 for 80FF)
        full_secs       = lr.get("full_sections", 0)
        last_lots       = lr.get("last_lots", 0)
        dev_cost_per_lot = lr.get("dev_cost_per_lot", 0)
        ls_per_lot       = safe(lr.get("landscaping_per_lot", 0))
        ff_lr            = lr.get("ff", 0)

        # Separate delivery-timed costs per lot (NOT in dev_cost, NOT phase-scheduled)
        fence_cost_ff  = safe(lr.get("fence_cost_per_ff", 0))
        urd_per_lot    = safe(lr.get("urd_per_lot", 0))
        lots_per_sl    = safe(lr.get("lots_per_streetlight", 0))  # "Lots per Street Light" spacing

        # Build section list: full sections + optional partial section
        sections = [(k, lots_18mo) for k in range(1, full_secs + 1)]
        if last_lots > 0:
            sections.append((full_secs + 1, float(last_lots)))

        for k, section_lots in sections:
            section_start = sm + (k - 1) * 18
            delivery_m    = section_start + 18

            # Dev cost with sectional contingency (Excel: lots * K_col * (1+B6))
            # K = FF*(WSD+Paving) only — no URD, no fence
            section_cost  = section_lots * dev_cost_per_lot * (1 + sectional_other_pct)
            total_dev_cost += section_cost

            # Phase 1: months 0-11 relative to section_start
            ph1_per_month = section_cost * 0.10 / 12
            for mo in range(12):
                m = section_start + mo
                if 1 <= m <= MAX_MONTHS:
                    lot_cost_by_month[m] += ph1_per_month

            # Phase 2: months 12-17 relative to section_start
            ph2_per_month = section_cost * 0.90 / 6
            for mo in range(12, 18):
                m = section_start + mo
                if 1 <= m <= MAX_MONTHS:
                    lot_cost_by_month[m] += ph2_per_month

            # Landscaping lump sum at delivery with landscaping contingency (Excel I col)
            if ls_per_lot:
                ls_amount = section_lots * ls_per_lot * (1 + landscaping_other_pct)
                total_lot_landscaping += ls_amount
                if 1 <= delivery_m <= MAX_MONTHS:
                    lot_landscaping_by_month[delivery_m] += ls_amount

            # Fencing cost: lump sum at delivery (Excel Calc_Revenues rows 244-479 / Cashflow B14)
            # = fence_cost_per_FF * lots * FF * fenced_pct
            if fence_cost_ff and ff_lr and fenced_pct:
                fence_amt = section_lots * fence_cost_ff * ff_lr * fenced_pct
                total_fencing_cost += fence_amt
                if 1 <= delivery_m <= MAX_MONTHS:
                    fencing_by_month[delivery_m] += fence_amt

            # URD (dry utilities): lump sum at delivery (Excel Calc_Costs rows 988-1003)
            # = lots * URD_per_lot
            if urd_per_lot:
                urd_amt = section_lots * urd_per_lot
                total_urd_cost += urd_amt
                if 1 <= delivery_m <= MAX_MONTHS:
                    urd_by_month[delivery_m] += urd_amt

            # Lot-level streetlights: lump sum at delivery (Excel Calc_Costs rows 988-1003)
            # = lots / lots_per_streetlight * cost_per_streetlight
            if lots_per_sl and lots_per_sl > 0:
                sl_amt = (section_lots / lots_per_sl) * cost_per_streetlight
                total_lot_streetlight_cost += sl_amt
                if 1 <= delivery_m <= MAX_MONTHS:
                    lot_streetlight_by_month[delivery_m] += sl_amt

            # Lot delivery tracking (lump sum at delivery month)
            if 1 <= delivery_m <= MAX_MONTHS:
                lot_count_by_month[delivery_m] += section_lots

    # Lot revenue: section lump-sum delivery matching Excel Calc_Revenues
    # Each section delivers ALL its lots as a lump sum at section_delivery (sm+(k+1)*18).
    # T1 = section_delivery, T2 = T1+6, T3 = T1+9.
    # This matches Excel Revenue Inputs cols AG/AH/AI and Calc_Revenues rows 723-739.
    lot_brokerage_by_month = [0.0] * (MAX_MONTHS + 1)
    lot_closing_by_month   = [0.0] * (MAX_MONTHS + 1)
    lot_tax_by_month       = [0.0] * (MAX_MONTHS + 1)
    lot_mailbox_by_month   = [0.0] * (MAX_MONTHS + 1)
    # Revenue sub-category accumulators (for Project Performance display)
    total_premium_rev     = 0.0
    total_escalation_rev  = 0.0
    total_fence_fee_rev   = 0.0
    total_mktg_fee_rev    = 0.0
    total_lot_base_rev    = 0.0   # base lot sale revenues (BEM + T1/T2/T3 of gross)

    for lr in lot_rows:
        if not safe(lr.get("on", 0)) or lr.get("total_lots", 0) == 0:
            continue
        pace = safe(lr.get("pace", 0))
        if pace <= 0:
            continue
        total        = int(lr["total_lots"])
        sm           = int(safe(lr.get("dev_start_month", default_start_month))) + 1
        ff           = lr.get("ff", 0)
        premium_pff  = safe(lr.get("premium_per_ff", 0))
        escalation   = safe(lr.get("escalation", 0))
        fence_fee_pff= safe(lr.get("fence_per_ff", 0))
        mktg_fee_lot = safe(lr.get("marketing_fee", 0))
        lot_av_pct_t = safe(lr.get("lot_av_pct", 0.5))
        lot_tax_rate = safe(lr.get("lot_tax_rate", 0.022))
        # lot_tax_per_lot uses base $/FF (Year 0), not year-specific — matches Excel N27 = $B$13*A27*K27*M27

        # Use exact section structure (matches dev cost loop above)
        lots_18mo_r  = pace * 18            # float
        full_secs_r  = lr.get("full_sections", 0)
        last_lots_r  = lr.get("last_lots", 0)
        rev_sections = [(k, lots_18mo_r) for k in range(1, full_secs_r + 1)]
        if last_lots_r > 0:
            rev_sections.append((full_secs_r + 1, float(last_lots_r)))

        for k, batch in rev_sections:
            t1_m  = sm + k * 18              # T1 = section delivery month
            t2_m  = t1_m + 6
            t3_m  = t1_m + 9

            year_idx = min(max(int((t1_m - 1) / 12), 0), 10)
            ff_rate = ff_by_year[year_idx] if year_idx < len(ff_by_year) else ff_by_year[-1]
            # Lot tax uses base $/FF (Year 0) per Excel N27 = $B$13 * A27 * K27 * M27
            lot_tax_per_lot = ff_by_year[0] * ff * lot_av_pct_t * lot_tax_rate
            gross_lot_rev = batch * ff * ff_rate

            # BEM: received bem_period months before T1
            bem_amount = gross_lot_rev * bem_pct
            bem_m = max(1, t1_m - bem_period)
            if bem_m <= MAX_MONTHS:
                lot_rev_by_month[bem_m] += bem_amount
                total_lot_base_rev += bem_amount

            # T1/T2/T3 revenues (gross minus BEM, split by take_pcts)
            gross_remainder = gross_lot_rev * (1 - bem_pct)
            if t1_m <= MAX_MONTHS:
                lot_rev_by_month[t1_m] += gross_remainder * take1_pct
                total_lot_base_rev += gross_remainder * take1_pct
            if t2_m <= MAX_MONTHS:
                lot_rev_by_month[t2_m] += gross_remainder * take2_pct
                total_lot_base_rev += gross_remainder * take2_pct
            if t3_m <= MAX_MONTHS:
                lot_rev_by_month[t3_m] += gross_remainder * take3_pct
                total_lot_base_rev += gross_remainder * take3_pct

            # Brokerage fees (cost)
            if brokerage_fees:
                if t1_m <= MAX_MONTHS:
                    lot_brokerage_by_month[t1_m] += gross_lot_rev * take1_pct * brokerage_fees
                if t2_m <= MAX_MONTHS:
                    lot_brokerage_by_month[t2_m] += gross_lot_rev * take2_pct * brokerage_fees
                if t3_m <= MAX_MONTHS:
                    lot_brokerage_by_month[t3_m] += gross_lot_rev * take3_pct * brokerage_fees

            # Lot closing costs (cost)
            if lot_closing_costs:
                if t1_m <= MAX_MONTHS:
                    lot_closing_by_month[t1_m] += gross_lot_rev * take1_pct * lot_closing_costs
                if t2_m <= MAX_MONTHS:
                    lot_closing_by_month[t2_m] += gross_lot_rev * take2_pct * lot_closing_costs
                if t3_m <= MAX_MONTHS:
                    lot_closing_by_month[t3_m] += gross_lot_rev * take3_pct * lot_closing_costs

            # Lot taxes (cost)
            if lot_tax_per_lot:
                if t1_m <= MAX_MONTHS:
                    lot_tax_by_month[t1_m] += batch * take1_pct * lot_tax_per_lot
                if t2_m <= MAX_MONTHS:
                    lot_tax_by_month[t2_m] += batch * take2_pct * lot_tax_per_lot
                if t3_m <= MAX_MONTHS:
                    lot_tax_by_month[t3_m] += batch * take3_pct * lot_tax_per_lot

            # Mailboxes (cost)
            if cost_per_mailbox:
                if t1_m <= MAX_MONTHS:
                    lot_mailbox_by_month[t1_m] += batch * take1_pct * cost_per_mailbox
                if t2_m <= MAX_MONTHS:
                    lot_mailbox_by_month[t2_m] += batch * take2_pct * cost_per_mailbox
                if t3_m <= MAX_MONTHS:
                    lot_mailbox_by_month[t3_m] += batch * take3_pct * cost_per_mailbox

            # Premiums (revenue)
            if premium_pff and ff:
                prem_total = batch * ff * premium_pff
                if t1_m <= MAX_MONTHS:
                    lot_rev_by_month[t1_m] += prem_total * take1_pct
                    total_premium_rev += prem_total * take1_pct
                if t2_m <= MAX_MONTHS:
                    lot_rev_by_month[t2_m] += prem_total * take2_pct
                    total_premium_rev += prem_total * take2_pct
                if t3_m <= MAX_MONTHS:
                    lot_rev_by_month[t3_m] += prem_total * take3_pct
                    total_premium_rev += prem_total * take3_pct

            # Escalation (revenue, at T2 and T3 relative to T1)
            if escalation:
                escal_t2 = (escalation / 12) * 6 * (gross_remainder * take2_pct)
                escal_t3 = (escalation / 12) * 9 * (gross_remainder * take3_pct)
                if escal_t2 > 0 and t2_m <= MAX_MONTHS:
                    lot_rev_by_month[t2_m] += escal_t2
                    total_escalation_rev += escal_t2
                if escal_t3 > 0 and t3_m <= MAX_MONTHS:
                    lot_rev_by_month[t3_m] += escal_t3
                    total_escalation_rev += escal_t3

            # Fence fees (revenue) — Excel: fence_rev_per_FF * lots * FF * fenced_pct
            if fence_fee_pff and ff and fenced_pct:
                fence_rev = batch * ff * fence_fee_pff * fenced_pct
                if t1_m <= MAX_MONTHS:
                    lot_rev_by_month[t1_m] += fence_rev * take1_pct
                    total_fence_fee_rev += fence_rev * take1_pct
                if t2_m <= MAX_MONTHS:
                    lot_rev_by_month[t2_m] += fence_rev * take2_pct
                    total_fence_fee_rev += fence_rev * take2_pct
                if t3_m <= MAX_MONTHS:
                    lot_rev_by_month[t3_m] += fence_rev * take3_pct
                    total_fence_fee_rev += fence_rev * take3_pct

            # Marketing fees (revenue)
            if mktg_fee_lot:
                if t1_m <= MAX_MONTHS:
                    lot_rev_by_month[t1_m] += batch * take1_pct * mktg_fee_lot
                    total_mktg_fee_rev += batch * take1_pct * mktg_fee_lot
                if t2_m <= MAX_MONTHS:
                    lot_rev_by_month[t2_m] += batch * take2_pct * mktg_fee_lot
                    total_mktg_fee_rev += batch * take2_pct * mktg_fee_lot
                if t3_m <= MAX_MONTHS:
                    lot_rev_by_month[t3_m] += batch * take3_pct * mktg_fee_lot
                    total_mktg_fee_rev += batch * take3_pct * mktg_fee_lot

    for m in range(1, MAX_MONTHS + 1):
        rev_monthly[m]  += lot_rev_by_month[m]
        cost_monthly[m] += lot_cost_by_month[m]
        cost_monthly[m] += lot_landscaping_by_month[m]
        cost_monthly[m] += fencing_by_month[m]
        cost_monthly[m] += urd_by_month[m]
        cost_monthly[m] += lot_streetlight_by_month[m]
        cost_monthly[m] += lot_brokerage_by_month[m]
        cost_monthly[m] += lot_closing_by_month[m]
        cost_monthly[m] += lot_tax_by_month[m]
        cost_monthly[m] += lot_mailbox_by_month[m]

    # Residential pod revenues — each pod uses its own sale_period (Excel K46, K47, ...)
    res_pod_revenue = 0.0
    for i, rp in enumerate(res_pods):
        if i >= res_pod_count:
            break
        ppa_pod = safe(rp.get("price_per_acre"))
        cc = safe(rp.get("closing_costs_pct", 0.045))
        impact_fee = safe(rp.get("impact_fee_per_lot", 0))
        lots_per_acre = safe(rp.get("implied_lots_per_acre", 0))
        sale_period = int(safe(rp.get("sale_period", proj_months)))
        # Excel J46: price_per_acre*(1-cc)*acres_per_pod + impact_fee*lots_per_acre*acres_per_pod
        pod_rev = acres_per_res_pod * ppa_pod * (1 - cc) + impact_fee * lots_per_acre * acres_per_res_pod
        res_pod_revenue += pod_rev
        if pod_rev > 0 and 1 <= sale_period <= MAX_MONTHS:
            rev_monthly[sale_period] += pod_rev

    # Commercial pod revenues — each pod uses its own sale_period (Excel I55, I56, ...)
    comm_pod_revenue = 0.0
    for i, cp in enumerate(comm_pods):
        if i >= comm_pod_count:
            break
        psf = safe(cp.get("price_per_sf"))
        cc = safe(cp.get("closing_costs_pct", 0.045))
        sale_period = int(safe(cp.get("sale_period", proj_months)))
        # Excel H55: price_per_sf * 43560 * (1-cc) * acres_per_pod
        pod_rev = acres_per_comm_pod * 43560 * psf * (1 - cc)
        comm_pod_revenue += pod_rev
        if pod_rev > 0 and 1 <= sale_period <= MAX_MONTHS:
            rev_monthly[sale_period] += pod_rev

    # ── AV BUILDUP (Excel-precise: section lump-sums + build_time + inventory) ─
    # Matches Calc_Revenues rows 742-838:
    #   1. Each section delivers all its lots as a lump sum at section_start+18
    #   2. Homes complete build_time months after lot delivery (row 742)
    #   3. Homes sell from completed inventory at pace/month (rows 763-798)
    #   4. Cumulative AV = running sum of homes_sold * av_per_unit (rows 803+823)
    _last_home_sale_m = 0
    av_by_month = [0.0] * (MAX_MONTHS + 2)
    home_sales_by_month = [0.0] * (MAX_MONTHS + 2)
    for lr in lot_rows:
        if not safe(lr.get("on", 0)) or lr.get("total_lots", 0) == 0:
            continue
        pace_lr    = safe(lr.get("pace", 0))
        if pace_lr <= 0:
            continue
        total_lr   = int(lr["total_lots"])
        sm_lr      = int(safe(lr.get("dev_start_month", default_start_month))) + 1
        build_time = max(0, int(safe(lr.get("build_time", 12))))
        av_pct_lr  = safe(lr.get("av_pct", 0.85))
        hp         = safe(lr.get("home_price", 0))
        av_per_lot = hp * av_pct_lr
        if av_per_lot <= 0:
            continue

        # Section-based home completion schedule — use same section structure as dev cost loop
        completions  = [0.0] * (MAX_MONTHS + 2)
        lots_18mo_av = pace_lr * 18              # float
        full_secs_av = lr.get("full_sections", 0)
        last_lots_av = lr.get("last_lots", 0)
        av_sections  = [(k, lots_18mo_av) for k in range(1, full_secs_av + 1)]
        if last_lots_av > 0:
            av_sections.append((full_secs_av + 1, float(last_lots_av)))

        for k, lots_this in av_sections:
            section_delivery = sm_lr + k * 18
            # Excel splits lot deliveries into takes (T1/T2/T3), homes complete
            # build_time after each take — matching Calc_Revenues rows 742+
            takes = [(section_delivery, take1_pct),
                     (section_delivery + 6, take2_pct),
                     (section_delivery + 9, take3_pct)]
            for take_m, take_frac in takes:
                if take_frac > 0:
                    comp_m = take_m + build_time
                    if 1 <= comp_m <= MAX_MONTHS:
                        completions[comp_m] += lots_this * take_frac

        # Inventory-based home sales: pace-limited, uses PRIOR period inventory (Excel row 783)
        # Excel formula: sales[P] = MIN(inventory[P-1], pace), where inventory = cum_comp - cum_sales
        cum_comp = 0.0
        cum_sold = 0.0
        prev_inventory = 0.0
        for m in range(1, MAX_MONTHS + 1):
            # Sales based on previous period's ending inventory (Excel: EL770 for period EM)
            sold_this  = min(prev_inventory, pace_lr)
            cum_sold  += sold_this
            av_by_month[m] += sold_this * av_per_lot
            home_sales_by_month[m] += sold_this
            if sold_this > 0:
                _last_home_sale_m = m
            # Update inventory after this period's completions
            cum_comp  += completions[m]
            prev_inventory = cum_comp - cum_sold

    lot_av = sum(av_by_month[1:MAX_MONTHS + 1])

    comm_av = 0.0
    for i, cp in enumerate(comm_pods):
        if i >= comm_pod_count:
            break
        av_per_acre = safe(cp.get("av_per_acre", 0))
        sale_period = int(safe(cp.get("sale_period", proj_months)))
        av_delay    = int(safe(cp.get("av_delay_months") or cp.get("av_delay", 18)))  # Excel K55=18 months after sale
        pod_av = acres_per_comm_pod * av_per_acre
        comm_av += pod_av
        av_period = sale_period + av_delay  # Excel L55 = I55 + K55
        if pod_av > 0 and 1 <= av_period <= MAX_MONTHS:
            av_by_month[av_period] += pod_av

    # Cumulative AV array (row 840 analogue: SUM of all prior months)
    cum_av_monthly = [0.0] * (MAX_MONTHS + 2)
    for m in range(1, MAX_MONTHS + 1):
        cum_av_monthly[m] = cum_av_monthly[m - 1] + av_by_month[m]

    def _compute_bond_issuances_excel(bond_cfg):
        """
        Excel-precise bond issuance (Calc_Revenues rows 866/864/865):
          bond_to_date(M) = cum_av_monthly[M-3] * debt_ratio   [at bond periods]
          bond_this(M)    = bond_to_date(M) - bond_to_date(M - bond_interval)
          proceeds(M)     = max(bond_this(M) * pct_to_dev, 0)
        Returns list of (month, proceeds) tuples.
        """
        if not bond_cfg:
            return []
        toggle = int(safe(bond_cfg.get("toggle", 1)))
        if not toggle:
            return []
        debt_ratio_b    = safe(bond_cfg.get("debt_ratio", 0.12))
        pct_to_dev_b    = safe(bond_cfg.get("pct_to_dev", 0.85))
        first_period    = int(safe(bond_cfg.get("first_bond_period", 0)))
        bond_interval_b = int(safe(bond_cfg.get("bond_interval", 12)))
        if first_period <= 0 or debt_ratio_b <= 0:
            return []
        result = []
        p = first_period
        prev_bond_to_date = 0.0
        while p <= MAX_MONTHS:
            av_lag = max(0, p - 3)
            bond_to_date = cum_av_monthly[av_lag] * debt_ratio_b
            bond_this = max(0.0, bond_to_date - prev_bond_to_date)
            proceeds = bond_this * pct_to_dev_b
            if proceeds > 0:
                result.append((p, proceeds))
            prev_bond_to_date = bond_to_date
            if bond_interval_b <= 0:
                break
            p += bond_interval_b
        return result

    # MUD/WCID bond revenues — Excel cumulative-AV method
    mud_issuances  = _compute_bond_issuances_excel(mud_row)
    wcid_issuances = _compute_bond_issuances_excel(wcid_row)
    mud_recv_fee_by_month  = [0.0] * (MAX_MONTHS + 1)
    wcid_recv_fee_by_month = [0.0] * (MAX_MONTHS + 1)
    mud_recv_fee_pct  = safe(mud_row.get("receivables_fee") or mud_row.get("receivables_fee_pct", 0.025)) if mud_row else 0
    wcid_recv_fee_pct = safe(wcid_row.get("receivables_fee") or wcid_row.get("receivables_fee_pct", 0.025)) if wcid_row else 0
    for bp, amt in mud_issuances:
        rev_monthly[bp] += amt
        # Receivables fee is a cost (Excel Calc_Revenues row 868)
        fee = amt * mud_recv_fee_pct
        mud_recv_fee_by_month[bp] += fee
        cost_monthly[bp] += fee
    for bp, amt in wcid_issuances:
        rev_monthly[bp] += amt
        fee = amt * wcid_recv_fee_pct
        wcid_recv_fee_by_month[bp] += fee
        cost_monthly[bp] += fee

    # Operating costs — Excel-precise end periods
    # Excel costs run from period 0 to period N inclusive = N+1 total months.
    # Our month 1 = Excel period 0, so to run N+1 months we need end = N+1.
    # D95 = last lot delivery period (0-indexed). Costs run D95+1 months → our end = D95+1.
    # App month M = Excel period M-1.  After dev_start_month +1 fix,
    # last_lot_rev_period already maps to Excel D95+1 total months.
    total_lot_revenue_gross = sum(lot_rev_by_month)
    last_lot_rev_period  = max((m for m in range(1, MAX_MONTHS + 1) if lot_rev_by_month[m] > 0), default=proj_months)
    last_delivery_period = last_lot_rev_period
    last_home_period     = _last_home_sale_m if _last_home_sale_m > 0 else last_delivery_period

    # Marketing cost = sum of per-lot marketing fees (Excel matches rev_mktg_fees)
    marketing_total  = total_mktg_fee_rev
    # Prof services = prof_svc_pct * total_revenue (includes pods+bonds, not just lot rev)
    total_revenue_pre = sum(rev_monthly[1:proj_months+1])
    prof_svc_total   = total_revenue_pre * prof_svc_pct

    total_det_landscaping = sum(r.get("total_landscaping", 0) for r in det_cost_rows)

    # Site work: per-section cost at BEM month (Excel Cashflow B15)
    # Excel: SUMIF(Revenue_Inputs!K71:K306, month, Revenue_Inputs!O71:O306)
    # K = BEM month (T1 - BEM_period), O = gross_section_revenue * site_work_pct
    site_work_by_month = [0.0] * (MAX_MONTHS + 1)
    site_work_total = 0.0
    for lr in lot_rows:
        if not safe(lr.get("on", 0)) or lr.get("total_lots", 0) == 0:
            continue
        pace_sw = safe(lr.get("pace", 0))
        if pace_sw <= 0:
            continue
        sm_sw = int(safe(lr.get("dev_start_month", default_start_month))) + 1
        ff_sw = lr.get("ff", 0)
        lots_18mo_sw = pace_sw * 18
        full_secs_sw = lr.get("full_sections", 0)
        last_lots_sw = lr.get("last_lots", 0)
        sw_sections = [(k, lots_18mo_sw) for k in range(1, full_secs_sw + 1)]
        if last_lots_sw > 0:
            sw_sections.append((full_secs_sw + 1, float(last_lots_sw)))
        for k, batch_sw in sw_sections:
            t1_sw = sm_sw + k * 18
            yr_sw = min(max(int((t1_sw - 1) / 12), 0), 10)
            ff_rate_sw = ff_by_year[yr_sw] if yr_sw < len(ff_by_year) else ff_by_year[-1]
            gross_sw = batch_sw * ff_sw * ff_rate_sw
            sw_amt = gross_sw * site_work_pct
            site_work_total += sw_amt
            bem_m_sw = max(1, t1_sw - bem_period)
            if 1 <= bem_m_sw <= MAX_MONTHS:
                site_work_by_month[bem_m_sw] += sw_amt
    for m in range(1, MAX_MONTHS + 1):
        cost_monthly[m] += site_work_by_month[m]

    # Marketing: C91 = total / (B91+1) months. B91 = last_home_period-1 (0-indexed).
    # Runs last_home_period months (months 1..last_home_period in Python).
    mkt_dur   = max(1, last_home_period)
    mkt_per_m = marketing_total / mkt_dur
    spread_cost(cost_monthly, marketing_total, 1, mkt_dur)

    # Prof Services: C95 = total / (D95+1) months. D95 = last_delivery_period-1 (0-indexed).
    # Runs last_delivery_period months (months 1..last_delivery_period in Python).
    ps_dur    = max(1, last_delivery_period)
    ps_per_m  = prof_svc_total / ps_dur
    spread_cost(cost_monthly, prof_svc_total, 1, ps_dur)

    # General Personnel: Excel D103=D95 (0-indexed). Runs (D95+1)=last_delivery_period months.
    gen_pers_end = max(1, last_delivery_period)
    for m in range(1, gen_pers_end + 1):
        if m <= MAX_MONTHS:
            cost_monthly[m] += personnel_mo

    # Marketing Personnel: Excel D104=B91 (0-indexed). Runs last_home_period months.
    mkt_pers_end = max(1, last_home_period)
    for m in range(1, mkt_pers_end + 1):
        if m <= MAX_MONTHS:
            cost_monthly[m] += marketing_personnel_mo

    # Legal: Excel D108=D103=D95 (0-indexed). Runs last_delivery_period months.
    legal_end = max(1, last_delivery_period)
    for m in range(1, legal_end + 1):
        if m <= MAX_MONTHS:
            cost_monthly[m] += legal_mo

    # Insurance: Excel D116=D120=D104=B91 (0-indexed). Runs last_home_period months.
    ins_end = max(1, last_home_period)
    for m in range(1, ins_end + 1):
        if m <= MAX_MONTHS:
            cost_monthly[m] += insurance_mo

    # Bookkeeping: Excel D120=D104=B91 (0-indexed). Runs last_home_period months.
    bk_end = max(1, last_home_period)
    for m in range(1, bk_end + 1):
        if m <= MAX_MONTHS:
            cost_monthly[m] += bookkeeping_mo

    # MUD & HOA Advances: Excel E112 = MROUND(D108 * D112, 1) where D108 is 0-indexed.
    # D108 = last_delivery_period - 1 (0-indexed). Runs (E112+1) months.
    mud_end_0idx = int(mround((last_delivery_period - 1) * mud_pct_duration, 1)) if mud_pct_duration > 0 else 0
    mud_run_months = max(1, mud_end_0idx + 1)
    for m in range(1, mud_run_months + 1):
        if m <= MAX_MONTHS:
            cost_monthly[m] += mud_mo
    mud_total = mud_mo * mud_run_months

    # Road-level streetlights: at road delivery period (Excel Cashflow B17, SUMIF term)
    road_streetlight_by_month = [0.0] * (MAX_MONTHS + 1)
    road_streetlight_total = 0.0
    for r in road_cost_rows:
        sl_cost = r["total_lights"] * cost_per_streetlight if r.get("total_lights") else 0
        if sl_cost > 0:
            road_streetlight_total += sl_cost
            dp = int(r["delivery_period"])
            if 1 <= dp <= MAX_MONTHS:
                road_streetlight_by_month[dp] += sl_cost
    for m in range(1, MAX_MONTHS + 1):
        cost_monthly[m] += road_streetlight_by_month[m]

    # Total streetlight cost = road-level + lot-level
    streetlight_total = road_streetlight_total + total_lot_streetlight_cost

    # DMF: monthly proportional — Excel Cashflow B18 = 2.5% * SUM(that month's other costs)
    # Included in DMF base: Plants + Amenities + Detention + Other + Roads + Fencing + SiteWork
    #   + Landscaping + MUD&HOA + Insurance + Legal + Taxes + ProfSvc + SectionDev
    # We need category-level monthly totals. Build them here for DMF + cashflow detail.
    # Most are already tracked. For operating costs we need per-month arrays:
    op_mud_m      = [0.0] * (MAX_MONTHS + 1)
    op_insurance_m = [0.0] * (MAX_MONTHS + 1)
    op_legal_m     = [0.0] * (MAX_MONTHS + 1)
    op_prof_svc_m  = [0.0] * (MAX_MONTHS + 1)
    op_mkt_exp_m   = [0.0] * (MAX_MONTHS + 1)
    for m in range(1, mud_run_months + 1):
        if m <= MAX_MONTHS: op_mud_m[m] = mud_mo
    for m in range(1, ins_end + 1):
        if m <= MAX_MONTHS: op_insurance_m[m] = insurance_mo
    for m in range(1, legal_end + 1):
        if m <= MAX_MONTHS: op_legal_m[m] = legal_mo
    if ps_dur > 0:
        for m in range(1, ps_dur + 1):
            if m <= MAX_MONTHS: op_prof_svc_m[m] = ps_per_m
    if mkt_dur > 0:
        for m in range(1, mkt_dur + 1):
            if m <= MAX_MONTHS: op_mkt_exp_m[m] = mkt_per_m

    # Reconstruct monthly infrastructure arrays (same as cashflow detail, but needed now for DMF)
    _Z = lambda: [0.0] * (MAX_MONTHS + 1)
    _cc_plants = _Z(); _cc_amen = _Z(); _cc_det = _Z(); _cc_other = _Z(); _cc_roads = _Z()
    for r in plant_cost_rows:
        if r["total_cost"] > 0:
            per = r["total_cost"] / r["duration"] if r["duration"] > 0 else 0
            for mm in range(int(r["start_month"]), int(r["start_month"]) + int(r["duration"])):
                if 1 <= mm <= MAX_MONTHS: _cc_plants[mm] += per
        if r["ph2_total_cost"] > 0 and r["ph2_duration"] > 0:
            per2 = r["ph2_total_cost"] / r["ph2_duration"]
            for mm in range(int(r["ph2_start_month"]), int(r["ph2_start_month"]) + int(r["ph2_duration"])):
                if 1 <= mm <= MAX_MONTHS: _cc_plants[mm] += per2
    for r in amenity_cost_rows:
        if r["total_cost"] > 0 and r["duration"] > 0:
            per = r["total_cost"] / r["duration"]
            for mm in range(int(r["start_month"]), int(r["start_month"]) + int(r["duration"])):
                if 1 <= mm <= MAX_MONTHS: _cc_amen[mm] += per
    for r in det_cost_rows:
        if r["total_cost"] > 0 and r["duration"] > 0:
            per = r["total_cost"] / r["duration"]
            for mm in range(int(r["start_month"]), int(r["start_month"]) + int(r["duration"])):
                if 1 <= mm <= MAX_MONTHS: _cc_det[mm] += per
    for r in other_cost_rows:
        dur_r = max(r["duration"], 1)
        if r["total_cost"] > 0:
            per = r["total_cost"] / dur_r
            for mm in range(int(r["start_month"]), int(r["start_month"]) + int(dur_r)):
                if 1 <= mm <= MAX_MONTHS: _cc_other[mm] += per
    for r in road_cost_rows:
        if r["total_cost"] > 0 and r["duration"] > 0:
            per = r["total_cost"] / r["duration"]
            for mm in range(int(r["start_month"]), int(r["start_month"]) + int(r["duration"])):
                if 1 <= mm <= MAX_MONTHS: _cc_roads[mm] += per

    # All landscaping monthly = det landscaping + road landscaping + sectional landscaping
    _cc_landscape_m = _Z()
    for r in det_cost_rows:
        dp = int(r["delivery_period"])
        if 1 <= dp <= MAX_MONTHS and r.get("total_landscaping", 0) > 0:
            _cc_landscape_m[dp] += r["total_landscaping"]
    for r in road_cost_rows:
        dp = int(r["delivery_period"])
        if 1 <= dp <= MAX_MONTHS and r["total_landscaping"] > 0:
            _cc_landscape_m[dp] += r["total_landscaping"]
    for m in range(1, MAX_MONTHS + 1):
        _cc_landscape_m[m] += lot_landscaping_by_month[m]

    dmf_by_month = _Z()
    dmf_total = 0.0
    for m in range(1, MAX_MONTHS + 1):
        # DMF base = plants + amenities + det + other + roads + fencing + site_work
        #          + landscaping + MUD&HOA + insurance + legal + taxes + prof_svc + section_dev
        dmf_base_m = (_cc_plants[m] + _cc_amen[m] + _cc_det[m] + _cc_other[m] + _cc_roads[m]
                      + fencing_by_month[m] + site_work_by_month[m] + _cc_landscape_m[m]
                      + op_mud_m[m] + op_insurance_m[m] + op_legal_m[m]
                      + lot_tax_by_month[m] + op_prof_svc_m[m]
                      + lot_cost_by_month[m])
        dmf_this = dmf_base_m * dmf_pct
        dmf_by_month[m] = dmf_this
        dmf_total += dmf_this
        cost_monthly[m] += dmf_this

    # Escalated land total (used for summaries)
    land_escalated = total_td_purchase + total_td_closing

    # Hard costs (for summary display, includes fencing/URD/streetlights)
    hard_costs = (land_escalated + total_plant_cost + total_amenity_cost + total_det_cost +
                  total_det_landscaping + total_other_cost + total_road_cost +
                  total_road_landscaping + total_dev_cost + total_lot_landscaping +
                  total_fencing_cost + total_urd_cost + total_lot_streetlight_cost +
                  road_streetlight_total + site_work_total)

    # Project Contingency — matches Excel (PP!F37, updated to exclude land)
    # Total base = plants + amenities + detention + sections + other + roads
    #            + fencing + dry_utilities + site_work + landscaping + prof_services
    # Excludes: land, taxes, marketing, mailboxes, brokerage, closing, all operating costs
    # Monthly spread: proportional to monthly costs (CF rows 9-17,26-30), 6-month offset
    contingency_base = (total_plant_cost + total_amenity_cost + total_det_cost +
                        total_dev_cost + total_other_cost + total_road_cost +
                        total_fencing_cost + (total_urd_cost + total_lot_streetlight_cost + road_streetlight_total) +
                        site_work_total + (total_lot_landscaping + total_det_landscaping + total_road_landscaping) +
                        prof_svc_total)
    contingency_total = contingency_base * contingency

    # Monthly costs used for proportional distribution (Excel CF rows 9-17 + 26-30)
    # Excludes land; includes taxes, prof svc, marketing, mailboxes, section dev
    monthly_cont_costs = [0.0] * (MAX_MONTHS + 1)
    for m in range(1, MAX_MONTHS + 1):
        monthly_cont_costs[m] = (_cc_plants[m] + _cc_amen[m] + _cc_det[m] + _cc_other[m] + _cc_roads[m]
                                 + fencing_by_month[m] + site_work_by_month[m]
                                 + _cc_landscape_m[m] + (urd_by_month[m] + lot_streetlight_by_month[m] + road_streetlight_by_month[m])
                                 + lot_tax_by_month[m] + op_prof_svc_m[m] + op_mkt_exp_m[m]
                                 + lot_mailbox_by_month[m] + lot_cost_by_month[m])
    total_monthly_cont_costs = sum(monthly_cont_costs)

    contingency_by_month = [0.0] * (MAX_MONTHS + 1)
    for m in range(1, MAX_MONTHS + 1):
        if total_monthly_cont_costs > 0 and monthly_cont_costs[m] > 0:
            share = monthly_cont_costs[m] / total_monthly_cont_costs * contingency_total
            target_m = m + 6
            if 1 <= target_m <= MAX_MONTHS:
                contingency_by_month[target_m] += share
                cost_monthly[target_m] += share

    # ── 5. SUMMARY OUTPUTS ────────────────────────────────────────────────────
    # total_revenue = sum of displayed sub-items (ensures table adds up)
    total_revenue = (total_lot_base_rev + total_premium_rev + total_escalation_rev +
                     total_fence_fee_rev + total_mktg_fee_rev +
                     sum(a for _, a in mud_issuances) + sum(a for _, a in wcid_issuances) +
                     res_pod_revenue + comm_pod_revenue)

    infra_cost = (total_plant_cost + total_amenity_cost + total_det_cost + total_det_landscaping +
                  total_other_cost + total_road_cost + total_road_landscaping +
                  total_dev_cost + total_lot_landscaping +
                  total_fencing_cost + total_urd_cost + streetlight_total + site_work_total)

    # Gross costs = sum of all above-the-line displayed cost items (ensures table adds up)
    gross_costs = (land_escalated + total_plant_cost + total_amenity_cost + total_det_cost +
                   total_dev_cost + total_other_cost + total_road_cost +
                   total_fencing_cost + (total_urd_cost + total_lot_streetlight_cost + road_streetlight_total) +
                   site_work_total + (total_lot_landscaping + total_det_landscaping + total_road_landscaping) +
                   legal_mo * legal_end + sum(lot_tax_by_month) + mud_total +
                   insurance_mo * ins_end + marketing_total +
                   sum(lot_brokerage_by_month) + sum(lot_closing_by_month) +
                   sum(lot_mailbox_by_month) + prof_svc_total + contingency_total)

    # Below-the-line items: DMF, Personnel, Bookkeeping, MUD/WCID Recv Fees
    below_line_dmf       = dmf_total
    below_line_personnel = personnel_mo * gen_pers_end + marketing_personnel_mo * mkt_pers_end
    below_line_bk        = bookkeeping_mo * bk_end
    total_recv_fees = sum(mud_recv_fee_by_month) + sum(wcid_recv_fee_by_month)
    below_line_recv_fees = total_recv_fees
    below_line = below_line_dmf + below_line_personnel + below_line_bk + below_line_recv_fees

    total_cost    = gross_costs + below_line
    gross_profit  = total_revenue - gross_costs       # Gross margin amount (revenue - gross costs)
    gross_margin  = gross_profit / total_revenue if total_revenue else 0  # Gross margin %
    roc           = gross_profit / gross_costs if gross_costs else 0      # Return on gross cost

    # 0-indexed final period values for display (match Excel Cost Inputs tab)
    gen_final_period_disp  = gen_pers_end - 1   # Excel D103
    mkt_final_period_disp  = mkt_pers_end - 1   # Excel D104 = B91
    legal_final_period_disp= legal_end - 1       # Excel D108
    mud_final_period_disp  = mud_end_0idx        # Excel E112
    ins_final_period_disp  = ins_end - 1         # Excel D116
    bk_final_period_disp   = bk_end - 1          # Excel D120

    gross_margin_amt = total_revenue - gross_costs
    net_profit    = total_revenue - total_cost
    net_margin    = net_profit / total_revenue if total_revenue else 0

    dev_ac = residential_dev_acres if residential_dev_acres > 0 else 1
    rev_per_dev_ac      = total_revenue / dev_ac
    cost_per_dev_ac     = total_cost / dev_ac
    infra_per_dev_ac    = infra_cost / dev_ac
    gm_per_ac           = gross_profit / dev_ac
    amenities_per_lot   = (total_amenity_cost + total_lot_landscaping + total_det_landscaping
                           + total_road_landscaping + total_fencing_cost) / total_lots if total_lots else 0
    infra_per_lot       = infra_cost / total_lots if total_lots else 0

    home_sales_per_year = sum(r.get("pace", 0) for r in lot_rows if safe(r.get("on", 0))) * 12
    lots_18mo = sum(r.get("lots_18mo", 0) for r in lot_rows if safe(r.get("on", 0)))

    # Cashflow range: extend past proj_months if bonds/revenue occur later
    last_cf_month = proj_months
    for m in range(proj_months + 1, MAX_MONTHS + 1):
        if rev_monthly[m] > 0 or cost_monthly[m] > 0:
            last_cf_month = m

    # XIRR (date-based, matching Excel's XIRR with 365-day convention)
    # App month 1 = closing date (Excel period 0), subsequent months = end-of-month dates
    closing_str = inp.get("closing_date", "")
    try:
        closing_dt = datetime.date.fromisoformat(closing_str) if closing_str else None
    except (ValueError, TypeError):
        closing_dt = None

    cf = [-(cost_monthly[m]) + rev_monthly[m] for m in range(1, last_cf_month + 1)]
    if closing_dt:
        # Generate end-of-month dates: month 1 = closing date, month N = N-1 months later
        cf_dates = [closing_dt]
        for m in range(2, last_cf_month + 1):
            months_offset = m - 1
            y = closing_dt.year + (closing_dt.month - 1 + months_offset) // 12
            mo = (closing_dt.month - 1 + months_offset) % 12 + 1
            cf_dates.append(_end_of_month(y, mo))
        irr = xirr(cf, cf_dates)
    else:
        irr = npv_irr(cf)

    # Yearly lots/homes for chart — dynamic range based on project length
    final_year = max(math.ceil(proj_months / 12), 1)
    yearly_lots  = {yr: 0 for yr in range(1, final_year + 1)}
    yearly_homes = {yr: 0 for yr in range(1, final_year + 1)}
    for m in range(1, MAX_MONTHS + 1):
        yr = (m - 1) // 12 + 1
        yearly_lots[yr]  = yearly_lots.get(yr, 0) + lot_count_by_month[m]
        yearly_homes[yr] = yearly_homes.get(yr, 0) + home_sales_by_month[m]
    out["yearly_lots"]  = [{"year": k, "lots":  round(v)} for k, v in sorted(yearly_lots.items())]
    out["yearly_homes"] = [{"year": k, "homes": round(v)} for k, v in sorted(yearly_homes.items())]

    # Cashflow by year for chart
    yearly_cf = {}
    yearly_rev = {}
    yearly_cost = {}
    for m in range(1, MAX_MONTHS + 1):
        yr = (m - 1) // 12 + 1
        yearly_rev[yr]  = yearly_rev.get(yr, 0) + rev_monthly[m]
        yearly_cost[yr] = yearly_cost.get(yr, 0) + cost_monthly[m]
        yearly_cf[yr]   = yearly_rev[yr] - yearly_cost[yr]
    out["yearly_cashflow"] = {k: round(v) for k, v in yearly_cf.items() if yearly_rev.get(k, 0) + yearly_cost.get(k, 0) > 0}
    out["yearly_revenue"]  = {k: round(v) for k, v in yearly_rev.items() if v > 0}
    out["yearly_cost_chart"] = {k: round(v) for k, v in yearly_cost.items() if v > 0}

    # Revenue breakdown — sub-categories matching Excel Project Performance
    out["rev_lot_sales"]    = round(total_lot_base_rev)
    out["rev_premiums"]     = round(total_premium_rev)
    out["rev_escalations"]  = round(total_escalation_rev)
    out["rev_fence_fees"]   = round(total_fence_fee_rev)
    out["rev_mktg_fees"]    = round(total_mktg_fee_rev)
    out["rev_mud"]          = round(sum(a for _, a in mud_issuances))
    out["rev_wcid"]         = round(sum(a for _, a in wcid_issuances))
    out["rev_res_pods"]     = round(res_pod_revenue)
    out["rev_comm_pods"]    = round(comm_pod_revenue)

    # Cost breakdown
    out["cost_land"]        = round(total_td_purchase + total_td_closing)
    out["cost_plants"]      = round(total_plant_cost)
    out["cost_amenities"]   = round(total_amenity_cost)
    out["cost_detention"]   = round(total_det_cost)
    out["cost_other"]       = round(total_other_cost)
    out["cost_roads"]       = round(total_road_cost)
    out["cost_lot_dev"]          = round(total_dev_cost)
    out["cost_lot_landscaping"]  = round(total_lot_landscaping + total_det_landscaping + total_road_landscaping)
    out["cost_marketing"]   = round(marketing_total)
    out["cost_prof_svc"]    = round(prof_svc_total)
    out["cost_dmf"]         = round(dmf_total)
    out["cost_site_work"]   = round(site_work_total)
    # Operating cost detail outputs for Cost Inputs tab display
    # Final period values are 0-indexed to match Excel (e.g. D103=118 not 119)
    out["marketing_total"]             = round(marketing_total)
    out["marketing_final_period"]      = mkt_final_period_disp
    out["marketing_per_month"]         = round(mkt_per_m)
    out["prof_services_total"]         = round(prof_svc_total)
    out["prof_services_per_month"]     = round(ps_per_m)
    out["prof_services_final_period"]  = gen_final_period_disp   # same as D95=D103
    out["dmf_total"]                   = round(dmf_total)
    out["personnel"] = {
        "general_total":   round(personnel_mo * gen_pers_end),
        "marketing_total": round(marketing_personnel_mo * mkt_pers_end),
    }
    out["general_final_period"]        = gen_final_period_disp
    out["marketing_pers_final_period"] = mkt_final_period_disp
    out["legal_total"]                 = round(legal_mo * legal_end)
    out["legal_final_period"]          = legal_final_period_disp
    out["mud_hoa_total"]               = round(mud_total)
    out["mud_hoa_final_period"]        = mud_final_period_disp
    out["mud_hoa_monthly"]             = round(mud_mo)
    out["insurance_total"]             = round(insurance_mo * ins_end)
    out["insurance_final_period"]      = ins_final_period_disp
    out["bookkeeping_total"]           = round(bookkeeping_mo * bk_end)
    out["bookkeeping_final_period"]    = bk_final_period_disp
    out["cost_personnel"]   = round(personnel_mo * gen_pers_end + marketing_personnel_mo * mkt_pers_end)
    out["cost_legal"]       = round(legal_mo * legal_end)
    out["cost_mud_hoa"]     = round(mud_total)
    out["cost_insurance"]   = round(insurance_mo * ins_end)
    out["cost_bookkeeping"] = round(bookkeeping_mo * bk_end)
    out["cost_streetlights"] = round(streetlight_total)
    out["cost_fencing"]     = round(total_fencing_cost)
    out["cost_dry_utilities"] = round(total_urd_cost + total_lot_streetlight_cost + road_streetlight_total)
    out["cost_contingency"] = round(contingency_total)
    out["cost_recv_fees"]   = round(total_recv_fees)
    out["cost_brokerage"]   = round(sum(lot_brokerage_by_month))
    out["cost_closing"]     = round(sum(lot_closing_by_month))
    out["cost_lot_taxes"]   = round(sum(lot_tax_by_month))
    out["cost_mailboxes"]   = round(sum(lot_mailbox_by_month))

    # KPIs
    out["gross_profit"]          = round(gross_profit)
    out["gross_margin_pct"]      = gross_margin
    out["return_on_cost"]        = roc
    out["return_on_cost_pct"]    = roc
    out["total_lots"]            = total_lots
    out["lot_av"]                = round(lot_av)
    out["comm_av"]               = round(comm_av)
    out["project_length_months"] = proj_months
    out["project_length_years"]  = round(proj_months / 12, 1)
    out["amenities_per_lot"]     = round(amenities_per_lot)
    out["infra_per_lot"]         = round(infra_per_lot)
    out["unlevered_irr"]         = irr
    out["total_lots_delivered"]  = total_lots
    out["lot_supply_18mo"]       = round(lots_18mo)
    out["home_sales_per_year"]   = round(home_sales_per_year)
    out["dev_acres"]             = round(dev_acres, 2)
    out["residential_dev_acres"] = round(residential_dev_acres, 2)
    out["rev_per_dev_acre"]      = round(rev_per_dev_ac)
    out["cost_per_dev_acre"]     = round(cost_per_dev_ac)
    out["infra_per_dev_acre"]    = round(infra_per_dev_ac)
    out["gm_per_acre"]           = round(gm_per_ac)
    # Compute displayed values so table math is exact (sum of rounded items = totals)
    _gc = (out["cost_land"] + out["cost_plants"] + out["cost_amenities"] + out["cost_detention"] +
           out["cost_lot_dev"] + out["cost_other"] + out["cost_roads"] + out["cost_fencing"] +
           out["cost_dry_utilities"] + out["cost_site_work"] + out["cost_lot_landscaping"] +
           out["cost_legal"] + out["cost_lot_taxes"] + out["cost_mud_hoa"] + out["cost_insurance"] +
           out["cost_marketing"] + out["cost_brokerage"] + out["cost_closing"] +
           out["cost_mailboxes"] + out["cost_prof_svc"] + out["cost_contingency"])
    _tr = (out["rev_lot_sales"] + out["rev_mud"] + out["rev_wcid"] + out["rev_premiums"] +
           out["rev_fence_fees"] + out["rev_escalations"] + out["rev_mktg_fees"] +
           out["rev_res_pods"] + out["rev_comm_pods"])
    _gm = _tr - _gc
    _btl_d = round(below_line_dmf)
    _btl_p = round(below_line_personnel)
    _btl_b = round(below_line_bk)
    _btl_r = round(below_line_recv_fees)
    _btl   = _btl_d + _btl_p + _btl_b + _btl_r
    _nm    = _gm - _btl
    out["total_revenue"]         = _tr
    out["gross_costs"]           = _gc
    out["gross_margin_amt"]      = _gm
    out["gross_margin_of_costs"] = _gm / _gc if _gc else 0
    out["gross_margin_of_rev"]   = _gm / _tr if _tr else 0
    out["below_line_dmf"]        = _btl_d
    out["below_line_personnel"]  = _btl_p
    out["below_line_bookkeeping"]= _btl_b
    out["below_line_recv_fees"]  = _btl_r
    out["below_line_total"]      = _btl
    _tc = _gc + _btl
    out["total_cost"]            = _tc
    out["net_profit"]            = _nm
    out["net_margin_amt"]        = _nm
    out["net_margin_of_costs"]   = _nm / _tc if _tc else 0
    out["net_margin_of_rev"]     = _nm / _tr if _tr else 0
    out["net_margin_pct"]        = _nm / _tr if _tr else 0

    # ── 6. CASHFLOW DETAIL (for Cashflows tab) ───────────────────────────────
    # Reuse pre-computed category arrays (_cc_plants, _cc_amen, etc.) from DMF section
    # Only need to build land and revenue category arrays here.
    Z = lambda: [0.0] * (MAX_MONTHS + 1)
    rc_lot = Z(); rc_res = Z(); rc_comm = Z(); rc_mud = Z()
    cc_land = Z()
    cc_lotdev = Z(); cc_lot_landscape = Z()

    for td in td_rows:
        mm = int(td["period"]) + 1
        if 1 <= mm <= MAX_MONTHS: cc_land[mm] += td["total"]
    for m in range(1, MAX_MONTHS + 1):
        cc_lotdev[m]        = lot_cost_by_month[m]
        cc_lot_landscape[m] = lot_landscaping_by_month[m]
        rc_lot[m]           = lot_rev_by_month[m]
    for i, rp in enumerate(res_pods):
        if i >= res_pod_count: break
        ppa_p = safe(rp.get("price_per_acre"))
        cc_p  = safe(rp.get("closing_costs_pct", 0.045))
        ifl   = safe(rp.get("impact_fee_per_lot", 0))
        lpa   = safe(rp.get("implied_lots_per_acre", 0))
        sp    = int(safe(rp.get("sale_period", proj_months)))
        pr    = acres_per_res_pod * ppa_p * (1 - cc_p) + ifl * lpa * acres_per_res_pod
        if pr > 0 and 1 <= sp <= MAX_MONTHS: rc_res[sp] += pr
    for i, cp in enumerate(comm_pods):
        if i >= comm_pod_count: break
        psf_p = safe(cp.get("price_per_sf"))
        cc_p  = safe(cp.get("closing_costs_pct", 0.045))
        sp    = int(safe(cp.get("sale_period", proj_months)))
        pr    = acres_per_comm_pod * 43560 * psf_p * (1 - cc_p)
        if pr > 0 and 1 <= sp <= MAX_MONTHS: rc_comm[sp] += pr
    for bp, amt in mud_issuances:
        rc_mud[bp] += amt
    for bp, amt in wcid_issuances:
        rc_mud[bp] += amt

    # Monthly detail covers full activity range (last_cf_month computed above for IRR)
    out["cf_monthly"] = [
        {
            "month": m,
            "yr": (m - 1) // 12 + 1,
            "revenue": round(rev_monthly[m]),
            "cost": round(cost_monthly[m]),
            "net": round(rev_monthly[m] - cost_monthly[m]),
            "rev_lot_sales": round(rc_lot[m]),
            "rev_res_pods":  round(rc_res[m]),
            "rev_comm_pods": round(rc_comm[m]),
            "rev_mud_wcid":  round(rc_mud[m]),
            "cost_land":     round(cc_land[m]),
            "cost_plants":   round(_cc_plants[m]),
            "cost_amenities":round(_cc_amen[m]),
            "cost_detention":round(_cc_det[m]),
            "cost_other":    round(_cc_other[m]),
            "cost_roads":    round(_cc_roads[m]),
            "cost_lot_dev":          round(cc_lotdev[m]),
            "cost_lot_landscaping":  round(_cc_landscape_m[m]),
            "cost_fencing":          round(fencing_by_month[m]),
            "cost_dry_utilities":    round(urd_by_month[m] + lot_streetlight_by_month[m] + road_streetlight_by_month[m]),
            "cost_site_work":        round(site_work_by_month[m]),
            "cost_dmf":              round(dmf_by_month[m]),
            "cost_operating":round(max(0, cost_monthly[m] - cc_land[m] - _cc_plants[m] - _cc_amen[m]
                                       - _cc_det[m] - _cc_other[m] - _cc_roads[m] - cc_lotdev[m]
                                       - _cc_landscape_m[m] - fencing_by_month[m]
                                       - urd_by_month[m] - lot_streetlight_by_month[m]
                                       - road_streetlight_by_month[m] - site_work_by_month[m]
                                       - dmf_by_month[m])),
        }
        for m in range(1, last_cf_month + 1)
    ]
    # Quarterly aggregation for chart
    q_data = {}
    for m in range(1, last_cf_month + 1):
        q = (m - 1) // 3 + 1
        if q not in q_data:
            q_data[q] = {"q": q, "yr": (q - 1) // 4 + 1, "qtr": (q - 1) % 4 + 1, "revenue": 0, "cost": 0}
        q_data[q]["revenue"] += rev_monthly[m]
        q_data[q]["cost"] += cost_monthly[m]
    for q in q_data:
        q_data[q]["net"] = round(q_data[q]["revenue"] - q_data[q]["cost"])
        q_data[q]["revenue"] = round(q_data[q]["revenue"])
        q_data[q]["cost"] = round(q_data[q]["cost"])
    out["cf_quarterly"] = [v for _, v in sorted(q_data.items())]

    return out
