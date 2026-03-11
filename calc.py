"""
Ember Tract Underwriting - Pure Python Calculation Engine
Faithfully ports all Excel formulas from the 10-sheet model.
No Excel or openpyxl required at runtime.
"""
import math
from typing import Any

# ─── LOOKUP TABLES (from Calc_Lookups sheet) ─────────────────────────────────
PLANT_LOOKUPS = {
    "WWTP":        {"acres": 10.0,  "duration": 8},
    "Water Plant": {"acres": 3.5,   "duration": 8},
    "Lift Station":{"acres": 0.75,  "duration": 3},
    "None":        {"acres": 0.0,   "duration": 0},
}
AMENITY_LOOKUPS = {
    "Pocket Park":          {"duration": 2},
    "Small Amenity Center": {"duration": 8},
    "Large Amenity Center": {"duration": 12},
    "None":                 {"duration": 0},
}
ROAD_LOOKUPS = {
    "2 Lane": {"wsd": 450, "paving": 343},
    "4 Lane": {"wsd": 460, "paving": 663},
}

def safe(v, default=0):
    """Return v if numeric, else default."""
    if v is None or v == "" or (isinstance(v, float) and math.isnan(v)):
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default

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


def calculate(inp: dict) -> dict:
    """
    Master calculation function.
    inp: flat dict of all user inputs (see schema in app.py)
    Returns: dict of all computed outputs for display.
    """
    out = {}

    # ── 1. TRACT INPUTS derived values ───────────────────────────────────────
    gross_ac      = safe(inp.get("gross_acreage"))
    ppa           = safe(inp.get("purchase_price_per_acre"))
    escalator     = safe(inp.get("land_escalator"))      # decimal e.g. 0.03
    closing_pct   = safe(inp.get("closing_costs_pct"))   # decimal e.g. 0.015

    purchase_price = ppa * gross_ac                        # B10
    closing_costs  = closing_pct * purchase_price          # B12
    total_land     = purchase_price + closing_costs         # B13

    # Plant acres lookup
    plants = inp.get("plants", [])  # list of {type, notes}
    plant_acres_list = [PLANT_LOOKUPS.get(p.get("type","None"), {"acres":0})["acres"] for p in plants]
    total_plant_acres = sum(plant_acres_list)

    # Detention
    det_rate        = safe(inp.get("det_storage_rate", 0.5))
    det_depth       = safe(inp.get("det_depth", 3))
    det_num         = safe(inp.get("det_num_projects", 1))
    parks_pct       = safe(inp.get("parks_pct", 0.02))
    drill_site_ac   = safe(inp.get("drill_site_acres", 0))

    # Net-out area (B72): plants + detention + amenities + other + roads
    amenities = inp.get("amenities", [])  # list of {type, acres}
    amenity_acres_total = sum(safe(a.get("acres")) for a in amenities)

    other_netouts = inp.get("other_netouts", [])  # list of {desc, acres}
    other_acres_total = sum(safe(o.get("acres")) for o in other_netouts)

    roads = inp.get("roads", [])  # list of {type, lf, width, road_setback, landscaping_setback}
    road_acres_list = []
    for r in roads:
        lf = safe(r.get("lf"))
        w  = safe(r.get("width"))
        rs = safe(r.get("road_setback"))
        ls = safe(r.get("landscaping_setback"))
        if lf:
            road_acres_list.append(lf * (w + (ls + rs) * 2) / 43560)
        else:
            road_acres_list.append(0)
    road_acres_total = sum(road_acres_list)

    # Detention acres: (det_rate * gross_ac) / det_depth * 1.3
    det_acres_each = (det_rate * gross_ac / det_depth * 1.3) if det_depth else 0
    det_total_acres = det_acres_each * det_num

    parks_acres = parks_pct * gross_ac
    net_out_total = total_plant_acres + det_total_acres + amenity_acres_total + parks_acres + drill_site_ac + other_acres_total + road_acres_total  # B72
    dev_acres = gross_ac - net_out_total  # B73

    comm_pod_acres = safe(inp.get("commercial_pod_acres"))
    res_pod_acres  = safe(inp.get("residential_pod_acres"))
    residential_dev_acres = dev_acres - comm_pod_acres - res_pod_acres  # B78

    out["purchase_price"] = purchase_price
    out["closing_costs"]  = closing_costs
    out["total_land_cost"]= total_land
    out["net_out_total"]  = net_out_total
    out["dev_acres"]      = dev_acres
    out["residential_dev_acres"] = residential_dev_acres
    out["det_acres_each"] = det_acres_each
    out["plant_acres_list"] = plant_acres_list
    out["road_acres_list"]  = road_acres_list

    # ── 2. COST INPUTS derived values ─────────────────────────────────────────
    default_other_pct      = safe(inp.get("default_other_pct", 0.17))
    contingency            = safe(inp.get("contingency", 0.05))
    cost_per_mailbox       = safe(inp.get("cost_per_mailbox", 300))
    cost_per_streetlight   = safe(inp.get("cost_per_streetlight", 5000))
    default_start_month    = safe(inp.get("default_start_month", 1))
    site_work_pct          = safe(inp.get("site_work_pct", 0.10))
    fenced_pct             = safe(inp.get("fenced_pct", 0.50))
    sectional_other_pct    = safe(inp.get("sectional_other_pct", 0.20))
    landscaping_other_pct  = safe(inp.get("landscaping_other_pct", 0.10))

    # Land cost takedowns — default to 1 takedown at 100% period 0 if none entered
    takedowns = inp.get("takedowns", [])
    if not takedowns or all(safe(td.get("pct", 0)) == 0 for td in takedowns):
        takedowns = [{"period": 0, "pct": 1.0}]
    td_rows = []
    for i, td in enumerate(takedowns):
        period = safe(td.get("period"))
        pct    = safe(td.get("pct"))
        if i == 0:
            purchase = purchase_price * pct
            closing  = closing_costs * pct
        else:
            purchase = purchase_price * pct * (1 + escalator) ** (period / 12)
            closing  = closing_costs * pct
        td_rows.append({"period": period, "pct": pct, "purchase": purchase, "closing": closing, "total": purchase + closing})
    out["takedown_rows"] = td_rows

    # Plant cost rows
    plant_costs = inp.get("plant_costs", [])  # {base_cost, other_pct, start_month, ph2_base_cost, ph2_other_pct, ph2_start_month}
    plant_cost_rows = []
    for i in range(max(len(plants), len(plant_costs), 8)):
        pc = plant_costs[i] if i < len(plant_costs) else {}
        ptype = plants[i].get("type", "None") if i < len(plants) else "None"
        dur = PLANT_LOOKUPS.get(ptype, {"duration": 0})["duration"]
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
        dur = AMENITY_LOOKUPS.get(atype, {"duration": 0})["duration"]
        base = safe(ac.get("base_cost"))
        other_p = safe(ac.get("other_pct", default_other_pct))
        total = base * (1 + other_p) if base else 0
        sm = safe(ac.get("start_month", default_start_month))
        amenity_cost_rows.append({
            "type": atype,
            "acres": safe(amenities[i].get("acres")) if i < len(amenities) else 0,
            "base_cost": base, "other_pct": other_p, "total_cost": total,
            "start_month": sm, "duration": dur,
        })
    out["amenity_cost_rows"] = amenity_cost_rows
    total_amenity_cost = sum(r["total_cost"] for r in amenity_cost_rows)

    # Detention cost rows
    det_costs = inp.get("det_costs", [])  # list per detention project
    det_sq_ft = det_acres_each * 43560
    det_cu_yd = det_sq_ft / 27
    det_cost_rows = []
    for idx in range(int(det_num)):
        dc = det_costs[idx] if idx < len(det_costs) else {}
        base_cost_per_cyd = safe(dc.get("base_cost_per_cyd", 0))
        other_p = safe(dc.get("other_pct", default_other_pct))
        base_cost = det_cu_yd * base_cost_per_cyd * (1 + safe(inp.get("land_escalator", 0))) * 0.3  # ~Excel J45 pattern
        total_cost = base_cost * (1 + other_p)
        sm_base = safe(inp.get("default_start_month", 1))
        sm = sm_base + idx * 15  # each project offset by 15 months
        dur = 9  # typical detention duration
        delivery = dur + sm - 1
        lsf = safe(dc.get("landscaping_per_foot", 0))
        perimeter = det_sq_ft ** 0.5 * 4 if det_sq_ft else 0
        total_landscaping = lsf * perimeter
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
        desc = other_netouts[i].get("desc", "") if i < len(other_netouts) else ""
        acres = safe(other_netouts[i].get("acres")) if i < len(other_netouts) else 0
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

    # Road cost rows
    road_costs = inp.get("road_costs", [])
    road_cost_rows = []
    for i, rc in enumerate(road_costs):
        rtype = roads[i].get("type", "") if i < len(roads) else ""
        lf = safe(roads[i].get("lf")) if i < len(roads) else 0
        ls_setback = safe(roads[i].get("landscaping_setback")) if i < len(roads) else 0
        rl = ROAD_LOOKUPS.get(rtype, {"wsd": 0, "paving": 0})
        wsd_per_lf  = rl["wsd"]
        pave_per_lf = rl["paving"]
        base_cost = lf * (wsd_per_lf + pave_per_lf) if lf else 0
        other_p = safe(rc.get("other_pct", default_other_pct))
        total_cost = base_cost * (1 + other_p) if base_cost else 0
        sm = safe(rc.get("start_month", default_start_month))
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

    # Lot size mix
    lot_sizes = inp.get("lot_sizes", [])  # [{on, lot_sf, yield_per_ac, pace, home_price, wsd_per_ff, paving_per_ff, dev_start_month, landscaping_per_lot, urd_per_lot, lots_per_streetlight, fence_cost_per_ff}]
    lot_rows = []
    for ls in lot_sizes:
        if not safe(ls.get("on", 0)):
            lot_rows.append({**ls, "lots_18mo": 0, "total_lots": 0, "acres_18mo": 0,
                             "dev_cost_per_lot": 0, "wsd_per_ff_calc": 0, "paving_per_ff_calc": 0})
            continue
        lot_sf   = safe(ls.get("lot_sf", 6000))
        ff       = lot_sf / safe(ls.get("depth", 120), 1)  # front footage = sf / depth
        yield_ac = safe(ls.get("yield_per_ac"))
        pace     = safe(ls.get("pace"))  # lots/month
        total_lots_this = int(residential_dev_acres * yield_ac) if yield_ac else 0
        lots_18mo = pace * 18 if pace else 0
        acres_18mo = lots_18mo / yield_ac if yield_ac else 0
        # Dev cost per lot
        wsd_ff    = safe(ls.get("wsd_per_ff", 0))
        pave_ff   = safe(ls.get("paving_per_ff", 0))
        ls_lot    = safe(ls.get("landscaping_per_lot", 0))
        urd_lot   = safe(ls.get("urd_per_lot", 0))
        fence_ff  = safe(ls.get("fence_cost_per_ff", 0))
        dev_cost  = (wsd_ff + pave_ff) * ff + ls_lot + urd_lot + fence_ff * ff * fenced_pct
        lot_rows.append({
            **ls,
            "total_lots": total_lots_this,
            "lots_18mo": lots_18mo,
            "acres_18mo": acres_18mo,
            "dev_cost_per_lot": dev_cost,
            "ff": ff,
        })
    out["lot_rows"] = lot_rows

    total_lots = sum(r.get("total_lots", 0) for r in lot_rows)
    total_dev_cost = sum(r.get("dev_cost_per_lot", 0) * r.get("total_lots", 0) for r in lot_rows)
    total_lot_landscaping = sum(safe(r.get("landscaping_per_lot")) * r.get("total_lots", 0) for r in lot_rows)

    out["total_lots"] = total_lots

    # Operating costs
    project_length_months = safe(inp.get("project_length_months", 60))  # estimated, refined below
    marketing_pct   = safe(inp.get("marketing_pct", 0.02))
    prof_svc_pct    = safe(inp.get("prof_svc_pct", 0.02))
    dmf_pct         = safe(inp.get("dmf_pct", 0.005))
    personnel_mo    = safe(inp.get("personnel_monthly", 0))
    legal_mo        = safe(inp.get("legal_monthly", 0))
    mud_mo          = safe(inp.get("mud_monthly", 0))
    mud_pct         = safe(inp.get("mud_pct", 0))
    insurance_mo    = safe(inp.get("insurance_monthly", 0))
    bookkeeping_mo  = safe(inp.get("bookkeeping_monthly", 0))

    # ── 3. REVENUE INPUTS derived values ──────────────────────────────────────
    timing_method = inp.get("timing_method", "1 Takedown")
    take1_pct = 1.0 if timing_method == "1 Takedown" else 0.5
    take2_pct = 0.0 if timing_method == "1 Takedown" else (0.5 if timing_method == "50/50" else 0.25)
    take3_pct = 0.25 if timing_method == "50/25/25" else 0.0

    brokerage_fees    = safe(inp.get("brokerage_fees", 0.03))
    lot_closing_costs = safe(inp.get("lot_closing_costs", 0.01))
    bem_period        = safe(inp.get("bem_period", 0))
    bem_pct           = safe(inp.get("bem_pct", 0))

    # $/FF by year
    ff_by_year = [safe(inp.get(f"ff_year_{y}", 0)) for y in range(11)]  # years 0–10

    # Home build table (lot sizes → home revenue)
    home_price_per_row = [safe(ls.get("home_price", 0)) for ls in lot_sizes]

    # Residential pods
    res_pods = inp.get("res_pods", [])
    res_pod_revenue = 0
    for rp in res_pods:
        acres = safe(rp.get("acres"))
        ppa_pod = safe(rp.get("price_per_acre"))
        cc = safe(rp.get("closing_costs_pct", 0.01))
        res_pod_revenue += acres * ppa_pod * (1 - cc)

    # Commercial pods
    comm_pods = inp.get("comm_pods", [])
    comm_pod_revenue = 0
    comm_av = 0
    for cp in comm_pods:
        acres = safe(cp.get("acres"))
        psf   = safe(cp.get("price_per_sf"))
        cc    = safe(cp.get("closing_costs_pct", 0.01))
        av_ac = safe(cp.get("av_per_acre"))
        comm_pod_revenue += acres * 43560 * psf * (1 - cc)
        comm_av += acres * av_ac

    # MUD / WCID bonds
    mud_row   = inp.get("mud_bond", {})
    wcid_row  = inp.get("wcid_bond", {})
    mud_bond_rev  = safe(mud_row.get("amount")) * safe(mud_row.get("reimbursement_pct", 0.8)) if mud_row else 0
    wcid_bond_rev = safe(wcid_row.get("amount")) * safe(wcid_row.get("reimbursement_pct", 0.8)) if wcid_row else 0

    # ── 4. CASHFLOW ENGINE ────────────────────────────────────────────────────
    # Build a monthly cashflow array (up to 360 months)
    MAX_MONTHS = 360

    # Determine project timeline
    # Earliest start = month 1, latest delivery = max delivery period across all cost items
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
                pace = safe(lr.get("pace", 0))
                sm   = safe(lr.get("dev_start_month", default_start_month))
                lots = lr.get("total_lots", 0)
                months_to_sell = math.ceil(lots / pace) + int(sm) + 18 if pace else 0
                max_lot_months = max(max_lot_months, months_to_sell)
        all_deliveries.append(max_lot_months)

    proj_months = min(max(all_deliveries + [60]), MAX_MONTHS) if all_deliveries else 60
    project_length_months = proj_months
    out["project_length_months"] = proj_months

    # Monthly revenue array
    rev_monthly = [0.0] * (MAX_MONTHS + 1)
    cost_monthly = [0.0] * (MAX_MONTHS + 1)

    # Land cost: takedowns
    for td in td_rows:
        m = max(1, int(td["period"]))
        if m <= MAX_MONTHS:
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
    for r in other_cost_rows:
        spread_cost(cost_monthly, r["total_cost"], r["start_month"], max(r["duration"], 1))
    for r in road_cost_rows:
        spread_cost(cost_monthly, r["total_cost"], r["start_month"], r["duration"])
        spread_cost(cost_monthly, r["total_landscaping"], r["start_month"], r["duration"])

    # Lot development costs + revenues
    lot_rev_by_month = [0.0] * (MAX_MONTHS + 1)
    lot_cost_by_month = [0.0] * (MAX_MONTHS + 1)
    lot_count_by_month = [0] * (MAX_MONTHS + 1)

    for lr in lot_rows:
        if not safe(lr.get("on", 0)) or lr.get("total_lots", 0) == 0:
            continue
        pace = safe(lr.get("pace", 0))
        if pace <= 0:
            continue
        total = int(lr["total_lots"])
        sm    = int(safe(lr.get("dev_start_month", default_start_month)))
        dev_cost = lr.get("dev_cost_per_lot", 0)
        ls_lot   = safe(lr.get("landscaping_per_lot", 0))

        # $/FF lookup for lot revenue
        delivered = 0
        m = sm
        while delivered < total and m <= MAX_MONTHS:
            batch = min(int(pace), total - delivered)
            lot_count_by_month[m] += batch
            # Development cost paid at delivery
            lot_cost_by_month[m] += batch * dev_cost
            lot_cost_by_month[m] += batch * ls_lot
            delivered += batch
            m += 1

    # Lot revenue: use $/FF × lot FF × takedown structure
    # Revenue comes in 18 months after dev start (finished lots sell after construction)
    revenue_start_offset = 18
    for lr in lot_rows:
        if not safe(lr.get("on", 0)) or lr.get("total_lots", 0) == 0:
            continue
        pace = safe(lr.get("pace", 0))
        if pace <= 0:
            continue
        total = int(lr["total_lots"])
        sm    = int(safe(lr.get("dev_start_month", default_start_month)))
        ff    = lr.get("ff", 0)
        lot_sf = safe(lr.get("lot_sf", 6000))

        delivered = 0
        m = sm + revenue_start_offset
        while delivered < total and m <= MAX_MONTHS:
            batch = min(int(pace), total - delivered)
            year_idx = min(int((m - 1) / 12), 10)
            ff_rate = ff_by_year[year_idx] if year_idx < len(ff_by_year) else ff_by_year[-1]
            gross_lot_rev = batch * ff * ff_rate
            net_lot_rev = gross_lot_rev * (1 - brokerage_fees - lot_closing_costs)
            # Apply takedown timing
            if m <= MAX_MONTHS:
                lot_rev_by_month[m] += net_lot_rev * take1_pct
            if m + 6 <= MAX_MONTHS:
                lot_rev_by_month[m + 6] += net_lot_rev * take2_pct
            if m + 9 <= MAX_MONTHS:
                lot_rev_by_month[m + 9] += net_lot_rev * take3_pct
            delivered += batch
            m += 1

    for m in range(1, MAX_MONTHS + 1):
        rev_monthly[m] += lot_rev_by_month[m]
        cost_monthly[m] += lot_cost_by_month[m]

    # Res/Comm pod revenues (lump sums)
    # Simplified: at month when all lots are delivered
    pod_month = proj_months
    rev_monthly[min(pod_month, MAX_MONTHS)] += res_pod_revenue + comm_pod_revenue + mud_bond_rev + wcid_bond_rev

    # Operating costs spread over project life
    total_lot_revenue_gross = sum(lot_rev_by_month)
    marketing_total = total_lot_revenue_gross * marketing_pct
    prof_svc_total  = total_lot_revenue_gross * prof_svc_pct
    dmf_total       = total_lot_revenue_gross * dmf_pct

    spread_cost(cost_monthly, marketing_total, 1, proj_months)
    spread_cost(cost_monthly, prof_svc_total,  1, proj_months)
    for m in range(1, proj_months + 1):
        if m <= MAX_MONTHS:
            cost_monthly[m] += personnel_mo + legal_mo + insurance_mo + bookkeeping_mo
            cost_monthly[m] += mud_mo

    # Streetlight cost
    streetlight_total = total_streetlights * cost_per_streetlight
    spread_cost(cost_monthly, streetlight_total, 1, max(proj_months, 1))

    # ── 5. SUMMARY OUTPUTS ────────────────────────────────────────────────────
    total_revenue = sum(rev_monthly[1:proj_months+1])
    total_cost    = sum(cost_monthly[1:proj_months+1])
    gross_profit  = total_revenue - total_cost
    gross_margin  = gross_profit / total_revenue if total_revenue else 0
    roc           = gross_profit / total_cost if total_cost else 0

    infra_cost = total_plant_cost + total_amenity_cost + total_det_cost + total_other_cost + total_road_cost + total_road_landscaping + total_dev_cost
    below_line = dmf_total + (personnel_mo + legal_mo + insurance_mo + bookkeeping_mo + mud_mo) * proj_months

    net_profit    = gross_profit - below_line
    net_margin    = net_profit / total_revenue if total_revenue else 0

    dev_ac = residential_dev_acres if residential_dev_acres > 0 else 1
    rev_per_dev_ac      = total_revenue / dev_ac
    cost_per_dev_ac     = total_cost / dev_ac
    infra_per_dev_ac    = infra_cost / dev_ac
    gm_per_ac           = gross_profit / dev_ac
    amenities_per_lot   = total_amenity_cost / total_lots if total_lots else 0
    infra_per_lot       = infra_cost / total_lots if total_lots else 0

    # Lot AV
    lot_av = 0
    for lr in lot_rows:
        if safe(lr.get("on", 0)) and lr.get("total_lots", 0):
            hp = safe(lr.get("home_price", 0))
            lot_av += lr["total_lots"] * hp * 0.01 * mround(safe(lr.get("lot_sf",6000)) * 0.01, 100)

    home_sales_per_year = sum(r.get("pace", 0) for r in lot_rows if safe(r.get("on", 0))) * 12
    lots_18mo = sum(r.get("lots_18mo", 0) for r in lot_rows if safe(r.get("on", 0)))

    # IRR (unlevered monthly cashflows)
    cf = [-(cost_monthly[m]) + rev_monthly[m] for m in range(1, proj_months + 1)]
    irr = npv_irr(cf)

    # Yearly lots/homes for chart
    yearly_lots  = {}
    yearly_homes = {}
    for m in range(1, MAX_MONTHS + 1):
        yr = (m - 1) // 12 + 1
        yearly_lots[yr]  = yearly_lots.get(yr, 0) + lot_count_by_month[m]
    out["yearly_lots"] = {k: v for k, v in yearly_lots.items() if v > 0}

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

    # Revenue breakdown
    out["rev_lot_sales"]    = round(sum(lot_rev_by_month))
    out["rev_res_pods"]     = round(res_pod_revenue)
    out["rev_comm_pods"]    = round(comm_pod_revenue)
    out["rev_mud_wcid"]     = round(mud_bond_rev + wcid_bond_rev)

    # Cost breakdown
    out["cost_land"]        = round(total_land)
    out["cost_plants"]      = round(total_plant_cost)
    out["cost_amenities"]   = round(total_amenity_cost)
    out["cost_detention"]   = round(total_det_cost)
    out["cost_other"]       = round(total_other_cost)
    out["cost_roads"]       = round(total_road_cost + total_road_landscaping)
    out["cost_lot_dev"]     = round(total_dev_cost + total_lot_landscaping)
    out["cost_marketing"]   = round(marketing_total)
    out["cost_prof_svc"]    = round(prof_svc_total)
    out["cost_dmf"]         = round(dmf_total)
    out["cost_personnel"]   = round(personnel_mo * proj_months)
    out["cost_legal"]       = round(legal_mo * proj_months)
    out["cost_mud_hoa"]     = round(mud_mo * proj_months)
    out["cost_insurance"]   = round(insurance_mo * proj_months)
    out["cost_bookkeeping"] = round(bookkeeping_mo * proj_months)
    out["cost_streetlights"]= round(streetlight_total)

    # KPIs
    out["gross_profit"]          = round(gross_profit)
    out["gross_margin_pct"]      = gross_margin
    out["return_on_cost"]        = roc
    out["net_profit"]            = round(net_profit)
    out["net_margin_pct"]        = net_margin
    out["total_revenue"]         = round(total_revenue)
    out["total_cost"]            = round(total_cost)
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
    out["below_line_total"]      = round(below_line)

    return out
