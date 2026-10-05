"""Tests for the acquisition -> underwriting derivation.

These numbers end up in a land offer, so the arithmetic is pinned here rather
than checked by eye in the browser. Run: python test_acq_to_uw.py
"""

import acq_to_uw as A


def _base():
    return {
        "gross_acreage": 0,
        "price_per_ff": {str(y): 1800 for y in range(11)},
        "other_netouts": [{"desc": "", "acres": 0, "notes": ""} for _ in range(6)],
        "lot_sizes": [{"front_footage": w, "on": 0, "yield_per_ac": 5,
                       "pace": 5, "home_price": 200000} for w in A.UW_LOT_WIDTHS],
    }


def check(name, got, want, tol=1e-6):
    ok = abs(got - want) <= tol if isinstance(want, (int, float)) else got == want
    print(("  PASS  " if ok else "  FAIL  ") + name)
    if not ok:
        print("          got %r, want %r" % (got, want))
    return ok


results = []

# --- pace comes off starts, not closings -----------------------------------
row = {"est_annual_starts": 168, "est_annual_closings": 999}
results.append(check("pace uses starts at 20% capture",
                     A.recommended_pace(row, 0.20), 168 * 0.20 / 12))
results.append(check("pace scales with capture",
                     A.recommended_pace(row, 0.40), 168 * 0.40 / 12))
results.append(check("no starts means no recommendation",
                     A.recommended_pace({"est_annual_starts": 0}, 0.2), None))

# --- lot value is a share of home price ------------------------------------
results.append(check("lot value at 22% of home price",
                     A.lot_value({"avg_price": 385000}, 0.22), 84700.0))

# --- the blend holds total revenue -----------------------------------------
# Two widths priced individually must produce the same revenue as the blend.
priced = [(40, 820, 318000), (50, 610, 398000), (60, 300, 489000), (70, 90, 612000)]
ratio = 0.22
blend = A.blended_price_per_ff(priced, ratio)
rev_individual = sum(lots * home * ratio for _ff, lots, home in priced)
rev_blended = sum(lots * ff * blend for ff, lots, _home in priced)
results.append(check("blended $/FF preserves total lot revenue",
                     rev_blended, rev_individual, tol=0.01))
results.append(check("blend sits between the per-width rates",
                     min(h * ratio / f for f, _l, h in priced) <= blend
                     <= max(h * ratio / f for f, _l, h in priced), True))

# --- off-grid market widths snap onto the model's table --------------------
mkt = A.market_by_width({"lot_widths": [
    {"lot_width_ff": 46, "lots": 240, "avg_price": 352000, "est_annual_starts": 54},
    {"lot_width_ff": 45, "lots": 160, "avg_price": 366000, "est_annual_starts": 36},
]})
results.append(check("46 FF and 45 FF merge into one 45 FF row", sorted(mkt), [45]))
results.append(check("merged lots add", A.market_by_width(
    {"lot_widths": [{"lot_width_ff": 46, "lots": 240},
                    {"lot_width_ff": 45, "lots": 160}]})[45]["lots"], 400))
results.append(check("merged price weights by lot count",
                     mkt[45]["avg_price"],
                     round((352000 * 240 + 366000 * 160) / 400, 2)))
results.append(check("merged starts add", mkt[45]["est_annual_starts"], 90))

# --- constraints carry marginal acreage, never footprints ------------------
analysis = {"netout_detail": [
    {"label": "Floodplain (100-yr)", "applied": True, "acres": 29.0, "acres_marginal": 29.0},
    {"label": "Wetlands (NWI)", "applied": True, "acres": 10.9, "acres_marginal": 9.9},
    {"label": "Stream buffer", "applied": True, "acres": 12.5, "acres_marginal": 5.8},
    {"label": "Not applied", "applied": False, "acres": 50.0, "acres_marginal": 50.0},
    {"label": "Layer failed", "applied": True, "error": True, "acres": 40.0},
]}
no = A.constraint_netouts(analysis)
results.append(check("marginal acreage, not footprint",
                     round(sum(r["acres"] for r in no), 2), 44.7))
results.append(check("kept-by-override layers are excluded",
                     any(r["desc"] == "Not applied" for r in no), False))
results.append(check("failed layers are excluded",
                     any(r["desc"] == "Layer failed" for r in no), False))
results.append(check("always fills the model's six slots", len(no), 6))

# more layers than slots: acreage is rolled up, never dropped
many = {"netout_detail": [
    {"label": "L%d" % i, "applied": True, "acres": 10 - i, "acres_marginal": 10 - i}
    for i in range(8)]}
rolled = A.constraint_netouts(many)
results.append(check("overflow rolls up rather than dropping acreage",
                     round(sum(r["acres"] for r in rolled), 2),
                     round(sum(10 - i for i in range(8)), 2)))

# --- end to end -------------------------------------------------------------
cbas = {"lot_widths": [
    {"lot_width_ff": 40, "lots": 820, "avg_price": 318000, "est_annual_starts": 168},
    {"lot_width_ff": 50, "lots": 610, "avg_price": 398000, "est_annual_starts": 132},
]}
full = {"gross_acres": 207.5, "gross_basis": "stated",
        "netout_detail": analysis["netout_detail"],
        "yield_estimates": {"breakdown": [
            {"label": "40 FF", "units_per_acre": 6.0},
            {"label": "50 FF", "units_per_acre": 5.0},
            {"label": "90 FF", "units_per_acre": 1.5}]}}
inp, bas = A.derive_uw_inputs(_base(), full, cbas)
on = {r["front_footage"]: r for r in inp["lot_sizes"] if r["on"]}
results.append(check("only the analyst's mix is switched on", sorted(on), [40, 50, 90]))
results.append(check("yield comes from the mix", on[40]["yield_per_ac"], 6.0))
results.append(check("pace comes from the market", on[40]["pace"], round(168 * .2 / 12, 2)))
results.append(check("a width with no market read keeps the model's pace",
                     on[90]["pace"], 5))
results.append(check("widths lacking a market read are reported",
                     bas["_widths_no_market"], [90]))
results.append(check("gross acreage carries its basis", inp["gross_acreage"], 207.5))

# Pricing is the underwriter's call. Nothing here may quietly set it.
results.append(check("home price is NOT applied", on[40]["home_price"], 200000))
results.append(check("$/FF is NOT applied", inp["price_per_ff"]["0"], 1800))
results.append(check("a $/FF is still suggested for the chart",
                     round(bas["_suggested_price_per_ff"], 2),
                     round(A.blended_price_per_ff(
                         [(40, 820, 318000), (50, 610, 398000)], 0.22), 2)))

# --- market evidence --------------------------------------------------------
ev_cbas = dict(cbas, builder_lot_widths=[
    {"name": "Builder A", "lot_width_ff": 40, "lots": 500, "avg_price": 309000,
     "min_price": 271000, "max_price": 358000, "avg_sqft": 1810, "plans": 7},
    {"name": "Builder B", "lot_width_ff": 40, "lots": 320, "avg_price": 332000,
     "min_price": 298000, "max_price": 391000, "avg_sqft": 1950, "plans": 5},
    {"name": "Builder C", "lot_width_ff": 50, "lots": 610, "avg_price": 398000,
     "min_price": 344000, "max_price": 470000, "avg_sqft": 2280, "plans": 9},
])
ev = A.market_evidence(ev_cbas, mix_widths={40, 50, 90}, lot_ratio=0.22)
w40 = next(w for w in ev["widths"] if w["ff"] == 40)
results.append(check("evidence breaks price out by builder",
                     [b["name"] for b in w40["builders"]], ["Builder A", "Builder B"]))
results.append(check("builders are ordered by lot count",
                     w40["builders"][0]["lots"], 500))
results.append(check("the range behind the average is carried",
                     (w40["builders"][0]["min_price"], w40["builders"][0]["max_price"]),
                     (271000, 358000)))
results.append(check("implied lot $/FF is shown per width",
                     w40["implied_lot_ff"], round(318000 * 0.22 / 40, 2)))
results.append(check("widths in the product mix are flagged", w40["in_mix"], True))
results.append(check("the suggestion blends only the mix widths",
                     round(ev["suggested_price_per_ff"], 2),
                     round(A.blended_price_per_ff(
                         [(40, 820, 318000), (50, 610, 398000)], 0.22), 2)))
results.append(check("no market means no suggestion",
                     A.market_evidence({}, {40}, 0.22)["suggested_price_per_ff"], None))

# nothing to go on at all: the model's defaults must come through untouched
empty_in, empty_bas = A.derive_uw_inputs(_base(), {}, {})
results.append(check("no analysis leaves gross at the model default",
                     empty_in["gross_acreage"], 0))
results.append(check("no market leaves $/FF at the model default",
                     empty_in["price_per_ff"]["0"], 1800))
results.append(check("no mix switches nothing on",
                     sum(r["on"] for r in empty_in["lot_sizes"]), 0))

print("\n%d passed, %d failed" % (sum(results), len(results) - sum(results)))
raise SystemExit(0 if all(results) else 1)
