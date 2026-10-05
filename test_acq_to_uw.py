"""Tests for the acquisition -> underwriting derivation.

These numbers end up in a land offer, so the arithmetic is pinned here rather
than checked by eye in the browser. Run: python test_acq_to_uw.py

CBAS below is the shape the endpoint actually returns. The first version of
this module read a top-level "lot_widths" that does not exist, so every width
came back with no market read and the market half silently did nothing while
reporting success. These tests exist so that cannot happen quietly again.
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
    numeric = isinstance(want, (int, float)) and not isinstance(want, bool)
    ok = (abs(got - want) <= tol) if (numeric and isinstance(got, (int, float))) \
        else got == want
    print(("  PASS  " if ok else "  FAIL  ") + name)
    if not ok:
        print("          got %r, want %r" % (got, want))
    return ok


results = []

B40 = "40–50 FF"
B50 = "50–60 FF"

CBAS = {
    "lot_bands": [
        {"label": "Under 40 FF", "min_ff": 20, "max_ff": 40, "lots": 310,
         "communities": 4, "builders": 3, "avg_price": 289000, "min_price": 241000,
         "max_price": 342000, "avg_sqft": 1650, "avg_ppsf": 175.2, "plans": 12},
        {"label": B40, "min_ff": 40, "max_ff": 50, "lots": 820,
         "communities": 9, "builders": 6, "avg_price": 318000, "min_price": 264000,
         "max_price": 402000, "avg_sqft": 1830, "avg_ppsf": 173.8, "plans": 24},
        {"label": B50, "min_ff": 50, "max_ff": 60, "lots": 610,
         "communities": 7, "builders": 5, "avg_price": 398000, "min_price": 331000,
         "max_price": 489000, "avg_sqft": 2240, "avg_ppsf": 177.7, "plans": 19},
        {"label": "60–70 FF", "min_ff": 60, "max_ff": 70, "lots": 300,
         "communities": 5, "builders": 4, "avg_price": 489000, "min_price": 412000,
         "max_price": 625000, "avg_sqft": 2810, "avg_ppsf": 174.0, "plans": 14},
        {"label": "70+ FF", "min_ff": 70, "max_ff": 200, "lots": 90,
         "communities": 2, "builders": 2, "avg_price": 612000, "min_price": 540000,
         "max_price": 735000, "avg_sqft": 3300, "avg_ppsf": 185.5, "plans": 6},
    ],
    "market_entry": {"capture": {
        "ring_annual_starts": 384, "active_communities": 9,
        "share_median_pct": 9.4, "share_p75_pct": 14.8, "share_max_pct": 22.0,
        "segment_share_pct": 78.0, "addressable_starts": 300,
        "target_ff": [40, 50], "target_bands": [B40, B50],
    }},
    "communities": [
        {"name": "Community A", "detail": {"lot_widths": [
            {"lot_width_ff": 45, "lots": 200, "avg_price": 309000, "min_price": 264000,
             "max_price": 358000, "avg_sqft": 1790, "plans": 7},
            {"lot_width_ff": 42, "lots": 120, "avg_price": 331000, "min_price": 290000,
             "max_price": 372000, "avg_sqft": 1850, "plans": 4},
            {"lot_width_ff": 55, "lots": 280, "avg_price": 412000, "min_price": 352000,
             "max_price": 489000, "avg_sqft": 2310, "plans": 9},
            {"lot_width_ff": 95, "lots": 40, "avg_price": 712000, "min_price": 640000,
             "max_price": 820000, "avg_sqft": 3600, "plans": 3},
        ], "builder_lot_widths": [
            {"name": "Lennar", "lot_width_ff": 45, "lots": 200, "avg_price": 309000,
             "min_price": 264000, "max_price": 358000, "avg_sqft": 1790, "plans": 7},
            {"name": "Builder TBD", "lot_width_ff": 45, "lots": 60, "avg_price": 300000},
        ]}},
        {"name": "Community B", "detail": {"builder_lot_widths": [
            {"name": "Lennar", "lot_width_ff": 42, "lots": 100, "avg_price": 333000,
             "min_price": 290000, "max_price": 380000, "avg_sqft": 1820, "plans": 4},
            {"name": "Perry Homes", "lot_width_ff": 55, "lots": 280, "avg_price": 412000,
             "min_price": 352000, "max_price": 489000, "avg_sqft": 2310, "plans": 9},
        ]}},
    ],
}

ANALYSIS_NETOUTS = [
    {"label": "Floodplain (100-yr)", "applied": True, "acres": 29.0, "acres_marginal": 29.0},
    {"label": "Wetlands (NWI)", "applied": True, "acres": 10.9, "acres_marginal": 9.9},
    {"label": "Stream buffer", "applied": True, "acres": 12.5, "acres_marginal": 5.8},
    {"label": "Not applied", "applied": False, "acres": 50.0, "acres_marginal": 50.0},
    {"label": "Layer failed", "applied": True, "error": True, "acres": 40.0},
]

# --- buckets are decades, built from real widths ---------------------------
# Builders talk in 40s, 50s, 60s. CBAS's own five bands ("40-50 FF") read as
# ranges that straddle the products a mix is built from, and a 45 is not a
# category of its own -- it is a 40.
results.append(check("45 is a 40", A.bucket_of(45)[1], "40 FF"))
results.append(check("47 is a 40", A.bucket_of(47)[1], "40 FF"))
results.append(check("50 starts the 50s", A.bucket_of(50)[1], "50 FF"))
results.append(check("55 is a 50", A.bucket_of(55)[1], "50 FF"))
results.append(check("80 is its own bucket", A.bucket_of(80)[1], "80 FF"))
results.append(check("everything under 40 groups", A.bucket_of(32)[1], "Under 40 FF"))
results.append(check("90 and up group", A.bucket_of(120)[1], "90+ FF"))
results.append(check("a zero width has no bucket", A.bucket_of(0), None))

bands = A.lot_bands(CBAS)
labels = [b["label"] for b in bands]
results.append(check("buckets come from the community detail, not the ring bands",
                     labels, ["40 FF", "50 FF", "90+ FF"]))
b40 = next(b for b in bands if b["label"] == "40 FF")
results.append(check("the 45s and 42s merge into the 40s", b40["lots"], 320))
results.append(check("the real widths behind a bucket are kept", b40["widths"], [42, 45]))
results.append(check("bucket price weights by lot count",
                     b40["avg_price"], round((309000 * 200 + 331000 * 120) / 320)))
results.append(check("bucket range spans its widths",
                     (b40["min_price"], b40["max_price"]), (264000, 372000)))
results.append(check("band_for_width finds the bucket",
                     A.band_for_width(bands, 47)["label"], "40 FF"))
results.append(check("no bands means no match", A.band_for_width([], 45), None))

# a payload with no community detail still charts off CBAS's own bands
legacy = A.lot_bands({"lot_bands": CBAS["lot_bands"]})
results.append(check("falls back to the ring bands when there is no detail",
                     len(legacy), 5))

# --- builders are merged across the ring, not counted per community --------
bb = A.builders_by_band(CBAS, bands)
lennar = next(r for r in bb["40 FF"] if r["name"] == "Lennar")
results.append(check("a builder in two communities is merged once",
                     lennar["communities"], 2))
results.append(check("merged builder lots add", lennar["lots"], 300))
results.append(check("merged builder price weights by lot count",
                     lennar["avg_price"], round((309000 * 200 + 333000 * 100) / 300)))
results.append(check("merged range spans both communities",
                     (lennar["min_price"], lennar["max_price"]), (264000, 380000)))
results.append(check("CBAS's placeholder builder is not a competitor",
                     any(r["name"] == "Builder TBD" for r in bb.get("40 FF", [])), False))
results.append(check("a builder lands in the band its width falls in",
                     [r["name"] for r in bb["50 FF"]], ["Perry Homes"]))

# --- pace off addressable starts, the way the endpoint models it -----------
pace, note = A.addressable_pace(CBAS, 0.20)
results.append(check("project pace = addressable starts x capture / 12",
                     pace, 300 * 0.20 / 12))
results.append(check("pace scales with capture",
                     A.addressable_pace(CBAS, 0.40)[0], 300 * 0.40 / 12))
results.append(check("the basis names the ring and the observed percentiles",
                     ("addressable" in note and "9.4" in note and "14.8" in note), True))
results.append(check("no addressable starts means no pace",
                     A.addressable_pace({}, 0.2)[0], None))

# --- the blend holds total revenue -----------------------------------------
priced = [(45, 820, 318000), (55, 610, 398000), (65, 300, 489000)]
ratio = 0.22
blend = A.blended_price_per_ff(priced, ratio)
rev_individual = sum(lots * home * ratio for _ff, lots, home in priced)
rev_blended = sum(lots * ff * blend for ff, lots, _home in priced)
results.append(check("blended $/FF preserves total lot revenue",
                     rev_blended, rev_individual, tol=0.01))
results.append(check("blend sits between the per-band rates",
                     min(h * ratio / f for f, _l, h in priced) <= blend
                     <= max(h * ratio / f for f, _l, h in priced), True))

# --- constraints carry marginal acreage, never footprints ------------------
no = A.constraint_netouts({"netout_detail": ANALYSIS_NETOUTS})
results.append(check("marginal acreage, not footprint",
                     round(sum(r["acres"] for r in no), 2), 44.7))
results.append(check("kept-by-override layers are excluded",
                     any(r["desc"] == "Not applied" for r in no), False))
results.append(check("failed layers are excluded",
                     any(r["desc"] == "Layer failed" for r in no), False))
results.append(check("always fills the model's six slots", len(no), 6))
rolled = A.constraint_netouts({"netout_detail": [
    {"label": "L%d" % i, "applied": True, "acres": 10 - i, "acres_marginal": 10 - i}
    for i in range(8)]})
results.append(check("overflow rolls up rather than dropping acreage",
                     round(sum(r["acres"] for r in rolled), 2),
                     round(sum(10 - i for i in range(8)), 2)))

# --- end to end -------------------------------------------------------------
full = {"gross_acres": 207.5, "gross_basis": "stated",
        "netout_detail": ANALYSIS_NETOUTS,
        "yield_estimates": {"breakdown": [
            {"label": "40 FF", "units_per_acre": 6.0, "allocation_pct": 50},
            {"label": "50 FF", "units_per_acre": 5.0, "allocation_pct": 50}]}}
inp, bas = A.derive_uw_inputs(_base(), full, CBAS)
on = {r["front_footage"]: r for r in inp["lot_sizes"] if r["on"]}
results.append(check("only the analyst's mix is switched on", sorted(on), [40, 50]))
results.append(check("yield comes from the mix", on[40]["yield_per_ac"], 6.0))
results.append(check("gross acreage carries over", inp["gross_acreage"], 207.5))
results.append(check("constraints carry over",
                     round(sum(r["acres"] for r in inp["other_netouts"]), 2), 44.7))

# One community's absorption, split across its products -- not repeated per product.
proj_pace = 300 * 0.20 / 12
results.append(check("pace is apportioned by mix allocation",
                     on[40]["pace"], round(proj_pace * 0.5, 2)))
results.append(check("the widths' paces sum to the project's",
                     round(on[40]["pace"] + on[50]["pace"], 2), round(proj_pace, 2)))
results.append(check("project pace is reported", bas["_project_pace"], round(proj_pace, 2)))
results.append(check("every mix width found a band", bas["_widths_no_market"], []))

# Pricing is the underwriter's call. Nothing may quietly set it.
results.append(check("home price is NOT applied", on[40]["home_price"], 200000))
results.append(check("$/FF is NOT applied", inp["price_per_ff"]["0"], 1800))

# --- evidence ---------------------------------------------------------------
ev = A.market_evidence(CBAS, mix_widths={40, 50}, lot_ratio=0.22)
ev40 = next(b for b in ev["bands"] if b["label"] == "40 FF")
results.append(check("evidence is bucketed by decade", [b["label"] for b in ev["bands"]], ["40 FF", "50 FF", "90+ FF"]))
results.append(check("the mix's buckets are flagged", ev40["in_mix"], True))
results.append(check("a bucket outside the mix is not flagged",
                     next(b for b in ev["bands"] if b["label"] == "90+ FF")["in_mix"], False))
results.append(check("builders come through on the bucket",
                     [r["name"] for r in ev40["builders"]], ["Lennar"]))
results.append(check("implied lot $/FF prices off the real average frontage",
                     ev40["implied_lot_ff"],
                     round(ev40["avg_price"] * 0.22 / 43.5, 2)))
results.append(check("the capture block travels with the evidence",
                     ev["capture"]["addressable_starts"], 300))
results.append(check("a suggestion is offered", ev["suggested_price_per_ff"] > 0, True))

# --- nothing to go on: the model's defaults must come through untouched -----
empty_in, _ = A.derive_uw_inputs(_base(), {}, {})
results.append(check("no analysis leaves gross at the model default",
                     empty_in["gross_acreage"], 0))
results.append(check("no market leaves $/FF at the model default",
                     empty_in["price_per_ff"]["0"], 1800))
results.append(check("no mix switches nothing on",
                     sum(r["on"] for r in empty_in["lot_sizes"]), 0))
results.append(check("no market means no evidence",
                     A.market_evidence({}, {40}, 0.22)["suggested_price_per_ff"], None))

print("\n%d passed, %d failed" % (sum(results), len(results) - sum(results)))
raise SystemExit(0 if all(results) else 1)
