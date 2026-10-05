"""Turn an acquisition analysis into MPC Underwriting inputs.

The two halves of the portal already know different things about the same
deal. The acquisition side knows the dirt: how many acres, what comes off for
floodplain and easements, what product mix the analyst chose. CBAS knows the
submarket: who is building what width, how fast it starts, what the homes
sell for. The underwriting model needs all of it, and until now it was
re-keyed by hand.

Everything here is a pure function over dicts -- no Flask, no database, no
network -- so the arithmetic that ends up driving a land offer can be tested
without a browser.

Two conventions are Carlos's, not defaults picked here:

  * Pace comes off STARTS, not closings. Starts lead closings and are the
    better read on what a market is absorbing right now.
  * Finished lot value is a share of home price (22%), divided by front
    footage to reach the $/FF the model prices lots in.

Pricing is the one thing this module will not decide. Pace, yield, acreage
and constraints are measurements and carry straight over; a price is a
judgement about a market, and silently writing one into a model that prices
3,000 lots is the wrong kind of automation. So `market_evidence` assembles
what the submarket actually shows -- price by width, by builder, with the
range -- and the underwriter types the number. The model's own defaults stand
until they do.

What this module deliberately does NOT do: it never touches the development
programme. Plants, amenities, detention, roads and parks are the
underwriter's netouts and the model computes acreage for them itself. The
only acreage handed over is gross, plus the physical constraints the GIS
measured -- floodplain, wetlands, easements -- which land in `other_netouts`
where they belong. Feeding net developable into a pod field would have
double-counted every constraint.
"""

import re

# Share of the submarket's annual starts a new community is assumed to take.
# A starting point, not a finding -- every deal argues its own capture.
CAPTURE_PCT_DEFAULT = 0.20

# Finished lot value as a share of the home price it carries.
LOT_RATIO_DEFAULT = 0.22

# The model's lot table is a fixed 16 rows, 25 FF to 100 FF in fives.
UW_LOT_WIDTHS = [25 + 5 * i for i in range(16)]


def _num(v, default=0.0):
    try:
        f = float(v)
        return default if f != f else f           # NaN
    except (TypeError, ValueError):
        return default


def _ff_of(label):
    """Front footage out of a product label like '40 FF' or '40FF'."""
    m = re.search(r"(\d+(?:\.\d+)?)", str(label or ""))
    return int(float(m.group(1))) if m else None


def _nearest_width(ff, available):
    """The modelled width closest to `ff`.

    CBAS reports what builders actually plat -- 46 FF, 52 FF -- while the
    model's table moves in fives. Snapping to the nearest row keeps a real
    market width from being dropped on the floor for not being round.
    """
    if not available:
        return None
    return min(available, key=lambda w: (abs(w - ff), w))


def market_by_width(cbas):
    """{front footage: CBAS row}, snapped onto the model's lot table."""
    out = {}
    for row in ((cbas or {}).get("lot_widths") or []):
        ff = _num(row.get("lot_width_ff"))
        if ff <= 0:
            continue
        w = _nearest_width(ff, UW_LOT_WIDTHS)
        if w is None:
            continue
        # Two CBAS widths can snap to one model row (46 and 47 both land on
        # 45). Merge them: lots add, prices weight by lot count.
        prev = out.get(w)
        if prev is None:
            out[w] = dict(row, _snapped_from=[ff])
            continue
        a, b = _num(prev.get("lots")), _num(row.get("lots"))
        tot = a + b
        for k in ("avg_price", "avg_sqft", "avg_ppsf"):
            pa, pb = _num(prev.get(k)), _num(row.get(k))
            if tot > 0 and (pa or pb):
                prev[k] = round((pa * a + pb * b) / tot, 2)
        for k in ("lots", "est_vdls", "est_futures",
                  "est_annual_starts", "est_annual_closings"):
            prev[k] = round(_num(prev.get(k)) + _num(row.get(k)), 1)
        prev["_snapped_from"].append(ff)
    return out


def mix_by_width(analysis):
    """{front footage: yield-mix row} from the analyst's chosen product mix."""
    out = {}
    br = (((analysis or {}).get("yield_estimates") or {}).get("breakdown")) or []
    for row in br:
        ff = _ff_of(row.get("label"))
        if ff:
            out[ff] = row
    return out


def recommended_pace(mkt_row, capture_pct):
    """Lots per month for one width, off the submarket's annual starts."""
    starts = _num((mkt_row or {}).get("est_annual_starts"))
    if starts <= 0:
        return None
    return starts * capture_pct / 12.0


def lot_value(mkt_row, lot_ratio):
    """Finished lot value implied by the home price that sits on it."""
    home = _num((mkt_row or {}).get("avg_price"))
    return home * lot_ratio if home > 0 else None


def blended_price_per_ff(priced_widths, lot_ratio):
    """One $/FF across every width, holding total lot revenue right.

    The model prices a lot as front_footage x $/FF with a single rate for the
    whole project, but a 40 FF lot and an 80 FF lot do not carry the same
    $/FF in the market. Averaging the rates would misprice the deal by
    whatever the mix is skewed toward; weighting by lot count and footage
    gives the rate whose revenue matches pricing each width on its own.

    `priced_widths` is (front_footage, lots, home_price) per width.
    """
    revenue = frontage = 0.0
    for ff, lots, home in priced_widths:
        if ff <= 0 or lots <= 0 or home <= 0:
            continue
        revenue += lots * home * lot_ratio
        frontage += lots * ff
    return (revenue / frontage) if frontage > 0 else None


def constraint_netouts(analysis, slots=6):
    """The GIS's measured constraints, shaped for the model's other_netouts.

    Marginal acreage, not footprint: overlapping layers are deducted once on
    the acquisition side, and carrying the footprints across would take a
    wetland sitting inside a floodplain out of the deal twice.
    """
    rows = []
    for d in ((analysis or {}).get("netout_detail") or []):
        if not d.get("applied") or d.get("error"):
            continue
        marg = d.get("acres_marginal")
        acres = _num(marg if marg is not None else d.get("acres"))
        if acres <= 0:
            continue
        rows.append({
            "desc": str(d.get("label") or d.get("key") or "Constraint")[:60],
            "acres": round(acres, 2),
            "notes": ("stated" if d.get("stated") else "measured by GIS"),
        })
    rows.sort(key=lambda r: -r["acres"])
    if len(rows) > slots:
        # More layers than the model has rows: keep the largest and roll the
        # tail into the last slot rather than silently dropping acreage.
        head, tail = rows[:slots - 1], rows[slots - 1:]
        head.append({"desc": "Other constraints (%d layers)" % len(tail),
                     "acres": round(sum(r["acres"] for r in tail), 2),
                     "notes": "measured by GIS"})
        rows = head
    while len(rows) < slots:
        rows.append({"desc": "", "acres": 0, "notes": ""})
    return rows


def derive_uw_inputs(base, analysis, cbas=None, *,
                     capture_pct=CAPTURE_PCT_DEFAULT,
                     lot_ratio=LOT_RATIO_DEFAULT):
    """Fill a default underwriting input set from the acquisition work.

    `base` is the model's own defaults (app.default_inputs), passed in rather
    than imported so this module stays free of the Flask app. Every field the
    acquisition side cannot speak to is left exactly as the model set it.

    Returns (inputs, basis). `basis` explains, in a sentence per field, where
    each derived number came from -- so an underwriter overriding one can see
    what they are overruling.
    """
    inputs = dict(base or {})
    basis = {}
    analysis = analysis or {}

    mkt = market_by_width(cbas)
    mix = mix_by_width(analysis)

    # ---- acreage ---------------------------------------------------------
    gross = _num(analysis.get("gross_acres"))
    if gross > 0:
        inputs["gross_acreage"] = round(gross, 2)
        src = {"override": "your acreage override",
               "stated": "the appraisal district",
               "measured": "the measured boundary"}.get(
                   analysis.get("gross_basis"), "the acquisition analysis")
        basis["gross_acreage"] = "%.2f ac per %s." % (gross, src)

    netouts = constraint_netouts(analysis)
    if any(r["acres"] for r in netouts):
        inputs["other_netouts"] = netouts
        named = ", ".join("%s %.1f ac" % (r["desc"], r["acres"])
                          for r in netouts if r["acres"])
        basis["other_netouts"] = (
            "Physical constraints measured by the GIS: %s. Marginal acreage, so "
            "overlapping layers are deducted once. The development programme "
            "(plants, amenities, detention, roads, parks) is still yours." % named)

    # ---- lot table -------------------------------------------------------
    rows = [dict(r) for r in (inputs.get("lot_sizes") or [])]
    priced = []
    for i, row in enumerate(rows):
        ff = int(_num(row.get("front_footage"),
                      UW_LOT_WIDTHS[i] if i < len(UW_LOT_WIDTHS) else 0))
        m, x = mkt.get(ff), mix.get(ff)

        # The analyst's mix decides what is built. The market decides how fast
        # it sells and what it sells for.
        row["on"] = 1 if x else 0
        if x:
            upa = _num(x.get("units_per_acre"))
            if upa > 0:
                row["yield_per_ac"] = round(upa, 2)
                basis["lot_sizes.%d.yield_per_ac" % i] = (
                    "%.2f u/ac from the acquisition product mix." % upa)

        if m:
            pace = recommended_pace(m, capture_pct)
            if pace:
                row["pace"] = round(pace, 2)
                basis["lot_sizes.%d.pace" % i] = (
                    "%.2f lots/mo = %.0f %d FF starts a year in the submarket "
                    "x %.0f%% capture / 12." % (pace, _num(m.get("est_annual_starts")),
                                                ff, capture_pct * 100))
            # Home price is NOT written. It prices every home in the deal and
            # feeds assessed value straight into MUD capacity, so it stays the
            # underwriter's call -- the evidence for it is assembled below.
            home = _num(m.get("avg_price"))
            if row["on"] and home > 0:
                priced.append((ff, _num(m.get("lots")), home))
        rows[i] = row
    inputs["lot_sizes"] = rows

    basis["_settings"] = {"capture_pct": capture_pct, "lot_ratio": lot_ratio}
    basis["_suggested_price_per_ff"] = blended_price_per_ff(priced, lot_ratio)
    basis["_widths_matched"] = sorted(set(mix) & set(mkt))
    basis["_widths_no_market"] = sorted(set(mix) - set(mkt))
    return inputs, basis


def market_evidence(cbas, mix_widths=None, lot_ratio=LOT_RATIO_DEFAULT):
    """What the submarket shows about pricing, shaped for a chart.

    Deliberately evidence and not a decision. Each width carries the average
    the market is achieving, the range behind that average, and the builders
    making it up -- because an average over two builders and an average over
    nine are not the same claim, and a width whose range is $180k wide is not
    really one price at all.

    `implied_lot_ff` is what the finished-lot share works out to per width. It
    is shown beside the home price rather than applied, so the underwriter can
    see whether one width is carrying the blend.
    """
    widths = []
    by_width = market_by_width(cbas)
    builder_rows = ((cbas or {}).get("builder_lot_widths") or [])

    for ff in sorted(by_width):
        m = by_width[ff]
        home = _num(m.get("avg_price"))
        lots = _num(m.get("lots"))
        builders = []
        for b in builder_rows:
            bff = _num(b.get("lot_width_ff"))
            if not bff or _nearest_width(bff, UW_LOT_WIDTHS) != ff:
                continue
            bp = _num(b.get("avg_price"))
            if bp <= 0:
                continue
            builders.append({
                "name": str(b.get("name") or "Builder")[:40],
                "lots": int(_num(b.get("lots"))),
                "avg_price": int(round(bp)),
                "min_price": int(round(_num(b.get("min_price")))) or None,
                "max_price": int(round(_num(b.get("max_price")))) or None,
                "avg_sqft": int(round(_num(b.get("avg_sqft")))) or None,
                "plans": int(_num(b.get("plans"))),
            })
        builders.sort(key=lambda r: -r["lots"])
        widths.append({
            "ff": ff,
            "in_mix": bool(mix_widths and ff in mix_widths),
            "lots": int(lots),
            "avg_price": int(round(home)) if home > 0 else None,
            "min_price": int(round(_num(m.get("min_price")))) or None,
            "max_price": int(round(_num(m.get("max_price")))) or None,
            "avg_sqft": int(round(_num(m.get("avg_sqft")))) or None,
            "avg_ppsf": _num(m.get("avg_ppsf")) or None,
            "annual_starts": _num(m.get("est_annual_starts")) or None,
            "vdls": int(_num(m.get("est_vdls"))) or None,
            "futures": int(_num(m.get("est_futures"))) or None,
            "implied_lot_value": int(round(home * lot_ratio)) if home > 0 else None,
            "implied_lot_ff": round(home * lot_ratio / ff, 2) if home > 0 and ff else None,
            "builders": builders[:10],
            "builder_count": len(builders),
            "snapped_from": sorted(set(m.get("_snapped_from") or [])),
        })

    in_mix = [(w["ff"], w["lots"], w["avg_price"]) for w in widths
              if w["in_mix"] and w["avg_price"]]
    return {
        "lot_ratio": lot_ratio,
        "widths": widths,
        # The blend over the widths actually being built, which is the number
        # that would hold total lot revenue if it were adopted.
        "suggested_price_per_ff": (
            round(blended_price_per_ff(in_mix, lot_ratio), 2)
            if in_mix else None),
        "suggested_basis": (
            "Blended over the %s FF in the product mix, weighted by lot count "
            "and frontage so total lot revenue matches pricing each width on "
            "its own." % ", ".join(str(ff) for ff, _l, _p in sorted(in_mix))
            if in_mix else None),
    }
