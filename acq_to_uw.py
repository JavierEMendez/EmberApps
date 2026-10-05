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



def lot_bands(cbas):
    """The ring's lot-width bands, as CBAS actually aggregates them.

    CBAS does not publish a row per front footage across the ring; it buckets
    into five bands (under 40, 40-50, 50-60, 60-70, 70+) because that is the
    grain the underlying survey supports. An earlier version of this module
    read a top-level `lot_widths` that does not exist, so every width came
    back with no market read and the whole market half silently did nothing.

    Working in bands keeps the evidence at the grain it was measured.
    """
    out = []
    for b in ((cbas or {}).get("lot_bands") or []):
        if not _num(b.get("lots")) and not _num(b.get("avg_price")):
            continue
        out.append({
            "label": b.get("label"),
            "min_ff": int(_num(b.get("min_ff"))),
            "max_ff": int(_num(b.get("max_ff"))),
            "lots": int(_num(b.get("lots"))),
            "communities": int(_num(b.get("communities"))),
            "builders": int(_num(b.get("builders"))),
            "avg_price": int(round(_num(b.get("avg_price")))) or None,
            "min_price": int(round(_num(b.get("min_price")))) or None,
            "max_price": int(round(_num(b.get("max_price")))) or None,
            "avg_sqft": int(round(_num(b.get("avg_sqft")))) or None,
            "avg_ppsf": _num(b.get("avg_ppsf")) or None,
            "plans": int(_num(b.get("plans"))),
        })
    return out


def band_for_width(bands, ff):
    """The band a front footage falls in. Bands are [min, max)."""
    for b in bands or []:
        if b["min_ff"] <= ff < b["max_ff"]:
            return b
    return None


def builders_by_band(cbas, bands):
    """{band label: [builder rows]} aggregated across the ring.

    Builder pricing is published per community, so a builder active in six
    communities appears six times. They are merged here, weighting price by
    lot count, because the question a chart answers is what a builder sells
    for in this submarket -- not what they sell for in one subdivision.
    """
    agg = {}
    for c in ((cbas or {}).get("communities") or []):
        for r in (((c.get("detail") or {}).get("builder_lot_widths")) or []):
            ff = _num(r.get("lot_width_ff"))
            band = band_for_width(bands, ff)
            price = _num(r.get("avg_price"))
            if not band or price <= 0:
                continue
            name = str(r.get("name") or "").strip()
            if not name or name.lower() == "builder tbd":
                continue           # CBAS's placeholder for "not yet assigned"
            key = (band["label"], name)
            cur = agg.get(key)
            lots = _num(r.get("lots"))
            if cur is None:
                agg[key] = {"name": name, "lots": lots, "_pw": price * max(lots, 1),
                            "_w": max(lots, 1),
                            "min_price": _num(r.get("min_price")) or price,
                            "max_price": _num(r.get("max_price")) or price,
                            "avg_sqft": _num(r.get("avg_sqft")) or None,
                            "plans": _num(r.get("plans")), "communities": 1}
            else:
                cur["lots"] += lots
                cur["_pw"] += price * max(lots, 1)
                cur["_w"] += max(lots, 1)
                lo = _num(r.get("min_price")) or price
                hi = _num(r.get("max_price")) or price
                cur["min_price"] = min(cur["min_price"], lo)
                cur["max_price"] = max(cur["max_price"], hi)
                cur["plans"] += _num(r.get("plans"))
                cur["communities"] += 1
    out = {}
    for (label, _name), r in agg.items():
        r["avg_price"] = int(round(r.pop("_pw") / r.pop("_w")))
        r["lots"] = int(r["lots"])
        r["min_price"] = int(round(r["min_price"]))
        r["max_price"] = int(round(r["max_price"]))
        r["plans"] = int(r["plans"])
        out.setdefault(label, []).append(r)
    for rows in out.values():
        rows.sort(key=lambda x: -x["lots"])
    return out


def addressable_pace(cbas, capture_pct):
    """Lots per month for the whole project, from the ring's addressable starts.

    The CBAS endpoint already works this out properly and the reasoning is its
    own: you do not compete for every start in the ring, only for starts in
    the widths you intend to build, so ring starts are apportioned by the
    share of ring lots sitting in the project's target bands. Capture of that
    addressable figure is the number worth arguing about -- capture of all
    starts understates it whenever a project targets part of the range.

    Returns (lots_per_month, note) or (None, None).
    """
    cap = (((cbas or {}).get("market_entry") or {}).get("capture")) or {}
    addressable = _num(cap.get("addressable_starts"))
    if addressable <= 0:
        return None, None
    pace = addressable * capture_pct / 12.0
    note = ("%.2f lots/mo = %.0f addressable starts a year x %.0f%% capture / 12. "
            "Addressable is the ring's %.0f starts apportioned to the %s FF this "
            "project targets. The median community in the ring runs %.1f%% share; "
            "the 75th runs %.1f%%."
            % (pace, addressable, capture_pct * 100,
               _num(cap.get("ring_annual_starts")),
               ", ".join(str(int(f)) for f in (cap.get("target_ff") or [])) or "targeted",
               _num(cap.get("share_median_pct")), _num(cap.get("share_p75_pct"))))
    return pace, note


def mix_by_width(analysis):
    """{front footage: yield-mix row} from the analyst's chosen product mix."""
    out = {}
    br = (((analysis or {}).get("yield_estimates") or {}).get("breakdown")) or []
    for row in br:
        ff = _ff_of(row.get("label"))
        if ff:
            out[ff] = row
    return out




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

    bands = lot_bands(cbas)
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

    # ---- pace ------------------------------------------------------------
    # One project-level absorption, split across the widths being built in
    # proportion to how much of the mix each carries. Pace is a claim about
    # the whole community competing for the ring's starts; apportioning it is
    # honest, whereas giving every width the ring's full capture would
    # multiply the project's absorption by the number of products in it.
    total_pace, pace_note = addressable_pace(cbas, capture_pct)
    alloc_total = sum(max(0.0, _num(r.get("allocation_pct"))) for r in mix.values()) or 0.0

    rows = [dict(r) for r in (inputs.get("lot_sizes") or [])]
    priced = []
    for i, row in enumerate(rows):
        ff = int(_num(row.get("front_footage"),
                      UW_LOT_WIDTHS[i] if i < len(UW_LOT_WIDTHS) else 0))
        x = mix.get(ff)
        band = band_for_width(bands, ff)

        row["on"] = 1 if x else 0
        if x:
            upa = _num(x.get("units_per_acre"))
            if upa > 0:
                row["yield_per_ac"] = round(upa, 2)
                basis["lot_sizes.%d.yield_per_ac" % i] = (
                    "%.2f u/ac from the acquisition product mix." % upa)

            if total_pace and alloc_total > 0:
                share = max(0.0, _num(x.get("allocation_pct"))) / alloc_total
                if share > 0:
                    row["pace"] = round(total_pace * share, 2)
                    basis["lot_sizes.%d.pace" % i] = (
                        "%.2f lots/mo = %.0f%% of the project's %.2f lots/mo. %s"
                        % (row["pace"], share * 100, total_pace, pace_note or ""))

        # Home price is NOT written. It prices every home in the deal and
        # feeds assessed value straight into MUD capacity, so it stays the
        # underwriter's call -- the evidence for it is assembled separately.
        if row["on"] and band and band.get("avg_price"):
            priced.append((ff, max(_num(x.get("allocation_pct")), 1.0),
                           band["avg_price"]))
        rows[i] = row
    inputs["lot_sizes"] = rows

    basis["_settings"] = {"capture_pct": capture_pct, "lot_ratio": lot_ratio}
    basis["_suggested_price_per_ff"] = blended_price_per_ff(priced, lot_ratio)
    basis["_project_pace"] = round(total_pace, 2) if total_pace else None
    basis["_widths_matched"] = sorted(w for w in mix if band_for_width(bands, w))
    basis["_widths_no_market"] = sorted(w for w in mix if not band_for_width(bands, w))
    return inputs, basis


def market_evidence(cbas, mix_widths=None, lot_ratio=LOT_RATIO_DEFAULT):
    """What the submarket shows about pricing, shaped for a chart.

    Deliberately evidence and not a decision. Each band carries the average
    the market is achieving, the range behind that average, and the builders
    making it up -- because an average over two builders and an average over
    nine are not the same claim, and a band whose range is $180k wide is not
    really one price at all.

    Bands, not front footages: that is the grain CBAS aggregates the ring at,
    and showing a per-FF number would imply a precision the survey does not
    carry.
    """
    bands = lot_bands(cbas)
    builders = builders_by_band(cbas, bands)
    mix_widths = set(mix_widths or [])

    rows = []
    for b in bands:
        home = _num(b.get("avg_price"))
        # Price a band at the midpoint of the widths it covers, which is what
        # a $/FF derived from it actually describes.
        mid = (b["min_ff"] + min(b["max_ff"], b["min_ff"] + 20)) / 2.0
        in_mix = sorted(w for w in mix_widths if b["min_ff"] <= w < b["max_ff"])
        rows.append({
            "label": b["label"],
            "min_ff": b["min_ff"], "max_ff": b["max_ff"],
            "mid_ff": round(mid, 1),
            "in_mix": bool(in_mix),
            "mix_widths": in_mix,
            "lots": b["lots"],
            "communities": b["communities"],
            "avg_price": b["avg_price"],
            "min_price": b["min_price"],
            "max_price": b["max_price"],
            "avg_sqft": b["avg_sqft"],
            "avg_ppsf": b["avg_ppsf"],
            "implied_lot_value": int(round(home * lot_ratio)) if home > 0 else None,
            "implied_lot_ff": round(home * lot_ratio / mid, 2) if home > 0 and mid else None,
            "builders": builders.get(b["label"], [])[:10],
            "builder_count": len(builders.get(b["label"], [])),
        })

    cap = (((cbas or {}).get("market_entry") or {}).get("capture")) or {}
    in_mix_rows = [(r["mid_ff"], r["lots"], r["avg_price"]) for r in rows
                   if r["in_mix"] and r["avg_price"]]
    return {
        "lot_ratio": lot_ratio,
        "bands": rows,
        "capture": {
            "ring_annual_starts": _num(cap.get("ring_annual_starts")) or None,
            "addressable_starts": _num(cap.get("addressable_starts")) or None,
            "active_communities": _num(cap.get("active_communities")) or None,
            "share_median_pct": _num(cap.get("share_median_pct")) or None,
            "share_p75_pct": _num(cap.get("share_p75_pct")) or None,
            "target_ff": cap.get("target_ff") or [],
        },
        "suggested_price_per_ff": (
            round(blended_price_per_ff(in_mix_rows, lot_ratio), 2)
            if in_mix_rows else None),
        "suggested_basis": (
            "Blended over the %s bands your mix falls in, weighted by lot count "
            "and frontage so total lot revenue matches pricing each band on its "
            "own." % ", ".join(r["label"] for r in rows if r["in_mix"])
            if in_mix_rows else None),
    }
