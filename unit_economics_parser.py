"""
unit_economics_parser.py — Parses the "Unit Economics" tab of a master-planned
community pro-forma model (one workbook = one entity) into structured JSON.

The tab holds one ~49-row block per section (Revenues / Costs line items across
To Date, Remaining, Total, $/FF, $/Lot, $/Acre, % of Costs, % of Rev), an
"Additional Info" panel per section (front feet, acreage, lots, land purchase),
per-phase stats tables, and an entity-wide "Community Rollup" block.

Everything is located by label, never by fixed row/column: sections get added,
removed, and re-phased between model versions. Only section-level data and the
entity rollup are trusted from the workbook — phase and cross-entity rollups
are recomputed by the app (verified to match the model's own rollups exactly).
Of those blocks, only the dollar columns are trusted: $/unit and % columns are
restated from each block's own Totals and units (see restate_units).

Layout facts the parser relies on (stable across model versions):
  * Section title cell matches "Section N" and the row below it, same column,
    is "Revenues" with "To Date" one cell right ("Total" header distinguishes
    the dollar block from the parallel "(Ks)" block).
  * The section's phase tag ("Phase N") sits in the same row, a few columns
    right of the title.
  * Sub-line-items (e.g. Lot Premiums under "Premiums, Escalations, Fence
    Fees") are indented one level; summary rows (Total, Gross Costs, Net
    Margin, ...) are bold — styles carry the hierarchy, so the parser loads
    the workbook with styles (not read_only).
"""

import io
import re
from datetime import date, datetime
from typing import Any

import openpyxl


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_SECTION_RE = re.compile(r"^Section\s+(\d+)$")
_PHASE_RE = re.compile(r"^Phase\s+(\d+)$")

# Sentinel that terminates a unit-economics block.
_LAST_ROW_LABEL = "net margin"

# Additional Info labels → JSON keys.
_INFO_KEYS = {
    "total front feet": "total_front_feet",
    "total acreage": "total_acreage",
    "total lots": "total_lots",
    "phase lots": "phase_lots",
    "blended price per acre": "blended_price_per_acre",
    "land purchase": "land_purchase",
    "price per acre": "price_per_acre",
    "phase acreage": "phase_acreage",
    "life of project acreage": "life_of_project_acreage",
    "life of project front feet": "life_of_project_front_feet",
    "phase front feet": "phase_front_feet",
}


# The same line item is worded differently across entity models (the
# Dennison model predates GPD's renames). Blends and actuals match rows on
# line_key(); a blended row whose members disagree on wording shows the
# canonical label — the names the TGP Phase 2+3 Unit Economics Summary
# uses when it blends Dennison and GPD sections into one phase.
_LINE_ALIASES = {
    "commercial site revenues": "commercial pod sales revenues",
    "residential pod revenues": "residential pod sales revenues",
    "residential pods + dc sales revenues": "residential pod sales revenues",
    "dry utilities, mailboxes": "dry utilities, mailboxes, site work",
    "legal": "legal & mud advances",
    "legal & mud/hoa deficit": "legal & mud advances",      # Windrose model
}
_CANONICAL_LABELS = {
    "commercial pod sales revenues": "Commercial Pod Sales Revenues",
    "residential pod sales revenues": "Residential Pod Sales Revenues",
    "dry utilities, mailboxes, site work": "Dry Utilities, Mailboxes, Site Work",
    "legal & mud advances": "Legal & MUD Advances",
}


def line_key(label: str) -> str:
    """Match key for a line item, stable across models' wording."""
    low = re.sub(r"\s+", " ", (label or "").strip().lower())
    low = re.sub(r"\s*,\s*", ", ", low)        # "Mailboxes,Site Work"
    return _LINE_ALIASES.get(low, low)


def _num(val: Any) -> float | int | None:
    """Cell value → number, preserving None (blank) as None."""
    if val is None:
        return None
    if isinstance(val, str):
        s = val.strip()
        if not s:
            return None
        try:
            val = float(s.replace(",", "").replace("$", ""))
        except ValueError:
            return None
    try:
        n = float(val)
    except (TypeError, ValueError):
        return None
    if n == int(n) and abs(n) < 1e15:
        return int(n)
    return round(n, 6)


def _str(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip()


def _date_iso(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.date().isoformat()
    if isinstance(val, date):
        return val.isoformat()
    return str(val).strip()


def _find_sheet(wb, needle: str = "unit economics"):
    for name in wb.sheetnames:
        if needle in name.strip().lower():
            return wb[name]
    return None


# ---------------------------------------------------------------------------
# Block reader
# ---------------------------------------------------------------------------

# Value columns of a Unit Economics block, as offsets right of the label.
_UE_COLS = {"to_date": 1, "remaining": 2, "total": 3, "per_ff": 4, "per_lot": 5,
            "per_acre": 6, "pct_costs": 7, "pct_rev": 8}


def _read_block_rows(ws, header_row: int, label_col: int, max_row: int,
                     cols: dict | None = None) -> list[dict]:
    """Read a unit-economics block starting at its "Revenues" header row.

    Returns the rows in sheet order. Each row carries a `group`:
      revenue / revenue_total / cost / summary
    plus indent + bold flags so the UI can mirror the Excel presentation.
    Stops at "Net Margin" (inclusive) or the first gap after the block.
    `cols` maps value fields to offsets right of the label column (default:
    the Unit Economics tab's layout); fields it leaves out read as blank.
    """
    cols = cols or _UE_COLS
    rows: list[dict] = []
    group = "revenue"
    r = header_row + 1
    blanks = 0
    while r <= max_row:
        cell = ws.cell(row=r, column=label_col)
        label = _str(cell.value)
        if not label:
            blanks += 1
            if blanks >= 3:
                break
            r += 1
            continue
        blanks = 0
        low = label.lower()
        if low == "costs":
            group = "cost"
            r += 1
            continue
        row = {
            "label": label,
            "group": group,
            "indent": 1 if (cell.alignment.indent or 0) >= 1 else 0,
            "bold": bool(cell.font.bold),
        }
        for field in _UE_COLS:
            off = cols.get(field)
            row[field] = _num(ws.cell(row=r, column=label_col + off).value) if off else None
        if group == "revenue" and low == "total":
            row["group"] = "revenue_total"
            group = "cost"          # "Costs" header follows; skip handled above
        elif low in ("gross costs", "gross margin", "operations & overhead",
                     "financing", "net costs", "net margin"):
            row["group"] = "summary"
            group = "summary"
        rows.append(row)
        if low == _LAST_ROW_LABEL:
            break
        r += 1
    return rows


def _read_additional_info(ws, header_row: int, info_col: int, max_row: int) -> dict:
    """Read the label/value pairs under a "Section N | Additional Info" cell."""
    info: dict = {}
    for r in range(header_row + 1, min(header_row + 20, max_row + 1)):
        label = _str(ws.cell(row=r, column=info_col).value).lower()
        if not label:
            continue
        key = _INFO_KEYS.get(label)
        if not key:
            continue
        raw = ws.cell(row=r, column=info_col + 1).value
        info[key] = _str(raw) if key == "land_purchase" else _num(raw)
    return info


def _find_phase_tag(ws, row: int, start_col: int, span: int = 12) -> str:
    """Phase tag ("Phase N") sits a few columns right of the section title."""
    for c in range(start_col + 1, start_col + span + 1):
        v = _str(ws.cell(row=row, column=c).value)
        if _PHASE_RE.match(v):
            return v
    return ""


# ---------------------------------------------------------------------------
# Allocation tab — per-project footprints for Drainage / Plant Facilities
# ---------------------------------------------------------------------------

def _parse_allocation_projects(wb) -> dict:
    """Per-project section footprints from the "Allocation" tab.

    Drainage Projects and Plant Facilities are the two Unit Economics lines
    whose totals come from this tab: each project spreads its cost at a flat
    $/FF over the sections flagged "Yes" in its column, and the tab's
    per-section total is what the Unit Economics tab pulls. Capturing the
    per-project split lets actuals follow each project's own footprint
    instead of the line-level aggregate.

    Returns {"drainage projects": [{name, total, sections: {"<num>": $}}],
             "plant facilities": [...]} — empty when the tab is absent
    (older models fall back to line-level pro-rata).
    """
    ws = None
    for name in wb.sheetnames:
        if name.strip().lower() == "allocation":
            ws = wb[name]
            break
    if ws is None:
        return {}
    max_row, max_col = ws.max_row, min(ws.max_column, 60)
    out = {}
    for r in range(1, max_row + 1):
        for c in range(1, 12):
            # A block's header row reads: ... | Section | Phase | Total Front
            # Feet | Total <X> Allocation | <project names...>
            if (_str(ws.cell(row=r, column=c).value) != "Section"
                    or _str(ws.cell(row=r, column=c + 1).value) != "Phase"):
                continue
            # Which line the block feeds, from the block title a few rows up.
            line = None
            for rr in range(r - 1, max(0, r - 25), -1):
                for cc in range(1, c + 2):
                    t = _str(ws.cell(row=rr, column=cc).value).lower()
                    if t.startswith("drainage project"):
                        line = "drainage projects"
                    elif t.startswith("plant facilities"):
                        line = "plant facilities"
                    if line:
                        break
                if line:
                    break
            if not line:
                continue
            ff_col, tot_col = c + 2, c + 3
            # Project columns run rightward from the total column until blank.
            proj_cols = []
            for pc in range(tot_col + 1, max_col + 1):
                pname = _str(ws.cell(row=r, column=pc).value)
                if not pname:
                    break
                proj_cols.append((pc, pname))
            if not proj_cols:
                continue
            # Per-project $/FF and Total Cost sit in labelled rows above.
            rate_row = cost_row = None
            for rr in range(r - 1, max(0, r - 10), -1):
                lab = _str(ws.cell(row=rr, column=ff_col).value).lower()
                if lab == "$/ff":
                    rate_row = rr
                elif lab == "total cost":
                    cost_row = rr
            projects = [{"name": pname,
                         "rate": _num(ws.cell(row=rate_row, column=pc).value) if rate_row else None,
                         "total": _num(ws.cell(row=cost_row, column=pc).value) if cost_row else None,
                         "sections": {}}
                        for pc, pname in proj_cols]
            blanks = 0
            for rr in range(r + 1, max_row + 1):
                sec = _str(ws.cell(row=rr, column=c).value)
                m = _SECTION_RE.match(sec)
                if not m:
                    if sec.lower() == "sections allocated":
                        continue
                    blanks += 1
                    if blanks >= 2:
                        break
                    continue
                blanks = 0
                ff = _num(ws.cell(row=rr, column=ff_col).value) or 0
                for (pc, _pname), proj in zip(proj_cols, projects):
                    flag = _str(ws.cell(row=rr, column=pc).value).lower()
                    if flag == "yes" and ff and proj["rate"]:
                        proj["sections"][m.group(1)] = round(ff * proj["rate"], 2)
            out[line] = [p for p in projects if p["sections"]]
    return out


# ---------------------------------------------------------------------------
# Returns tab — LP IRR / equity multiple / promote
# ---------------------------------------------------------------------------

_RETURN_LABELS = {"lp irr": "irr", "irr": "irr", "lp equity multiple": "multiple",
                  "multiple": "multiple", "promote": "promote"}


def _parse_returns(wb) -> dict:
    """LP IRR, LP equity multiple and promote from the model's "Returns"
    tab (labels in one column, the value two columns right). {} when the
    tab or the labels are missing."""
    ws = next((wb[n] for n in wb.sheetnames if n.strip().lower() == "returns"), None)
    if ws is None:
        return {}
    out = {}
    for r in range(1, min(ws.max_row, 80) + 1):
        for c in range(1, 6):
            key = _RETURN_LABELS.get(_str(ws.cell(row=r, column=c).value).lower())
            if key and key not in out:
                val = _num(ws.cell(row=r, column=c + 2).value)
                if val is not None:
                    out[key] = val
    return out


# ---------------------------------------------------------------------------
# Pro-forma comparison tables — entity-level statements without sections
# ---------------------------------------------------------------------------

_DATE_RE = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})")
_EXTRA_LABELS = {"total front feet": "front_feet", "total acreage": "acreage",
                 "total lots": "lots", "total projected av": "projected_av",
                 "irr": "irr", "multiple": "multiple", "promote": "promote"}


def _header_field(text: str) -> str | None:
    """Comparison-table column header -> statement field."""
    t = text.strip().lower()
    if t in ("total", "actual + forecast", "actuals + forecast"):
        return "total"
    if t.startswith("to date") or t.startswith("actuals"):
        return "to_date"
    if t in ("remaining", "forecast"):
        return "remaining"
    return None


def parse_comparison_table(file_bytes: bytes) -> dict:
    """Parse a "Pro-Forma Comparison Table" workbook: side-by-side statement
    columns (e.g. "Original Pro Forma - Jordan", "Current Pro-Forma") in the
    Unit Economics line-item format, each with To Date / Remaining / Total,
    $/FF... and an "Acreage & Yield Assumptions" + "Returns" footer.

    A statement column is a "Base Lot Revenue" label whose header row has a
    "$/FF" column; the side comparison blocks (Total / % of Rev / Delta)
    have none and are skipped. Returns {"blocks": [{title, as_of, rows,
    units, returns, projected_av}]}. Raises ValueError when none is found.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    blocks = []
    for ws in wb.worksheets:
        max_row, max_col = ws.max_row, min(ws.max_column, 120)
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                if _str(ws.cell(row=r, column=c).value).lower() != "base lot revenue":
                    continue
                # Header row: the nearest row above with "$/FF" to the right.
                header = None
                for h in range(r - 1, max(0, r - 4), -1):
                    if any(_str(ws.cell(row=h, column=cc).value).lower() == "$/ff"
                           for cc in range(c + 1, c + 12)):
                        header = h
                        break
                if header is None:
                    continue
                cols, as_of = {}, ""
                for cc in range(c + 1, c + 12):
                    text = _str(ws.cell(row=header, column=cc).value)
                    if text.lower() == "$/ff":
                        break
                    field = _header_field(text)
                    if field and field not in cols:
                        cols[field] = cc - c
                        m = _DATE_RE.search(text)
                        if field == "to_date" and m:
                            as_of = "%s-%02d-%02d" % (m.group(3), int(m.group(1)), int(m.group(2)))
                if "total" not in cols:
                    continue
                title = ""
                for h in range(header, max(0, header - 3), -1):
                    t = _str(ws.cell(row=h, column=c).value)
                    if t and t.lower() != "revenues":
                        title = t
                        break
                rows = _read_block_rows(ws, header, c, max_row, cols)
                for row in rows:
                    if row["remaining"] is None and row["total"] is not None and row["to_date"] is not None:
                        row["remaining"] = round(row["total"] - row["to_date"], 6)
                # Footer: units and returns, labels in the same column.
                end = r + len(rows) + 4
                extras = {}
                for rr in range(r, min(max_row, end + 20) + 1):
                    key = _EXTRA_LABELS.get(_str(ws.cell(row=rr, column=c).value).lower())
                    if key and key not in extras:
                        for cc in range(c + 1, c + 4):
                            val = _num(ws.cell(row=rr, column=cc).value)
                            if val is not None:
                                extras[key] = val
                                break
                blocks.append({
                    "sheet": ws.title, "title": title, "as_of": as_of, "rows": rows,
                    "units": {k: extras.get(k) or 0 for k in ("front_feet", "acreage", "lots")},
                    "returns": {k: extras[k] for k in ("irr", "multiple", "promote") if k in extras},
                    "projected_av": extras.get("projected_av"),
                })
    if not blocks:
        raise ValueError("No pro-forma statement columns found (expected a "
                         "\"Base Lot Revenue\" column with a $/FF header)")
    return {"blocks": blocks}


# ---------------------------------------------------------------------------
# Main entry
# ---------------------------------------------------------------------------

def parse_unit_economics(file_bytes: bytes) -> dict:
    """Parse the Unit Economics tab. Raises ValueError on a missing tab or
    if no section blocks are found."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = _find_sheet(wb)
    if ws is None:
        raise ValueError('Workbook has no "Unit Economics" tab')
    max_row = ws.max_row
    max_col = min(ws.max_column, 100)

    # Actuals date: labelled cell near the top ("Actuals Date:").
    actuals_date = ""
    for r in range(1, min(6, max_row) + 1):
        for c in range(1, 12):
            if _str(ws.cell(row=r, column=c).value).lower().startswith("actuals date"):
                actuals_date = _date_iso(ws.cell(row=r, column=c + 1).value)
                break
        if actuals_date:
            break

    # ── Section blocks ──────────────────────────────────────────────────
    # A section anchor is a "Section N" title whose next row (same column)
    # is "Revenues" followed by "To Date", and whose Total header is plain
    # "Total" — that separates the dollar block from the "(Ks)" mirror.
    sections = []
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            m = _SECTION_RE.match(_str(ws.cell(row=r, column=c).value))
            if not m:
                continue
            if _str(ws.cell(row=r + 1, column=c).value).lower() != "revenues":
                continue
            if _str(ws.cell(row=r + 1, column=c + 1).value).lower() != "to date":
                continue
            if _str(ws.cell(row=r + 1, column=c + 3).value).lower() != "total":
                continue
            phase = _find_phase_tag(ws, r, c)
            rows = _read_block_rows(ws, r + 1, c, max_row)
            if not rows:
                continue
            # Additional Info panel: same row, first matching cell right of
            # the block ("Section N | Additional Info").
            info = {}
            for ic in range(c + 9, min(c + 16, max_col + 1)):
                v = _str(ws.cell(row=r, column=ic).value)
                if v.lower().endswith("additional info"):
                    info = _read_additional_info(ws, r, ic, max_row)
                    break
            sections.append({
                "key": f"Section {m.group(1)}",
                "number": int(m.group(1)),
                "phase": phase,
                "phase_num": int(_PHASE_RE.match(phase).group(1)) if _PHASE_RE.match(phase) else 0,
                "info": info,
                "rows": rows,
            })
            break  # one section per row; skip the "(Ks)" mirror block
    if not sections:
        raise ValueError('No section blocks found on the "Unit Economics" tab')
    sections.sort(key=lambda s: s["number"])

    # ── Entity rollup ("Community Rollup" block in the model) ───────────
    # The model's rollup includes to-date history from closed-out sections
    # that no longer appear as blocks, so it is parsed rather than summed.
    entity_rollup = None
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if _str(ws.cell(row=r, column=c).value).lower() != "community rollup":
                continue
            if _str(ws.cell(row=r + 1, column=c).value).lower() == "revenues":
                entity_rollup = _read_block_rows(ws, r + 1, c, max_row)
            break
        if entity_rollup:
            break

    # ── Per-phase stats tables (lots / front feet / acreage) ────────────
    # Parsed for entity-level denominators and validation; phase-level
    # denominators are recomputed from section info at render time.
    phase_stats: dict[str, dict] = {}
    stat_titles = {
        "total lots per phase": "lots",
        "total ff per phase": "front_feet",
        "total acreage per phase": "acreage",
    }
    entity_stats: dict[str, float] = {}
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            key = stat_titles.get(_str(ws.cell(row=r, column=c).value).lower())
            if not key:
                continue
            for rr in range(r + 1, min(r + 16, max_row + 1)):
                label = _str(ws.cell(row=rr, column=c).value)
                val = _num(ws.cell(row=rr, column=c + 1).value)
                if _PHASE_RE.match(label):
                    phase_stats.setdefault(label, {})[key] = val or 0
                elif label.lower().startswith("total"):
                    entity_stats[key] = val or 0
                    break
                elif not label:
                    break

    # Entity-level unit denominators. Life-of-project figures cover sold-out
    # sections too, matching what the model's own rollup divides by.
    info0 = sections[0]["info"] if sections else {}
    entity_units = {
        "front_feet": info0.get("life_of_project_front_feet") or entity_stats.get("front_feet") or 0,
        "acreage": info0.get("life_of_project_acreage") or entity_stats.get("acreage") or 0,
        "lots": entity_stats.get("lots") or sum((s["info"].get("total_lots") or 0) for s in sections),
    }

    return {
        "actuals_date": actuals_date,
        "sections": sections,
        "entity_rollup": entity_rollup,
        "entity_units": entity_units,
        "phase_stats": phase_stats,
        "alloc_projects": _parse_allocation_projects(wb),
        "returns": _parse_returns(wb),
    }


# ---------------------------------------------------------------------------
# Aggregation — blends N section blocks (or N entity rollups) into one block
# ---------------------------------------------------------------------------

def restate_pcts(rows: list[dict]) -> list[dict]:
    """Recompute both % columns in place on the rollup basis: % of Costs is
    each row over Net Costs and % of Rev each row over Total revenue, for
    revenue and cost rows alike. That is how the model's Phase Rollup and
    Community Rollup blocks (and the Dennison model's section blocks)
    compute them; GPD's section blocks divide cost lines by Gross Costs and
    reuse the revenue share as % of Costs, so parsed section values are
    restated to keep every level on one basis."""
    rev_total = next((r["total"] for r in rows if r["group"] == "revenue_total"), None) or 0
    net_costs = next((r["total"] for r in rows if r["group"] == "summary"
                      and line_key(r["label"]) == "net costs"), None) or 0
    for row in rows:
        total = row.get("total") or 0
        row["pct_costs"] = round(total / net_costs, 6) if net_costs else None
        row["pct_rev"] = round(total / rev_total, 6) if rev_total else None
    return rows


def restate_units(rows: list[dict], units: dict) -> list[dict]:
    """Recompute $/FF, $/Lot and $/Acre in place as each row's Total over
    the block's own units — the Phase Rollup method — then both % columns.

    The models' section blocks were built by copying the first section's
    block, and absolute references came along: in GPD, Sections 2-33 divide
    eleven sub-lines' $/Lot and $/Acre by Section 1's 15 lots and 5.74
    acres; in Dennison, Sections 10-12 divide sixteen sub-lines by Section
    9's units. A denominator the block doesn't carry leaves the parsed value
    in place."""
    for unit, field in (("front_feet", "per_ff"), ("lots", "per_lot"), ("acreage", "per_acre")):
        denom = units.get(unit) or 0
        for row in rows:
            if denom:
                row[field] = round((row.get("total") or 0) / denom, 2)
            else:
                row.setdefault(field, None)
    return restate_pcts(rows)


def blend_blocks(row_sets: list[list[dict]], units: dict) -> list[dict]:
    """Sum dollar columns across blocks and recompute per-unit and percentage
    columns against the blended denominators. Row identity is (group,
    line_key); ordering follows the first block, and a row only a later
    block carries (e.g. Dennison's Impact Fee) slots in after the row that
    precedes it in its own block.

    Verified against the model: its phase rollups equal this blend of their
    member sections, and so does the TGP Phase 2 Unit Economics Summary,
    which blends Dennison and GPD sections into one phase.
    """
    order: list[tuple] = []
    merged: dict[tuple, dict] = {}
    labels: dict[tuple, set] = {}
    for rows in row_sets:
        prev = None
        for row in rows:
            key = (row["group"], line_key(row["label"]))
            if key not in merged:
                merged[key] = {
                    "label": row["label"], "group": row["group"],
                    "indent": row["indent"], "bold": row["bold"],
                    "to_date": None, "remaining": None, "total": None,
                }
                order.insert(order.index(prev) + 1 if prev in merged else 0, key)
            labels.setdefault(key, set()).add(row["label"].strip())
            prev = key
            tgt = merged[key]
            for f in ("to_date", "remaining", "total"):
                if row[f] is not None:
                    tgt[f] = (tgt[f] or 0) + row[f]

    out = []
    for key in order:
        row = merged[key]
        if len(labels[key]) > 1:
            row["label"] = _CANONICAL_LABELS.get(key[1], row["label"])
        out.append(row)
    return restate_units(out, units)


def sum_units(infos: list[dict]) -> dict:
    """Per-unit denominators for a set of sections (their info panels)."""
    return {
        "front_feet": sum((i.get("total_front_feet") or 0) for i in infos),
        "acreage": round(sum((i.get("total_acreage") or 0) for i in infos), 4),
        "lots": sum((i.get("total_lots") or 0) for i in infos),
    }
