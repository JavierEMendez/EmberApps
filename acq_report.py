"""Executive acquisition report — data assembly and figure rendering.

The report is HTML/CSS rendered by WeasyPrint, not hand-positioned on a PDF
canvas. That decision was already made in this codebase: WeasyPrint is in
requirements.txt and nixpacks.toml already installs Pango, Cairo and the
Liberation fonts on Railway for the Ember Capital Report, which ships the same
way. Following it means the layout lives in a template anyone can edit rather
than in a thousand lines of absolute coordinates.

Nothing here recomputes the acquisition analysis. Every number comes from the
endpoints that already feed the project page; this module calls those view
functions inside the request context and reshapes what they return. If a figure
looks wrong, it is wrong on the page too.

Figures (site map, competitor map, trend charts) are matplotlib PNGs inlined as
data URIs. They are deliberately NOT screenshots of the Leaflet map: a screenshot
carries the zoom buttons, the attribution strip and whatever the user had panned
into view, none of which belong in a document going to an investment committee.
"""
import base64
import io
import math
import re

# Palette — matches the report template and the rest of the acquisitions tab.
NAVY = "#13344E"
NAVY_DEEP = "#0B2233"
ORANGE = "#F25929"
BLUE = "#3B5BA5"
GREEN = "#2E7D4F"
GREY = "#6B7B8B"
GREY_LINE = "#DDE3E8"
CREAM = "#F7F4EF"


# ---------------------------------------------------------------------------
# Figure helpers
# ---------------------------------------------------------------------------
def _fig_to_uri(fig, dpi=170, pad=0.0, transparent=False):
    """PNG data URI with no surrounding padding.

    bbox_inches="tight" plus a white facecolor is what put a white band around
    the map in the reference deck: the axes keep an equal aspect, so any figure
    whose shape does not match the geography gets letterboxed and the padding
    is painted white. Everything here is sized to fill its frame instead.
    """
    import matplotlib.pyplot as plt
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, pad_inches=pad,
                bbox_inches=None if pad == 0 else "tight",
                facecolor="none" if transparent else fig.get_facecolor(),
                edgecolor="none", transparent=transparent)
    plt.close(fig)
    buf.seek(0)
    return "data:image/png;base64," + base64.b64encode(buf.read()).decode("ascii")


def _geo_figsize(bounds, target_w=7.4, min_h=3.0, max_h=5.2):
    """Figure size whose shape matches the ground, so nothing is letterboxed.

    A degree of longitude is shorter than a degree of latitude everywhere but
    the equator, so the true aspect uses cos(lat) at the map's centre.
    """
    minx, miny, maxx, maxy = bounds
    mid_lat = (miny + maxy) / 2.0
    w_deg = max(maxx - minx, 1e-6)
    h_deg = max(maxy - miny, 1e-6)
    ground_w = w_deg * math.cos(math.radians(mid_lat))
    ground_h = h_deg
    h = target_w * (ground_h / ground_w) if ground_w else target_w * 0.7
    return (target_w, max(min_h, min(max_h, h)))


def _padded_bounds(bounds, frac=0.12):
    minx, miny, maxx, maxy = bounds
    px = (maxx - minx) * frac or 0.004
    py = (maxy - miny) * frac or 0.004
    return minx - px, miny - py, maxx + px, maxy + py


def _fit_bounds_to_figure(ax, bounds, figsize):
    """Widen the shorter axis so the data fills the frame exactly.

    With set_aspect("equal") matplotlib will letterbox whenever the data's
    aspect and the figure's aspect disagree. Rather than drop equal aspect --
    which would stretch the parcels out of shape -- the view is grown on
    whichever axis has slack, so the map reaches every edge and the geometry
    stays true.
    """
    minx, miny, maxx, maxy = bounds
    mid_lat = (miny + maxy) / 2.0
    k = math.cos(math.radians(mid_lat)) or 1.0
    fig_ar = figsize[0] / figsize[1]                 # width / height
    data_ar = ((maxx - minx) * k) / max(maxy - miny, 1e-9)
    if data_ar < fig_ar:                             # too tall -> widen
        want = (maxy - miny) * fig_ar / k
        cx = (minx + maxx) / 2.0
        minx, maxx = cx - want / 2.0, cx + want / 2.0
    else:                                            # too wide -> heighten
        want = ((maxx - minx) * k) / fig_ar
        cy = (miny + maxy) / 2.0
        miny, maxy = cy - want / 2.0, cy + want / 2.0
    ax.set_xlim(minx, maxx)
    ax.set_ylim(miny, maxy)
    return minx, miny, maxx, maxy


def _strip_axes(ax):
    """A map is a picture, not a plot: no ticks, no frame, no labels."""
    ax.set_xticks([])
    ax.set_yticks([])
    for s in ax.spines.values():
        s.set_visible(False)
    ax.set_xlabel("")
    ax.set_ylabel("")


def _scale_bar(ax, bounds, colour="#FFFFFF"):
    """Distance scale, chosen from the map's own width at its centre latitude."""
    minx, miny, maxx, maxy = bounds
    span_ft = (maxx - minx) * 364000.0 * math.cos(math.radians((miny + maxy) / 2))
    if span_ft <= 0:
        return
    nice = min([264, 528, 1320, 2640, 5280, 10560, 26400, 52800],
               key=lambda v: abs(v - span_ft * 0.24))
    frac = nice / span_ft
    bx = minx + (maxx - minx) * 0.045
    by = miny + (maxy - miny) * 0.055
    ax.plot([bx, bx + (maxx - minx) * frac], [by, by], color="#00000055",
            lw=4.4, solid_capstyle="butt", zorder=60)
    ax.plot([bx, bx + (maxx - minx) * frac], [by, by], color=colour,
            lw=2.0, solid_capstyle="butt", zorder=61)
    ax.text(bx + (maxx - minx) * frac / 2, by + (maxy - miny) * 0.022,
            (f"{nice/5280:g} mi" if nice >= 5280 else f"{nice:,} ft"),
            ha="center", va="bottom", fontsize=7.5, color="#FFFFFF",
            fontweight="bold", zorder=62,
            bbox=dict(boxstyle="round,pad=0.22", facecolor=NAVY_DEEP,
                      edgecolor="none", alpha=0.82))


def _north_arrow(ax):
    ax.text(0.972, 0.965, "N\n▲", transform=ax.transAxes, ha="center", va="top",
            fontsize=10.5, color="#FFFFFF", fontweight="bold", zorder=62,
            bbox=dict(boxstyle="round,pad=0.28", facecolor=NAVY_DEEP,
                      edgecolor="white", linewidth=0.7))


def _esri_basemap(ax, bounds, kind="imagery"):
    """Composite Esri XYZ tiles behind the map. Silent no-op if unreachable."""
    import numpy as np
    try:
        import requests
        from PIL import Image
    except Exception:
        return False
    minx, miny, maxx, maxy = bounds

    def to_tile(lon, lat, z):
        n = 2 ** z
        lat_r = math.radians(lat)
        xt = (lon + 180.0) / 360.0 * n
        yt = (1 - math.log(math.tan(lat_r) + 1 / math.cos(lat_r)) / math.pi) / 2 * n
        return xt, yt

    # Imagery alone has no place names, so a reader cannot tell where the site
    # is. Esri publishes the labels and roads as a separate transparent
    # reference layer; drawing it over the imagery gives the hybrid view.
    services = (["World_Imagery", "Reference/World_Boundaries_and_Places"]
                if kind == "imagery" else ["World_Topo_Map"])
    service = services[0]
    for z in range(16, 9, -1):
        x0, y1 = to_tile(minx, miny, z)
        x1, y0 = to_tile(maxx, maxy, z)
        tx0, tx1 = int(math.floor(x0)), int(math.floor(x1))
        ty0, ty1 = int(math.floor(y0)), int(math.floor(y1))
        if (tx1 - tx0 + 1) * (ty1 - ty0 + 1) <= 36:
            break
    else:
        return False
    try:
        cols, rows = tx1 - tx0 + 1, ty1 - ty0 + 1
        canvas = Image.new("RGB", (cols * 256, rows * 256), "#DfE6EC")

        # Fetched in parallel. Serially this was 36 round trips at roughly
        # 0.7s each -- 25 seconds for one map, and the report draws two.
        from concurrent.futures import ThreadPoolExecutor

        def grab(ij):
            """One tile, retried. A single dropped tile leaves a grey
            rectangle sitting in the middle of the aerial -- it happened on
            the Story Lindsey cover -- and these are transient often enough
            that one retry usually closes it."""
            i, j, tx, ty = ij
            url = (f"https://server.arcgisonline.com/ArcGIS/rest/services/"
                   f"{service}/MapServer/tile/{z}/{ty}/{tx}")
            for attempt in range(3):
                try:
                    r = requests.get(url, timeout=8 + attempt * 4)
                    if r.status_code == 200:
                        return i, j, Image.open(io.BytesIO(r.content)).convert("RGB")
                except Exception:
                    pass
            return None

        jobs = [(i, j, tx, ty)
                for i, tx in enumerate(range(tx0, tx1 + 1))
                for j, ty in enumerate(range(ty0, ty1 + 1))]
        got, filled, missing = 0, [], []
        with ThreadPoolExecutor(max_workers=12) as pool:
            for res in pool.map(grab, jobs):
                if res is None:
                    continue
                i, j, im = res
                canvas.paste(im, (i * 256, j * 256))
                filled.append((i, j, im))
                got += 1
        if not got:
            return False
        # Any tile still missing is painted with the average colour of the
        # ones that arrived, so a gap reads as haze rather than as a grey
        # box someone will ask about.
        have = {(i, j) for i, j, _ in filled}
        missing = [(i, j) for i, _, _ in [(x, 0, 0) for x in range(cols)]
                   for j in range(rows) if (i, j) not in have]
        if missing and filled:
            sample = filled[len(filled) // 2][2].resize((1, 1)).getpixel((0, 0))
            patch = Image.new("RGB", (256, 256), sample)
            for i, j in missing:
                canvas.paste(patch, (i * 256, j * 256))
            print(f"[report] basemap: {len(missing)} tile(s) unavailable, "
                  f"filled from neighbours", flush=True)

        def tile_to_lonlat(xt, yt, z):
            n = 2 ** z
            lon = xt / n * 360.0 - 180.0
            lat = math.degrees(math.atan(math.sinh(math.pi * (1 - 2 * yt / n))))
            return lon, lat

        w_lon, n_lat = tile_to_lonlat(tx0, ty0, z)
        e_lon, s_lat = tile_to_lonlat(tx1 + 1, ty1 + 1, z)
        ax.imshow(np.asarray(canvas), extent=[w_lon, e_lon, s_lat, n_lat],
                  origin="upper", zorder=0, interpolation="bilinear")
        return True
    except Exception:
        return False


def _draw_geom(ax, geom, **kw):
    """Draw a shapely geometry of any type without caring which type it is."""
    from shapely.geometry import (Polygon, MultiPolygon, LineString,
                                  MultiLineString, GeometryCollection)
    if geom is None or geom.is_empty:
        return
    if isinstance(geom, (MultiPolygon, MultiLineString, GeometryCollection)):
        for g in geom.geoms:
            _draw_geom(ax, g, **kw)
        return
    if isinstance(geom, Polygon):
        xs, ys = geom.exterior.xy
        ax.fill(xs, ys, facecolor=kw.get("face", "none"),
                edgecolor=kw.get("edge", "none"),
                linewidth=kw.get("lw", 0.8), alpha=kw.get("alpha", 1.0),
                zorder=kw.get("z", 5))
        for ring in geom.interiors:
            xs, ys = ring.xy
            ax.fill(xs, ys, facecolor=CREAM, edgecolor="none",
                    alpha=kw.get("alpha", 1.0), zorder=kw.get("z", 5) + 0.1)
    elif isinstance(geom, LineString):
        xs, ys = geom.xy
        ax.plot(xs, ys, color=kw.get("edge", BLUE), linewidth=kw.get("lw", 1.0),
                alpha=kw.get("alpha", 1.0), zorder=kw.get("z", 5))


def render_site_map(union_geom, constraint_geoms=None, tracts=None,
                    width_in=7.4):
    """The subject site with its physical constraints, edge to edge.

    No zoom controls, no attribution strip, no letterboxing -- the three things
    that gave away the reference deck's map as a screenshot of the web app.
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from shapely.geometry import shape as shp_shape

    if union_geom is None:
        return None
    bounds = _padded_bounds(union_geom.bounds, 0.16)
    figsize = _geo_figsize(bounds, target_w=width_in, min_h=3.2, max_h=4.9)
    fig, ax = plt.subplots(figsize=figsize, dpi=170)
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    ax.set_position([0, 0, 1, 1])
    ax.set_facecolor("#E7ECF0")
    ax.set_aspect("equal")
    bounds = _fit_bounds_to_figure(ax, bounds, figsize)
    _esri_basemap(ax, bounds, "imagery")

    cg = constraint_geoms or {}
    layers = [
        ("floodplain", "#5B6FD6", 0.42, "Floodplain (100-yr)"),
        ("wetlands", "#2E9E6B", 0.48, "Wetlands (NWI)"),
        ("stream_buffers", "#2F7FD6", 0.55, "Streams"),
        ("pipeline_easements", "#C99A2E", 0.55, "Pipeline easement"),
        ("transmission_row", "#B0552E", 0.55, "Transmission ROW"),
    ]
    legend = []
    for key, colour, alpha, label in layers:
        gj = cg.get(key)
        if not gj:
            continue
        try:
            g = shp_shape(gj) if isinstance(gj, dict) else gj
        except Exception:
            continue
        if g is None or g.is_empty:
            continue
        _draw_geom(ax, g, face=colour, edge=colour, alpha=alpha, lw=0.6, z=3)
        legend.append((label, colour))

    # Individual tracts hairlined inside the assembly outline, so a multi-tract
    # deal reads as an assembly rather than one blob.
    for t in (tracts or []):
        try:
            g = shp_shape(t.get("geometry")) if t.get("geometry") else None
        except Exception:
            g = None
        if g is not None and not g.is_empty:
            _draw_geom(ax, g, face="none", edge="#FFD9A0", lw=0.7, alpha=0.85, z=6)

    _draw_geom(ax, union_geom, face="none", edge=ORANGE, lw=2.1, z=8)
    _strip_axes(ax)
    _north_arrow(ax)
    _scale_bar(ax, bounds)

    if legend:
        from matplotlib.patches import Patch
        ax.legend(handles=[Patch(facecolor=c, edgecolor="none", alpha=0.75,
                                 label=l) for l, c in legend],
                  loc="lower right", fontsize=6.8, frameon=True,
                  facecolor="#FFFFFFEE", edgecolor=GREY_LINE,
                  borderpad=0.5, handlelength=1.1).set_zorder(63)
    return _fig_to_uri(fig, dpi=170, pad=0.0)


def render_competitor_map(center, communities, radius_mi=None, width_in=7.4):
    """Subject site with the competing communities around it."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    pts = [(c.get("lon"), c.get("lat"), c) for c in (communities or [])
           if c.get("lat") and c.get("lon")]
    if not center or not center.get("lat"):
        return None
    clon, clat = float(center["lon"]), float(center["lat"])
    xs = [p[0] for p in pts] + [clon]
    ys = [p[1] for p in pts] + [clat]
    if len(xs) < 2:
        span = (radius_mi or 5) / 55.0
        bounds = (clon - span, clat - span, clon + span, clat + span)
    else:
        bounds = _padded_bounds((min(xs), min(ys), max(xs), max(ys)), 0.14)

    figsize = _geo_figsize(bounds, target_w=width_in, min_h=2.8, max_h=3.7)
    fig, ax = plt.subplots(figsize=figsize, dpi=170)
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    ax.set_position([0, 0, 1, 1])
    ax.set_facecolor("#EEF1F4")
    ax.set_aspect("equal")
    bounds = _fit_bounds_to_figure(ax, bounds, figsize)
    _esri_basemap(ax, bounds, "topo")

    if radius_mi:
        from matplotlib.patches import Circle
        r_deg = float(radius_mi) / 69.0
        ax.add_patch(Circle((clon, clat), r_deg, fill=False, linestyle=(0, (4, 3)),
                            edgecolor=ORANGE, linewidth=1.1, alpha=0.75, zorder=4))

    status_style = {
        "active": (GREEN, "o", "Active"),
        "future": ("#C99A2E", "^", "Future"),
        "closed": (GREY, "s", "Built out"),
    }
    seen = {}
    for lon, lat, c in pts:
        st = str(c.get("status") or "active").lower()
        key = ("future" if "future" in st else
               "closed" if ("close" in st or "built" in st or "sold" in st) else "active")
        colour, marker, label = status_style[key]
        lots = float(c.get("total_lots") or c.get("lots") or 0)
        size = 18 + min(math.sqrt(max(lots, 0)) * 3.2, 90)
        ax.scatter([lon], [lat], s=size, c=colour, marker=marker,
                   edgecolors="#FFFFFF", linewidths=0.6, alpha=0.9, zorder=6)
        seen[label] = (colour, marker)

    ax.scatter([clon], [clat], s=230, marker="*", c=ORANGE,
               edgecolors="#FFFFFF", linewidths=1.1, zorder=9)
    ax.annotate("SUBJECT", (clon, clat), textcoords="offset points",
                xytext=(0, -15), ha="center", fontsize=7.4, fontweight="bold",
                color="#FFFFFF", zorder=10,
                bbox=dict(boxstyle="round,pad=0.24", facecolor=ORANGE,
                          edgecolor="none"))

    _strip_axes(ax)
    _north_arrow(ax)
    _scale_bar(ax, bounds, colour="#FFFFFF")
    if seen:
        from matplotlib.lines import Line2D
        handles = [Line2D([], [], marker=m, color="none", markerfacecolor=c,
                          markeredgecolor="#FFFFFF", markersize=7, label=l)
                   for l, (c, m) in seen.items()]
        ax.legend(handles=handles, loc="lower right", fontsize=6.8, frameon=True,
                  facecolor="#FFFFFFEE", edgecolor=GREY_LINE,
                  borderpad=0.5).set_zorder(63)
    return _fig_to_uri(fig, dpi=170, pad=0.0)


def render_roads_map(center, projects, site_geom=None, radius_mi=None, width_in=7.4):
    """Programmed roadway work around the site, coloured by how soon it lets.

    The table gives the roadway names; the map is what shows whether the
    investment actually surrounds this site or sits on the far side of the
    county. Projects with no let year are drawn faintly rather than dropped --
    unfunded capacity work is still context.
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import datetime as _dt

    lines = []
    for p in (projects or []):
        g = p.get("geometry") or {}
        if g.get("type") != "LineString" or not g.get("coordinates"):
            continue
        lines.append((p, [c for c in g["coordinates"] if len(c) >= 2]))
    if not lines or not center or not center.get("lat"):
        return None
    clon, clat = float(center["lon"]), float(center["lat"])

    xs = [c[0] for _, cs in lines for c in cs] + [clon]
    ys = [c[1] for _, cs in lines for c in cs] + [clat]
    bounds = _padded_bounds((min(xs), min(ys), max(xs), max(ys)), 0.05)
    figsize = _geo_figsize(bounds, target_w=width_in, min_h=3.0, max_h=4.4)
    fig, ax = plt.subplots(figsize=figsize, dpi=170)
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    ax.set_position([0, 0, 1, 1])
    ax.set_facecolor("#EEF1F4")
    ax.set_aspect("equal")
    bounds = _fit_bounds_to_figure(ax, bounds, figsize)
    _esri_basemap(ax, bounds, "topo")

    this_year = _dt.date.today().year
    seen = {}
    for p, cs in lines:
        yr = _n(p.get("let_year"))
        if yr is None:
            colour, lw, alpha, lab = GREY, 1.1, 0.55, "Unscheduled"
        elif yr <= this_year + 4:
            colour, lw, alpha, lab = ORANGE, 2.4, 0.95, "Lets within 4 years"
        else:
            colour, lw, alpha, lab = BLUE, 1.7, 0.8, "Later horizon"
        ax.plot([c[0] for c in cs], [c[1] for c in cs], color=colour,
                linewidth=lw, alpha=alpha, solid_capstyle="round", zorder=5)
        seen[lab] = colour

    if site_geom is not None:
        _draw_geom(ax, site_geom, face=ORANGE, edge=ORANGE, alpha=0.35, lw=1.4, z=8)
    ax.scatter([clon], [clat], s=150, marker="*", c=ORANGE, edgecolors="#FFFFFF",
               linewidths=1.0, zorder=9)
    ax.annotate("SITE", (clon, clat), textcoords="offset points", xytext=(0, -13),
                ha="center", fontsize=7, fontweight="bold", color="#FFFFFF", zorder=10,
                bbox=dict(boxstyle="round,pad=0.22", facecolor=NAVY_DEEP, edgecolor="none"))

    _strip_axes(ax)
    _north_arrow(ax)
    _scale_bar(ax, bounds)
    if seen:
        from matplotlib.lines import Line2D
        ax.legend(handles=[Line2D([], [], color=c, lw=2.4, label=l)
                           for l, c in seen.items()],
                  loc="lower right", fontsize=6.8, frameon=True,
                  facecolor="#FFFFFFEE", edgecolor=GREY_LINE,
                  borderpad=0.5).set_zorder(63)
    return _fig_to_uri(fig, dpi=170, pad=0.0)


# ---------------------------------------------------------------------------
# Formatting — one place, so "1,064.7 ac" and "$326,881" look the same
# everywhere and a missing value never prints as None, NaN or undefined.
# ---------------------------------------------------------------------------
def _n(v):
    try:
        f = float(v)
        return None if f != f else f          # NaN check
    except (TypeError, ValueError):
        return None


def ac(v, dash="—"):
    f = _n(v)
    return dash if f is None else f"{f:,.1f} ac"


def num(v, dash="—"):
    f = _n(v)
    return dash if f is None else f"{f:,.0f}"


def pct(v, dash="—", dp=1):
    f = _n(v)
    return dash if f is None else f"{f:,.{dp}f}%"


def money(v, dash="—"):
    f = _n(v)
    if f is None:
        return dash
    if abs(f) >= 1_000_000_000:
        return f"${f/1e9:,.2f}B"
    if abs(f) >= 1_000_000:
        return f"${f/1e6:,.0f}M"
    if abs(f) >= 10_000:
        return f"${f/1000:,.0f}k"
    return f"${f:,.0f}"


def miles(v, direction=None, dash="—"):
    f = _n(v)
    if f is None:
        return dash
    return f"{f:,.2f} mi" + (f" {direction}" if direction else "")


def _first(d, *keys, default=None):
    """First key that carries a usable value. The upstream payloads spell the
    same idea several ways depending on which service answered."""
    for k in keys:
        if isinstance(d, dict) and d.get(k) not in (None, "", [], {}):
            return d[k]
    return default


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
def render_quarter_chart(quarters, width_in=6.9):
    """Starts vs closings by quarter — grouped bars, no chart junk."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np
    if not quarters:
        return None
    labels = [q.get("label") or q.get("quarter") or "" for q in quarters]
    starts = [_n(q.get("starts")) or 0 for q in quarters]
    clos = [_n(q.get("closings")) or 0 for q in quarters]
    fig, ax = plt.subplots(figsize=(width_in, 1.45), dpi=170)
    fig.patch.set_facecolor("white")
    x = np.arange(len(labels))
    ax.bar(x - 0.19, starts, 0.36, color=NAVY, label="Starts")
    ax.bar(x + 0.19, clos, 0.36, color=ORANGE, label="Closings")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=7, color=GREY)
    ax.tick_params(axis="y", labelsize=7, colors=GREY)
    ax.grid(axis="y", linestyle=":", color=GREY_LINE, linewidth=0.6)
    ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(GREY_LINE)
    ax.legend(fontsize=7, frameon=False, ncol=2, loc="upper left")
    fig.tight_layout(pad=0.4)
    return _fig_to_uri(fig, dpi=170, pad=0.02)


def render_enrollment_chart(series, width_in=6.9):
    """District enrolment history. Takes one series or several.

    A tract on a district line belongs to both, so both trajectories matter --
    charting only the larger one told half the story. Each district gets its
    own line; the callouts stay on the largest so two sets of labels do not
    collide.

    `series` is a list of (name, trend) pairs; a bare trend list is accepted
    for the single-district case.
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter

    if series and isinstance(series[0], dict):        # a bare trend
        series = [(None, series)]
    cleaned = []
    for name, trend in (series or []):
        pts = sorted([(int(y), e) for y, e in
                      ((p.get("year"), _n(p.get("enrollment"))) for p in (trend or []))
                      if y and e])
        if len(pts) >= 3:
            cleaned.append((name, pts))
    if not cleaned:
        return None
    cleaned.sort(key=lambda t: -t[1][-1][1])          # largest district first

    fig, ax = plt.subplots(figsize=(width_in, 1.85), dpi=170)
    fig.patch.set_facecolor("white")
    colours = [ORANGE, BLUE, GREEN]
    all_y = []
    for idx, (name, pts) in enumerate(cleaned):
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        all_y += ys
        col = colours[idx % len(colours)]
        if idx == 0:
            ax.fill_between(xs, ys, color=col, alpha=0.13)
        ax.plot(xs, ys, color=col, linewidth=1.6, zorder=3, label=(name or None))
        ax.scatter([xs[-1]], [ys[-1]], s=24, color=col, zorder=7)
        if idx == 0:
            first, last = xs[0], xs[-1]
            marks = [i for i, y in enumerate(xs)
                     if y in (first, last) or (y % 5 == 0 and last - y >= 3)]
            if len(marks) > 9:
                marks = [i for i in marks if xs[i] % 10 == 0 or xs[i] in (first, last)]
            for i in marks:
                end = xs[i] in (first, last)
                ax.annotate(f"{ys[i]:,.0f}", (xs[i], ys[i]),
                            textcoords="offset points", xytext=(0, 7), ha="center",
                            fontsize=6.6, fontweight="bold" if end else "normal",
                            color=col if end else NAVY, zorder=6)
        else:
            ax.annotate(f"{ys[-1]:,.0f}", (xs[-1], ys[-1]),
                        textcoords="offset points", xytext=(0, -11), ha="right",
                        fontsize=6.6, fontweight="bold", color=col, zorder=6)

    ax.set_ylim(min(all_y) * 0.82, max(all_y) * 1.16)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.tick_params(labelsize=6.8, colors=GREY)
    ax.grid(axis="y", linestyle=":", color=GREY_LINE, linewidth=0.6)
    ax.set_axisbelow(True)
    for sp in ("top", "right", "left"):
        ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color(GREY_LINE)
    if len(cleaned) > 1:
        ax.legend(fontsize=6.6, frameon=False, ncol=len(cleaned), loc="upper left")
    fig.tight_layout(pad=0.4)
    return _fig_to_uri(fig, dpi=170, pad=0.02)


def _thesis(r, a, data):
    """The two or three sentences an executive reads first.

    Every clause rests on a figure printed elsewhere in the document -- scale,
    conversion, ring velocity, supply -- so the summary cannot claim something
    the report does not show.
    """
    bits = []
    gross = _n(a.get("gross_acres")) or 0
    scale = ("Large-scale" if gross >= 500 else
             "Mid-scale" if gross >= 150 else "Infill-scale")
    where = (r.get("schools", {}) or {}).get("district") or ""
    where = where.title().replace(" Isd", " ISD") if where else ""
    where = where or (r.get("project", {}).get("location") or "the submarket")
    opening = f"{scale} residential opportunity in the {where} corridor"
    cb = data.get("cbas") or {}
    starts = _n(cb.get("ring_annual_starts")) or _n(
        (cb.get("aggregate") or {}).get("annual_starts"))
    if starts:
        opening += f", where the competitive ring runs {starts:,.0f} starts a year"
    bits.append(opening + ".")

    conv = _n(a.get("net_saleable_pct"))
    if conv is not None:
        lead = ("The central site issue is development efficiency"
                if conv < 55 else "Development efficiency is favourable")
        fl = next((d for d in (a.get("netout_detail") or [])
                   if d.get("key") == "flood"), {})
        fl_ac = _n(fl.get("acres")) or 0
        why = (f", driven by {fl_ac:,.0f} acres of floodplain"
               if gross and fl_ac / gross > 0.15 else "")
        bits.append(f"{lead}{why}: gross-to-saleable conversion is {conv:,.1f}%.")

    mos = _mos_of(r)
    fut = _n((cb.get("aggregate") or {}).get("futures"))
    if mos is not None or fut:
        line = "Supply warrants caution"
        if mos is not None:
            line += f": {mos:,.1f} months of lot supply"
        if fut:
            line += f"{' and' if mos is not None else ':'} {fut:,.0f} future lots in the ring"
        bits.append(line + ".")
    return " ".join(bits)


def _mos_of(r):
    """Months of lot supply as a number, back off the formatted market card."""
    v = (r.get("market") or {}).get("mos")
    if v in (None, "", "-"):
        return None
    return _n(str(v).replace(",", ""))


def _inline_name(name):
    """District names are stored upper-case for the card label; mid-sentence
    that reads as shouting, so give it back its title case."""
    n = str(name or "").strip()
    if n and n.isupper():
        n = n.title().replace(" Isd", " ISD").replace(" Hs", " HS")
    return n or "The district"


def _pct_of_gross(a, key):
    gross = _n(a.get("gross_acres")) or 0
    for d in (a.get("netout_detail") or []):
        if d.get("key") == key and gross:
            acres = _n(d.get("acres")) or 0
            return acres, acres / gross * 100
    return None, None


def _positives(r, a, data):
    """What supports the deal, drawn from every section that rendered.

    Each line names the figure it rests on, so a reader can find it elsewhere
    in the document. Nothing is asserted that the report does not also show.
    """
    out = []
    # Read the MAPPED schools block, not the raw payload: since a tract can
    # straddle a district line that payload is a list of districts, and
    # calling .get on it raised AttributeError and took the whole report down.
    sch = r.get("schools") or {}
    g5txt = next((g.get("value") for g in (sch.get("growth") or [])
                  if "5" in str(g.get("label", ""))), None)
    g5 = _n(str(g5txt).strip("+%")) if g5txt else None
    if g5 is not None and g5 > 8:
        out.append(f"{_inline_name(sch.get('district'))} enrolment is "
                   f"up {g5:,.1f}% over five years, a direct read on household formation.")

    mk = r.get("market") or {}
    starts = _n((data.get("cbas") or {}).get("ring_annual_starts")) or _n(
        ((data.get("cbas") or {}).get("aggregate") or {}).get("annual_starts"))
    if starts and starts >= 400:
        out.append(f"{starts:,.0f} annual starts in the competitive ring: builders are "
                   "already proving demand here, which lowers market-proof risk.")
    if mk.get("builders"):
        names = ", ".join(b["name"] for b in mk["builders"][:3])
        out.append(f"National and regional builders are active on the ring -- {names} "
                   "among them -- so lot takedown has credible counterparties.")
    if mk.get("lot_bands"):
        top = mk["lot_bands"][0]
        out.append(f"The deepest lot product is {top['width']} at {top['lots']} lots, "
                   "which the preliminary mix can be aimed at.")

    slope = _n((a.get("_topo") or {}).get("mean_slope_pct"))
    if slope is not None and slope < 2:
        out.append(f"Gentle topography at {slope:,.2f}% mean slope implies standard "
                   "grading rather than heavy earthwork.")
    conv = _n(a.get("net_saleable_pct"))
    if conv is not None and conv >= 55:
        out.append(f"{conv:,.1f}% of gross converts to saleable land, an efficient site.")

    acc = r.get("access") or {}
    fw = next((c for c in acc.get("cards", []) if c["label"] == "Nearest freeway"), None)
    if fw:
        out.append(f"{fw['value']} is {fw['note']}, giving a credible regional access story.")
    if acc.get("funded_note"):
        out.append(f"Roadway investment is committed nearby: {acc['funded_note'].lower()}.")

    dem = r.get("demographics") or {}
    pop = next((d for d in dem.get("rows", []) if d["label"] == "Population"), None)
    inc = next((d for d in dem.get("rows", []) if d["label"] == "Median HH income"), None)
    if pop and inc:
        out.append(f"County demand context: {pop['value']} residents at "
                   f"{inc['value']} median household income.")
    return out[:5]


def _risks(r, a, data):
    """What could break the deal. Same rule: every line cites a figure shown."""
    out = []
    fl_ac, fl_pct = _pct_of_gross(a, "flood")
    if fl_pct is not None and fl_pct > 15:
        out.append(f"{fl_ac:,.1f} acres of floodplain -- {fl_pct:,.1f}% of gross -- "
                   "materially reduces land efficiency and implies mitigation and "
                   "detention cost.")
    conv = _n(a.get("net_saleable_pct"))
    if conv is not None and conv < 50:
        out.append(f"Only {conv:,.1f}% of gross reaches net saleable, so land basis has "
                   "to be underwritten on saleable acres rather than gross.")

    mk = r.get("market") or {}
    mos = _n(str(mk.get("mos", "")).replace(",", "") or None)
    fut = _n(((data.get("cbas") or {}).get("aggregate") or {}).get("futures"))
    if mos is not None and mos >= 18:
        line = f"{mos:,.1f} months of lot supply"
        if fut:
            line += f" plus {fut:,.0f} future lots in the ring"
        out.append(line + " creates absorption and pricing pressure.")
    elif fut and fut > 10000:
        out.append(f"{fut:,.0f} future lots in the ring will compete with later phases.")

    sd = r.get("schools") or {}
    rating = str(sd.get("rating") or "")
    if rating[:1] in ("C", "D", "F"):
        out.append(f"{_inline_name(sd.get('district'))} carries a TEA rating of "
                   f"{rating}, which can cap pricing power against stronger "
                   "school-district submarkets.")

    slope = _n((a.get("_topo") or {}).get("max_slope_pct"))
    if slope is not None and slope > 8:
        out.append(f"Maximum slope of {slope:,.1f}% indicates areas needing grading design.")

    am = r.get("amenities") or {}
    hosp = next((n for n in am.get("nearest", []) if n["label"] == "Nearest hospital"), None)
    groc = next((n for n in am.get("nearest", []) if n["label"] == "Nearest grocery"), None)
    far = [n for n in (hosp, groc) if n and (_n(n["value"].split()[0]) or 0) > 7]
    if far:
        out.append("Daily-needs and healthcare convenience is dispersed ("
                   + "; ".join(f"{n['label'].lower()} {n['value']}" for n in far)
                   + "), which lengthens the growth thesis.")
    return out[:5]


def _site_read(r, a):
    """The two or three lines that sit under the site section."""
    out = []
    fl_ac, fl_pct = _pct_of_gross(a, "flood")
    if fl_pct is not None and fl_pct > 5:
        out.append(f"Floodplain is the primary land-efficiency issue at {fl_ac:,.1f} ac "
                   f"({fl_pct:,.1f}% of gross); wetlands and stream overlap add limited "
                   "incremental loss because they largely sit inside it.")
    topo = r.get("topo") or {}
    y = r.get("yield") or {}
    if topo.get("character") and y.get("total_lots"):
        out.append(f"Topography is {topo['character'].lower()}. The current mix implies "
                   f"{y['total_lots']} lots at {y.get('density', 'the stated density')} "
                   "before civil refinement.")
    infra = _n(a.get("infrastructure_acres"))
    if infra:
        out.append(f"Infrastructure and landscaping is carried at "
                   f"{_n(a.get('infrastructure_pct')) or 30:,.0f}% of net developable "
                   f"({infra:,.1f} ac), which should be tested against a real "
                   "roads-and-detention layout.")
    return out[:3] or None


NEXT_STEPS = [
    "Civil: confirm floodplain reclamation assumptions, detention need, drainage "
    "outfalls and off-site utility availability.",
    "Market: builder calls and LOIs by lot width, phase and price; reconcile "
    "capture against current community absorption.",
    "Phasing: model a conservative lot takedown against prevailing months of "
    "supply and the future pipeline.",
    "Land basis: derive maximum land value from net saleable acreage and "
    "finished-lot economics, not gross acreage.",
    "Entitlements: validate jurisdiction, MUD / PID / utility-district path, "
    "school boundaries and roadway obligations.",
    "IC output: return with low / base / high yield, absorption and cost cases "
    "plus a clear walk-away land basis.",
]



# ---------------------------------------------------------------------------
# Height budget
#
# A section either fits its sheet or it does not. If a table spills three rows
# onto a fresh page, the page it came from ends short AND the spill page sits
# two-thirds empty, because the next section always starts fresh -- so the only
# real fix is to not spill. These are the template's own measurements in
# points; keep them in step if the CSS padding changes.
#
# Letter is 792pt tall; the @page margins take 0.52in and 0.62in, leaving this.
# ---------------------------------------------------------------------------
PAGE_PT = 709.9
SECTION_CHROME_PT = 74.0        # running header + section title + subtitle
KPI_ROW_PT = 56.0               # a row of metric cards, incl. its margin
TABLE_CHROME_PT = 42.0          # card title + table header + card margin
TABLE_ROW_PT = 18.0
CARD_CHROME_PT = 45.0           # padding + title + margin on a plain card
READ_LINE_PT = 16.0
IMAGE_W_PT = 532.8              # content width: 8.5in less 0.55in each side


def _image_pt(aspect_w, aspect_h, card=True):
    """Height of a full-width figure, plus its card if it sits in one."""
    return IMAGE_W_PT * (aspect_h / aspect_w) + (CARD_CHROME_PT if card else 0)


def _rows_that_fit(used_pt, n_wanted, min_rows=3):
    """How many table rows are left after everything else on the sheet."""
    spare = PAGE_PT - used_pt - TABLE_CHROME_PT
    return max(min_rows, min(n_wanted, int(spare // TABLE_ROW_PT)))

# ---------------------------------------------------------------------------
# Context assembly
# ---------------------------------------------------------------------------
BRIDGE_COLOURS = {
    "gross": NAVY, "flood": "#5B6FD6", "wetlands": "#2E9E6B",
    "streams": "#2F7FD6", "pipelines": "#C99A2E", "transmission": "#B0552E",
    "netdev": NAVY, "infra": ORANGE, "saleable": ORANGE,
}
CONSTRAINT_COLOURS = {
    "flood": "#5B6FD6", "wetlands": "#2E9E6B", "streams": "#2F7FD6",
    "pipelines": "#C99A2E", "transmission": "#B0552E",
}


def build_context(proj, analysis, data=None, elevation=None):
    """Everything the template needs, from data the app already produced.

    `data` carries the payloads the route pulled from the existing endpoints.
    Any of them may be missing or carry an "error" -- in that case the section
    is left out of the context entirely and the template drops it, rather than
    printing an empty card or the word None.
    """
    data = data or {}
    a = dict(analysis or {})
    if elevation:
        a["_topo"] = elevation
    tracts = proj.get("tracts") or []
    gross = _n(a.get("gross_acres")) or 0.0

    counties = sorted({(t.get("county") or "").strip() for t in tracts if t.get("county")})
    location = ", ".join(counties) + (" County, Texas" if len(counties) == 1
                                      else " Counties, Texas") if counties else "Texas"
    import datetime as _dt
    today = _dt.date.today()

    r = {
        "project": {
            "name": proj.get("name") or "Untitled project",
            "name_upper": (proj.get("name") or "Untitled project").upper(),
            "location": location,
            "tract_line": f"{len(tracts)} tract" + ("" if len(tracts) == 1 else "s"),
            "date_long": today.strftime("%B %Y"),
        },
        "kpi": {
            "gross": ac(a.get("gross_acres")),
            "net_dev": ac(a.get("net_developable_acres")),
            "net_saleable": ac(a.get("net_saleable_acres")),
            "lots": num((a.get("yield_estimates") or {}).get("total_lots")) + " lots",
        },
        "missing": [],
        # Where the appraisal district's boundary replaced StratMap's, the
        # document has to say so: the acreage a reader is pricing on moved.
        "geometry_notes": list(a.get("geometry_notes") or []),
    }

    # ---- tract composition -------------------------------------------------
    if tracts:
        r["tracts"] = [{
            "owner": t.get("owner_name") or "—",
            "prop_id": t.get("prop_id") or "—",
            "county": t.get("county") or "—",
            "acres": ac(t.get("acres")),
            "pct": pct((_n(t.get("acres")) or 0) / gross * 100 if gross else None),
        } for t in tracts]

    # ---- gross-to-saleable bridge -----------------------------------------
    detail = a.get("netout_detail") or []
    if gross:
        rows = [{"label": "Gross", "value": ac(gross), "pct": 100,
                 "colour": BRIDGE_COLOURS["gross"], "total": True}]
        for d in detail:
            if not d.get("applied"):
                continue
            marg = _n(d.get("acres_marginal"))
            if marg is None:
                marg = _n(d.get("acres")) or 0
            if marg <= 0:
                continue
            rows.append({
                "label": d.get("label", "").replace(" (100-yr)", "").replace(" (NWI)", ""),
                "value": "-" + ac(marg),
                "pct": max(1.0, marg / gross * 100),
                "colour": BRIDGE_COLOURS.get(d.get("key"), BLUE), "total": False})
        nd = _n(a.get("net_developable_acres"))
        rows.append({"label": "Net developable", "value": ac(nd),
                     "pct": (nd / gross * 100) if nd else 0,
                     "colour": BRIDGE_COLOURS["netdev"], "total": True})
        infra = _n(a.get("infrastructure_acres"))
        if infra:
            rows.append({"label": "Infra / landscape", "value": "-" + ac(infra),
                         "pct": max(1.0, infra / gross * 100),
                         "colour": BRIDGE_COLOURS["infra"], "total": False})
        ns = _n(a.get("net_saleable_acres"))
        rows.append({"label": "Net saleable", "value": ac(ns),
                     "pct": (ns / gross * 100) if ns else 0,
                     "colour": BRIDGE_COLOURS["saleable"], "total": True})
        r["bridge"] = rows

        cons = []
        for d in detail:
            acres = _n(d.get("acres"))
            if acres is None:
                continue
            cons.append({
                "label": d.get("label") or d.get("key"),
                "acres": ac(acres),
                "pct": pct(acres / gross * 100 if gross else None),
                "bar": min(100.0, (acres / gross * 100) if gross else 0),
                "colour": CONSTRAINT_COLOURS.get(d.get("key"), BLUE),
            })
        r["constraints"] = cons

    # ---- topography --------------------------------------------------------
    e = elevation or {}
    if _n(e.get("min_ft")) is not None:
        rng = (_n(e.get("max_ft")) or 0) - (_n(e.get("min_ft")) or 0)
        char = e.get("site_character") or e.get("character")
        r["topo"] = {
            "range": f"{rng:,.1f} ft",
            "min_max": f"{_n(e.get('min_ft')):,.1f} - {_n(e.get('max_ft')):,.1f} ft",
            "mean_slope": pct(e.get("mean_slope_pct"), dp=2),
            "max_slope": pct(e.get("max_slope_pct"), dp=2),
            "drainage": e.get("drainage") or e.get("drainage_dir"),
            "character": (str(char).upper() if char else None),
        }

    # ---- yield -------------------------------------------------------------
    y = a.get("yield_estimates") or {}
    if y.get("total_lots") is not None:
        r["yield"] = {
            "total_lots": num(y.get("total_lots")),
            "density": (f"{_n(y.get('weighted_density')):,.2f} u/ac"
                        if _n(y.get("weighted_density")) else "—"),
            "conversion": pct(a.get("net_saleable_pct")),
            "products": [{
                "product": p.get("label") or p.get("product") or "—",
                "density": (f"{_n(p.get('density')):,.1f} u/ac"
                            if _n(p.get("density")) else "—"),
                "allocation": pct(p.get("allocation_pct"), dp=0),
                "acres": ac(p.get("acres")),
                "lots": num(p.get("lots")),
            } for p in (y.get("breakdown") or [])],
        }

    r["next_steps"] = NEXT_STEPS

    # Each section is mapped in isolation. One upstream payload shaped
    # differently than expected should cost its own page, not the whole
    # document -- an executive report that fails entirely because the macro
    # feed renamed a field is worse than one that comes back a page short and
    # says so.
    import traceback
    sections = (
        ("market", _map_market, (data.get("cbas"),)),
        ("competition", _map_comps, (data.get("cbas"), data.get("comp_map"))),
        ("schools", _map_schools, (data.get("schools"),)),
        ("demographics", _map_demographics, (data.get("market"),)),
        ("access", _map_access, (data.get("roads"), data.get("amenities"),
                                 data.get("roads_map"))),
        ("amenities", _map_amenities, (data.get("amenities"),)),
        ("news", _map_news, (data.get("news"),)),
        ("macro", _map_macro, (data.get("fred"),)),
    )
    for name, fn, args in sections:
        try:
            fn(r, *args)
        except Exception as e:
            r["missing"].append(f"{name}: {type(e).__name__}: {e}")
            print(f"[report] section {name!r} failed: {e}{chr(10)}"
                  f"{traceback.format_exc()}", flush=True)

    # The synthesis runs LAST, once every section has been mapped. Run before,
    # it could only see the raw analysis and produced a single positive and a
    # single risk -- schools, demographics, access and the market read were all
    # sitting there unused.
    r["thesis"] = _thesis(r, a, data)
    r["why_works"] = _positives(r, a, data)
    r["what_breaks"] = _risks(r, a, data)
    r["dev_read"] = _site_read(r, a)
    return r


# ---------------------------------------------------------------------------
# Section mappers. Each is a no-op when its payload is absent, so the template
# omits that page rather than printing an empty shell.
# ---------------------------------------------------------------------------
def _substantial(*values, need=3):
    """True when enough of these are real values rather than placeholders.

    A section whose service answered with a shell still rendered a full page:
    six of eight market metrics as em-dashes, no tables and no read, or a
    schools card carrying nothing but the district name. In a document going
    to a committee an empty page reads as "there is nothing here" rather than
    "the source had nothing to give", so a section has to earn its page.
    """
    real = 0
    for v in values:
        if v in (None, "", "-", "—", 0, "0"):
            continue
        real += 1
    return real >= need


def _map_market(r, cb):
    if not cb:
        return
    agg = cb.get("aggregate") or {}
    starts = _first(agg, "annual_starts") or cb.get("ring_annual_starts")
    closings = _first(agg, "annual_closings")
    vdls = _first(agg, "vdls", "finished_lots")
    radius = _n(cb.get("radius_mi"))

    # Months of lot supply is derived when the field is absent rather than
    # printed as a dash. It is a definition, not a lookup: finished lots over
    # monthly closings. The first run showed "-" here beside 3,349 finished
    # lots and 1,857 annual closings, which is 21.6 months sitting in plain
    # sight on the same row.
    mos = _n(cb.get("months_lot_supply"))
    if mos is None and _n(vdls) and _n(closings):
        mos = _n(vdls) / (_n(closings) / 12.0)

    # price_band is {"min": ..., "max": ...}, not a string.
    pb = cb.get("price_band")
    if isinstance(pb, dict):
        # Exact figures here, not the abbreviated form used elsewhere: a price
        # band is a range someone will quote, and "$185k - $1M" loses the ends.
        lo, hi = _n(pb.get("min")), _n(pb.get("max"))
        price_range = (f"${lo:,.0f} - ${hi:,.0f}" if lo and hi
                       else (f"${(lo or hi):,.0f}" if (lo or hi) else "-"))
    else:
        price_range = str(pb) if pb else "-"

    ctx = {
        "subtitle": ("New-home velocity, lot supply, future pipeline and product depth"
                     + (f" within the {radius:g}-mile competitive ring" if radius else "")),
        "starts": num(starts),
        "closings": num(closings),
        "vdl": num(vdls),
        "mos": (f"{mos:,.1f}" if mos is not None else "-"),
        "mos_hot": bool(mos is not None and mos >= 18),
        "uc": num(_first(agg, "under_construction")),
        # The payload calls it complete_vacant; there is no finished_vacant.
        "fv": num(_first(agg, "complete_vacant")),
        "future": num(_first(agg, "futures")),
        "price_range": price_range,
    }
    bits = []
    if cb.get("quarter_label"):
        bits.append(str(cb["quarter_label"]))
    for key, word in (("community_count", "communities"),
                      ("active_count", "actively closing"),
                      ("builder_count", "builders")):
        if _n(cb.get(key)) is not None:
            bits.append(f"{_n(cb[key]):,.0f} {word}")
    ctx["context"] = "  |  ".join(bits)

    qs = cb.get("quarter_series") or []
    if qs:
        ctx["quarterly_chart"] = render_quarter_chart([
            {"label": q.get("label") or q.get("quarter"),
             "starts": _first(q, "starts", "start"),
             "closings": _first(q, "closings", "closing")} for q in qs])

    # lot_bands ship pre-aggregated: label ("40-50 FF"), avg_price, avg_ppsf.
    # Reading prices/ppsfs/lot_width_ff -- the internal accumulator's names --
    # put a dash in Width, Avg price and $/SF on every row.
    bands = []
    for b in (cb.get("lot_bands") or []):
        ppsf = _n(b.get("avg_ppsf"))
        bands.append({
            "width": str(b.get("label") or "-")[:12],
            "lots": num(b.get("lots")),
            "avg_price": money(b.get("avg_price")),
            "psf": (f"${ppsf:,.0f}" if ppsf is not None else "-"),
            "_sort": _n(b.get("lots")) or 0,
        })
    # Budget the sheet before choosing row counts. Two KPI rows, the quarterly
    # chart and the section chrome already spend 382 of 710 points; two full
    # tables plus a read block came to 793 and spilled.
    used = SECTION_CHROME_PT + 2 * KPI_ROW_PT
    if ctx.get("quarterly_chart"):
        used += _image_pt(6.9, 1.45)   # a shorter chart buys three table rows
    used += CARD_CHROME_PT + 3 * READ_LINE_PT          # the market read
    # With sections flowing, a table that runs past the page foot costs an inch
    # of the next one rather than a whole sheet, so the budget is a floor for
    # readability rather than a hard ceiling. Depth matters more than tidiness
    # here: a six-row builder table says little about a 36-builder market.
    per_table = max(0.0, (PAGE_PT - used) / 2.0)
    n_rows = int((per_table - TABLE_CHROME_PT) // TABLE_ROW_PT)
    n_rows = max(6, min(9, n_rows))
    ctx["lot_bands"] = sorted(bands, key=lambda x: -x["_sort"])[:n_rows]
    ctx["_row_budget"] = max(8, n_rows)

    blds = []
    for b in (cb.get("builders") or []):
        prices = [p for p in (b.get("prices") or []) if _n(p)]
        avg = _n(b.get("avg_price"))
        if avg is None and prices:
            avg = sum(prices) / len(prices)
        blds.append({
            "name": b.get("name") or "-",
            "starts": num(b.get("est_annual_starts")),
            "lots": num(b.get("lots")),
            "avg_price": money(avg),
            "_sort": _n(b.get("est_annual_starts")) or 0,
        })
    blds.sort(key=lambda x: -x["_sort"])
    nb = ctx.get("_row_budget") or 6
    ctx["builders"] = blds[:nb]
    if len(blds) > nb:
        n = _n(cb.get("builder_count")) or len(blds)
        ctx["builder_note"] = f"{n:,.0f} builders active within the competitive study area."

    read = []
    st, cl = _n(starts), _n(_first(agg, "annual_closings"))
    if st and cl:
        read.append(
            f"Starts ({st:,.0f}) and closings ({cl:,.0f}) are close, so absorption is "
            "keeping pace with delivery."
            if cl >= st * 0.95 else
            f"Starts ({st:,.0f}) run ahead of closings ({cl:,.0f}), a signal of "
            "inventory build in the ring.")
    if mos is not None:
        read.append(f"{mos:,.1f} months of lot supply is elevated; phase conservatively."
                    if mos >= 18 else
                    f"{mos:,.1f} months of lot supply indicates a constrained lot market.")
    # Only claim a deepest product when it is actually named. The first run
    # printed "The deepest product is -, which is where the preliminary mix
    # should concentrate", which is worse than saying nothing.
    top_band = next((b for b in ctx["lot_bands"] if b["width"] not in ("-", "")), None)
    if top_band:
        read.append(f"The deepest product is {top_band['width']} at "
                    f"{top_band['lots']} lots, which is where the preliminary mix "
                    "should concentrate.")
    ctx["read"] = read or None
    if not (_substantial(ctx["starts"], ctx["closings"], ctx["vdl"], ctx["mos"],
                         ctx["uc"], ctx["fv"], ctx["future"], need=3)
            or ctx["lot_bands"] or ctx["builders"]):
        return                       # a page of dashes helps nobody
    r["market"] = ctx


def _lot_range(c):
    """Lot widths as a range, not the dict the payload actually carries.

    lot_type_range is {"min": .., "max": ..}; printing it produced literal
    "{'max': 65," in the table. Values above 200 are square footage rather
    than frontage and are left out.
    """
    v = c.get("lot_type_range")
    if isinstance(v, dict):
        lo, hi = _n(v.get("min")), _n(v.get("max"))
        lo = lo if (lo and lo < 200) else None
        hi = hi if (hi and hi < 200) else None
        if lo and hi and lo != hi:
            return f"{lo:,.0f}-{hi:,.0f}"
        if lo or hi:
            return f"{(lo or hi):,.0f}"
        return "-"
    if isinstance(v, (list, tuple, set)):
        nums = sorted({int(x) for x in v if _n(x) and _n(x) < 200})
        return (f"{nums[0]}-{nums[-1]}" if len(nums) > 1 else
                (str(nums[0]) if nums else "-"))
    ff = c.get("lot_types_ff")
    if isinstance(ff, (list, tuple, set)):
        nums = sorted({int(x) for x in ff if _n(x) and _n(x) < 200})
        return (f"{nums[0]}-{nums[-1]}" if len(nums) > 1 else
                (str(nums[0]) if nums else "-"))
    return str(v or "-")[:12]


def _map_comps(r, cb, comp_map=None):
    """The competitive set, with the dead entries left out.

    The first run listed the fourteen nearest communities whatever their state,
    so half the table was rows of zeroes -- subdivisions with no starts, no
    closings and no pipeline tell a reader nothing about the competitive
    environment and push the ones that matter off the page. A community earns
    a row only if it is actually doing something.

    Field names are the endpoint's: lot_type_range, months_lot_supply and
    pct_built_out. The first attempt guessed lot_widths / months_supply /
    pct_built and printed a dash in all three columns for every row.
    """
    comms = (cb or {}).get("communities") or []
    if not comms:
        return

    def dist(c):
        return _n(c.get("distance_mi")) or 999.0

    def activity(c):
        return ((_n(c.get("annual_closings")) or 0) + (_n(c.get("annual_starts")) or 0)
                + (_n(c.get("vdls")) or 0) + (_n(c.get("futures")) or 0))

    live = [c for c in comms if activity(c) > 0]
    dropped = len(comms) - len(live)
    # Nearest first, but a community with real velocity outranks a closer one
    # that is only sitting on future lots.
    live.sort(key=lambda c: (0 if (_n(c.get("annual_closings")) or 0) > 0 else 1, dist(c)))

    # Same budget: the competitor map and a KPI row come first, the table gets
    # what is left rather than a fixed count that spilled onto a near-empty
    # page.
    used = SECTION_CHROME_PT + KPI_ROW_PT + CARD_CHROME_PT + READ_LINE_PT * 2
    if comp_map:
        used += _image_pt(7.4, 3.2)
    # The competitive set is the point of this page; show it properly.
    n_rows = max(12, _rows_that_fit(used, 14, min_rows=12))

    rows = []
    for c in live[:n_rows]:
        bl = c.get("builders")
        if isinstance(bl, (list, set, tuple)):
            bl = ", ".join(sorted(str(x) for x in bl)[:2])
        # A community with two closings a year and a hundred finished lots
        # computes to hundreds of months. That is arithmetic, not a market
        # signal, and it made "highest MOS 374.2" the headline figure.
        # Months of supply divides finished lots by monthly closings, so a
        # community selling seven homes a year produces 41 months and one
        # selling two produces hundreds. Below roughly one sale a month the
        # figure is an artefact of a small denominator rather than a read on
        # supply, so it is left blank instead of printed.
        mos = _n(c.get("months_lot_supply"))
        if mos is not None and ((_n(c.get("annual_closings")) or 0) < 12 or mos > 120):
            mos = None
        rows.append({
            "name": str(c.get("name") or "-")[:26],
            "distance": miles(c.get("distance_mi"), c.get("direction")),
            "builders": (str(bl)[:24] if bl else "-"),
            "widths": _lot_range(c),
            "closings": num(c.get("annual_closings")),
            "starts": num(c.get("annual_starts")),
            "mos": (f"{mos:,.1f}" if mos is not None else "-"),
            "pipeline": num(c.get("futures")),
            "built": pct(c.get("pct_built_out"), dp=0),
        })
    if not rows:
        return

    kpis = []
    nearest = min((dist(c) for c in live), default=None)
    if nearest and nearest < 900:
        kpis.append({"label": "Nearest active comp", "value": miles(nearest)})
    for label, key in (("Top closings", "annual_closings"),
                       ("Largest pipeline", "futures")):
        v = max((_n(c.get(key)) or 0 for c in live), default=0)
        if v:
            kpis.append({"label": label, "value": num(v)})
    mx = [_n(c.get("months_lot_supply")) for c in live
          if _n(c.get("months_lot_supply")) is not None
          and (_n(c.get("annual_closings")) or 0) >= 12
          and _n(c.get("months_lot_supply")) <= 120]
    if mx:
        kpis.append({"label": "Highest MOS", "value": f"{max(mx):,.1f}"})

    note = []
    if len(live) > len(rows):
        note.append(f"{len(live):,} active communities in the study area; the "
                    f"{len(rows)} most relevant are listed.")
    if dropped:
        note.append(f"{dropped:,} with no starts, closings, lots or pipeline omitted.")
    r["comps"] = {
        "map": comp_map, "rows": rows, "kpis": kpis[:4],
        "note": " ".join(note) or None,
        "read": ("nearest communities are already proving demand, but several early-stage "
                 "projects carry large pipelines and long remaining buildout."),
    }


def _district_block(sd):
    """One district's headline figures."""
    win = (sd.get("growth") or {}).get("windows") or {}
    growth = []
    for key, label in (("5_year", "Enrolment / 5 yr"), ("10_year", "10-year"),
                       ("20_year", "20-year"), ("all_time", "All-time")):
        blk = win.get(key) or {}
        p = _n(blk.get("total_pct"))
        if p is not None:
            cagr = _n(blk.get("cagr_pct"))
            growth.append({"label": label, "value": f"{p:+,.1f}%",
                           "note": (f"{cagr:+,.2f}% CAGR" if cagr is not None else "")})
    tea = sd.get("tea") or {}
    rating = tea.get("overall_rating")
    score = _n(tea.get("overall_score"))
    return {
        "district": str(sd.get("name") or "School district").upper(),
        "rating": (f"{rating} / {score:,.0f}" if rating and score is not None
                   else (str(rating) if rating else "-")),
        "rating_note": ("TEA " + str(sd.get("tea_year") or "")).strip(),
        "growth": growth[:3],
        "enrollment": num(sd.get("enrollment")),
        "school_count": num(sd.get("schools_count")),
        "teachers": num(sd.get("teachers_fte")),
        "ratio": (f"{_n(sd.get('student_teacher_ratio')):,.1f} : 1"
                  if _n(sd.get("student_teacher_ratio")) is not None else "-"),
        "_enrol": _n(sd.get("enrollment")) or 0,
    }


def _map_schools(r, sd):
    """Every district the tract touches, not just the first.

    A tract on a district line sits in two; the app's own schools card shows
    both, and taking the first made the export disagree with the screen. The
    largest district by enrolment leads and carries the trend chart; the
    others follow as their own cards, and the campus table merges all of them
    by distance with a district column so a reader can tell them apart.
    """
    payloads = sd if isinstance(sd, list) else ([sd] if sd else [])
    payloads = [p for p in payloads if isinstance(p, dict) and not p.get("error")]
    if not payloads:
        return

    blocks, per_district = [], []
    for p in payloads:
        blk = _district_block(p)
        mine = []
        for c in (p.get("schools") or []):
            mine.append({
                "name": str(c.get("name") or "-")[:28],
                "district": blk["district"].title().replace(" Isd", " ISD"),
                "tea": str(c.get("tea_rating") or "Not rated")[:12],
                "level": str(c.get("level") or "-")[:18],
                "enrollment": num(c.get("enrollment")),
                "distance": miles(c.get("distance_mi"), c.get("direction")),
                "_d": _n(c.get("distance_mi")) or 999.0,
            })
        mine.sort(key=lambda c: c["_d"])
        per_district.append((blk, mine))
        if (_substantial(blk["enrollment"], blk["school_count"], blk["teachers"],
                         blk["ratio"], blk["rating"], need=2) or blk["growth"]):
            blocks.append(blk)

    # Take the nearest few from EACH district before merging. A plain distance
    # sort filled all ten rows with the nearer district, so the second one
    # vanished from a table that names it in its own column.
    per = 6 if len(per_district) > 1 else 10
    campuses = [c for _, mine in per_district for c in mine[:per]]
    campuses.sort(key=lambda c: c["_d"])
    if not blocks and not campuses:
        return
    blocks.sort(key=lambda b: -b["_enrol"])
    out = dict(blocks[0]) if blocks else {}
    out.update({
        "others": blocks[1:],
        "multi": len(blocks) > 1,
        "trend_chart": render_enrollment_chart(
            [(str(p.get("name") or ""), p.get("enrollment_trend")) for p in payloads]),
        "campuses": campuses[:12],
    })
    r["schools"] = out


ACS = [
    ("B01003_001E", "Population", num), ("B11001_001E", "Households", num),
    ("B19013_001E", "Median HH income", money), ("B19301_001E", "Per-capita income", money),
    ("B25077_001E", "Median home value", money), ("B25064_001E", "Median rent", money),
    ("B25001_001E", "Housing units", num), ("B23025_004E", "Employed", num),
]


def _map_demographics(r, mk):
    if not mk or mk.get("error"):
        return
    src = mk.get("current") if isinstance(mk.get("current"), dict) else mk
    rows = []
    for code, label, fmt in ACS:
        v = src.get(code, mk.get(code))
        if _n(v) is None:
            continue
        rows.append({"label": label, "value": fmt(v), "note": ""})
    occ, vac = _n(src.get("B25002_002E")), _n(src.get("B25002_003E"))
    if occ and vac is not None and (occ + vac):
        rows.append({"label": "Vacancy rate", "value": pct(vac / (occ + vac) * 100), "note": ""})
    own, rent = _n(src.get("B25003_002E")), _n(src.get("B25003_003E"))
    if own and rent is not None and (own + rent):
        rows.append({"label": "Owner occupancy", "value": pct(own / (own + rent) * 100, dp=0),
                     "note": ""})
    tot = _n(src.get("B15003_001E"))
    ba = sum(_n(src.get(k)) or 0 for k in
             ("B15003_022E", "B15003_023E", "B15003_024E", "B15003_025E"))
    if tot and ba:
        rows.append({"label": "Bachelor's degree +", "value": pct(ba / tot * 100), "note": ""})
    if rows:
        r["demographics"] = {
            "title": str(mk.get("county_name") or "County") + " - demand context",
            "rows": rows[:12],
        }


def _map_access(r, roads, am, roads_map=None):
    am = am or {}
    cards = []
    hw = [h for h in (am.get("highways") or []) if _n(h.get("distance_mi")) is not None]
    if hw:
        h = min(hw, key=lambda x: _n(x.get("distance_mi")))
        # OSM packs concurrent routes into one ref separated by semicolons
        # ("US 290;TX 6"), which reads as a typo on a printed page.
        ref = str(h.get("ref") or h.get("name") or "-").replace(";", " / ")
        cards.append({"label": "Nearest freeway", "value": ref[:20],
                      "note": miles(h.get("distance_mi"), h.get("direction"))})
    if _n(am.get("nearest_ramp_mi")) is not None:
        cards.append({"label": "Nearest on-ramp", "value": miles(am["nearest_ramp_mi"]),
                      "note": "closest interchange"})
    ap = [a for a in (am.get("airports") or []) if _n(a.get("distance_mi")) is not None]
    if ap:
        comm = [a for a in ap if a.get("commercial")] or ap
        a0 = min(comm, key=lambda x: _n(x.get("distance_mi")))
        cards.append({"label": "Airport", "value": str(a0.get("iata") or a0.get("name") or "-")[:16],
                      "note": miles(a0.get("distance_mi"), a0.get("direction"))})
    if not cards and not roads:
        return

    projects, note, funded = [], None, None
    if roads:
        planned = (roads.get("planned") or []) + (roads.get("programmed") or [])

        def yr(p):
            return _n(_first(p, "let_year", "start_year")) or 9999

        # Budget the sheet: three access cards, the read block, and the map
        # when there is one. The table takes what is left.
        used = SECTION_CHROME_PT + KPI_ROW_PT + CARD_CHROME_PT + 3 * READ_LINE_PT
        if roads_map:
            used += _image_pt(7.4, 3.4)
        rows = sorted(planned, key=yr)[:max(6, _rows_that_fit(used, 12, min_rows=6))]
        for p in rows:
            frm = _first(p, "from", default="") or ""
            to = _first(p, "to", default="") or ""
            projects.append({
                "road": str(_first(p, "roadway", "highway", default="-"))[:16],
                "work": str(_first(p, "work", "description", default="-"))[:38],
                "limits": (f"{frm} - {to}".strip(" -") or "-")[:38],
                "start": num(_first(p, "let_year", "start_year")),
            })
        near = _n((roads.get("counts") or {}).get("near_term"))
        if near:
            funded = f"{near:,.0f} funded / dated in the next four years"
        if len(planned) > len(rows):
            note = (f"{len(planned):,} programmed projects within "
                    f"{roads.get('radius_mi', '')} mi; nearest-term shown.")

    read = []
    if cards and cards[0]["label"] == "Nearest freeway":
        read.append(f"{cards[0]['value']} at {cards[0]['note']} gives a credible regional "
                    "access story.")
    if projects:
        read.append("Programmed capacity work on the surrounding network reinforces the "
                    "longer-term growth corridor.")
        read.append("Off-site road obligations and timing should be tied directly into "
                    "phase-level development underwriting.")
    r["access"] = {"cards": cards[:4], "projects": projects, "funded_note": funded,
                   "note": note, "read": read or None, "map": roads_map}


AMENITY_GROUPS = [
    ("Grocery & retail", ("grocery_stores",)),
    ("Healthcare", ("hospitals",)),
    ("Pharmacies & parks", ("pharmacies", "parks")),
    ("Fuel & convenience", ("fuel",)),
]


def _map_amenities(r, am):
    if not am:
        return
    groups = []
    for title, keys in AMENITY_GROUPS:
        items = []
        for k in keys:
            for x in (am.get(k) or []):
                if _n(x.get("distance_mi")) is None:
                    continue
                items.append({"name": str(x.get("name") or x.get("brand") or "-")[:32],
                              "distance": miles(x.get("distance_mi"), x.get("direction")),
                              "_d": _n(x.get("distance_mi"))})
        items.sort(key=lambda i: i["_d"])
        if items:
            # NOT "items": in Jinja, `group.items` resolves to dict.items --
            # the bound method -- before it falls back to the key, and the
            # template then tries to iterate a builtin_function_or_method.
            groups.append({"title": title, "places": items[:6]})
    nearest = []
    for label, key in (("Nearest fuel", "fuel"), ("Nearest pharmacy", "pharmacies"),
                       ("Nearest grocery", "grocery_stores"),
                       ("Nearest hospital", "hospitals")):
        pool = [x for x in (am.get(key) or []) if _n(x.get("distance_mi")) is not None]
        if pool:
            x = min(pool, key=lambda i: _n(i["distance_mi"]))
            nearest.append({"label": label, "value": miles(x["distance_mi"]),
                            "note": str(x.get("name") or x.get("brand") or "")[:20]})
    if groups or nearest:
        r["amenities"] = {"groups": groups, "nearest": nearest[:4]}


NEWS_MAX_AGE_DAYS = 550          # roughly eighteen months

# Why a story matters to an acquisition, keyed off what it is about. A headline
# on its own makes a reader do the work; the point of this section is to say
# what the signal is.
NEWS_SIGNALS = [
    (("jobs", "hiring", "employment", "workforce", "manufactur", "plant"),
     "Employment signal supporting regional household growth."),
    (("distribution", "industrial", "warehouse", "logistics", "data center"),
     "Industrial absorption -- a demand driver for nearby rooftops."),
    (("hospital", "medical", "health", "clinic"),
     "Healthcare investment adds an institutional growth signal."),
    (("highway", "road", "interchange", "corridor", "expansion", "infrastructure",
      "utility", "water"),
     "Infrastructure investment affecting access and development timing."),
    (("home", "housing", "subdivision", "master-planned", "master planned",
      "lots", "builder", "residential", "development"),
     "Reinforces the residential growth thesis and the future-supply picture."),
    (("school", "isd", "enrollment", "campus"),
     "District growth pressure, a proxy for household formation."),
    (("acres", "land", "ranch", "acquisition", "sold", "purchase"),
     "Land transaction comparable to the subject."),
]


def _news_relevance(title):
    """Match on whole words, not substrings.

    A plain `"water" in title` tagged "New Homes Now Selling in Attwater" as
    infrastructure, because the place name contains the keyword. Place names
    swallow short keywords constantly, so every term is anchored.
    """
    t = " " + re.sub(r"[^a-z0-9 ]+", " ", (title or "").lower()) + " "
    for words, why in NEWS_SIGNALS:
        for w in words:
            # prefix match so "manufactur" still catches manufacturing
            if re.search(r"\b" + re.escape(w), t):
                return why
    return None


def _headline(title, source):
    """Google News appends " - Outlet" to every title.

    The outlet is already printed on the line above, so the raw title read
    "...after first day of school - Houston Chronicle" directly under
    "AUG 2026 - HOUSTON CHRONICLE".
    """
    t = str(title or "").strip()
    src = str(source or "").strip()
    if src and t.lower().endswith(" - " + src.lower()):
        t = t[: -(len(src) + 3)].rstrip()
    else:
        # Some feeds carry a slightly different outlet name than the source
        # field; fall back to trimming a short trailing " - Something".
        head, sep, tail = t.rpartition(" - ")
        if sep and 0 < len(tail) <= 34 and not tail.endswith("."):
            t = head.rstrip()
    return t[:150]


def _map_news(r, nw):
    """Recent, relevant stories only.

    The feed returns whatever the query matched, oldest included, in query
    order. A report dated this month carrying a two-year-old headline reads as
    stale research, so anything past NEWS_MAX_AGE_DAYS is dropped and the rest
    are newest first.
    """
    import datetime as _dt
    from email.utils import parsedate_to_datetime

    stories = (nw or {}).get("stories") or []
    if not stories:
        return
    now = _dt.datetime.now(_dt.timezone.utc)
    dated = []
    for st in stories:
        raw = st.get("published")
        when = None
        if raw:
            try:                                  # RSS: RFC-2822
                when = parsedate_to_datetime(str(raw))
            except Exception:
                try:                              # or ISO
                    when = _dt.datetime.fromisoformat(str(raw)[:19])
                except Exception:
                    when = None
        if when is not None and when.tzinfo is None:
            when = when.replace(tzinfo=_dt.timezone.utc)
        # An undated story is kept but sorts last -- dropping it would hide a
        # relevant item purely because the feed omitted a timestamp.
        if when is not None and (now - when).days > NEWS_MAX_AGE_DAYS:
            continue
        dated.append((when, st))
    dated.sort(key=lambda p: (p[0] is not None, p[0]), reverse=True)

    # The aggregator runs several overlapping queries, so one event arrives as
    # three near-identical headlines from three outlets -- "Waller ISD faces
    # registration backlog", "Waller ISD experiences application surge", "'All
    # hands on deck': Waller ISD experiences application..." all in one list.
    # Keeping the first of each cluster is what a person would do.
    STOP = {"the", "a", "an", "of", "in", "at", "to", "for", "and", "on", "as",
            "with", "after", "new", "its", "is", "are", "from", "by"}

    def sig(title):
        return {w for w in re.findall(r"[a-z0-9]+", (title or "").lower())
                if len(w) > 2 and w not in STOP}

    # Token overlap alone does not cluster these: "Waller ISD faces
    # registration backlog", "Waller ISD experiences application backlog" and
    # "'All hands on deck': Waller ISD experiences application backlog" share
    # only two or three words once stop words are gone. Capping each signal
    # category at two is the rule that actually produces a varied page.
    kept, per_topic = [], {}
    for when, st in dated:
        s = sig(st.get("title"))
        if not s:
            continue
        if any(len(s & k) / max(len(s | k), 1) >= 0.34 for k in kept):
            continue                       # same story, different outlet
        why = _news_relevance(st.get("title"))
        topic = why or "other"
        if per_topic.get(topic, 0) >= 2:
            continue                       # already covered this signal
        per_topic[topic] = per_topic.get(topic, 0) + 1
        kept.append(s)
        r.setdefault("news", []).append({
            "headline": _headline(st.get("title"), st.get("source")),
            "source": str(st.get("source") or "")[:34],
            "date": (when.strftime("%b %Y") if when else ""),
            "why": why,
        })
        if len(r["news"]) >= 6:
            break


def _macro_value(cur, fmt):
    """Format a FRED reading using the series' declared unit.

    The units are not decorative: TXNA and TXPOP are published in THOUSANDS,
    so printing them raw gave "14,468" for 14.5 million jobs and "31,710" for
    the population of Texas. TXNQGSP is in billions.
    """
    fmt = str(fmt or "").lower()
    if fmt == "percent":
        return f"{cur:,.2f}%"
    if fmt == "dollars":
        return f"${cur:,.0f}"
    if fmt == "dollars_b":
        # The series is labelled "$B" but FRED publishes TXNQGSP in MILLIONS
        # of dollars, so dividing by a thousand printed Texas GDP as
        # "$3,034.82T". Choose the unit from the magnitude instead of trusting
        # the label, which is the thing that was wrong.
        if abs(cur) >= 1e6:
            return f"${cur/1e6:,.2f}T"
        if abs(cur) >= 1e3:
            return f"${cur/1e3:,.0f}B"
        return f"${cur:,.0f}M"
    if fmt == "thousands":
        v = cur * 1000.0
        return (f"{v/1e6:,.1f}M" if abs(v) >= 1e6 else f"{v:,.0f}")
    if fmt == "count":
        return f"{cur:,.0f}"
    return f"{cur:,.1f}"          # index and anything unrecognised


def _macro_label(label):
    """Shorten on a word boundary. Cutting at a fixed width produced
    "TEXAS HOME PRICE INDEX (FH" and "HOUSTON MSA HOME PRICE IND"."""
    txt = re.sub(r"\s*\([^)]*\)", "", str(label or "")).strip()
    txt = (txt.replace("Texas ", "TX ").replace("Houston MSA ", "Houston ")
              .replace("per capita personal income", "per-capita income")
              .replace("total private permits", "private permits")
              .replace("non-farm employment", "non-farm jobs")
              .replace("resident population", "population")
              .replace("median household income", "median HH income")
              .replace("leading economic index", "leading econ. index")
              .replace("US Treasury yield", "Treasury yield"))
    if len(txt) > 28:
        cut = txt[:28].rsplit(" ", 1)[0]
        txt = cut or txt[:28]
    return txt.upper()


def _map_macro(r, fr):
    out = []
    for i in ((fr or {}).get("indicators") or []):
        d = i.get("data") or {}
        cur = _n(d.get("current"))
        if d.get("error") or cur is None:
            continue
        note = []
        if _n(d.get("yoy_pct")) is not None:
            note.append(f"YoY {_n(d['yoy_pct']):+,.1f}%")
        if _n(d.get("five_year_total_pct")) is not None:
            note.append(f"5Y {_n(d['five_year_total_pct']):+,.1f}%")
        out.append({"label": _macro_label(i.get("label") or i.get("series_id")),
                    "value": _macro_value(cur, i.get("format")),
                    "note": "   ".join(note)})
    if out:
        r["macro"] = out[:12]
