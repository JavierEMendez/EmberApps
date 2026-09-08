"""Render the report locally so layout can be checked without a round trip.

WeasyPrint needs GTK and will not install on Windows, so every page-break
change was a hypothesis someone else had to test by exporting. Chrome is
already on this machine and honours the same @page, break-inside and
break-after properties, which is what these bugs live in. It is not
byte-identical to WeasyPrint -- fonts and hyphenation differ -- but a section
that strands its heading here strands it there.

    python devtools/render_local.py out.pdf
"""
import io
import json
import os
import subprocess
import sys
import tempfile

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# run from devtools/, import from the repo root
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _cad_tracts(accounts):
    """Tracts straight from HCAD, for a machine with no parcel cache."""
    import requests
    from shapely.geometry import Polygon, mapping
    from shapely.ops import unary_union
    out = []
    for acct in accounts:
        try:
            r = requests.get(
                "https://www.gis.hctx.net/arcgis/rest/services/HCAD/Parcels/"
                "MapServer/0/query",
                params={"where": f"HCAD_NUM='{acct}'", "outFields": "*",
                        "returnGeometry": "true", "outSR": 4326, "f": "json"},
                timeout=40)
            f = (r.json().get("features") or [None])[0]
            if not f:
                continue
            rings = [Polygon(g) for g in f["geometry"]["rings"] if len(g) >= 4]
            g = unary_union([p for p in rings if p.is_valid and not p.is_empty])
            a = f["attributes"]
            out.append({"prop_id": acct, "owner_name": a.get("owner_name_1"),
                        "acres": float(str(a.get("Acreage") or "0").split()[0] or 0),
                        "county": "Harris", "geometry": mapping(g)})
        except Exception as e:
            print(f"  HCAD {acct} skipped: {e}")
    return out


def build_html():
    import jinja2
    import acq_gis
    import acq_parcels as pc
    import acq_report
    from shapely.geometry import shape as shp_shape
    from shapely.ops import unary_union

    here = os.path.dirname(os.path.abspath(__file__))
    fixture = os.path.join(here, "report_fixture.json")
    data = json.load(io.open(fixture, encoding="utf-8"))

    try:
        r = pc.find_parcels_by_owner(data["owner"], include_geometry=True)
        ps = [p for p in (r.get("parcels") or []) if p.get("geometry")]
    except Exception:
        ps = []
    if not ps:
        # The parcel cache is a few GB on a Railway volume and is not on a dev
        # machine, so fall back to the appraisal district's own service for the
        # accounts the fixture names. Same geometry the analysis reconciles to.
        ps = _cad_tracts(data.get("cad_accounts") or [])
    if not ps:
        raise SystemExit("no tract geometry: cache is empty and the fixture "
                         "lists no cad_accounts to fall back on")
    proj = {"name": data["name"],
            "tracts": [{"prop_id": str(p["prop_id"]), "owner_name": p.get("owner_name"),
                        "acres": p.get("acres"), "county": p.get("county"),
                        "geometry": p["geometry"]} for p in ps]}
    analysis = acq_gis.run_analysis(proj)
    ctx = acq_report.build_context(proj, analysis, data["payloads"], data["elevation"])
    try:
        u, tl = acq_gis.analysis_geometry(proj, analysis)
        ctx["site_map"] = acq_report.render_site_map(
            u, analysis.get("constraint_geoms"), tl)
    except Exception as e:
        print("site map skipped:", e)
    env = jinja2.Environment(loader=jinja2.FileSystemLoader("templates"), autoescape=True)
    return env.get_template("acq_report.html").render(r=ctx)


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "report_local.pdf"
    html = build_html()
    tmp = os.path.join(tempfile.gettempdir(), "_acq_report_local.html")
    io.open(tmp, "w", encoding="utf-8").write(html)
    subprocess.run([CHROME, "--headless", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={out}", "file:///" + tmp.replace("\\", "/")],
                   check=True, capture_output=True, timeout=180)
    print(f"rendered {len(html):,} chars -> {out}")


if __name__ == "__main__":
    main()
