"""Step 1 of v2: assign every camera to its New York county (point-in-polygon).

Usage: python assign_counties.py counties-10m.json ALPR_Camera_Cleaned_NY.csv with_county.csv
  counties-10m.json: U.S. Census county boundaries (1:10m) from the us-atlas npm package.
Also writes ny_counties.csv (all 62 NY counties), used by build_workbook.py.
"""
import json, sys
import pandas as pd
from shapely.geometry import shape, Polygon, MultiPolygon, Point
from shapely.strtree import STRtree

TOPO, CSV, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
t = json.load(open(TOPO))
sc, tr = t["transform"]["scale"], t["transform"]["translate"]

def decode_arc(arc):
    x = y = 0; out = []
    for dx, dy in arc:
        x += dx; y += dy
        out.append((x * sc[0] + tr[0], y * sc[1] + tr[1]))
    return out
arcs = [decode_arc(a) for a in t["arcs"]]

def ring(idx):
    pts = []
    for i in idx:
        a = arcs[i] if i >= 0 else arcs[~i][::-1]
        pts.extend(a if not pts else a[1:])
    return pts

def geom(g):
    if g["type"] == "Polygon":
        r = [ring(x) for x in g["arcs"]]; return Polygon(r[0], r[1:])
    if g["type"] == "MultiPolygon":
        return MultiPolygon([Polygon(ring(p[0]), [ring(h) for h in p[1:]]) for p in g["arcs"]])

ny = [g for g in t["objects"]["counties"]["geometries"] if str(g["id"]).startswith("36")]
polys = [geom(g) for g in ny]; names = [g["properties"]["name"] for g in ny]
tree = STRtree(polys)

d = pd.read_csv(CSV, encoding="utf-8-sig")
def county(lat, lon):
    p = Point(lon, lat)
    hits = tree.query(p, predicate="intersects")
    if len(hits): return names[hits[0]]
    i = tree.nearest(p)  # coastal points just outside the low-res outline
    return names[i] + "*"
d["County"] = [county(a, b) for a, b in zip(d.LAT, d.LONG)]
d["near_only"] = d.County.str.endswith("*"); d["County"] = d.County.str.rstrip("*")
nyc = {"New York", "Kings", "Queens", "Bronx", "Richmond"}
d["Region"] = d.County.map(lambda c: "NYC" if c in nyc else "Long Island" if c in {"Nassau", "Suffolk"} else "Rest of State")
d["Unknown_Op"] = d.Operator.eq("Unknown")
d.to_csv(OUT, index=False)
pd.DataFrame({"County": sorted(names)}).to_csv("ny_counties.csv", index=False)
print(f"{len(d)} cameras; {int(d.near_only.sum())} coastal points assigned to the nearest county")

