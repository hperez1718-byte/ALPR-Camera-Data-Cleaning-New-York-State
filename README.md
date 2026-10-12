Flock/ALPR Camera Data Extraction Script
=========================================
Pulls fresh ALPR (automatic license plate reader) camera location data
for New York State from the DeFlock CDN.

Data source: DeFlock (https://deflock.me)
  - Underlying data: OpenStreetMap (crowdsourced, ODbL license)
  - DeFlock publishes pre-built static JSON tiles on their CDN
  - Index: https://cdn.deflock.me/regions/index.json
  - Tiles:  https://cdn.deflock.me/regions/{lat}/{lon}.json

Method:
  1. Fetch the region index to get tile URLs
  2. Download the tile covering the northeastern US (40/-80)
  3. Filter to New York State bounding box (lat 40.4-45.1, lon -79.9 to -71.7)
  4. Flatten OSM tags into columns and export to Excel

Usage:
  python3 pull_flock_cameras.py

Output:
  flock-cameras-ny-fresh.xlsx — ID, LAT, LON, Manufacturer, Operator, Direction
"""

import json
import urllib.request
import urllib.parse
import openpyxl

INDEX_URL = "https://cdn.deflock.me/regions/index.json"
TILE_URL = "https://cdn.deflock.me/regions/40/-80.json"
HEADERS = {"User-Agent": "ALPR-DataSchool-Project/1.0"}

# New York State approximate bounding box
NY_LAT_MIN, NY_LAT_MAX = 40.4, 45.1
NY_LON_MIN, NY_LON_MAX = -79.9, -71.7


def fetch_json(url):
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.load(r)


def main():
    print("Fetching DeFlock region index...")
    index = fetch_json(INDEX_URL)
    print(f"  Tile size: {index['tile_size_degrees']} degrees")
    print(f"  Regions available: {len(index['regions'])}")

    print("Downloading northeastern US tile (40/-80)...")
    cameras = fetch_json(TILE_URL)
    print(f"  Cameras in tile: {len(cameras)}")

    print("Filtering to New York State...")
    ny = [
        c for c in cameras
        if NY_LAT_MIN <= c["lat"] <= NY_LAT_MAX
        and NY_LON_MIN <= c["lon"] <= NY_LON_MAX
    ]
    print(f"  Cameras in NY State: {len(ny)}")

    print("Building Excel workbook...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "cameras"
    ws.append(["ID", "LAT", "LON", "Manufacturer", "Operator", "Direction"])
    for c in ny:
        tags = c.get("tags", {})
        ws.append([
            c["id"],
            c["lat"],
            c["lon"],
            tags.get("manufacturer", "Unknown"),
            tags.get("operator", "Unknown"),
            tags.get("direction", "Unknown"),
        ])

    out = "flock-cameras-ny-fresh.xlsx"
    wb.save(out)
    print(f"  Saved: {out} ({len(ny)} rows)")



