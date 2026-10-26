# ALPR Camera Data Cleaning — New York State

## Overview
Extracted and cleaned a dataset of automatic license plate reader (ALPR) / surveillance camera locations across **New York State** using Python and Excel Power Query. This cleaned dataset feeds a Tableau visualization project (in progress) exploring camera distribution by manufacturer, operator, and region.

## Tools
- Python (data extraction)
- Microsoft Excel — Power Query (data cleaning)
- Excel Copilot (AI-assisted audit & verification — coordinate point-in-polygon testing, direction validation, operator variant detection)

## The Dataset
**Raw:** `flock-cameras-ny-fresh.xlsx` — 8,689 records pulled October 11, 2026 from the DeFlock CDN (crowdsourced OpenStreetMap data) via `pull_flock_cameras.py`.

**Cleaned:** 4,483 verified New York State records, 10 columns:

| Column | Type | Description |
|---|---|---|
| Manufacturer_1 | Text | Camera manufacturer |
| State | Text | NY |
| Operator | Text | Operating agency |
| ID | Whole Number | Unique record identifier |
| LAT / LONG | Decimal Number | GPS coordinates |
| Direction_1–4 | Whole Number | Compass facing direction(s) in degrees |

## Data Quality Issues Found
- **Geographic scope creep:** the extraction bounding box pulled in thousands of rows from Connecticut, Massachusetts, Pennsylvania, New Jersey, Vermont, Rhode Island, and Ontario (Canada). Operator-name filtering caught ~1,004 — but 3,196 more slipped through, most with "Unknown" operators, invisible to name-based filtering
- Blank cells across Manufacturer, Operator, and Direction columns
- Inconsistent manufacturer capitalization and acronym casing ("Mav (Iq:350Xr)", "Ndi", "Turing Ai", "Unv")
- Stray characters in values ("Motorola?", "Lowe'S", trailing whitespace)
- Non-manufacturer values in the Manufacturer column ("Municipal Parking Services", "Oswego City")
- Operator field combining multiple agencies/programs in a single string
- Typo in a locality name ("Lewsiboro")
- Direction stored as text with 'None' strings instead of nulls
- 9 direction values outside the valid 0–360 compass range (source typos, e.g. -70, 93420)
- Duplicate rows

## Cleaning Steps
1. **Extracted** fresh data from the DeFlock CDN with `pull_flock_cameras.py` (8,689 rows, 6 source columns)
2. **Scoped to New York State (pass 1)** — removed ~1,004 rows with out-of-state operators (CT, MA, PA, NJ, VT, RI) and Ontario records → 7,685 rows
3. **Split Manufacturer** on "/" into `Manufacturer_1/2/3` (some cameras list multiple manufacturers)
4. **Split Operator** on "/" into `Operator_1/2/3`, preserving the full original titles (e.g. "Tarrytown Westchester County Police RTC") since the full wording carries useful information
5. **Split Direction** on ";" into `Direction_1`–`Direction_4` (cameras can face multiple directions, e.g. "360;180")
6. **Standardized blanks** → "Unknown" for transparency in text columns
7. **Removed duplicate rows**
8. **Standardized manufacturers:** "Mav (Iq:350Xr)" → "MAV Systems" · "Ndi" → "NDI Recognition Systems" · "Motorola?" → "Motorola" · "Unv" → "UNV" · "Turing Ai" → "Turing AI" · "Lowe'S" → "Lowe's" · trimmed trailing whitespace
9. **Removed non-manufacturer values** ("Municipal Parking Services", "Oswego City") from the Manufacturer column
10. **Fixed the typo** "Lewsiboro" → "Lewisboro"
11. **Typed the Direction columns** as Whole Number; replaced 'None' strings with true nulls — *not* 0, because 0° is a valid compass direction (due north) and filling blanks with 0 would corrupt directional analysis
12. **Fixed impossible direction values** — 9 cells held values outside the valid 0–360° compass range (source typos, e.g. -70, 93420). 7 belonged to out-of-state rows removed in step 13; the remaining 2 were blanked to true nulls (not 0, since 0° is valid north)
13. **AI-assisted verification & final cleanup (pass 2)** — used Excel Copilot to audit the full dataset: point-in-polygon testing on the LAT/LON coordinates against New York State boundaries, validity checks on Direction_1–4 (must be 0–360 or null), and operator spelling-variant detection. Coordinates don't lie about location: this caught **3,196 out-of-state rows** that operator-name filtering had missed (NJ 1,306 · PA 822 · CT 800 · MA 174 · Ontario 61 · VT 28 · RI 4 · other 1). Flagged rows were moved to a separate audit-trail sheet, then removed → **4,483 verified NY records**. Every flag was independently spot-checked against the coordinates before deletion

## Cleanup pass 2 — October 26, 2026
- **Flattened to single Manufacturer and Operator columns** — Manufacturer_2/3 and Operator_2/3 were nearly empty (one known manufacturer value across both columns; Operator_2 4.0% and Operator_3 2.4% filled), so they were dropped. The data now carries one Manufacturer_1 and one Operator column.
- **Merged operator spelling variants** — "New York City Department Of Transportation" (29 rows) was folded into "NYC Department of Transportation", giving a single NYC DOT value covering 642 rows.
- **Renamed LON → LONG** and dropped Region (every row is NY; the State column is retained).
- Result: same 4,483 verified rows and unique IDs, now in 10 columns.

## Results
- 8,689 → 7,685 → **4,483 verified New York State records** (after AI coordinate verification)
- **25 distinct manufacturers** — Flock Safety is the largest (1,987 cameras, 44.3%)
- **74 distinct operators** — though 62.2% of cameras still have Unknown operators
- Zero duplicate IDs; all columns correctly typed

## Key Takeaway
Real-world data is messy. Methodical cleaning — scoping geography, handling nulls, standardizing text, fixing data types — is what makes analysis possible. And cleaning is never one-and-done: AI-assisted verification (point-in-polygon on raw coordinates) caught 3,196 bad rows that name-based filtering missed. Trust, then verify — especially with coordinates. Every decision here is documented so the process is repeatable.

## What's Next
- Tableau dashboard: camera density by region, manufacturer share, operator coverage
- The narrative is emerging from the clean data — transparency and surveillance density are the leading angles
