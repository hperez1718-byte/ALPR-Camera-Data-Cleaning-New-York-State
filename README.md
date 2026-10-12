[README.md](https://github.com/user-attachments/files/33318201/README.md)
# ALPR Camera Data Cleaning — New York State

## Overview
Extracted and cleaned a dataset of automatic license plate reader (ALPR) / surveillance camera locations across **New York State** using Python and Excel Power Query. This cleaned dataset feeds a Tableau visualization project (in progress) exploring camera distribution by manufacturer, operator, and region.

## Tools
- Python (data extraction)
- Microsoft Excel — Power Query (data cleaning)

## The Dataset
**Raw:** `flock-cameras-ny-fresh.xlsx` — 8,689 records pulled October 11, 2026 from the DeFlock CDN (crowdsourced OpenStreetMap data) via `pull_flock_cameras.py`.

**Cleaned:** 7,685 records, 13 columns:

| Column | Type | Description |
|---|---|---|
| ID | Whole Number | Unique record identifier |
| LAT / LON | Decimal Number | GPS coordinates |
| Manufacturer_1/2/3 | Text | Camera manufacturer(s), split on "/" |
| Operator_1/2/3 | Text | Operating agency, split on "/" |
| Direction_1–4 | Whole Number | Compass facing direction(s) in degrees, split on ";" |

## Data Quality Issues Found
- **Geographic scope creep:** the extraction bounding box pulled in ~1,000 rows from Connecticut, Massachusetts, Pennsylvania, New Jersey, Vermont, Rhode Island, and Ontario (Canada)
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
2. **Scoped to New York State** — removed ~1,004 rows with out-of-state operators (CT, MA, PA, NJ, VT, RI) and Ontario records → 7,685 rows
3. **Split Manufacturer** on "/" into `Manufacturer_1/2/3` (some cameras list multiple manufacturers)
4. **Split Operator** on "/" into `Operator_1/2/3`, preserving the full original titles (e.g. "Tarrytown Westchester County Police RTC") since the full wording carries useful information
5. **Split Direction** on ";" into `Direction_1`–`Direction_4` (cameras can face multiple directions, e.g. "360;180")
6. **Standardized blanks** → "Unknown" for transparency in text columns
7. **Removed duplicate rows**
8. **Standardized manufacturers:** "Mav (Iq:350Xr)" → "MAV Systems" · "Ndi" → "NDI Recognition Systems" · "Motorola?" → "Motorola" · "Unv" → "UNV" · "Turing Ai" → "Turing AI" · "Lowe'S" → "Lowe's" · trimmed trailing whitespace
9. **Removed non-manufacturer values** ("Municipal Parking Services", "Oswego City") from the Manufacturer column
10. **Fixed the typo** "Lewsiboro" → "Lewisboro"
11. **Typed the Direction columns** as Whole Number; replaced 'None' strings with true nulls — *not* 0, because 0° is a valid compass direction (due north) and filling blanks with 0 would corrupt directional analysis
12. **Flagged 9 impossible direction values** (outside 0–360°) as source-data typos; left in place and documented

## Results
- 8,689 → **7,685 clean New York State records**
- **31 distinct manufacturers** — Flock Safety is the largest single manufacturer
- Zero duplicate IDs; all columns correctly typed

## Key Takeaway
Real-world data is messy. Methodical cleaning — scoping geography, handling nulls, standardizing text, fixing data types — is what makes analysis possible. Every decision here is documented so the process is repeatable.

## What's Next
- Tableau dashboard: camera density by region, manufacturer share, operator coverage
- The narrative is emerging from the clean data — transparency and surveillance density are the leading angles
