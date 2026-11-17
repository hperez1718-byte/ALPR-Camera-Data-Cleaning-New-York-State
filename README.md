# ALPR Camera Data Cleaning — New York State

Cleaning and analysis of automated license plate reader (ALPR) camera locations across **New York State**, built from crowdsourced OpenStreetMap data. This dataset feeds my Tableau Public story: [ALPR Camera Analysis – New York State](https://public.tableau.com/app/profile/hyndrani.perez/viz/ALPRCameraAnalysis-NewYorkState/ALPRStory).

**Tools:** Python (extraction, county assignment, workbook build) · Excel Power Query (first-pass cleaning) · Excel Copilot (point-in-polygon audit) · Tableau Public (visualization)

## Data source and license
- **Source:** [DeFlock](https://deflock.me), which publishes ALPR locations mapped by volunteers in OpenStreetMap. Pulled **October 11, 2026** with `pull_flock_cameras.py`.
- **License:** © OpenStreetMap contributors. Data available under the [Open Database License (ODbL)](https://opendatacommons.org/licenses/odbl/). The cleaned files in this repo are derived from it and shared under the ODbL as well.
- **County boundaries:** U.S. Census Bureau cartographic boundary files (1:10m), via the `us-atlas` package.
- **County population:** U.S. Census Bureau, American Community Survey 2024 5-Year Estimates, table B01003 (Total Population), from data.census.gov.

## What this data can and can't tell you
- **"Operator not documented" ≠ secret.** 62% of cameras have no operator tag. That means volunteers didn't record who runs the camera. Flock Safety cameras are easy to recognize on sight, but their operator usually isn't visible, so 87% of Flock cameras have no documented operator.
- **Counts reflect where volunteers have mapped.** Treat them as a lower bound, not a full inventory.
- **"ALPR" covers different uses.** The data includes traffic-enforcement cameras (e.g., NYC DOT red-light and speed cameras) alongside investigative plate readers (e.g., Flock Safety).

## Files
| File | What it is |
|---|---|
| `ALPR_Camera_Data_v2.xlsx` | **Current version.** Main table plus operator and county lookups, summaries, change log, and flagged rows |
| `ALPR_Camera_Cleaned_NY_v2.csv` | Main table from the workbook as a plain CSV (4,483 rows, 18 columns) |
| `county_pop_2024.csv` | Population of all 62 NY counties (ACS 2024 5-Year, B01003) |
| `pull_flock_cameras.py` | Pulls raw data from the DeFlock CDN |
| `assign_counties.py` | Assigns each camera to its NY county (point-in-polygon) |
| `build_workbook.py` | Builds the v2 workbook (lookups, formulas, summaries, change log) |

The v1 output (`ALPR_Camera_Cleaned_NY.csv`) was removed in favor of v2. It is still available in this repo's commit history for before-and-after comparison.

## Pipeline
**8,689** raw rows (rectangle around NY) → **7,685** after removing out-of-state operators (Power Query) → **4,483** after a point-in-polygon check against the NY State boundary.

### v1: Power Query cleaning (Oct 2026)
- Split multi-value Manufacturer, Operator and Direction fields; standardized blanks, casing and stray characters; removed duplicate rows
- Converted Direction to whole numbers with true nulls instead of 0, because 0° is a valid compass direction (due north)
- **Scope check:** filtering by operator name removed ~1,004 out-of-state rows, but a coordinate check later found **3,196 more** (NJ 1,306 · PA 822 · CT 800 · MA 174 · Ontario 61 · VT 28 · RI 4 · other 1), mostly with blank operators that name filtering couldn't catch. Lesson: verify location with coordinates, not names.

### v2: Audit and restructure (Nov 2026)
A review of v1 found problems in the cleaning itself:
- **Lookup tables instead of find-and-replace.** Replacing partial text caused errors (e.g., "Yorktown" → "Yorktown Heights" → "Yorktown Heights Heights"). All 74 raw operator values now map through `Operator_Lookup` to a clean name, a sector (Police / Sheriff, Transportation, Municipality, University, Retail, Other Government) and a type.
- **Merged variants:** 3 MTA spellings → MTA (127 cameras); NYC DOT red-light and speed programs run by Verra Mobility → NYC DOT (651 cameras).
- **Fixed typos** that v1 listed as fixed but were still in the data (Lewsiboro), or that a later Power Query step brought back (Pleasantiville).
- **Real geography:** each camera is matched to its county by coordinates. `Region` (NYC / Long Island / Rest of State) and `NY_Region` (New York's 10 official economic development regions) come from County. The earlier Tableau region rule used lat/long rectangles, which put 129 Westchester and Nassau cameras in "NYC" and 28 Westchester cameras in "Long Island."
- **Direction:** blanked 2 impossible values still in the v1 CSV (93,420 and 2,060).
- **Flagged, not deleted:** 31 cameras that share exact coordinates with another camera (could be two cameras on one pole), and 11 coastal points assigned to the nearest county.
- **Known open issue:** two Power Query steps turned every 0° (due north) direction into blank, and a "Removed errors" step can delete whole rows. Fixing this means re-running from the raw file; tracked in the workbook's `Change_Log`.

## Results
| Region | Cameras | Cameras per 100k residents | Operator documented | Flock Safety share |
|---|---|---|---|---|
| New York City | 1,320 | 15.6 | 81% | 7% |
| Long Island | 720 | 24.7 | 24% | 76% |
| Rest of State | 2,443 | 28.9 | 19% | 55% |
| **New York State** | **4,483** | **22.6** | **38%** | **44%** |

New York has two very different camera landscapes in this data. In NYC, the mapped cameras are mostly run by city agencies (NYC DOT, NYPD, MTA), use a mix of vendors, and usually have a documented operator. Outside NYC, Flock Safety is the main vendor and most cameras have no documented operator.

**Raw counts vs. per resident:** NYC has the most mapped cameras but the lowest rate per resident. Among New York's 10 official economic development regions, Mid-Hudson has the highest rate (39.6 per 100k) and NYC the lowest (15.6). Counts reflect volunteer mapping coverage, so these rates are a lower bound.

## Next steps
- Compare against EFF's Atlas of Surveillance to measure how many "not documented" operators are actually knowable
- Use NYC Open Data to separate traffic-enforcement cameras from investigative ALPRs
