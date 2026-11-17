"""Step 2 of v2: build ALPR_Camera_Data_v2.xlsx from the county-tagged data.

Usage: python build_workbook.py with_county.csv ALPR_Camera_Data_v2.xlsx ny_counties.csv
Then open in Excel (or run LibreOffice recalculation) so formula values are cached.
"""
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

SRC, OUT = sys.argv[1], sys.argv[2]
d = pd.read_csv(SRC)

# ---------------- operator lookup (raw -> clean, sector, type) ----------------
G, R, N = "Government/Agency", "Retail/Private", "Not Documented"
LOOK = [  # raw, clean, sector, type, note
 ("Unknown", "Not documented", "Not documented", N, "No operator tag in OpenStreetMap. Means 'not recorded by mappers', not 'secret'."),
 ("NYC Department of Transportation", "NYC Department of Transportation", "Transportation", G, ""),
 ("NYC Dot Red Light Camera", "NYC Department of Transportation", "Transportation", G, "Red-light camera program; merged into NYC DOT."),
 ("Verra Mobility Operated Nyc Dot Red Light Camera", "NYC Department of Transportation", "Transportation", G, "Run by vendor Verra Mobility for NYC DOT; merged into NYC DOT."),
 ("Verra Mobility Nyc Dot Red Light Camera", "NYC Department of Transportation", "Transportation", G, "Run by vendor Verra Mobility for NYC DOT; merged into NYC DOT."),
 ("Verra Mobility Nyc  Dot Red Light Camera", "NYC Department of Transportation", "Transportation", G, "Double space in source; merged into NYC DOT."),
 ("Verra Mobility Operated Nyc Dot Speed Camera", "NYC Department of Transportation", "Transportation", G, "Speed camera program; merged into NYC DOT."),
 ("Verra Mobility Nyc Dot Speed Camera", "NYC Department of Transportation", "Transportation", G, "Speed camera program; merged into NYC DOT."),
 ("New York City Police Department", "New York City Police Department", "Police / Sheriff", G, ""),
 ("Metropolitan Transportation Authority", "Metropolitan Transportation Authority (MTA)", "Transportation", G, ""),
 ("MTA Bridges And Tunnels", "Metropolitan Transportation Authority (MTA)", "Transportation", G, "MTA division; merged into MTA (matches Tableau's 127)."),
 ("Nyc M.T.A", "Metropolitan Transportation Authority (MTA)", "Transportation", G, "Variant spelling; merged into MTA."),
 ("New York State Thruway", "New York State Thruway Authority", "Transportation", G, ""),
 ("Niagara Frontier Transportation Authority", "Niagara Frontier Transportation Authority", "Transportation", G, ""),
 ("New York City Department Of Environmental Protection", "NYC Department of Environmental Protection", "Other Government", G, ""),
 ("Buffalo Police Department", "Buffalo Police Department", "Police / Sheriff", G, ""),
 ("Westchester County Police", "Westchester County Police", "Police / Sheriff", G, ""),
 ("Greenburgh Westchester County Police", "Greenburgh", "Municipality", G, "Source pairs the town with Westchester County Police (shared Real Time Crime Center). Kept as the town, like other Westchester towns."),
 ("Tarrytown Westchester County Police", "Tarrytown", "Municipality", G, "Same pattern as Greenburgh."),
 ("Hempstead Police Department", "Hempstead Police Department", "Police / Sheriff", G, ""),
 ("Nassau County Police Department", "Nassau County Police Department", "Police / Sheriff", G, ""),
 ("Suffolk County Police Department", "Suffolk County Police Department", "Police / Sheriff", G, ""),
 ("Syracuse Police Department", "Syracuse Police Department", "Police / Sheriff", G, ""),
 ("Clarkstown Police Department", "Clarkstown Police Department", "Police / Sheriff", G, ""),
 ("Putnam Police Department", "Putnam Police Department", "Police / Sheriff", G, "Earlier cleaning merged 'Putnam County Sheriff' and 'Putnam Valley' into this value; judgment call worth revisiting with the raw file."),
 ("Garden City Police Department", "Garden City Police Department", "Police / Sheriff", G, ""),
 ("New York State Police", "New York State Police", "Police / Sheriff", G, ""),
 ("Onondaga County Sheriff’S Office", "Onondaga County Sheriff's Office", "Police / Sheriff", G, "Fixed curly apostrophe + capital S."),
 ("Broome County Sheriffs Office", "Broome County Sheriff's Office", "Police / Sheriff", G, "Added apostrophe."),
 ("Dutchess County Sheriff's Office", "Dutchess County Sheriff's Office", "Police / Sheriff", G, ""),
 ("Tompkins County Sheriff", "Tompkins County Sheriff's Office", "Police / Sheriff", G, ""),
 ("Village Of Saint Johnsville Police Department", "Village of St. Johnsville Police Department", "Police / Sheriff", G, ""),
 ("Batavia Police Department", "Batavia Police Department", "Police / Sheriff", G, ""),
 ("Village Of Kenmore Police Department", "Village of Kenmore Police Department", "Police / Sheriff", G, ""),
 ("Old Brookville Police Department", "Old Brookville Police Department", "Police / Sheriff", G, ""),
 ("Village Of Wappigners Pd", "Village of Wappingers Falls Police Department", "Police / Sheriff", G, "Fixed typo 'Wappigners'."),
 ("Amherst Police Department", "Amherst Police Department", "Police / Sheriff", G, ""),
 ("Loyd Harbor Police Department", "Lloyd Harbor Police Department", "Police / Sheriff", G, "Fixed typo 'Loyd'."),
 ("Glen Cove Police Department", "Glen Cove Police Department", "Police / Sheriff", G, ""),
 ("Tonawanda Police Department", "Tonawanda Police Department", "Police / Sheriff", G, ""),
 ("Saranac Lake Police Department", "Saranac Lake Police Department", "Police / Sheriff", G, ""),
 ("Niagara Falls Police Department", "Niagara Falls Police Department", "Police / Sheriff", G, ""),
 ("Niskayuna Police Department", "Niskayuna Police Department", "Police / Sheriff", G, ""),
 ("Stonybrook University", "Stony Brook University", "University", G, "Official name is two words; earlier cleaning had reversed this."),
 ("New York State University Police", "SUNY University Police", "University", G, ""),
 ("Onondaga Community Collage Public Safety (Occ Ps)", "Onondaga Community College Public Safety", "University", G, "Fixed typo 'Collage'."),
 ("Binghamton University PD", "Binghamton University Police", "University", G, ""),
 ("Mount Pleasant", "Mount Pleasant", "Municipality", G, "Town name only; department not specified in source."),
 ("Ossining", "Ossining", "Municipality", G, "Town name only."),
 ("Bedford", "Bedford", "Municipality", G, "Town name only."),
 ("Yorktown Heights", "Yorktown", "Municipality", G, "Earlier find-and-replace turned 'Yorktown' into 'Yorktown Heights'; the town is Yorktown."),
 ("North Castle", "North Castle", "Municipality", G, ""),
 ("White Plains", "White Plains", "Municipality", G, ""),
 ("Briarcliff Manor", "Briarcliff Manor", "Municipality", G, ""),
 ("Town Of Huntington", "Town of Huntington", "Municipality", G, ""),
 ("Carmel", "Carmel", "Municipality", G, ""),
 ("Lewsiboro", "Lewisboro", "Municipality", G, "Fixed typo (README said fixed; it was still in the data)."),
 ("Peekskill", "Peekskill", "Municipality", G, ""),
 ("Pleasantville", "Pleasantville", "Municipality", G, ""),
 ("Pleasantiville", "Pleasantville", "Municipality", G, "Typo re-introduced by Power Query step order; fixed."),
 ("Kent", "Kent", "Municipality", G, ""),
 ("Ardsley", "Ardsley", "Municipality", G, ""),
 ("Croton-On-Hudson", "Croton-on-Hudson", "Municipality", G, ""),
 ("Town Of Amherst", "Town of Amherst", "Municipality", G, ""),
 ("Town Of Newburgh", "Town of Newburgh", "Municipality", G, ""),
 ("Brookhaventown", "Town of Brookhaven", "Municipality", G, "Fixed run-together name."),
 ("Elmsford", "Elmsford", "Municipality", G, ""),
 ("The Home Depot", "The Home Depot", "Retail", R, ""),
 ("Lowe's", "Lowe's", "Retail", R, ""),
 ("Simon Property Group", "Simon Property Group", "Retail", R, "Mall owner."),
 ("Destiny Usa", "Destiny USA", "Retail", R, "Mall."),
 ("Wegmans Food Markets", "Wegmans Food Markets", "Retail", R, ""),
 ("Walden Galeria", "Walden Galleria", "Retail", R, "Mall; fixed spelling 'Galeria'."),
 ("Mcdonald'S", "McDonald's", "Retail", R, "Fixed casing; Tableau's Operator Type calc used this exact misspelling."),
]
look = pd.DataFrame(LOOK, columns=["Operator_Raw", "Operator", "Operator_Sector", "Operator_Type", "Note"])
missing = set(d.Operator) - set(look.Operator_Raw)
assert not missing, missing

# ---------------- row-level fixes & flags ----------------
dirs = ["Direction_1", "Direction_2", "Direction_3", "Direction_4"]
flags = {i: [] for i in d.index}
oor = []
for c in dirs:
    bad = d[c].notna() & ((d[c] < 0) | (d[c] > 360))
    for i in d.index[bad]:
        oor.append((int(d.at[i, "ID"]), c, d.at[i, c]))
        flags[i].append(f"{c} was {int(d.at[i, c])} (outside 0-360); blanked")
    d.loc[bad, c] = pd.NA
dup = d.duplicated(["LAT", "LONG"], keep=False)
for i in d.index[dup]: flags[i].append("Same coordinates as another camera (kept: may be 2 cameras on one pole)")
for i in d.index[d.near_only]: flags[i].append("Just outside county outline (coast); assigned nearest county")
d["Data_Flag"] = [("; ".join(v)) for v in flags.values()]

counties = sorted(pd.read_csv(sys.argv[3]).County) if len(sys.argv) > 3 else sorted(d.County.unique())

# ---------------- workbook ----------------
F = "Arial"
H_FILL = PatternFill("solid", fgColor="1F3864"); H_FONT = Font(name=F, bold=True, color="FFFFFF")
BODY = Font(name=F, size=10); BOLD = Font(name=F, bold=True)
INPUT = PatternFill("solid", fgColor="FFFF00"); BLUE = Font(name=F, size=10, color="0000FF")
thin = Side(style="thin", color="BFBFBF")

def header(ws, row, cols):
    for j, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=j, value=c); cell.fill = H_FILL; cell.font = H_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)

def widths(ws, w):
    for j, x in enumerate(w, 1): ws.column_dimensions[get_column_letter(j)].width = x

def add_table(ws, name, nrows, ncols, start_row=1):
    ref = f"A{start_row}:{get_column_letter(ncols)}{start_row + nrows}"
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)

wb = Workbook()
# ---- About
ab = wb.active; ab.title = "About"
about = [
 ("ALPR Camera Locations - New York State (cleaned, v2)", BOLD),
 ("", None),
 ("Source", BOLD), ("DeFlock (deflock.me), built on OpenStreetMap - crowdsourced, volunteer-mapped camera locations. Pulled Oct 11, 2026 with pull_flock_cameras.py.", None),
 ("License: © OpenStreetMap contributors. Data available under the Open Database License (ODbL). This derived dataset is shared under the ODbL as well.", None),
 ("County boundaries: U.S. Census Bureau cartographic boundary files (1:10m), via the us-atlas package.", None),
 ("", None),
 ("Row history", BOLD),
 ("8,689 raw rows (rectangle around NY) -> 7,685 after removing out-of-state operators -> 4,483 after point-in-polygon check against NY State.", None),
 ("", None),
 ("How to read this data", BOLD),
 ("'Not documented' operator = the OpenStreetMap record has no operator tag. It means volunteers did not record who runs the camera; it does NOT mean the operator is secret.", None),
 ("Camera counts reflect where volunteers have mapped, not a complete inventory. Treat them as a lower bound.", None),
 ("'ALPR' here includes traffic-enforcement cameras (e.g. NYC DOT red-light and speed cameras) as well as investigative plate readers (e.g. Flock Safety).", None),
 ("", None),
 ("Sheets", BOLD),
 ("ALPR_Cameras - one row per camera (main table for Tableau). Operator, Sector, Type and Region are formulas pulled from the lookup sheets.", None),
 ("Operator_Lookup - raw operator text -> clean name, sector, type. Edit here to reclassify; ALPR_Cameras updates automatically.", None),
 ("County_Lookup - 62 NY counties -> Region and NY_Region (10 official REDC regions). Yellow Population column is for you to fill in from the Census (see Change_Log).", None),
 ("County_Summary / Region_Summary - counts and shares, all formulas.", None),
 ("Change_Log - every fix made in v2 and why. Data_Issues - rows that were changed or flagged.", None),
]
for i, (t, f) in enumerate(about, 1):
    c = ab.cell(row=i, column=1, value=t); c.font = f or BODY; c.alignment = Alignment(wrap_text=True, vertical="top")
ab["A1"].font = Font(name=F, bold=True, size=14)
ab.column_dimensions["A"].width = 120

# ---- Operator_Lookup
ol = wb.create_sheet("Operator_Lookup")
cols = list(look.columns); header(ol, 1, cols)
for i, r in enumerate(look.itertuples(index=False), 2):
    for j, v in enumerate(r, 1): ol.cell(row=i, column=j, value=v).font = BODY
add_table(ol, "OperatorLookup", len(look), len(cols)); widths(ol, [48, 44, 18, 20, 80]); ol.freeze_panes = "A2"
OL_N = len(look) + 1

# ---- County_Lookup
cl = wb.create_sheet("County_Lookup")
REDC = {"Capital Region": ["Albany","Columbia","Greene","Rensselaer","Saratoga","Schenectady","Warren","Washington"],
 "Central New York": ["Cayuga","Cortland","Madison","Onondaga","Oswego"],
 "Finger Lakes": ["Genesee","Livingston","Monroe","Ontario","Orleans","Seneca","Wayne","Wyoming","Yates"],
 "Long Island": ["Nassau","Suffolk"],
 "Mid-Hudson": ["Dutchess","Orange","Putnam","Rockland","Sullivan","Ulster","Westchester"],
 "Mohawk Valley": ["Fulton","Herkimer","Montgomery","Oneida","Otsego","Schoharie"],
 "New York City": ["Bronx","Kings","New York","Queens","Richmond"],
 "North Country": ["Clinton","Essex","Franklin","Hamilton","Jefferson","Lewis","St. Lawrence"],
 "Southern Tier": ["Broome","Chemung","Chenango","Delaware","Schuyler","Steuben","Tioga","Tompkins"],
 "Western New York": ["Allegany","Cattaraugus","Chautauqua","Erie","Niagara"]}
C2R = {c: r for r, cs in REDC.items() for c in cs}
assert set(C2R) == set(counties), set(counties) ^ set(C2R)
NYC = {"New York": "Manhattan", "Kings": "Brooklyn", "Queens": "Queens", "Bronx": "Bronx", "Richmond": "Staten Island"}
header(cl, 1, ["County", "Borough", "Region", "Population", "NY_Region"])
for i, c in enumerate(counties, 2):
    reg = "NYC" if c in NYC else "Long Island" if c in ("Nassau", "Suffolk") else "Rest of State"
    for j, v in enumerate([c, NYC.get(c, ""), reg, None], 1):
        cell = cl.cell(row=i, column=j, value=v); cell.font = BODY
    cl.cell(row=i, column=5, value=C2R[c]).font = BODY
    cl.cell(row=i, column=4).fill = INPUT; cl.cell(row=i, column=4).font = BLUE
    cl.cell(row=i, column=4).number_format = "#,##0"
CL_N = len(counties) + 1
add_table(cl, "CountyLookup", len(counties), 5); widths(cl, [18, 16, 16, 14, 20]); cl.freeze_panes = "A2"
cl["D1"].comment = Comment("Fill in: U.S. Census Bureau, ACS 5-year table B01003 (Total Population), all counties in New York State, from data.census.gov.", "Claude")

# ---- ALPR_Cameras
ac = wb.create_sheet("ALPR_Cameras", 1)
cols = ["ID", "LAT", "LONG", "Manufacturer", "Operator_Raw", "Operator", "Operator_Sector", "Operator_Type",
        "Operator_Documented", "County", "Region", "NY_Region", "Direction_1", "Direction_2", "Direction_3", "Direction_4",
        "Has_Direction", "Data_Flag"]
header(ac, 1, cols)
lk = lambda col, r: f"=INDEX(Operator_Lookup!${col}$2:${col}${OL_N},MATCH($E{r},Operator_Lookup!$A$2:$A${OL_N},0))"
for k, r in enumerate(d.itertuples(index=False), 2):
    vals = [int(r.ID), float(r.LAT), float(r.LONG), r.Manufacturer_1, r.Operator,
            lk("B", k), lk("C", k), lk("D", k), f'=IF($E{k}="Unknown","No","Yes")', r.County,
            f"=INDEX(County_Lookup!$C$2:$C${CL_N},MATCH($J{k},County_Lookup!$A$2:$A${CL_N},0))",
            f"=INDEX(County_Lookup!$E$2:$E${CL_N},MATCH($J{k},County_Lookup!$A$2:$A${CL_N},0))"]
    vals += [None if pd.isna(getattr(r, c)) else int(getattr(r, c)) for c in dirs]
    vals += [f'=IF(COUNT(M{k}:P{k})>0,"Yes","No")', r.Data_Flag or None]
    for j, v in enumerate(vals, 1):
        cell = ac.cell(row=k, column=j, value=v); cell.font = BODY
    ac.cell(row=k, column=1).number_format = "0"
    ac.cell(row=k, column=2).number_format = "0.000000"; ac.cell(row=k, column=3).number_format = "0.000000"
N = len(d) + 1
add_table(ac, "ALPRCameras", len(d), len(cols)); ac.freeze_panes = "B2"
widths(ac, [14, 11, 12, 22, 34, 34, 18, 18, 12, 14, 14, 18, 9, 9, 9, 9, 10, 60])

# ---- County_Summary
cs = wb.create_sheet("County_Summary")
cols = ["County", "Region", "Cameras", "Operator Documented", "% Documented", "Flock Safety", "% Flock", "Population", "Cameras per 100k"]
header(cs, 1, cols)
A = "ALPR_Cameras"
for i, c in enumerate(counties, 2):
    row = [f"=County_Lookup!A{i}", f"=County_Lookup!C{i}",
           f"=COUNTIFS({A}!$J$2:$J${N},$A{i})",
           f'=COUNTIFS({A}!$J$2:$J${N},$A{i},{A}!$I$2:$I${N},"Yes")',
           f'=IF(C{i}=0,"",D{i}/C{i})',
           f'=COUNTIFS({A}!$J$2:$J${N},$A{i},{A}!$D$2:$D${N},"Flock Safety")',
           f'=IF(C{i}=0,"",F{i}/C{i})',
           f'=IF(County_Lookup!D{i}="","",County_Lookup!D{i})',
           f'=IF(H{i}="","",C{i}/H{i}*100000)']
    for j, v in enumerate(row, 1):
        cell = cs.cell(row=i, column=j, value=v); cell.font = BODY
        if j in (1, 2, 8): cell.font = Font(name=F, size=10, color="008000")
    cs.cell(row=i, column=5).number_format = "0.0%"; cs.cell(row=i, column=7).number_format = "0.0%"
    cs.cell(row=i, column=8).number_format = "#,##0"; cs.cell(row=i, column=9).number_format = "0.0"
t = CL_N + 1
cs.cell(row=t, column=1, value="Total").font = BOLD
for j, col in [(3, "C"), (4, "D"), (6, "F")]:
    cs.cell(row=t, column=j, value=f"=SUM({col}2:{col}{CL_N})").font = BOLD
cs.cell(row=t, column=5, value=f"=D{t}/C{t}").number_format = "0.0%"
cs.cell(row=t, column=7, value=f"=F{t}/C{t}").number_format = "0.0%"
cs.cell(row=t + 2, column=1, value="Population and Cameras per 100k fill in automatically once County_Lookup!D (yellow) is filled.").font = Font(name=F, italic=True, size=9)
widths(cs, [16, 14, 10, 12, 12, 11, 10, 13, 13]); cs.freeze_panes = "B2"

# ---- Region_Summary
rs = wb.create_sheet("Region_Summary")
header(rs, 1, ["Region", "Cameras", "% of State", "% Operator Documented", "% Flock Safety", "Population", "Cameras per 100k"])
for i, reg in enumerate(["NYC", "Long Island", "Rest of State"], 2):
    vals = [reg, f"=COUNTIFS({A}!$K$2:$K${N},A{i})", f"=B{i}/$B$5",
            f'=COUNTIFS({A}!$K$2:$K${N},A{i},{A}!$I$2:$I${N},"Yes")/B{i}',
            f'=COUNTIFS({A}!$K$2:$K${N},A{i},{A}!$D$2:$D${N},"Flock Safety")/B{i}',
            f'=IF(SUMIFS(County_Lookup!$D$2:$D${CL_N},County_Lookup!$C$2:$C${CL_N},A{i})=0,"",SUMIFS(County_Lookup!$D$2:$D${CL_N},County_Lookup!$C$2:$C${CL_N},A{i}))',
            f'=IF(F{i}="","",B{i}/F{i}*100000)']
    for j, v in enumerate(vals, 1): rs.cell(row=i, column=j, value=v).font = BODY
    for j in (3, 4, 5): rs.cell(row=i, column=j).number_format = "0.0%"
    rs.cell(row=i, column=6).number_format = "#,##0"; rs.cell(row=i, column=7).number_format = "0.0"
rs["A5"] = "Total"; rs["A5"].font = BOLD
rs["B5"] = "=SUM(B2:B4)"; rs["B5"].font = BOLD
rs["C5"] = "=SUM(C2:C4)"; rs["C5"].number_format = "0.0%"
rs["D5"] = f'=COUNTIFS({A}!$I$2:$I${N},"Yes")/B5'; rs["D5"].number_format = "0.0%"
rs["E5"] = f'=COUNTIFS({A}!$D$2:$D${N},"Flock Safety")/B5'; rs["E5"].number_format = "0.0%"
rs["A8"] = "By NY State economic development region (REDC)"; rs["A8"].font = BOLD
header(rs, 9, ["NY_Region", "Cameras", "% of State", "% Operator Documented", "% Flock Safety", "Population", "Cameras per 100k"])
for i, reg in enumerate(sorted(REDC), 10):
    vals = [reg, f"=COUNTIFS({A}!$L$2:$L${N},A{i})", f"=B{i}/$B$5",
            f'=IF(B{i}=0,"",COUNTIFS({A}!$L$2:$L${N},A{i},{A}!$I$2:$I${N},"Yes")/B{i})',
            f'=IF(B{i}=0,"",COUNTIFS({A}!$L$2:$L${N},A{i},{A}!$D$2:$D${N},"Flock Safety")/B{i})',
            f'=IF(SUMIFS(County_Lookup!$D$2:$D${CL_N},County_Lookup!$E$2:$E${CL_N},A{i})=0,"",SUMIFS(County_Lookup!$D$2:$D${CL_N},County_Lookup!$E$2:$E${CL_N},A{i}))',
            f'=IF(F{i}="","",B{i}/F{i}*100000)']
    for j, v in enumerate(vals, 1): rs.cell(row=i, column=j, value=v).font = BODY
    for j in (3, 4, 5): rs.cell(row=i, column=j).number_format = "0.0%"
    rs.cell(row=i, column=6).number_format = "#,##0"; rs.cell(row=i, column=7).number_format = "0.0"
rs["A20"] = "Total"; rs["A20"].font = BOLD; rs["B20"] = "=SUM(B10:B19)"; rs["B20"].font = BOLD
rs["A22"] = "REDC regions: the 10 Regional Economic Development Council regions defined by New York State (Empire State Development)."; rs["A22"].font = Font(name=F, italic=True, size=9)
widths(rs, [20, 10, 11, 14, 12, 13, 13])

# ---- Change_Log
lg = wb.create_sheet("Change_Log")
LOG = [
 ("Operators", "Built Operator_Lookup (raw -> clean name, sector, type) instead of one-off find-and-replace steps", "Find-and-replace on partial text caused errors like 'Yorktown Heights Heights'. A lookup table is visible, editable and repeatable.", f"{len(look)} raw values -> {look.Operator.nunique()} clean operators"),
 ("Operators", "Merged MTA variants (Metropolitan Transportation Authority, MTA Bridges And Tunnels, Nyc M.T.A)", "Same agency. This was done in Tableau but not in the data.", "127 rows"),
 ("Operators", "Merged NYC DOT red-light/speed camera variants (incl. 'Verra Mobility operated') into NYC Department of Transportation", "Verra Mobility runs these cameras on NYC DOT's behalf; DOT is the accountable operator.", f"{d.Operator.str.contains('Verra|Red Light', case=False).sum()} rows"),
 ("Operators", "Fixed typos: Lewsiboro, Pleasantiville, Stonybrook, Collage, Wappigners, Loyd, Brookhaventown, Galeria, Sheriff'S, Mcdonald'S, Destiny Usa", "Lewsiboro was listed as fixed in the README but remained; Pleasantiville was re-introduced by Power Query step order.", "see Operator_Lookup Note column"),
 ("Operators", "Added Operator_Sector (Police / Sheriff, Transportation, Municipality, University, Other Government, Retail, Not documented) and Operator_Documented (Yes/No)", "Tableau's Operator Type was a hard-coded list of 7 retailers with everything else defaulting to Government; this makes the rule explicit.", "all rows"),
 ("Operators", "Renamed 'Unknown' operator to 'Not documented' in the clean Operator column (raw text kept in Operator_Raw)", "'Unknown' / 'Unidentified' implied secrecy. The data only shows that mappers did not record an operator.", "2,789 rows"),
 ("Geography", "Added County via point-in-polygon against Census county boundaries; Region now comes from County", "Tableau's Region used lat/long rectangles that put 129 Westchester/Nassau cameras in 'NYC' and 28 Westchester cameras in 'Long Island'. It also used the raw Long field, which is broken in Tableau's copy of the data.", f"4,483 rows; {int(d.near_only.sum())} coastal points assigned to nearest county"),
 ("Geography", "Added NY_Region: New York State's 10 official economic development (REDC) regions, via County_Lookup", "Splits the large 'Rest of State' group (55 counties) using an official, defensible grouping instead of a made-up one.", "all rows"),
 ("Geography", "Region values renamed: NYC / Long Island / Rest of State (was 'Upstate')", "Westchester, Rockland and the Hudson Valley are not usually called 'Upstate'.", "all rows"),
 ("Direction", f"Blanked {len(oor)} direction values outside 0-360", "Impossible compass values (source typos). README said these were fixed, but they were still in the CSV.", "; ".join(f"ID {i} {c}={int(v)}" for i, c, v in oor)),
 ("Direction", "NOT FIXABLE YET: Power Query steps 'Replaced value 47/48' turned every 0 (due north) into blank, and 'Removed errors' deleted whole rows whose Direction_1 could not convert to a whole number", "This contradicts the README (0 is valid north). Restoring it needs the raw file 'NEW- Flok Cam.xlsx'. In Power Query: delete those 2 steps, and replace 'Removed errors' with 'Replace errors -> null'.", "unknown until raw file is re-run"),
 ("Duplicates", "Flagged (not removed) cameras sharing exact coordinates with another camera", "Could be 2 cameras on one pole (real) or a double-mapped camera. Kept, but flagged so the choice is visible.", f"{int(dup.sum())} rows"),
 ("Tableau", "Data feeding Tableau has broken longitudes (needed the 'Longitude Fixed' calc). This workbook has correct LONG values", "Fix at the source instead of patching downstream. Point Tableau at this workbook and drop 'Longitude Fixed', 'Is NYC Metro' and the old 'Region' calc.", "all rows"),
 ("To do", "Fill County_Lookup!D (Population)", "Needed for cameras per 100k residents. Source: data.census.gov, ACS 5-year table B01003, all counties in New York State.", "62 counties"),
]
header(lg, 1, ["Area", "Change", "Why", "Rows affected"])
for i, r in enumerate(LOG, 2):
    for j, v in enumerate(r, 1):
        c = lg.cell(row=i, column=j, value=v); c.font = BODY; c.alignment = Alignment(wrap_text=True, vertical="top")
widths(lg, [12, 60, 70, 40]); lg.freeze_panes = "A2"

# ---- Data_Issues
di = wb.create_sheet("Data_Issues")
iss = d[d.Data_Flag != ""][["ID", "County", "Manufacturer_1", "Operator", "Data_Flag"]]
header(di, 1, ["ID", "County", "Manufacturer", "Operator_Raw", "Issue"])
for i, r in enumerate(iss.itertuples(index=False), 2):
    for j, v in enumerate(r, 1): di.cell(row=i, column=j, value=v).font = BODY
    di.cell(row=i, column=1).number_format = "0"
add_table(di, "DataIssues", len(iss), 5); widths(di, [14, 14, 20, 30, 80]); di.freeze_panes = "A2"

wb.save(OUT)
print("rows", len(d), "lookup", len(look), "issues", len(iss), "oor", oor)
