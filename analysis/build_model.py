"""Build the Rippling NB staffing model workbook (formula-driven).

Usage: python3 -I analysis/build_model.py data/rippling_nb_dataset.csv model/Rippling_NB_Staffing_Model.xlsx
Conventions: blue font = editable input, black = formula, green = link to another sheet.
"""
import sys
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL

SRC, OUT = sys.argv[1], sys.argv[2]

# ---------------------------------------------------------------- styles
FONT = "Arial"
BLUE, BLACK, GREEN, GREY, WHITE = "0000FF", "000000", "008000", "666666", "FFFFFF"
NAVY = "1F2A44"
F_IN = Font(name=FONT, size=10, color=BLUE)
F_FX = Font(name=FONT, size=10, color=BLACK)
F_LK = Font(name=FONT, size=10, color=GREEN)
F_B = Font(name=FONT, size=10, bold=True)
F_NOTE = Font(name=FONT, size=9, italic=True, color=GREY)
F_TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
F_H = Font(name=FONT, size=10, bold=True, color=WHITE)
F_SEC = Font(name=FONT, size=11, bold=True, color=NAVY)
FILL_H = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor="E8ECF4")
FILL_IN = PatternFill("solid", fgColor="FFF9DB")   # light yellow = key assumption
FILL_OUT = PatternFill("solid", fgColor="E6F4EA")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

USD = '$#,##0;($#,##0);"-"'
USDK = '$#,##0,"K";($#,##0,"K");"-"'
USDM = '$#,##0.00,,"M";($#,##0.00,,"M");"-"'
PCT = '0.0%;(0.0%);"-"'
NUM0 = '#,##0;(#,##0);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
NUM2 = '0.00;(0.00);"-"'
MULT = '0.00"x";(0.00"x");"-"'
MO = '0.0" mo";(0.0" mo");"-"'

wb = Workbook()


def sheet(name, title, subtitle=None, widths=None):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = F_TITLE
    if subtitle:
        ws["A2"] = subtitle
        ws["A2"].font = F_NOTE
    ws.sheet_view.showGridLines = False
    for col, w in (widths or {}).items():
        ws.column_dimensions[col].width = w
    return ws


def put(ws, ref, val, kind="fx", fmt=None, bold=False, fill=None, note=None):
    c = ws[ref]
    c.value = val
    font = {"in": F_IN, "fx": F_FX, "lk": F_LK, "lbl": F_FX, "b": F_B, "note": F_NOTE}[kind]
    if bold:
        font = Font(name=FONT, size=10, bold=True, color=font.color)
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if note:
        c.comment = Comment(note, "Model")
    return c


def header(ws, row, col, labels, width=None):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font, c.fill = F_H, FILL_H
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def section(ws, row, text, span=8):
    c = ws.cell(row=row, column=1, value=text)
    c.font = F_SEC
    for i in range(1, span + 1):
        ws.cell(row=row, column=i).fill = FILL_SEC


# ---------------------------------------------------------------- data
raw = pd.read_csv(SRC)
raw.columns = ["month", "fte", "leads", "demos", "qo", "logos", "arr"]
raw["arr"] = raw["arr"].str.replace(r"[$,]", "", regex=True).astype(float)
raw["month"] = pd.to_datetime(raw["month"])
BUCKETS = ["1-10", "11-50", "51-100", "101-500", "501+"]
MONTHS = sorted(raw["month"].unique())
N = len(raw)  # 60
D0, D1 = 2, N + 1  # data rows

SEGS = ["Velocity (1-50)", "Emerging MM (51-100)", "Core MM (101-500)", "Enterprise (501+)"]
SEGSHORT = ["Velocity", "Emerging MM", "Core MM", "Enterprise"]
SEGCOLS = ["C", "D", "E", "F"]  # per-segment columns on Inputs / Demand

# ================================================================ README
ws = wb.active
ws.title = "README"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 28
ws.column_dimensions["C"].width = 110
ws["B1"] = "Rippling New-Business Revenue Staffing Model"
ws["B1"].font = F_TITLE
ws["B2"] = "Case-study model built on the fictional dataset in the take-home exercise (Feb-2021 to Jan-2022, 5 employee-size buckets)."
ws["B2"].font = F_NOTE
rows = [
    ("PURPOSE", "Segment prospects by employee count and size the SDR / AE / SC capacity each segment needs. Then evaluate cost, expected ARR and scenarios."),
    ("COLOR CODE", "Blue font = editable input / assumption. Black font = formula. Green font = link pulled from another sheet. Yellow fill = key assumption to review."),
    ("ASSUMPTION STATUS", "Every capacity, ramp, quota, compensation and uplift input is ILLUSTRATIVE. None is a Rippling fact. Dataset-derived figures are calculated live from the 'Data' tab."),
    ("FLOW", "Data -> Analysis (diagnostics) -> Segmentation (option comparison) -> Demand (baseline by segment) -> Scenario_Engine (all scenario math) -> Staffing (base recommendation) -> ENT_Economics / Sensitivity -> Checks."),
    ("Data", "60 monthly source rows (blue = source values from the CSV). Helper columns map each bucket to segments and flag the time windows set on the Inputs tab."),
    ("Inputs", "ALL editable assumptions: time windows, capacity, utilization, ramp, quota, comp, scenario levers. Change these to re-run the model."),
    ("Analysis", "Funnel by bucket, monthly trend, growth, mix effect, volume/conversion/deal-size decomposition, L2D-vs-volume correlation."),
    ("Segmentation", "Compares the recommended 4-tier / 3-motion structure with two conventional 3-tier alternatives."),
    ("Demand", "Monthly demand and observed funnel rates per segment for the baseline window (Aug-Oct-21), the stress month (Jan-22), the H1 reference (Feb-Jul-21) and the downside case."),
    ("Scenario_Engine", "One column per segment per scenario. Shows every intermediate step: leads -> demos -> opps -> logos -> ARR, FTE required, cost, efficiency."),
    ("Staffing", "Base-case recommendation. Fractional FTE -> pooled and rounded headcount, ratios, cost, unit economics, and an optional current-headcount gap (enter your own)."),
    ("ENT_Economics", "Enterprise pod unit economics, breakeven win rate, pilot design and payback sensitivity grids."),
    ("Sensitivity", "One-at-a-time flex of each major assumption and its effect on total FTE, cost, ARR and payback."),
    ("Checks", "Reconciliation of the model to the source data plus internal integrity checks. All should read TRUE."),
    ("KEY CAVEATS", "(1) Funnel rates divide same-month counts and are not cohort conversion. (2) No current headcount is in the data, so the model sizes REQUIRED capacity, not net-new hires. (3) The lead-to-demo decline has competing explanations (SDR capacity, lead quality, speed-to-lead, scheduling, AE availability). Its recovery is a scenario lever, not a forecast. (4) The Nov-Dec dip is seen in one year only and is not confirmed seasonality."),
]
r = 4
for k, v in rows:
    ws.cell(row=r, column=2, value=k).font = F_B
    c = ws.cell(row=r, column=3, value=v)
    c.font = F_FX
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30 if len(v) > 110 else 16
    r += 1
ws.cell(row=r + 1, column=2, value="Legend sample").font = F_B
put(ws, f"C{r+2}", "Blue = input (e.g., AE demo capacity 30 / month)", "in")
put(ws, f"C{r+3}", "Black = formula (e.g., AE FTE = demos / (capacity x utilization))", "fx")
put(ws, f"C{r+4}", "Green = cross-sheet link (e.g., Staffing pulls FTE from Scenario_Engine)", "lk")

# ================================================================ Seg_Map
sm = sheet("Seg_Map", "Segment mapping (editable)", "Maps each dataset bucket to a segment. Change the blue labels to test a different structure; Segmentation tab compares options.",
           {"A": 12, "B": 24, "C": 22, "D": 22, "E": 22, "F": 10})
header(sm, 3, 1, ["FTE bucket", "Recommended segment (4-tier)", "Sales motion (3)", "Option A: Conventional 3-tier", "Option B: Alt 3-tier", "Sort"])
segmap = {
    "1-10": (SEGS[0], "Velocity", "SMB (1-100)", "SMB (1-50)"),
    "11-50": (SEGS[0], "Velocity", "SMB (1-100)", "SMB (1-50)"),
    "51-100": (SEGS[1], "Mid-Market", "SMB (1-100)", "MM (51-500)"),
    "101-500": (SEGS[2], "Mid-Market", "MM (101-500)", "MM (51-500)"),
    "501+": (SEGS[3], "Enterprise", "ENT (501+)", "ENT (501+)"),
}
for i, b in enumerate(BUCKETS):
    rr = 4 + i
    put(sm, f"A{rr}", b, "in")
    for j, v in enumerate(segmap[b]):
        put(sm, f"{CL(2+j)}{rr}", v, "in")
    put(sm, f"F{rr}", i + 1, "in")
SM_RNG = "Seg_Map!$A$4:$A$8"

# ================================================================ Inputs
inp = sheet("Inputs", "Inputs & assumptions (all blue cells are editable; ILLUSTRATIVE unless stated)",
            "SF-based comp per case guidance. Validate comp with People/Comp team and capacity with Sales leadership before use.",
            {"A": 2, "B": 46, "C": 16, "D": 16, "E": 16, "F": 16, "G": 70})
IN = {}

def inrow(r, key, label, val, fmt, note, key_assump=False, kind="in"):
    put(inp, f"B{r}", label, "lbl")
    put(inp, f"C{r}", val, kind, fmt, fill=FILL_IN if key_assump else None)
    c = inp[f"G{r}"]
    c.value, c.font = note, F_NOTE
    IN[key] = f"Inputs!$C${r}"

section(inp, 4, "A. Time windows (drive all dataset-derived baselines)", 7)
inrow(5, "win_s", "Baseline window start", datetime(2021, 8, 1), "mmm-yy", "Primary demand baseline = average of Aug-Oct 2021. These are the latest 3 months before the Nov-Dec dip.", True)
inrow(6, "win_e", "Baseline window end", datetime(2021, 10, 1), "mmm-yy", "Excluding Nov-Dec means planning to non-trough demand. See Demand tab note.")
inrow(7, "stress", "Stress-test month", datetime(2022, 1, 1), "mmm-yy", "Jan-22 = highest-demand month observed. Used as the capacity stress test.")
inrow(8, "h1_s", "Reference period start (L2D recovery target)", datetime(2021, 2, 1), "mmm-yy", "Feb-Jul 2021 L2D is the 'pre-erosion' reference.")
inrow(9, "h1_e", "Reference period end", datetime(2021, 7, 1), "mmm-yy", "")
inrow(10, "tr_s", "Trough window start", datetime(2021, 11, 1), "mmm-yy", "Nov-Dec 2021 dip. Seen ONCE, so it is not confirmed seasonality.")
inrow(11, "tr_e", "Trough window end", datetime(2021, 12, 1), "mmm-yy", "")
inrow(12, "dc_s", "Decomposition early window start", datetime(2021, 2, 1), "mmm-yy", "Volume / conversion / deal-size decomposition compares Feb-Apr with the baseline window.")
inrow(13, "dc_e", "Decomposition early window end", datetime(2021, 4, 1), "mmm-yy", "")

section(inp, 15, "B. Global operating assumptions", 7)
inrow(16, "util", "Target utilization of capacity", 0.85, PCT, "Illustrative. 15% buffer for PTO, training, admin, internal meetings and demand spikes. Applied to every role.", True)
inrow(17, "wdays", "Working days per year", 250, NUM0, "Illustrative (52 wks x 5 days less ~10 holidays). Used to express monthly capacity as per-day activity.")
inrow(18, "months", "Months per year", 12, NUM0, "Annualization of monthly run-rates.")
inrow(19, "load", "Fully loaded cost multiplier on OTE", 1.25, NUM2, "Illustrative. Benefits, payroll tax, equipment, T&E. Excludes management overhead and recruiting.", True)
inrow(20, "gm", "Gross margin on ARR (for payback)", 0.80, PCT, "Illustrative SaaS gross margin. Used only for payback months.")
inrow(21, "hurdle", "Payback hurdle (months)", 18, NUM0, "Illustrative investment hurdle for ENT gating.")
inrow(22, "round", "Rounding threshold for headcount", 0.30, NUM2, "A pool's fractional FTE is rounded UP when the fraction is >= this value, otherwise DOWN. The 85% utilization already gives a buffer.")
inrow(23, "rampprod", "Productivity during ramp (avg)", 0.50, PCT, "Illustrative. A new hire averages 50% productivity during the ramp months.")
inrow(24, "sdr_ote", "SDR OTE (SF)", 90000, USD, "Illustrative SF SDR on-target earnings. Validate with Comp.", True)
inrow(25, "sdr_ramp", "SDR ramp (months)", 2, NUM0, "Illustrative.")
inrow(26, "sc_ramp", "SC ramp (months)", 3, NUM0, "Illustrative.")
inrow(27, "down", "Downside demand haircut vs baseline", 0.15, PCT, "Illustrative. Assumed lead pull-back (e.g., lower marketing spend). This is NOT a forecast.", True)
inrow(28, "tr_mo", "Months per year in trough (if trough applied)", 2, NUM0, "Used only in scenarios with the trough toggle on.")

section(inp, 30, "C. Segment assumptions", 7)
header(inp, 31, 2, ["Assumption"] + SEGS + ["Rationale / source"])
SEG_IN = {}

def segrow(r, key, label, vals, fmt, note, key_assump=False):
    put(inp, f"B{r}", label, "lbl")
    for col, v in zip(SEGCOLS, vals):
        put(inp, f"{col}{r}", v, "in", fmt, fill=FILL_IN if key_assump else None)
    inp[f"G{r}"].value, inp[f"G{r}"].font = note, F_NOTE
    SEG_IN[key] = r

segrow(32, "sdr_cap", "SDR capacity: high-intent leads worked / ramped month", [300, 200, 175, 125], NUM0,
       "Illustrative. Covers speed-to-lead outreach, qualification and booking. Larger accounts need more research and multi-threading. 200/mo = ~10 new leads/day.", True)
segrow(33, "ae_cap", "AE capacity: demos held / ramped month", [45, 30, 20, 12], NUM0,
       "Illustrative. Includes follow-up and pipeline management for the demos held. Velocity ~2/day; Enterprise demos are multi-session.", True)
segrow(34, "sc_inv", "SC involvement: share of demos needing an SC", [0.05, 0.15, 0.50, 1.00], PCT,
       "Illustrative. Rises with complexity proxies: D->Q falls from 80% to 65% and demos per logo rise from 3.7 to 7.8.", True)
segrow(35, "sc_cap", "SC capacity: SC-supported demos / ramped month", [30, 30, 30, 20], NUM0,
       "Illustrative. ENT engagements are deeper (discovery, custom demo, security review). Standard ENT coverage = ~1 SC : 1.7 AE.", True)
segrow(36, "ae_ramp", "AE ramp (months)", [3, 3, 4, 6], NUM0, "Illustrative. Longer cycles mean longer ramp.")
segrow(37, "ae_ote", "AE OTE (SF)", [160000, 190000, 220000, 300000], USD, "Illustrative SF AE on-target earnings by segment. Validate with Comp.", True)
segrow(38, "sc_ote", "SC OTE (SF)", [210000, 210000, 210000, 240000], USD, "Illustrative SF Solution Consultant OTE. ENT SCs are more senior.", True)
segrow(39, "quota", "AE annual new-ARR quota (cross-check only)", [400000, 650000, 600000, 900000], USD,
       "Illustrative. Used ONLY for the quota cross-check. It does not drive workload headcount.")
segrow(40, "attain", "Expected quota attainment (cross-check only)", [0.80, 0.80, 0.80, 0.80], PCT, "Illustrative planning attainment.")
segrow(41, "l2d_flag", "L2D-recovery lever applies? (1/0)", [1, 1, 1, 0], NUM0,
       "ENT = 0. Its L2D did not erode as volume tripled, so its bottleneck is win rate.")
segrow(42, "wr_flag", "Win-rate (SC pilot) lever applies? (1/0)", [0, 0, 0, 1], NUM0, "The SC-paired pilot is scoped to Enterprise only.")

section(inp, 44, "D. Enterprise SC-paired pilot", 7)
inrow(45, "pil_cov", "Share of ENT demos covered by SC-paired pilot", 0.30, PCT, "Illustrative. ~2 of ~6 ENT AEs (see pilot AEs below); the rest form the control group.", True)
inrow(46, "pil_sccap", "Pilot SC capacity (demos / month, ~1:1 pairing)", 12, NUM0, "Illustrative. Dedicated SC per AE. This is the cost of the pilot.")
inrow(47, "pil_ae", "Pilot AEs (for pilot sizing)", 2, NUM0, "Illustrative pod size.")
inrow(48, "pil_mo", "Pilot duration (months)", 6, NUM0, "Two quarters, about one ENT sales cycle.")
inrow(49, "ent_gate", "ENT scale gate passed? (1 = pilot met gates)", 0, NUM0, "0 = ENT AE/SC pools round DOWN (fund to the floor of required capacity). Set to 1 once the pilot clears the gates on ENT_Economics.", True)

section(inp, 50, "E. Scenario levers", 7)
SCN = ["Conservative", "Base", "Upside", "Stress test"]
SCN_COLS = ["C", "D", "E", "F"]
header(inp, 51, 2, ["Lever"] + SCN + ["Definition"])
SC_IN = {}

def scrow(r, key, label, vals, fmt, note):
    put(inp, f"B{r}", label, "lbl")
    for col, v in zip(SCN_COLS, vals):
        put(inp, f"{col}{r}", v, "in", fmt, fill=FILL_IN)
    inp[f"G{r}"].value, inp[f"G{r}"].font = note, F_NOTE
    SC_IN[key] = r

scrow(52, "case", "Demand case (1=Downside, 2=Baseline Aug-Oct, 3=Stress Jan-22)", [1, 2, 2, 3], NUM0,
      "Upside holds demand at baseline, so the upside comes from conversion only. Stress tests capacity at Jan-22 demand.")
scrow(53, "rec", "L2D recovery: % of gap to Feb-Jul reference closed", [0.25, 0.50, 1.00, 0.50], PCT,
      "SCENARIO, not forecast. The decline may be lead quality, not capacity. At 0% there is no uplift (see Sensitivity).")
scrow(54, "wr", "ENT win-rate uplift for pilot-covered opps (pts)", [0.00, 0.03, 0.06, 0.03], PCT,
      "SCENARIO. Applied to the pilot-covered share of ENT opps only.")
scrow(55, "trough", "Apply Nov-Dec trough to annual ARR? (1/0)", [1, 0, 0, 0], NUM0,
      "Conservative assumes the trough recurs. The other scenarios annualize the non-trough run-rate.")
for col, v in zip(SCN_COLS, ["Downside demand, partial recovery, no win-rate gain, trough recurs",
                              "Aug-Oct demand, half of L2D gap recovered, +3pt ENT win rate on pilot",
                              "Aug-Oct demand, full L2D recovery, +6pt ENT win rate on pilot",
                              "Jan-22 demand & rates, base levers (capacity stress test)"]):
    c = inp[f"{col}56"]
    c.value, c.font = v, F_NOTE
    c.alignment = Alignment(wrap_text=True, vertical="top")
inp.row_dimensions[56].height = 52
put(inp, "B56", "Scenario description", "lbl")

section(inp, 58, "F. Current headcount (optional; NOT in dataset)", 7)
put(inp, "B59", "Enter current ramped HC in the Staffing tab (Section F) to see the incremental hiring gap.", "note")

# ================================================================ Data
dt = sheet("Data", "Source data (60 monthly rows) + helper columns", "Columns A-G are the source CSV values (blue). Columns H-O are formulas.",
           {"A": 10, "B": 9, "C": 8, "D": 8, "E": 10, "F": 8, "G": 12, "H": 22, "I": 12, "J": 15, "K": 14, "L": 9, "M": 9, "N": 9, "O": 9, "P": 9})
dt.delete_rows(1, 2)
hdr = ["Month", "FTE bucket", "Leads", "Demos", "Qual Opps", "Logos", "ARR ($)", "Segment (4-tier)", "Motion", "Option A seg", "Option B seg",
       "Baseline win (1/0)", "Stress (1/0)", "Reference H1 (1/0)", "Trough (1/0)", "Early win (1/0)"]
header(dt, 1, 1, hdr)
for i, row in raw.iterrows():
    r = i + 2
    put(dt, f"A{r}", row.month.to_pydatetime(), "in", "mmm-yy")
    put(dt, f"B{r}", row.fte, "in")
    for col, key in zip("CDEF", ["leads", "demos", "qo", "logos"]):
        put(dt, f"{col}{r}", int(row[key]), "in", NUM0)
    put(dt, f"G{r}", float(row.arr), "in", USD)
    for col, idx in zip("HIJK", [2, 3, 4, 5]):
        put(dt, f"{col}{r}", f"=INDEX(Seg_Map!${CL(idx)}$4:${CL(idx)}$8,MATCH($B{r},{SM_RNG},0))", "lk")
    put(dt, f"L{r}", f"=IF(AND($A{r}>={IN['win_s']},$A{r}<={IN['win_e']}),1,0)")
    put(dt, f"M{r}", f"=IF($A{r}={IN['stress']},1,0)")
    put(dt, f"N{r}", f"=IF(AND($A{r}>={IN['h1_s']},$A{r}<={IN['h1_e']}),1,0)")
    put(dt, f"O{r}", f"=IF(AND($A{r}>={IN['tr_s']},$A{r}<={IN['tr_e']}),1,0)")
    put(dt, f"P{r}", f"=IF(AND($A{r}>={IN['dc_s']},$A{r}<={IN['dc_e']}),1,0)")
dt.freeze_panes = "C2"
R = lambda col: f"Data!${col}${D0}:${col}${D1}"
DCOL = {"leads": "C", "demos": "D", "qo": "E", "logos": "F", "arr": "G", "seg": "H", "mot": "I", "optA": "J", "optB": "K",
        "win": "L", "str": "M", "h1": "N", "tr": "O", "early": "P", "fte": "B", "month": "A"}

# ================================================================ Analysis
an = sheet("Analysis", "Funnel diagnostics (all formulas on the Data tab)",
           "Rates = aggregated numerator / aggregated denominator. Same-month ratios, NOT cohort conversion.",
           {"A": 30, **{CL(i): 12 for i in range(2, 20)}})
# --- A. Funnel by bucket
section(an, 4, "A. 12-month funnel by employee-size bucket", 17)
cols = ["Bucket", "Leads", "Demos", "Qual Opps", "Logos", "ARR", "L->Demo", "Demo->Opp", "Opp->Win", "Lead->Win", "ARR / logo", "ARR / lead",
        "ARR share", "Lead share", "Logo share", "Leads / logo", "Demos / logo"]
header(an, 5, 1, cols)
for i, b in enumerate(BUCKETS + ["Total"]):
    r = 6 + i
    put(an, f"A{r}", b, "b" if b == "Total" else "lbl")
    for j, key in enumerate(["leads", "demos", "qo", "logos", "arr"]):
        col = CL(2 + j)
        if b == "Total":
            put(an, f"{col}{r}", f"=SUM({col}6:{col}10)", fmt=USD if key == "arr" else NUM0, bold=True)
        else:
            put(an, f"{col}{r}", f"=SUMIFS({R(DCOL[key])},{R('B')},$A{r})", fmt=USD if key == "arr" else NUM0)
    put(an, f"G{r}", f"=C{r}/B{r}", fmt=PCT)
    put(an, f"H{r}", f"=D{r}/C{r}", fmt=PCT)
    put(an, f"I{r}", f"=E{r}/D{r}", fmt=PCT)
    put(an, f"J{r}", f"=E{r}/B{r}", fmt=PCT)
    put(an, f"K{r}", f"=F{r}/E{r}", fmt=USD)
    put(an, f"L{r}", f"=F{r}/B{r}", fmt=USD)
    put(an, f"M{r}", f"=F{r}/F$11", fmt=PCT)
    put(an, f"N{r}", f"=B{r}/B$11", fmt=PCT)
    put(an, f"O{r}", f"=E{r}/E$11", fmt=PCT)
    put(an, f"P{r}", f"=B{r}/E{r}", fmt=NUM1)
    put(an, f"Q{r}", f"=C{r}/E{r}", fmt=NUM1)

# --- B. Monthly totals
section(an, 13, "B. Monthly totals (all buckets)", 17)
header(an, 14, 1, ["Month", "Leads", "Demos", "Qual Opps", "Logos", "ARR", "L->Demo", "Demo->Opp", "Opp->Win", "Lead->Win", "ARR / logo", "ARR / lead", "Trough?"])
for i, m in enumerate(MONTHS):
    r = 15 + i
    put(an, f"A{r}", pd.Timestamp(m).to_pydatetime(), "in", "mmm-yy")
    for j, key in enumerate(["leads", "demos", "qo", "logos", "arr"]):
        put(an, f"{CL(2+j)}{r}", f"=SUMIFS({R(DCOL[key])},{R('A')},$A{r})", fmt=USD if key == "arr" else NUM0)
    put(an, f"G{r}", f"=C{r}/B{r}", fmt=PCT)
    put(an, f"H{r}", f"=D{r}/C{r}", fmt=PCT)
    put(an, f"I{r}", f"=E{r}/D{r}", fmt=PCT)
    put(an, f"J{r}", f"=E{r}/B{r}", fmt=PCT)
    put(an, f"K{r}", f"=F{r}/E{r}", fmt=USD)
    put(an, f"L{r}", f"=F{r}/B{r}", fmt=USD)
    put(an, f"M{r}", f"=SUMIFS({R('O')},{R('A')},$A{r})/5", fmt=NUM0)
MROW0, MROW1 = 15, 26
put(an, "A27", "Total", "b")
for col in "BCDEF":
    put(an, f"{col}27", f"=SUM({col}15:{col}26)", fmt=USD if col == "F" else NUM0, bold=True)

# --- C. Matrices by bucket x month
def matrix(top, title, num, den, fmt, blank_trough=False):
    section(an, top, title, 17)
    header(an, top + 1, 1, ["Month"] + BUCKETS + ["Total"])
    for i in range(12):
        r = top + 2 + i
        put(an, f"A{r}", f"=A{15+i}", fmt="mmm-yy")
        for j, b in enumerate(BUCKETS + ["Total"]):
            col = CL(2 + j)
            bcrit = f",{R('B')},\"{b}\"" if b != "Total" else ""
            n = f"SUMIFS({R(DCOL[num])},{R('A')},$A{r}{bcrit})"
            f = n if den is None else f"{n}/SUMIFS({R(DCOL[den])},{R('A')},$A{r}{bcrit})"
            if blank_trough:
                f = f"IF($M{15+i}=1,\"\",{f})"
            put(an, f"{col}{r}", "=" + f, fmt=fmt)
    return top + 2, top + 13

L2D0, L2D1 = matrix(29, "C1. Lead->Demo by bucket by month", "demos", "leads", PCT)
LD0, LD1 = matrix(44, "C2. Leads by bucket by month", "leads", None, NUM0)
Q2W0, Q2W1 = matrix(59, "C3. Opp->Win by bucket by month", "logos", "qo", PCT)
ASP0, ASP1 = matrix(74, "C4. ARR per logo by bucket by month", "arr", "logos", USD)
L2X0, L2X1 = matrix(89, "C5. Lead->Demo excl. trough months (blank = trough)", "demos", "leads", PCT, True)
LX0, LX1 = matrix(104, "C6. Leads excl. trough months (blank = trough)", "leads", None, NUM0, True)

# --- D. Growth & correlations
section(an, 119, "D. Growth, mix and volume-sensitivity", 17)
header(an, 120, 1, ["Metric"] + BUCKETS + ["Total", "Note"])
first, last = MONTHS[0], MONTHS[-1]
lines = [
    ("Leads: first month (Feb-21)", lambda c: f"={c}{LD0}", NUM0),
    ("Leads: last month (Jan-22)", lambda c: f"={c}{LD1}", NUM0),
    ("Leads growth Feb-21 -> Jan-22", lambda c: f"={c}122/{c}121-1", PCT),
    ("ARR: Feb-21", lambda c, b=None: None, USD),
    ("ARR: Jan-22", None, USD),
    ("ARR growth Feb-21 -> Jan-22", lambda c: f"={c}125/{c}124-1", PCT),
    ("Corr(leads, L2D) excl. trough", lambda c: f"=CORREL({c}{LX0}:{c}{LX1},{c}{L2X0}:{c}{L2X1})", NUM2),
    ("L2D reference (Feb-Jul)", None, PCT),
    ("L2D baseline window (Aug-Oct)", None, PCT),
    ("L2D change (pts)", lambda c: f"={c}129-{c}128", PCT),
]
for i, (lab, fn, fmt) in enumerate(lines):
    r = 121 + i
    put(an, f"A{r}", lab, "lbl")
    for j, b in enumerate(BUCKETS + ["Total"]):
        col = CL(2 + j)
        bcrit = f",{R('B')},\"{b}\"" if b != "Total" else ""
        if lab.startswith("ARR: Feb"):
            f = f"=SUMIFS({R('G')},{R('A')},A15{bcrit})".replace("A15", "$A$15")
        elif lab.startswith("ARR: Jan"):
            f = f"=SUMIFS({R('G')},{R('A')},$A$26{bcrit})"
        elif lab.startswith("L2D reference"):
            f = f"=SUMIFS({R('D')},{R('N')},1{bcrit})/SUMIFS({R('C')},{R('N')},1{bcrit})"
        elif lab.startswith("L2D baseline"):
            f = f"=SUMIFS({R('D')},{R('L')},1{bcrit})/SUMIFS({R('C')},{R('L')},1{bcrit})"
        else:
            f = fn(col)
        put(an, f"{col}{r}", f, fmt=fmt)
put(an, "H127", "Negative = L2D falls in high-volume months, consistent with a capacity ceiling (hypothesis). Not proof of causation.", "note")
put(an, "H130", "Within-segment erosion in buckets <=500. 501+ is flat.", "note")

# --- E. Mix effect
section(an, 132, "E. Mix effect on blended Lead->Demo (Feb-21 vs Jan-22)", 17)
put(an, "A133", "Blended L2D Feb-21", "lbl"); put(an, "B133", f"=C15/B15", fmt=PCT)
put(an, "A134", "Blended L2D Jan-22", "lbl"); put(an, "B134", f"=C26/B26", fmt=PCT)
put(an, "A135", "Jan-22 bucket rates at Feb-21 lead mix", "lbl")
put(an, "B135", f"=SUMPRODUCT(B{LD0}:F{LD0}/G{LD0},B{L2D1}:F{L2D1})", fmt=PCT)
put(an, "A136", "Total change (pts)", "lbl"); put(an, "B136", "=B134-B133", fmt=PCT)
put(an, "A137", "  of which within-bucket rate decline", "lbl"); put(an, "B137", "=B135-B133", fmt=PCT)
put(an, "A138", "  of which mix shift to larger accounts", "lbl"); put(an, "B138", "=B134-B135", fmt=PCT)

# --- F. LMDI decomposition
section(an, 140, "F. ARR change decomposition: early window (Feb-Apr-21) vs baseline window (Aug-Oct-21), LMDI by bucket", 17)
header(an, 141, 1, ["Bucket", "Leads early", "Logos early", "ARR early", "Leads base", "Logos base", "ARR base", "Delta ARR", "Volume effect", "Conversion effect", "Deal-size effect", "Check"])
for i, b in enumerate(BUCKETS + ["Total"]):
    r = 142 + i
    put(an, f"A{r}", b, "b" if b == "Total" else "lbl")
    if b == "Total":
        for col in "BCDEFGHIJKL":
            put(an, f"{col}{r}", f"=SUM({col}142:{col}146)", fmt=USD if col in "DGHIJKL" else NUM0, bold=True)
        continue
    for col, key, flag in [("B", "C", "P"), ("C", "F", "P"), ("D", "G", "P"), ("E", "C", "L"), ("F", "F", "L"), ("G", "G", "L")]:
        put(an, f"{col}{r}", f"=SUMIFS({R(key)},{R('B')},$A{r},{R(flag)},1)", fmt=USD if key == "G" else NUM0)
    put(an, f"H{r}", f"=G{r}-D{r}", fmt=USD)
    k = f"(G{r}-D{r})/LN(G{r}/D{r})"
    put(an, f"I{r}", f"={k}*LN(E{r}/B{r})", fmt=USD)
    put(an, f"J{r}", f"={k}*LN((F{r}/E{r})/(C{r}/B{r}))", fmt=USD)
    put(an, f"K{r}", f"={k}*LN((G{r}/F{r})/(D{r}/C{r}))", fmt=USD)
    put(an, f"L{r}", f"=I{r}+J{r}+K{r}-H{r}", fmt=USD)
put(an, "A148", "Reading: ARR growth is volume-led. Conversion is a drag and deal size is a tailwind (upmarket ASP). Window sums are 3-month totals.", "note")

# --- G. Demo-gap scenario
section(an, 150, "G. Illustrative demo gap: H2 non-trough months at Feb-Jul L2D (scenario, not a forecast)", 17)
header(an, 151, 1, ["Bucket", "H2 leads (non-trough)", "H2 demos actual", "Ref L2D", "Demos at ref L2D", "Demo gap", "Ref Demo->Win", "H2 ARR/logo", "Implied ARR gap"])
for i, b in enumerate(BUCKETS + ["Total"]):
    r = 152 + i
    put(an, f"A{r}", b, "b" if b == "Total" else "lbl")
    if b == "Total":
        for col in "BCEFI":
            put(an, f"{col}{r}", f"=SUM({col}152:{col}156)", fmt=USD if col == "I" else NUM0, bold=True)
        continue
    h2 = f",{R('N')},0,{R('O')},0"
    put(an, f"B{r}", f"=SUMIFS({R('C')},{R('B')},$A{r}{h2})", fmt=NUM0)
    put(an, f"C{r}", f"=SUMIFS({R('D')},{R('B')},$A{r}{h2})", fmt=NUM0)
    put(an, f"D{r}", f"=SUMIFS({R('D')},{R('B')},$A{r},{R('N')},1)/SUMIFS({R('C')},{R('B')},$A{r},{R('N')},1)", fmt=PCT)
    put(an, f"E{r}", f"=B{r}*D{r}", fmt=NUM0)
    put(an, f"F{r}", f"=E{r}-C{r}", fmt=NUM0)
    put(an, f"G{r}", f"=SUMIFS({R('F')},{R('B')},$A{r},{R('N')},1)/SUMIFS({R('D')},{R('B')},$A{r},{R('N')},1)", fmt=PCT)
    put(an, f"H{r}", f"=SUMIFS({R('G')},{R('B')},$A{r}{h2})/SUMIFS({R('F')},{R('B')},$A{r}{h2})", fmt=USD)
    put(an, f"I{r}", f"=F{r}*G{r}*H{r}", fmt=USD)
put(an, "A158", "4 months (Aug, Sep, Oct, Jan). Assumes downstream conversion would hold for the extra demos; 501+ shows no gap.", "note")

# charts
ch = LineChart(); ch.title = "Monthly leads vs demos (all segments)"; ch.height, ch.width = 7.5, 16
ch.add_data(Reference(an, min_col=2, max_col=3, min_row=14, max_row=26), titles_from_data=True)
ch.set_categories(Reference(an, min_col=1, min_row=15, max_row=26)); ch.y_axis.title = "Count"
an.add_chart(ch, "S4")
ch2 = LineChart(); ch2.title = "Lead->Demo by bucket"; ch2.height, ch2.width = 7.5, 16
ch2.add_data(Reference(an, min_col=2, max_col=6, min_row=30, max_row=L2D1), titles_from_data=True)
ch2.set_categories(Reference(an, min_col=1, min_row=L2D0, max_row=L2D1)); ch2.y_axis.numFmt = "0%"
an.add_chart(ch2, "S21")
ch3 = BarChart(); ch3.title = "12-month ARR by bucket"; ch3.height, ch3.width = 7.5, 16
ch3.add_data(Reference(an, min_col=6, min_row=5, max_row=10), titles_from_data=True)
ch3.set_categories(Reference(an, min_col=1, min_row=6, max_row=10)); ch3.legend = None
an.add_chart(ch3, "S38")

# ================================================================ Segmentation
sg = sheet("Segmentation", "Segmentation options compared (formula-driven from Data)",
           "Boundaries can only fall on dataset bucket edges (10 / 50 / 100 / 500). Homogeneity = within-segment spread of bucket ARR/logo and Lead->Win.",
           {"A": 30, **{CL(i): 13 for i in range(2, 18)}})
# bucket helper (ASP & L2W) for spreads
section(sg, 4, "Bucket reference (from Analysis)", 16)
header(sg, 5, 1, ["Bucket", "ARR / logo", "Lead->Win", "Rec segment", "Motion", "Option A", "Option B"])
for i, b in enumerate(BUCKETS):
    r = 6 + i
    put(sg, f"A{r}", f"=Analysis!A{6+i}", "lk")
    put(sg, f"B{r}", f"=Analysis!K{6+i}", "lk", USD)
    put(sg, f"C{r}", f"=Analysis!J{6+i}", "lk", PCT)
    for col, idx in zip("DEFG", "BCDE"):
        put(sg, f"{col}{r}", f"=Seg_Map!{idx}{4+i}", "lk")

OPTS = [("Recommended: 4 segments", "seg", "D", SEGS),
        ("Recommended: 3 motions (shared resources)", "mot", "E", ["Velocity", "Mid-Market", "Enterprise"]),
        ("Option A: Conventional 3-tier", "optA", "F", ["SMB (1-100)", "MM (101-500)", "ENT (501+)"]),
        ("Option B: Alt 3-tier", "optB", "G", ["SMB (1-50)", "MM (51-500)", "ENT (501+)"])]
r = 13
mcols = ["Segment", "Leads", "Logos", "ARR", "Lead share", "ARR share", "L->Demo", "Opp->Win", "Lead->Win", "ARR / logo", "ARR / lead",
         "Demos / logo", "ASP spread (max/min bucket)", "L2W spread (max/min bucket)", "Avg demos / mo"]
SEG_ROWS = {}
for title, key, refcol, names in OPTS:
    section(sg, r, title, 16)
    header(sg, r + 1, 1, mcols)
    rows_here = []
    for nm in names:
        rr = r + 2 + len(rows_here)
        rows_here.append(rr)
        crit = f"{R(DCOL[key])},$A{rr}"
        put(sg, f"A{rr}", nm, "lbl")
        put(sg, f"B{rr}", f"=SUMIFS({R('C')},{crit})", fmt=NUM0)
        put(sg, f"C{rr}", f"=SUMIFS({R('F')},{crit})", fmt=NUM0)
        put(sg, f"D{rr}", f"=SUMIFS({R('G')},{crit})", fmt=USD)
        put(sg, f"E{rr}", f"=B{rr}/SUM({R('C')})", fmt=PCT)
        put(sg, f"F{rr}", f"=D{rr}/SUM({R('G')})", fmt=PCT)
        put(sg, f"G{rr}", f"=SUMIFS({R('D')},{crit})/B{rr}", fmt=PCT)
        put(sg, f"H{rr}", f"=C{rr}/SUMIFS({R('E')},{crit})", fmt=PCT)
        put(sg, f"I{rr}", f"=C{rr}/B{rr}", fmt=PCT)
        put(sg, f"J{rr}", f"=D{rr}/C{rr}", fmt=USD)
        put(sg, f"K{rr}", f"=D{rr}/B{rr}", fmt=USD)
        put(sg, f"L{rr}", f"=SUMIFS({R('D')},{crit})/C{rr}", fmt=NUM1)
        put(sg, f"M{rr}", f"=_xlfn.MAXIFS($B$6:$B$10,${refcol}$6:${refcol}$10,$A{rr})/_xlfn.MINIFS($B$6:$B$10,${refcol}$6:${refcol}$10,$A{rr})", fmt=MULT)
        put(sg, f"N{rr}", f"=_xlfn.MAXIFS($C$6:$C$10,${refcol}$6:${refcol}$10,$A{rr})/_xlfn.MINIFS($C$6:$C$10,${refcol}$6:${refcol}$10,$A{rr})", fmt=MULT)
        put(sg, f"O{rr}", f"=SUMIFS({R('D')},{crit})/12", fmt=NUM1)
    SEG_ROWS[key] = rows_here
    r = r + 3 + len(names)
r += 1
section(sg, r, "Assessment (judgement, informed by the metrics above)", 16)
assess = [
    ("Criterion", "Recommended 4-tier / 3-motion", "Option A: 1-100 / 101-500 / 501+", "Option B: 1-50 / 51-500 / 501+"),
    ("Deal-size homogeneity", "High. Max ASP spread 1.3x (1-50)", "Low. SMB mixes $3K and $9K deals (3.1x)", "Medium. MM mixes $9K and $12K (1.3x), but conversion differs"),
    ("Conversion homogeneity", "High. Each segment sits within one conversion regime", "High within tiers", "Low. MM mixes ~9.6% and ~4.6% Lead->Win (2.1x)"),
    ("Protects best-ROI segment", "Yes. 51-100 gets dedicated AEs and speed-to-lead SLAs", "No. 51-100 competes with $3K deals in one queue", "Partly. 51-100 is pooled with complex 101-500 deals"),
    ("Operational complexity", "4 AE teams, but only 3 SDR pools and 2 SC pools", "Lowest", "Low"),
    ("Main tradeoff", "More routing rules and quota plans to run", "AE time diluted on low-ASP deals", "Different motions forced into one pod"),
]
for i, row in enumerate(assess):
    rr = r + 1 + i
    for j, (v, col) in enumerate(zip(row, ["A", "B", "F", "J"])):
        c = sg[f"{col}{rr}"]
        c.value = v
        c.font = F_H if i == 0 else (F_B if j == 0 else F_FX)
        if i == 0:
            c.fill = FILL_H
    for a, b in [("B", "E"), ("F", "I"), ("J", "M")]:
        sg.merge_cells(f"{a}{rr}:{b}{rr}")
SEG_ASSESS_END = r + len(assess)

# ================================================================ Demand
dm = sheet("Demand", "Demand & observed funnel rates by recommended segment (from Data)",
           "Monthly averages within each window. These feed the Scenario_Engine.",
           {"A": 4, "B": 44, "C": 17, "D": 17, "E": 17, "F": 17, "G": 15, "H": 60})
header(dm, 4, 2, ["Metric"] + SEGS + ["Total", "Note"])
for col, nm in zip(SEGCOLS, SEGS):
    put(dm, f"{col}5", nm, "lk")
put(dm, "B5", "Segment key", "lbl")
DM = {}

def dmrow(r, key, label, ffn, fmt, total="sum", note=""):
    put(dm, f"B{r}", label, "lbl")
    for col in SEGCOLS:
        put(dm, f"{col}{r}", ffn(col), fmt=fmt)
    if total == "sum":
        put(dm, f"G{r}", f"=SUM(C{r}:F{r})", fmt=fmt, bold=True)
    elif total:
        put(dm, f"G{r}", total, fmt=fmt, bold=True)
    if note:
        dm[f"H{r}"].value, dm[f"H{r}"].font = note, F_NOTE
    DM[key] = r

def avg(metric, flag):
    return lambda c: f"=SUMIFS({R(DCOL[metric])},{R('H')},{c}$5,{R(flag)},1)/(COUNTIFS({R(flag)},1)/5)"

section(dm, 7, "A. Baseline window (Aug-Oct-21 by default): monthly average", 8)
put(dm, "B8", "Months in window", "lbl"); put(dm, "C8", f"=COUNTIFS({R('L')},1)/5", fmt=NUM0)
dmrow(9, "b_leads", "Leads / month", avg("leads", "L"), NUM1)
dmrow(10, "b_demos", "Demos / month", avg("demos", "L"), NUM1)
dmrow(11, "b_qo", "Qual opps / month", avg("qo", "L"), NUM1)
dmrow(12, "b_logos", "Logos / month", avg("logos", "L"), NUM1)
dmrow(13, "b_arr", "ARR / month", avg("arr", "L"), USD)
dmrow(14, "b_l2d", "Lead->Demo", lambda c: f"={c}10/{c}9", PCT, "=G10/G9")
dmrow(15, "b_d2q", "Demo->Opp", lambda c: f"={c}11/{c}10", PCT, "=G11/G10")
dmrow(16, "b_q2w", "Opp->Win", lambda c: f"={c}12/{c}11", PCT, "=G12/G11")
dmrow(17, "b_asp", "ARR / logo", lambda c: f"={c}13/{c}12", USD, "=G13/G12")

section(dm, 19, "B. Stress month (Jan-22)", 8)
dmrow(20, "s_leads", "Leads / month", avg("leads", "M"), NUM1)
dmrow(21, "s_demos", "Demos / month", avg("demos", "M"), NUM1)
dmrow(22, "s_qo", "Qual opps / month", avg("qo", "M"), NUM1)
dmrow(23, "s_logos", "Logos / month", avg("logos", "M"), NUM1)
dmrow(24, "s_arr", "ARR / month", avg("arr", "M"), USD)
dmrow(25, "s_l2d", "Lead->Demo", lambda c: f"={c}21/{c}20", PCT, "=G21/G20")
dmrow(26, "s_d2q", "Demo->Opp", lambda c: f"={c}22/{c}21", PCT, "=G22/G21")
dmrow(27, "s_q2w", "Opp->Win", lambda c: f"={c}23/{c}22", PCT, "=G23/G22")
dmrow(28, "s_asp", "ARR / logo", lambda c: f"={c}24/{c}23", USD, "=G24/G23")
dmrow(29, "s_vs_b", "Stress leads vs baseline", lambda c: f"={c}20/{c}9-1", PCT, "=G20/G9-1")

section(dm, 31, "C. Reference period (Feb-Jul-21): L2D recovery target", 8)
dmrow(32, "h_leads", "Leads / month", avg("leads", "N"), NUM1)
dmrow(33, "h_demos", "Demos / month", avg("demos", "N"), NUM1)
dmrow(34, "h_l2d", "Lead->Demo (reference)", lambda c: f"={c}33/{c}32", PCT, "=G33/G32")
dmrow(35, "gap_b", "L2D gap vs baseline (pts, + = erosion)", lambda c: f"={c}34-{c}14", PCT, "=G34-G14")

section(dm, 37, "D. Downside demand & trough", 8)
dmrow(38, "d_leads", "Downside leads / month (baseline x (1 - haircut))", lambda c: f"={c}9*(1-{IN['down']})", NUM1)
dmrow(39, "t_arr", "Trough-window ARR / month (observed)", avg("arr", "O"), USD)
put(dm, "B40", "Trough ARR factor (trough ARR / baseline ARR)", "lbl")
put(dm, "G40", "=G39/G13", fmt=PCT, bold=True)
DM["t_fac"] = 40
put(dm, "B41", "Trough leads factor (trough leads / baseline leads)", "lbl")
put(dm, "G41", f"=(SUMIFS({R('C')},{R('O')},1)/(COUNTIFS({R('O')},1)/5))/G9", fmt=PCT, bold=True)
notes = [
    "WHY EXCLUDE NOV-DEC FROM THE BASELINE: capacity must be sized to non-trough demand. Sizing to a 12-month average would leave the team",
    "short in 10 of 12 months. Leads fell ~41% and ARR ~61% vs the Aug-Oct baseline, but there is only ONE year, so the dip is not confirmed seasonality.",
    "Implications: (1) Annualizing the Aug-Oct run-rate x12 OVERSTATES annual ARR if the trough recurs. The Conservative scenario applies the observed trough to show the size of that effect.",
    "(2) If the trough recurs, the team runs at ~40-60% utilization for 2 months. Plan Q4 for hiring, onboarding, enablement and pipeline building.",
    "(3) If the dip was a one-off (e.g., a data or process issue), the baseline is correct and the Jan-22 rebound supports it.",
]
for i, n in enumerate(notes):
    put(dm, f"B{43+i}", n, "note")

# ================================================================ Scenario_Engine
en = sheet("Scenario_Engine", "Scenario engine: every step from demand to headcount, cost and ARR",
           "One column per segment per scenario. Capacity is sized to a non-trough month. Revenue impact is split into run-rate vs lever-driven increments (no double counting).",
           {"A": 4, "B": 50})
NSEG = 4
blocks = []  # (scenario idx, [seg cols], total col)
col = 3
for s in range(4):
    segc = [CL(col + k) for k in range(NSEG)]
    totc = CL(col + NSEG)
    blocks.append((s, segc, totc))
    for c in segc:
        en.column_dimensions[c].width = 13
    en.column_dimensions[totc].width = 14
    col += NSEG + 1
EN = {}
lines = []  # define rows

def L(key, label, fmt, seg_fn, tot="sum", sec=None):
    lines.append((key, label, fmt, seg_fn, tot, sec))

# helper getters: seg index k, scenario index s
def demand(metric_b, metric_s=None):
    return lambda k, s, c: f"=IF({sc('case', s)}=3,Demand!{SEGCOLS[k]}{DM[metric_s]},Demand!{SEGCOLS[k]}{DM[metric_b]})"

def sc(key, s):
    return f"Inputs!${SCN_COLS[s]}${SC_IN[key]}"

def si(key, k):
    return f"Inputs!${SEGCOLS[k]}${SEG_IN[key]}"

def r_(key):
    return EN[key]

L("hdr_s", "Scenario", None, None, None)
L("hdr_k", "Segment", None, None, None)
L(None, "DEMAND & FUNNEL (per non-trough month)", None, None, None, True)
L("case", "Demand case (1 Downside / 2 Baseline / 3 Stress)", NUM0, lambda k, s, c: f"={sc('case', s)}", None)
L("leads", "Leads / month", NUM1, lambda k, s, c: f"=CHOOSE({c}{r_('case')},Demand!{SEGCOLS[k]}{DM['d_leads']},Demand!{SEGCOLS[k]}{DM['b_leads']},Demand!{SEGCOLS[k]}{DM['s_leads']})")
L("o_l2d", "Observed Lead->Demo (stress uses Jan-22 rates)", PCT, demand("b_l2d", "s_l2d"), "w:o_demos0/leads")
L("o_d2q", "Observed Demo->Opp", PCT, demand("b_d2q", "s_d2q"), None)
L("o_q2w", "Observed Opp->Win", PCT, demand("b_q2w", "s_q2w"), None)
L("asp", "Observed ARR / logo", USD, demand("b_asp", "s_asp"), None)
L("h_l2d", "Reference Lead->Demo (Feb-Jul)", PCT, lambda k, s, c: f"=Demand!{SEGCOLS[k]}{DM['h_l2d']}", None)
L("l2d_flag", "L2D lever applies (1/0)", NUM0, lambda k, s, c: f"={si('l2d_flag', k)}", None)
L("gap", "L2D gap eligible for recovery (pts)", PCT, lambda k, s, c: f"=MAX(0,{c}{r_('h_l2d')}-{c}{r_('o_l2d')})*{c}{r_('l2d_flag')}", None)
L("rec", "Recovery: % of gap closed (scenario)", PCT, lambda k, s, c: f"={sc('rec', s)}", None)
L("t_l2d", "Target Lead->Demo", PCT, lambda k, s, c: f"={c}{r_('o_l2d')}+{c}{r_('gap')}*{c}{r_('rec')}", "w:demos/leads")
L("demos", "Demos / month", NUM1, lambda k, s, c: f"={c}{r_('leads')}*{c}{r_('t_l2d')}")
L("qo", "Qualified opps / month", NUM1, lambda k, s, c: f"={c}{r_('demos')}*{c}{r_('o_d2q')}")
L("wr_flag", "Win-rate lever applies (1/0)", NUM0, lambda k, s, c: f"={si('wr_flag', k)}", None)
L("wr_up", "Effective win-rate uplift (pts x pilot coverage)", PCT, lambda k, s, c: f"={sc('wr', s)}*{IN['pil_cov']}*{c}{r_('wr_flag')}", None)
L("t_q2w", "Target Opp->Win", PCT, lambda k, s, c: f"={c}{r_('o_q2w')}+{c}{r_('wr_up')}", "w:logos/qo")
L("logos", "Logos / month", NUM1, lambda k, s, c: f"={c}{r_('qo')}*{c}{r_('t_q2w')}")
L("arr_m", "New ARR / non-trough month", USD, lambda k, s, c: f"={c}{r_('logos')}*{c}{r_('asp')}")
L(None, "ANNUALIZATION & REVENUE ATTRIBUTION", None, None, None, True)
L("trough", "Trough applied (1/0)", NUM0, lambda k, s, c: f"={sc('trough', s)}", None)
L("ann", "Effective months / year", NUM2, lambda k, s, c: f"={IN['months']}-{c}{r_('trough')}*{IN['tr_mo']}*(1-Demand!$G${DM['t_fac']})", None)
L("arr_y", "Annual new ARR (total)", USD, lambda k, s, c: f"={c}{r_('arr_m')}*{c}{r_('ann')}")
L("arr_rr", "  Run-rate ARR (observed rates, no levers)", USD, lambda k, s, c: f"={c}{r_('leads')}*{c}{r_('o_l2d')}*{c}{r_('o_d2q')}*{c}{r_('o_q2w')}*{c}{r_('asp')}*{c}{r_('ann')}")
L("arr_l2d", "  + Incremental: L2D recovery lever", USD, lambda k, s, c: f"={c}{r_('leads')}*({c}{r_('t_l2d')}-{c}{r_('o_l2d')})*{c}{r_('o_d2q')}*{c}{r_('o_q2w')}*{c}{r_('asp')}*{c}{r_('ann')}")
L("arr_wr", "  + Incremental: ENT win-rate lever (on target demos)", USD, lambda k, s, c: f"={c}{r_('leads')}*{c}{r_('t_l2d')}*{c}{r_('o_d2q')}*({c}{r_('t_q2w')}-{c}{r_('o_q2w')})*{c}{r_('asp')}*{c}{r_('ann')}")
L("arr_chk", "  Attribution check (should be 0)", USD, lambda k, s, c: f"=ROUND({c}{r_('arr_rr')}+{c}{r_('arr_l2d')}+{c}{r_('arr_wr')}-{c}{r_('arr_y')},2)")
L(None, "WORKLOAD CAPACITY (sized to a non-trough month)", None, None, None, True)
L("util", "Target utilization", PCT, lambda k, s, c: f"={IN['util']}", None)
L("sdr_cap", "SDR capacity (leads / ramped month)", NUM0, lambda k, s, c: f"={si('sdr_cap', k)}", None)
L("sdr", "SDR FTE required = leads / (capacity x util)", NUM2, lambda k, s, c: f"={c}{r_('leads')}/({c}{r_('sdr_cap')}*{c}{r_('util')})")
L("ae_cap", "AE capacity (demos / ramped month)", NUM0, lambda k, s, c: f"={si('ae_cap', k)}", None)
L("ae", "AE FTE required = demos / (capacity x util)", NUM2, lambda k, s, c: f"={c}{r_('demos')}/({c}{r_('ae_cap')}*{c}{r_('util')})")
L("sc_inv", "SC involvement (% of demos)", PCT, lambda k, s, c: f"={si('sc_inv', k)}", None)
L("sc_cap", "SC capacity (supported demos / month)", NUM0, lambda k, s, c: f"={si('sc_cap', k)}", None)
L("pil", "Pilot coverage (ENT only)", PCT, lambda k, s, c: f"={IN['pil_cov']}*{c}{r_('wr_flag')}", None)
L("sc", "SC FTE required = SC demos x mix of std/pilot capacity / util", NUM2,
  lambda k, s, c: f"={c}{r_('demos')}*{c}{r_('sc_inv')}*((1-{c}{r_('pil')})/{c}{r_('sc_cap')}+{c}{r_('pil')}/{IN['pil_sccap']})/{c}{r_('util')}")
L("fte", "Total FTE required", NUM2, lambda k, s, c: f"={c}{r_('sdr')}+{c}{r_('ae')}+{c}{r_('sc')}")
L(None, "QUOTA CROSS-CHECK (separate from workload; does not drive headcount)", None, None, None, True)
L("quota", "AE annual quota", USD, lambda k, s, c: f"={si('quota', k)}", None)
L("attain", "Planning attainment", PCT, lambda k, s, c: f"={si('attain', k)}", None)
L("ae_q", "AE FTE implied by quota = annual ARR / (quota x attainment)", NUM2, lambda k, s, c: f"={c}{r_('arr_y')}/({c}{r_('quota')}*{c}{r_('attain')})")
L("ae_gap", "Workload AE minus quota AE (+ = workload needs more)", NUM2, lambda k, s, c: f"={c}{r_('ae')}-{c}{r_('ae_q')}")
L("att_impl", "Implied attainment if staffed to workload", PCT, lambda k, s, c: f"={c}{r_('arr_y')}/({c}{r_('ae')}*{c}{r_('quota')})", "w:arr_y/ae_quota")
L(None, "COST & EFFICIENCY (fully loaded, steady-state annual)", None, None, None, True)
L("c_sdr1", "SDR loaded cost / FTE", USD, lambda k, s, c: f"={IN['sdr_ote']}*{IN['load']}", None)
L("c_ae1", "AE loaded cost / FTE", USD, lambda k, s, c: f"={si('ae_ote', k)}*{IN['load']}", None)
L("c_sc1", "SC loaded cost / FTE", USD, lambda k, s, c: f"={si('sc_ote', k)}*{IN['load']}", None)
L("c_sdr", "SDR cost", USD, lambda k, s, c: f"={c}{r_('sdr')}*{c}{r_('c_sdr1')}")
L("c_ae", "AE cost", USD, lambda k, s, c: f"={c}{r_('ae')}*{c}{r_('c_ae1')}")
L("c_sc", "SC cost", USD, lambda k, s, c: f"={c}{r_('sc')}*{c}{r_('c_sc1')}")
L("cost", "Total staffing cost", USD, lambda k, s, c: f"={c}{r_('c_sdr')}+{c}{r_('c_ae')}+{c}{r_('c_sc')}")
L("eff", "Annual new ARR / staffing cost", MULT, lambda k, s, c: f"={c}{r_('arr_y')}/{c}{r_('cost')}", "w:arr_y/cost")
L("payback", "Gross-margin payback (months)", MO, lambda k, s, c: f"={c}{r_('cost')}/({c}{r_('arr_y')}*{IN['gm']}/12)", "w:payback")
L("arr_fte", "Annual new ARR per FTE", USD, lambda k, s, c: f"={c}{r_('arr_y')}/{c}{r_('fte')}", "w:arr_y/fte")
L(None, "INCREMENTAL VIEW (levers vs no-lever capacity at same demand)", None, None, None, True)
L("ae0", "AE FTE without levers", NUM2, lambda k, s, c: f"={c}{r_('leads')}*{c}{r_('o_l2d')}/({c}{r_('ae_cap')}*{c}{r_('util')})")
L("sc0", "SC FTE without levers or pilot", NUM2, lambda k, s, c: f"={c}{r_('leads')}*{c}{r_('o_l2d')}*{c}{r_('sc_inv')}/({c}{r_('sc_cap')}*{c}{r_('util')})")
L("d_fte", "Incremental FTE for levers (AE + SC)", NUM2, lambda k, s, c: f"={c}{r_('ae')}-{c}{r_('ae0')}+{c}{r_('sc')}-{c}{r_('sc0')}")
L("d_cost", "Incremental cost for levers", USD, lambda k, s, c: f"=({c}{r_('ae')}-{c}{r_('ae0')})*{c}{r_('c_ae1')}+({c}{r_('sc')}-{c}{r_('sc0')})*{c}{r_('c_sc1')}")
L("d_arr", "Incremental ARR from levers", USD, lambda k, s, c: f"={c}{r_('arr_l2d')}+{c}{r_('arr_wr')}")
L("d_arr_fte", "Incremental ARR per incremental FTE", USD, lambda k, s, c: f"=IFERROR({c}{r_('d_arr')}/{c}{r_('d_fte')},0)", "w:d_arr/d_fte")
L("d_eff", "Incremental ARR / incremental cost", MULT, lambda k, s, c: f"=IFERROR({c}{r_('d_arr')}/{c}{r_('d_cost')},0)", "w:d_arr/d_cost")

# assign row numbers
row = 4
for key, label, fmt, fn, tot, sec in lines:
    if key:
        EN[key] = row
    row += 1
ROW_LAST = row - 1
# write
row = 4
WEIGHTED = {
    "w:o_demos0/leads": lambda c: f"=SUMPRODUCT({{s}}{EN['leads']}:{{e}}{EN['leads']},{{s}}{EN['o_l2d']}:{{e}}{EN['o_l2d']})/{c}{EN['leads']}",
}
for key, label, fmt, fn, tot, sec in lines:
    if sec:
        section(en, row, label, 3 + 5 * 4)
        row += 1
        continue
    put(en, f"B{row}", label, "b" if key in ("hdr_s", "hdr_k") else "lbl")
    for s, segc, totc in blocks:
        if key == "hdr_s":
            for c in segc + [totc]:
                put(en, f"{c}{row}", SCN[s], "b", fill=FILL_SEC)
            continue
        if key == "hdr_k":
            for k, c in enumerate(segc):
                put(en, f"{c}{row}", SEGSHORT[k], "b", fill=FILL_SEC)
            put(en, f"{totc}{row}", "Total", "b", fill=FILL_SEC)
            continue
        for k, c in enumerate(segc):
            v = fn(k, s, c)
            kind = "lk" if ("Inputs!" in v or "Demand!" in v) and not any(op in v.split("!")[-1] for op in ["*", "/"]) and fmt != NUM2 else "fx"
            if key in ("case", "rec", "trough", "util", "sdr_cap", "ae_cap", "sc_inv", "sc_cap", "quota", "attain", "l2d_flag", "wr_flag", "h_l2d", "o_d2q", "o_q2w", "asp", "o_l2d", "leads"):
                kind = "lk"
            put(en, f"{c}{row}", v, kind, fmt)
        s0, s1 = segc[0], segc[-1]
        t = totc
        if tot == "sum":
            put(en, f"{t}{row}", f"=SUM({s0}{row}:{s1}{row})", fmt=fmt, bold=True)
        elif tot and tot.startswith("w:"):
            spec = tot[2:]
            if spec == "o_demos0/leads":
                f = f"=SUMPRODUCT({s0}{EN['leads']}:{s1}{EN['leads']},{s0}{EN['o_l2d']}:{s1}{EN['o_l2d']})/{t}{EN['leads']}"
            elif spec == "demos/leads":
                f = f"={t}{EN['demos']}/{t}{EN['leads']}"
            elif spec == "logos/qo":
                f = f"={t}{EN['logos']}/{t}{EN['qo']}"
            elif spec == "arr_y/ae_quota":
                f = f"={t}{EN['arr_y']}/SUMPRODUCT({s0}{EN['ae']}:{s1}{EN['ae']},{s0}{EN['quota']}:{s1}{EN['quota']})"
            elif spec == "arr_y/cost":
                f = f"={t}{EN['arr_y']}/{t}{EN['cost']}"
            elif spec == "payback":
                f = f"={t}{EN['cost']}/({t}{EN['arr_y']}*{IN['gm']}/12)"
            elif spec == "arr_y/fte":
                f = f"={t}{EN['arr_y']}/{t}{EN['fte']}"
            elif spec == "d_arr/d_fte":
                f = f"=IFERROR({t}{EN['d_arr']}/{t}{EN['d_fte']},0)"
            elif spec == "d_arr/d_cost":
                f = f"=IFERROR({t}{EN['d_arr']}/{t}{EN['d_cost']},0)"
            put(en, f"{t}{row}", f, fmt=fmt, bold=True)
    row += 1
en.freeze_panes = "C6"
for s, segc, totc in blocks:
    for rr in range(4, ROW_LAST + 1):
        if en[f"{totc}{rr}"].fill.fgColor.rgb in (None, "00000000"):
            en[f"{totc}{rr}"].fill = PatternFill("solid", fgColor="F3F5F9")
# Notes under engine
nr = ROW_LAST + 2
for i, n in enumerate([
    "No double counting: run-rate + L2D lever + win-rate lever = total ARR, checked per column. SDR FTE is sized to leads, so it does not change with the L2D lever.",
    "The L2D lever assumes added SDR/AE capacity and process fixes restore part of the Feb-Jul conversion. That causal link is UNPROVEN, so the lever is a scenario.",
    "Volume growth is NOT credited to headcount. Run-rate ARR changes across scenarios only because the demand case changes.",
]):
    put(en, f"B{nr+i}", n, "note")

BASE = blocks[1]  # (s, segc, totc)

def eng(key, s, k=None):
    blk = blocks[s]
    c = blk[2] if k is None else blk[1][k]
    return f"Scenario_Engine!${c}${EN[key]}"

# ================================================================ Staffing
st = sheet("Staffing", "Recommended staffing: Base scenario (Aug-Oct-21 demand)",
           "REQUIRED capacity, not net-new hires (current headcount is not in the dataset). Fractional FTE comes from Scenario_Engine; pools are rounded per the Inputs rule.",
           {"A": 4, "B": 40, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 58})
section(st, 4, "A. Required capacity by segment (fractional FTE, base scenario)", 8)
header(st, 5, 2, ["Role"] + SEGS + ["Total"])
ST = {}
for i, (role, key) in enumerate([("SDR", "sdr"), ("AE", "ae"), ("SC", "sc")]):
    r = 6 + i
    put(st, f"B{r}", role, "lbl")
    for k, col in enumerate(SEGCOLS):
        put(st, f"{col}{r}", f"={eng(key, 1, k)}", "lk", NUM2)
    put(st, f"G{r}", f"=SUM(C{r}:F{r})", fmt=NUM2, bold=True)
    ST[key] = r
put(st, "B9", "Total", "b")
for col in SEGCOLS + ["G"]:
    put(st, f"{col}9", f"=SUM({col}6:{col}8)", fmt=NUM2, bold=True)

section(st, 11, "B. Pooled & rounded headcount (shared resources across the 3 motions)", 8)
header(st, 12, 2, ["Pool", "Segments covered", "Fractional FTE", "Recommended HC", "Motion", "", "Why pooled"])
RND = lambda x: f"=MAX(1,IF({x}-INT({x})>={IN['round']},INT({x})+1,INT({x})))"
pools = [
    ("SDR: Commercial inbound pool", "Velocity + Emerging MM", "=C6+D6", "Velocity / MM", "Same high-intent inbound profile (L2D ~36-40%). One pool smooths volume. Routing still prioritizes 51-100 (highest ARR/lead)."),
    ("SDR: Core MM", "Core MM", "=E6", "Mid-Market", "Higher-volume, lower-L2D segment that needs qualification rigor."),
    ("SDR: Enterprise", "Enterprise", "=F6", "Enterprise", "Account research and multi-threading. Candidate for an outbound/ABM test."),
    ("AE: Velocity", "Velocity", "=C7", "Velocity", "Dedicated. High-velocity, $3-4K ASP. 1-10 gets digital-assist scheduling."),
    ("AE: Emerging MM", "Emerging MM", "=D7", "Mid-Market", "Dedicated. Highest ARR/lead; protect with speed-to-lead SLAs."),
    ("AE: Core MM", "Core MM", "=E7", "Mid-Market", "Dedicated. Complex buying; SC-supported."),
    ("AE: Enterprise", "Enterprise", "=F7", "Enterprise", "Hold to base demand. Expansion gated on pilot (see ENT_Economics)."),
    ("SC: Commercial pool", "Velocity + Emerging + Core MM", "=C8+D8+E8", "Velocity / MM", "Shared pool. Coverage on request below 100 employees, ~50% of Core MM demos."),
    ("SC: Enterprise (incl. pilot)", "Enterprise", "=F8", "Enterprise", "Standard ~1:1.7 coverage plus 1:1 pairing on the pilot half."),
]
POOL0 = 13
for i, (nm, segs_, fx, mot, why) in enumerate(pools):
    r = POOL0 + i
    put(st, f"B{r}", nm, "lbl")
    put(st, f"C{r}", segs_, "lbl")
    put(st, f"D{r}", fx, fmt=NUM2)
    rnd = RND(f"D{r}")
    if "Enterprise" in nm and not nm.startswith("SDR"):
        rnd = f"=IF({IN['ent_gate']}=1,{rnd[1:]},MAX(1,INT(D{r})))"
    put(st, f"E{r}", rnd, fmt=NUM0, bold=True, fill=FILL_OUT)
    put(st, f"F{r}", mot, "lbl")
    st[f"H{r}"].value, st[f"H{r}"].font = why, F_NOTE
POOL1 = POOL0 + len(pools) - 1
r = POOL1 + 1
put(st, f"B{r}", "Total SDR", "b"); put(st, f"D{r}", f"=SUM(D{POOL0}:D{POOL0+2})", fmt=NUM2); put(st, f"E{r}", f"=SUM(E{POOL0}:E{POOL0+2})", fmt=NUM0, bold=True); ST["t_sdr"] = r
put(st, f"B{r+1}", "Total AE", "b"); put(st, f"D{r+1}", f"=SUM(D{POOL0+3}:D{POOL0+6})", fmt=NUM2); put(st, f"E{r+1}", f"=SUM(E{POOL0+3}:E{POOL0+6})", fmt=NUM0, bold=True); ST["t_ae"] = r + 1
put(st, f"B{r+2}", "Total SC", "b"); put(st, f"D{r+2}", f"=SUM(D{POOL0+7}:D{POOL0+8})", fmt=NUM2); put(st, f"E{r+2}", f"=SUM(E{POOL0+7}:E{POOL0+8})", fmt=NUM0, bold=True); ST["t_sc"] = r + 2
put(st, f"B{r+3}", "Total quota-carrying + support HC", "b"); put(st, f"D{r+3}", f"=SUM(D{r}:D{r+2})", fmt=NUM2); put(st, f"E{r+3}", f"=SUM(E{r}:E{r+2})", fmt=NUM0, bold=True); ST["t_all"] = r + 3

# C. allocation by segment of rounded HC
c0 = r + 5
section(st, c0, "C. Recommended HC allocated to segments (pooled roles split pro rata to workload)", 8)
header(st, c0 + 1, 2, ["Role"] + SEGS + ["Total"])
sdr_pool = {0: POOL0, 1: POOL0, 2: POOL0 + 1, 3: POOL0 + 2}
sc_pool = {0: POOL0 + 7, 1: POOL0 + 7, 2: POOL0 + 7, 3: POOL0 + 8}
for i, (role, src, poolmap) in enumerate([("SDR", 6, sdr_pool), ("AE", 7, None), ("SC", 8, sc_pool)]):
    rr = c0 + 2 + i
    put(st, f"B{rr}", role, "lbl")
    for k, col in enumerate(SEGCOLS):
        if poolmap is None:
            put(st, f"{col}{rr}", f"=E{POOL0+3+k}", fmt=NUM1)
        else:
            p = poolmap[k]
            put(st, f"{col}{rr}", f"={col}{src}/$D${p}*$E${p}", fmt=NUM1)
    put(st, f"G{rr}", f"=SUM(C{rr}:F{rr})", fmt=NUM1, bold=True)
ALLOC = {"sdr": c0 + 2, "ae": c0 + 3, "sc": c0 + 4}
rr = c0 + 5
put(st, f"B{rr}", "Total HC", "b")
for col in SEGCOLS + ["G"]:
    put(st, f"{col}{rr}", f"=SUM({col}{c0+2}:{col}{c0+4})", fmt=NUM1, bold=True)
ALLOC["tot"] = rr
rr += 1
put(st, f"B{rr}", "SDR : AE ratio (1 SDR per x AE)", "lbl")
for col in SEGCOLS + ["G"]:
    put(st, f"{col}{rr}", f"=IFERROR({col}{ALLOC['ae']}/{col}{ALLOC['sdr']},0)", fmt=NUM1)
ALLOC["sdr_ae"] = rr
rr += 1
put(st, f"B{rr}", "AE : SC ratio (x AE per 1 SC)", "lbl")
for col in SEGCOLS + ["G"]:
    put(st, f"{col}{rr}", f"=IFERROR({col}{ALLOC['ae']}/{col}{ALLOC['sc']},0)", fmt=NUM1)
ALLOC["ae_sc"] = rr

# D. Cost & ARR of recommended HC
d0 = rr + 2
section(st, d0, "D. Annual cost and expected ARR of recommended (rounded) HC: base scenario", 8)
header(st, d0 + 1, 2, ["Metric"] + SEGS + ["Total"])
D = {}
def drow(i, key, label, fn, fmt, tot="sum"):
    rr = d0 + 2 + i
    put(st, f"B{rr}", label, "lbl")
    for k, col in enumerate(SEGCOLS):
        put(st, f"{col}{rr}", fn(k, col), fmt=fmt)
    if tot == "sum":
        put(st, f"G{rr}", f"=SUM(C{rr}:F{rr})", fmt=fmt, bold=True)
    else:
        put(st, f"G{rr}", tot, fmt=fmt, bold=True)
    D[key] = rr
drow(0, "c_sdr", "SDR cost", lambda k, col: f"={col}{ALLOC['sdr']}*{eng('c_sdr1', 1, k)}", USD)
drow(1, "c_ae", "AE cost", lambda k, col: f"={col}{ALLOC['ae']}*{eng('c_ae1', 1, k)}", USD)
drow(2, "c_sc", "SC cost", lambda k, col: f"={col}{ALLOC['sc']}*{eng('c_sc1', 1, k)}", USD)
drow(3, "cost", "Total annual fully loaded cost", lambda k, col: f"=SUM({col}{d0+2}:{col}{d0+4})", USD)
drow(4, "arr", "Expected annual new ARR (base scenario)", lambda k, col: f"={eng('arr_y', 1, k)}", USD)
drow(5, "arr_rr", "  of which run-rate (no levers)", lambda k, col: f"={eng('arr_rr', 1, k)}", USD)
drow(6, "arr_inc", "  of which lever-driven (scenario)", lambda k, col: f"={eng('arr_l2d', 1, k)}+{eng('arr_wr', 1, k)}", USD)
drow(7, "eff", "ARR / cost", lambda k, col: f"={col}{d0+6}/{col}{d0+5}", MULT, f"=G{d0+6}/G{d0+5}")
drow(8, "payback", "Gross-margin payback (months)", lambda k, col: f"={col}{d0+5}/({col}{d0+6}*{IN['gm']}/12)", MO, f"=G{d0+5}/(G{d0+6}*{IN['gm']}/12)")
drow(9, "arr_ae", "ARR per AE", lambda k, col: f"={col}{d0+6}/{col}{ALLOC['ae']}", USD, f"=G{d0+6}/G{ALLOC['ae']}")
drow(10, "arr_share", "Share of ARR", lambda k, col: f"={col}{d0+6}/$G${d0+6}", PCT)
drow(11, "cost_share", "Share of cost", lambda k, col: f"={col}{d0+5}/$G${d0+5}", PCT)
for k, col in enumerate(SEGCOLS + ["G"]):
    st[f"{col}{D['eff']}"].fill = FILL_OUT
    st[f"{col}{D['payback']}"].fill = FILL_OUT

# E. Unit economics per incremental head
e0 = D["cost_share"] + 2
section(st, e0, "E. Unit economics of one additional ramped head (base rates; assumes demand exists to absorb it)", 8)
header(st, e0 + 1, 2, ["Metric"] + SEGS + ["Note"])
U = {}
def urow(i, key, label, fn, fmt, note=""):
    rr = e0 + 2 + i
    put(st, f"B{rr}", label, "lbl")
    for k, col in enumerate(SEGCOLS):
        put(st, f"{col}{rr}", fn(k, col), fmt=fmt)
    if note:
        st[f"G{rr}"].value, st[f"G{rr}"].font = note, F_NOTE
    U[key] = rr
urow(0, "ae_arr", "AE: ramped annual ARR capacity", lambda k, col: f"={eng('ae_cap',1,k)}*{IN['util']}*{eng('o_d2q',1,k)}*{eng('t_q2w',1,k)}*{eng('asp',1,k)}*{IN['months']}", USD, "Capacity x util x D->Q x Q->W x ASP x 12")
urow(1, "ae_rf", "AE: year-1 ramp factor", lambda k, col: f"=({IN['months']}-{si('ae_ramp',k)}*(1-{IN['rampprod']}))/{IN['months']}", PCT, "(12 - ramp x (1 - ramp productivity)) / 12")
urow(2, "ae_y1", "AE: year-1 ARR per new hire", lambda k, col: f"={col}{e0+2}*{col}{e0+3}", USD)
urow(3, "ae_c", "AE: loaded cost", lambda k, col: f"={eng('c_ae1',1,k)}", USD)
urow(4, "ae_x", "AE: ramped ARR / cost", lambda k, col: f"={col}{e0+2}/{col}{e0+5}", MULT)
urow(5, "ae_pod", "AE + attached SC & SDR: ramped ARR / cost", lambda k, col: f"={col}{e0+2}/({col}{e0+5}+{eng('sc',1,k)}/{eng('ae',1,k)}*{eng('c_sc1',1,k)}+{eng('sdr',1,k)}/{eng('ae',1,k)}*{eng('c_sdr1',1,k)})", MULT, "Includes SC and SDR support per AE at the base ratios")
urow(6, "sdr_stake", "SDR: ARR at stake per SDR IF the L2D gap is capacity-driven", lambda k, col: f"={eng('sdr_cap',1,k)}*{IN['util']}*{eng('gap',1,k)}*{eng('o_d2q',1,k)}*{eng('o_q2w',1,k)}*{eng('asp',1,k)}*{IN['months']}", USD, "HYPOTHESIS. Full gap, ramped. Zero for ENT (no erosion).")
urow(7, "sdr_c", "SDR: loaded cost", lambda k, col: f"={eng('c_sdr1',1,k)}", USD)
urow(8, "sdr_x", "SDR: ARR at stake / cost", lambda k, col: f"={col}{e0+8}/{col}{e0+9}", MULT)

# F. Current HC gap
f0 = U["sdr_x"] + 2
section(st, f0, "F. OPTIONAL: enter current ramped headcount to see the hiring gap (blank = unknown)", 8)
header(st, f0 + 1, 2, ["Role"] + SEGS + ["Total"])
for i, role in enumerate(["SDR", "AE", "SC"]):
    rr = f0 + 2 + i
    put(st, f"B{rr}", f"Current {role} (enter)", "lbl")
    for col in SEGCOLS:
        put(st, f"{col}{rr}", None, "in", NUM1, fill=FILL_IN)
    rg = f0 + 5 + i
    put(st, f"B{rg}", f"{role} gap = recommended - current", "lbl")
    src = [ALLOC["sdr"], ALLOC["ae"], ALLOC["sc"]][i]
    for col in SEGCOLS:
        put(st, f"{col}{rg}", f'=IF({col}{rr}="","enter current HC",{col}{src}-{col}{rr})', fmt=NUM1)
put(st, f"B{f0+8}", "The dataset has no headcount, so the model does not state a net-new hiring number. Required capacity minus current capacity = hiring need (before attrition backfill).", "note")

# G. Stress-test increment (gated)
g0 = f0 + 10
section(st, g0, "G. Jan-22 stress test: additional fractional FTE vs base. Hire only if demand persists (gated)", 8)
header(st, g0 + 1, 2, ["Role"] + SEGS + ["Total"])
for i, (role, key) in enumerate([("SDR", "sdr"), ("AE", "ae"), ("SC", "sc")]):
    rr = g0 + 2 + i
    put(st, f"B{rr}", role, "lbl")
    for k, col in enumerate(SEGCOLS):
        put(st, f"{col}{rr}", f"={eng(key,3,k)}-{eng(key,1,k)}", "lk", NUM2)
    put(st, f"G{rr}", f"=SUM(C{rr}:F{rr})", fmt=NUM2, bold=True)
ST["stress0"] = g0 + 2
put(st, f"B{g0+5}", "Trigger: 2 consecutive non-trough months at >= Jan-22 lead volume, with L2D holding. Use a contractor/overflow SDR first; AE hires lag by one ramp period.", "note")

h0 = g0 + 7
section(st, h0, "H. Scenario summary (fractional FTE, steady-state annual; from Scenario_Engine)", 8)
header(st, h0 + 1, 2, ["Metric"] + SCN + [""])
SUMM = [("Leads / month", "leads", NUM0), ("Demos / month", "demos", NUM0), ("Annual new ARR", "arr_y", USD), ("  Run-rate (no levers)", "arr_rr", USD),
        ("  L2D recovery lever", "arr_l2d", USD), ("  ENT win-rate lever", "arr_wr", USD), ("SDR FTE", "sdr", NUM1), ("AE FTE", "ae", NUM1), ("SC FTE", "sc", NUM1),
        ("Total FTE", "fte", NUM1), ("Total staffing cost", "cost", USD), ("ARR / cost", "eff", MULT), ("GM payback (months)", "payback", MO),
        ("Incremental ARR per incremental FTE (levers)", "d_arr_fte", USD), ("Incremental ARR / incremental cost (levers)", "d_eff", MULT)]
ST["summ"] = {}
for i, (lab, key, fmt) in enumerate(SUMM):
    rr = h0 + 2 + i
    put(st, f"B{rr}", lab, "lbl")
    for sidx, col in enumerate(SEGCOLS):
        put(st, f"{col}{rr}", f"={eng(key, sidx)}", "lk", fmt)
    ST["summ"][key] = rr

# ================================================================ ENT_Economics
ee = sheet("ENT_Economics", "Enterprise economics: test before scaling",
           "Per-AE pod economics at base rates, breakeven win rate for the payback hurdle, pilot design and sensitivities.",
           {"A": 4, "B": 46, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14, "I": 14, "J": 50})
K = 3  # ENT index
section(ee, 4, "A. Enterprise pod at base rates (1 ramped AE + attached SC & SDR)", 9)
EE = {}
def erow(r, key, label, f, fmt, kind="fx", note=""):
    put(ee, f"B{r}", label, "lbl")
    put(ee, f"C{r}", f, kind, fmt)
    if note:
        ee[f"D{r}"].value, ee[f"D{r}"].font = note, F_NOTE
    EE[key] = r
erow(5, "cap", "AE demos / month", f"={si('ae_cap',K)}", NUM0, "lk")
erow(6, "util", "Utilization", f"={IN['util']}", PCT, "lk")
erow(7, "d2q", "Demo->Opp (baseline window)", f"=Demand!F{DM['b_d2q']}", PCT, "lk")
erow(8, "q2w", "Opp->Win (baseline window)", f"=Demand!F{DM['b_q2w']}", PCT, "lk")
erow(9, "asp", "ARR / logo (baseline window)", f"=Demand!F{DM['b_asp']}", USD, "lk")
erow(10, "l2d", "Lead->Demo (baseline window)", f"=Demand!F{DM['b_l2d']}", PCT, "lk")
erow(11, "opps", "Qualified opps / AE / year", "=C5*C6*C7*12", NUM1)
erow(12, "arr", "ARR / AE / year", "=C11*C8*C9", USD)
erow(13, "sc_per", "SC per AE (standard coverage)", f"=C5*{si('sc_inv',K)}/{si('sc_cap',K)}", NUM2, note="AE demos x SC involvement / SC capacity")
erow(14, "sdr_per", "SDR per AE", f"=(C5/C10)/{si('sdr_cap',K)}", NUM2, note="Leads needed per AE / SDR capacity (both at the same utilization)")
erow(15, "cost", "Pod loaded cost / year", f"={si('ae_ote',K)}*{IN['load']}+C13*{si('sc_ote',K)}*{IN['load']}+C14*{IN['sdr_ote']}*{IN['load']}", USD)
erow(16, "x", "ARR / cost", "=C12/C15", MULT)
erow(17, "pb", "GM payback (months)", f"=C15/(C12*{IN['gm']}/12)", MO)
erow(18, "be_q2w", "Breakeven Opp->Win for payback hurdle", f"=C15*12/({IN['gm']}*{IN['hurdle']})/(C11*C9)", PCT, note="Win rate at which the pod pays back within the hurdle")
erow(19, "be_up", "Required uplift vs today (pts)", "=C18-C8", PCT)
erow(20, "be_asp", "Breakeven ARR/logo at today's win rate", f"=C15*12/({IN['gm']}*{IN['hurdle']})/(C11*C8)", USD)
erow(21, "cmp", "Compare: Emerging MM pod ARR / cost", f"=Staffing!D{U['ae_pod']}", MULT, "lk", "Same method from the Staffing tab. Shows where the marginal dollar works hardest.")

section(ee, 23, "B. SC-paired pilot design (illustrative)", 9)
erow(24, "p_ae", "Pilot AEs", f"={IN['pil_ae']}", NUM0, "lk")
erow(25, "p_sc", "Pilot SCs (1:1)", "=C24", NUM0)
erow(26, "p_mo", "Duration (months)", f"={IN['pil_mo']}", NUM0, "lk")
erow(27, "p_opps", "Pilot qualified opps in window", "=C24*C5*C6*C7*C26", NUM1)
erow(28, "p_se", "Std error of win rate at that sample (pts)", "=SQRT(C8*(1-C8)/C27)", PCT, note="With ~80 opps, only a ~+9pt effect is statistically clear. Read it as directional, or extend the pilot.")
erow(29, "p_mde", "Approx. minimum detectable uplift (2 x SE)", "=2*C28", PCT)
erow(30, "p_cost", "Incremental pilot cost (extra SC above standard coverage)", f"=(C25-C24*C13)*{si('sc_ote',K)}*{IN['load']}*C26/12", USD)
erow(31, "p_gate1", "Scale gate 1: pilot win rate >=", "=C18", PCT, note="Equals the breakeven win rate. Below this, do not scale ENT headcount.")
erow(32, "p_gate2", "Scale gate 2: ARR/logo >=", "=C9", USD, note="No ASP erosion vs the baseline window")
put(ee, "B33", "Other gates: sales cycle no longer than the control group's; SC hours per opp tracked; win rate measured on opps created during the pilot (cohort basis).", "note")

section(ee, 35, "C. Sensitivity: GM payback (months) by Opp->Win (rows) and ARR/logo (cols), pod at base capacity & cost", 9)
asps = [18000, 21000, 24000, 27000, 30000, 35000]
q2ws = [0.15, 0.175, 0.20, 0.225, 0.25, 0.275, 0.30]
put(ee, "B36", "Opp->Win  \\  ARR/logo", "b")
for j, a in enumerate(asps):
    put(ee, f"{CL(3+j)}36", a, "in", USD)
for i, q in enumerate(q2ws):
    rr = 37 + i
    put(ee, f"B{rr}", q, "in", PCT)
    for j in range(len(asps)):
        cc = CL(3 + j)
        put(ee, f"{cc}{rr}", f"=$C$15/($C$11*$B{rr}*{cc}$36*{IN['gm']}/12)", fmt=MO)
section(ee, 45, "D. Sensitivity: GM payback (months) by Opp->Win (rows) and AE demo capacity (cols), today's ASP", 9)
caps = [8, 10, 12, 15, 18, 20]
put(ee, "B46", "Opp->Win  \\  demos/AE/mo", "b")
for j, a in enumerate(caps):
    put(ee, f"{CL(3+j)}46", a, "in", NUM0)
for i, q in enumerate(q2ws):
    rr = 47 + i
    put(ee, f"B{rr}", q, "in", PCT)
    for j in range(len(caps)):
        cc = CL(3 + j)
        cost = f"({si('ae_ote',K)}*{IN['load']}+{cc}$46*{si('sc_inv',K)}/{si('sc_cap',K)}*{si('sc_ote',K)}*{IN['load']}+({cc}$46/$C$10)/{si('sdr_cap',K)}*{IN['sdr_ote']}*{IN['load']})"
        put(ee, f"{cc}{rr}", f"={cost}/({cc}$46*$C$6*$C$7*12*$B{rr}*$C$9*{IN['gm']}/12)", fmt=MO)
put(ee, "B55", "Cells at or below the payback hurdle are where scaling ENT pays off. Today's ENT sits near 19-20% win rate and ~$23K ARR/logo.", "note")

# ================================================================ Sensitivity
se = sheet("Sensitivity", "One-at-a-time sensitivity on the Base scenario (exact, formula-driven)",
           "FTE scales with 1/capacity and 1/utilization. Cost scales with OTE x load. ARR levers are linear. Recommended HC uses fractional FTE here (before rounding).",
           {"A": 4, "B": 44, "C": 12, "D": 12, "E": 14, "F": 14, "G": 14, "H": 14, "I": 15, "J": 15, "K": 12, "L": 12})
t = BASE[2]
B_ = lambda key: f"Scenario_Engine!${t}${EN[key]}"
put(se, "B4", "Base totals", "b")
base_items = [("SDR FTE", "sdr", NUM2), ("AE FTE", "ae", NUM2), ("SC FTE", "sc", NUM2), ("SDR cost", "c_sdr", USD), ("AE cost", "c_ae", USD), ("SC cost", "c_sc", USD),
              ("Annual ARR", "arr_y", USD), ("Run-rate ARR", "arr_rr", USD), ("L2D-lever ARR", "arr_l2d", USD), ("Win-rate-lever ARR", "arr_wr", USD),
              ("AE FTE w/o levers", "ae0", NUM2), ("SC FTE w/o levers & pilot", "sc0", NUM2)]
SB = {}
for i, (lab, key, fmt) in enumerate(base_items):
    rr = 5 + i
    put(se, f"B{rr}", lab, "lbl")
    put(se, f"C{rr}", f"={B_(key)}", "lk", fmt)
    SB[key] = f"$C${rr}"
put(se, "B17", "Total FTE", "lbl"); put(se, "C17", "=C5+C6+C7", fmt=NUM2); SB["fte"] = "$C$17"
put(se, "B18", "Total cost", "lbl"); put(se, "C18", "=C8+C9+C10", fmt=USD); SB["cost"] = "$C$18"
put(se, "B19", "Base recovery %", "lbl"); put(se, "C19", f"={sc('rec',1)}", "lk", PCT); SB["rec"] = "$C$19"
put(se, "B20", "Base ENT win-rate uplift", "lbl"); put(se, "C20", f"={sc('wr',1)}", "lk", PCT); SB["wr"] = "$C$20"

header(se, 22, 2, ["Assumption flexed", "Low", "High", "FTE @ low", "FTE @ high", "Cost @ low", "Cost @ high", "ARR @ low", "ARR @ high", "Payback @ low", "Payback @ high"])
S = SB
def fte_cost(sdr_m="1", ae_m="1", sc_m="1", c_m="1", ae_cost_m="1", arr_extra="0", ae_add="0", sc_add="0"):
    fte = f"({S['sdr']}*{sdr_m}+({S['ae']}+{ae_add})*{ae_m}+({S['sc']}+{sc_add})*{sc_m})"
    cost = f"(({S['c_sdr']}*{sdr_m}+{S['c_ae']}*{ae_m}*{ae_cost_m}*(1+({ae_add})/{S['ae']})+{S['c_sc']}*{sc_m}*(1+({sc_add})/{S['sc']}))*{c_m})"
    arr = f"({S['arr_y']}+{arr_extra})"
    return fte, cost, arr
rows = []
# (label, low val, high val, fmt, builder(v) -> (fte,cost,arr)) ; v is cell ref to low/high
rows.append(("SDR capacity (x base)", 0.8, 1.2, "0.00x", lambda v: fte_cost(sdr_m=f"(1/{v})")))
rows.append(("AE capacity (x base)", 0.8, 1.2, "0.00x", lambda v: fte_cost(ae_m=f"(1/{v})")))
rows.append(("SC capacity (x base)", 0.8, 1.2, "0.00x", lambda v: fte_cost(sc_m=f"(1/{v})")))
rows.append(("Utilization", 0.75, 0.95, PCT, lambda v: fte_cost(sdr_m=f"({IN['util']}/{v})", ae_m=f"({IN['util']}/{v})", sc_m=f"({IN['util']}/{v})")))
rows.append(("Loaded-cost multiplier", 1.15, 1.35, NUM2, lambda v: fte_cost(c_m=f"({v}/{IN['load']})")))
rows.append(("AE OTE (x base)", 0.85, 1.15, "0.00x", lambda v: fte_cost(ae_cost_m=v)))
rows.append(("L2D recovery (% of gap)", 0.0, 1.0, PCT, lambda v: fte_cost(arr_extra=f"{S['arr_l2d']}*({v}/{S['rec']}-1)", ae_add=f"({S['ae']}-{S['ae0']})*({v}/{S['rec']}-1)",
                                                                            sc_add=f"0")))
rows.append(("ENT win-rate uplift (pts)", 0.0, 0.06, PCT, lambda v: fte_cost(arr_extra=f"{S['arr_wr']}*({v}/{S['wr']}-1)")))
rows.append(("Demand (x baseline leads)", 0.85, 1.15, "0.00x", lambda v: fte_cost(sdr_m=v, ae_m=v, sc_m=v, arr_extra=f"{S['arr_y']}*({v}-1)")))
for i, (lab, lo, hi, fmt, fn) in enumerate(rows):
    rr = 23 + i
    put(se, f"B{rr}", lab, "lbl")
    put(se, f"C{rr}", lo, "in", fmt)
    put(se, f"D{rr}", hi, "in", fmt)
    for j, v in enumerate([f"$C{rr}", f"$D{rr}"]):
        fte, cost, arr = fn(v)
        put(se, f"{CL(5+j)}{rr}", "=" + fte, fmt=NUM1)
        put(se, f"{CL(7+j)}{rr}", "=" + cost, fmt=USD)
        put(se, f"{CL(9+j)}{rr}", "=" + arr, fmt=USD)
        put(se, f"{CL(11+j)}{rr}", f"={CL(7+j)}{rr}/({CL(9+j)}{rr}*{IN['gm']}/12)", fmt=MO)
SENS_END = 23 + len(rows) - 1
put(se, f"B{SENS_END+2}", "Notes: the L2D-recovery row moves AE FTE with the extra demos (SC effect ignored, <0.2 FTE). Demand scales all roles and ARR proportionally.", "note")
put(se, f"B{SENS_END+3}", "The SC-pilot cost is held constant in the win-rate row: the pilot is a decision, the uplift is the outcome.", "note")

# ================================================================ Checks
ck = sheet("Checks", "Reconciliation & integrity checks", "Every check should read TRUE.", {"A": 4, "B": 70, "C": 18, "D": 18, "E": 10})
header(ck, 4, 2, ["Check", "Value A", "Value B", "Pass?"])
checks = [
    ("Data rows = 60 (12 months x 5 buckets)", f"=COUNTA({R('B')})", "=60"),
    ("Analysis total leads = SUM(Data leads)", "=Analysis!B11", f"=SUM({R('C')})"),
    ("Analysis total ARR = SUM(Data ARR)", "=Analysis!F11", f"=SUM({R('G')})"),
    ("Monthly table ARR = bucket table ARR", "=Analysis!F27", "=Analysis!F11"),
    ("Segmentation (4-tier) ARR = total ARR", f"=SUM(Segmentation!D{SEG_ROWS['seg'][0]}:D{SEG_ROWS['seg'][-1]})", "=Analysis!F11"),
    ("Option A ARR = total ARR", f"=SUM(Segmentation!D{SEG_ROWS['optA'][0]}:D{SEG_ROWS['optA'][-1]})", "=Analysis!F11"),
    ("Baseline-window ARR/mo (Demand) = Data window ARR / months", f"=Demand!G{DM['b_arr']}", f"=SUMIFS({R('G')},{R('L')},1)/Demand!C8"),
    ("Stress ARR/mo (Demand) = Data Jan-22 ARR", f"=Demand!G{DM['s_arr']}", "=Analysis!F26"),
    ("Base run-rate ARR (engine) = baseline ARR/mo x 12", f"={eng('arr_rr',1)}", f"=Demand!G{DM['b_arr']}*12"),
    ("Stress run-rate ARR (engine) = Jan-22 ARR x 12", f"={eng('arr_rr',3)}", "=Analysis!F26*12"),
    ("Engine attribution sums to total (all scenarios)", f"=SUM(Scenario_Engine!C{EN['arr_chk']}:{blocks[-1][2]}{EN['arr_chk']})", "=0"),
    ("Engine base demos at 0% recovery = observed demos (L2D x leads)", f"=Scenario_Engine!{BASE[2]}{EN['leads']}*Scenario_Engine!{BASE[2]}{EN['o_l2d']}", f"=Demand!G{DM['b_demos']}"),
    ("Staffing rounded SDR HC >= fractional - 0.3", f"=Staffing!E{ST['t_sdr']}", f"=Staffing!D{ST['t_sdr']}-3*{IN['round']}"),
    ("Staffing allocated HC total = pooled HC total", f"=Staffing!G{ALLOC['tot']}", f"=Staffing!E{ST['t_all']}"),
    ("Staffing base ARR = engine base ARR", f"=Staffing!G{D['arr']}", f"={eng('arr_y',1)}"),
    ("Analysis LMDI effects sum to delta ARR", "=Analysis!L147", "=0"),
    ("Funnel monotonic in every row (leads>=demos>=opps>=logos)", f"=SUMPRODUCT(({R('D')}>{R('C')})+({R('E')}>{R('D')})+({R('F')}>{R('E')}))", "=0"),
]
for i, (lab, a, b) in enumerate(checks):
    rr = 5 + i
    put(ck, f"B{rr}", lab, "lbl")
    put(ck, f"C{rr}", a, "lk", NUM2)
    put(ck, f"D{rr}", b, "fx", NUM2)
    if "fractional" in lab:
        put(ck, f"E{rr}", f"=C{rr}>=D{rr}")
    else:
        put(ck, f"E{rr}", f"=ABS(C{rr}-D{rr})<0.01")
CK_END = 5 + len(checks) - 1
put(ck, f"B{CK_END+2}", "ALL CHECKS PASS", "b")
put(ck, f"E{CK_END+2}", f"=AND(E5:E{CK_END})", bold=True)

# order sheets
order = ["README", "Inputs", "Data", "Seg_Map", "Analysis", "Segmentation", "Demand", "Scenario_Engine", "Staffing", "ENT_Economics", "Sensitivity", "Checks"]
wb._sheets = [wb[n] for n in order]
for wsx in wb.worksheets:
    wsx.sheet_properties.tabColor = {"Inputs": "0000FF", "Staffing": "008000", "Checks": "C00000"}.get(wsx.title, "1F2A44")
wb.save(OUT)
print("saved", OUT)
import json
pass  # json.dump({"EN": EN, "DM": DM, "ST": ST, "ALLOC": ALLOC, "D": D, "U": U, "EE": EE, "POOL0": POOL0, "blocks": blocks, "SEG_ROWS": SEG_ROWS}, open(OUT + ".map.json", "w"), default=str)
