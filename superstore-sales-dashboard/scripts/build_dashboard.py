"""Build the Superstore Sales Dashboard workbook (formulas, named ranges, charts).

Usage: python scripts/build_dashboard.py [raw_csv] [output_xlsx]
Then recalculate (LibreOffice) so cached values exist:  python recalc.py <output_xlsx>
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import CharacterProperties, Font as DFont, Paragraph, ParagraphProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parents[1]
RAW = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data" / "superstore_train_raw.csv"
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "dashboard" / "Superstore_Sales_Dashboard.xlsx"

# ----------------------------------------------------------------------------- 1. clean
df = pd.read_csv(RAW)
log = []  # (step, finding, action, rows affected)
log.append(("Import", f"{len(df):,} rows x {df.shape[1]} columns loaded from train.csv", "None", len(df)))
for c in ("Order Date", "Ship Date"):
    df[c] = pd.to_datetime(df[c], format="%d/%m/%Y")
log.append(("Dates", "Order/Ship Date stored as text in dd/mm/yyyy", "Converted to true Excel dates (day-first parsing)", len(df)))
assert (df["Ship Date"] >= df["Order Date"]).all()
log.append(("Date logic", "Ship Date earlier than Order Date", "None found - no action", 0))
txt = df.select_dtypes(include=["object", "string"]).columns
n_ws = int(sum((df[c] != df[c].str.strip()).sum() for c in txt))
df[txt] = df[txt].apply(lambda s: s.str.strip())
log.append(("Whitespace", "Leading/trailing spaces in text fields", "Trimmed (none found)" if n_ws == 0 else "Trimmed", n_ws))
miss = df["Postal Code"].isna()
df.loc[miss & (df.City == "Burlington") & (df.State == "Vermont"), "Postal Code"] = 5401
log.append(("Missing values", "Postal Code blank for Burlington, Vermont", "Filled with 05401 (Burlington, VT ZIP). Not used in any metric", int(miss.sum())))
df["Postal Code"] = df["Postal Code"].astype(int)
dup_cols = [c for c in df.columns if c != "Row ID"]
n_dup = int(df.duplicated(dup_cols, keep=False).sum())
log.append(("Duplicates", "Rows identical in every field except Row ID (Order US-2015-150119, 2 lines, $281.37 each)",
            "KEPT and flagged: file has no Quantity column, so these may be 2 units of one item. Impact 0.01% of revenue", n_dup))
log.append(("Columns", "'Country' is constant (United States); no Quantity, Discount or Profit columns exist in this extract",
            "Country kept; 'units sold' is proxied by order lines (see Notes)", 0))
for k in ("Region", "Segment", "Order Date"):
    assert (df.groupby("Order ID")[k].nunique() == 1).all()
df = df.sort_values(["Order Date", "Order ID", "Category", "Row ID"]).reset_index(drop=True)
N = len(df)
LAST = N + 1

# ----------------------------------------------------------------------------- styles
NAVY, BLUE, TEAL, GREY, LIGHT, WHITE = "1F3A5F", "2E86AB", "F18F01", "6B7280", "EEF3F8", "FFFFFF"
AR = "Arial"
f = lambda **k: Font(name=AR, **k)
thin = Side(style="thin", color="D0D7E2")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
fill = lambda c: PatternFill("solid", start_color=c, end_color=c)
USD, USD2, PCT, INT = '"$"#,##0', '"$"#,##0.00', '0.0%', '#,##0'
PCT_COL = '[Color10]+0.0%;[Red]-0.0%;0.0%'

wb = Workbook()
ws_d, ws_a, ws_i, ws_s, ws_n = wb.active, wb.create_sheet("Analysis"), wb.create_sheet("Insights"), wb.create_sheet("Data"), wb.create_sheet("Notes")
ws_d.title = "Dashboard"

# ----------------------------------------------------------------------------- 2. Data sheet
cols = ["Row ID", "Order ID", "Order Date", "Ship Date", "Ship Mode", "Customer ID", "Customer Name", "Segment", "Country",
        "City", "State", "Postal Code", "Region", "Product ID", "Category", "Sub-Category", "Product Name", "Sales"]
helpers = ["Year", "Month", "Month No", "Ship Days", "New Order", "New Order (Cat)"]
ws_s.append(cols + helpers)
for i, row in enumerate(df[cols].itertuples(index=False), start=2):
    ws_s.append(list(row[:2]) + [row[2].to_pydatetime(), row[3].to_pydatetime()] + list(row[4:]) + [
        f"=YEAR(C{i})", f"=DATE(YEAR(C{i}),MONTH(C{i}),1)", f"=MONTH(C{i})", f"=D{i}-C{i}",
        f"=IF(B{i}=B{i-1},0,1)", f"=IF(AND(B{i}=B{i-1},O{i}=O{i-1}),0,1)"])
for c in range(1, len(cols) + len(helpers) + 1):
    h = ws_s.cell(1, c)
    h.font, h.fill, h.alignment = f(bold=True, color=WHITE), fill(NAVY if c <= len(cols) else BLUE), Alignment(horizontal="center", vertical="center", wrap_text=True)
widths = [8, 16, 12, 12, 14, 12, 22, 13, 14, 18, 16, 10, 10, 17, 16, 13, 48, 11, 7, 11, 9, 9, 10, 12]
for i, w in enumerate(widths, 1):
    ws_s.column_dimensions[L(i)].width = w
for r in ws_s.iter_rows(min_row=2, max_row=LAST):
    for c in r:
        c.font = f(size=9)
    r[2].number_format = r[3].number_format = "dd-mmm-yyyy"
    r[11].number_format = "00000"
    r[17].number_format = USD2
    r[19].number_format = "mmm-yyyy"
ws_s.freeze_panes = "C2"
ws_s.row_dimensions[1].height = 30
tbl = Table(displayName="tblSales", ref=f"A1:{L(len(cols)+len(helpers))}{LAST}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
ws_s.add_table(tbl)
ws_s.sheet_properties.tabColor = GREY

def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)
for n, col in dict(d_Order="B", d_Segment="H", d_Region="M", d_Category="O", d_Sub="P", d_State="K", d_Sales="R", d_Year="S",
                   d_Month="T", d_MonthNo="U", d_Ship="V", d_NewOrder="W", d_NewOrderCat="X").items():
    name(n, f"Data!${col}$2:${col}${LAST}")

# ----------------------------------------------------------------------------- 3. Dashboard controls (referenced by names)
D = ws_d
name("sel_Region", "Dashboard!$D$4"); name("sel_Category", "Dashboard!$H$4"); name("sel_Segment", "Dashboard!$L$4"); name("sel_Year", "Dashboard!$P$4")
name("crit_Region", "Analysis!$B$2"); name("crit_Category", "Analysis!$B$3"); name("crit_Segment", "Analysis!$B$4")
name("crit_YearMin", "Analysis!$E$2"); name("crit_YearMax", "Analysis!$E$3"); name("yoy_Year", "Analysis!$E$4")

FIL = 'd_Region,crit_Region,d_Category,crit_Category,d_Segment,crit_Segment,d_Year,">="&crit_YearMin,d_Year,"<="&crit_YearMax'
FIL_NY = 'd_Region,crit_Region,d_Category,crit_Category,d_Segment,crit_Segment'
REV = lambda fil, extra="": f"SUMIFS(d_Sales,{fil}{',' + extra if extra else ''})"
def ORD(fil, extra="", cat=False):
    a, b = (f"SUMIFS(d_NewOrder,{fil}{',' + extra if extra else ''})", f"SUMIFS(d_NewOrderCat,{fil}{',' + extra if extra else ''})")
    return b if cat else f"IF(sel_Category=\"All\",{a},{b})"
LIN = lambda fil, extra="": f"COUNTIFS({fil}{',' + extra if extra else ''})"

# ----------------------------------------------------------------------------- 4. Analysis sheet
A = ws_a
A.sheet_view.showGridLines = False
A["A1"] = "Analysis - pivot-style summary tables (live formulas, respond to the Dashboard filters)"
A["A1"].font = f(bold=True, size=14, color=NAVY)
for r, (lab, src, fm) in enumerate([("Region criterion", "sel_Region", '=IF(sel_Region="All","*",sel_Region)'),
                                    ("Category criterion", "sel_Category", '=IF(sel_Category="All","*",sel_Category)'),
                                    ("Segment criterion", "sel_Segment", '=IF(sel_Segment="All","*",sel_Segment)')], start=2):
    A.cell(r, 1, lab).font = f(size=9, color=GREY); A.cell(r, 2, fm).font = f(size=9)
for r, (lab, fm) in enumerate([("Year from", '=IF(sel_Year="All",MIN(d_Year),VALUE(sel_Year))'),
                               ("Year to", '=IF(sel_Year="All",MAX(d_Year),VALUE(sel_Year))'),
                               ("YoY focus year", '=IF(sel_Year="All",MAX(d_Year),VALUE(sel_Year))')], start=2):
    A.cell(r, 4, lab).font = f(size=9, color=GREY); A.cell(r, 5, fm).font = f(size=9)

def header(ws, r, labels, c0=1):
    for i, t in enumerate(labels):
        c = ws.cell(r, c0 + i, t)
        c.font, c.fill, c.alignment, c.border = f(bold=True, color=WHITE, size=10), fill(NAVY), Alignment(horizontal="center", vertical="center", wrap_text=True), box
def title(ws, r, t, note=None):
    ws.cell(r, 1, t).font = f(bold=True, size=12, color=NAVY)
    if note:
        ws.cell(r + 1, 1, note).font = f(italic=True, size=9, color=GREY)
def put(ws, r, c, v, fmt=None, bold=False, color="000000"):
    x = ws.cell(r, c, v); x.font = f(size=10, bold=bold, color=color); x.border = box
    if fmt: x.number_format = fmt
    return x

pos = {}
# 4a monthly trend ------------------------------------------------------------
r0 = 7
title(A, r0, "1. Monthly trend (always the full period; Region / Category / Segment filters apply)")
header(A, r0 + 1, ["Month", "Revenue ($)", "Orders", "Order Lines", "Avg Order Value ($)", "MoM %", "YoY %", "3-Mo Moving Avg ($)"])
m0 = r0 + 2
NM = 48
for k in range(NM):
    r = m0 + k
    put(A, r, 1, "=MIN(d_Month)" if k == 0 else f"=EDATE(A{r-1},1)", "mmm-yyyy")
    ex = f"d_Month,$A{r}"
    put(A, r, 2, "=" + REV(FIL_NY, ex), USD)
    put(A, r, 3, "=" + ORD(FIL_NY, ex), INT)
    put(A, r, 4, "=" + LIN(FIL_NY, ex), INT)
    put(A, r, 5, f"=IF(C{r}=0,0,B{r}/C{r})", USD)
    put(A, r, 6, f'=IF(B{r-1}=0,"",B{r}/B{r-1}-1)' if k >= 1 else "", PCT_COL)
    put(A, r, 7, f'=IF(B{r-12}=0,"",B{r}/B{r-12}-1)' if k >= 12 else "", PCT_COL)
    put(A, r, 8, f"=AVERAGE(B{r-2}:B{r})" if k >= 2 else "", USD)
pos["month"] = (m0, m0 + NM - 1)
rt = m0 + NM
put(A, rt, 1, "Total", bold=True); put(A, rt, 2, f"=SUM(B{m0}:B{rt-1})", USD, True); put(A, rt, 3, f"=SUM(C{m0}:C{rt-1})", INT, True); put(A, rt, 4, f"=SUM(D{m0}:D{rt-1})", INT, True)
put(A, rt, 5, f"=IF(C{rt}=0,0,B{rt}/C{rt})", USD, True)

# 4b yearly ---------------------------------------------------------------------
r0 = rt + 3
title(A, r0, "2. Yearly summary (Region / Category / Segment filters apply)")
header(A, r0 + 1, ["Year", "Revenue ($)", "Orders", "Order Lines", "Avg Order Value ($)", "YoY %"])
y0 = r0 + 2
for k in range(4):
    r = y0 + k
    put(A, r, 1, "=MIN(d_Year)" if k == 0 else f"=A{r-1}+1", "0")
    ex = f"d_Year,$A{r}"
    put(A, r, 2, "=" + REV(FIL_NY, ex), USD); put(A, r, 3, "=" + ORD(FIL_NY, ex), INT); put(A, r, 4, "=" + LIN(FIL_NY, ex), INT)
    put(A, r, 5, f"=IF(C{r}=0,0,B{r}/C{r})", USD)
    put(A, r, 6, f'=IF(B{r-1}=0,"",B{r}/B{r-1}-1)' if k else "", PCT_COL)
pos["year"] = (y0, y0 + 3)

# 4c category / region / segment ---------------------------------------------------
def dim_table(r0, ttl, dimname, labels, nm, catflag=False):
    title(A, r0, ttl)
    header(A, r0 + 1, [dimname, "Revenue ($)", "% of Revenue", "Orders", "Order Lines", "Avg Order Value ($)"])
    s = r0 + 2
    for k, lab in enumerate(labels):
        r = s + k
        put(A, r, 1, lab)
        ex = f"{nm},$A{r}"
        put(A, r, 2, "=" + REV(FIL, ex), USD)
        put(A, r, 3, f"=IF(SUM($B${s}:$B${s+len(labels)-1})=0,0,B{r}/SUM($B${s}:$B${s+len(labels)-1}))", PCT)
        put(A, r, 4, "=" + ORD(FIL, ex, cat=catflag), INT)
        put(A, r, 5, "=" + LIN(FIL, ex), INT)
        put(A, r, 6, f"=IF(D{r}=0,0,B{r}/D{r})", USD)
    return s, s + len(labels) - 1
cats = sorted(df.Category.unique()); regs = sorted(df.Region.unique()); segs = sorted(df.Segment.unique())
pos["cat"] = dim_table(y0 + 7, "3. Revenue by Category (all filters apply)", "Category", cats, "d_Category", True)
pos["reg"] = dim_table(pos["cat"][1] + 3, "4. Revenue by Region (all filters apply)", "Region", regs, "d_Region")
pos["seg"] = dim_table(pos["reg"][1] + 3, "5. Revenue by Customer Segment (all filters apply)", "Segment", segs, "d_Segment")

# 4d ranked tables (sub-category, state) ---------------------------------------------
def ranked(r0, ttl, labels, nm, topn, hdr):
    title(A, r0, ttl, "Left block = unsorted calculation; right block = ranked with LARGE/INDEX/MATCH (no dynamic-array functions needed).")
    header(A, r0 + 2, [hdr, "Revenue ($)", "Rank key"] , 1)
    header(A, r0 + 2, ["Rank", hdr, "Revenue ($)", "% of Revenue"], 5)
    s = r0 + 3; n = len(labels)
    for k, lab in enumerate(labels):
        r = s + k
        put(A, r, 1, lab); put(A, r, 2, "=" + REV(FIL, f"{nm},$A{r}"), USD)
        put(A, r, 3, f"=B{r}+ROW()/1000000", '0.000000')
    for k in range(topn):
        r = s + k
        put(A, r, 5, k + 1, "0")
        put(A, r, 6, f"=INDEX($A${s}:$A${s+n-1},MATCH(LARGE($C${s}:$C${s+n-1},E{r}),$C${s}:$C${s+n-1},0))")
        put(A, r, 7, f"=INDEX($B${s}:$B${s+n-1},MATCH(LARGE($C${s}:$C${s+n-1},E{r}),$C${s}:$C${s+n-1},0))", USD)
        put(A, r, 8, f"=IF(SUM($B${s}:$B${s+n-1})=0,0,G{r}/SUM($B${s}:$B${s+n-1}))", PCT)
    return s, s + topn - 1, s + n - 1
subs = sorted(df["Sub-Category"].unique()); states = sorted(df.State.unique())
pos["sub"] = ranked(pos["seg"][1] + 3, "6. Sub-Category ranking (all filters apply)", subs, "d_Sub", len(subs), "Sub-Category")
pos["state"] = ranked(pos["sub"][2] + 3, "7. State ranking (all filters apply) - top 10 shown", states, "d_State", 10, "State")

# 4e seasonality ----------------------------------------------------------------------
r0 = pos["state"][2] + 3
title(A, r0, "8. Seasonality - revenue by calendar month (all filters apply; with Year = All it sums all years)")
header(A, r0 + 1, ["Month", "Revenue ($)", "% of Revenue"])
s0 = r0 + 2
mn = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
for k in range(12):
    r = s0 + k
    put(A, r, 1, mn[k]); put(A, r, 2, "=" + REV(FIL, f"d_MonthNo,{k+1}"), USD)
    put(A, r, 3, f"=IF(SUM($B${s0}:$B${s0+11})=0,0,B{r}/SUM($B${s0}:$B${s0+11}))", PCT)
pos["season"] = (s0, s0 + 11)
for i, w in enumerate([22, 16, 14, 14, 18, 22, 16, 18], 1):
    A.column_dimensions[L(i)].width = w
A.sheet_properties.tabColor = BLUE
A.freeze_panes = "A7"

# ----------------------------------------------------------------------------- 5. Dashboard
D.sheet_view.showGridLines = False
D.sheet_properties.tabColor = NAVY
for i in range(1, 19):
    D.column_dimensions[L(i)].width = 12.3
D.column_dimensions["A"].width = 2; D.column_dimensions["R"].width = 2
D.merge_cells("B1:Q1"); D.merge_cells("B2:Q2")
D["B1"] = "Superstore Sales Performance Dashboard"; D["B1"].font = f(bold=True, size=22, color=WHITE)
D["B2"] = "Jan 2015 - Dec 2018  |  9,800 order lines  |  use the dropdowns below to filter every KPI, table and chart"; D["B2"].font = f(size=10, color="DCE6F2")
for r in (1, 2):
    for c in range(1, 19):
        D.cell(r, c).fill = fill(NAVY)
D.row_dimensions[1].height = 36; D.row_dimensions[2].height = 20
D.row_dimensions[3].height = 8
filters = [("B", "C", "D", "E", "REGION", ["All"] + regs), ("F", "G", "H", "I", "CATEGORY", ["All"] + cats),
           ("J", "K", "L", "M", "SEGMENT", ["All"] + segs), ("N", "O", "P", "Q", "YEAR", ["All"] + sorted(df["Order Date"].dt.year.unique().astype(str)))]
D.row_dimensions[4].height = 24
for lc, _, vc, ve, lab, opts in filters:
    D.merge_cells(f"{lc}4:{_}4")
    D[f"{lc}4"] = lab + " ▾"; D[f"{lc}4"].font = f(bold=True, size=10, color=NAVY); D[f"{lc}4"].alignment = Alignment(horizontal="right", vertical="center")
    D.merge_cells(f"{vc}4:{ve}4")
    D[f"{vc}4"] = "All"
    D[f"{vc}4"].font = f(bold=True, size=11, color=NAVY); D[f"{vc}4"].alignment = Alignment(horizontal="center", vertical="center")
    for c in (vc, ve):
        D[f"{c}4"].fill = fill("FFF4D6"); D[f"{c}4"].border = Border(top=Side(style="medium", color=TEAL), bottom=Side(style="medium", color=TEAL),
                                                                        left=Side(style="medium", color=TEAL) if c == vc else None, right=Side(style="medium", color=TEAL) if c == ve else None)
    dv = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=False, showDropDown=False)
    dv.error, dv.errorTitle = "Pick a value from the list", "Invalid filter"
    D.add_data_validation(dv); dv.add(f"{vc}4")
D.row_dimensions[5].height = 8

# KPI tiles (6 tiles x ~2.67 cols): use column pairs B:C, E:F... -> 6 tiles over B..Q = 16 cols; use 2 cols each + spacers
tiles = [("TOTAL REVENUE", "=" + REV(FIL), USD, '="Filtered selection"'),
         ("ORDERS", "=" + ORD(FIL), INT, '="Distinct Order IDs"'),
         ("ORDER LINES (UNITS PROXY)", "=" + LIN(FIL), INT, '="No quantity field in data"'),
         ("AVG ORDER VALUE", '=IFERROR(' + REV(FIL) + "/" + ORD(FIL) + ",0)", USD, '="Revenue / orders"'),
         ("AVG DAYS TO SHIP", f'=IFERROR(AVERAGEIFS(d_Ship,{FIL}),0)', '0.0', '="Per order line"'),
         ("YoY REVENUE GROWTH", f'=IFERROR({REV(FIL_NY, "d_Year,yoy_Year")}/{REV(FIL_NY, "d_Year,yoy_Year-1")}-1,"n/a")', PCT_COL, '=yoy_Year&" vs "&(yoy_Year-1)')]
# 16 columns B..Q: four tiles of 3 columns + two of 2 columns
tcols = [("B", "D"), ("E", "G"), ("H", "J"), ("K", "M"), ("N", "O"), ("P", "Q")]
for (c1, c2), (lab, fm, nf, sub) in zip(tcols, tiles):
    for r, val, ft, h in ((6, lab, f(bold=True, size=8, color=GREY), 18), (7, fm, f(bold=True, size=20, color=NAVY), 34), (8, sub, f(size=8, color=GREY, italic=True), 16)):
        D.merge_cells(f"{c1}{r}:{c2}{r}")
        D[f"{c1}{r}"] = val; D[f"{c1}{r}"].font = ft; D[f"{c1}{r}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        D.row_dimensions[r].height = h
    D[f"{c1}7"].number_format = nf
    from openpyxl.utils import column_index_from_string as ci
    for r in (6, 7, 8):
        for c in range(ci(c1), ci(c2) + 1):
            D.cell(r, c).fill = fill(LIGHT)
            D.cell(r, c).border = Border(top=Side(style="thick", color=BLUE) if r == 6 else None, left=Side(style="thin", color="FFFFFF") if c == ci(c1) else None,
                                         right=Side(style="thin", color="FFFFFF") if c == ci(c2) else None)
D.row_dimensions[9].height = 10
D["B10"] = '="Showing: Region = "&sel_Region&"  |  Category = "&sel_Category&"  |  Segment = "&sel_Segment&"  |  Year = "&sel_Year&"     (monthly and yearly charts always span all years)"'
D["B10"].font = f(italic=True, size=9, color=GREY)

# charts ---------------------------------------------------------------------------------
cp = lambda sz=900, b=False, col="404040": CharacterProperties(latin=DFont(typeface=AR), sz=sz, b=b, solidFill=col)
def style(ch, ttl, w, h, legend=False, yfmt=None):
    ch.title = ttl
    ch.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=cp(1100, True, NAVY))
    for run in ch.title.tx.rich.p[0].r:
        run.rPr = cp(1100, True, NAVY)
    ch.title.overlay = False
    ch.width, ch.height = w, h
    ch.x_axis.delete = False; ch.y_axis.delete = False
    for ax in (ch.x_axis, ch.y_axis):
        ax.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp()), endParaRPr=cp())])
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=None); ch.y_axis.majorGridlines.spPr.line.solidFill = "E5E7EB"
    if yfmt: ch.y_axis.number_format = yfmt; ch.y_axis.numFmt.sourceLinked = False
    if legend:
        ch.legend.position = "b"; ch.legend.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp()), endParaRPr=cp())])
    else:
        ch.legend = None
    return ch
def color(s, c, line=False, w=28000):
    if line:
        s.graphicalProperties.line.solidFill = c; s.graphicalProperties.line.width = w; s.smooth = False
        s.marker.symbol = "none"
    else:
        s.graphicalProperties.solidFill = c; s.graphicalProperties.line.solidFill = c
def labels(s, fmt):
    s.dLbls = DataLabelList(); s.dLbls.showVal = True; s.dLbls.showSerName = s.dLbls.showCatName = s.dLbls.showLegendKey = False
    s.dLbls.numFmt = fmt; s.dLbls.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp(800)), endParaRPr=cp(800))])

W, H, HALF = 37.6, 8.2, 18.6
ref = lambda col, a, b, hdr=False: Reference(A, min_col=col, min_row=a - (1 if hdr else 0), max_row=b)

m0, m1 = pos["month"]
ch = LineChart(); ch.add_data(ref(2, m0, m1, True), titles_from_data=True); ch.add_data(ref(8, m0, m1, True), titles_from_data=True)
ch.set_categories(ref(1, m0, m1)); style(ch, "Monthly revenue with 3-month moving average", W, H, True, '"$"#,##0')
color(ch.series[0], "9DB7D5", True, 22000); color(ch.series[1], BLUE, True, 34000)
ch.x_axis.number_format = "mmm-yy"; ch.x_axis.numFmt.sourceLinked = False if ch.x_axis.numFmt else None; ch.x_axis.tickLblSkip = 3; ch.x_axis.tickMarkSkip = 3
D.add_chart(ch, "B11")

def bar(title_, col, rng, anchor, c, horiz=False, fmt='"$"#,##0', lbl='"$"#,##0,"K"', cat_col=1, sheet_hdr=True):
    b = BarChart(); b.type = "bar" if horiz else "col"
    b.add_data(Reference(A, min_col=col, min_row=rng[0] - 1, max_row=rng[1]), titles_from_data=True)
    b.set_categories(Reference(A, min_col=cat_col, min_row=rng[0], max_row=rng[1]))
    style(b, title_, HALF, H, False, fmt); color(b.series[0], c); labels(b.series[0], lbl); b.gapWidth = 60
    b.y_axis.scaling.min = 0   # bars must start at zero
    if horiz: b.x_axis.scaling.orientation = "maxMin"; b.y_axis.crosses = "max"
    D.add_chart(b, anchor)
R1, R2, R3, R4 = 28, 45, 62, 79
bar("Revenue by year", 2, pos["year"], f"B{R1}", NAVY)
bar("Revenue by category", 2, pos["cat"], f"J{R1}", BLUE)
bar("Revenue by region", 2, pos["reg"], f"B{R2}", BLUE)
s0, s1, _ = pos["sub"]
bar("Top 10 sub-categories by revenue", 7, (s0, s0 + 9), f"J{R2}", NAVY, horiz=True, cat_col=6)
ch = LineChart(); ch.add_data(Reference(A, min_col=7, min_row=pos["month"][0] + 12 - 1, max_row=pos["month"][1]), titles_from_data=True)
ch.set_categories(Reference(A, min_col=1, min_row=pos["month"][0] + 12, max_row=pos["month"][1]))
style(ch, "Year-over-year revenue growth by month", HALF, H, False, "0%"); color(ch.series[0], TEAL, True, 26000)
ch.x_axis.number_format = "mmm-yy"; ch.x_axis.tickLblSkip = 6; ch.x_axis.tickMarkSkip = 6; ch.x_axis.tickLblPos = "low"
D.add_chart(ch, f"B{R3}")
bar("Seasonality: revenue by calendar month", 2, pos["season"], f"J{R3}", BLUE)
st0, st1, _ = pos["state"]
bar("Top 10 states by revenue", 7, (st0, st1), f"B{R4}", NAVY, horiz=True, cat_col=6)
bar("Revenue by customer segment", 2, pos["seg"], f"J{R4}", BLUE)
D["B96"] = "Source: Superstore Sales Dataset (Kaggle, train.csv). Revenue = Sales column. 'Units' are order lines because the extract has no Quantity field. See Notes sheet."
D["B96"].font = f(italic=True, size=8, color=GREY)
D.page_setup.orientation = "landscape"; D.page_setup.fitToWidth = 1; D.page_setup.fitToHeight = 0
D.sheet_properties.pageSetUpPr.fitToPage = True
D.print_area = "A1:R96"

# ----------------------------------------------------------------------------- 6. Insights (static, from pandas at default filters) and Notes
rev = df.Sales.sum(); yr = df.groupby(df["Order Date"].dt.year).Sales.sum()
cat_s = df.groupby("Category").Sales.sum().sort_values(ascending=False); reg_s = df.groupby("Region").Sales.sum().sort_values(ascending=False)
sub_s = df.groupby("Sub-Category").Sales.sum().sort_values(ascending=False); st_s = df.groupby("State").Sales.sum().sort_values(ascending=False)
mo = df.groupby(df["Order Date"].dt.month).Sales.sum(); seg_s = df.groupby("Segment").Sales.sum().sort_values(ascending=False)
mon = df.groupby(df["Order Date"].dt.to_period("M")).Sales.sum()
q4 = mo.loc[[9, 10, 11, 12]].sum() / rev
cagr = (yr[2018] / yr[2015]) ** (1 / 3) - 1
nov_dec = mo.loc[[11, 12]].sum() / rev
ins = [
    ("Headline", f"Total revenue is ${rev:,.0f} from {df['Order ID'].nunique():,} orders and {len(df):,} order lines ({df['Customer ID'].nunique()} customers). Average order value is ${rev/df['Order ID'].nunique():,.0f}."),
    ("Growth", f"Revenue fell {abs(yr[2016]/yr[2015]-1):.1%} in 2016, then grew {yr[2017]/yr[2016]-1:+.1%} in 2017 and {yr[2018]/yr[2017]-1:+.1%} in 2018 (${yr[2018]:,.0f}). 2015-2018 CAGR is {cagr:.1%}."),
    ("Category mix", f"{cat_s.index[0]} leads with {cat_s.iloc[0]/rev:.1%} of revenue, then {cat_s.index[1]} ({cat_s.iloc[1]/rev:.1%}) and {cat_s.index[2]} ({cat_s.iloc[2]/rev:.1%}). Office Supplies has the most order lines but the lowest revenue per line."),
    ("Top sub-categories", f"{sub_s.index[0]} (${sub_s.iloc[0]:,.0f}), {sub_s.index[1]} (${sub_s.iloc[1]:,.0f}) and {sub_s.index[2]} (${sub_s.iloc[2]:,.0f}) are the top three; the top 5 account for {sub_s.iloc[:5].sum()/rev:.0%} of revenue."),
    ("Regions", f"{reg_s.index[0]} is the largest region ({reg_s.iloc[0]/rev:.1%}) and {reg_s.index[-1]} the smallest ({reg_s.iloc[-1]/rev:.1%}). The top-2 regions ({reg_s.index[0]} + {reg_s.index[1]}) generate {reg_s.iloc[:2].sum()/rev:.0%} of revenue."),
    ("States", f"{st_s.index[0]} (${st_s.iloc[0]:,.0f}, {st_s.iloc[0]/rev:.1%}) and {st_s.index[1]} (${st_s.iloc[1]:,.0f}) lead; the top 5 states generate {st_s.iloc[:5].sum()/rev:.0%} of revenue."),
    ("Seasonality", f"Revenue is strongly seasonal: November and December together are {nov_dec:.0%} of all revenue and Sep-Dec is {q4:.0%}. The peak month is {mon.idxmax().strftime('%b %Y')} (${mon.max():,.0f}); the weakest is {mon.idxmin().strftime('%b %Y')} (${mon.min():,.0f})."),
    ("Segments", f"{seg_s.index[0]} contributes {seg_s.iloc[0]/rev:.0%} of revenue, {seg_s.index[1]} {seg_s.iloc[1]/rev:.0%}, {seg_s.index[2]} {seg_s.iloc[2]/rev:.0%}."),
    ("Recommendation", "Plan inventory and ad spend around the Sep-Dec peak, defend the top sub-categories, and investigate why the weakest region and bottom sub-categories lag before adding promotions there."),
    ("Caveat", "Findings describe Revenue (Sales) only. The data has no Profit, Discount or Quantity fields, so profitability and true unit volumes cannot be assessed."),
]
ws_i.sheet_view.showGridLines = False
ws_i["A1"] = "Key Insights (full dataset, no filters)"; ws_i["A1"].font = f(bold=True, size=16, color=NAVY)
ws_i["A2"] = "Static text computed from the cleaned data; the Dashboard shows the live, filterable figures."; ws_i["A2"].font = f(italic=True, size=9, color=GREY)
for k, (a, b) in enumerate(ins, start=4):
    ws_i.cell(k, 1, a).font = f(bold=True, size=10, color=WHITE); ws_i.cell(k, 1).fill = fill(BLUE)
    ws_i.cell(k, 1).alignment = Alignment(vertical="center", horizontal="center")
    c = ws_i.cell(k, 2, b); c.font = f(size=10); c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = box
    ws_i.row_dimensions[k].height = 44
ws_i.column_dimensions["A"].width = 20; ws_i.column_dimensions["B"].width = 120
ws_i.sheet_properties.tabColor = TEAL

N_ = ws_n
N_.sheet_view.showGridLines = False
N_["A1"] = "Notes: methodology, data cleaning log and how to use"; N_["A1"].font = f(bold=True, size=16, color=NAVY)
N_["A3"] = "Data cleaning log"; N_["A3"].font = f(bold=True, size=12, color=NAVY)
header(N_, 4, ["Step", "Finding", "Action taken", "Rows affected"])
for k, row in enumerate(log, start=5):
    for c, v in enumerate(row, 1):
        x = N_.cell(k, c, v); x.font = f(size=10); x.alignment = Alignment(wrap_text=True, vertical="top"); x.border = box
    N_.row_dimensions[k].height = 44
r = 5 + len(log) + 1
N_.cell(r, 1, "Metric definitions & workbook guide").font = f(bold=True, size=12, color=NAVY)
defs = [("Revenue", "Sum of the Sales column (the only monetary field in the extract)."),
        ("Orders", "Distinct Order IDs. Helper columns 'New Order' / 'New Order (Cat)' on the Data sheet flag the first line of each order (or of each order-category pair when a Category filter is on), so a distinct count can be done with SUMIFS."),
        ("Order lines (units proxy)", "Row count. The dataset has no Quantity column, so true units sold cannot be computed."),
        ("Avg order value", "Revenue / distinct orders."),
        ("MoM % / YoY %", "Revenue vs. the previous month / the same month one year earlier (Analysis section 1). YoY KPI tile compares the selected year with the prior year (latest year when Year = All)."),
        ("Filters", "Dashboard dropdowns (Region, Category, Segment, Year) drive criteria cells on the Analysis sheet; every KPI, table and chart recalculates. They behave like slicers but are data-validation lists."),
        ("Pivot tables / slicers", "The summary tables are formula-driven pivot-style tables (SUMIFS over named ranges). To add native Excel PivotTables and Slicers: click any cell in the Data sheet > Insert > PivotTable (source: tblSales), then PivotTable Analyze > Insert Slicer. Step-by-step in the project README."),
        ("Data sheet", "Columns A-R are the cleaned source fields; S-X are formula helpers (Year, Month, Month No, Ship Days, order flags). Sorted by Order Date, Order ID, Category."),
        ("Source", "Superstore Sales Dataset (Kaggle), train.csv: 9,800 rows, Jan 2015 - Dec 2018, U.S. only.")]
for k, (a, b) in enumerate(defs, start=r + 1):
    N_.cell(k, 1, a).font = f(bold=True, size=10); c = N_.cell(k, 2, b); c.font = f(size=10); c.alignment = Alignment(wrap_text=True, vertical="top")
    N_.merge_cells(start_row=k, start_column=2, end_row=k, end_column=4); N_.row_dimensions[k].height = 44
    N_.cell(k, 1).alignment = Alignment(vertical="top", wrap_text=True)
for i, w in enumerate([24, 60, 60, 14], 1):
    N_.column_dimensions[L(i)].width = w
N_.sheet_properties.tabColor = GREY

OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
df.to_csv(ROOT / "data" / "superstore_cleaned.csv", index=False, date_format="%Y-%m-%d")
print("saved", OUT)
