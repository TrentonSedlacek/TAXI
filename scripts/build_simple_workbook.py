"""Build analysis/Theta Xi GPA and Recruitment.xlsx, the short version to share with the chapter.

Run from the repo root:  python3 scripts/build_simple_workbook.py
"""
import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.properties import PageSetupProperties

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "analysis" / "Theta Xi GPA and Recruitment.xlsx"
US = "Theta Xi"
FIRST = "Fall 2018"
SPARE = 4          # empty rows left for future semesters / years

F = Font(name="Calibri", size=11)
B = Font(name="Calibri", size=11, bold=True)
T = Font(name="Calibri", size=16, bold=True)
SUB = Font(name="Calibri", size=10, color="595959")
HF = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="1F4E78")
USFILL = PatternFill("solid", fgColor="DDEBF7")
LINE = Side(style="thin", color="BFBFBF")
BOX = Border(top=LINE, bottom=LINE, left=LINE, right=LINE)
WRAPC = Alignment(horizontal="center", vertical="center", wrap_text=True)


def key(t):
    s, y = t.split()
    return int(y) * 2 + (s == "Fall")


def num(v):
    if v in (None, ""):
        return None
    f = float(v)
    return int(f) if f.is_integer() else f


with open(ROOT / "data" / "chapter_semester_data.csv", newline="") as fh:
    rows = [r for r in csv.DictReader(fh)
            if r["council"] == "IFC" and key(r["semester"]) >= key(FIRST) and r["source_file"].startswith("source-data")]
with open(ROOT / "data" / "campus_stats.csv", newline="") as fh:
    campus = [r for r in csv.DictReader(fh) if key(r["semester"]) >= key(FIRST)]
rows.sort(key=lambda r: (key(r["semester"]), r["chapter"]))
terms = sorted({r["semester"] for r in rows}, key=key)
latest = terms[-1]
last_fall_year = max(int(t.split()[1]) for t in terms if t.startswith("Fall"))

wb = Workbook()
sem_ws = wb.active
sem_ws.title = "By Semester"
year_ws = wb.create_sheet("By Year")
house_ws = wb.create_sheet("IFC Houses")
data_ws = wb.create_sheet("UNL Data")
avg_ws = wb.create_sheet("UNL Averages")


def cell(ws, r, c, v, font=F, fmt=None, fill=None, border=True):
    x = ws.cell(r, c, v)
    x.font = font
    if fmt:
        x.number_format = fmt
    if fill:
        x.fill = fill
    if border:
        x.border = BOX
    return x


def header(ws, r, labels, height=32):
    for i, text in enumerate(labels, start=1):
        x = cell(ws, r, i, text, HF, fill=HFILL)
        x.alignment = WRAPC
    ws.row_dimensions[r].height = height


def title(ws, text, sub):
    cell(ws, 1, 1, text, T, border=False)
    cell(ws, 2, 1, sub, SUB, border=False)


# ------------------------------------------------------------------ UNL Data
header(data_ws, 1, ["Semester", "House", "Members", "New Members", "GPA"], height=20)
for i, r in enumerate(rows, start=2):
    cell(data_ws, i, 1, r["semester"])
    cell(data_ws, i, 2, r["chapter"])
    cell(data_ws, i, 3, num(r["total_members"]), fmt="0")
    cell(data_ws, i, 4, num(r["new_members"]), fmt="0")
    cell(data_ws, i, 5, num(r["chapter_gpa"]), fmt="0.000")
data_ws.freeze_panes = "A2"
data_ws.auto_filter.ref = f"A1:E{len(rows) + 1}"
for col, w in zip("ABCDE", [13, 22, 10, 13, 8]):
    data_ws.column_dimensions[col].width = w
data_ws.column_dimensions["G"].width = 70
notes = ["IFC houses only, straight from UNL's semester report cards and scorecards.",
         "Fall 2023 has no scorecard, so it comes from UNL's grade report (no new-member counts).",
         "Spring 2020 has no GPAs (pass/no-pass semester). Spring 2020 and Spring 2022 have no new-member counts.",
         "To add a semester, paste the IFC rows at the bottom."]
for k, text in enumerate(notes):
    cell(data_ws, 1 + k, 7, text, SUB, border=False)

N = 1500
SEM, HOUSE, MEM, NM, GPA = (f"'UNL Data'!${c}$2:${c}${N}" for c in "ABCDE")

# ------------------------------------------------------------------ UNL Averages
header(avg_ws, 1, ["Semester", "All UNL Men", "All Greek Men", "All UNL Students"], height=20)
for i, r in enumerate(campus, start=2):
    cell(avg_ws, i, 1, r["semester"])
    cell(avg_ws, i, 2, num(r["All Male"]), fmt="0.000")
    cell(avg_ws, i, 3, num(r["Greek Male"]), fmt="0.000")
    cell(avg_ws, i, 4, num(r["All UNL"]), fmt="0.000")
for col, w in zip("ABCD", [13, 13, 14, 16]):
    avg_ws.column_dimensions[col].width = w
cell(avg_ws, 1, 6, "From UNL's All-Community Grade Report each semester.", SUB, border=False)
avg_ws.freeze_panes = "A2"


def us_value(rng, sem_expr):
    return (f'IF(COUNTIFS({SEM},{sem_expr},{HOUSE},"{US}",{rng},"<>")=0,"",'
            f'SUMIFS({rng},{SEM},{sem_expr},{HOUSE},"{US}"))')


def avg_lookup(col, sem_cell):
    return (f"IFERROR(IF(INDEX('UNL Averages'!${col}:${col},MATCH({sem_cell},'UNL Averages'!$A:$A,0))=\"\",\"\","
            f"INDEX('UNL Averages'!${col}:${col},MATCH({sem_cell},'UNL Averages'!$A:$A,0))),\"\")")


# ------------------------------------------------------------------ By Semester
title(sem_ws, "Theta Xi vs the IFC", "Every semester since Fall 2018, from UNL's report cards. "
                                     "IFC Avg = the average IFC house that semester.")
header(sem_ws, 4, ["Semester", "TX GPA", "All Greek Men GPA", "TX vs Greek Men", "All UNL Men GPA",
                   "TX GPA Rank in IFC", "TX Members", "IFC Avg Members", "TX New Members",
                   "IFC Avg New Members", "TX New Member Rank in IFC"], height=48)
sem_ws.freeze_panes = "B5"
for i in range(len(terms) + SPARE):
    r = 5 + i
    a = f"$A{r}"
    cell(sem_ws, r, 1, terms[i] if i < len(terms) else None, B)
    cell(sem_ws, r, 2, f'=IF({a}="","",{us_value(GPA, a)})', fmt="0.000")
    cell(sem_ws, r, 3, f'=IF({a}="","",{avg_lookup("C", a)})', fmt="0.000")
    cell(sem_ws, r, 4, f'=IF(OR(B{r}="",C{r}=""),"",B{r}-C{r})', fmt="+0.000;-0.000;0.000")
    cell(sem_ws, r, 5, f'=IF({a}="","",{avg_lookup("B", a)})', fmt="0.000")
    cell(sem_ws, r, 6, f'=IF(B{r}="","",COUNTIFS({SEM},{a},{GPA},">"&B{r})+1&" of "&COUNTIFS({SEM},{a},{GPA},">0"))')
    cell(sem_ws, r, 7, f'=IF({a}="","",{us_value(MEM, a)})', fmt="0")
    cell(sem_ws, r, 8, f'=IF({a}="","",IFERROR(AVERAGEIFS({MEM},{SEM},{a}),""))', fmt="0")
    cell(sem_ws, r, 9, f'=IF({a}="","",{us_value(NM, a)})', fmt="0")
    cell(sem_ws, r, 10, f'=IF({a}="","",IFERROR(AVERAGEIFS({NM},{SEM},{a}),""))', fmt="0.0")
    cell(sem_ws, r, 11, f'=IF(I{r}="","",COUNTIFS({SEM},{a},{NM},">"&I{r})+1&" of "&COUNTIFS({SEM},{a},{NM},">=0"))')
    for c in range(6, 12, 5):
        sem_ws.cell(r, c).alignment = Alignment(horizontal="right")
last = 4 + len(terms) + SPARE
sem_ws.conditional_formatting.add(f"D5:D{last}", CellIsRule(operator="lessThan", formula=["0"],
                                                            font=Font(name="Calibri", color="C00000")))
sem_ws.conditional_formatting.add(f"D5:D{last}", CellIsRule(operator="greaterThan", formula=["0"],
                                                            font=Font(name="Calibri", color="00B050")))
for col, w in zip("ABCDEFGHIJK", [13, 9, 11, 11, 11, 11, 10, 10, 10, 11, 12]):
    sem_ws.column_dimensions[col].width = w
known = 4 + len(terms)
cats = Reference(sem_ws, min_col=1, min_row=5, max_row=known)
c1 = BarChart()
c1.title, c1.height, c1.width = "GPA: Theta Xi vs All Greek Men", 7.5, 17
c1.y_axis.scaling.min, c1.y_axis.scaling.max = 2.6, 3.6
c1.y_axis.number_format = "0.0"
for col in (2, 3):
    c1.add_data(Reference(sem_ws, min_col=col, min_row=4, max_row=known), titles_from_data=True)
c1.set_categories(cats)
sem_ws.add_chart(c1, f"A{last + 3}")
c2 = BarChart()
c2.title, c2.height, c2.width = "New Members: Theta Xi vs the Average IFC House", 7.5, 17
for col in (9, 10):
    c2.add_data(Reference(sem_ws, min_col=col, min_row=4, max_row=known), titles_from_data=True)
c2.set_categories(cats)
sem_ws.add_chart(c2, f"G{last + 3}")
cell(sem_ws, last + 1, 1, "Blank = UNL didn't report it that semester.", SUB, border=False)

# ------------------------------------------------------------------ By Year
title(year_ws, "Recruitment by Year", "Year = spring + fall recruitment, the same exec board "
                                      "(elected in November, trained in December).")
header(year_ws, 4, ["Year", "TX New Members", "IFC Avg New Members", "TX vs Avg House",
                    "TX Members (Fall)", "Change From Last Fall", "IFC Avg Members (Fall)"], height=48)
year_ws.freeze_panes = "B5"
first_year = int(FIRST.split()[1])
years = list(range(first_year, last_fall_year + 1 + SPARE))
for i, y in enumerate(years):
    r = 5 + i
    sp, fa = f'"Spring "&$A{r}', f'"Fall "&$A{r}'
    cell(year_ws, r, 1, y, B)
    cell(year_ws, r, 2, f'=IF(OR(COUNTIFS({SEM},{sp},{HOUSE},"{US}",{NM},"<>")=0,'
                        f'COUNTIFS({SEM},{fa},{HOUSE},"{US}",{NM},"<>")=0),"",'
                        f'SUMIFS({NM},{SEM},{sp},{HOUSE},"{US}")+SUMIFS({NM},{SEM},{fa},{HOUSE},"{US}"))', fmt="0")
    cell(year_ws, r, 3, f'=IFERROR(AVERAGEIFS({NM},{SEM},{sp})+AVERAGEIFS({NM},{SEM},{fa}),"")', fmt="0.0")
    cell(year_ws, r, 4, f'=IF(OR(B{r}="",C{r}=""),"",B{r}/C{r})', fmt="0%")
    cell(year_ws, r, 5, f'={us_value(MEM, fa)}', fmt="0")
    cell(year_ws, r, 6, "" if i == 0 else f'=IF(OR(E{r}="",E{r - 1}=""),"",E{r}-E{r - 1})', fmt="+0;-0;0")
    cell(year_ws, r, 7, f'=IFERROR(AVERAGEIFS({MEM},{SEM},{fa}),"")', fmt="0")
ylast = 4 + len(years)
year_ws.conditional_formatting.add(f"F5:F{ylast}", CellIsRule(operator="lessThan", formula=["0"],
                                                              font=Font(name="Calibri", color="C00000")))
for col, w in zip("ABCDEFG", [8, 11, 12, 11, 11, 11, 12]):
    year_ws.column_dimensions[col].width = w
cell(year_ws, ylast + 1, 1, "TX vs Avg House: 100% = recruited as many as the average IFC house. "
                            "Blank = UNL left new members off one of that year's reports.", SUB, border=False)
yknown = 4 + (last_fall_year - first_year + 1)
ycats = Reference(year_ws, min_col=1, min_row=5, max_row=yknown)
c3 = BarChart()
c3.title, c3.height, c3.width = "New Members per Year", 7.5, 14
for col in (2, 3):
    c3.add_data(Reference(year_ws, min_col=col, min_row=4, max_row=yknown), titles_from_data=True)
c3.set_categories(ycats)
year_ws.add_chart(c3, f"A{ylast + 3}")
c4 = BarChart()
c4.title, c4.height, c4.width = "Members Each Fall", 7.5, 14
for col in (5, 7):
    c4.add_data(Reference(year_ws, min_col=col, min_row=4, max_row=yknown), titles_from_data=True)
c4.set_categories(ycats)
year_ws.add_chart(c4, f"F{ylast + 3}")

# ------------------------------------------------------------------ IFC Houses
title(house_ws, "Every IFC House", "Change the two blue cells to look at a different semester or year.")
cell(house_ws, 3, 1, "GPA and members for", B, border=False)
cell(house_ws, 3, 3, latest, Font(name="Calibri", size=11, bold=True, color="0000FF"))
cell(house_ws, 4, 1, "New members for year", B, border=False)
cell(house_ws, 4, 3, last_fall_year, Font(name="Calibri", size=11, bold=True, color="0000FF"))
header(house_ws, 6, ["House", "", "GPA", "GPA Rank", "Members", "New Members (Spring + Fall)", "New Member Rank"],
       height=48)
house_ws.merge_cells("A6:B6")
current = [r for r in rows if r["semester"] == latest]
current.sort(key=lambda r: -(num(r["chapter_gpa"]) or 0))
h0, h1 = 7, 7 + len(current) - 1
sp, fa = '"Spring "&$C$4', '"Fall "&$C$4'
for i, r in enumerate(current):
    rr = h0 + i
    fill = USFILL if r["chapter"] == US else None
    font = B if r["chapter"] == US else F
    house_ws.merge_cells(f"A{rr}:B{rr}")
    cell(house_ws, rr, 1, r["chapter"], font, fill=fill)
    cell(house_ws, rr, 2, None, font, fill=fill)
    a = f"$A{rr}"
    cell(house_ws, rr, 3, f'=IF(COUNTIFS({SEM},$C$3,{HOUSE},{a},{GPA},"<>")=0,"",SUMIFS({GPA},{SEM},$C$3,{HOUSE},{a}))',
         font, "0.000", fill)
    cell(house_ws, rr, 4, f'=IF(C{rr}="","",COUNTIF($C${h0}:$C${h1},">"&C{rr})+1)', font, "0", fill)
    cell(house_ws, rr, 5, f'=IF(COUNTIFS({SEM},$C$3,{HOUSE},{a},{MEM},"<>")=0,"",SUMIFS({MEM},{SEM},$C$3,{HOUSE},{a}))',
         font, "0", fill)
    cell(house_ws, rr, 6, f'=IF(OR(COUNTIFS({SEM},{sp},{HOUSE},{a},{NM},"<>")=0,COUNTIFS({SEM},{fa},{HOUSE},{a},{NM},"<>")=0),"",'
                          f'SUMIFS({NM},{SEM},{sp},{HOUSE},{a})+SUMIFS({NM},{SEM},{fa},{HOUSE},{a}))', font, "0", fill)
    cell(house_ws, rr, 7, f'=IF(F{rr}="","",COUNTIF($F${h0}:$F${h1},">"&F{rr})+1)', font, "0", fill)
for col, w in zip("ABCDEFG", [14, 12, 9, 9, 10, 14, 11]):
    house_ws.column_dimensions[col].width = w
cell(house_ws, h1 + 2, 1, f"Houses listed are the IFC houses on the {latest} report, sorted by {latest} GPA.",
     SUB, border=False)

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = ws.title in ("UNL Data", "UNL Averages")
    for ch in ws._charts:
        ch.legend.position = "b"
        ch.x_axis.delete = False
        ch.y_axis.delete = False
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
wb.save(OUT)
print("saved", OUT, len(rows), "rows", terms[0], "to", latest)
