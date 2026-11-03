"""
build_template.py — builds the Paycheck Planner budget template (.xlsx), version 2.

Run:
    python tools/build_template.py                 # every theme x every style in config/product.json
    python tools/build_template.py sage            # just the "sage" theme (all styles)
    python tools/build_template.py sage bold       # just sage, bold style
    python tools/build_template.py --push          # build, then update the master Google Sheet

Tabs buyers see:  Setup · Spending · Dashboard · Bills · Goals · Accounts · Debts · Reports
Hidden tab:       Engine  (tax tables, budget presets, chart data)

Everything you'd normally change (name, colors, groups, categories, presets, tax
brackets, inflation) lives in config/product.json, NOT in this file.

Google Sheets rule we follow: every dropdown either has its choices typed right
into it ("Weekly,Biweekly,...") or reads a list on the SAME tab. Dropdowns that
read another tab can disappear when Google converts the file.
"""
import json
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import DoughnutChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
CONFIG = json.loads((ROOT / "config" / "product.json").read_text())
OUT_DIR = ROOT / "output"

# ------------------------------------------------------------------ shared look
FONT = "Arial"
MONEY = '$#,##0.00;-$#,##0.00;"-"'
MONEY0 = '$#,##0;-$#,##0;"$0"'
PCT = '0%;-0%;"-"'
PCT1 = '0.0%;-0.0%;"-"'
DATE_FMT = "mmm d, yyyy"
INPUT_FILL = PatternFill("solid", fgColor="FFF4C2")   # yellow = you type here
WHITE = PatternFill("solid", fgColor="FFFFFF")
LINE = Side(style="thin", color="DADDE1")
THICK_WHITE = Side(style="medium", color="FFFFFF")
GRID = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
UNDER = Border(bottom=LINE)
INK = "1F2328"          # main text
MUTED = "6B7280"        # hints
RED, RED_BG = "B42318", "FDE8E7"
AMBER_BG = "FFF4D6"
GREEN, GREEN_BG = "1E7B4F", "E3F4EA"
BLUE, BLUE_BG = "2B4C7E", "E6EEF8"   # neutral status (not good or bad yet)

# ------------------------------------------------------------------ fixed cell map
# Other tabs point at these Setup cells, so keep them in one place.
G = CONFIG["group_slots"]
C_SLOTS = CONFIG["category_slots"]
S = dict(freq="$C$8", first_pay="$C$9", pay="$C$10", monthly="$C$11", pretax="$C$12", first="$C$13",
         per_year="$K$8", pay_days="$C$14", weekend="$C$15", early_label="$C$16",
         tax_mode="$C$18", status="$C$19", state="$C$20", fica="$C$21", manual="$C$22",
         infl_mode="$C$25", infl_auto="$C$26", infl_manual="$C$27", infl_used="$C$28",
         take_y="$C$39", take_m="$D$39", eff="$C$40", cost1="$C$41", per_check="$C$42",
         plan="$C$46")
# Twice a month: which days? (only used when you're paid Semi-monthly)
DAYS_FROM_FIRST = "From first payday"
PAY_DAY_CHOICES = {DAYS_FROM_FIRST: None, "1st & 15th": (1, 15), "15th & 30th": (15, 30), "15th & last day": (15, 31)}
# Payday lands on a Saturday or Sunday? (only used for Semi-monthly and Monthly pay)
WEEKEND_BEFORE, WEEKEND_AFTER, WEEKEND_SAME = "Workday before", "Workday after", "No change"
OLD_BEFORE = "Friday before"        # v2.11 name; copies imported from v2.11 may still have it
EARLY_CHOICES = ["No, on payday", "1 day early", "2 days early", "3 days early", "4 days early", "5 days early"]
EARLY_DAYS = "Engine!$B$18"           # how many days early the money lands (0-5), worked out from Setup C16
GROUP_FIRST = 48                       # group rows 48-53
GROUP_LAST = GROUP_FIRST + G - 1
GROUP_TOTAL = GROUP_LAST + 1
CAT_FIRST = 60                         # category rows 60-89
CAT_LAST = CAT_FIRST + C_SLOTS - 1
SP_FIRST, SP_LAST = 5, 5 + CONFIG["spending_rows"] - 1
MI_FIRST, MI_LAST = 5, 5 + CONFIG["money_in_rows"] - 1      # "Money in" log, right side of Spending
PAYDAY_FIRST = 60                                           # Engine: list of every payday
PAYDAY_YEARS = (CONFIG["calendar_months"] + 11) // 12 + 1   # 4 years: covers Jan 2026-Dec 2028 even if you start late
PAYDAY_LAST = PAYDAY_FIRST + PAYDAY_YEARS * CONFIG["paycheck_rows"] - 1
KIND_BILLS, KIND_FLEX, KIND_SAVE = CONFIG["budget_kinds"]
TAX_MANUAL_LABEL = "Use my own rate"
INFL_MANUAL_LABEL = "My own rate"


def setup(key):
    return f"Setup!{S[key]}"


def rng(sheet, col, first, last):
    return f"{sheet}!${col}${first}:${col}${last}"


SP_AMT = rng("Spending", "E", SP_FIRST, SP_LAST)
SP_DATE = rng("Spending", "B", SP_FIRST, SP_LAST)
SP_CAT = rng("Spending", "D", SP_FIRST, SP_LAST)
SP_OK = rng("Spending", "G", SP_FIRST, SP_LAST)       # OK'd? (reason you went over)
SP_COVER = rng("Spending", "H", SP_FIRST, SP_LAST)    # Cover with (extra money in / last month's leftover)
SP_GROUP = rng("Spending", "J", SP_FIRST, SP_LAST)    # hidden: the category's group
SP_KIND = rng("Spending", "K", SP_FIRST, SP_LAST)     # hidden: bills / flexible / saving
SP_KEY = rng("Spending", "L", SP_FIRST, SP_LAST)      # hidden: which budget the row checks (category or group)
SP_OKCAT = rng("Spending", "M", SP_FIRST, SP_LAST)    # hidden: "Groceries|202610" when OK'd
SP_OKGRP = rng("Spending", "N", SP_FIRST, SP_LAST)    # hidden: "Fun|202610" when OK'd
OK_REASONS = ["Planned it", "Emergency", "Sale / good deal", "Bill was higher", "Special day / gift", "Other"]
COVER_LEFTOVER = "Last month's leftover"
COVER_CHOICES = ["Extra money in", COVER_LEFTOVER]
G_NAME = rng("Setup", "B", GROUP_FIRST, GROUP_LAST)
G_KIND = rng("Setup", "C", GROUP_FIRST, GROUP_LAST)
G_MONTH = rng("Setup", "F", GROUP_FIRST, GROUP_LAST)
C_NAME = rng("Setup", "B", CAT_FIRST, CAT_LAST)
C_GROUP = rng("Setup", "C", CAT_FIRST, CAT_LAST)
C_GOAL = rng("Setup", "D", CAT_FIRST, CAT_LAST)
C_RANK = rng("Setup", "J", CAT_FIRST, CAT_LAST)
C_KIND = rng("Setup", "K", CAT_FIRST, CAT_LAST)
C_ALPHA = rng("Setup", "L", CAT_FIRST, CAT_LAST)    # hidden: each category's place in A-to-Z order (for dropdowns)
G_ALPHA_LIST = f"$M${GROUP_FIRST}:$M${GROUP_FIRST + CONFIG['group_slots'] - 1}"   # Setup: group names A to Z
MI_DATE = rng("Spending", "Q", MI_FIRST, MI_LAST)
MI_AMT = rng("Spending", "S", MI_FIRST, MI_LAST)
MI_FOR = rng("Spending", "T", MI_FIRST, MI_LAST)      # extra money "for" a category or group
MI_FORGRP = rng("Spending", "V", MI_FIRST, MI_LAST)   # hidden: the group that extra money lands in
PAYDAYS = rng("Engine", "B", PAYDAY_FIRST, PAYDAY_LAST)
HOL_FIRST = PAYDAY_FIRST                                   # Engine E:G, same rows as the payday list
HOL_LAST = HOL_FIRST + CONFIG["bank_holidays"]["years"] * len(CONFIG["bank_holidays"]["list"]) - 1
HOLIDAYS = rng("Engine", "G", HOL_FIRST, HOL_LAST)          # every bank holiday (the day banks are closed)
FIRST_NET = "Engine!$B$20"          # take-home of your first paycheck (smaller if you started mid pay period)


def payday(n):
    """Formula for the OFFICIAL date of paycheck number n (before any "bank pays early" days).
    Weekly / Biweekly: every 7 / 14 days from your first payday.
    Monthly: the same date every month.  Semi-monthly: twice a month, either 15 days apart
    (counted from your first payday) or on fixed days like the 15th & 30th.
    Semi-monthly and Monthly paydays that land on a weekend move to the Friday before / Monday after."""
    freq, first, days, wk = setup("freq"), setup("first"), setup("pay_days"), setup("weekend")
    base = f"DATE(YEAR({first}),MONTH({first}),1)"          # the 1st of your first payday's month

    def fixed(d1, d2, slot0):
        """Paycheck n on fixed days d1/d2. slot0 = which slot your first payday was (0 = d1, 1 = d2, 2 = next month's d1)."""
        s_ = f"({slot0}+{n}-1)"                              # slot number counting from your first payday's month
        mo = f"INT({s_}/2)"                                   # months after that month
        day = f"IF(MOD({s_},2)=0,{d1},{d2})"
        # MIN(...) keeps "the 30th" or "the 31st" inside short months (Feb 30 -> Feb 28)
        return f"MIN(DATE(YEAR({base}),MONTH({base})+{mo},{day}),EOMONTH({base},{mo}))"

    d = f"DAY({first})"
    semi = (f'IF({days}="1st & 15th",{fixed(1, 15, f"IF({d}>=24,2,IF({d}<=7,0,1))")},'
            f'IF({days}="15th & 30th",{fixed(15, 30, f"IF({d}<=22,0,1)")},'
            f'IF({days}="15th & last day",{fixed(15, 31, f"IF({d}<=22,0,1)")},'
            f'IF(MOD({n},2)=1,EDATE({first},({n}-1)/2),EDATE({first}+15,({n}-2)/2)))))')
    on_paper = f'IF({freq}="Monthly",EDATE({first},{n}-1),{semi})'
    return f'IF({freq}="Weekly",{first}+({n}-1)*7,IF({freq}="Biweekly",{first}+({n}-1)*14,{on_paper}))'


def weekend_moved(cell):
    """The payday in `cell`, moved off a weekend or bank holiday when Setup C15 says so.
    Semi-monthly / Monthly: moved off weekends AND bank holidays.
    Weekly / Biweekly: only moved off bank holidays (they always land on the same weekday you picked).
    WORKDAY(date, -1, holidays) = the business day before; WORKDAY(date, 1, holidays) = the business day after.
    WORKDAY(cell+1, -1, ...) = "this day if it's a business day, otherwise the one before"."""
    freq, wk = setup("freq"), setup("weekend")
    before = f'OR({wk}="{WEEKEND_BEFORE}",{wk}="{OLD_BEFORE}")'
    every_n_days = (f'IF(COUNTIF({HOLIDAYS},{cell})=0,{cell},'
                    f'IF({before},WORKDAY({cell},-1,{HOLIDAYS}),WORKDAY({cell},1,{HOLIDAYS})))')
    set_days = f'IF({before},WORKDAY({cell}+1,-1,{HOLIDAYS}),WORKDAY({cell}-1,1,{HOLIDAYS}))'
    return (f'(IF(OR({wk}="{WEEKEND_SAME}",{wk}=""),{cell},'
            f'IF(OR({freq}="Weekly",{freq}="Biweekly"),{every_n_days},{set_days})))')


def holiday_date(rule, y):
    """Formula for one holiday in year cell `y`. rule = [name, kind, month, ...] from config.
    nth: the nth weekday (1 = Monday ... 7 = Sunday) of the month.  last: the last that weekday of the month.
    A holiday on Sunday is observed Monday. On Saturday, banks stay open Friday (Federal Reserve rule)."""
    kind, m = rule[1], rule[2]
    if kind == "fixed":
        d = f"DATE({y},{m},{rule[3]})"
    elif kind == "nth":
        wd, n = rule[3], rule[4]
        d = f"(DATE({y},{m},1)+MOD({wd}-WEEKDAY(DATE({y},{m},1),2),7)+{7 * (n - 1)})"
    else:   # last
        wd = rule[3]
        last = f"EOMONTH(DATE({y},{m},1),0)"
        d = f"({last}-MOD(WEEKDAY({last},2)-{wd},7))"
    return f"{d}+IF(WEEKDAY({d},2)=7,1,0)"


def paydays_in(start, end_exclusive):
    """How many paydays land between two dates (uses the Engine payday list)."""
    return f'COUNTIFS({PAYDAYS},">="&{start},{PAYDAYS},"<"&{end_exclusive})'


def cal_start():
    """First month of every month list (config "calendar_start", e.g. Jan 2026).
    Every list (Dashboard picker, Reports, Goals scorecard, net worth log) runs
    "calendar_months" months from here: Jan 2026 to Dec 2028."""
    y, m = CONFIG["calendar_start"].split("-")
    return f"DATE({int(y)},{int(m)},1)"


CAL_MONTHS = CONFIG["calendar_months"]          # 36
CAL_YEARS = (CAL_MONTHS + 11) // 12             # 3 (2026, 2027, 2028)
_cy, _cm = (int(x) for x in CONFIG["calendar_start"].split("-"))
_ly, _lm = _cy + (_cm - 1 + CAL_MONTHS - 1) // 12, (_cm - 1 + CAL_MONTHS - 1) % 12 + 1
CAL_LABEL = f"{date(_cy, _cm, 1):%b %Y} to {date(_ly, _lm, 1):%b %Y}"   # "Jan 2026 to Dec 2028"
THIS_MONTH = "DATE(YEAR(TODAY()),MONTH(TODAY()),1)"


def year_start(k):
    """Formula for Jan 1 of calendar year number k (0 = the first year)."""
    return f"DATE(YEAR({cal_start()})+{k},1,1)"


def expected_pay(start, end_exclusive):
    """Take-home you expect between two dates: paydays x take-home per paycheck,
    minus the difference if your smaller FIRST paycheck lands in this window (Setup C13)."""
    first, check = f'N(INDEX({PAYDAYS},1))', setup("per_check")      # the day your first check LANDS (Engine list)
    return (f'({paydays_in(start, end_exclusive)}*{check}'
            f'-IF(AND({first}>={start},{first}<{end_exclusive}),{check}-{FIRST_NET},0))')


def received_in(start, end_exclusive):
    """Money actually logged as received (Money in log) between two dates."""
    return f'SUMIFS({MI_AMT},{MI_DATE},">="&{start},{MI_DATE},"<"&{end_exclusive})'


def between(start, end_exclusive):
    """SUMIFS date filter: on or after start, before end."""
    return f'{SP_DATE},">="&{start},{SP_DATE},"<"&{end_exclusive}'


# ------------------------------------------------------------------ style helpers
class Look:
    """Colors for one theme + style combination."""

    def __init__(self, theme, style):
        self.style = style
        self.primary, self.accent, self.soft = theme["primary"], theme["accent"], theme["soft"]
        self.group_colors = CONFIG["group_colors"]

    def group_color(self, i):
        return self.group_colors[i % len(self.group_colors)] if self.style == "bold" else self.primary


def font(size=11, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def put(ws, ref, value=None, f=None, fill=None, fmt=None, align=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = f or font()
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    elif value is not None and not c.alignment.vertical:
        c.alignment = Alignment(vertical="center")     # Excel puts text at the BOTTOM by default; center it so rows line up
    if border:
        c.border = border
    return c


def merged(ws, cell_range, value=None, **kw):
    ws.merge_cells(cell_range)
    return put(ws, cell_range.split(":")[0], value, **kw)


def fill_range(ws, cell_range, fill):
    for row in ws[cell_range]:
        for c in row:
            c.fill = fill


def outline(ws, cell_range, color="DADDE1", weight="thin"):
    """Draw a box around a range (the 'card' look)."""
    side = Side(style=weight, color=color)
    rows = list(ws[cell_range])
    for r_i, row in enumerate(rows):
        for c_i, c in enumerate(row):
            b = c.border
            c.border = Border(left=side if c_i == 0 else b.left,
                              right=side if c_i == len(row) - 1 else b.right,
                              top=side if r_i == 0 else b.top,
                              bottom=side if r_i == len(rows) - 1 else b.bottom)


def section_bar(ws, row, first_col, last_col, text, look, number=None):
    label = f"{number}  ·  {text}" if number else text
    ref = f"{first_col}{row}:{last_col}{row}"
    if look.style == "bold":
        fill_range(ws, ref, PatternFill("solid", fgColor=look.primary))
        merged(ws, ref, label, f=font(12, True, "FFFFFF"), align=Alignment(vertical="center", indent=1))
    else:
        fill_range(ws, ref, PatternFill("solid", fgColor=look.soft))
        merged(ws, ref, label, f=font(12, True, look.primary), align=Alignment(vertical="center", indent=1))
        for c in ws[ref][0]:
            c.border = Border(bottom=Side(style="medium", color=look.accent))
    ws.row_dimensions[row].height = 24


def header_row(ws, row, cols_labels, look, color=None):
    color = color or look.primary
    for col, label in cols_labels:
        if ":" in col:
            a, b = col.split(":")
            fill_range(ws, f"{a}{row}:{b}{row}", PatternFill("solid", fgColor=color))
            merged(ws, f"{a}{row}:{b}{row}", label, f=font(10, True, "FFFFFF"),
                   align=Alignment(horizontal="center", vertical="center", wrap_text=True))
        else:
            put(ws, f"{col}{row}", label, f=font(10, True, "FFFFFF"),
                fill=PatternFill("solid", fgColor=color),
                align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[row].height = 30


def as_input(ws, ref, value=None, fmt=None):
    if fmt == MONEY:
        fmt = "$#,##0.00"          # inputs show $0.00, not a dash
    put(ws, ref, value, f=font(11, color="1D4ED8"), fill=INPUT_FILL, fmt=fmt, border=GRID,
        align=Alignment(horizontal="left" if fmt is None else "right", vertical="center", indent=1))


def as_calc(ws, ref, value=None, fmt=None, bold=False, size=11, color=INK, align="right"):
    put(ws, ref, value, f=font(size, bold, color), fmt=fmt, border=GRID,
        align=Alignment(horizontal=align, vertical="center", indent=1 if align in ("left", "right") else 0))


def total_row(ws, look, row, first_col, last_col):
    """Same look for every Total row: bold, tinted band across the table."""
    for c in range(ord(first_col) - 64, ord(last_col) - 63):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill("solid", fgColor=look.soft)
        cell.font = Font(name=FONT, size=cell.font.size or 11, bold=True, color=cell.font.color)
        h = cell.alignment.horizontal if cell.alignment else None
        cell.alignment = Alignment(horizontal=h, vertical="center", indent=1 if h in ("left", "right", None) else 0)
    ws.row_dimensions[row].height = 22


def red_if_negative(ws, cell_range):
    """Money below zero turns red (left over, left to spend...)."""
    first = cell_range.split(":")[0]
    col = "".join(ch for ch in first if ch.isalpha())
    row = "".join(ch for ch in first if ch.isdigit())
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"AND(ISNUMBER(${col}{row}),${col}{row}<0)"], font=Font(color=RED, bold=True)))


def this_month_highlight(ws, look, cell_range, first_b):
    """Soft-color the row for the current month, so it's easy to find in a long list."""
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"{first_b}={THIS_MONTH}"], fill=PatternFill("solid", fgColor=look.soft)))


def hint(ws, ref, text):
    put(ws, ref, text, f=font(9, italic=True, color=MUTED), align=Alignment(vertical="center", indent=1))


def dropdown(ws, cell_range, choices=None, same_sheet_range=None, keep_order=False):
    """choices = list typed into the dropdown; same_sheet_range = '$J$5:$J$34' on this tab.
    Typed lists are shown A to Z (owner's rule), unless keep_order=True (like 'on payday, 1 day early, 2 days...')."""
    if choices and not keep_order:
        choices = sorted(choices, key=str.lower)
    # No "=" in front of a range: Excel shrugs it off, but Google Sheets silently DROPS the dropdown.
    src = '"' + ",".join(choices) + '"' if choices else same_sheet_range.lstrip("=")
    dv = DataValidation(type="list", formula1=src, allow_blank=True, showDropDown=False)
    dv.error, dv.errorTitle = "Please pick from the list.", "Not on the list"
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(cell_range)


def date_picker(ws, cell_range):
    """Only real dates allowed. In Google Sheets, double-clicking the cell opens a calendar."""
    dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)  # after Jan 1, 2000
    dv.error, dv.errorTitle = "Type a date like 10/22/2026, or double-click for a calendar (Google Sheets).", "Not a date"
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(cell_range)


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def words(cell, *ws_):
    """Formula: TRUE if the cell contains any of these words (not case-sensitive)."""
    return "OR(" + ",".join(f'ISNUMBER(SEARCH("{w}",{cell}))' for w in ws_) + ")"


def status_colors(ws, cell_range, first_cell):
    """Color a status text cell: red if over, amber if almost, green if on track, blue for anything else."""
    rules = [
        (f'LEFT({first_cell},1)="✔"', GREEN, GREEN_BG),          # anything starting with ✔ is good ("✔ Covered" has "over" in it)
        (words(first_cell, "over", "more than", "missed", "not paid", "too low", "never", "short", "late"), RED, RED_BG),
        (words(first_cell, "almost", "behind", "to go", "partly", "due in", "due today", "date passed", "fill in", "close:"), "8A5A00", AMBER_BG),
        (words(first_cell, "on track", "reached", "on plan", "met", "paid", "ready", "✔"), GREEN, GREEN_BG),
        # everything else ("Still to put away", "Coming up"...) gets a calm blue, so every row has a color
        (f'LEN({first_cell})>0', BLUE, BLUE_BG),
    ]
    for formula, fg, bg in rules:
        ws.conditional_formatting.add(cell_range, FormulaRule(
            formula=[formula], stopIfTrue=True, font=Font(color=fg, bold=True), fill=PatternFill("solid", fgColor=bg)))


# ================================================================== ENGINE (hidden)
def build_engine(wb, look):
    ws = wb.create_sheet("Engine")
    t = CONFIG["tax"]
    put(ws, "A1", "ENGINE — behind-the-scenes math. Buyers don't need this tab.", f=font(12, True))
    statuses = list(t["bracket_floors"].keys())
    for i, label in enumerate(["Rate", "Rate step"] + statuses):
        put(ws, f"{col_letter(1 + i)}3", label, f=font(10, True))
    for i, rate in enumerate(t["rates"]):
        r = 4 + i
        put(ws, f"A{r}", rate, fmt="0%")
        put(ws, f"B{r}", "=A4" if i == 0 else f"=A{r}-A{r-1}", fmt="0%")
        for j, s in enumerate(statuses):
            put(ws, f"{col_letter(3 + j)}{r}", t["bracket_floors"][s][i], fmt="$#,##0")
    put(ws, "A20", "First paycheck take-home", f=font(10, True))
    put(ws, "B20", f"=IF(N({setup('first_pay')})>0,{setup('per_check')}*N({setup('first_pay')})/MAX(1,N({setup('pay')})),"
                   f"{setup('per_check')})", fmt=MONEY)
    put(ws, "A12", "Standard deduction", f=font(10, True))
    for j, s in enumerate(statuses):
        put(ws, f"{col_letter(3 + j)}12", t["standard_deduction"][s], fmt="$#,##0")
    put(ws, "A14", "Social Security + Medicare")
    put(ws, "B14", t["fica_rate"], fmt="0.00%")
    put(ws, "A16", "Latest US inflation (auto)")
    put(ws, "B16", CONFIG["inflation"]["auto_rate"], fmt="0.0%")
    put(ws, "A17", "Inflation as of")
    put(ws, "B17", CONFIG["inflation"]["as_of"])
    put(ws, "A19", "Federal taxable income (per year)")
    put(ws, "A18", "Days early the money lands")
    put(ws, "B18", f'=IF(LEFT({setup("early_label")},1)="N",0,IFERROR(VALUE(LEFT({setup("early_label")},1)),0))')
    put(ws, "B19", f"=MAX(0,Setup!$C$32-Setup!$C$33-IFERROR(INDEX($C$12:$E$12,"
                   f"MATCH({setup('status')},$C$3:$E$3,0)),0))", fmt="$#,##0")
    put(ws, "A21", "Pay frequency", f=font(10, True))
    put(ws, "B21", "Paychecks / year", f=font(10, True))
    for i, (n, k) in enumerate([("Weekly", 52), ("Biweekly", 26), ("Semi-monthly", 24), ("Monthly", 12)]):
        put(ws, f"A{22 + i}", n)
        put(ws, f"B{22 + i}", k)
    put(ws, "A28", "Budget plan", f=font(10, True))
    for j in range(G):
        put(ws, f"{col_letter(2 + j)}28", f"Group {j + 1} %", f=font(10, True))
    for i, p in enumerate(CONFIG["presets"]):
        r = 29 + i
        put(ws, f"A{r}", p["name"])
        for j, pct in enumerate(p["pct"]):
            put(ws, f"{col_letter(2 + j)}{r}", pct / 100, fmt="0%")
    # Donut chart data (charts read this, so it lives on a sheet, not in hidden columns)
    put(ws, "A40", "Dashboard donut data", f=font(10, True))
    put(ws, "B40", "Left")
    put(ws, "C40", "Spent")
    for i in range(G):
        r, dr = 41 + i, 11 + i
        put(ws, f"A{r}", f"=Dashboard!$B${dr}")
        # Left = colored part of the ring, Spent = grey part (the ring "drains" as you spend)
        put(ws, f"B{r}", f'=IF(A{r}="",0,MAX(0,Dashboard!$D${dr}-Dashboard!$E${dr}))', fmt="0.00")
        put(ws, f"C{r}", f'=IF(A{r}="",1,MIN(Dashboard!$E${dr},Dashboard!$D${dr}))', fmt="0.00")
    # Every payday for 4 years (feeds "how many paydays land in this month")
    # C = the official payday (pay stub date); B = the day the money lands (weekend rule, then bank-early days)
    put(ws, f"A{PAYDAY_FIRST - 1}", "Paycheck #", f=font(10, True))
    put(ws, f"B{PAYDAY_FIRST - 1}", "Money lands", f=font(10, True))
    put(ws, f"C{PAYDAY_FIRST - 1}", "Official payday", f=font(10, True))
    for k in range(1, PAYDAY_LAST - PAYDAY_FIRST + 2):
        r = PAYDAY_FIRST + k - 1
        put(ws, f"A{r}", k)
        # blank until a first payday is typed on Setup (a fresh template has no paydays yet)
        put(ws, f"C{r}", f'=IF(OR(N({setup("first")})=0,A{r}>{PAYDAY_YEARS}*{setup("per_year")}),"",{payday(f"A{r}")})',
            fmt="mm/dd/yyyy")
        put(ws, f"B{r}", f'=IF(C{r}="","",{weekend_moved(f"C{r}")}-{EARLY_DAYS})', fmt="mm/dd/yyyy")
    # Bank holidays (E name, F year, G date banks are closed) for 5 years from Jan 2026
    put(ws, f"E{HOL_FIRST - 1}", "Bank holiday", f=font(10, True))
    put(ws, f"F{HOL_FIRST - 1}", "Year", f=font(10, True))
    put(ws, f"G{HOL_FIRST - 1}", "Banks closed", f=font(10, True))
    rules = CONFIG["bank_holidays"]["list"]
    for yi in range(CONFIG["bank_holidays"]["years"]):
        for hi, rule in enumerate(rules):
            r = HOL_FIRST + yi * len(rules) + hi
            put(ws, f"E{r}", rule[0])
            put(ws, f"F{r}", f"=YEAR({cal_start()})+{yi}")
            put(ws, f"G{r}", f"={holiday_date(rule, f'F{r}')}", fmt="ddd mm/dd/yyyy")
    widths(ws, {"A": 34, "B": 16, "C": 16, "D": 22, "E": 20, "F": 14, "G": 14})
    wb.defined_names["InflationAuto"] = DefinedName("InflationAuto", attr_text="Engine!$B$16")
    wb.defined_names["InflationAsOf"] = DefinedName("InflationAsOf", attr_text="Engine!$B$17")
    ws.sheet_state = "hidden"


# ================================================================== SETUP
def build_setup(wb, look):
    ws = wb.create_sheet("Setup")
    p = CONFIG["product"]
    widths(ws, {"A": 2, "B": 44, "C": 29, "D": 15, "E": 13, "F": 15, "G": 15, "H": 16, "I": 3, "J": 4})
    merged(ws, "B1:H1", p["name"], f=font(22, True, look.primary), align=Alignment(vertical="center"))
    ws.row_dimensions[1].height = 36
    merged(ws, "B2:H2", "Answer the yellow boxes. Everything else fills in by itself.",
           f=font(11, color=MUTED))
    steps = ("HOW IT WORKS:   1. Answer the questions on this page    "
             "2. Log what you spend on the Spending tab    3. Check your Dashboard and Goals")
    fill_range(ws, "B4:H4", PatternFill("solid", fgColor=look.soft))
    merged(ws, "B4:H4", steps, f=font(10, True, look.primary), align=Alignment(vertical="center", indent=1))
    ws.row_dimensions[4].height = 26
    put(ws, "B5", "Yellow box = you type   ·   White box = automatic, don't type over it",
        f=font(9, italic=True, color=MUTED))

    def question(r, text, value, fmt=None, tip="", calc=False):
        put(ws, f"B{r}", text, f=font(11), align=Alignment(vertical="center", wrap_text=True))
        (as_calc if calc else as_input)(ws, f"C{r}", value, fmt)
        hint(ws, f"D{r}", tip)
        ws.row_dimensions[r].height = 21

    section_bar(ws, 7, "B", "H", "YOUR PAY", look, 1)
    question(8, "How often do you get paid?", "Biweekly", None, "")
    put(ws, "D8", '="Pick from the list. That\'s "&K8&" paychecks a year."',
        f=font(9, italic=True, color=MUTED), align=Alignment(vertical="center", indent=1))
    dropdown(ws, "C8", ["Weekly", "Biweekly", "Semi-monthly", "Monthly"])
    question(9, "First paycheck, if it was smaller (before taxes)", None, MONEY,
             "Your very first check was only part of a full one? Type it. Blank = it was full.")
    question(10, "Full paycheck, before taxes (ONE paycheck)", None, MONEY,
             "One check, not the whole month. Paid monthly? Type your monthly pay.")
    question(11, "Monthly income, before taxes", f"=C10*K8/12", MONEY,
             "", calc=True)
    put(ws, "D11", f'=IF(N(C10)=0,"Automatic: full paycheck × paychecks per year ÷ 12.",'
                   f'"Automatic. After taxes: "&TEXT($D$42,"$#,##0")&" a month (see Take-home below).")',
        f=font(9, italic=True, color=MUTED), align=Alignment(vertical="center", indent=1))
    question(12, "Taken out before taxes, per paycheck", None, MONEY,
             "401k, health insurance, etc. Leave blank if you don't have any.")
    question(13, "When was your first payday this year?", None, DATE_FMT,
             "The payday on your FIRST pay stub, even if it was a half check.")
    date_picker(ws, "C13")
    # paychecks per year: hidden helper (K8), shown in the tip next to "How often"
    put(ws, "K8", "=IFERROR(INDEX(Engine!$B$22:$B$25,MATCH(C8,Engine!$A$22:$A$25,0)),26)")
    question(14, "Paid twice a month? Which days?", DAYS_FROM_FIRST, None,
             "Like the 15th & 30th. Only used if you're paid Semi-monthly.")
    dropdown(ws, "C14", list(PAY_DAY_CHOICES))
    question(15, "Payday on a weekend or bank holiday?", WEEKEND_BEFORE, None,
             "Most jobs pay the business day before. Holidays: Christmas, Labor Day, etc.")
    dropdown(ws, "C15", [WEEKEND_BEFORE, WEEKEND_AFTER, WEEKEND_SAME])
    question(16, "Does your bank pay you early?", EARLY_CHOICES[0], None,
             "Chime, SoFi and many credit unions pay 1–2 days early.")
    dropdown(ws, "C16", EARLY_CHOICES, keep_order=True)
    section_bar(ws, 17, "B", "H", "TAXES", look, 2)
    question(18, "How should we figure your taxes?", "Estimate for me", None,
             "'Estimate for me' uses 2026 US tax rules. Or type your own rate below.")
    dropdown(ws, "C18", ["Estimate for me", TAX_MANUAL_LABEL])
    question(19, "Filing status", "Single", None, "Most students and first-jobbers are 'Single'.")
    dropdown(ws, "C19", list(CONFIG["tax"]["standard_deduction"].keys()))
    question(20, "State + local income tax rate", None, PCT1, "Leave blank if your state has no income tax.")
    question(21, "Take out Social Security + Medicare?", "Yes", None, "Almost every job does (7.65%).")
    dropdown(ws, "C21", ["Yes", "No"])
    question(22, "Your own total tax rate", None, PCT1,
             "Only used if you picked 'Use my own rate'. Tip: all taxes on a pay stub ÷ gross pay.")

    section_bar(ws, 24, "B", "H", "INFLATION (prices going up)", look, 3)
    question(25, "Which inflation rate should we use?", "Latest US rate", None,
             "Inflation makes the same stuff cost more next year.")
    dropdown(ws, "C25", ["Latest US rate", INFL_MANUAL_LABEL])
    question(26, "Latest US rate", "=Engine!$B$16", PCT1, "", calc=True)
    put(ws, "D26", '="as of "&Engine!$B$17', f=font(9, italic=True, color=MUTED), align=Alignment(vertical="center", indent=1))
    question(27, "Your own inflation guess", None, PCT1, "Only used if you picked 'My own rate'.")
    question(28, "Rate being used", f'=IF(C25="{INFL_MANUAL_LABEL}",C27,C26)', PCT1, "", calc=True)

    # ---- result card: what you really take home
    section_bar(ws, 30, "B", "H", "WHAT YOU REALLY TAKE HOME", look)
    for col, label in (("C", "Per year"), ("D", "Per month")):
        put(ws, f"{col}31", label, f=font(10, True, look.primary), align=Alignment(horizontal="right"))
    manual = f'{S["tax_mode"]}="{TAX_MANUAL_LABEL}"'
    status_col = f"MATCH({S['status']},Engine!$C$3:$E$3,0)"
    lines = [
        (32, "Pay before taxes", f"={S['pay']}*{S['per_year']}"),
        (33, "Taken out before taxes (401k, health)", f"={S['pretax']}*{S['per_year']}"),
        (34, "Federal income tax (estimate)",
         f"=IF({manual},0,SUMPRODUCT((Engine!$B$19>INDEX(Engine!$C$4:$E$10,0,{status_col}))"
         f"*(Engine!$B$19-INDEX(Engine!$C$4:$E$10,0,{status_col}))*Engine!$B$4:$B$10))"),
        (35, "Social Security + Medicare", f'=IF(OR({manual},{S["fica"]}<>"Yes"),0,C32*Engine!$B$14)'),
        (36, "State + local tax", f"=IF({manual},0,MAX(0,C32-C33)*{S['state']})"),
        (37, "Taxes at your own rate", f"=IF({manual},C32*{S['manual']},0)"),
        (38, "Total taxes", "=SUM(C34:C37)"),
        (39, "TAKE-HOME PAY", "=C32-C33-C38"),
    ]
    for r in range(31, 43):                 # same row height as the questions above
        ws.row_dimensions[r].height = 21
    for r, label, formula in lines:
        big = r == 39
        put(ws, f"B{r}", label, f=font(13 if big else 11, big or r == 38))
        as_calc(ws, f"C{r}", formula, MONEY, bold=big or r == 38, size=13 if big else 11)
        as_calc(ws, f"D{r}", f"=C{r}/12", MONEY, bold=big or r == 38, size=13 if big else 11)
    for col in "BCD":
        ws[f"{col}39"].fill = PatternFill("solid", fgColor=look.soft)
    ws.row_dimensions[39].height = 24
    put(ws, "B40", "Share of your pay that goes to taxes")
    as_calc(ws, "C40", "=IF(C32=0,0,C38/C32)", PCT1)
    put(ws, "B41", "Each $1 you spend costs this much pay")
    as_calc(ws, "C41", "=IF(C40>=1,0,1/(1-C40))", "$0.00")
    put(ws, "B42", "Take-home per paycheck")
    as_calc(ws, "C42", f"=IF({S['per_year']}=0,0,C39/{S['per_year']})", MONEY)
    as_calc(ws, "D42", "=D39", MONEY)                   # monthly take-home again, for the Monthly income tip
    ws["D42"].font = font(9, color="FFFFFF")            # (hidden helper: white on white)
    ws["D42"].border = Border()
    hint(ws, "E39", "This is the money you actually get to budget.")

    # ---- budget plan (groups)
    section_bar(ws, 44, "B", "H", "YOUR BUDGET PLAN", look, 4)
    put(ws, "B45", "Groups are the big buckets your money goes into. Rename them, or pick a plan to fill in the %.",
        f=font(9, italic=True, color=MUTED))
    plans = [p_["name"] for p_ in CONFIG["presets"]] + [CONFIG["custom_plan_label"]]
    question(46, "Pick a budget plan", CONFIG["default_plan"], None,
             "Pick 'Custom' to type your own % in the 'My own %' column.")
    dropdown(ws, "C46", plans)
    header_row(ws, GROUP_FIRST - 1, [("B", "Group name"), ("C", "What kind?"), ("D", "My own %\n(Custom only)"),
                        ("E", "% in use"), ("F", "Per month"), ("G", "Grows with\ninflation?"),
                        ("H", "Next year\nper month")], look)
    preset_rows = f"Engine!$A$29:$A${28 + len(CONFIG['presets'])}"
    for i in range(G):
        r = GROUP_FIRST + i
        d = CONFIG["default_groups"][i] if i < len(CONFIG["default_groups"]) else {"name": "", "kind": "", "inflation": ""}
        as_input(ws, f"B{r}", d["name"] or None)
        as_input(ws, f"C{r}", d["kind"] or None)
        as_input(ws, f"D{r}", CONFIG["default_custom_pct"][i] / 100, PCT)
        slot_col = col_letter(2 + i)
        preset_vals = f"Engine!${slot_col}$29:${slot_col}${28 + len(CONFIG['presets'])}"
        as_calc(ws, f"E{r}", f'=IF(B{r}="","",IF({S["plan"]}="{CONFIG["custom_plan_label"]}",D{r},'
                             f'IFERROR(INDEX({preset_vals},MATCH({S["plan"]},{preset_rows},0)),0)))', PCT, bold=True)
        as_calc(ws, f"F{r}", f'=IF(B{r}="","",E{r}*{S["take_m"]})', MONEY)
        as_input(ws, f"G{r}", d["inflation"] or None)
        as_calc(ws, f"H{r}", f'=IF(B{r}="","",IF(G{r}="Yes",F{r}*(1+{S["infl_used"]}),F{r}))', MONEY)
        if look.style == "bold":
            ws[f"B{r}"].border = Border(left=Side(style="thick", color=look.group_color(i)),
                                        right=LINE, top=LINE, bottom=LINE)
    dropdown(ws, f"C{GROUP_FIRST}:C{GROUP_LAST}", CONFIG["budget_kinds"])
    dropdown(ws, f"G{GROUP_FIRST}:G{GROUP_LAST}", ["Yes", "No"])
    t = GROUP_TOTAL
    put(ws, f"B{t}", "Total", f=font(11, True))
    as_calc(ws, f"E{t}", f"=SUM(E{GROUP_FIRST}:E{GROUP_LAST})", PCT, bold=True)
    as_calc(ws, f"F{t}", f"=SUM(F{GROUP_FIRST}:F{GROUP_LAST})", MONEY, bold=True)
    as_calc(ws, f"H{t}", f"=SUM(H{GROUP_FIRST}:H{GROUP_LAST})", MONEY, bold=True)
    total_row(ws, look, t, "B", "H")
    merged(ws, f"B{t+1}:H{t+1}",
           f'=IF(ROUND(E{t},4)=1,"✔ All of your take-home has a job. Nice!",'
           f'IF(E{t}>1,"⚠ Your groups add up to "&TEXT(E{t},"0%")&". Lower one so it equals 100%.",'
           f'"ℹ "&TEXT(1-E{t},"0%")&" of your take-home has no group yet. Add it to one."))',
           f=font(11, True, look.primary))
    ws.conditional_formatting.add(f"B{t+1}", FormulaRule(formula=[f"ROUND($E${t},4)<>1"],
                                                          font=Font(color=RED, bold=True)))

    # ---- categories
    section_bar(ws, CAT_FIRST - 3, "B", "H", "YOUR CATEGORIES", look, 5)
    put(ws, f"B{CAT_FIRST - 2}", "List what you spend on and pick its group. Monthly goal is optional. Bills: add the day "
                   "of the month it's due, and it shows on the Bills tab with a reminder.", f=font(9, italic=True, color=MUTED))
    header_row(ws, CAT_FIRST - 1, [("B", "Category"), ("C", "Group"), ("D", "Monthly goal\n(optional)"),
                        ("E", "Due day\n(bills, 1–31)")], look)
    ws.row_dimensions[CAT_FIRST - 1].height = 34
    for i in range(C_SLOTS):
        r = CAT_FIRST + i
        d = CONFIG["default_categories"][i] if i < len(CONFIG["default_categories"]) else (None, None)
        as_input(ws, f"B{r}", d[0])
        as_input(ws, f"C{r}", d[1])
        as_input(ws, f"D{r}", d[2] if len(d) > 2 else None, MONEY)
        as_input(ws, f"E{r}", d[3] if len(d) > 3 else None, "0")
        # Hidden helper: "Needs#3" = third category in the Needs group (feeds the Dashboard lists)
        put(ws, f"J{r}", f'=IF(OR(B{r}="",C{r}=""),"",C{r}&"#"&COUNTIFS($C${CAT_FIRST}:C{r},C{r},'
                         f'$B${CAT_FIRST}:B{r},"<>"))', f=font(8, color=MUTED))
        # Hidden helper: place in A-to-Z order (ties keep their row order), so dropdowns list categories alphabetically
        put(ws, f"L{r}", f'=IF(B{r}="","",COUNTIF($B${CAT_FIRST}:$B${CAT_LAST},"<"&B{r})+COUNTIF($B${CAT_FIRST}:B{r},B{r}))',
            f=font(8, color=MUTED))
        # Hidden helper: the category's kind, looked up from its group (feeds the Goals tab)
        put(ws, f"K{r}", f'=IF(C{r}="","",IFERROR(INDEX({G_KIND},MATCH(C{r},{G_NAME},0)),""))',
            f=font(8, color=MUTED))
    # group names A to Z for the Group dropdown (hidden column M next to the groups)
    for i in range(G):
        r = GROUP_FIRST + i
        put(ws, f"N{r}", f'=IF(B{r}="","",COUNTIF($B${GROUP_FIRST}:$B${GROUP_LAST},"<"&B{r})+COUNTIF($B${GROUP_FIRST}:B{r},B{r}))',
            f=font(8, color=MUTED))
        put(ws, f"M{r}", f'=IFERROR(INDEX($B${GROUP_FIRST}:$B${GROUP_LAST},MATCH({i + 1},$N${GROUP_FIRST}:$N${GROUP_LAST},0)),"")',
            f=font(8, color=MUTED))
    dropdown(ws, f"C{CAT_FIRST}:C{CAT_LAST}", same_sheet_range=G_ALPHA_LIST)
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="31", allow_blank=True)
    dv.error, dv.errorTitle, dv.showErrorMessage = "Type a day of the month, 1 to 31.", "Due day", True
    ws.add_data_validation(dv)
    dv.add(f"E{CAT_FIRST}:E{CAT_LAST}")
    for col in "JKLMN":
        ws.column_dimensions[col].hidden = True
    ws.freeze_panes = "A3"


# ================================================================== SPENDING
def build_spending(wb, look):
    """Spending: one row per purchase. Each row shows how its budget is doing that month (F), and if it puts you
    over, you can OK it with a reason (G) and/or cover it (H). Money in (right side) can be marked 'For' a
    category or group: that extra money adds to its budget that month."""
    ws = wb.create_sheet("Spending")
    widths(ws, {"A": 2, "B": 14, "C": 28, "D": 20, "E": 12, "F": 44, "G": 20, "H": 23, "I": 22,
                "J": 4, "K": 4, "L": 4, "M": 4, "N": 4, "O": 4, "P": 3,
                "Q": 14, "R": 24, "S": 12, "T": 18, "U": 20, "V": 4, "W": 4, "X": 4, "Y": 4})
    merged(ws, "B1:I1", "Spending", f=font(22, True, look.primary))
    ws.row_dimensions[1].height = 36
    merged(ws, "B2:I2", "One row per purchase or bill. Pick the category: column F tells you how that budget is doing. "
                        "Added something late? Budget Tools → Sort by date.", f=font(11, color=MUTED))
    # row 3: how much is safe to spend before the next paycheck (from the Bills tab), so you see it while logging
    merged(ws, "B3:I3", f'=IF(Bills!$AK$2="","Add your pay and first payday on Setup to see what is safe to spend.",'
                        f'IF(Bills!$E${SAFE_ROW}<0,"⚠ Careful: you\'re "&TEXT(-Bills!$E${SAFE_ROW},"$#,##0")&" short for bills due before "'
                        f'&TEXT(Bills!$AK$3,"mmm d")&". Check the Bills tab.",'
                        f'"✔ Safe to spend until "&TEXT(Bills!$AK$3,"ddd, mmm d")&": "&TEXT(Bills!$E${SAFE_ROW},"$#,##0")'
                        f'&"  (your bills before then are covered). Planning a purchase? Bills tab → Can I afford it?"))',
           f=font(10, True, GREEN), align=Alignment(vertical="center", indent=1))
    ws.row_dimensions[3].height = 24
    for formula, fg, bg in ((f"Bills!$E${SAFE_ROW}<0", RED, RED_BG), ("TRUE", GREEN, GREEN_BG)):
        ws.conditional_formatting.add("B3:I3", FormulaRule(formula=[formula], stopIfTrue=True,
                                      font=Font(color=fg, bold=True), fill=PatternFill("solid", fgColor=bg)))
    header_row(ws, 4, [("B", "Date"), ("C", "What was it?"), ("D", "Category"), ("E", "Amount"),
                       ("F", "Budget check"), ("G", "OK'd?"), ("H", "Cover with"), ("I", "Notes (optional)")], look)
    almost = CONFIG["alert_almost_pct"]
    for r in range(SP_FIRST, SP_LAST + 1):
        for col, fmt in zip("BCDE", [DATE_FMT, None, None, MONEY]):
            as_input(ws, f"{col}{r}", None, fmt)
        for col in "GHI":
            as_input(ws, f"{col}{r}", None, None)
        m0 = f"DATE(YEAR(B{r}),MONTH(B{r}),1)"
        m1, p0 = f"EDATE({m0},1)", f"EDATE({m0},-1)"
        goal = f"N(IFERROR(INDEX({C_GOAL},MATCH(D{r},{C_NAME},0)),0))"
        # hidden helpers
        ws[f"J{r}"] = f'=IF(D{r}="","",IFERROR(INDEX({C_GROUP},MATCH(D{r},{C_NAME},0)),"(no group)"))'      # group
        ws[f"K{r}"] = f'=IF(J{r}="","",IFERROR(INDEX({G_KIND},MATCH(J{r},{G_NAME},0)),""))'                  # kind
        ws[f"L{r}"] = f'=IF(D{r}="","",IF({goal}>0,D{r},J{r}))'      # which budget it checks: the category (if it has a goal) or its group
        ws[f"M{r}"] = f'=IF(OR(G{r}="",B{r}=""),"",D{r}&"|"&TEXT(B{r},"yyyymm"))'    # an OK'd purchase (category, month)
        ws[f"N{r}"] = f'=IF(OR(G{r}="",B{r}=""),"",J{r}&"|"&TEXT(B{r},"yyyymm"))'    # an OK'd purchase (group, month)
        cat_or_grp = lambda d0, d1, op1="<": (f'IF(L{r}=D{r},SUMIFS({SP_AMT},{SP_CAT},D{r},{SP_DATE},">="&{d0},{SP_DATE},"{op1}"&{d1}),'
                                              f'SUMIFS({SP_AMT},{SP_GROUP},J{r},{SP_DATE},">="&{d0},{SP_DATE},"{op1}"&{d1}))')
        base = f'IF(L{r}=D{r},{goal},N(IFERROR(INDEX({G_MONTH},MATCH(J{r},{G_NAME},0)),0)))'
        extra = f'SUMIFS({MI_AMT},{MI_FOR},L{r},{MI_DATE},">="&{m0},{MI_DATE},"<"&{m1})'
        roll_on = f'COUNTIFS({SP_COVER},"{COVER_LEFTOVER}",{SP_KEY},L{r},{SP_DATE},">="&{m0},{SP_DATE},"<"&{m1})>0'
        roll = f'IF({roll_on},MAX(0,{base}-{cat_or_grp(p0, m0)}),0)'
        ws[f"X{r}"] = f'=IF(L{r}="","",{base}+{extra}+{roll})'                       # budget this month (with covers)
        ws[f"Y{r}"] = f'=IF(L{r}="","",{cat_or_grp(m0, f"B{r}", "<=")})'               # spent up to this purchase
        ws[f"W{r}"] = f'=IF(L{r}="","",{base})'                                         # plain budget (no covers)
        # whole dollars, but show cents when it's under $1 (so "Over by $0" becomes "Over by $0.36")
        money = lambda x: f'TEXT({x},IF(ABS({x})<1,"$0.00","$#,##0"))'
        as_calc(ws, f"F{r}", f'=IF(OR(B{r}="",D{r}="",N(E{r})=0),"",IF(K{r}="{KIND_SAVE}","✔ Into savings",'
                             f'IF(N(X{r})=0,"No budget yet for "&L{r}&" (Setup)",IF(Y{r}>X{r},IF(G{r}<>"",'
                             f'"Noted: "&{money(f"Y{r}-X{r}")}&" over · "&G{r},"⚠ Over by "&{money(f"Y{r}-X{r}")}&" in "&L{r}&": OK it or cover it"),'
                             f'IF(ABS(X{r}-Y{r})<0.005,"✔ Right on budget in "&L{r},IF(X{r}-Y{r}<={1 - almost}*X{r},"Close: "&{money(f"X{r}-Y{r}")}&" left in "&L{r},'
                             f'IF(Y{r}>W{r},"✔ Covered: "&{money(f"X{r}-Y{r}")}&" left in "&L{r},"✔ "&{money(f"X{r}-Y{r}")}&" left in "&L{r})))))))',
                align="left")
        ws.row_dimensions[r].height = 20
    rows = f"{SP_FIRST}:{SP_LAST}"
    f_rng = f"F{SP_FIRST}:F{SP_LAST}"
    ws.conditional_formatting.add(f_rng, FormulaRule(formula=[f'LEFT($F{SP_FIRST},5)="Noted"'], stopIfTrue=True,
                                  font=Font(color=BLUE, bold=True), fill=PatternFill("solid", fgColor=BLUE_BG)))
    ws.conditional_formatting.add(f_rng, FormulaRule(formula=[f'LEFT($F{SP_FIRST},1)="✔"'], stopIfTrue=True,   # "✔ Covered" has "over" in it
                                  font=Font(color=GREEN, bold=True), fill=PatternFill("solid", fgColor=GREEN_BG)))
    status_colors(ws, f_rng, f"$F{SP_FIRST}")
    # an over-budget row lights up its OK'd? and Cover with boxes until you fill one in
    ws.conditional_formatting.add(f"G{SP_FIRST}:H{SP_LAST}", FormulaRule(
        formula=[f'AND(LEFT($F{SP_FIRST},1)="⚠",$G{SP_FIRST}="",$H{SP_FIRST}="")'],
        fill=PatternFill("solid", fgColor="FBCFCB"), font=Font(color=RED, bold=True)))
    for i in range(C_SLOTS):
        ws[f"O{5 + i}"] = f'=IFERROR(INDEX({C_NAME},MATCH({i + 1},{C_ALPHA},0)),"")'     # categories A to Z
    dropdown(ws, f"D{SP_FIRST}:D{SP_LAST}", same_sheet_range=f"$O$5:$O${4 + C_SLOTS}")
    dropdown(ws, f"G{SP_FIRST}:G{SP_LAST}", OK_REASONS)
    dropdown(ws, f"H{SP_FIRST}:H{SP_LAST}", COVER_CHOICES)
    date_picker(ws, f"B{SP_FIRST}:B{SP_LAST}")
    for col in "JKLMNOVWXY":
        ws.column_dimensions[col].hidden = True

    # ---- Money in (optional): every deposit, even part of a paycheck. "For" = extra money for one budget.
    merged(ws, "Q1:U1", "Money in", f=font(22, True, look.primary))
    merged(ws, "Q2:U2", "Optional. Log each deposit when it lands. Extra money (side gig, gift) for one budget? "
                        "Pick it under 'For' and that budget gets it this month.",
           f=font(9, italic=True, color=MUTED), align=Alignment(wrap_text=True, vertical="top", indent=1))
    ws.row_dimensions[2].height = 30
    header_row(ws, 4, [("Q", "Date received"), ("R", "From (job, gig, gift...)"), ("S", "Amount"),
                       ("T", "For (optional)"), ("U", "Notes (optional)")], look,
               color=look.accent if look.style == "card" else "2E7D32")
    for r in range(MI_FIRST, MI_LAST + 1):
        for col, fmt in zip("QRSTU", [DATE_FMT, None, MONEY, None, None]):
            as_input(ws, f"{col}{r}", None, fmt)
        # hidden: which group the extra money goes to (so the Dashboard group adds it)
        ws[f"V{r}"] = (f'=IF(T{r}="","",IF(ISNUMBER(MATCH(T{r},{G_NAME},0)),T{r},'
                       f'IFERROR(INDEX({C_GROUP},MATCH(T{r},{C_NAME},0)),"")))')
    # 'For' list: groups A to Z, then categories A to Z (hidden column at the far right)
    n_groups = f"SUMPRODUCT(--(LEN(Setup!{G_ALPHA_LIST})>0))"
    for i in range(G + C_SLOTS):
        ws[f"Z{5 + i}"] = (f'=IF({i + 1}<={n_groups},INDEX(Setup!{G_ALPHA_LIST},{i + 1}),'
                           f'IFERROR(INDEX({C_NAME},MATCH({i + 1}-{n_groups},{C_ALPHA},0)),""))')
    ws.column_dimensions["Z"].hidden = True
    dropdown(ws, f"T{MI_FIRST}:T{MI_LAST}", same_sheet_range=f"$Z$5:$Z${4 + G + C_SLOTS}")
    date_picker(ws, f"Q{MI_FIRST}:Q{MI_LAST}")
    ws.freeze_panes = "C5"          # header rows + the Date column stay put while you swipe (great on a phone)


# ================================================================== DASHBOARD
def build_dashboard(wb, look):
    ws = wb.create_sheet("Dashboard")
    widths(ws, {"A": 2, **{col_letter(c): 13 for c in range(2, 14)}, "N": 2, "U": 12})
    merged(ws, "B1:H1", "Dashboard", f=font(22, True, look.primary))
    ws.row_dimensions[1].height = 36
    put(ws, "B2", "Month:", f=font(11, True), align=Alignment(horizontal="right", vertical="center"))
    ws.merge_cells("C2:D2")
    as_input(ws, "C2", "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)", "mmmm yyyy")
    hint(ws, "E2", f"Pick any month from {CAL_LABEL} (starts on this month).")
    # Hidden helpers in column U
    put(ws, "U2", "=DATE(YEAR(C2),MONTH(C2),1)", fmt="mm/dd/yyyy")       # month start
    put(ws, "U3", "=EDATE(U2,1)", fmt="mm/dd/yyyy")                       # next month start
    put(ws, "U4", "=IF(AND(TODAY()>=U2,TODAY()<U3),U3-TODAY(),0)")       # days left incl. today
    for i in range(CAL_MONTHS):                                           # U6:U41 = Jan 2026 ... Dec 2028
        put(ws, f"U{6 + i}", f"=EDATE({cal_start()},{i})", fmt="mmmm yyyy")
    dropdown(ws, "C2", same_sheet_range=f"$U$6:$U${5 + CAL_MONTHS}")
    ws.column_dimensions["U"].hidden = True

    month = between("$U$2", "$U$3")
    flex_plan = f'SUMIFS({G_MONTH},{G_KIND},"{KIND_FLEX}")'
    flex_spent = f'SUMIFS({SP_AMT},{SP_KIND},"{KIND_FLEX}",{month})'
    # Hidden helpers in column V: this month's paydays, expected pay, money actually received
    put(ws, "V2", f"={paydays_in('$U$2', '$U$3')}")
    put(ws, "V3", f"={expected_pay('$U$2', '$U$3')}", fmt=MONEY)
    put(ws, "V4", f"={received_in('$U$2', '$U$3')}", fmt=MONEY)
    ws.column_dimensions["V"].hidden = True

    # ---- big number cards (row 4 label, row 5 number, row 6 subtext)
    cards = [
        ("B:D", "PAY THIS MONTH", "=$V$3",
         f'=$V$2&IF($V$2=1," payday"," paydays")&" this month  ·  average month: "&TEXT({setup("take_m")},"$#,##0")'),
        ("E:G", "SPENT SO FAR", f"=SUMIFS({SP_AMT},{month})",
         '="Everything logged in "&TEXT($U$2,"mmmm")'),
        ("H:J", "LEFT THIS MONTH", "=B5-E5",
         '=IF($V$4>0,"Cash in hand: "&TEXT($V$4-E5,"$#,##0")&" (received − spent)",'
         'IF(H5<0,"You spent more than you expect to get","Based on the pay you expect"))'),
        ("K:M", "YOU CAN SPEND / DAY", f'=IF($U$4>0,MAX(0,{flex_plan}-{flex_spent})/$U$4,"—")',
         '=IF($U$4>0,"Flexible money, for the next "&$U$4&" days","Shows for the current month")'),
    ]
    big_cards(ws, look, [(sp, lab, val, MONEY0, sub) for sp, lab, val, sub in cards])
    ws.conditional_formatting.add("H5", FormulaRule(formula=["$H$5<0"], font=Font(color=RED, bold=True, size=24)))
    merged(ws, "B7:M7",
           f'=IF($U$4>0,"Spent today: "&TEXT(SUMIFS({SP_AMT},{SP_DATE},TODAY()),"$#,##0.00")&"   ·   "&$U$4&'
           f'" days left in "&TEXT($U$2,"mmmm"),"You are looking at "&TEXT($U$2,"mmmm yyyy")&". '
           f'Pick this month to see today\'s numbers.")',
           f=font(10, True, look.primary), align=Alignment(horizontal="center", vertical="center"))
    ws.row_dimensions[7].height = 22
    merged(ws, "B8:M8",
           '=IF($V$4>0,"Money in: "&TEXT($V$4,"$#,##0.00")&" received of "&TEXT($V$3,"$#,##0.00")&" expected   ·   "'
           '&IF($V$3-$V$4>0,TEXT($V$3-$V$4,"$#,##0.00")&" still coming","all of it is in ✔"),'
           '"Tip: log deposits under Money in (Spending tab, right side) to see what has actually arrived.")',
           f=font(10, color=look.primary), align=Alignment(horizontal="center", vertical="center"))

    # ---- groups table
    section_bar(ws, 9, "B", "M", "YOUR GROUPS THIS MONTH", look, "A")
    header_row(ws, 10, [("B:C", "Group"), ("D", "Plan\n(+ extra)"), ("E", "Spent / put away"), ("F", "Left"),
                        ("G", "% used"), ("H:J", "Progress"), ("K:M", "How it's going")], look)
    for i in range(G):
        r, sr = 11 + i, GROUP_FIRST + i
        ws.merge_cells(f"B{r}:C{r}")
        as_calc(ws, f"B{r}", f'=IF(Setup!$B${sr}="","",Setup!$B${sr})', bold=True, align="left",
                color=look.group_color(i) if look.style == "bold" else INK)
        # plan + extra money in marked "For" this group (or one of its categories) + last month's leftover if you used it
        prev = f'SUMIFS({SP_AMT},{SP_GROUP},B{r},{SP_DATE},">="&EDATE($U$2,-1),{SP_DATE},"<"&$U$2)'
        roll = (f'IF(COUNTIFS({SP_COVER},"{COVER_LEFTOVER}",{SP_KEY},B{r},{SP_DATE},">="&$U$2,{SP_DATE},"<"&$U$3)>0,'
                f'MAX(0,Setup!$F${sr}-{prev}),0)')
        as_calc(ws, f"D{r}", f'=IF(B{r}="","",Setup!$F${sr}+SUMIFS({MI_AMT},{MI_FORGRP},B{r},{MI_DATE},">="&$U$2,{MI_DATE},"<"&$U$3)+{roll})', MONEY)
        as_calc(ws, f"E{r}", f'=IF(B{r}="","",SUMIFS({SP_AMT},{SP_GROUP},B{r},{month}))', MONEY)
        as_calc(ws, f"F{r}", f'=IF(B{r}="","",D{r}-E{r})', MONEY)
        as_calc(ws, f"G{r}", f'=IF(OR(B{r}="",N(D{r})=0),"",E{r}/D{r})', PCT)
        ws.merge_cells(f"H{r}:J{r}")
        as_calc(ws, f"H{r}", f'=IF(G{r}="","",REPT("█",MIN(10,ROUND(G{r}*10,0)))&'
                             f'REPT("░",10-MIN(10,ROUND(G{r}*10,0))))',
                color=look.group_color(i), align="left")
        ws.merge_cells(f"K{r}:M{r}")
        kind = f"Setup!$C${sr}"
        as_calc(ws, f"K{r}",
                f'=IF(B{r}="","",IF(N(D{r})=0,"Add your pay on Setup",IF({kind}="{KIND_SAVE}",'
                f'IF(E{r}>=D{r},"✔ Goal reached","Still to put away: "&TEXT(F{r},"$#,##0")),'
                f'IF(E{r}>D{r},"Over by "&TEXT(-F{r},"$#,##0"),'
                f'IF(AND(D{r}>0,E{r}>={CONFIG["alert_almost_pct"]}*D{r}),"Almost out: "&TEXT(F{r},"$#,##0")&" left",'
                f'"On track: "&TEXT(F{r},"$#,##0")&" left")))))', align="left")
        for c in ws[f"B{r}:M{r}"][0]:
            c.border = UNDER
    status_colors(ws, "K11:M16", "$K11")
    put(ws, "B17", "Total", f=font(11, True))
    for col in "DEF":
        as_calc(ws, f"{col}17", f"=SUM({col}11:{col}16)", MONEY, bold=True)
    total_row(ws, look, 17, "B", "M")
    red_if_negative(ws, "F11:F17")

    build_alerts(ws, look, month, flex_plan, flex_spent)

    # ---- donuts: one per group (data lives on Engine)
    section_bar(ws, 19, "B", "M", "LEFT TO SPEND, BY GROUP", look, "B")
    spots = [("B", 20), ("F", 20), ("J", 20), ("B", 31), ("F", 31), ("J", 31)]
    for i, (col, row) in enumerate(spots[:G]):
        end_col = col_letter(ord(col) - 64 + 3)
        merged(ws, f"{col}{row}:{end_col}{row}",
               f'=IF(B{11+i}="","(Empty: add a group on the Setup tab)",IF(N(D{11+i})=0,B{11+i}&": add your pay on Setup",'
               f'B{11+i}&":  "&TEXT(MAX(0,F{11+i}),"$#,##0")&" left of "&TEXT(D{11+i},"$#,##0")))',
               f=font(11, True, look.group_color(i)), align=Alignment(horizontal="center"))
        chart = DoughnutChart(holeSize=60)
        data = Reference(wb["Engine"], min_col=2, max_col=3, min_row=41 + i, max_row=41 + i)
        chart.add_data(data, from_rows=True, titles_from_data=False)
        chart.set_categories(Reference(wb["Engine"], min_col=2, max_col=3, min_row=40, max_row=40))
        chart.legend = None
        chart.plotVisOnly = False
        series = chart.series[0]
        for idx, color in enumerate([look.group_color(i) if look.style == "bold" else look.accent, "E5E7EB"]):
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = color
            pt.graphicalProperties.line.solidFill = "FFFFFF"
            series.dPt.append(pt)
        chart.width, chart.height = 8.5, 4.8
        ws.add_chart(chart, f"{col}{row + 1}")

    # ---- category cards: each group's categories
    lines = CONFIG["dashboard_lines_per_group"]
    section_bar(ws, 43, "B", "M", "YOUR CATEGORIES THIS MONTH", look, "C")
    card_cols = ["B", "F", "J"]
    for i in range(G):
        top = 45 + (i // 3) * (lines + 4)
        c0 = card_cols[i % 3]
        ci = ord(c0) - 64
        c1, c2, c3 = (col_letter(ci + k) for k in (1, 2, 3))
        name = f"$B${11 + i}"
        color = look.group_color(i) if look.style == "bold" else look.primary
        fill_range(ws, f"{c0}{top}:{c3}{top}", PatternFill("solid", fgColor=color))
        merged(ws, f"{c0}{top}:{c3}{top}", f'=IF({name}="","(EMPTY: ADD A GROUP ON SETUP)",UPPER({name}))',
               f=font(11, True, "FFFFFF"), align=Alignment(horizontal="center", vertical="center"))
        for col, label in ((f"{c0}:{c1}", "Category"), (c2, "Spent"), (c3, "Goal")):
            if ":" in col:
                a, b = col.split(":")
                merged(ws, f"{a}{top+1}:{b}{top+1}", label, f=font(9, True, MUTED), align=Alignment(indent=1))
            else:
                put(ws, f"{col}{top+1}", label, f=font(9, True, MUTED),
                    align=Alignment(horizontal="right", indent=1 if col == c3 else 0))
        for k in range(1, lines + 1):
            r = top + 1 + k
            ws.merge_cells(f"{c0}{r}:{c1}{r}")
            put(ws, f"{c0}{r}", f'=IF({name}="","",IFERROR(INDEX({C_NAME},MATCH({name}&"#{k}",{C_RANK},0)),""))',
                f=font(10), align=Alignment(indent=1))
            put(ws, f"{c2}{r}", f'=IF({c0}{r}="","",SUMIFS({SP_AMT},{SP_CAT},{c0}{r},{month}))', f=font(10), fmt=MONEY)
            put(ws, f"{c3}{r}", f'=IF({c0}{r}="","",IFERROR(INDEX({C_GOAL},MATCH({c0}{r},{C_NAME},0)),""))',
                f=font(10, color=MUTED), fmt=MONEY, align=Alignment(horizontal="right", indent=1))
            for col in (c0, c1, c2, c3):
                ws[f"{col}{r}"].border = UNDER
            ws.conditional_formatting.add(f"{c2}{r}", FormulaRule(
                formula=[f"AND(N({c3}{r})>0,N({c2}{r})>N({c3}{r}))"], font=Font(color=RED, bold=True)))
        outline(ws, f"{c0}{top}:{c3}{top + 1 + lines}", "DADDE1")
    bottom = 45 + 2 * (lines + 4)
    put(ws, f"B{bottom - 1}", "Red = over that category's goal. Only the first "
                             f"{lines} categories of each group show here; every category is in Reports.",
        f=font(9, italic=True, color=MUTED))

    # ---- taxes & inflation in plain words
    p = bottom + 1
    section_bar(ws, p, "B", "M", "TAXES & INFLATION, IN PLAIN WORDS", look)
    sentences = [
        f'="Out of every $100 you earn, about $"&TEXT({setup("eff")}*100,"0")&" goes to taxes."',
        f'="So a $50 purchase really takes about $"&TEXT(50*{setup("cost1")},"0.00")&" of your pay."',
        f'="With "&TEXT({setup("infl_used")},"0.0%")&" inflation, your take-home will buy about $"'
        f'&TEXT({setup("take_m")}*{setup("infl_used")}/(1+{setup("infl_used")}),"#,##0")&" less each month next year."',
        f'="To keep the same life next year, your groups would need about $"&TEXT(Setup!$H${GROUP_TOTAL},"#,##0")'
        f'&" a month (now: $"&TEXT(Setup!$F${GROUP_TOTAL},"#,##0")&")."',
    ]
    for k, s in enumerate(sentences):
        merged(ws, f"B{p+1+k}:M{p+1+k}", s, f=font(11), align=Alignment(vertical="center", indent=1))
        ws.row_dimensions[p + 1 + k].height = 20
    put(ws, f"B{p+6}", "Tax numbers are estimates to help you budget, not tax advice. Your pay stub is the final word.",
        f=font(9, italic=True, color=MUTED))


# ================================================================== REPORTS
def build_reports(wb, look):
    ws = wb.create_sheet("Reports")
    widths(ws, {"A": 2, "B": 24, "C": 15, **{col_letter(c): 12 for c in range(4, 17)}})
    merged(ws, "B1:J1", "Reports", f=font(22, True, look.primary))
    ws.row_dimensions[1].height = 36
    merged(ws, "B2:J2", "Scroll down: Month by month  ·  Each category by month  ·  "
                        "Day by day  ·  Paycheck by paycheck", f=font(11, color=MUTED))
    spend_plan = f'SUMIFS({G_MONTH},{G_KIND},"<>{KIND_SAVE}")'

    # ---- A. month by month: every month from Jan 2026 to Dec 2028 (rows 6-41), then one total per year
    a_first, a_last = 6, 5 + CAL_MONTHS
    section_bar(ws, 4, "B", "L", f"MONTH BY MONTH ({CAL_LABEL.upper()})", look, "A")
    header_row(ws, 5, [("B", "Month"), ("C", "Paydays"), ("D", "Expected pay"), ("E", "Actually received"),
                       ("F", "Bills & needs"), ("G", "Flexible"), ("H", "Saved & invested"), ("I", "Left over"),
                       ("J:L", "How it went")], look)
    for i in range(CAL_MONTHS):
        r = a_first + i
        nxt = f"EDATE($B{r},1)"
        as_calc(ws, f"B{r}", f"=EDATE({cal_start()},{i})", "mmmm yyyy", align="left", bold=True)
        as_calc(ws, f"C{r}", f"={paydays_in(f'$B{r}', nxt)}", "0", align="center")
        as_calc(ws, f"D{r}", f"={expected_pay(f'$B{r}', nxt)}", MONEY)
        as_calc(ws, f"E{r}", f"={received_in(f'$B{r}', nxt)}", MONEY)
        m = between(f"$B{r}", nxt)
        for col, kind in (("F", KIND_BILLS), ("G", KIND_FLEX), ("H", KIND_SAVE)):
            as_calc(ws, f"{col}{r}", f'=SUMIFS({SP_AMT},{SP_KIND},"{kind}",{m})', MONEY)
        # Left over uses money received if you log it, otherwise expected pay
        as_calc(ws, f"I{r}", f"=IF(E{r}>0,E{r},D{r})-F{r}-G{r}-H{r}", MONEY)
        ws.merge_cells(f"J{r}:L{r}")
        as_calc(ws, f"J{r}", f'=IF(COUNTIFS({SP_DATE},">="&$B{r},{SP_DATE},"<"&{nxt})=0,'
                             f'"Nothing logged yet",IF(I{r}<0,"Spent more than you made",'
                             f'IF(F{r}+G{r}>{spend_plan},"Over plan by "&TEXT(F{r}+G{r}-{spend_plan},"$#,##0"),'
                             f'"✔ On plan")))', align="left")
    status_colors(ws, f"J{a_first}:L{a_last}", f"$J{a_first}")
    this_month_highlight(ws, look, f"B{a_first}:I{a_last}", f"$B{a_first}")
    # one total row per year (2026 total, 2027 total, 2028 total)
    months_b = f"$B${a_first}:$B${a_last}"
    for k in range(CAL_YEARS):
        r = a_last + 1 + k
        in_year = f'{months_b},">="&{year_start(k)},{months_b},"<"&{year_start(k + 1)}'
        as_calc(ws, f"B{r}", f'=YEAR({year_start(k)})&" total"', align="left", bold=True)
        as_calc(ws, f"C{r}", f"=SUMIFS(C${a_first}:C${a_last},{in_year})", "0", bold=True, align="center")
        for col in "DEFGHI":
            as_calc(ws, f"{col}{r}", f"=SUMIFS({col}${a_first}:{col}${a_last},{in_year})", MONEY, bold=True)
        total_row(ws, look, r, "B", "L")
    a_end = a_last + CAL_YEARS
    red_if_negative(ws, f"I{a_first}:I{a_end}")
    put(ws, f"B{a_end + 1}", "Paid every 2 weeks? Two months a year have 3 paydays. Those are great months to save extra.",
        f=font(9, italic=True, color=MUTED))

    # ---- B. each category by month: Jan-Dec of the year you picked on the Dashboard
    b_top = a_end + 3
    section_bar(ws, b_top, "B", "P", "EACH CATEGORY BY MONTH (the year picked on the Dashboard)", look, "B")
    header_row(ws, b_top + 1, [("B", "Category"), ("C", "Group")], look)
    hr = b_top + 1                                         # header row with the 12 month names
    for j in range(12):
        col = col_letter(4 + j)
        put(ws, f"{col}{hr}", f"=DATE(YEAR(Dashboard!$U$2),{j + 1},1)", f=font(10, True, "FFFFFF"),
            fill=PatternFill("solid", fgColor=look.primary), fmt="mmm yy", align=Alignment(horizontal="center"))
    put(ws, f"P{hr}", "Total", f=font(10, True, "FFFFFF"), fill=PatternFill("solid", fgColor=look.primary),
        align=Alignment(horizontal="center"))
    cat_first = hr + 1
    for i in range(C_SLOTS):
        r, sr = cat_first + i, CAT_FIRST + i
        as_calc(ws, f"B{r}", f'=IF(Setup!$B${sr}="","",Setup!$B${sr})', align="left")
        as_calc(ws, f"C{r}", f'=IF(B{r}="","",Setup!$C${sr})', align="left", color=MUTED)
        for j in range(12):
            col = col_letter(4 + j)
            as_calc(ws, f"{col}{r}", f'=IF($B{r}="","",SUMIFS({SP_AMT},{SP_CAT},$B{r},'
                                     f'{between(f"{col}${hr}", f"EDATE({col}${hr},1)")}))', MONEY)
        as_calc(ws, f"P{r}", f'=IF($B{r}="","",SUM(D{r}:O{r}))', MONEY, bold=True)

    # ---- C. day by day for the Dashboard month
    top2 = cat_first + C_SLOTS + 2
    section_bar(ws, top2, "B", "G", "DAY BY DAY (month picked on the Dashboard)", look, "C")
    header_row(ws, top2 + 1, [("B", "Date"), ("C", "Spent that day"), ("D", "Flexible so far"),
                              ("E", "Even pace"), ("F:G", "Ahead or behind?")], look)
    flex_plan = f'SUMIFS({G_MONTH},{G_KIND},"{KIND_FLEX}")'
    for d in range(1, 32):
        r = top2 + 1 + d
        as_calc(ws, f"B{r}", f'=IF({d}<=DAY(EOMONTH(Dashboard!$U$2,0)),Dashboard!$U$2+{d}-1,"")', "ddd mmm d",
                align="left")
        as_calc(ws, f"C{r}", f'=IF(B{r}="","",SUMIFS({SP_AMT},{SP_DATE},B{r}))', MONEY)
        as_calc(ws, f"D{r}", f'=IF(B{r}="","",SUMIFS({SP_AMT},{SP_KIND},"{KIND_FLEX}",{SP_DATE},">="&Dashboard!$U$2,'
                             f'{SP_DATE},"<="&B{r}))', MONEY)
        as_calc(ws, f"E{r}", f'=IF(B{r}="","",{flex_plan}*{d}/DAY(EOMONTH(Dashboard!$U$2,0)))', MONEY)
        ws.merge_cells(f"F{r}:G{r}")
        as_calc(ws, f"F{r}", f'=IF(OR(B{r}="",N(E{r})=0),"",IF(D{r}>E{r},"Over pace by "&TEXT(D{r}-E{r},"$#,##0"),'
                             f'"On track ("&TEXT(E{r}-D{r},"$#,##0")&" under)"))', align="left")
    status_colors(ws, f"F{top2+2}:G{top2+32}", f"$F{top2+2}")
    # ---- D. paycheck by paycheck: every paycheck up to Dec 2028 (it is the longest list, so it goes last)
    top = top2 + 1 + 31 + 3
    section_bar(ws, top, "B", "L", f"PAYCHECK BY PAYCHECK (UP TO {CAL_LABEL.upper()[-8:]})", look, "D")
    header_row(ws, top + 1, [("B", "Paycheck #"), ("C", "Money lands"), ("D", "Covers until"), ("E", "Expected"),
                             ("F", "Received"), ("G", "Bills & needs"), ("H", "Flexible"), ("I", "Saved"),
                             ("J", "Left over"), ("K:L", "How it went")], look)
    per_year = setup("per_year")

    rows = CAL_YEARS * CONFIG["paycheck_rows"]                      # room for weekly pay for 3 years
    cal_end = f"EDATE({cal_start()},{CAL_MONTHS})"                   # Jan 1, 2029
    for k in range(1, rows + 1):
        r = top + 1 + k
        # blank once the paycheck would land after Dec 2028
        lands = f"INDEX({PAYDAYS},{k})"                                 # day paycheck k lands (Engine list)
        as_calc(ws, f"B{r}", f'=IF(OR({lands}="",N({lands})>={cal_end}),"",{k})', "0",
                align="center")
        as_calc(ws, f"C{r}", f'=IF(B{r}="","",{lands})', DATE_FMT)                      # day the money lands
        as_calc(ws, f"D{r}", f'=IF(B{r}="","",INDEX({PAYDAYS},{k + 1})-1)', DATE_FMT)    # day before the next one
        as_calc(ws, f"E{r}", f'=IF(B{r}="","",IF(B{r}=1,{FIRST_NET},{setup("per_check")}))', MONEY)
        p_range = f'{SP_DATE},">="&$C{r},{SP_DATE},"<="&$D{r}'
        as_calc(ws, f"F{r}", f'=IF(B{r}="","",{received_in(f"$C{r}", f"($D{r}+1)")})', MONEY)
        for col, kind in (("G", KIND_BILLS), ("H", KIND_FLEX), ("I", KIND_SAVE)):
            as_calc(ws, f"{col}{r}", f'=IF(B{r}="","",SUMIFS({SP_AMT},{SP_KIND},"{kind}",{p_range}))', MONEY)
        as_calc(ws, f"J{r}", f'=IF(B{r}="","",IF(F{r}>0,F{r},E{r})-G{r}-H{r}-I{r})', MONEY)
        ws.merge_cells(f"K{r}:L{r}")
        as_calc(ws, f"K{r}", f'=IF(B{r}="","",IF(C{r}>TODAY(),"Coming up",'
                             f'IF(AND(F{r}>0,F{r}<E{r}),"Partly paid: "&TEXT(E{r}-F{r},"$#,##0")&" still coming",'
                             f'IF(J{r}<0,"Overspent","✔ On track"))))', align="left")
    status_colors(ws, f"K{top+2}:L{top+1+rows}", f"$K{top+2}")
    red_if_negative(ws, f"J{top+2}:J{top+1+rows}")
    ws.conditional_formatting.add(f"K{top+2}:K{top+1+rows}", FormulaRule(
        formula=[f'ISNUMBER(SEARCH("partly",$K{top+2}))'], font=Font(color="8A5A00", bold=True),
        fill=PatternFill("solid", fgColor=AMBER_BG)))

    ws.freeze_panes = "A3"           # keep only the title rows frozen (no line through the sections)


# ================================================================== GOALS
SIGNED = '+$#,##0;-$#,##0;$0'          # shows +$120 / -$80, so ahead/behind reads at a glance
GOAL_FIRST = 12                        # goal rows 12-19
MONTH_FIRST = 24                       # scorecard rows 24-59 (Jan 2026 to Dec 2028), then one total per year


def build_goals(wb, look):
    """Goals tab: are you on track with your savings, and by how much?
    Top: 4 big cards for THIS calendar year (they start over every January). Middle: your own savings goals.
    Bottom: a month-by-month savings scorecard (bills paid, saved, vs goal, running total)."""
    ws = wb.create_sheet("Goals")
    n_goals = CONFIG["goal_slots"]
    goal_last = GOAL_FIRST + n_goals - 1
    m_last = MONTH_FIRST + CAL_MONTHS - 1
    # same width for every money/date column so the tab looks even
    widths(ws, {"A": 2, "B": 22, **{c: 14 for c in "CDEFHKLMN"}, "G": 20, "I": 15, "J": 10,
                "O": 13, "P": 13, "Q": 13, "S": 4, "U": 12})
    merged(ws, "B1:J1", "Goals", f=font(22, True, look.primary))
    ws.row_dimensions[1].height = 36
    merged(ws, "B2:Q2", "Are you on track with your savings, and by how much? Everything here fills in "
                        "from your Setup and Spending tabs.", f=font(11, color=MUTED))

    # Hidden helpers (column U) + category list for the "counts" dropdown (column S)
    #  U2 = Jan 1 of THIS year (the cards count this year only, so they start over every January)
    #  U5 = the first month that counts at all: the latest of Jan 2026, your first payday's month, and G9
    #  U4 = the first month that counts THIS year (the later of U2 and U5)
    #  U3 = how many months of this year count so far
    put(ws, "U2", "=DATE(YEAR(TODAY()),1,1)", fmt="mm/dd/yyyy")
    first = setup("first")
    put(ws, "U5", f'=MAX({cal_start()},IF(N({first})=0,0,DATE(YEAR({first}),MONTH({first}),1)),'
                  f'IF(N($G$9)=0,0,DATE(YEAR($G$9),MONTH($G$9),1)))', fmt="mm/dd/yyyy")
    put(ws, "U4", "=MAX(U2,U5)", fmt="mm/dd/yyyy")
    put(ws, "U3", "=MAX(0,MIN(12,(YEAR(TODAY())-YEAR(U4))*12+MONTH(TODAY())-MONTH(U4)+1))")
    put(ws, "S4", "(All savings)")
    for i in range(C_SLOTS):
        ws[f"S{5 + i}"] = f'=IFERROR(INDEX({C_NAME},MATCH({i + 1},{C_ALPHA},0)),"")'     # categories A to Z
    for col in "STU":
        ws.column_dimensions[col].hidden = True

    months = f"$B${MONTH_FIRST}:$B${m_last}"
    counted = f"$I${MONTH_FIRST}:$I${m_last}"
    goal_m = f"$J${MONTH_FIRST}:$J${m_last}"
    # SUMIFS piece: months of THIS year that count (on/after U4, before next January)
    since = f'{months},">="&$U$4,{months},"<"&EDATE($U$2,12)'
    tracked = f'COUNTIFS({since})'                      # how many months of this year you'll track
    goal_year = f'SUMIFS({goal_m},{since})'             # savings goal for the months you track

    # ---- 4 big cards (row 4 label, row 5 number, row 6 explanation)
    cards = [
        ("B:E", "SAVED THIS YEAR", f'=SUMIFS({counted},{months},"<="&TODAY(),{since})', MONEY0,
         '="Over "&$U$3&IF($U$3=1," month"," months")&"  ·  about "&TEXT(IF($U$3=0,0,B5/$U$3),"$#,##0")&" a month"'),
        ("F:I", "YOUR GOAL SO FAR", f'=SUMIFS({goal_m},{months},"<="&TODAY(),{since})', MONEY0,
         f'="Goal for "&YEAR($U$2)&": "&TEXT({goal_year},"$#,##0")'),
        ("J:M", "AHEAD OR BEHIND", "=B5-F5", SIGNED,
         f'=IF({goal_year}=0,"Add your pay and plan on Setup",IF(J5>=0,"✔ Ahead of your savings goal","Save about "&TEXT(-J5/MAX(1,{tracked}-$U$3),"$#,##0")'
         '&" extra a month to catch up"))'),
        ("N:Q", "ON PACE TO SAVE", f"=IF($U$3=0,0,B5/$U$3*{tracked})", MONEY0,
         f'="by "&TEXT(EDATE($U$2,12)-1,"mmm d, yyyy")&IF({goal_year}=0,"",IF(N5>={goal_year},"  ✔ beats your goal",'
         f'"  ·  "&TEXT({goal_year}-N5,"$#,##0")&" short of your goal"))'),
    ]
    big_cards(ws, look, cards)
    if look.style == "card":   # on the bold colored blocks, red text would be hard to read
        ws.conditional_formatting.add("J5", FormulaRule(formula=["$J$5<0"], font=Font(color=RED, bold=True, size=24)))

    # ---- one question: does leftover money count as saved?
    merged(ws, "B8:F8", "Count money left over at the end of each month as saved?", f=font(11, True),
           align=Alignment(vertical="center"))
    as_input(ws, "G8", "No")
    dropdown(ws, "G8", ["No", "Yes"])
    merged(ws, "H8:Q8", "Pick Yes if you leave unspent money in savings instead of moving it yourself. "
                        "It only counts once the month is over.", f=font(9, italic=True, color=MUTED),
           align=Alignment(vertical="center", wrap_text=True))
    ws.row_dimensions[8].height = 28
    merged(ws, "B9:F9", "Start tracking savings from (month)", f=font(11, True), align=Alignment(vertical="center"))
    as_input(ws, "G9", None, "mmm yyyy")
    date_picker(ws, "G9")
    merged(ws, "H9:Q9", "Started mid-year? Type the first day of the month you began (like 10/1/2026). "
                        "Earlier months won't count against you. Blank = from your first payday.",
           f=font(9, italic=True, color=MUTED), align=Alignment(vertical="center", wrap_text=True, indent=1))
    ws.row_dimensions[9].height = 28

    # ---- A. savings goals
    section_bar(ws, 10, "B", "Q", "YOUR SAVINGS GOALS", look, "A")
    header_row(ws, 11, [("B", "Goal"), ("C", "Target"), ("D", "Start\ndate"), ("E", "Target\ndate"),
                        ("F", "Saved at\nstart"), ("G", "Counts\nfrom"), ("H", "Saved\nso far"),
                        ("I", "Progress"), ("J", "% done"), ("K", "Should have\nby now"),
                        ("L", "Ahead (+)\nBehind (−)"), ("M", "Needed\nper month"),
                        ("N:Q", "How it's going")], look)
    ws.row_dimensions[11].height = 34          # every label fits on 2 lines
    for i in range(n_goals):
        r = GOAL_FIRST + i
        g = CONFIG["default_goals"][i] if i < len(CONFIG["default_goals"]) else None
        as_input(ws, f"B{r}", g["name"] if g else None)
        as_input(ws, f"C{r}", g["target"] if g else None, MONEY)
        as_input(ws, f"D{r}", "=$U$4" if g else None, DATE_FMT)            # example goals start when you start tracking
        as_input(ws, f"E{r}", "=EDATE($U$4,12)-1" if g else None, DATE_FMT)
        as_input(ws, f"F{r}", g["already"] if g else None, MONEY)
        as_input(ws, f"G{r}", g["counts"] if g else None)
        saved = (f'N(F{r})+IF(G{r}="(All savings)",SUMIFS({SP_AMT},{SP_KIND},"{KIND_SAVE}",{SP_DATE},">="&N(D{r})),'
                 f'IF(G{r}="",0,SUMIFS({SP_AMT},{SP_CAT},G{r},{SP_DATE},">="&N(D{r}))))')
        as_calc(ws, f"H{r}", f'=IF(B{r}="","",{saved})', MONEY, bold=True)
        bars = f"MAX(0,MIN(10,ROUND(H{r}/C{r}*10,0)))"
        as_calc(ws, f"I{r}", f'=IF(OR(B{r}="",N(C{r})=0),"",REPT("█",{bars})&REPT("░",10-{bars}))',
                align="left", color=look.group_color(i % 6) if look.style == "bold" else look.accent)
        as_calc(ws, f"J{r}", f'=IF(OR(B{r}="",N(C{r})=0),"",H{r}/C{r})', PCT)
        # straight line from "already saved" on the start date to the target on the target date
        as_calc(ws, f"K{r}", f'=IF(OR(B{r}="",N(D{r})=0,N(E{r})<=N(D{r})),"",IF(TODAY()<=D{r},N(F{r}),'
                             f'IF(TODAY()>=E{r},C{r},N(F{r})+(C{r}-N(F{r}))*(TODAY()-D{r})/(E{r}-D{r}))))', MONEY)
        as_calc(ws, f"L{r}", f'=IF(K{r}="","",H{r}-K{r})', SIGNED, bold=True)
        as_calc(ws, f"M{r}", f'=IF(OR(B{r}="",N(E{r})=0),"",IF(H{r}>=C{r},0,(C{r}-H{r})/'
                             f'MAX(1,(YEAR(E{r})-YEAR(TODAY()))*12+MONTH(E{r})-MONTH(TODAY()))))', MONEY)
        ws.merge_cells(f"N{r}:Q{r}")
        as_calc(ws, f"N{r}",
                f'=IF(B{r}="","",IF(H{r}>=C{r},"✔ Done! Goal reached",IF(K{r}="","Add a start and target date",'
                f'IF(TODAY()>E{r},"Missed the date: "&TEXT(C{r}-H{r},"$#,##0")&" short",'
                f'IF(L{r}>=0,"✔ On track (ahead by "&TEXT(L{r},"$#,##0")&")",'
                f'"Behind by "&TEXT(-L{r},"$#,##0")&"  ·  save "&TEXT(M{r},"$#,##0")&"/mo")))))',
                align="left")
    for r in range(GOAL_FIRST, goal_last + 1):
        ws.row_dimensions[r].height = 22
    dropdown(ws, f"G{GOAL_FIRST}:G{goal_last}", same_sheet_range=f"$S$4:$S${4 + C_SLOTS}")
    date_picker(ws, f"D{GOAL_FIRST}:E{goal_last}")
    status_colors(ws, f"N{GOAL_FIRST}:Q{goal_last}", f"$N{GOAL_FIRST}")
    for r in range(GOAL_FIRST, goal_last + 1):
        ws.conditional_formatting.add(f"L{r}", FormulaRule(formula=[f'AND(L{r}<>"",L{r}<0)'],
                                                           font=Font(color=RED, bold=True)))
    put(ws, f"B{goal_last + 1}", "How saving counts: on the Spending tab, log money you move to savings using a "
                                 "savings category (like 'Emergency fund'). Pick that category in 'Counts money logged to'.",
        f=font(9, italic=True, color=MUTED))

    # ---- B. savings scorecard, month by month
    top = 22
    section_bar(ws, top, "B", "Q", "SAVINGS SCORECARD, MONTH BY MONTH", look, "B")
    header_row(ws, top + 1, [("B", "Month"), ("C", "Money in"), ("D", "Bills\nexpected"), ("E", "Bills\npaid"),
                             ("F", "Flexible\nspent"), ("G", "Put into\nsavings"), ("H", "Left over"),
                             ("I", "Counted\nas saved"), ("J", "Savings\ngoal"), ("K", "Over (+)\nUnder (−)"),
                             ("L", "Savings\nrate"), ("M", "Saved\nthis year"), ("N", "Goal\nthis year"),
                             ("O:Q", "How it went")], look)
    ws.row_dimensions[top + 1].height = 34
    bills_goals = f'SUMIFS({C_GOAL},{C_KIND},"{KIND_BILLS}")'
    bills_plan = f'IF({bills_goals}>0,{bills_goals},SUMIFS({G_MONTH},{G_KIND},"{KIND_BILLS}"))'
    for i in range(CAL_MONTHS):
        r = MONTH_FIRST + i
        nxt = f"EDATE($B{r},1)"
        m = between(f"$B{r}", nxt)
        before = f"$B{r}<$U$5"                               # a month before you started tracking
        future = f"OR($B{r}>TODAY(),{before})"
        as_calc(ws, f"B{r}", f"=EDATE({cal_start()},{i})", "mmmm yyyy", align="left", bold=True)
        as_calc(ws, f"C{r}", f"=IF({received_in(f'$B{r}', nxt)}>0,{received_in(f'$B{r}', nxt)},"
                             f"{expected_pay(f'$B{r}', nxt)})", MONEY)
        as_calc(ws, f"D{r}", f'=IF({before},"",{bills_plan})', MONEY, color=MUTED)     # blank before you started
        for col, kind in (("E", KIND_BILLS), ("F", KIND_FLEX), ("G", KIND_SAVE)):
            as_calc(ws, f"{col}{r}", f'=SUMIFS({SP_AMT},{SP_KIND},"{kind}",{m})', MONEY)
        as_calc(ws, f"H{r}", f'=IF({future},"",C{r}-E{r}-F{r}-G{r})', MONEY)
        as_calc(ws, f"I{r}", f'=G{r}+IF(AND($G$8="Yes",{nxt}<=TODAY()),MAX(0,H{r}),0)', MONEY, bold=True)
        as_calc(ws, f"J{r}", f'=IF({before},"",SUMIFS({G_MONTH},{G_KIND},"{KIND_SAVE}"))', MONEY, color=MUTED)
        as_calc(ws, f"K{r}", f'=IF({future},"",I{r}-J{r})', SIGNED, bold=True)
        as_calc(ws, f"L{r}", f'=IF(OR({future},C{r}=0),"",I{r}/C{r})', PCT)
        # running totals that start over every January (only months you track)
        this_yr = f'{months},"<="&$B{r},{months},">="&MAX($U$5,DATE(YEAR($B{r}),1,1))'
        as_calc(ws, f"M{r}", f'=IF({future},"",SUMIFS({counted},{this_yr}))', MONEY)
        as_calc(ws, f"N{r}", f'=IF({future},"",SUMIFS({goal_m},{this_yr}))', MONEY)
        ws.row_dimensions[r].height = 22
        ws.merge_cells(f"O{r}:Q{r}")
        logged = f'COUNTIFS({SP_DATE},">="&$B{r},{SP_DATE},"<"&{nxt})'
        as_calc(ws, f"O{r}",
                f'=IF({before},"Before you started",IF({future},"Coming up",IF({logged}=0,"Nothing logged yet",'
                f'IF({nxt}>TODAY(),IF(E{r}<D{r},"Bills left: "&TEXT(D{r}-E{r},"$#,##0")&"  ·  ","✔ Bills paid  ·  ")'
                f'&IF(I{r}>=J{r},"savings goal met",TEXT(J{r}-I{r},"$#,##0")&" to go on savings"),'
                f'IF(I{r}>=J{r},"✔ Goal met (+"&TEXT(I{r}-J{r},"$#,##0")&")",'
                f'"Missed goal by "&TEXT(J{r}-I{r},"$#,##0"))))))', align="left")
        ws.conditional_formatting.add(f"K{r}", FormulaRule(formula=[f'AND(K{r}<>"",K{r}<0)'],
                                                           font=Font(color=RED, bold=True)))
    status_colors(ws, f"O{MONTH_FIRST}:Q{m_last}", f"$O{MONTH_FIRST}")
    this_month_highlight(ws, look, f"B{MONTH_FIRST}:N{m_last}", f"$B{MONTH_FIRST}")
    # one total row per year: 2026 total, 2027 total, 2028 total
    for k in range(CAL_YEARS):
        t = m_last + 1 + k
        in_year = f'{months},">="&{year_start(k)},{months},"<"&{year_start(k + 1)}'
        as_calc(ws, f"B{t}", f'=YEAR({year_start(k)})&" total"', align="left", bold=True)
        for col in "CDEFGHIJ":
            as_calc(ws, f"{col}{t}", f"=SUMIFS({col}${MONTH_FIRST}:{col}${m_last},{in_year})", MONEY, bold=True)
        as_calc(ws, f"K{t}", f"=SUMIFS(K${MONTH_FIRST}:K${m_last},{in_year})", SIGNED, bold=True)
        past = f'{in_year},{months},"<="&TODAY(),{months},">="&$U$5'      # months of that year that already count
        as_calc(ws, f"L{t}", f'=IF(SUMIFS(C${MONTH_FIRST}:C${m_last},{past})=0,"",'
                             f'SUMIFS({counted},{past})/SUMIFS(C${MONTH_FIRST}:C${m_last},{past}))', PCT, bold=True)
        total_row(ws, look, t, "B", "Q")
    t_end = m_last + CAL_YEARS
    red_if_negative(ws, f"H{MONTH_FIRST}:H{t_end}")
    put(ws, f"B{t_end + 1}", "Bills expected = the monthly goals you set for bill categories on Setup (or your bills group "
                             "plan if you didn't set any). Savings goal = your 'Saving & investing' groups' monthly plan. "
                             "This month counts as it goes.", f=font(9, italic=True, color=MUTED))
    put(ws, f"B{t_end + 2}", "Money in and Bills expected for future months are what you expect, so each year's total "
                             "shows your whole year.", f=font(9, italic=True, color=MUTED))
    ws.freeze_panes = "A3"


# ================================================================== SHARED: the 4 big number cards
def big_cards(ws, look, cards, row=4):
    """cards = [(span "B:D", LABEL, value formula, number format, subtext formula), ...]
    Same look on every tab: label / big number / small line, colored in bold style, soft in card style."""
    for i, (span, label, value, fmt, sub) in enumerate(cards):
        a, b = span.split(":")
        if look.style == "bold":
            bg, label_c, num_c, sub_c = look.group_color(i), "FFFFFF", "FFFFFF", "FFFFFF"
        else:
            bg, label_c, num_c, sub_c = look.soft, MUTED, look.primary, MUTED
        fill_range(ws, f"{a}{row}:{b}{row + 2}", PatternFill("solid", fgColor=bg))
        merged(ws, f"{a}{row}:{b}{row}", label, f=font(9, True, label_c),
               align=Alignment(horizontal="center", vertical="bottom"))
        merged(ws, f"{a}{row + 1}:{b}{row + 1}", value, f=font(24, True, num_c), fmt=fmt,
               align=Alignment(horizontal="center", vertical="center", shrink_to_fit=True))
        merged(ws, f"{a}{row + 2}:{b}{row + 2}", sub, f=font(9, color=sub_c),
               align=Alignment(horizontal="center", vertical="top", wrap_text=True))
        for r in range(row, row + 3):                       # white line BETWEEN cards only, so outer edges line up
            for c in range(ord(a) - 64, ord(b) - 63):
                ws.cell(row=r, column=c).border = Border(left=THICK_WHITE if (c == ord(a) - 64 and i > 0) else None,
                                                         right=THICK_WHITE if (c == ord(b) - 64 and i < len(cards) - 1) else None)
        if look.style == "card":
            outline(ws, f"{a}{row}:{b}{row + 2}", look.accent)
    ws.row_dimensions[row].height = 20
    ws.row_dimensions[row + 1].height = 40
    ws.row_dimensions[row + 2].height = 30


def tab_title(ws, look, title, subtitle, last_col="M"):
    merged(ws, f"B1:{last_col}1", title, f=font(22, True, look.primary), align=Alignment(vertical="center"))
    ws.row_dimensions[1].height = 36
    merged(ws, f"B2:{last_col}2", subtitle, f=font(11, color=MUTED), align=Alignment(vertical="center"))


# ================================================================== BILLS
BILL_HELP_FIRST = 5                     # hidden helper rows on the Bills tab (one per category slot)
SAFE_ROW = 12                           # "safe to spend until your next paycheck" (Bills!E12)
AFFORD_ROW = 13                         # "Can I afford it?" box (Bills!E13)
BILL_ROWS_FIRST = 19                    # visible bill list (section B)
BILL_LIST = CONFIG["bill_rows"]          # how many bills the list shows
BILL_ROWS_LAST = BILL_ROWS_FIRST - 1 + BILL_LIST
ARR_SLOTS = CONFIG["arrangement_slots"]
ARR_FIRST = BILL_ROWS_LAST + 6          # payment arrangements (section C)
ARR_LAST = ARR_FIRST + ARR_SLOTS - 1
PAYPLAN_FIRST = ARR_LAST + 6            # "which paycheck pays which bill" rows (section D, 6 of them)
ARR_FREQ = ["Weekly", "Every 2 weeks", "Monthly", "One time"]
YEARLY_SLOTS = 8


def bills_month():
    return "Dashboard!$U$2", "Dashboard!$U$3"   # the month picked on the Dashboard (start, next start)


def build_bills(wb, look):
    """Bills tab: every bill with a due day (Setup → Categories → Due day), whether it's paid,
    which paycheck pays it, and yearly costs to set money aside for."""
    ws = wb.create_sheet("Bills")
    n = C_SLOTS
    soon = CONFIG["bill_due_soon_days"]
    m0, m1 = bills_month()
    widths(ws, {"A": 2, "B": 30, "C": 14, "D": 14, "E": 16, "F": 14, "G": 15, "H": 13, "I": 13, "J": 13,
                "K": 2})
    tab_title(ws, look, "Bills", f'="Your bills for "&TEXT({m0},"mmmm yyyy")&". Change the month at the top of the Dashboard. '
                                 f'Add due days on Setup (Categories)."', "J")

    # ---- hidden helpers (columns AA:AI), one row per category slot
    H = BILL_HELP_FIRST
    hl = H + n - 1
    for i in range(n):
        r, cr = H + i, CAT_FIRST + i
        ws[f"AA{r}"] = f'=IF(OR(Setup!$B${cr}="",N(Setup!$E${cr})=0),"",Setup!$B${cr})'          # bill name
        ws[f"AB{r}"] = f'=IF(AA{r}="","",MIN(31,MAX(1,ROUND(Setup!$E${cr},0))))'                  # due day
        ws[f"AC{r}"] = f'=IF(AA{r}="","",DATE(YEAR({m0}),MONTH({m0}),MIN(AB{r},DAY({m1}-1))))'   # due date this month
        ws[f"AD{r}"] = f'=IF(AA{r}="",0,N(Setup!$D${cr}))'                                         # amount = monthly goal
        ws[f"AE{r}"] = f'=IF(AA{r}="",0,SUMIFS({SP_AMT},{SP_CAT},AA{r},{between(m0, m1)}))'        # paid this month
        ws[f"AF{r}"] = (f'=IF(AA{r}="","",IF(COUNTIF({PAYDAYS},"<="&AC{r})=0,"",'                   # paycheck that pays it
                        f'INDEX({PAYDAYS},COUNTIF({PAYDAYS},"<="&AC{r}))))')
        ws[f"AG{r}"] = f'=IF(AA{r}="","",AC{r}+ROW()/1000)'                                         # sort key: due date
        ws[f"AH{r}"] = f'=IF(AA{r}="","",COUNTIF($AG${H}:$AG${hl},"<"&AG{r})+1)'                    # 1 = due first
        ws[f"AI{r}"] = (f'=IF(AA{r}="","",IF(OR(AND(AD{r}>0,AE{r}>=AD{r}-0.005),AND(AD{r}=0,AE{r}>0)),"✔ Paid",'
                        f'IF(AE{r}>0,"Partly paid: "&TEXT(AD{r}-AE{r},"$#,##0")&" left",'
                        f'IF(Dashboard!$U$4=0,IF({m0}<TODAY(),"Not paid","Coming up"),'
                        f'IF(AC{r}<TODAY(),"Overdue by "&(TODAY()-AC{r})&IF(TODAY()-AC{r}=1," day"," days"),'
                        f'IF(AC{r}=TODAY(),"Due today",IF(AC{r}-TODAY()<={soon},"Due in "&(AC{r}-TODAY())'
                        f'&IF(AC{r}-TODAY()=1," day"," days"),"Due "&TEXT(AC{r},"mmm d"))))))))')
    for c in ("AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI"):
        ws.column_dimensions[c].hidden = True
    NAME, DUE, AMT, PAID, PAYS, RANK, STAT = (f"$A{c}${H}:$A{c}${hl}" for c in "ABDEFHI")
    DUE = f"$AC${H}:$AC${hl}"

    # ---- cards
    unpaid_rank = "$AJ$3"
    # AJ3 = rank of the first unpaid bill due today or later (helper)
    ws["AJ2"] = "first unpaid, due from today"
    ws["AJ3"] = (f'=IFERROR(SMALL($AJ${H}:$AJ${hl},1),"")')
    for i in range(n):
        r = H + i
        ws[f"AJ{r}"] = f'=IF(OR(AA{r}="",LEFT(AI{r},1)="✔",AC{r}<TODAY()),"",AH{r})'
    ws.column_dimensions["AJ"].hidden = True
    # window helpers for "before your next paycheck" (AN:AP), one row per category slot
    cur, nxt = "$AK$2", "$AK$3"
    for i in range(n):
        r = H + i
        d1 = f'DATE(YEAR({cur}),MONTH({cur}),MIN(AB{r},DAY(EOMONTH({cur},0))))'
        d2 = f'DATE(YEAR(EDATE({cur},1)),MONTH(EDATE({cur},1)),MIN(AB{r},DAY(EOMONTH({cur},1))))'
        ws[f"AN{r}"] = f'=IF(OR(AA{r}="",{cur}=""),"",IF({d1}<{cur},{d2},{d1}))'           # next due date on/after this paycheck
        ws[f"AO{r}"] = (f'=IF(AN{r}="",0,IF(AN{r}<{nxt},MAX(0,AD{r}-SUMIFS({SP_AMT},{SP_CAT},AA{r},'
                        f'{SP_DATE},">="&{cur})),0))')                                         # still to pay before next paycheck
        ws[f"AP{r}"] = f'=IF(AO{r}>0,AN{r}+ROW()/1000,"")'                                     # sort key for "first one due"
    for c in ("AK", "AN", "AO", "AP"):
        ws.column_dimensions[c].hidden = True
    nxt_name = f'INDEX({NAME},MATCH({unpaid_rank},$AH${H}:$AH${hl},0))'
    nxt_date = f'INDEX({DUE},MATCH({unpaid_rank},$AH${H}:$AH${hl},0))'
    big_cards(ws, look, [
        ("B:C", "BILLS THIS MONTH", f"=SUM({AMT})", MONEY0, f'=COUNT({DUE})&" bills with a due day"'),
        ("D:E", "PAID SO FAR", f"=SUMPRODUCT(({PAID}>={AMT})*{AMT}+({PAID}<{AMT})*{PAID})",
         MONEY0, f'=COUNTIF({STAT},"✔*")&" of "&COUNT({DUE})&" paid"'),
        ("F:G", "STILL TO PAY", "=MAX(0,B5-D5)", MONEY0,
         f'=IF(COUNTIF({STAT},"Overdue*")>0,COUNTIF({STAT},"Overdue*")&" overdue!","Nothing overdue ✔")'),
        ("H:J", "NEXT BILL", f'=IF(COUNT({DUE})=0,"No bills yet",IF({unpaid_rank}="","All set ✔",{nxt_name}))', None,
         f'=IF(COUNT({DUE})=0,"Add due days on Setup (Categories)",IF({unpaid_rank}="","No unpaid bills left this month",TEXT({nxt_date},"ddd, mmm d")&IF({nxt_date}=TODAY(),'
         f'"  (today!)","  ·  in "&({nxt_date}-TODAY())&" days")))'),
    ])
    ws["H5"].font = font(18, True, "FFFFFF" if look.style == "bold" else look.primary)
    ws.conditional_formatting.add("F6", FormulaRule(formula=[f'COUNTIF({STAT},"Overdue*")>0'],
                                                    font=Font(color="FFFFFF" if look.style == "bold" else RED, bold=True)))

    build_safe_to_spend(ws, look, H, hl)

    # ---- B. bill list, sorted by due date
    section_bar(ws, BILL_ROWS_FIRST - 3, "B", "J", "BILLS THIS MONTH (by due date)", look, "B")
    header_row(ws, BILL_ROWS_FIRST - 2, [("B", "Bill"), ("C", "Due"), ("D", "Amount"), ("E", "Paid"), ("F", "Left to pay"),
                        ("G", "Paid from\npaycheck on"), ("H:J", "Status")], look)
    ws.row_dimensions[BILL_ROWS_FIRST - 2].height = 34
    rows_last = BILL_ROWS_LAST
    put(ws, f"B{BILL_ROWS_FIRST - 1}", "Amount = the category's monthly goal on Setup. Log the payment on Spending and it shows as paid.",
        f=font(9, italic=True, color=MUTED))
    for k in range(1, BILL_LIST + 1):
        r = BILL_ROWS_FIRST - 1 + k
        pick = lambda col: f'INDEX($A{col}${H}:$A{col}${hl},MATCH({k},$AH${H}:$AH${hl},0))'
        as_calc(ws, f"B{r}", f'=IFERROR({pick("A")},"")', align="left", bold=True)
        as_calc(ws, f"C{r}", f'=IF(B{r}="","",{pick("C")})', "ddd mmm d", align="center")
        as_calc(ws, f"D{r}", f'=IF(B{r}="","",{pick("D")})', MONEY)
        as_calc(ws, f"E{r}", f'=IF(B{r}="","",{pick("E")})', MONEY)
        as_calc(ws, f"F{r}", f'=IF(B{r}="","",MAX(0,D{r}-E{r}))', MONEY)
        as_calc(ws, f"G{r}", f'=IF(B{r}="","",IF({pick("F")}="","Last month",{pick("F")}))', "mmm d", align="center")
        ws.merge_cells(f"H{r}:J{r}")
        as_calc(ws, f"H{r}", f'=IF(B{r}="","",{pick("I")})', align="left")
        ws.row_dimensions[r].height = 20
    status_colors(ws, f"H{BILL_ROWS_FIRST}:J{rows_last}", f"$H{BILL_ROWS_FIRST}")
    put(ws, f"B{rows_last + 1}", f'=IF(COUNT({DUE})=0,"No bills yet: on Setup → Categories, type a Due day (1–31) next to each bill.",'
                                 f'IF(COUNT({DUE})>{BILL_LIST},"+ "&(COUNT({DUE})-{BILL_LIST})&" more bills (the list shows the first {BILL_LIST})",""))',
        f=font(10, True, look.primary))

    build_arrangements(ws, look)

    # ---- D. which paycheck pays which bill (and payment arrangements)
    top = PAYPLAN_FIRST - 3
    section_bar(ws, top, "B", "J", "WHICH PAYCHECK PAYS WHICH BILL", look, "D")
    put(ws, f"B{top + 1}", "Each bill (and payment arrangement) is paid from the last paycheck before its due date.",
        f=font(9, italic=True, color=MUTED))
    header_row(ws, top + 2, [("B", "Paycheck"), ("C", "Take-home"), ("D", "Bills it pays"), ("E", "Left after\nbills"),
                             ("F:J", "Which bills")], look)
    ws.row_dimensions[top + 2].height = 34
    first_idx = f'(COUNTIF({PAYDAYS},"<"&{m0})+1)'
    for k in range(0, 6):                       # 0 = last paycheck of the month before
        r = top + 3 + k
        idx = f"({first_idx}+{k - 1})"
        in_month = f'COUNTIFS({PAYDAYS},">="&{m0},{PAYDAYS},"<"&{m1})'
        pd = (f'IF({idx}<1,"",INDEX({PAYDAYS},{idx}))' if k == 0
              else f'IF({k}>{in_month},"",INDEX({PAYDAYS},{idx}))')
        ws[f"AL{r}"] = f"={pd}"                                                 # hidden: this paycheck's date
        ws[f"AM{r}"] = (f'=IF(AL{r}="","",' + "&".join(
            f'IF(AF{H + i}=AL{r},AA{H + i}&", ","")' for i in range(n)) + "&" + "&".join(
            f'IF(AT{ARR_FIRST + j}=AL{r},B{ARR_FIRST + j}&" (plan), ","")' for j in range(ARR_SLOTS)) + ")")   # hidden: names joined
        as_calc(ws, f"B{r}", f'=IF(AL{r}="","",TEXT(AL{r},"ddd, mmm d")&IF({k}=0,"  (last month)",""))', align="left", bold=True)
        as_calc(ws, f"C{r}", f'=IF(AL{r}="","",IF({idx}=1,{FIRST_NET},{setup("per_check")}))', MONEY)
        as_calc(ws, f"D{r}", f'=IF(AL{r}="","",SUMIFS({AMT},$AF${H}:$AF${hl},AL{r})'
                             f'+SUMIFS($D${ARR_FIRST}:$D${ARR_LAST},$AT${ARR_FIRST}:$AT${ARR_LAST},AL{r}))', MONEY)
        as_calc(ws, f"E{r}", f'=IF(AL{r}="","",C{r}-D{r})', MONEY, bold=True)
        ws.merge_cells(f"F{r}:J{r}")
        as_calc(ws, f"F{r}", f'=IF(AL{r}="","",IF(LEN(AM{r})<3,"No bills",LEFT(AM{r},LEN(AM{r})-2)))', align="left")
        ws[f"F{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[r].height = 30
    pc_first, pc_last = top + 3, top + 8
    ws.conditional_formatting.add(f"E{pc_first}:E{pc_last}", FormulaRule(
        formula=[f'AND(ISNUMBER(E{pc_first}),E{pc_first}<0)'], font=Font(color=RED, bold=True),
        fill=PatternFill("solid", fgColor=RED_BG)))
    ws.conditional_formatting.add(f"E{pc_first}:E{pc_last}", FormulaRule(
        formula=[f'AND(ISNUMBER(E{pc_first}),E{pc_first}>=0)'], font=Font(color=GREEN, bold=True),
        fill=PatternFill("solid", fgColor=GREEN_BG)))
    for c in ("AL", "AM"):
        ws.column_dimensions[c].hidden = True
    put(ws, f"B{pc_last + 1}", "Red = that paycheck can't cover its bills. Move a due date, or set money aside from the paycheck before.",
        f=font(9, italic=True, color=MUTED))

    # ---- C. yearly costs: set a little aside each month
    yt = pc_last + 3
    section_bar(ws, yt, "B", "J", "YEARLY COSTS: SET A LITTLE ASIDE EACH MONTH", look, "E")
    put(ws, f"B{yt + 1}", "Car registration, holidays, renewals, insurance paid once a year... Split them into small monthly amounts.",
        f=font(9, italic=True, color=MUTED))
    header_row(ws, yt + 2, [("B", "What"), ("C", "Next due"), ("D", "Cost"), ("E", "Set aside\nso far"),
                            ("F", "Months\nleft"), ("G", "Set aside\nper month"), ("H", "Ready"), ("I:J", "Status")], look)
    ws.row_dimensions[yt + 2].height = 34
    examples = CONFIG.get("default_yearly_costs", [])
    yf = yt + 3
    for i in range(YEARLY_SLOTS):
        r = yf + i
        ex = examples[i] if i < len(examples) else None
        as_input(ws, f"B{r}", ex["name"] if ex else None)
        as_input(ws, f"C{r}", f'=EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),{ex["months_away"]})' if ex else None, DATE_FMT)
        as_input(ws, f"D{r}", ex["cost"] if ex else None, MONEY)
        as_input(ws, f"E{r}", 0 if ex else None, MONEY)
        months_left = f'MAX(1,(YEAR(C{r})-YEAR(TODAY()))*12+MONTH(C{r})-MONTH(TODAY()))'
        as_calc(ws, f"F{r}", f'=IF(OR(B{r}="",N(C{r})=0),"",{months_left})', "0", align="center")
        as_calc(ws, f"G{r}", f'=IF(F{r}="","",MAX(0,(N(D{r})-N(E{r}))/F{r}))', MONEY, bold=True)
        as_calc(ws, f"H{r}", f'=IF(OR(B{r}="",N(D{r})=0),"",MIN(1,N(E{r})/D{r}))', PCT, align="center")
        ws.merge_cells(f"I{r}:J{r}")
        as_calc(ws, f"I{r}", f'=IF(B{r}="","",IF(N(D{r})>0,IF(N(E{r})>=N(D{r}),"✔ Ready",'
                             f'IF(N(C{r})<TODAY(),"Date passed: update it",'
                             f'IF(F{r}<=1,"Due soon: "&TEXT(D{r}-E{r},"$#,##0")&" to go","On track"))),""))', align="left")
        ws.row_dimensions[r].height = 20
    date_picker(ws, f"C{yf}:C{yf + YEARLY_SLOTS - 1}")
    status_colors(ws, f"I{yf}:J{yf + YEARLY_SLOTS - 1}", f"$I{yf}")
    tot = yf + YEARLY_SLOTS
    put(ws, f"B{tot}", "Total per month", f=font(11, True))
    as_calc(ws, f"G{tot}", f"=SUM(G{yf}:G{tot - 1})", MONEY, bold=True)
    total_row(ws, look, tot, "B", "J")
    merged(ws, f"B{tot + 1}:J{tot + 1}",
           f'="Tip: on Setup, give a savings category (like \'Yearly costs\') a monthly goal of "&TEXT(G{tot},"$#,##0")'
           f'&", and log what you set aside there."',
           f=font(10, True, look.primary), align=Alignment(vertical="center", wrap_text=True))
    ws.row_dimensions[tot + 1].height = 30
    ws.freeze_panes = "A3"
    return {"pay_first": pc_first, "pay_last": pc_last}


def build_safe_to_spend(ws, look, H, hl):
    """Section A on Bills: money from the paycheck you're living on now, minus what you've spent since it
    landed, minus bills and payment-plan payments still due before the next paycheck = SAFE TO SPEND.
    Plus a 'Can I afford it?' box: type an amount, it tells you if your bills are still covered."""
    soon_days = f"MAX(1,$AK$3-TODAY())"
    k = f'COUNTIF({PAYDAYS},"<="&TODAY())'
    ws["AK1"] = "this paycheck / next paycheck / paycheck #"
    ws["AK2"] = f'=IF({k}=0,"",INDEX({PAYDAYS},{k}))'                                       # paycheck you're on now
    ws["AK3"] = (f'=IF($AK$2="","",IF(INDEX({PAYDAYS},{k}+1)="",$AK$2+ROUND(365/{setup("per_year")},0),'
                 f'INDEX({PAYDAYS},{k}+1)))')                                              # next paycheck
    ws["AK4"] = f"={k}"
    # first unpaid bill/payment before the next paycheck (for the "you'd be short for ..." message)
    ws["AK5"] = f'=IFERROR(SMALL($AP${H}:$AP${hl},1),"")'
    ws["AK6"] = f'=IF($AK$5="","",INDEX($AA${H}:$AA${hl},MATCH($AK$5,$AP${H}:$AP${hl},0)))'
    ws["AK7"] = f'=IF($AK$5="","",INT($AK$5))'

    section_bar(ws, 8, "B", "J", "BEFORE YOUR NEXT PAYCHECK", look, "A")
    win = f'{MI_DATE},">="&$AK$2,{MI_DATE},"<"&$AK$3'
    received = f'SUMIFS({MI_AMT},{win},{MI_FOR},"")'        # paychecks you logged (Money in with no "For")
    extra_in = f'SUMIFS({MI_AMT},{win},{MI_FOR},"<>")'      # extra money marked "For" a budget: added on top
    check = f'IF($AK$4=1,{FIRST_NET},{setup("per_check")})'
    window_bills = f"SUM($AO${H}:$AO${hl})+SUM($AS${ARR_FIRST}:$AS${ARR_LAST})"
    lines = [
        (9, '=IF($AK$2="","Money from your paycheck","Money that landed "&TEXT($AK$2,"ddd, mmm d"))',
         f'=IF($AK$2="",0,IF({received}>0,{received},{check})+{extra_in})',
         '=IF($AK$2="","Add your pay and first payday on Setup to see this.",'
         f'IF({received}>0,"What you logged under Money in since then.","Your take-home for one paycheck (from Setup)."))'),
        (10, '="Spent since then"', f'=IF($AK$2="",0,SUMIFS({SP_AMT},{SP_DATE},">="&$AK$2,{SP_DATE},"<"&$AK$3))',
         '="Everything on the Spending tab since that paycheck, bills too."'),
        (11, '=IF($AK$3="","Bills + payments still due","Bills + payments still due before "&TEXT($AK$3,"ddd, mmm d"))',
         f"={window_bills}",
         '=IF($AK$6="","Nothing else due before your next paycheck. ✔","First one: "&$AK$6&" on "&TEXT($AK$7,"mmm d")&".")'),
    ]
    for r, label, value, note in lines:
        merged(ws, f"B{r}:D{r}", label, f=font(11), align=Alignment(vertical="center", indent=1))
        as_calc(ws, f"E{r}", value, MONEY)
        merged(ws, f"F{r}:J{r}", note, f=font(9, italic=True, color=MUTED), align=Alignment(vertical="center", indent=1))
        ws.row_dimensions[r].height = 22
    r = SAFE_ROW
    merged(ws, f"B{r}:D{r}", '=IF($AK$3="","SAFE TO SPEND","SAFE TO SPEND until "&TEXT($AK$3,"ddd, mmm d"))',
           f=font(13, True), align=Alignment(vertical="center", indent=1))
    as_calc(ws, f"E{r}", "=E9-E10-E11", MONEY, bold=True, size=13)
    merged(ws, f"F{r}:J{r}", f'=IF($AK$2="","",IF(E{r}<0,"⚠ You\'re "&TEXT(-E{r},"$#,##0")&" short for bills due before your next paycheck.",'
                             f'"About "&TEXT(E{r}/{soon_days},"$#,##0")&" a day for the next "&{soon_days}&" days, and your bills are covered."))',
           f=font(10, True, GREEN), align=Alignment(vertical="center", indent=1))
    for c in "BCDE":
        ws[f"{c}{r}"].fill = PatternFill("solid", fgColor=look.soft)
    ws.row_dimensions[r].height = 26
    ws.conditional_formatting.add(f"E{r}:J{r}", FormulaRule(formula=[f"$E${r}<0"], font=Font(color=RED, bold=True)))
    r = AFFORD_ROW
    merged(ws, f"B{r}:D{r}", "Can I afford it? Type the amount →", f=font(11, True, look.primary),
           align=Alignment(vertical="center", indent=1))
    as_input(ws, f"E{r}", None, MONEY)
    merged(ws, f"F{r}:J{r}", f'=IF(N(E{r})=0,"Type what you want to spend (or plan to) and this checks your bills.",'
                             f'IF(N(E{r})<=E{SAFE_ROW},"✔ Yes. You\'d still have "&TEXT(E{SAFE_ROW}-E{r},"$#,##0")&" after your bills.",'
                             f'"✖ Not safe: you\'d be "&TEXT(E{r}-MAX(0,E{SAFE_ROW}),"$#,##0")&" short"&IF($AK$6="","",'
                             f'" for "&$AK$6&" (due "&TEXT($AK$7,"mmm d")&")")&"."))',
           f=font(10, True, look.primary), align=Alignment(vertical="center", indent=1, wrap_text=True))
    ws.row_dimensions[r].height = 30
    ws.conditional_formatting.add(f"F{r}:J{r}", FormulaRule(formula=[f'LEFT($F${r},1)="✔"'], stopIfTrue=True,
                                  font=Font(color=GREEN, bold=True), fill=PatternFill("solid", fgColor=GREEN_BG)))
    ws.conditional_formatting.add(f"F{r}:J{r}", FormulaRule(formula=[f'LEFT($F${r},1)="✖"'], stopIfTrue=True,
                                  font=Font(color=RED, bold=True), fill=PatternFill("solid", fgColor=RED_BG)))


def build_arrangements(ws, look):
    """Section C on Bills: payment arrangements (a plan you agreed to: $50 every 2 weeks until $400 is paid).
    Hidden helpers: AS = still due before your next paycheck, AT = paycheck that pays the next payment."""
    soon = CONFIG["bill_due_soon_days"]
    top = ARR_FIRST - 3
    section_bar(ws, top, "B", "J", "PAYMENT ARRANGEMENTS", look, "C")
    put(ws, f"B{top + 1}", "A payment plan you agreed to (rent catch-up, a medical bill, a card in hardship...). "
                           "Update 'Paid so far' when you pay.", f=font(9, italic=True, color=MUTED))
    header_row(ws, top + 2, [("B", "Who / what"), ("C", "Total owed"), ("D", "Each payment"), ("E", "How often"),
                             ("F", "First payment"), ("G", "Paid so far"), ("H", "Next payment"), ("I:J", "Status")], look)
    ws.row_dimensions[top + 2].height = 34
    for i in range(ARR_SLOTS):
        r = ARR_FIRST + i
        for col, fmt in (("B", None), ("C", MONEY), ("D", MONEY), ("E", None), ("F", DATE_FMT), ("G", MONEY)):
            as_input(ws, f"{col}{r}", None, fmt)
        ws.row_dimensions[r].height = 22

        def count_by(x):           # how many scheduled payments fall on or before date x
            return (f'IF(OR(N(F{r})=0,{x}<F{r}),0,IF(E{r}="Weekly",INT(({x}-F{r})/7)+1,IF(E{r}="Every 2 weeks",'
                    f'INT(({x}-F{r})/14)+1,IF(E{r}="Monthly",(YEAR({x})-YEAR(F{r}))*12+MONTH({x})-MONTH(F{r})'
                    f'+IF(DAY({x})>=DAY(F{r}),1,0),1))))')
        n_pay = f"ROUNDUP(N(C{r})/N(D{r}),0)"
        due_by = lambda x: f"MIN(N(C{r}),N(D{r})*MIN({n_pay},{count_by(x)}))"      # what you should have paid by date x
        k = f"(INT(N(G{r})/N(D{r}))+1)"                                          # number of the next payment
        nxt = (f'IF(E{r}="Weekly",F{r}+({k}-1)*7,IF(E{r}="Every 2 weeks",F{r}+({k}-1)*14,'
               f'IF(E{r}="Monthly",EDATE(F{r},{k}-1),F{r})))')
        ok = f'AND(B{r}<>"",N(C{r})>0,N(D{r})>0,N(F{r})>0)'
        as_calc(ws, f"H{r}", f'=IF(NOT({ok}),"",IF(N(G{r})>=N(C{r}),"",{nxt}))', "ddd mmm d", align="center")
        ws.merge_cells(f"I{r}:J{r}")
        late = f"{due_by('TODAY()')}-N(G{r})"
        as_calc(ws, f"I{r}", f'=IF(B{r}="","",IF(NOT({ok}),"Fill in the yellow boxes",IF(N(G{r})>=N(C{r}),"✔ Paid off",'
                             f'IF({late}>0.5,"Late: "&TEXT({late},"$#,##0")&" behind",'
                             f'IF(H{r}=TODAY(),"Due today",IF(H{r}-TODAY()<={soon},"Due in "&(H{r}-TODAY())&" days",'
                             f'"On track · "&ROUNDUP((N(C{r})-N(G{r}))/N(D{r}),0)&" left"))))))', align="left")
        ws[f"AS{r}"] = f'=IF(OR(NOT({ok}),$AK$2=""),0,MAX(0,{due_by("($AK$3-1)")}-N(G{r})))'   # due before next paycheck
        ws[f"AT{r}"] = (f'=IF(H{r}="","",IF(COUNTIF({PAYDAYS},"<="&H{r})=0,"",'
                        f'INDEX({PAYDAYS},COUNTIF({PAYDAYS},"<="&H{r}))))')                      # paycheck that pays it
    dropdown(ws, f"E{ARR_FIRST}:E{ARR_LAST}", sorted(ARR_FREQ))
    date_picker(ws, f"F{ARR_FIRST}:F{ARR_LAST}")
    status_colors(ws, f"I{ARR_FIRST}:J{ARR_LAST}", f"$I{ARR_FIRST}")
    for c in ("AS", "AT"):
        ws.column_dimensions[c].hidden = True
    put(ws, f"B{ARR_LAST + 1}", "Payments due before your next paycheck count in section A, and show in alerts.",
        f=font(9, italic=True, color=MUTED))

# ================================================================== DEBTS (payoff planner)
DEBT_FIRST = 12
DEBT_MONTHS = 360                       # simulate up to 30 years
GRID_COL = 40                           # hidden month-by-month grid starts at column AN
GRID_TOP = 30                           # row 30 = today (month 0); rows 31.. = month 1, 2, ...
METHODS = ["Snowball: smallest balance first", "Avalanche: highest interest first"]


def build_debts(wb, look):
    """Debt payoff planner. Everyone pays at least the minimums; the extra money goes to ONE debt at a time
    (smallest balance first = snowball, or highest interest first = avalanche). When a debt is paid off,
    its minimum rolls into the next one. A hidden month-by-month grid (columns AN onward) does the math."""
    ws = wb.create_sheet("Debts")
    n = CONFIG["debt_slots"]
    last = DEBT_FIRST + n - 1
    tot = last + 1
    widths(ws, {"A": 2, "B": 26, "C": 16, "D": 13, "E": 11, "F": 13, "G": 9, "H": 12, "I": 13, "J": 13,
                "K": 13, "L": 13, "M": 2})
    tab_title(ws, look, "Debts", "Pay off debt faster: list what you owe, add a little extra each month, pick a method.", "L")

    # ---- settings
    merged(ws, "B8:E8", "Extra you can pay each month (on top of minimums)", f=font(11, True),
           align=Alignment(vertical="center"))
    ws.merge_cells("F8:H8")
    as_input(ws, "F8", CONFIG["default_debt_extra"] or None, MONEY)
    merged(ws, "I8:L8", "Tip: this is the 'Extra debt payment' category on Setup. Even $25 helps.",
           f=font(9, italic=True, color=MUTED), align=Alignment(vertical="center", indent=1))
    merged(ws, "B9:E9", "Which debt gets the extra first?", f=font(11, True), align=Alignment(vertical="center"))
    ws.merge_cells("F9:H9")
    as_input(ws, "F9", METHODS[0])
    dropdown(ws, "F9", METHODS)
    merged(ws, "I9:L9", "Snowball = quick wins (feels good). Avalanche = least interest (saves the most).",
           f=font(9, italic=True, color=MUTED), align=Alignment(vertical="center", wrap_text=True, indent=1))
    ws.row_dimensions[8].height = 24
    ws.row_dimensions[9].height = 34

    # ---- A. your debts
    section_bar(ws, 10, "B", "L", "YOUR DEBTS", look, "A")
    header_row(ws, 11, [("B", "Debt"), ("C", "Type"), ("D", "Balance\nnow"), ("E", "Interest\n(APR)"),
                        ("F", "Minimum\npayment"), ("G", "Order"), ("H", "Months to\npay off"), ("I", "Paid off by"),
                        ("J", "Interest\nyou'll pay"), ("K:L", "Status")], look)
    ws.row_dimensions[11].height = 34
    # hidden per-debt helpers: AA = valid?, AB = sort key, AC = payoff rank (1 = gets the extra first)
    for i in range(n):
        r = DEBT_FIRST + i
        d = CONFIG["default_debts"][i] if i < len(CONFIG["default_debts"]) else None
        as_input(ws, f"B{r}", d["name"] if d else None)
        as_input(ws, f"C{r}", d["type"] if d else None)
        as_input(ws, f"D{r}", d["balance"] if d else None, MONEY)
        as_input(ws, f"E{r}", d["apr"] if d else None, "0.00%")
        as_input(ws, f"F{r}", d["min"] if d else None, MONEY)
        ws[f"AA{r}"] = f'=IF(AND(B{r}<>"",N(D{r})>0),1,0)'
        ws[f"AB{r}"] = (f'=IF(AA{r}=0,"",IF(LEFT($F$9,4)="Snow",N(D{r}),-N(E{r}))+ROW()/1000000)')
        ws[f"AC{r}"] = f'=IF(AA{r}=0,"",COUNTIF($AB${DEBT_FIRST}:$AB${last},"<"&AB{r})+1)'
        ws.row_dimensions[r].height = 22
    dropdown(ws, f"C{DEBT_FIRST}:C{last}", CONFIG["debt_types"])
    for c in ("AA", "AB", "AC"):
        ws.column_dimensions[c].hidden = True

    # ---- hidden grid. For payoff rank j (0-7): MIN paid, EXTRA paid, BALANCE left. Then L = extra pool,
    #      then minimums-only balances (to compare).
    gmin = lambda j: col_letter(GRID_COL + 3 * j)
    gext = lambda j: col_letter(GRID_COL + 3 * j + 1)
    gbal = lambda j: col_letter(GRID_COL + 3 * j + 2)
    pool = col_letter(GRID_COL + 3 * n)
    gmo = lambda j: col_letter(GRID_COL + 3 * n + 1 + j)
    ranks = f"$AC${DEBT_FIRST}:$AC${last}"
    P_SLOT, P_RATE, P_MIN, P_BAL0, R_MONTHS, R_INT, R_MO_MONTHS, R_MO_INT = 21, 22, 23, 24, 25, 26, 27, 28
    for j in range(n):
        b = gbal(j)
        ws[f"{b}{P_SLOT}"] = f'=IFERROR(MATCH({j + 1},{ranks},0),"")'
        ws[f"{b}{P_RATE}"] = f'=IF({b}{P_SLOT}="",0,N(INDEX($E${DEBT_FIRST}:$E${last},{b}{P_SLOT}))/12)'
        ws[f"{b}{P_MIN}"] = f'=IF({b}{P_SLOT}="",0,N(INDEX($F${DEBT_FIRST}:$F${last},{b}{P_SLOT})))'
        ws[f"{b}{P_BAL0}"] = f'=IF({b}{P_SLOT}="",0,N(INDEX($D${DEBT_FIRST}:$D${last},{b}{P_SLOT})))'
        ws[f"{b}{GRID_TOP}"] = f"={b}{P_BAL0}"
        ws[f"{gmo(j)}{GRID_TOP}"] = f"={b}{P_BAL0}"
    budget = f"${pool}${P_MIN}"
    ws[f"{pool}{P_MIN}"] = "=" + "+".join(f"{gbal(j)}{P_MIN}" for j in range(n)) + "+N($F$8)"   # monthly debt budget
    first_row, last_row = GRID_TOP + 1, GRID_TOP + DEBT_MONTHS
    for t in range(1, DEBT_MONTHS + 1):
        r, p = GRID_TOP + t, GRID_TOP + t - 1
        for j in range(n):
            b = gbal(j)
            due = f"{b}{p}*(1+{b}${P_RATE})"
            ws[f"{gmin(j)}{r}"] = f"=MIN({due},{b}${P_MIN})"
        ws[f"{pool}{r}"] = f"={budget}-(" + "+".join(f"{gmin(j)}{r}" for j in range(n)) + ")"
        for j in range(n):
            b = gbal(j)
            due = f"{b}{p}*(1+{b}${P_RATE})"
            used = "+".join(f"{gext(x)}{r}" for x in range(j)) or "0"
            ws[f"{gext(j)}{r}"] = f"=MAX(0,MIN({due}-{gmin(j)}{r},{pool}{r}-({used})))"
            ws[f"{b}{r}"] = f"=IF({due}-{gmin(j)}{r}-{gext(j)}{r}<0.005,0,{due}-{gmin(j)}{r}-{gext(j)}{r})"
            mo = gmo(j)
            ws[f"{mo}{r}"] = f"=IF({mo}{p}*(1+{b}${P_RATE})-{b}${P_MIN}<0.005,0,{mo}{p}*(1+{b}${P_RATE})-{b}${P_MIN})"
    # results per payoff rank
    for j in range(n):
        b, mo = gbal(j), gmo(j)
        col_b = f"{b}{first_row}:{b}{last_row}"
        col_mo = f"{mo}{first_row}:{mo}{last_row}"
        ws[f"{b}{R_MONTHS}"] = f'=IF({b}{P_BAL0}=0,0,IF(COUNTIF({col_b},">0")>={DEBT_MONTHS},"never",COUNTIF({col_b},">0")+1))'
        ws[f"{b}{R_INT}"] = f"={b}{P_RATE}*SUM({b}{GRID_TOP}:{b}{last_row - 1})"
        ws[f"{b}{R_MO_MONTHS}"] = f'=IF({b}{P_BAL0}=0,0,IF(COUNTIF({col_mo},">0")>={DEBT_MONTHS},"never",COUNTIF({col_mo},">0")+1))'
        ws[f"{b}{R_MO_INT}"] = f"={b}{P_RATE}*SUM({mo}{GRID_TOP}:{mo}{last_row - 1})"
    for c in range(GRID_COL, GRID_COL + 4 * n + 1):
        ws.column_dimensions[col_letter(c)].hidden = True
    res = lambda row: f"${gbal(0)}${row}:${gbal(n - 1)}${row}"
    pick = lambda row, rank: f"INDEX({res(row)},1,3*({rank}-1)+1)"
    start = "DATE(YEAR(TODAY()),MONTH(TODAY()),1)"
    for i in range(n):
        r = DEBT_FIRST + i
        rank = f"AC{r}"
        as_calc(ws, f"G{r}", f'=IF({rank}="","","#"&{rank})', align="center", bold=True)
        as_calc(ws, f"H{r}", f'=IF({rank}="","",{pick(R_MONTHS, rank)})', "0", align="center")
        as_calc(ws, f"I{r}", f'=IF(OR(H{r}="",H{r}="never"),"",EDATE({start},H{r}))', "mmm yyyy", align="center")
        as_calc(ws, f"J{r}", f'=IF(OR({rank}="",H{r}="never"),"",{pick(R_INT, rank)})', MONEY)
        ws.merge_cells(f"K{r}:L{r}")
        as_calc(ws, f"K{r}", f'=IF(B{r}="","",IF(N(D{r})=0,"✔ Paid off",IF(H{r}="never","Payment too low: never paid off",'
                             f'IF(N(F{r})<=N(D{r})*N(E{r})/12,"Minimum is too low","On track: "&H{r}&" months"))))',
                align="left")
    status_colors(ws, f"K{DEBT_FIRST}:L{last}", f"$K{DEBT_FIRST}")
    put(ws, f"B{tot}", "Total", f=font(11, True))
    as_calc(ws, f"D{tot}", f"=SUM(D{DEBT_FIRST}:D{last})", MONEY, bold=True)
    as_calc(ws, f"F{tot}", f"=SUM(F{DEBT_FIRST}:F{last})", MONEY, bold=True)
    as_calc(ws, f"J{tot}", f"=SUM(J{DEBT_FIRST}:J{last})", MONEY, bold=True)
    total_row(ws, look, tot, "B", "L")

    # ---- summary numbers (hidden helpers in AE)
    never = f'COUNTIF({res(R_MONTHS)},"never")>0'
    ws["AE21"] = f'=IF({never},"never",MAX({res(R_MONTHS)}))'                 # months to debt-free (your plan)
    ws["AE22"] = f'=IF(COUNTIF({res(R_MO_MONTHS)},"never")>0,"never",MAX({res(R_MO_MONTHS)}))'   # minimums only
    ws["AE23"] = f"=SUM({res(R_INT)})"
    ws["AE24"] = f"=SUM({res(R_MO_INT)})"
    ws.column_dimensions["AE"].hidden = True
    big_cards(ws, look, [
        ("B:C", "TOTAL DEBT NOW", f"=D{tot}", MONEY0, f'=COUNTIF(AA{DEBT_FIRST}:AA{last},1)&" debt(s)  ·  minimums "&TEXT(F{tot},"$#,##0")&"/mo"'),
        ("D:F", "DEBT-FREE BY", f'=IF(D{tot}=0,"No debt ✔",IF($AE$21="never","Never",TEXT(EDATE({start},$AE$21),"mmm yyyy")))', None,
         f'=IF(D{tot}=0,"Add your debts below to see your plan",IF($AE$21="never","Raise a minimum or the extra",$AE$21&" months, paying "&TEXT(F{tot}+N($F$8),"$#,##0")&"/mo"))'),
        ("G:I", "INTEREST YOU'LL PAY", "=IF($AE$21=\"never\",\"—\",$AE$23)", MONEY0, '="With your plan"'),
        ("J:L", "YOU SAVE", f'=IF(OR($AE$21="never",$AE$22="never"),"—",MAX(0,$AE$24-$AE$23))', MONEY0,
         f'=IF(OR($AE$21="never",$AE$22="never"),"vs. minimums only","in interest and "&MAX(0,$AE$22-$AE$21)&" months, vs. minimums only")'),
    ])
    for ref in ("D5",):
        ws[ref].font = font(20, True, "FFFFFF" if look.style == "bold" else look.primary)

    # ---- B. side by side
    sb = tot + 2
    section_bar(ws, sb, "B", "L", "MINIMUMS ONLY vs. YOUR PLAN", look, "B")
    header_row(ws, sb + 1, [("B", ""), ("C:D", "Debt-free in"), ("E:F", "Debt-free by"), ("G:H", "Total interest"),
                            ("I:L", "What it means")], look)
    rows = [("Minimums only", "$AE$22", "$AE$24", "Paying just the minimums."),
            ("Your plan", "$AE$21", "$AE$23", '="Minimums + "&TEXT(N($F$8),"$#,##0")&" extra, "&IF(LEFT($F$9,4)="Snow","smallest balance first","highest interest first")&"."')]
    for k, (label, months, interest, note) in enumerate(rows):
        r = sb + 2 + k
        put(ws, f"B{r}", label, f=font(11, True), border=GRID, align=Alignment(vertical="center", indent=1))
        ws.merge_cells(f"C{r}:D{r}")
        as_calc(ws, f"C{r}", f'=IF(D${tot}=0,"",IF({months}="never","Never",{months}&" months"))', align="center")
        ws.merge_cells(f"E{r}:F{r}")
        as_calc(ws, f"E{r}", f'=IF(OR(D${tot}=0,{months}="never"),"",TEXT(EDATE({start},{months}),"mmm yyyy"))', align="center")
        ws.merge_cells(f"G{r}:H{r}")
        as_calc(ws, f"G{r}", f'=IF(OR(D${tot}=0,{months}="never"),"",{interest})', MONEY)
        ws.merge_cells(f"I{r}:L{r}")
        as_calc(ws, f"I{r}", note if note.startswith("=") else note, align="left", color=MUTED)
        ws.row_dimensions[r].height = 22
    for c in "BCDEFGHIJKL":
        ws[f"{c}{sb + 3}"].fill = PatternFill("solid", fgColor=GREEN_BG)

    # ---- C. total left at the end of each year
    yb = sb + 6
    section_bar(ws, yb, "B", "L", "WHAT YOU OWE AT THE END OF EACH YEAR", look, "C")
    header_row(ws, yb + 1, [("B", "After"), ("C", "Your plan"), ("D", "Minimums only"), ("E:L", "Your plan (bar)")], look)
    for y in range(1, 11):
        r = yb + 1 + y
        gr = GRID_TOP + 12 * y
        put(ws, f"B{r}", f"{y} year{'s' if y > 1 else ''}", f=font(11, True), border=GRID, align=Alignment(vertical="center", indent=1))
        as_calc(ws, f"C{r}", "=" + "+".join(f"{gbal(j)}{gr}" for j in range(n)), MONEY0, bold=True)
        as_calc(ws, f"D{r}", "=" + "+".join(f"{gmo(j)}{gr}" for j in range(n)), MONEY0, color=MUTED)
        ws.merge_cells(f"E{r}:L{r}")
        bar = f"IF($D${tot}=0,0,ROUND(MIN(1,C{r}/$D${tot})*40,0))"
        as_calc(ws, f"E{r}", f'=IF($D${tot}=0,"",IF(C{r}=0,"✔ Debt-free",REPT("█",{bar})))', align="left", color=RED)
        ws[f"E{r}"].alignment = Alignment(horizontal="left", vertical="center")
    ws.conditional_formatting.add(f"E{yb + 2}:E{yb + 11}", FormulaRule(formula=[f'LEFT($E{yb + 2},1)="✔"'],
                                  font=Font(color=GREEN, bold=True)))
    put(ws, f"B{yb + 13}", "Estimates. Interest is figured monthly (APR ÷ 12). Card rates and minimums can change. Not financial advice.",
        f=font(9, italic=True, color=MUTED))
    ws.freeze_panes = "A3"


# ================================================================== TODAY (phone-friendly)
def build_today(wb, look):
    """A narrow tab made for a phone or tablet: everything you check most, in one column
    (fits a phone screen without zooming). It only READS other tabs, so it never needs typing,
    except the 'Can I afford it?' box."""
    ws = wb.create_sheet("Phone view")
    widths(ws, {"A": 1, "B": 25, "C": 15, "D": 3})
    ws.column_dimensions["D"].hidden = True             # hidden: status text that colors each row
    put(ws, "B1", "Phone view", f=font(20, True, look.primary))
    ws.row_dimensions[1].height = 32
    put(ws, "B2", '=TEXT(TODAY(),"dddd, mmm d")', f=font(11, color=MUTED))
    ws.merge_cells("B4:C4")
    as_calc(ws, "B4", "=Dashboard!$B$3", align="left", bold=True)
    ws["B4"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    ws.row_dimensions[4].height = 48
    for formula, fg, bg in ((f"N({setup('pay')})=0", BLUE, BLUE_BG), ("Dashboard!$Z$1>0", RED, RED_BG),
                            ("Dashboard!$Y$1>0", "8A5A00", AMBER_BG), ("TRUE", GREEN, GREEN_BG)):
        ws.conditional_formatting.add("B4:C4", FormulaRule(formula=[formula], stopIfTrue=True,
                                      font=Font(color=fg, bold=True), fill=PatternFill("solid", fgColor=bg)))

    def section(row, text):
        fill_range(ws, f"B{row}:C{row}", PatternFill("solid", fgColor=look.primary))
        merged(ws, f"B{row}:C{row}", text, f=font(10, True, "FFFFFF"), align=Alignment(vertical="center", indent=1))
        ws.row_dimensions[row].height = 22

    def line(row, label, value, fmt=MONEY0, big=False):
        put(ws, f"B{row}", label, f=font(13 if big else 11, big), align=Alignment(vertical="center", indent=1))
        as_calc(ws, f"C{row}", value, fmt, bold=big, size=16 if big else 11)
        ws.row_dimensions[row].height = 30 if big else 22

    # ---- safe to spend
    section(6, "SAFE TO SPEND")
    line(7, '=IF(Bills!$AK$3="","Until next payday","Until "&TEXT(Bills!$AK$3,"ddd, mmm d"))', f"=Bills!$E${SAFE_ROW}", MONEY0, True)
    ws.conditional_formatting.add("C7", FormulaRule(formula=["$C$7<0"], font=Font(color=RED, bold=True, size=16)))
    line(8, "About per day", f'=IF(Bills!$AK$3="",0,MAX(0,Bills!$E${SAFE_ROW})/MAX(1,Bills!$AK$3-TODAY()))')
    put(ws, "B9", "Can I afford it? →", f=font(11, True, look.primary), align=Alignment(vertical="center", indent=1))
    as_input(ws, "C9", None, MONEY)
    ws.row_dimensions[9].height = 26
    ws.merge_cells("B10:C10")
    as_calc(ws, "B10", f'=IF(N(C9)=0,"Type an amount to check your bills.",IF(N(C9)<=Bills!$E${SAFE_ROW},'
                       f'"✔ Yes. You\'d still have "&TEXT(Bills!$E${SAFE_ROW}-C9,"$#,##0")&" after bills.",'
                       f'"✖ Not safe: "&TEXT(C9-MAX(0,Bills!$E${SAFE_ROW}),"$#,##0")&" short"&IF(Bills!$AK$6="","",'
                       f'" for "&Bills!$AK$6)&"."))', align="left")
    ws["B10"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    ws.row_dimensions[10].height = 34
    ws.conditional_formatting.add("B10:C10", FormulaRule(formula=['LEFT($B$10,1)="✔"'], stopIfTrue=True,
                                  font=Font(color=GREEN, bold=True), fill=PatternFill("solid", fgColor=GREEN_BG)))
    ws.conditional_formatting.add("B10:C10", FormulaRule(formula=['LEFT($B$10,1)="✖"'], stopIfTrue=True,
                                  font=Font(color=RED, bold=True), fill=PatternFill("solid", fgColor=RED_BG)))

    # ---- this month
    section(12, '="THIS MONTH  ·  "&UPPER(TEXT(Dashboard!$U$2,"mmmm"))')
    line(13, "Pay this month", "=Dashboard!$B$5")
    line(14, "Spent so far", "=Dashboard!$E$5")
    line(15, "Left this month", "=Dashboard!$H$5")
    line(16, "You can spend / day", "=Dashboard!$K$5")
    red_if_negative(ws, "C15")

    # ---- groups
    section(18, "YOUR GROUPS: LEFT TO SPEND")
    for i in range(G):
        r = 19 + i
        put(ws, f"B{r}", f"=Dashboard!$B${11 + i}", f=font(11, True), align=Alignment(vertical="center", indent=1))
        as_calc(ws, f"C{r}", f'=IF(Dashboard!$B${11 + i}="","",Dashboard!$F${11 + i})', MONEY0)
        ws[f"D{r}"] = f"=Dashboard!$K${11 + i}"
        ws.row_dimensions[r].height = 22
    status_colors(ws, f"B19:C{18 + G}", "$D19")

    # ---- bills
    b0 = 20 + G
    section(b0, "BILLS THIS MONTH")
    for k in range(6):
        r, src = b0 + 1 + k, BILL_ROWS_FIRST + k
        put(ws, f"B{r}", f'=IF(Bills!$B${src}="","",Bills!$B${src}&"  ·  "&TEXT(Bills!$C${src},"mmm d"))',
            f=font(11), align=Alignment(vertical="center", indent=1))
        as_calc(ws, f"C{r}", f'=IF(Bills!$B${src}="","",IF(LEFT(Bills!$H${src},1)="✔","Paid",Bills!$F${src}))', MONEY0)
        ws[f"D{r}"] = f"=Bills!$H${src}"
        ws.row_dimensions[r].height = 22
    status_colors(ws, f"B{b0 + 1}:C{b0 + 6}", f"$D{b0 + 1}")
    put(ws, f"B{b0 + 7}", f'=IF(COUNT(Bills!$C${BILL_ROWS_FIRST}:$C${BILL_ROWS_LAST})=0,"Add due days on Setup to see bills here.","")',
        f=font(9, italic=True, color=MUTED), align=Alignment(vertical="center", indent=1))

    # ---- alerts
    a0 = b0 + 9
    section(a0, '="ALERTS ("&Dashboard!$Y$1&")"')
    for k in range(1, 6):
        r = a0 + k
        ws.merge_cells(f"B{r}:C{r}")
        as_calc(ws, f"B{r}", f"=Dashboard!$O${10 + k}", align="left")
        ws[f"B{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws[f"B{r}"].font = font(10, True)
        ws[f"D{r}"] = f"=IFERROR(INDEX(Dashboard!$Y$2:$Y$400,MATCH({k},Dashboard!$Z$2:$Z$400,0)),0)"
        ws.row_dimensions[r].height = 36
    for formula, fg, bg in ((f"$D{a0 + 1}=1", RED, RED_BG), (f"$D{a0 + 1}=2", "8A5A00", AMBER_BG),
                            (f"LEN($B{a0 + 1})>0", GREEN, GREEN_BG)):
        ws.conditional_formatting.add(f"B{a0 + 1}:C{a0 + 5}", FormulaRule(formula=[formula], stopIfTrue=True,
                                      font=Font(color=fg, bold=True), fill=PatternFill("solid", fgColor=bg)))
    merged(ws, f"B{a0 + 7}:C{a0 + 8}", "Made for phones: one column, no zooming. On a computer, use the Dashboard. Log purchases on Spending.",
           f=font(9, italic=True, color=MUTED), align=Alignment(vertical="top", wrap_text=True, indent=1))


# ================================================================== ALERTS (on the Dashboard)
def build_alerts(ws, look, month, flex_plan, flex_spent):
    """Overspending alerts for the month picked on the Dashboard.
    Hidden columns X:Z hold one row per POSSIBLE alert: X = the message ("" if it isn't happening),
    Y = 1 red / 2 yellow, Z = 1, 2, 3... numbering only the real ones. The visible list shows alert #1, #2...
    with INDEX/MATCH, so there are no blank gaps. Red ones are listed first."""
    almost, buffer, slots = CONFIG["alert_almost_pct"], CONFIG["alert_pace_buffer"], CONFIG["alert_slots"]
    money = lambda x: f'TEXT({x},"$#,##0")'
    cands = []                                    # (message formula, severity)
    # 1. spent more than you'll get this month
    cands.append((f'IF(AND(E5>0,H5<0),"Overspent: you\'ve spent "&{money("-H5")}&" more than you expect to get this month.","")', 1))
    # 2. a bills/flexible group is over its plan (and that costs your savings). OK'd on Spending = "Noted" (yellow)
    noted = []
    for i in range(G):
        r, sr = 11 + i, GROUP_FIRST + i
        key = f'B{r}&"|"&TEXT($U$2,"yyyymm")'
        okd = f'COUNTIF({SP_OKGRP},{key})>0'
        over = f'AND(B{r}<>"",Setup!$C${sr}<>"{KIND_SAVE}",N(D{r})>0,N(E{r})>N(D{r}))'
        cands.append((f'IF(AND({over},NOT({okd})),'
                      f'"Over budget: "&B{r}&" is "&{money(f"E{r}-D{r}")}&" over. That\'s money your savings goals don\'t get.","")', 1))
        noted.append((f'IF(AND({over},{okd}),"Noted: "&B{r}&" is "&{money(f"E{r}-D{r}")}&" over. You OK\'d it ("'
                       f'&INDEX({SP_OK},MATCH({key},{SP_OKGRP},0))&").","")', 2))
    # 2a. spending since payday leaves you short for bills due before the next paycheck (today, any month)
    cands.append((f'IF(AND(Bills!$AK$2<>"",Bills!$E${SAFE_ROW}<0),"Short: spending since your "&TEXT(Bills!$AK$2,"mmm d")'
                  f'&" paycheck leaves you "&TEXT(-Bills!$E${SAFE_ROW},"$#,##0")&" short for bills due before "&TEXT(Bills!$AK$3,"mmm d")&".","")', 1))
    # 2a2. a payment arrangement is behind
    for j in range(ARR_SLOTS):
        r = ARR_FIRST + j
        cands.append((f'IF(LEFT(Bills!$I${r},4)="Late","Late: your "&Bills!$B${r}&" payment plan is "&MID(Bills!$I${r},7,40)&".","")', 1))
    # 2b. a bill is overdue (Bills tab, this month only)
    for i in range(C_SLOTS):
        h = BILL_HELP_FIRST + i
        cands.append((f'IF(AND($U$4>0,LEFT(Bills!$AI${h},7)="Overdue"),"Overdue: "&Bills!$AA${h}&" was due "'
                       f'&TEXT(Bills!$AC${h},"mmm d")&IF(Bills!$AD${h}>0," ("&TEXT(MAX(0,Bills!$AD${h}-Bills!$AE${h}),"$#,##0")'
                       f'&" left to pay).","."),"")', 1))
    # 2c. a paycheck can't cover the bills it has to pay
    for k in range(6):
        r = PAYPLAN_FIRST + k
        cands.append((f'IF(AND(ISNUMBER(Bills!$E${r}),Bills!$E${r}<0,N(Bills!$AL${r})>=TODAY()),"Short: your "'
                       f'&TEXT(Bills!$AL${r},"mmm d")&" paycheck is "&TEXT(-Bills!$E${r},"$#,##0")&" short for the bills it pays.","")', 1))
    # 3. a category is over its own monthly goal (after any extra money in / last month's leftover covering it)
    for i in range(C_SLOTS):
        cr = CAT_FIRST + i
        name = f"Setup!$B${cr}"
        spent = f'SUMIFS({SP_AMT},{SP_CAT},{name},{month})'
        prev = f'SUMIFS({SP_AMT},{SP_CAT},{name},{SP_DATE},">="&EDATE($U$2,-1),{SP_DATE},"<"&$U$2)'
        budget = (f'(N(Setup!$D${cr})+SUMIFS({MI_AMT},{MI_FOR},{name},{MI_DATE},">="&$U$2,{MI_DATE},"<"&$U$3)'
                  f'+IF(COUNTIFS({SP_COVER},"{COVER_LEFTOVER}",{SP_KEY},{name},{month})>0,MAX(0,N(Setup!$D${cr})-{prev}),0))')
        key = f'{name}&"|"&TEXT($U$2,"yyyymm")'
        okd = f'COUNTIF({SP_OKCAT},{key})>0'
        over = f'AND({name}<>"",N(Setup!$D${cr})>0,{spent}>{budget})'
        cands.append((f'IF(AND({over},NOT({okd})),"Over goal: "&{name}&" is "&{money(f"{spent}-{budget}")}&" over its "'
                      f'&{money(budget)}&" budget.","")', 1))
        noted.append((f'IF(AND({over},{okd}),"Noted: "&{name}&" is "&{money(f"{spent}-{budget}")}&" over. You OK\'d it ("'
                       f'&INDEX({SP_OK},MATCH({key},{SP_OKCAT},0))&").","")', 2))
    cands.extend(noted)                       # "Noted" overages come first among the yellow ones
    # 4. a group is almost out
    for i in range(G):
        r, sr = 11 + i, GROUP_FIRST + i
        cands.append((f'IF(AND(B{r}<>"",Setup!$C${sr}<>"{KIND_SAVE}",N(D{r})>0,N(E{r})>={almost}*N(D{r}),N(E{r})<=N(D{r})),'
                      f'"Heads up: "&B{r}&" is "&TEXT(E{r}/D{r},"0%")&" used, only "&{money(f"F{r}")}&" left.","")', 2))
    # 4b. a bill is due soon (Bills tab)
    for i in range(C_SLOTS):
        h = BILL_HELP_FIRST + i
        cands.append((f'IF(AND($U$4>0,LEFT(Bills!$AI${h},3)="Due"),"Heads up: "&Bills!$AA${h}&" is due "'
                       f'&IF(Bills!$AC${h}=TODAY(),"today",TEXT(Bills!$AC${h},"ddd, mmm d"))'
                       f'&IF(Bills!$AD${h}>0," ("&TEXT(MAX(0,Bills!$AD${h}-Bills!$AE${h}),"$#,##0")&")","")&".","")', 2))
    # 4c. a payment arrangement payment is due soon
    for j in range(ARR_SLOTS):
        r = ARR_FIRST + j
        cands.append((f'IF(LEFT(Bills!$I${r},3)="Due","Heads up: "&TEXT(Bills!$D${r},"$#,##0")&" for your "&Bills!$B${r}'
                       f'&" payment plan is due "&IF(Bills!$H${r}=TODAY(),"today",TEXT(Bills!$H${r},"ddd, mmm d"))&".","")', 2))
    # 5. flexible spending is running ahead of the calendar (this month only)
    used, gone = f"({flex_spent})/({flex_plan})", "(($U$3-$U$2-$U$4+1)/($U$3-$U$2))"
    cands.append((f'IF($U$4=0,"",IF(({flex_plan})<=0,"",IF(AND({used}>{gone}+{buffer},{used}<=1),'
                  f'"Heads up: you\'ve used "&TEXT({used},"0%")&" of your flexible money with "&TEXT({gone},"0%")'
                  f'&" of the month gone. "&IF(N(K5)<1,"Try to pause extras until next month.","Aim for "&TEXT(K5,"$#,##0")&"/day."),"")))', 2))
    # 6. not enough left to finish this month's savings
    to_save = (f'MAX(0,SUMIFS($D$11:$D$16,Setup!$C${GROUP_FIRST}:$C${GROUP_LAST},"{KIND_SAVE}")'
               f'-SUMIFS($E$11:$E$16,Setup!$C${GROUP_FIRST}:$C${GROUP_LAST},"{KIND_SAVE}"))')
    cands.append((f'IF(AND($U$4>0,H5>=0,{to_save}>H5),"Heads up: only "&{money("H5")}&" left, but "&{money(to_save)}'
                  f'&" still to save this month. Hold off on extras.","")', 2))
    # 7. a savings goal is falling behind (from the Goals tab). Grace period: not in the goal's first month.
    for i in range(CONFIG["goal_slots"]):
        gr = GOAL_FIRST + i
        cands.append((f'IF(AND(Goals!$B${gr}<>"",ISNUMBER(Goals!$L${gr}),N(Goals!$L${gr})<=-1,N(Goals!$E${gr})>=TODAY(),N(Goals!$D${gr})<DATE(YEAR(TODAY()),MONTH(TODAY()),1)),'
                      f'"Heads up: "&Goals!$B${gr}&" is behind by "&{money(f"-Goals!$L${gr}")}&". Save about "'
                      f'&{money(f"N(Goals!$M${gr})")}&"/month to catch up.","")', 2))

    first, last = 2, 1 + len(cands)
    for k, (formula, sev) in enumerate(cands):
        r = first + k
        ws[f"X{r}"] = "=" + formula
        ws[f"Y{r}"] = sev
        ws[f"Z{r}"] = f'=IF(X{r}="","",SUMPRODUCT(--(LEN($X${first}:X{r})>0)))'
    X, Y, Z = f"$X${first}:$X${last}", f"$Y${first}:$Y${last}", f"$Z${first}:$Z${last}"
    ws["X1"] = "Alerts (every possible one)"
    ws["Y1"] = f"=SUMPRODUCT(--(LEN({X})>0))"                  # how many alerts
    ws["Z1"] = f"=SUMPRODUCT((LEN({X})>0)*({Y}=1))"            # how many are red
    for col in "XYZ":
        ws.column_dimensions[col].hidden = True

    # ---- row 3: one banner across the whole Dashboard (cards + alerts), most important alert first
    ws.merge_cells("B3:R3")
    as_calc(ws, "B3", f'=IF(N({setup("pay")})=0,"👋 Start here: type your pay and first payday on the Setup tab.",IF($Y$1=0,"✔ No alerts: your spending is on track this month.",'
                      f'IF($Y$1=1,"1 alert: ",$Y$1&" alerts: ")&INDEX({X},MATCH(1,{Z},0))))', align="left", bold=True)
    ws["B3"].alignment = Alignment(horizontal="left", vertical="center", indent=1, shrink_to_fit=True)
    ws.row_dimensions[3].height = 26
    for formula, fg, bg in ((f"N({setup('pay')})=0", BLUE, BLUE_BG), ("$Z$1>0", RED, RED_BG), ("$Y$1>0", "8A5A00", AMBER_BG), ("TRUE", GREEN, GREEN_BG)):
        ws.conditional_formatting.add("B3:R3", FormulaRule(formula=[formula], stopIfTrue=True,
                                      font=Font(color=fg, bold=True), fill=PatternFill("solid", fgColor=bg)))

    # ---- rows 4-6: an ALERTS card, the 5th card, on exactly the same rows as the other four
    for col, w in (("N", 2), ("O", 16), ("P", 16), ("Q", 16), ("R", 16), ("S", 2)):
        ws.column_dimensions[col].width = w
    bold = look.style == "bold"
    fill_range(ws, "O4:R6", PatternFill("solid", fgColor="2E7D32" if bold else GREEN_BG))
    ink = "FFFFFF" if bold else GREEN
    merged(ws, "O4:R4", "ALERTS", f=font(9, True, ink), align=Alignment(horizontal="center", vertical="bottom"))
    merged(ws, "O5:R5", f'=IF($Y$1=0,"All clear",$Y$1)', f=font(24, True, ink),
           align=Alignment(horizontal="center", vertical="center", shrink_to_fit=True))
    merged(ws, "O6:R6", f'=IF($Y$1=0,"Nothing to worry about ✔",IF($Z$1>0,$Z$1&" over budget","")'
                        f'&IF(AND($Z$1>0,$Y$1-$Z$1>0),"  ·  ","")&IF($Y$1-$Z$1>0,($Y$1-$Z$1)&" heads up","")&"  ·  list below")',
           f=font(9, color=ink), align=Alignment(horizontal="center", vertical="top", wrap_text=True))
    for formula, solid, soft, txt in (("$Z$1>0", RED, RED_BG, RED), ("$Y$1>0", "B7791F", AMBER_BG, "8A5A00")):
        ws.conditional_formatting.add("O4:R6", FormulaRule(formula=[formula], stopIfTrue=True,
            font=Font(color="FFFFFF" if bold else txt, bold=True),
            fill=PatternFill("solid", fgColor=solid if bold else soft)))
    if look.style == "card":
        outline(ws, "O4:R6", look.accent)

    # ---- rows 9-17: the full list, one alert per row, lined up with the groups table
    section_bar(ws, 9, "O", "R", "ALERTS THIS MONTH", look)
    header_row(ws, 10, [("O:R", "Most important first")], look)
    first_row = 11
    for k in range(1, 7):                          # rows 11-16 = the 6 group rows next to it
        r = first_row + k - 1
        ws.merge_cells(f"O{r}:R{r}")
        c = ws[f"O{r}"]
        empty = 'IF($Y$1=0,"✔ Nothing to worry about.","")' if k == 1 else '""'
        c.value = f'=IFERROR(INDEX({X},MATCH({k},{Z},0)),{empty})'
        c.font = font(10)
        c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        for col in "OPQR":
            ws[f"{col}{r}"].border = UNDER
        ws.row_dimensions[r].height = 30          # room for 2 lines; the group row beside it gets the same height
    sev = f"IFERROR(INDEX({Y},MATCH(ROW()-{first_row - 1},{Z},0)),0)"     # 1 red, 2 yellow, 0 = empty row
    for formula, fg, bg in ((f"{sev}=1", RED, RED_BG), (f"{sev}=2", "8A5A00", AMBER_BG),
                            (f'LEN($O{first_row})>0', GREEN, GREEN_BG)):
        ws.conditional_formatting.add(f"O{first_row}:R16", FormulaRule(formula=[formula], stopIfTrue=True,
                                      font=Font(color=fg, bold=True), fill=PatternFill("solid", fgColor=bg)))
    merged(ws, "O17:R17", f'=IF($Y$1>6,"+ "&($Y$1-6)&" more: fix the ones above first","Red = over budget  ·  Yellow = heads up")',
           f=font(9, italic=True, color=MUTED), align=Alignment(horizontal="left", vertical="center", indent=1))
    total_row(ws, look, 17, "O", "R")
    ws["O17"].font = font(9, True, MUTED)


# ================================================================== ACCOUNTS (HYSA + investing)
ACC_FIRST = 12                                   # account rows 12-19
ACCOUNTS_INPUTS = {}                             # filled while building (for the importer list)
YEARS = [1, 2, 3, 4, 5, 10, 15, 20, 25, 30]      # growth table years (hidden helper columns W..AF)


def build_accounts(wb, look):
    """Accounts tab: high-yield savings + investing accounts (SoFi, Fidelity, Robinhood...).
    What you have now, what it earns each month, and what it could grow to with your monthly deposits."""
    ws = wb.create_sheet("Accounts")
    n = CONFIG["account_slots"]
    last = ACC_FIRST + n - 1
    tot = last + 1
    widths(ws, {"A": 2, "B": 29, "C": 21, "D": 24, "E": 13, "F": 12, "G": 10, "H": 12, "I": 17,
                "J": 13, "K": 12, "L": 13, "M": 13, "N": 13, "O": 14, "P": 13, "Q": 13, "S": 4, "U": 12})
    merged(ws, "B1:J1", "Accounts", f=font(22, True, look.primary))
    ws.row_dimensions[1].height = 36
    merged(ws, "B2:Q2", "Your high-yield savings and investing accounts: what you have, what it earns, "
                        "and what it could grow to.", f=font(11, color=MUTED))

    # hidden helpers: category list (S), national average savings rate (U5), inflation in use (U6)
    put(ws, "S4", "")
    for i in range(C_SLOTS):
        ws[f"S{5 + i}"] = f'=IFERROR(INDEX({C_NAME},MATCH({i + 1},{C_ALPHA},0)),"")'     # categories A to Z
    put(ws, "U5", CONFIG["national_avg_savings_rate"], fmt="0.00%")
    put(ws, "U6", f"={setup('infl_used')}", fmt="0.0%")
    for col in "STU":
        ws.column_dimensions[col].hidden = True

    hysa = CONFIG["account_types"][0]
    rng_ = lambda c: f"${c}${ACC_FIRST}:${c}${last}"
    # ---- 4 big cards
    cards = [
        ("B:E", "IN SAVINGS & INVESTING NOW", f"=J{tot}", MONEY0,
         f'="Across "&COUNTA({rng_("B")})&" account(s)  ·  adding "&TEXT(H{tot},"$#,##0")&" a month"'),
        ("F:I", "EARNING PER MONTH", f"=K{tot}", MONEY0,
         f'="About "&TEXT(K{tot}*12,"$#,##0")&" a year from interest + growth"'),
        ("J:M", "IN 10 YEARS", f"=N{tot}", MONEY0,
         f'="You put in "&TEXT(J{tot}+H{tot}*120,"$#,##0")&"  ·  growth "&TEXT(N{tot}-J{tot}-H{tot}*120,"$#,##0")'
         f'&IF($G$8="Yes","  (today\'s dollars)","")'),
        ("N:Q", "HIGH-YIELD BONUS / YEAR",
         f'=SUMPRODUCT(({rng_("C")}="{hysa}")*{rng_("J")}*({rng_("G")}-$U$5))', MONEY0,
         '="More than a regular savings account at the "&TEXT($U$5,"0.00%")&" national average"'),
    ]
    big_cards(ws, look, cards)

    merged(ws, "B8:F8", "Show future amounts in today's dollars (after inflation)?", f=font(11, True),
           align=Alignment(vertical="center"))
    as_input(ws, "G8", "No")
    dropdown(ws, "G8", ["No", "Yes"])
    merged(ws, "H8:Q8", "Yes = what the money will really buy, after prices go up (uses the inflation rate on Setup).",
           f=font(9, italic=True, color=MUTED), align=Alignment(vertical="center", wrap_text=True, indent=1))
    ws.row_dimensions[8].height = 28

    # ---- A. your accounts
    section_bar(ws, 10, "B", "Q", "YOUR ACCOUNTS", look, "A")
    header_row(ws, 11, [("B", "Account"), ("C", "Type"), ("D", "Where\n(bank / app)"), ("E", "Balance"),
                        ("F", "Balance\nas of"), ("G", "Rate\nper year"), ("H", "Adding\nper month"),
                        ("I", "Deposits\nlogged to"), ("J", "Now\n(estimate)"), ("K", "Earns\nper month"),
                        ("L", "In 1 year"), ("M", "In 5 years"), ("N", "In 10 years"), ("O", "In 30 years"),
                        ("P:Q", "Good to know")], look)
    ws.row_dimensions[11].height = 34
    year_cols = [col_letter(23 + k) for k in range(len(YEARS))]          # W, X, ... AF
    for k, y in enumerate(YEARS):
        put(ws, f"{year_cols[k]}11", y, f=font(8, color=MUTED))
    notes = {CONFIG["account_types"][0]: "FDIC-insured. Rate can change.",
             CONFIG["account_types"][1]: "Safe, earns very little.",
             CONFIG["account_types"][2]: "Can go up and down.",
             CONFIG["account_types"][3]: "Penalty if taken out early."}
    for i in range(n):
        r = ACC_FIRST + i
        a = CONFIG["default_accounts"][i] if i < len(CONFIG["default_accounts"]) else None
        as_input(ws, f"B{r}", a["name"] if a else None)
        as_input(ws, f"C{r}", a["type"] if a else None)
        as_input(ws, f"D{r}", a["where"] if a else None)
        as_input(ws, f"E{r}", 0 if a else None, MONEY)
        as_input(ws, f"F{r}", None, DATE_FMT)
        as_input(ws, f"G{r}", a["rate"] if a else None, "0.00%")
        as_input(ws, f"H{r}", a["deposit"] if a else None, MONEY)
        as_input(ws, f"I{r}", a["fed_by"] if a else None)
        since = f"IF(N(F{r})=0,TODAY(),F{r})"
        yrs = f"MAX(0,TODAY()-{since})/365"
        logged = f'IF(I{r}="",0,SUMIFS({SP_AMT},{SP_CAT},I{r},{SP_DATE},">"&{since}))'
        # balance grows at its rate since the "as of" date, plus deposits you logged since then
        as_calc(ws, f"J{r}", f'=IF(B{r}="",0,N(E{r})*(1+N(G{r}))^({yrs})+{logged})', MONEY, bold=True)
        as_calc(ws, f"K{r}", f'=IF(B{r}="",0,J{r}*((1+N(G{r}))^(1/12)-1))', MONEY)
        # monthly rate; in today's dollars the growth is reduced by inflation
        rate = f'IF($G$8="Yes",(1+N(G{r}))/(1+$U$6)-1,N(G{r}))'
        monthly = f"((1+{rate})^(1/12)-1)"
        for k, y in enumerate(YEARS):
            put(ws, f"{year_cols[k]}{r}", f'=IF(B{r}="",0,FV({monthly},{12 * y},-N(H{r}),-J{r}))', fmt=MONEY0,
                f=font(8, color=MUTED))
        for col, y in (("L", 1), ("M", 5), ("N", 10), ("O", 30)):
            as_calc(ws, f"{col}{r}", f"={year_cols[YEARS.index(y)]}{r}", MONEY0)
        ws.merge_cells(f"P{r}:Q{r}")
        type_note = "".join(f'IF(C{r}="{t}","{msg}",' for t, msg in notes.items())
        as_calc(ws, f"P{r}", f'=IF(B{r}="","",{type_note}""{")" * len(notes)})', align="left", color=MUTED)
        ws.row_dimensions[r].height = 22
    dropdown(ws, f"C{ACC_FIRST}:C{last}", CONFIG["account_types"])
    dropdown(ws, f"I{ACC_FIRST}:I{last}", same_sheet_range=f"$S$4:$S${4 + C_SLOTS}")
    date_picker(ws, f"F{ACC_FIRST}:F{last}")
    put(ws, f"B{tot}", "Total", f=font(11, True))
    for col in "EHJKLMNO":
        as_calc(ws, f"{col}{tot}", f"=SUM({col}{ACC_FIRST}:{col}{last})", MONEY0 if col in "LMNO" else MONEY, bold=True)
    for k in range(len(YEARS)):
        c = year_cols[k]
        put(ws, f"{c}{tot}", f"=SUM({c}{ACC_FIRST}:{c}{last})", fmt=MONEY0, f=font(8, color=MUTED))
    total_row(ws, look, tot, "B", "Q")
    put(ws, f"B{tot + 1}", "Type your real balance + the date you checked it. Moving money in? Log it on Spending with the "
                           "category in 'Deposits logged to', and 'Now' adds it.", f=font(9, italic=True, color=MUTED))

    # ---- B. how it grows
    top = tot + 3
    section_bar(ws, top, "B", "Q", "HOW YOUR MONEY GROWS (ALL ACCOUNTS)", look, "B")
    header_row(ws, top + 1, [("B", "After"), ("C", "You put in"), ("D", "Worth"), ("E", "Growth"),
                             ("F:J", "Growth vs. what you put in")], look)
    for k, y in enumerate(YEARS):
        r = top + 2 + k
        c = year_cols[k]
        as_calc(ws, f"B{r}", f"{y} year{'s' if y > 1 else ''}", align="left", bold=True)
        as_calc(ws, f"C{r}", f"=J{tot}+H{tot}*{12 * y}", MONEY0)
        as_calc(ws, f"D{r}", f"={c}{tot}", MONEY0, bold=True)
        as_calc(ws, f"E{r}", f"=D{r}-C{r}", MONEY0, color=GREEN)
        ws.merge_cells(f"F{r}:J{r}")
        bar = f"MAX(0,MIN(30,ROUND(IF(D{r}=0,0,C{r}/D{r})*30,0)))"
        as_calc(ws, f"F{r}", f'=IF(D{r}<=0,"",REPT("█",{bar})&REPT("▒",30-{bar}))', align="left",
                color=look.primary)
    put(ws, f"B{top + 3 + len(YEARS)}", "█ = money you put in   ▒ = growth (interest + investment returns).",
        f=font(9, italic=True, color=MUTED))

    # ---- C. plain-words facts
    f0 = top + 5 + len(YEARS)
    section_bar(ws, f0, "B", "Q", "GOOD TO KNOW", look, "C")
    facts = [
        f'="High-yield savings: top accounts pay about 4% to 5% a year right now; the national average is only "&TEXT($U$5,"0.00%")&". Check your bank app for your real rate."',
        "=\"Savings accounts are usually FDIC-insured (up to $250,000 per bank). Great for an emergency fund. Rates can change any time.\"",
        f'="Investing (stocks, index funds) can lose money. The S&P 500 has averaged about {CONFIG["sp500_avg_return"]:.1%} a year since 1928 (about {CONFIG["sp500_real_return"]:.1%} after inflation), with big ups and downs."',
        "=\"Many people plan with 6% to 7% for long-term investing to be careful. These are estimates, not promises or financial advice.\"",
    ]
    for k, text in enumerate(facts):
        merged(ws, f"B{f0 + 1 + k}:Q{f0 + 1 + k}", text, f=font(10), align=Alignment(vertical="center", indent=1, wrap_text=True))
        ws.row_dimensions[f0 + 1 + k].height = 20
    # ---- D. net worth: what you own minus what you owe
    nw = f0 + 6
    section_bar(ws, nw, "B", "Q", "NET WORTH: WHAT YOU OWN MINUS WHAT YOU OWE", look, "D")
    debt_total = f"Debts!$D${DEBT_FIRST + CONFIG['debt_slots']}"
    lines = [
        (nw + 1, "Savings & investing (your accounts above)", f"=J{tot}", False, "Automatic."),
        (nw + 2, "Other things you own (car value, cash at home...)", None, True, "Optional. A rough guess is fine."),
        (nw + 3, "Debts (from the Debts tab)", f"={debt_total}", False, "Automatic."),
        (nw + 4, "Other money you owe (not on the Debts tab)", None, True, "Optional. Type it as a positive number."),
    ]
    for r, label, value, typed, tip in lines:
        merged(ws, f"B{r}:E{r}", label, f=font(11), align=Alignment(vertical="center"))
        if typed:
            as_input(ws, f"F{r}", value, MONEY)
        else:
            as_calc(ws, f"F{r}", value, MONEY if r == nw + 1 else '"−"$#,##0.00;"+"$#,##0.00;$0.00')
        merged(ws, f"G{r}:Q{r}", tip, f=font(9, italic=True, color=MUTED), align=Alignment(vertical="center", indent=1))
        ws.row_dimensions[r].height = 21
    ws[f"F{nw + 4}"].number_format = '"−"$#,##0.00;"+"$#,##0.00;$0.00'
    r = nw + 5
    merged(ws, f"B{r}:E{r}", "NET WORTH", f=font(13, True), align=Alignment(vertical="center"))
    as_calc(ws, f"F{r}", f"=F{nw + 1}+N(F{nw + 2})-N(F{nw + 3})-N(F{nw + 4})", MONEY, bold=True, size=13)
    merged(ws, f"G{r}:Q{r}", f'=IF(AND(F{nw + 1}=0,N(F{nw + 2})=0,N(F{nw + 3})=0,N(F{nw + 4})=0),"Add your accounts above (and debts on the Debts tab) to see this.",IF(F{r}>=0,"✔ You own more than you owe.","You owe more than you own right now. '
                             f'Paying down debt raises this every month."))',
           f=font(10, True, look.primary), align=Alignment(vertical="center"))
    for c in "BCDEF":
        ws[f"{c}{r}"].fill = PatternFill("solid", fgColor=look.soft)
    ws.row_dimensions[r].height = 26
    ws.conditional_formatting.add(f"F{r}", FormulaRule(formula=[f"$F${r}<0"], font=Font(color=RED, bold=True, size=13)))
    ws.conditional_formatting.add(f"G{r}", FormulaRule(formula=[f"$F${r}<0"], font=Font(color=RED, bold=True)))

    # monthly log: type it, or Budget Tools → "Save this month's net worth"
    lt = nw + 7
    header_row(ws, lt, [("B", "Month"), ("C", "Net worth"), ("D", "Change"), ("E:Q", "Net worth over time")], look)
    put(ws, f"B{lt - 1}", "Write it down once a month (or Budget Tools → Save this month's net worth) to watch it grow.",
        f=font(9, italic=True, color=MUTED))
    log_first, log_last = lt + 1, lt + CAL_MONTHS                   # Jan 2026 ... Dec 2028
    biggest = f"MAX(1,MAX($C${log_first}:$C${log_last}),-MIN($C${log_first}:$C${log_last}))"
    for i in range(CAL_MONTHS):
        r = log_first + i
        as_calc(ws, f"B{r}", f"=EDATE({cal_start()},{i})", "mmmm yyyy", align="left", bold=True)
        as_input(ws, f"C{r}", None, MONEY)
        as_calc(ws, f"D{r}", "" if i == 0 else f'=IF(OR(C{r}="",C{r - 1}=""),"",C{r}-C{r - 1})', '+$#,##0;-$#,##0;$0')
        ws.merge_cells(f"E{r}:Q{r}")
        as_calc(ws, f"E{r}", f'=IF(C{r}="","",REPT("█",MAX(1,ROUND(ABS(C{r})/{biggest}*40,0))))', align="left", color=GREEN)
        ws[f"E{r}"].alignment = Alignment(horizontal="left", vertical="center")
    ws.conditional_formatting.add(f"E{log_first}:E{log_last}", FormulaRule(formula=[f"N($C{log_first})<0"], font=Font(color=RED)))
    ws.conditional_formatting.add(f"D{log_first}:D{log_last}", FormulaRule(formula=[f"N($D{log_first})<0"], font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(f"D{log_first}:D{log_last}", FormulaRule(formula=[f"N($D{log_first})>0"], font=Font(color=GREEN, bold=True)))
    this_month_highlight(ws, look, f"B{log_first}:D{log_last}", f"$B{log_first}")
    ACCOUNTS_INPUTS["net_worth"] = (f"F{nw + 2}", f"F{nw + 4}", f"C{log_first}:C{log_last}")
    for c in year_cols:
        ws.column_dimensions[c].hidden = True
    ws.freeze_panes = "A3"


# ================================================================== HOW TO USE (the guide tab)
# Each section: (title, [(label, text), ...]). Plain words, one action per step.
GUIDE = [
    ("SET IT UP ONCE (about 5 minutes, on the Setup tab)", [
        ("Step 1", "How often do you get paid? Pick Weekly, Biweekly (every 2 weeks), Semi-monthly (twice a month) or Monthly."),
        ("Step 2", "First paycheck, if it was smaller: started in the middle of a pay period, so your very first check was "
                   "only part of a full one? Type that check here. Was it a full check? Leave it blank."),
        ("Step 3", "Full paycheck: type ONE normal paycheck before taxes (the 'gross' line on your pay stub). "
                   "Not two checks added together, and not the whole month."),
        ("Step 4", "When was your first payday? Pick the payday printed on your FIRST pay stub, even if it was a "
                   "small or half check. Every later payday is counted from this date."),
        ("Step 5", "Paid twice a month on set days, like the 15th & 30th? Pick those days. Then pick what happens when "
                   "payday lands on a weekend or a bank holiday like Christmas (most jobs pay the business day before)."),
        ("Step 6", "Does your bank pay you early? Chime, SoFi and many credit unions pay 1 or 2 days early. Pick that "
                   "here, and leave the payday in Step 4 as the date on your pay stub."),
        ("Step 7", "Taxes: leave 'Estimate for me' and pick your filing status and state rate. Already typed your amounts "
                   "AFTER taxes (what hits your bank)? Pick your own rate and type 0%."),
        ("Step 8", "Pick a budget plan, then look at your groups and categories. Type a monthly goal next to any category "
                   "you want to watch, and a due day (1 to 31) next to each bill."),
    ]),
    ("EVERY TIME MONEY COMES IN (Spending tab, right side: MONEY IN)", [
        ("Date received", "The day it landed in your bank. Type it like 8/28/2026 or click the box for a calendar. "
                          "A real date lines up on the RIGHT side of the box. If it sits on the left, it's text and won't count."),
        ("From", "Who paid you: your job, a side gig, a gift..."),
        ("Amount", "What actually hit your bank, after taxes. Two deposits on the same day (like a split paycheck)? "
                   "Type each one on its own row."),
        ("For (optional)", "Leave this BLANK for paychecks. Only pick a budget if the money is meant for one thing "
                           "(like $50 birthday money for Fun). It then adds to that budget for the month."),
        ("Sort by date", "Budget Tools → Sort Money in by date. It also sorts itself every time you open the sheet."),
    ]),
    ("EVERY TIME YOU SPEND (Spending tab, left side)", [
        ("Date", "The day you spent it."),
        ("What was it?", "A few words, like 'Publix' or 'Phone bill'."),
        ("Category", "Pick from the list. The category decides where it counts (see below)."),
        ("Amount", "How much you spent."),
        ("Budget check", "Fills in by itself and tells you how that budget is doing after this purchase."),
        ("Over budget?", "Pick a reason under OK'd? (Planned it, Emergency...) or pick what covers it under Cover with. "
                         "The red warning turns into a calm note."),
        ("Saving money?", "Moving money to savings counts as a row too: pick a savings category (like 'Emergency fund')."),
        ("Got a refund?", "Type it here as a minus amount (like -30) in the same category you bought it in. Don't put "
                          "refunds under Money in: that counts them as pay."),
        ("Sort by date", "Budget Tools → Sort Spending by date. It also sorts itself every time you open the sheet."),
    ]),
    ("WHERE EACH NUMBER COMES FROM (you never type in these)", [
        ("Expected", "What your plan says you should get, from the Setup tab: your first (smaller) check for paycheck #1, "
                     "then your full paycheck after estimated taxes."),
        ("Received", "The Money in rows whose Date received falls in that paycheck's dates (or that month)."),
        ("Bills & needs", "Spending rows whose category belongs to a 'Bills & needs' group (like Needs: rent, phone, groceries)."),
        ("Flexible", "Spending rows whose category belongs to a 'Flexible spending' group (like Wants and Fun: shopping, eating out)."),
        ("Saved", "Spending rows whose category belongs to a 'Saving & investing' group (like Savings and Investing: emergency fund, investments)."),
        ("Left over", "Received minus Bills & needs, Flexible and Saved. Haven't logged money in yet? It uses Expected."),
        ("Which group?", "Setup tab → each category's Group, and each group's kind. That's what sorts a row into these columns."),
    ]),
    ("WHAT EACH TAB IS FOR", [
        ("Setup", "Your pay, taxes, groups and categories. Fill it in once and change it when your pay changes."),
        ("Spending", "Where you type what you spend (left) and the money you get (right). This is the tab you use most."),
        ("Dashboard", "This month at a glance: pay, spent, left, each group and your alerts. Pick another month at the top."),
        ("Phone view", "The same things in one narrow column that fits a phone screen."),
        ("Bills", "Safe to spend until next payday, bills coming up, payment arrangements and yearly costs."),
        ("Goals", "Your savings goals and a month-by-month savings scorecard."),
        ("Accounts", "Savings and investing accounts, what they could grow to, and your net worth."),
        ("Debts", "Your debts and when you'll pay them off."),
        ("Reports", "Month by month, each category, day by day and paycheck by paycheck."),
    ]),
    ("SOMETHING LOOKS WRONG?", [
        ("Money in too low", "Check that every Date received is a real date (right side of the box) in the right month and year."),
        ("Pay too high", "The full paycheck box is ONE check before taxes, not two checks added together."),
        ("Paydays are off", "Check 'When was your first payday?' is the day your very first check landed, and that "
                            "'How often' is right."),
        ("A purchase is missing", "Check it has a real date and a category picked from the list."),
    ]),
]


def build_guide(wb, look):
    """'How to use' tab: a short, plain-English guide. Text only, nothing to type here."""
    ws = wb.create_sheet("How to use")
    widths(ws, {"A": 2, "B": 20, "C": 100})
    merged(ws, "B1:C1", "How to use Paycheck Planner", f=font(22, True, look.primary))
    ws.row_dimensions[1].height = 36
    merged(ws, "B2:C2", "Yellow boxes are for you to type in. Everything else fills in by itself.",
           f=font(11, color=MUTED))
    r = 4
    for n, (title, rows) in enumerate(GUIDE):
        section_bar(ws, r, "B", "C", title, look, "ABCDEFGH"[n])
        r += 1
        for label, text in rows:
            put(ws, f"B{r}", label, f=font(11, True, look.primary),
                align=Alignment(vertical="top", wrap_text=True))
            put(ws, f"C{r}", text, f=font(11), align=Alignment(vertical="top", wrap_text=True))
            lines = max(1, -(-len(text) // 105))           # about 105 letters fit on one line of column C
            ws.row_dimensions[r].height = 6 + 16 * lines
            for c in "BC":
                ws[f"{c}{r}"].border = UNDER
            r += 1
        r += 1                                               # one empty row between sections
    ws.freeze_panes = "A3"


# ================================================================== BUILD
def build(theme, style):
    look = Look(theme, style)
    wb = Workbook()
    wb.remove(wb.active)
    build_setup(wb, look)
    build_spending(wb, look)
    build_engine(wb, look)       # before Dashboard: the donut charts read Engine
    build_dashboard(wb, look)
    build_reports(wb, look)
    build_goals(wb, look)
    build_bills(wb, look)
    build_debts(wb, look)
    build_accounts(wb, look)
    build_today(wb, look)
    build_guide(wb, look)
    order = ["How to use", "Setup", "Spending", "Dashboard", "Phone view", "Bills", "Goals", "Accounts", "Debts", "Reports", "Engine"]
    wb._sheets.sort(key=lambda w: order.index(w.title))     # the order buyers see the tabs in
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = look.primary
        ws.sheet_view.zoomScale = 100
    wb.active = 1                 # open on Setup; "How to use" is the first tab, one click away
    OUT_DIR.mkdir(exist_ok=True)
    p = CONFIG["product"]
    suffix = "" if style == "card" else f"_{style}"
    path = OUT_DIR / f"{p['name'].replace(' ', '_')}_v{p['version']}_{theme['id']}{suffix}.xlsx"
    wb.save(path)
    return path


if __name__ == "__main__":
    args = sys.argv[1:]
    themes = [t for t in CONFIG["themes"] if not set(args) & {x["id"] for x in CONFIG["themes"]} or t["id"] in args]
    styles = [s for s in CONFIG["styles"] if not set(args) & set(CONFIG["styles"]) or s in args]
    for t in themes:
        for s in styles:
            try:
                print("Built", build(t, s).relative_to(ROOT))
            except PermissionError as e:
                # Windows locks a file while Excel has it open, so we can't save over it
                print(f"\n✖ Can't save {Path(e.filename).name}: it's open in Excel (or another app).")
                print("  Close that file, then run this again.")
                sys.exit(1)
    # Tidy the output folder: keep the newest versions only (owner's standing permission, see CLAUDE.md)
    sys.path.insert(0, str(Path(__file__).parent))
    import clean_output
    clean_output.main(["--yes"])
    if "--push" in args:                      # also send the result to the master Google Sheet
        sys.path.insert(0, str(Path(__file__).parent))
        import push_to_sheets
        push_to_sheets.push()
