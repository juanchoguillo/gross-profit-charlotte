"""
Branded, formula-dynamic Excel exports for the Gross Profit dashboard.

Every derived number in the workbook is a live Excel formula (with its current
value cached so pandas / the saved-report viewer still read plain numbers):

  Orders sheet   — Net Sales, Direct Cost, Gross/Net Profit, margins and
                   Contribution recompute from the editable base columns.
  Summary sheet  — every rep row is a SUMIFS over the orders sheet; shop /
                   manager rows sum their reps; the grand total sums the reps.
  Fixed rate     — one editable cell on the cover sheet (named "FixedRate")
                   drives every Fixed Cost / Overhead formula in the book.
  Other sheets   — customer-type and operational tables aggregate the orders
                   sheet; ratio columns (Avg $, margins) are formulas.

Styling mirrors the hand-formatted "07. July Gross Profit Full Report" workbook:
MC Granite navy #194052 + gold #989027, tinted branch bands, data bars, and the
company logo on the cover sheet.
"""
import io
import math
import os
from datetime import datetime

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

NAVY = "#194052"
GOLD = "#989027"
GOLD_TINT = "#F3F0DC"      # shop/manager subtotal rows + one branch band
BLUE_TINT = "#EAF0F3"      # zebra stripe + one branch band
GRAY_TINT = "#F2F4F5"      # third branch band
RED = "#C0392B"
TEXT = "#333333"
BAND_TINTS = [BLUE_TINT, GOLD_TINT, GRAY_TINT]

FMT = {
    "money": '"$"#,##0;\\("$"#,##0\\);"-"',
    "money2": '"$"#,##0.00',
    "sqft": '#,##0.0',
    "pct": '0.0%',
    "int": '#,##0',
    "num": '#,##0.00',
    "text": None,
}

COVER_SHEET = "Report"
LOGO_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "assets", "logo-export.png")
RATE_CELL = "$D$13"        # FixedRate lives here on the cover sheet


# --------------------------------------------------------------------------
# format cache — one xlsxwriter Format per unique property combination
# --------------------------------------------------------------------------
class _Formats:
    def __init__(self, wb):
        self.wb, self.cache = wb, {}

    def get(self, kind="text", fill=None, bold=False, color=None, size=None,
            align=None, border_color=None, header=False, italic=False):
        key = (kind, fill, bold, color, size, align, border_color, header, italic)
        if key not in self.cache:
            p = {}
            if header:
                p.update({"bold": True, "font_color": "white", "bg_color": NAVY,
                          "align": "center", "valign": "vcenter", "text_wrap": True})
            else:
                if FMT.get(kind):
                    p["num_format"] = FMT[kind]
                p["font_color"] = color or TEXT
                if fill:
                    p["bg_color"] = fill
                if bold:
                    p["bold"] = True
            if size:
                p["font_size"] = size
            if align:
                p["align"] = align
            if italic:
                p["italic"] = True
            if border_color:
                p.update({"border": 1, "border_color": border_color})
            self.cache[key] = self.wb.add_format(p)
        return self.cache[key]


def _isnum(v):
    """A real, finite number — including numpy scalars (np.int64 is NOT a
    Python int). NaN/±inf must never reach the XML: Excel would flag the file
    as corrupt and 'repair' it, dropping content."""
    if v is None or isinstance(v, (bool, str, bytes)):
        return False
    try:
        f = float(v)
    except (TypeError, ValueError):
        return False
    return math.isfinite(f)


def _kind_of(col: str) -> str:
    """Number-format kind inferred from a header name (same rules the app's
    saved-report viewer uses)."""
    n = str(col).lower()
    if any(w in n for w in ("id", "number", "#", "name", "rep", "shop", "manager",
                            "branch", "type", "customer", "material", "color",
                            "group", "level", "loaded", "ref", "person")):
        return "text"
    if "margin" in n or ("%" in n and "avg" not in n):
        return "pct"
    if any(w in n for w in ("avg", "price")):
        return "money2"
    if any(w in n for w in ("orders", "crew", "tabs")):
        return "int"
    if any(w in n for w in ("$", "income", "cost", "revenue", "invoice", "tax",
                            "profit", "overhead", "amount", "memo", "sales",
                            "contribution")):
        return "money"
    if "sqft" in n.replace(" ", "") or "quantity" in n:
        return "sqft"
    return "num"


def _write_sheet(wb, F, name, df, *, tab_color=GOLD, kinds=None, widths=None,
                 header_h=28, freeze=True, fill_of=None, formula_of=None,
                 text_cols_note=True):
    """Generic styled table: navy header row, per-row fills, optional per-cell
    formulas as (formula, cached_value). Returns the worksheet."""
    ws = wb.add_worksheet(name)
    ws.set_tab_color(tab_color)
    cols = list(df.columns)
    kinds = kinds or {}
    for j, c in enumerate(cols):
        ws.write(0, j, str(c), F.get(header=True))
        w = (widths or {}).get(c)
        ws.set_column(j, j, w if w else max(10.5, min(len(str(c)) * 1.15 + 2, 26)))
    ws.set_row(0, header_h)
    if freeze:
        ws.freeze_panes(1, 0)

    for i in range(len(df)):
        row = df.iloc[i]
        style = fill_of(i, row) if fill_of else {}
        for j, c in enumerate(cols):
            kind = kinds.get(c) or _kind_of(c)
            fmt = F.get(kind=kind, **style)
            fx = formula_of(i, j, c, i + 2) if formula_of else None
            v = row[c]
            if fx:
                if _isnum(v):
                    cached = float(v)
                elif isinstance(v, str):
                    cached = v
                else:
                    cached = ""       # NaN/inf/None — recalculates on open
                ws.write_formula(i + 1, j, fx, fmt, cached)
            elif _isnum(v):
                ws.write_number(i + 1, j, float(v), fmt)
            elif isinstance(v, str) and v and v != "nan":
                ws.write_string(i + 1, j, v, fmt)
            elif v is None or pd.isna(v) \
                    or str(v) in ("", "nan", "NaT", "<NA>", "None", "inf", "-inf"):
                ws.write_blank(i + 1, j, None, fmt)
            else:
                ws.write_string(i + 1, j, str(v), fmt)
    if text_cols_note and len(df):
        # ids stored as text are intentional — hide Excel's green triangles
        last = xl_col_to_name(len(cols) - 1)
        ws.ignore_errors({"number_stored_as_text": f"A2:{last}{len(df) + 1}"})
    return ws


def _zebra(i, _row):
    return {"fill": BLUE_TINT} if i % 2 == 0 else {}


def _databar(ws, df, col, color, *, negative=None):
    if col not in df.columns or not len(df):
        return
    j = list(df.columns).index(col)
    opts = {"type": "data_bar", "bar_color": color, "bar_border_color": color}
    if negative:
        opts["bar_negative_color"] = negative
    ws.conditional_format(1, j, len(df), j, opts)


def _q(sheet: str) -> str:
    return "'" + sheet.replace("'", "''") + "'"


# --------------------------------------------------------------------------
# cover sheet — logo, report info, the editable FixedRate cell, live KPIs
# --------------------------------------------------------------------------
def _cover(wb, F, *, subtitle, period, group_by, basis_label, fixed_rate, kpis):
    ws = wb.add_worksheet(COVER_SHEET)
    ws.set_tab_color(GOLD)
    ws.hide_gridlines(2)
    for col, w in enumerate([2.5, 30, 6, 18, 18, 18]):
        ws.set_column(col, col, w)

    if os.path.exists(LOGO_FILE):
        # 480px source ≈ 128px tall on sheet
        ws.insert_image("B2", LOGO_FILE, {"x_scale": 0.267, "y_scale": 0.267})

    ws.write("D3", "MC GRANITE", F.get(bold=True, color=NAVY, size=24))
    ws.write("D4", "Gross Profit Report", F.get(bold=True, color=GOLD, size=16))
    ws.write("D5", subtitle, F.get(color=TEXT, italic=True))
    meta = [("Period", period), ("Grouped by", group_by),
            ("Income basis", basis_label),
            ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M"))]
    for i, (lbl, val) in enumerate(meta):
        ws.write(6 + i, 3, lbl, F.get(bold=True, color=NAVY))
        ws.write(6 + i, 4, str(val), F.get(color=TEXT))

    # gold divider
    ws.set_row(10, 4)
    for col in range(1, 6):
        ws.write_blank(10, col, None, F.get(fill=GOLD))

    ws.write("B13", "Fixed cost $ / SqFt (stone)  →", F.get(bold=True, color=NAVY))
    ws.write_number(RATE_CELL.replace("$", ""), float(fixed_rate),
                    F.get(kind="money2", fill=GOLD_TINT, bold=True, color=NAVY,
                          border_color=NAVY, align="center"))
    wb.define_name("FixedRate", f"={COVER_SHEET}!{RATE_CELL}")
    ws.write("B14", "Edit this rate (or any base number on the order sheet) — every "
                    "Fixed Cost, Overhead, profit and margin formula recalculates.",
             F.get(color="#777777", italic=True))

    ws.write("B16", "PERIOD TOTALS — live formulas", F.get(header=True))
    ws.write_blank("C16", None, F.get(header=True))
    r = 16
    for lbl, formula, cached, kind in kpis:
        fill = GOLD_TINT if (r - 16) % 2 else None
        ws.write(r, 1, lbl, F.get(bold=True, color=TEXT, fill=fill))
        cached = float(cached) if _isnum(cached) else ""
        ws.write_formula(r, 2, formula, F.get(kind=kind, fill=fill, bold=True,
                                              color=NAVY), cached)
        r += 1
    ws.set_column(2, 2, 14)
    return ws


# --------------------------------------------------------------------------
# hierarchical summary sheet (Level / Name / value columns)
# --------------------------------------------------------------------------
# summary column -> the matching orders-sheet column, by (display, internal) name
_SUM_FIELDS = ["SqFtBilled", "SqFtAllocated", "Income", "MaterialCost",
               "TemplateCost", "InstallCost", "FabricationCost"]


def _summary_sheet(wb, F, name, df, *, orders_sheet, orders_col, crit_col,
                   income_dynamic):
    """orders_col: {summary field: orders-sheet column letter}; crit_col: the
    orders-sheet letter holding the sales-rep name."""
    n = len(df)
    levels = list(df["Level"])
    # excel row of the last non-total data row (data starts at excel row 2)
    last_data = n if levels and levels[-1] == "Total" else n + 1
    letters = {c: xl_col_to_name(j) for j, c in enumerate(df.columns)}
    OQ = _q(orders_sheet)

    def sumifs(field, xr):
        L = orders_col.get(field)
        if not L:
            return None
        return (f"=SUMIFS({OQ}!${L}:${L},{OQ}!${crit_col}:${crit_col},"
                f"$B{xr})")

    # rep-block ranges for each group row: group at i, reps i+1..k
    block_end = {}
    for i, lvl in enumerate(levels):
        if lvl in ("Shop", "Manager"):
            k = i + 1
            while k < n and levels[k] == "Rep":
                k += 1
            block_end[i] = (i + 3, k + 1)     # excel rows of first..last rep

    def formula_of(i, j, c, xr):
        lvl = levels[i]
        cl = letters[c]
        if c in _SUM_FIELDS:
            if c == "Income" and not income_dynamic and lvl == "Rep":
                return None                    # paid basis: rep income stays static
            if lvl == "Rep":
                return sumifs(c, xr)
            if lvl in ("Shop", "Manager") and i in block_end:
                a, b = block_end[i]
                return f"=SUM({cl}{a}:{cl}{b})"
            if lvl == "Total":
                return (f'=SUMIF($A$2:$A${last_data},"Rep",'
                        f'{cl}$2:{cl}${last_data})')
            return None
        if c == "OperationalCost":
            return (f"={letters['TemplateCost']}{xr}+{letters['InstallCost']}{xr}"
                    f"+{letters['FabricationCost']}{xr}")
        if c == "Overhead":
            return f"={letters['SqFtBilled']}{xr}*FixedRate"
        if c == "TotalCost":
            return (f"={letters['MaterialCost']}{xr}+{letters['OperationalCost']}{xr}"
                    f"+{letters['Overhead']}{xr}")
        if c == "Profit":
            return f"={letters['Income']}{xr}-{letters['TotalCost']}{xr}"
        if c == "Margin":
            return (f'=IFERROR({letters["Profit"]}{xr}/{letters["Income"]}{xr},"")')
        return None

    def fill_of(_i, row):
        lvl = row["Level"]
        if lvl == "Total":
            return {"fill": NAVY, "bold": True, "color": "white"}
        if lvl in ("Shop", "Manager"):
            return {"fill": GOLD_TINT, "bold": True, "color": NAVY}
        return {}

    kinds = {"Level": "text", "Name": "text", "SqFtBilled": "sqft",
             "SqFtAllocated": "sqft", "Margin": "pct"}
    widths = {"Level": 8, "Name": 21, "SqFtBilled": 12, "SqFtAllocated": 13}
    ws = _write_sheet(wb, F, name, df, tab_color=NAVY, kinds=kinds, widths=widths,
                      header_h=30, fill_of=fill_of, formula_of=formula_of)
    return ws


# --------------------------------------------------------------------------
# the full export — one sheet per dashboard tab
# --------------------------------------------------------------------------
ORDERS_SHEET = "Orders — Net Margin"

_ORDER_WIDTHS = {"Order ID": 8, "Order #": 11, "Ref #": 25.5, "Cust #": 9,
                 "Customer Name": 27.3, "Branch": 10, "Sales Person": 17,
                 "Customer Type": 22.6, "Sq Ft (Stone)": 14.2}


def _orders_sheet(wb, F, df):
    """The Orders — Net Margin sheet: branch-tinted bands, derived columns as
    formulas, data bars on the profit columns."""
    cols = list(df.columns)
    L = {c: xl_col_to_name(j) for j, c in enumerate(cols)}

    def formula_of(i, j, c, xr):
        if c == "Net Sales":
            return (f"={L['Total Invoice']}{xr}-{L['Sales Tax']}{xr}"
                    f"-{L['Credit Memo']}{xr}")
        if c == "Total Direct Cost":
            return (f"={L['Material Cost']}{xr}+{L['Template Cost']}{xr}"
                    f"+{L['Install Cost']}{xr}+{L['Fab Cost']}{xr}")
        if c == "Gross Profit ($)":
            return f"={L['Net Sales']}{xr}-{L['Total Direct Cost']}{xr}"
        if c == "GP Margin %":
            return f'=IFERROR({L["Gross Profit ($)"]}{xr}/{L["Net Sales"]}{xr},"")'
        if c == "Fixed Cost":
            return f"={L['Sq Ft (Stone)']}{xr}*FixedRate"
        if c == "Net Profit ($)":
            return f"={L['Gross Profit ($)']}{xr}-{L['Fixed Cost']}{xr}"
        if c == "Net Margin %":
            return f'=IFERROR({L["Net Profit ($)"]}{xr}/{L["Net Sales"]}{xr},"")'
        if c == "Contribution / Sq Ft":
            return (f'=IFERROR({L["Gross Profit ($)"]}{xr}'
                    f'/{L["Sq Ft (Stone)"]}{xr},"")')
        return None

    # one tint per branch block (Blue Ridge / Jasper / Kennesaw …)
    tint_of, tint, prev = {}, -1, object()
    branches = list(df["Branch"]) if "Branch" in df.columns else [""] * len(df)
    for i, b in enumerate(branches):
        if b != prev:
            tint, prev = tint + 1, b
        tint_of[i] = BAND_TINTS[tint % len(BAND_TINTS)]

    kinds = {c: ("pct" if c.endswith("%") else
                 "sqft" if c.startswith("Sq Ft") else
                 "money2" if c == "Contribution / Sq Ft" else
                 "text" if _kind_of(c) == "text" else "money") for c in cols}
    ws = _write_sheet(wb, F, ORDERS_SHEET, df, tab_color=NAVY, kinds=kinds,
                      widths=_ORDER_WIDTHS, header_h=34,
                      fill_of=lambda i, _r: {"fill": tint_of[i]},
                      formula_of=formula_of)
    _databar(ws, df, "Gross Profit ($)", NAVY, negative=RED)
    _databar(ws, df, "Net Margin %", GOLD, negative=RED)
    if "Net Profit ($)" in cols and len(df):
        j = cols.index("Net Profit ($)")
        ws.conditional_format(1, j, len(df), j, {
            "type": "cell", "criteria": "<", "value": 0,
            "format": wb.add_format({"font_color": RED, "bold": True})})
    return L


def _avg_formula(df, letters):
    """Ratio column (Avg $/unit, Avg $/SqFt, AvgPrice) = Revenue / base col."""
    avg = next((c for c in df.columns if str(c).lower().startswith("avg")), None)
    base = next((c for c in ("Quantity", "SqFt") if c in df.columns), None)
    if not avg or not base or "Revenue" not in df.columns:
        return None
    return avg, lambda xr: (f'=IFERROR({letters["Revenue"]}{xr}'
                            f'/{letters[base]}{xr},"")')


def _simple_sheet(wb, F, name, df, *, databars=(), avg=True, fill_of=_zebra,
                  kinds=None, widths=None):
    letters = {c: xl_col_to_name(j) for j, c in enumerate(df.columns)}
    av = _avg_formula(df, letters) if avg else None

    def formula_of(_i, _j, c, xr):
        if av and c == av[0]:
            return av[1](xr)
        return None

    ws = _write_sheet(wb, F, name, df, kinds=kinds, widths=widths,
                      fill_of=fill_of, formula_of=formula_of)
    for col, color in databars:
        _databar(ws, df, col, color)
    return ws


def build_full_export(orders_nm, summary, by_material, by_color, by_customer_type,
                      by_service_group, op_diag_df, mapping_df, group_by, *,
                      fixed_rate=0.0, income_basis="billed", period="",
                      basis_label="", subtitle=None):
    """One workbook mirroring every dashboard tab — cover + one sheet per tab,
    with live formulas throughout (see module docstring)."""
    orders = orders_nm.copy()
    if {"Total Invoice", "Sales Tax", "Net Sales"} <= set(orders.columns) \
            and "Credit Memo" not in orders.columns:
        cm = (orders["Total Invoice"] - orders["Sales Tax"]
              - orders["Net Sales"]).round(2)
        orders.insert(list(orders.columns).index("Sales Tax") + 1,
                      "Credit Memo", cm.mask(cm.abs() < 0.005, 0.0))

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    F = _Formats(wb)

    OQ = _q(ORDERS_SHEET)
    ocols = {c: xl_col_to_name(j) for j, c in enumerate(orders.columns)}

    def osum(col):
        return f"SUM({OQ}!${ocols[col]}:${ocols[col]})"

    kpis = [
        ("Net Sales", f"={osum('Net Sales')}", orders["Net Sales"].sum(), "money"),
        ("Total Direct Cost", f"={osum('Total Direct Cost')}",
         orders["Total Direct Cost"].sum(), "money"),
        ("Gross Profit", f"={osum('Gross Profit ($)')}",
         orders["Gross Profit ($)"].sum(), "money"),
        ("Fixed Cost", f"={osum('Fixed Cost')}", orders["Fixed Cost"].sum(), "money"),
        ("Net Profit", f"={osum('Net Profit ($)')}",
         orders["Net Profit ($)"].sum(), "money"),
        ("Net Margin %", f"=IFERROR({osum('Net Profit ($)')}/{osum('Net Sales')},\"\")",
         (orders["Net Profit ($)"].sum() / orders["Net Sales"].sum())
         if orders["Net Sales"].sum() else None, "pct"),
        ("Sq Ft (Stone)", f"={osum('Sq Ft (Stone)')}",
         orders["Sq Ft (Stone)"].sum(), "sqft"),
    ]
    _cover(wb, F, subtitle=subtitle or "Full report — one sheet per dashboard tab",
           period=period, group_by=group_by, basis_label=basis_label or income_basis,
           fixed_rate=fixed_rate, kpis=kpis)

    if summary is not None and len(summary):
        _summary_sheet(wb, F, f"Summary by {group_by}", summary,
                       orders_sheet=ORDERS_SHEET,
                       orders_col={"SqFtBilled": ocols["Sq Ft (Stone)"],
                                   "SqFtAllocated": ocols["Sq Ft Allocated"],
                                   "Income": ocols["Net Sales"],
                                   "MaterialCost": ocols["Material Cost"],
                                   "TemplateCost": ocols["Template Cost"],
                                   "InstallCost": ocols["Install Cost"],
                                   "FabricationCost": ocols["Fab Cost"]},
                       crit_col=ocols["Sales Person"],
                       income_dynamic=(income_basis == "billed"))

    _orders_sheet(wb, F, orders)

    if by_material is not None and len(by_material):
        _simple_sheet(wb, F, "By Material", by_material,
                      databars=[("Revenue", GOLD)])
    if by_color is not None and len(by_color):
        _simple_sheet(wb, F, "Countertop Colors", by_color,
                      databars=[("Revenue", GOLD)], widths={"Color": 35.5})

    if by_customer_type is not None and len(by_customer_type):
        ct = by_customer_type.reset_index(drop=True)
        tcol = ocols.get("Customer Type")
        billed = income_basis == "billed"

        def ct_formula(_i, _j, c, xr):
            crit = f"{OQ}!${tcol}:${tcol},$A{xr}"
            if c == "Orders":
                return f"=COUNTIFS({crit})"
            if not billed:
                return None
            if c == "Income":
                return f"=SUMIFS({OQ}!${ocols['Net Sales']}:${ocols['Net Sales']},{crit})"
            if c == "Profit":
                return (f"=SUMIFS({OQ}!${ocols['Net Profit ($)']}:"
                        f"${ocols['Net Profit ($)']},{crit})")
            return None

        ws = _write_sheet(wb, F, "By Customer Type", ct,
                          widths={"CustomerType": 22.6},
                          fill_of=_zebra, formula_of=ct_formula)
        _databar(ws, ct, "Income", NAVY)

    if by_service_group is not None and len(by_service_group):
        sg = by_service_group.reset_index(drop=True)
        col_by_group = {"Template": "Template Cost", "Install": "Install Cost",
                        "Fabrication": "Fab Cost"}

        def sg_formula(i, _j, c, _xr):
            oc = col_by_group.get(str(sg.iloc[i].get("Group", "")))
            return f"={osum(oc)}" if (c == "Amount" and oc) else None

        _write_sheet(wb, F, "Operational by Type", sg, freeze=False,
                     widths={"Amount": 20}, fill_of=_zebra, formula_of=sg_formula)

    if op_diag_df is not None and len(op_diag_df):
        od = op_diag_df.reset_index(drop=True)
        odL = {c: xl_col_to_name(j) for j, c in enumerate(od.columns)}
        applied = "Cost applied (period)"
        col_by_type = {"Template": "Template Cost", "Install": "Install Cost",
                       "Fabrication": "Fab Cost"}

        def od_formula(i, _j, c, xr):
            t = str(od.iloc[i].get("Type", ""))
            if t == "TOTAL" and c not in ("Type", "Loaded"):
                return f"=SUM({odL[c]}2:{odL[c]}{xr - 1})"
            if c == applied and t in col_by_type:
                return f"={osum(col_by_type[t])}"
            return None

        def od_fill(i, row):
            if str(row.get("Type")) == "TOTAL":
                return {"fill": NAVY, "bold": True, "color": "white"}
            return {}

        _write_sheet(wb, F, "Operational Files", od, freeze=False, header_h=34,
                     kinds={"Cost in file": "money", applied: "money"},
                     fill_of=od_fill, formula_of=od_formula)

    if mapping_df is not None and len(mapping_df):
        _write_sheet(wb, F, "Shops & Managers", mapping_df,
                     widths={"Sales Rep": 17, "Manager": 16}, fill_of=_zebra)

    wb.close()
    return buf.getvalue()


# --------------------------------------------------------------------------
# the summary-tab report — same workbook minus the colors/diagnostics/mapping
# sheets, and the orders sheet in the SAME original "Orders — Net Margin"
# layout (columns + names) as the full report.
# --------------------------------------------------------------------------
def build_report_excel(orders_nm, summary, by_material, by_customer_type,
                       by_service_group, group_by, *, fixed_rate=0.0,
                       income_basis="billed", period="", basis_label=""):
    return build_full_export(
        orders_nm, summary, by_material, None, by_customer_type,
        by_service_group, None, None, group_by,
        fixed_rate=fixed_rate, income_basis=income_basis, period=period,
        basis_label=basis_label, subtitle=f"Summary report by {group_by}")
