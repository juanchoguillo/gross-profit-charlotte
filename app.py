"""
MC Granite — Dynamic Gross Profit Dashboard  (v2)
Run with:  streamlit run app.py
"""
import glob
import io
import os
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import gp_core as gp
from report_xlsx import COVER_SHEET, build_full_export, build_report_excel

st.set_page_config(page_title="MC Granite — Gross Profit", layout="wide", page_icon="📊")
PALETTE = ["#1f4e79", "#2e75b6", "#9dc3e6", "#c55a11", "#ed7d31", "#70ad47", "#a6a6a6", "#7030a0"]

_LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "assets", "LOGO-MC-GRANITE.png")
if os.path.exists(_LOGO):
    st.logo(_LOGO, size="large")

# Max table height. Tables taller than this scroll INSIDE the grid, which keeps
# the column-header row frozen on screen — instead of growing the page so the
# headers scroll away.
TABLE_H = 600

st.markdown(
    "<style>.block-container{padding-top:1.4rem}"
    "[data-testid='stMetricValue']{font-size:1.45rem}</style>",
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------
# Login — static users, dropdown + PIN. shop=None ⇒ admin (sees/edits all);
# otherwise the whole dashboard is scoped to that user's shop.
# --------------------------------------------------------------------------
USERS = {
    "maria":   {"pin": "1024", "shop": None},
    "carlos":  {"pin": "1021", "shop": "Jasper"},
    "manuel":  {"pin": "1022", "shop": "Blue Ridge"},
    "luciano": {"pin": "1023", "shop": "Kennesaw"},
}

if "auth_user" not in st.session_state:
    _, mid, _ = st.columns([1, 1.1, 1])
    with mid:
        if os.path.exists(_LOGO):
            st.image(_LOGO, width=170)
        st.title("🔐 Gross Profit")
        st.caption("MC Granite — sign in to continue")
        with st.form("login"):
            login_user = st.selectbox("User", list(USERS), format_func=str.title)
            login_pin = st.text_input("PIN", type="password", max_chars=8,
                                      placeholder="4-digit PIN")
            submitted = st.form_submit_button("Sign in", type="primary")
        if submitted:
            if login_pin.strip() == USERS[login_user]["pin"]:
                st.session_state["auth_user"] = login_user
                st.rerun()
            else:
                st.error("Wrong PIN — try again.")
    st.stop()

AUTH_USER = st.session_state["auth_user"]
USER_SHOP = USERS[AUTH_USER]["shop"]          # None = all shops
IS_ADMIN = USER_SHOP is None


@st.cache_data(show_spinner="Reading files…")
def _load(blobs: tuple):
    srcs = [(io.BytesIO(b), n) for n, b in blobs]
    return gp.load_files(srcs)


def money(x):
    return f"${x:,.2f}"


# Orders tab — mirrors the example workbook's "Orders — Net Margin" sheet, in
# its exact column order. (internal column, display header).
NET_MARGIN_COLS = [
    ("OrderID", "Order ID"), ("OrderNumber", "Order #"), ("RefNumber", "Ref #"),
    ("CustomerNumber", "Cust #"), ("Customer", "Customer Name"), ("Shop", "Branch"),
    ("Rep", "Sales Person"), ("CustomerType", "Customer Type"),
    ("SqFtBilled", "Sq Ft (Stone)"), ("SqFtAllocated", "Sq Ft Allocated"),
    ("TotalInvoice", "Total Invoice"), ("SalesTax", "Sales Tax"), ("NetSales", "Net Sales"),
    ("TemplateCost", "Template Cost"), ("InstallCost", "Install Cost"),
    ("FabricationCost", "Fab Cost"), ("MaterialCost", "Material Cost"),
    ("DirectCost", "Total Direct Cost"), ("GrossProfit", "Gross Profit ($)"),
    ("GPMargin", "GP Margin %"), ("FixedCost", "Fixed Cost"),
    ("NetProfit", "Net Profit ($)"), ("NetMargin", "Net Margin %"),
    ("Contribution", "Contribution / Sq Ft"),
]
NM_MONEY = ["TotalInvoice", "SalesTax", "NetSales", "TemplateCost", "InstallCost",
            "FabricationCost", "MaterialCost", "DirectCost", "GrossProfit",
            "FixedCost", "NetProfit", "Contribution"]
NM_SQFT = ["SqFtBilled", "SqFtAllocated"]
NM_PCT = ["GPMargin", "NetMargin"]


def build_net_margin(df: pd.DataFrame) -> pd.DataFrame:
    """Add the Net-Margin derived columns to a per-order frame.
    Direct cost = material + template + install + fab; Gross Profit = Net Sales −
    Direct; Fixed Cost = stone sqft × the fixed-cost rate (the app's Overhead
    column); Net Profit = Gross − Fixed."""
    nm = df.copy()
    nm["DirectCost"] = nm["MaterialCost"] + nm["OperationalCost"]
    nm["GrossProfit"] = nm["NetSales"] - nm["DirectCost"]
    nm["FixedCost"] = nm["Overhead"]
    nm["NetProfit"] = nm["GrossProfit"] - nm["FixedCost"]
    ns = nm["NetSales"].replace(0, pd.NA)
    nm["GPMargin"] = nm["GrossProfit"] / ns
    nm["NetMargin"] = nm["NetProfit"] / ns
    nm["Contribution"] = nm["GrossProfit"] / nm["SqFtBilled"].replace(0, pd.NA)
    return nm


def op_diag_table(od: dict) -> pd.DataFrame:
    """The Operational tab's per-file diagnostics, one row per cost type plus a
    TOTAL row. Totals only — individual crew tabs are counted, not listed."""
    rows = []
    for key in ["template", "install", "fabrication"]:
        v = od[key]
        rows.append({
            "Type": v["label"],
            "Loaded": "✅" if v["loaded"] else "—",
            "Crew tabs": len(v["sheets"]),
            "Orders in file": v["orders_in_file"],
            "Cost in file": v["total_in_file"],
            "Orders matched": v["matched_orders"],
            "Cost applied (period)": v["matched_total"],
        })
    df = pd.DataFrame(rows)
    total = {c: df[c].sum() for c in
             ["Crew tabs", "Orders in file", "Cost in file",
              "Orders matched", "Cost applied (period)"]}
    total.update({"Type": "TOTAL", "Loaded": ""})
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


# --------------------------------------------------------------------------
# Saved-report viewer — re-open an Excel report exported from this app
# --------------------------------------------------------------------------
@st.cache_data(show_spinner="Reading report…")
def _load_report(blob: bytes) -> dict:
    """All sheets of a previously downloaded report: {sheet name: DataFrame}."""
    return pd.read_excel(io.BytesIO(blob), sheet_name=None)


# Pretty headers for sheets that were exported with internal column names
# (the Summary and Order Detail sheets); already-pretty sheets pass through.
_VIEW_RENAME = {**dict(NET_MARGIN_COLS), "Income": "Income",
                "OperationalCost": "Operational Cost", "TotalCost": "Total Cost",
                "Overhead": "Overhead", "Margin": "Margin", "AvgPrice": "Avg $"}
_ROW_DARK = "background-color:#1f4e79;color:white;font-weight:700"
_ROW_GROUP = "background-color:#dde7f2;font-weight:700"


def _sheet_fmt(df: pd.DataFrame) -> dict:
    """Number formats per column, inferred from the header name."""
    fmt = {}
    for c in df.columns:
        if not pd.api.types.is_numeric_dtype(df[c]):
            continue
        n = str(c).lower()
        if any(w in n for w in ("id", "number", "#")):
            fmt[c] = "{:.0f}"
        elif any(w in n for w in ("rank", "orders", "crew", "tabs")):
            fmt[c] = "{:,.0f}"
        elif "margin" in n or "%" in n:
            fmt[c] = "{:.2%}"
        elif any(w in n for w in ("$", "income", "cost", "revenue", "invoice", "sales",
                                  "tax", "profit", "overhead", "contribution",
                                  "amount", "price", "avg")):
            fmt[c] = "${:,.2f}"
        else:
            fmt[c] = "{:,.2f}"
    return fmt


def _tab_icon(sheet: str) -> str:
    sl = sheet.lower()
    for key, icon in (("summary", "📋"), ("order", "📑"), ("color", "🎨"),
                      ("material", "🪨"), ("customer", "👥"),
                      ("operational", "🛠️"), ("shop", "⚙️")):
        if key in sl:
            return icon + " "
    return ""


def render_saved_report(name: str, book: dict):
    st.title("📂 Saved Report")
    st.caption(f"Viewing **{name}** — a read-only snapshot of a downloaded report. "
               "Sidebar settings and filters don't apply here; switch **Load data "
               "from** back to *Upload files* for the live dashboard.")

    sheets = {}
    for sname, df in book.items():
        if sname == COVER_SHEET:      # branding/settings cover page — not a table
            continue
        if df is None or df.empty:
            continue
        sdf = df
        if not IS_ADMIN:
            if {"Level", "Name"} <= set(sdf.columns):
                # Hierarchical summary sheet: keep this shop's subtotal row and
                # the rep rows underneath it (grand total is company-wide).
                lv = sdf["Level"].astype(str).str.strip()
                nmv = sdf["Name"].astype(str).str.strip().str.lower()
                starts = sdf.index[(lv == "Shop") & (nmv == USER_SHOP.lower())]
                if not len(starts):
                    continue            # grouped by Manager etc. — can't scope
                keep = [starts[0]]
                for i in sdf.index[sdf.index > starts[0]]:
                    if str(sdf.at[i, "Level"]).strip() != "Rep":
                        break
                    keep.append(i)
                sdf = sdf.loc[keep]
            else:
                shop_col = next((c for c in ("Shop", "Branch")
                                 if c in sdf.columns), None)
                if shop_col is None:
                    continue            # company-wide sheet — admin only
                sdf = sdf[sdf[shop_col].astype(str).str.strip().str.lower()
                          == USER_SHOP.lower()]
            if sdf.empty:
                continue
        sheets[sname] = sdf.reset_index(drop=True)

    if not sheets:
        st.warning("Nothing to show in this file"
                   + ("." if IS_ADMIN else f" for **{USER_SHOP}**."))
        return
    if not IS_ADMIN:
        st.caption(f"🔒 Showing **{USER_SHOP}** rows only — company-wide sheets "
                   "are hidden.")

    # KPI row from the per-order sheet, when the report has one.
    ord_sheet = next((s for s in ("Orders — Net Margin", "Order Detail")
                      if s in sheets), None)
    if ord_sheet:
        od = sheets[ord_sheet]

        def _s(col):
            return od[col].sum() if col in od.columns else 0.0

        k = st.columns(6)
        k[0].metric("Orders", f"{len(od):,}")
        if ord_sheet == "Orders — Net Margin":
            ns, gpv, npv = _s("Net Sales"), _s("Gross Profit ($)"), _s("Net Profit ($)")
            k[1].metric("Net Sales", money(ns))
            k[2].metric("Gross Profit", money(gpv))
            k[3].metric("Net Profit", money(npv))
            k[4].metric("GP Margin", f"{(gpv / ns * 100) if ns else 0:,.2f}%")
            k[5].metric("SqFt (Stone)", f"{_s('Sq Ft (Stone)'):,.2f}")
        else:
            inc, pr = _s("Income"), _s("Profit")
            k[1].metric("Job Income", money(inc))
            k[2].metric("Material Cost", money(_s("MaterialCost")))
            k[3].metric("Operational Cost", money(_s("OperationalCost")))
            k[4].metric("Profit", money(pr))
            k[5].metric("Margin", f"{(pr / inc * 100) if inc else 0:,.2f}%")

    tabs = st.tabs([_tab_icon(s) + s for s in sheets])
    for tab, (sname, sdf) in zip(tabs, sheets.items()):
        with tab:
            lvl = sdf["Level"] if "Level" in sdf.columns else None
            disp = sdf.drop(columns=["Level"]) if lvl is not None else sdf
            disp = disp.rename(columns=_VIEW_RENAME)
            first = disp.iloc[:, 0].astype(str).str.strip().str.lower()

            def rowstyle(row, _lvl=lvl, _first=first):
                if _lvl is not None:
                    v = _lvl.iloc[row.name]
                    if v == "Total":
                        return [_ROW_DARK] * len(row)
                    if v != "Rep":
                        return [_ROW_GROUP] * len(row)
                if _first.iloc[row.name] in ("total", "grand total"):
                    return [_ROW_DARK] * len(row)
                return [""] * len(row)

            st.dataframe(disp.style.format(_sheet_fmt(disp), na_rep="–")
                         .apply(rowstyle, axis=1),
                         width="stretch", hide_index=True,
                         height=min(60 + 35 * len(disp), TABLE_H))
            st.caption(f"{len(disp):,} rows")


# --------------------------------------------------------------------------
# Sidebar — inputs
# --------------------------------------------------------------------------
st.sidebar.title("📊 Gross Profit")
st.sidebar.caption("MC Granite")

uc1, uc2 = st.sidebar.columns([3, 1.4])
uc1.markdown(f"👤 **{AUTH_USER.title()}** · "
             f"{'All shops' if IS_ADMIN else USER_SHOP + ' only'}")
if uc2.button("Log out"):
    st.session_state.pop("auth_user", None)
    st.rerun()
st.sidebar.divider()

source = st.sidebar.radio(
    "Load data from", ["Upload files", "Folder path", "Saved report"], horizontal=True,
    help="**Saved report** re-opens an Excel report you downloaded from this app "
         "before — e.g. last month's — without needing the source export files.")
blobs = ()
if source == "Saved report":
    rups = st.sidebar.file_uploader(
        "Saved report (.xlsx)", type=["xlsx"], accept_multiple_files=True,
        help="Upload one or more reports downloaded from this app — either the "
             "Summary-tab report or the full one-sheet-per-tab report. With several "
             "uploaded, pick which one to view below.")
    if not rups:
        st.title("📂 Saved Report Viewer")
        st.info("👈 Upload a report you downloaded from this app earlier — e.g. "
                "**Gross Profit 2026-05 by Shop.xlsx** — to review a past period "
                "without reloading the source export files.")
        st.stop()
    rnames = [u.name for u in rups]
    rsel = st.sidebar.selectbox("Report to view", rnames) if len(rnames) > 1 else rnames[0]
    rblob = next(u for u in rups if u.name == rsel).getvalue()
    try:
        book = _load_report(rblob)
    except Exception as e:  # noqa: BLE001
        st.error(f"Couldn't read that Excel file — is it a report downloaded from "
                 f"this app? ({e})")
        st.stop()
    render_saved_report(rsel, book)
    st.stop()
if source == "Upload files":
    ups = st.sidebar.file_uploader(
        "Drop your export files",
        type=["csv", "xlsx", "xlsm", "xls", "xlsb"],
        accept_multiple_files=True,
        help="Sales By SKU · Invoice List · Inventory Allocation · Sales Person Summary · "
        "Customer (ReportAdHoc) · Products · Template Schedule · Install Schedule (.xlsb) · "
        "Production LOG (operational cost) · and optionally a rep→shop→manager mapping file.",
    )
    if ups:
        blobs = tuple((u.name, u.getvalue()) for u in ups)
else:
    folder = st.sidebar.text_input("Folder with the files",
                                   value="/Users/juancardona/Downloads/gross_profit_inputs")
    extra = st.sidebar.text_input("Extra files (schedules / mapping), comma-separated", value="")
    paths = []
    if folder and os.path.isdir(folder):
        paths += sorted(glob.glob(os.path.join(folder, "*")))
    paths += [p.strip() for p in extra.split(",") if p.strip() and os.path.isfile(p.strip())]
    blobs = tuple((os.path.basename(p), open(p, "rb").read()) for p in paths)

REQUIRED = {
    "sales_by_sku": "Sales By SKU", "invoices": "Invoice List",
    "sales_person": "Sales Person Summary", "allocations": "Inventory Allocation",
    "customers": "Customers (ReportAdHoc)", "products": "Products",
}
OPTIONAL = {
    "template_cost": "Template Schedule (template cost)",
    "install_cost": "Install Schedule (install cost)",
    "fabrication_cost": "Production LOG (fabrication cost)",
    "rep_map": "Shop/Manager mapping",
}

if not blobs:
    st.title("Dynamic Gross Profit Report")
    st.info("👈 Load your export files to begin. The app recognises each file by its "
            "columns, so names can vary.")
    c = st.columns(2)
    c[0].markdown("**Required**\n\n" + "\n".join(f"- {v}" for v in REQUIRED.values()))
    c[1].markdown("**Optional**\n\n" + "\n".join(f"- {v}" for v in OPTIONAL.values()))
    st.stop()

data, notes = _load(blobs)
detected = set(data)

with st.sidebar.expander("Detected files", expanded=bool(REQUIRED.keys() - detected)):
    for k, label in {**REQUIRED, **OPTIONAL}.items():
        ok = k in detected
        if ok and isinstance(data[k], dict) and "map" in data[k]:
            extra = f" · {data[k]['map'].size:,} orders"
        elif ok:
            extra = f" · {len(data[k]):,} rows"
        else:
            extra = ""
        st.markdown(f"{'✅' if ok else '⬜'} {label}{extra}")
    for n in notes:
        st.caption("• " + n)

critical = {"sales_by_sku", "invoices", "sales_person"}
if critical - detected:
    st.error("Missing essential file(s): " + ", ".join(REQUIRED[k] for k in critical - detected))
    st.stop()

# --------------------------------------------------------------------------
# Sidebar — settings
# --------------------------------------------------------------------------
st.sidebar.divider()
st.sidebar.subheader("Settings")

basis_label = st.sidebar.radio("Income basis", ["Billed (Total − Credit Memo − Tax)", "Amount Paid"], index=0)
income_basis = "paid" if basis_label == "Amount Paid" else "billed"
overhead_rate = st.sidebar.number_input(
    "Fixed cost $/SqFt (stone)", value=0.00, min_value=0.00, step=0.01, format="%.2f",
    help="Enter your fixed-cost allocation per stone sqft — nothing is applied "
    "until you type a rate (the example Net Margin report used $28.36). "
    "At 0 the report shows Gross Profit only.")
if overhead_rate == 0:
    st.sidebar.caption("⚠️ Fixed cost rate is 0 — Net Profit = Gross Profit. "
                       "Enter a $/SqFt rate above to allocate fixed costs.")

# ---- Shop / Manager mapping: persisted + editable in-app -------------------
# Discover every rep that appears in the data, plus saved + default + uploaded.
disc_reps = set()
if "sales_person" in data:
    disc_reps |= {r for r in data["sales_person"]["SalesPersonName"].dropna().astype(str).str.strip()
                  if r and r.lower() not in ("nan", "none")}
saved_map = gp.load_mapping()
file_map = {}
if "rep_map" in data:
    sm, mm = gp.parse_rep_map(data["rep_map"])
    for r in set(sm) | set(mm):
        file_map[r] = {"shop": sm.get(r, ""), "manager": mm.get(r, "")}

all_reps = sorted(disc_reps | set(saved_map) | set(gp.DEFAULT_SHOP_MAP) | set(file_map))
if "mapping" not in st.session_state:
    st.session_state["mapping"] = {}
mp = st.session_state["mapping"]
for r in all_reps:                       # add any rep not yet in the working map
    if r not in mp:
        mp[r] = {
            "shop": saved_map.get(r, {}).get("shop") or file_map.get(r, {}).get("shop")
                    or gp.DEFAULT_SHOP_MAP.get(r, ""),
            "manager": saved_map.get(r, {}).get("manager") or file_map.get(r, {}).get("manager")
                       or gp.DEFAULT_MANAGER_MAP.get(r, ""),
        }

shop_map = {r: v["shop"] for r, v in mp.items() if v.get("shop")}
manager_map = {r: v["manager"] for r, v in mp.items() if v.get("manager")}

cfg = gp.Config(shop_map=shop_map, manager_map=manager_map,
                overhead_rate=overhead_rate, income_basis=income_basis)

probe = gp.compute(data, cfg, period=None)
period_opts = ["All periods"] + probe.periods
idx = len(period_opts) - 1 if probe.periods else 0
period_label = st.sidebar.selectbox("Reporting period (invoice month)", period_opts, index=idx)
period = None if period_label == "All periods" else period_label

group_by = st.sidebar.radio("Group report by", ["Shop", "Manager"], horizontal=True)

rep = gp.compute(data, cfg, period=period)

# Saved in-app order edits overlay the computed numbers, for every user.
overrides = gp.load_order_overrides()
orders_all = gp.apply_order_overrides(rep.orders, overrides, cfg)
rd = rep.rep_detail
if not IS_ADMIN:                 # shop users only ever see their own shop
    _shop_lc = USER_SHOP.lower()
    orders_all = orders_all[
        orders_all["Shop"].astype(str).str.strip().str.lower() == _shop_lc]
    rd = rd[rd["Shop"].astype(str).str.strip().str.lower() == _shop_lc]

# Filters
st.sidebar.subheader("Filters")
if IS_ADMIN:
    shops = sorted(rd["Shop"].unique())
    sel_shops = st.sidebar.multiselect("Shop", shops, default=shops)
else:
    sel_shops = sorted(orders_all["Shop"].unique())
    st.sidebar.caption(f"🔒 Shop locked to **{USER_SHOP}**")
mgrs = sorted(rd["Manager"].unique())
sel_mgrs = st.sidebar.multiselect("Manager", mgrs, default=mgrs)
reps_av = sorted(rd[rd["Shop"].isin(sel_shops)]["Rep"].unique())
sel_reps = st.sidebar.multiselect("Sales rep", reps_av, default=reps_av)
ctypes = sorted(orders_all["CustomerType"].dropna().unique())
sel_ct = st.sidebar.multiselect("Customer type", ctypes, default=ctypes)

fo = orders_all[
    orders_all["Shop"].isin(sel_shops) & orders_all["Manager"].isin(sel_mgrs)
    & orders_all["Rep"].isin(sel_reps) & orders_all["CustomerType"].isin(sel_ct)
].copy()


def aggregate(o: pd.DataFrame) -> pd.DataFrame:
    g = o.groupby(["Shop", "Manager", "Rep"], as_index=False)[gp.VALUE_COLS].sum()
    g["Margin"] = g["Profit"] / g["Income"].replace(0, pd.NA)
    return g[g[["SqFtBilled", "Income", "MaterialCost", "OperationalCost"]].abs().sum(axis=1) > 0]


frep = aggregate(fo)
fsummary = gp.build_summary(frep, group_by) if len(frep) else frep

# --------------------------------------------------------------------------
# Header + KPIs
# --------------------------------------------------------------------------
st.title("Gross Profit Report")
st.caption(f"Period **{rep.meta['period']}** · income basis **{basis_label}** · "
           f"{len(fo):,} orders · grouped by **{group_by}**")

inc = fo["Income"].sum()
mat = fo["MaterialCost"].sum()
op = fo["OperationalCost"].sum()
oh = fo["Overhead"].sum()
profit = inc - mat - op - oh
k = st.columns(6)
k[0].metric("Job Income", money(inc))
k[1].metric("Material Cost", money(mat))
k[2].metric("Operational Cost", money(op))
k[3].metric("Gross Profit", money(profit))
k[4].metric("Margin", f"{(profit/inc*100) if inc else 0:,.2f}%")
k[5].metric("SqFt Billed", f"{fo['SqFtBilled'].sum():,.2f}")

# Data-quality banners — company-wide diagnostics, admin only.
op_diag = rep.meta["op_diag"]
if IS_ADMIN:
    if not rep.meta["has_opcost"]:
        st.info("ℹ️ No operational-cost files loaded → Operational Cost is $0. Add the "
                "**Template Schedule**, **Install Schedule** (.xlsb) and **Production LOG** to "
                "capture template / installation / fabrication cost.")
    else:
        unmatched = [v["label"] for v in op_diag.values()
                     if v["loaded"] and v["matched_orders"] == 0]
        if unmatched:
            st.warning("⚠️ " + ", ".join(unmatched) + " cost file(s) loaded but **0 orders "
                       "matched** this period — check the order numbers line up with the "
                       "invoiced OrderIDs, or that the file's dates overlap the selected period.")
    if rep.meta["has_allocations"]:
        n_ord = rep.meta["n_orders"]
        matched = rep.meta["alloc_matched_orders"]
        if matched < n_ord:
            st.warning(
                f"⚠️ Only **{matched} of {n_ord}** invoiced orders this period have matching "
                f"**Allocation** lines ({rep.meta['alloc_matched_rows']} of {rep.meta['alloc_rows']} "
                f"allocation rows matched on **OrderID**). Material Cost is $0 for the rest. "
                f"This usually means the Allocation export covers a **different date range** than the "
                f"invoices — material is allocated before the order is invoiced, so the two exports "
                f"must span the same orders.")
        with st.expander("🔎 Allocation diagnostics (why is Material Cost what it is?)"):
            m = rep.meta
            c1, c2, c3 = st.columns(3)
            c1.metric("Material cost IN FILE (before matching)", f"${m['alloc_linecost_all']:,.2f}")
            c2.metric("Material cost MATCHED to this period", f"${m['material_total']:,.2f}")
            c3.metric("Orders matched", f"{m['alloc_matched_orders']} / {m['n_orders']}")
            if m["alloc_linecost_all"] < 1:
                st.error("The allocation file itself sums to ~$0 of material cost. This is NOT a "
                         "date/matching issue — the **ExtendedCost / UnitCost columns aren't being "
                         "read**. Check the column headers in the allocation export.")
            elif m["material_total"] < 0.5 * m["alloc_linecost_all"]:
                st.warning("The file HAS material cost, but most of it doesn't match this period's "
                           "orders — so it's an **OrderID overlap** problem (different date range or "
                           "ID format), not a formula problem. Compare the two ID samples below.")
            st.caption(f"Allocation file: {m['alloc_rows']:,} rows, "
                       f"{m['alloc_distinct_orders']:,} distinct OrderIDs.")
            st.write("**Sample allocation OrderIDs:**", m.get("alloc_id_sample", []))
            st.write("**Sample invoice OrderIDs (this period):**", m.get("inv_id_sample", []))
if group_by == "Manager" and not rep.meta["has_manager_map"]:
    st.info("ℹ️ No managers assigned yet — open the **⚙️ Shops & Managers** tab, fill in the "
            "Manager column, and Save. Then the Manager grouping populates.")

# --------------------------------------------------------------------------
# Tabs
# --------------------------------------------------------------------------
# Customer-type table over the filtered orders — used by the Customer Type tab
# and both Excel exports (for shop users it is scoped to their shop).
ct_table = fo.groupby("CustomerType", as_index=False).agg(
    Income=("Income", "sum"), Profit=("Profit", "sum"), Orders=("OrderID", "count")
).sort_values("Income", ascending=False)

# One canonical per-order export frame — the original workbook's
# "Orders — Net Margin" layout over ALL sidebar-scoped orders. Both Excel
# downloads use this full list; the Orders tab's quick filters (order #, rep,
# shop) only narrow the table on screen, never the downloaded workbook.
nm_export = (build_net_margin(fo)
             .sort_values(["Shop", "NetSales"], ascending=[True, False])
             [[c for c, _ in NET_MARGIN_COLS]].reset_index(drop=True)
             .rename(columns=dict(NET_MARGIN_COLS)))

if IS_ADMIN:
    (tab_summary, tab_orders, tab_charts, tab_drill, tab_material, tab_ct,
     tab_op, tab_map) = st.tabs(
        ["📋 Summary", "📑 Orders", "📈 Charts", "🔎 Drill-down", "🪨 Material",
         "👥 Customer Type", "🛠️ Operational", "⚙️ Shops & Managers"])
else:
    # Shop users: the company-wide tabs (Material, Operational) and the
    # mapping editor are admin-only.
    tab_summary, tab_orders, tab_charts, tab_drill, tab_ct = st.tabs(
        ["📋 Summary", "📑 Orders", "📈 Charts", "🔎 Drill-down", "👥 Customer Type"])
    tab_material = tab_op = tab_map = None

LABELS = {"SqFtBilled": "SqFt Billed", "SqFtAllocated": "SqFt Alloc.", "Income": "Income",
          "MaterialCost": "Material", "TemplateCost": "Template", "InstallCost": "Install",
          "FabricationCost": "Fabrication", "OperationalCost": "Operational",
          "Overhead": "Overhead", "TotalCost": "Total Cost", "Profit": "Profit",
          "Margin": "Margin"}
# Same columns, same order, as the exported Summary sheet.
SHOWN = ["SqFtBilled", "SqFtAllocated", "Income", "MaterialCost", "TemplateCost",
         "InstallCost", "FabricationCost", "OperationalCost", "Overhead",
         "TotalCost", "Profit", "Margin"]


def style_summary(df: pd.DataFrame):
    out = df[["Name"] + SHOWN].rename(columns=LABELS).rename(columns={"Name": f"{group_by} / Rep"})
    fmt = {"SqFt Billed": "{:,.2f}", "SqFt Alloc.": "{:,.2f}"}
    fmt.update({LABELS[c]: "${:,.2f}" for c in ["Income", "MaterialCost", "TemplateCost",
                                                "InstallCost", "FabricationCost",
                                                "OperationalCost", "Overhead",
                                                "TotalCost", "Profit"]})
    fmt["Margin"] = "{:.2%}"

    def rowstyle(row):
        lvl = df.iloc[row.name]["Level"]
        if lvl == "Total":
            return ["background-color:#1f4e79;color:white;font-weight:700"] * len(row)
        if lvl in ("Shop", "Manager"):
            return ["background-color:#dde7f2;font-weight:700"] * len(row)
        return [""] * len(row)

    return (out.style.format(fmt, na_rep="–").apply(rowstyle, axis=1)
            .map(lambda v: "color:#c00" if isinstance(v, (int, float)) and v < 0 else "",
                 subset=["Profit"]))


with tab_summary:
    if len(fsummary):
        st.dataframe(style_summary(fsummary), width="stretch", hide_index=True,
                     height=min(60 + 35 * len(fsummary), TABLE_H))

        report_xlsx = build_report_excel(
            nm_export, fsummary,
            rep.by_material.rename(columns={"AvgPrice": "Avg $/unit"})
            if IS_ADMIN else pd.DataFrame(),
            ct_table,
            rep.by_service_group if IS_ADMIN else pd.DataFrame(),
            group_by, fixed_rate=overhead_rate, income_basis=income_basis,
            period=rep.meta["period"], basis_label=basis_label)
        st.download_button("⬇️ Download report (Excel)", data=report_xlsx,
                           file_name=f"Gross Profit {rep.meta['period']} by {group_by}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No data for the current filters.")

with tab_orders:
    st.caption("One row per order, in the **Orders — Net Margin** layout of the example "
               "workbook: revenue → direct costs → Gross Profit → fixed cost → Net Profit. "
               "Sort any column by clicking its header.")

    fc1, fc2, fc3, fc4 = st.columns([2, 2.6, 2.6, 2.2])
    order_q = fc1.text_input("🔎 Filter by order number",
                             placeholder="type part of an order #, e.g. 76926")
    all_orders = sorted(fo["OrderID"].astype(str).unique())
    pick = fc2.multiselect("…or pick specific orders", all_orders,
                           help="Leave empty to show all orders.")
    reps_here = sorted(fo["Rep"].dropna().astype(str).unique())
    pick_reps = fc3.multiselect("…or pick sales person(s)", reps_here,
                                help="Leave empty to show all sales people.")
    shops_here = sorted(fo["Shop"].dropna().astype(str).unique())
    pick_shops = fc4.multiselect("…or pick shop(s)", shops_here,
                                 help="Leave empty to show all shops. Filters "
                                      "here only narrow the table on screen — "
                                      "the Excel downloads always contain the "
                                      "full list.")
    base = fo
    if order_q.strip():
        base = base[base["OrderID"].astype(str).str.contains(
            order_q.strip(), case=False, na=False, regex=False)]
    if pick:
        base = base[base["OrderID"].astype(str).isin(pick)]
    if pick_reps:
        base = base[base["Rep"].astype(str).isin(pick_reps)]
    if pick_shops:
        base = base[base["Shop"].astype(str).isin(pick_shops)]

    if len(base):
        nm = build_net_margin(base).sort_values(["Shop", "NetSales"],
                                                ascending=[True, False])
        internal = [c for c, _ in NET_MARGIN_COLS]
        ordr = nm[internal].reset_index(drop=True)

        if msg := st.session_state.pop("edit_msg", None):
            st.success(msg)
        edit_mode = st.toggle(
            "✏️ **Edit mode** — type new values directly in the table",
            key="orders_edit_mode",
            help="Editable: Total Invoice, Sales Tax, the two Sq Ft columns and the four "
                 "direct costs. Net Sales, Gross Profit, Net Profit and the margins "
                 "recompute when you Save."
                 + ("" if IS_ADMIN else f" You can edit {USER_SHOP} orders only."))

        if edit_mode:
            # ---- Editable grid: same layout, the base columns accept typing ----
            st.caption("Change any highlighted-column value, then click **💾 Save edits**. "
                       "Edits are stored locally in `order_overrides.json` and apply for "
                       "every user; the source export files are never modified. The derived "
                       "columns (Net Sales → margins) refresh after saving.")
            EDIT_COLS = gp.EDITABLE_ORDER_FIELDS
            ed_cols = internal[:1] + ["Edited"] + internal[1:]
            ed_base = nm[ed_cols].reset_index(drop=True)
            nm_labels = dict(NET_MARGIN_COLS)
            colcfg = {"Edited": st.column_config.CheckboxColumn(
                "Edited?", help="Order has saved edits")}
            for c in internal:
                label = nm_labels[c]
                if c in NM_MONEY:
                    hl = "✏️ " if c in EDIT_COLS else ""
                    colcfg[c] = st.column_config.NumberColumn(hl + label, format="$%.2f",
                                                              step=0.01)
                elif c in NM_SQFT:
                    colcfg[c] = st.column_config.NumberColumn("✏️ " + label, format="%.2f",
                                                              step=0.01)
                elif c in NM_PCT:
                    colcfg[c] = st.column_config.NumberColumn(label, format="percent")
                else:
                    colcfg[c] = st.column_config.TextColumn(label)
            edited_df = st.data_editor(
                ed_base, key="orders_editor", hide_index=True, width="stretch",
                column_config=colcfg,
                disabled=[c for c in ed_cols if c not in EDIT_COLS],
                height=min(60 + 35 * len(ed_base), TABLE_H))

            changes = {}
            for i in range(len(ed_base)):
                oid = str(ed_base.at[i, "OrderID"])
                for c in EDIT_COLS:
                    old = float(ed_base.at[i, c]) if pd.notna(ed_base.at[i, c]) else 0.0
                    new = edited_df.at[i, c]
                    if pd.notna(new) and abs(float(new) - old) > 1e-9:
                        changes.setdefault(oid, {})[c] = float(new)

            stored = gp.load_order_overrides()
            shown_ids = {str(x) for x in ed_base["OrderID"]}
            n_saved = sum(1 for k in stored if k in shown_ids)
            b1, b2, b3 = st.columns([1.3, 2.4, 3])
            if b1.button("💾 Save edits", type="primary", disabled=not changes):
                now = datetime.now().isoformat(timespec="seconds")
                for oid, fields in changes.items():
                    entry = stored.get(oid, {"fields": {}})
                    entry.setdefault("fields", {}).update(fields)
                    entry["by"] = AUTH_USER
                    entry["at"] = now
                    stored[oid] = entry
                gp.save_order_overrides(stored)
                st.session_state.pop("orders_editor", None)
                st.session_state["edit_msg"] = f"Saved edits to {len(changes)} order(s)."
                st.rerun()
            if b2.button(f"🗑️ Remove saved edits for shown orders ({n_saved})",
                         disabled=not n_saved):
                for k in list(stored):
                    if k in shown_ids:
                        del stored[k]
                gp.save_order_overrides(stored)
                st.session_state.pop("orders_editor", None)
                st.session_state["edit_msg"] = (
                    f"Removed saved edits from {n_saved} order(s).")
                st.rerun()
            if changes:
                b3.info(f"✏️ {len(changes)} order(s) changed — click **Save edits** "
                        "to keep them.")
        else:
            # ---- Read-only view with TOTAL row ----
            sum_cols = [c for c in NM_MONEY + NM_SQFT if c in ordr.columns]
            total = {c: "" for c in internal}
            total["OrderID"] = "TOTAL"
            for c in sum_cols:
                total[c] = ordr[c].sum()
            ns_t = ordr["NetSales"].sum()
            sqft_t = ordr["SqFtBilled"].sum()
            total["GPMargin"] = (ordr["GrossProfit"].sum() / ns_t) if ns_t else float("nan")
            total["NetMargin"] = (ordr["NetProfit"].sum() / ns_t) if ns_t else float("nan")
            total["Contribution"] = (ordr["GrossProfit"].sum() / sqft_t) if sqft_t else float("nan")

            full = pd.concat([ordr, pd.DataFrame([total])], ignore_index=True)
            full.insert(0, "#", [str(i + 1) for i in range(len(ordr))] + [""])
            rename = {"#": "#", **{c: lbl for c, lbl in NET_MARGIN_COLS}}
            show = full.rename(columns=rename)

            ofmt = {rename[c]: "${:,.2f}" for c in NM_MONEY}
            ofmt.update({rename[c]: "{:,.2f}" for c in NM_SQFT})
            ofmt.update({rename[c]: "{:.2%}" for c in NM_PCT})

            def order_rowstyle(row):
                if show.iloc[row.name]["Order ID"] == "TOTAL":
                    return ["background-color:#1f4e79;color:white;font-weight:700"] * len(row)
                return [""] * len(row)

            styled = (show.style.format(ofmt, na_rep="–").apply(order_rowstyle, axis=1)
                      .map(lambda v: "color:#c00" if isinstance(v, (int, float)) and v < 0 else "",
                           subset=["Gross Profit ($)", "Net Profit ($)"]))
            st.dataframe(styled, width="stretch", hide_index=True,
                         height=min(60 + 35 * len(show), TABLE_H))
        m = rep.meta
        st.caption(
            f"{len(ordr):,} orders · Net Sales = Total Invoice − Sales Tax (− credit memos) · "
            f"Fixed Cost = Sq Ft (Stone) × **${overhead_rate:,.2f}** · "
            f"Net Profit = Gross Profit − Fixed Cost. "
            f"**Sq Ft (Stone) counts stone slab lines only** "
            f"({m['sqft_stone_total']:,.2f} stone sqft kept, "
            f"{m['sqft_nonstone_qty']:,.2f} non-stone units excluded this period)."
            + (f" {m['sqft_unmatched_qty']:,.2f} sqft on SKUs not in the Products catalog "
               f"are treated as non-stone — check Products coverage if this is large."
               if m['sqft_unmatched_qty'] > 0 else ""))

        # Excel download = the whole dashboard, one sheet per tab — always the
        # FULL order list (the quick filters above are on-screen only).
        mapping_df = pd.DataFrame(
            [{"Sales Rep": r, "Shop": v.get("shop", ""), "Manager": v.get("manager", "")}
             for r, v in sorted(mp.items())]) if IS_ADMIN else pd.DataFrame()
        full_xlsx = build_full_export(
            nm_export, fsummary,
            rep.by_material.rename(columns={"AvgPrice": "Avg $/unit"})
            if IS_ADMIN else pd.DataFrame(),
            rep.by_color.rename(columns={"AvgPrice": "Avg $/SqFt"})
            if IS_ADMIN else pd.DataFrame(),
            ct_table,
            rep.by_service_group if IS_ADMIN else pd.DataFrame(),
            op_diag_table(rep.meta["op_diag"])
            if (IS_ADMIN and rep.meta["has_opcost"]) else None,
            mapping_df, group_by,
            fixed_rate=overhead_rate, income_basis=income_basis,
            period=rep.meta["period"], basis_label=basis_label)
        st.download_button(
            "⬇️ Download full report (Excel) — one sheet per tab", data=full_xlsx,
            file_name=f"Gross Profit Full Report {rep.meta['period']} by {group_by}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # ---- Round-trip: re-upload an edited export to apply edits in bulk ----
        with st.expander("⬆️ Upload an edited export to update orders in bulk"):
            st.caption(
                "Download **either** report above, edit an order sheet (the "
                "**Orders — Net Margin** or **Order Detail** sheet) in Excel, then "
                "upload it back here. Only these columns are read: **Total Invoice, "
                "Sales Tax, Sq Ft (Stone), Sq Ft Allocated, Material / Template / "
                "Install / Fab Cost** — every other column (Net Sales, Gross Profit, "
                "the margins) is recomputed. A cell set back to its original value "
                "clears that edit. Other sheets are ignored and the source export "
                "files are never modified."
                + ("" if IS_ADMIN else f" Only **{USER_SHOP}** orders are applied."))
            up = st.file_uploader("Edited report (.xlsx)", type=["xlsx"],
                                  key="orders_roundtrip")
            if up is not None:
                nm_labels = dict(NET_MARGIN_COLS)
                # Accept both the internal name and the pretty label for each column,
                # so either export (Order Detail / Orders — Net Margin) round-trips.
                id_aliases = {"OrderID", nm_labels["OrderID"]}
                field_aliases = {c: {c, nm_labels[c]} for c in gp.EDITABLE_ORDER_FIELDS}

                def _pick(cols, aliases):
                    for a in aliases:
                        if a in cols:
                            return a
                    return None

                try:
                    book = pd.read_excel(up, sheet_name=None)  # {sheet: DataFrame}
                except Exception as e:  # noqa: BLE001
                    book, err = None, str(e)
                    st.error(f"Couldn't read that Excel file. ({err})")
                sheet_df = id_col = None
                if book:
                    pref = ["Orders — Net Margin", "Order Detail"]
                    ordered = ([s for s in pref if s in book]
                               + [x for x in book if x not in pref])
                    for s in ordered:
                        df = book[s]
                        idc = _pick(df.columns, id_aliases)
                        if idc and any(_pick(df.columns, al)
                                       for al in field_aliases.values()):
                            sheet_df, id_col = df, idc
                            break
                    if sheet_df is None:
                        st.error("None of the sheets have an order-ID column plus at "
                                 "least one editable column — upload a report exported "
                                 "from this app.")
                if sheet_df is not None:
                    col_of = {c: _pick(sheet_df.columns, al)
                              for c, al in field_aliases.items()}
                    col_of = {c: cc for c, cc in col_of.items() if cc}
                    parsed = {}
                    for _, r in sheet_df.iterrows():
                        oid = str(r[id_col]).strip()
                        if not oid or oid.lower() in ("nan", "total"):
                            continue
                        fields = {}
                        for c, cc in col_of.items():
                            val = gp.clean_number(r[cc])
                            if val is not None:
                                fields[c] = val
                        if fields:
                            parsed[oid] = fields
                    allowed = None if IS_ADMIN else set(
                        orders_all["OrderID"].astype(str))
                    merged, info = gp.diff_order_overrides(
                        parsed, rep.orders, gp.load_order_overrides(),
                        allowed_ids=allowed, by=AUTH_USER,
                        at=datetime.now().isoformat(timespec="seconds"))
                    chg = info["changes"]
                    if info["skipped_perm"]:
                        st.caption(f"⚠️ {info['skipped_perm']} order(s) outside "
                                   f"{USER_SHOP} were skipped.")
                    if info["skipped_missing"]:
                        st.caption(f"⚠️ {info['skipped_missing']} order(s) in the file "
                                   "aren't in the current data and were skipped.")
                    if not chg:
                        st.info(f"No changes detected ({len(parsed)} order(s) read).")
                    else:
                        prev = pd.DataFrame([{
                            "Order ID": c["OrderID"],
                            "Field": nm_labels.get(c["field"], c["field"]),
                            "From": c["old"], "To": c["new"],
                            "Change": "↩︎ revert" if c["action"] == "revert"
                            else "✏️ edit"} for c in chg])
                        st.write(f"**{len(chg)} change(s)** across "
                                 f"{prev['Order ID'].nunique()} order(s):")
                        st.dataframe(prev.style.format(
                            {"From": "{:,.2f}", "To": "{:,.2f}"}),
                            hide_index=True, width="stretch")
                        if st.button("✅ Apply these edits", type="primary",
                                     key="apply_roundtrip"):
                            gp.save_order_overrides(merged)
                            st.session_state.pop("orders_roundtrip", None)
                            st.session_state["edit_msg"] = (
                                f"Applied {len(chg)} edit(s) from the uploaded file.")
                            st.rerun()
    elif not len(fo):
        st.info("No orders for the current sidebar filters.")
    else:
        st.info("No order matches the current order-number / sales-person filter — "
                "clear it to see all orders.")

with tab_charts:
    if len(frep):
        gcol = group_by
        c1, c2 = st.columns(2)
        gru = frep.groupby(gcol, as_index=False).agg(Income=("Income", "sum"), Profit=("Profit", "sum"))
        fig = go.Figure()
        fig.add_bar(x=gru[gcol], y=gru["Income"], name="Income", marker_color=PALETTE[1])
        fig.add_bar(x=gru[gcol], y=gru["Profit"], name="Profit", marker_color=PALETTE[5])
        fig.update_layout(barmode="group", title=f"Income vs Profit by {gcol}", height=370,
                          legend=dict(orientation="h"))
        c1.plotly_chart(fig, width="stretch")

        fig2 = px.bar(frep.sort_values("Profit"), x="Profit", y="Rep", orientation="h", color=gcol,
                      color_discrete_sequence=PALETTE, title="Profit by Sales Rep", height=370)
        c2.plotly_chart(fig2, width="stretch")

        c3, c4 = st.columns(2)
        fig3 = px.bar(frep.sort_values("Margin"), x="Margin", y="Rep", orientation="h", color="Margin",
                      color_continuous_scale=["#c00", "#f4f4f4", "#2e8b57"], range_color=[-0.3, 0.6],
                      title="Margin by Sales Rep")
        fig3.update_layout(height=370, xaxis_tickformat=".0%")
        c3.plotly_chart(fig3, width="stretch")

        if IS_ADMIN:        # material mix is computed company-wide
            mat_df = rep.by_material[rep.by_material["Material"].isin(list(gp.STONE_CATALOGS))]
            if len(mat_df):
                fig4 = px.pie(mat_df, values="Revenue", names="Material", hole=0.45,
                              title="Revenue by stone material", color_discrete_sequence=PALETTE)
                fig4.update_layout(height=370)
                c4.plotly_chart(fig4, width="stretch")
    else:
        st.info("No data for the current filters.")

with tab_drill:
    st.caption("Click a sales rep to see the underlying orders.")
    gcol = group_by
    for g in sorted(frep[gcol].unique()):
        st.markdown(f"### {g}")
        for _, prow in frep[frep[gcol] == g].sort_values("Income", ascending=False).iterrows():
            with st.expander(f"**{prow['Rep']}** — Income {money(prow['Income'])} · "
                             f"Profit {money(prow['Profit'])} · Margin {prow['Margin']*100:,.2f}%"):
                od = fo[fo["Rep"] == prow["Rep"]][
                    ["OrderID", "Customer", "CustomerType", "SqFtBilled", "Income",
                     "MaterialCost", "TemplateCost", "InstallCost", "FabricationCost",
                     "OperationalCost", "Profit"]].sort_values("Income", ascending=False)
                st.dataframe(od.style.format({
                    "SqFtBilled": "{:,.2f}", "Income": "${:,.2f}", "MaterialCost": "${:,.2f}",
                    "TemplateCost": "${:,.2f}", "InstallCost": "${:,.2f}",
                    "FabricationCost": "${:,.2f}", "OperationalCost": "${:,.2f}",
                    "Profit": "${:,.2f}"}),
                    width="stretch", hide_index=True)

if IS_ADMIN:
    with tab_material:
        st.subheader("Top countertop colors")
        st.caption("Stone slab lines only, named by the Products **Title** field, ranked by "
                   "revenue. SqFt here is billed stone square footage.")
        topn = rep.by_color.head(21).copy()
        if len(topn):
            topn.insert(0, "Rank", range(1, len(topn) + 1))
            tc = topn.rename(columns={"AvgPrice": "Avg $/SqFt"})
            st.dataframe(tc.style.format({"SqFt": "{:,.2f}", "Revenue": "${:,.2f}",
                                          "Avg $/SqFt": "${:,.2f}"}),
                         width="stretch", hide_index=True,
                         height=min(60 + 35 * len(tc), TABLE_H))
            st.caption(f"Top {len(tc)} of {len(rep.by_color):,} colors sold this period — "
                       "the full list is in the Excel export.")
        else:
            st.info("No stone slab lines this period.")

        st.subheader("By material category")
        st.caption("Material type comes from the Products **Catalogs** field (joined by SKU). "
                   "**Other** = the SKU exists in Products but its Catalogs field is blank. "
                   "Quantity mixes units: sqft for stone, piece counts for sinks/services.")
        bm = rep.by_material
        if len(bm):
            bm_rank = bm.sort_values("Revenue")        # largest ends up on top
            figm = go.Figure(go.Bar(
                x=bm_rank["Revenue"], y=bm_rank["Material"], orientation="h",
                marker_color=PALETTE[1],
                text=[f"${v:,.0f}" for v in bm_rank["Revenue"]],
                textposition="outside", cliponaxis=False,
                hovertemplate="%{y}: $%{x:,.2f}<extra></extra>"))
            figm.update_layout(
                title="Revenue by material category",
                height=max(340, 36 * len(bm_rank) + 120),
                xaxis=dict(tickprefix="$", tickformat=",.0f"),
                margin=dict(r=90))
            st.plotly_chart(figm, width="stretch")

            bmt = bm.rename(columns={"AvgPrice": "Avg $/unit"})
            st.dataframe(bmt.style.format({"Quantity": "{:,.2f}", "Revenue": "${:,.2f}",
                                           "Avg $/unit": "${:,.2f}"}),
                         width="stretch", hide_index=True)
        else:
            st.info("No categorised material lines this period.")
        st.caption("ℹ️ The ERP's generic roll-up lines (the `COUNTERTOPS` SKU and other "
                   "SKUs missing from the Products export) are excluded from this "
                   "report — they restate each job's total next to the itemised lines "
                   "and would double-count revenue.")

with tab_ct:
    ct = ct_table
    if len(ct):
        c1, c2 = st.columns([3, 2])
        fig = px.bar(ct, x="CustomerType", y="Income", color="CustomerType",
                     title="Job Income by Customer Type", color_discrete_sequence=PALETTE)
        fig.update_layout(showlegend=False, height=380)
        c1.plotly_chart(fig, width="stretch")
        fig2 = px.pie(ct, values="Income", names="CustomerType", hole=0.45,
                      color_discrete_sequence=PALETTE, title="Income share")
        fig2.update_layout(height=380)
        c2.plotly_chart(fig2, width="stretch")
        st.dataframe(ct.style.format({"Income": "${:,.2f}", "Profit": "${:,.2f}"}),
                     width="stretch", hide_index=True)
    else:
        st.info("No data for the current filters.")

if IS_ADMIN:
    with tab_op:
        od = rep.meta["op_diag"]
        if not rep.meta["has_opcost"]:
            st.info("No operational-cost files loaded. Add the **Template Schedule**, "
                    "**Install Schedule** (.xlsb) and **Production LOG** to capture "
                    "template / installation / fabrication cost per order. Cost attaches to an "
                    "order when it is invoiced in the selected period.")
        else:
            st.subheader("Operational cost by type")
            if len(rep.by_service_group):
                fig = px.bar(rep.by_service_group, x="Group", y="Amount", color="Group",
                             color_discrete_sequence=PALETTE,
                             title="Operational cost applied this period")
                fig.update_layout(showlegend=False, height=360)
                st.plotly_chart(fig, width="stretch")

            diag_df = op_diag_table(od)

            def diag_rowstyle(row):
                if diag_df.iloc[row.name]["Type"] == "TOTAL":
                    return ["background-color:#1f4e79;color:white;font-weight:700"] * len(row)
                return [""] * len(row)

            st.dataframe(
                diag_df.style.format({"Cost in file": "${:,.2f}",
                                      "Cost applied (period)": "${:,.2f}"})
                .apply(diag_rowstyle, axis=1),
                width="stretch", hide_index=True)
            st.caption("One row per cost type — crew tabs are summed together, not "
                       "broken out. **In file** = everything in the schedule workbook "
                       "(all crew tabs, double-counting tabs excluded). **Applied (period)** "
                       "= only orders invoiced in the selected period — operational cost "
                       "attaches to an order when it is invoiced, the same way material "
                       "cost is counted. Order numbers like `SOU-76926` are matched on "
                       "the numeric part (`76926`).")

    with tab_map:
        st.caption("Assign each sales rep to a **Shop** and a **Manager**. Edits update the report "
                   "immediately; **Save** keeps them for next time. Add a row for any new rep.")
        map_df = pd.DataFrame(
            [{"Sales Rep": r, "Shop": v.get("shop", ""), "Manager": v.get("manager", "")}
             for r, v in sorted(mp.items())]
        )
        edited = st.data_editor(
            map_df, key="map_editor", num_rows="dynamic", width="stretch", hide_index=True,
            column_config={
                "Sales Rep": st.column_config.TextColumn("Sales Rep", required=True),
                "Shop": st.column_config.TextColumn("Shop", help="e.g. Blue Ridge, Jasper, Kennesaw"),
                "Manager": st.column_config.TextColumn("Manager"),
            },
        )
        # Write the edited table back to the working map (drives the report on rerun).
        newmap = {}
        for _, row in edited.iterrows():
            rep_name = str(row.get("Sales Rep", "")).strip()
            if rep_name and rep_name.lower() not in ("nan", "none"):
                newmap[rep_name] = {"shop": str(row.get("Shop") or "").strip(),
                                    "manager": str(row.get("Manager") or "").strip()}
        # keep any discovered rep that a delete would otherwise drop
        for r in all_reps:
            newmap.setdefault(r, {"shop": gp.DEFAULT_SHOP_MAP.get(r, ""),
                                  "manager": gp.DEFAULT_MANAGER_MAP.get(r, "")})
        st.session_state["mapping"] = newmap

        c1, c2, _ = st.columns([1, 1, 3])
        if c1.button("💾 Save mapping", type="primary"):
            gp.save_mapping(newmap)
            st.success(f"Saved {len(newmap)} reps to config.json")
        n_mgr = sum(1 for v in newmap.values() if v.get("manager"))
        n_shop = sum(1 for v in newmap.values() if v.get("shop"))
        c2.metric("Mapped", f"{n_shop} shops · {n_mgr} mgrs")
        if n_mgr == 0:
            st.info("No managers assigned yet — fill the Manager column to enable the **Manager** "
                    "grouping, then Save.")
