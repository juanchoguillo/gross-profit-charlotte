"""
Gross Profit calculation engine for MC Granite  (v2).

Ingests the ERP export files (identified by their column signatures, so file
names can vary) and computes a gross-profit report that can be grouped by
Shop or by Manager, down to the individual sales rep.

Data model (per Juan's spec):
  Income (per order)   = Billed = TotalAmount - CreditMemoTotal - SalesTax
                         on non-cancelled invoices (toggle: Paid = AmountPaid)
  SqFt Billed          = Sales-By-SKU Quantity for the period's invoiced orders,
                         STONE lines only (sinks/edges/services/accessories excluded)
  Material Cost        = Inventory Allocation, per line:
                            stone  -> ExtendedCost          (sqft = AllocationMeasure)
                            other  -> AllocationQty * UnitCost (sqft = 0)
                         stone vs other comes from Products "Catalogs"
  SqFt Allocated       = stone AllocationMeasure
  Operational Cost     = Template + Install + Fabrication cost per order, pulled
                         from the three schedule workbooks (one tab per crew
                         member), joined on the numeric OrderID
  Overhead             = SqFt Billed * overhead_rate (default 0, editable)
  Total Cost           = Material + Operational + Overhead
  Profit               = Income - Total Cost ;  Margin = Profit / Income
  Customer Type        = customer master "Type"
  Material category    = first word of Products "Catalogs"
"""
from __future__ import annotations

import io
import json
import os
import re
import warnings
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# Persisted rep -> {shop, manager} mapping, edited in-app.
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def load_mapping() -> dict:
    """Return {rep: {'shop':.., 'manager':..}} from config.json (empty if none)."""
    try:
        with open(CONFIG_PATH, encoding="utf-8") as fh:
            return json.load(fh).get("reps", {})
    except Exception:  # noqa: BLE001
        return {}


def save_mapping(reps: dict) -> None:
    with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
        json.dump({"reps": reps}, fh, indent=2, ensure_ascii=False)


# Persisted per-order manual edits made in the Orders tab (the source export
# files are never modified — edits overlay them at report time).
OVERRIDES_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "order_overrides.json")

# Order fields that may be edited in-app; every derived column (Net Sales,
# Operational Cost, Overhead, Profit, Margin) is recomputed from these.
EDITABLE_ORDER_FIELDS = [
    "TotalInvoice", "SalesTax", "SqFtBilled", "SqFtAllocated", "MaterialCost",
    "TemplateCost", "InstallCost", "FabricationCost",
]


def load_order_overrides() -> dict:
    """Return {order_id: {"fields": {col: value}, "by": user, "at": iso}}."""
    try:
        with open(OVERRIDES_PATH, encoding="utf-8") as fh:
            return json.load(fh).get("orders", {})
    except Exception:  # noqa: BLE001
        return {}


def save_order_overrides(orders: dict) -> None:
    with open(OVERRIDES_PATH, "w", encoding="utf-8") as fh:
        json.dump({"orders": orders}, fh, indent=2, ensure_ascii=False)


def clean_number(v):
    """Coerce a possibly Excel-formatted cell ('$1,234.50', '(50)', '') to a float,
    or None when the cell is blank / not a number. Used when reading an edited
    export back in — users may leave currency symbols, thousands separators or
    accounting-style negatives in the cells."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("$", "").replace("−", "-")
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").strip()
    if s in ("", "-", "–"):
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    return -f if neg else f


def diff_order_overrides(parsed: dict, base_orders: pd.DataFrame, overrides: dict,
                         allowed_ids=None, by: str = "", at: str = ""):
    """Merge a batch of edits read from an edited export into the saved overrides.

    ``parsed`` is ``{order_id: {editable_field: value}}`` taken from the uploaded
    "Orders — Net Margin" sheet. Each field is compared against the ORIGINAL value
    in ``base_orders`` (the per-order frame *before* any overrides): a value equal
    to the original clears that field's override (a revert); any other value is
    stored as an override. ``allowed_ids`` limits which orders may be touched
    (others are skipped) so a shop user can't edit another shop's orders.

    Returns ``(merged_overrides, info)`` where ``info`` has ``changes`` (a list of
    ``{OrderID, field, old, new, action}`` rows for a preview) plus the
    ``skipped_perm`` / ``skipped_missing`` counts.
    """
    merged = {k: {"fields": dict(v.get("fields", {})),
                  "by": v.get("by", ""), "at": v.get("at", "")}
              for k, v in overrides.items()}
    ids = _norm_id(base_orders["OrderID"])
    allowed = None if allowed_ids is None else set(_norm_id(list(allowed_ids)))
    changes, skipped_perm, skipped_missing = [], 0, 0
    for oid, fields in parsed.items():
        soid = _norm_id([oid]).iloc[0]
        if allowed is not None and soid not in allowed:
            skipped_perm += 1
            continue
        mask = ids == soid
        if not mask.any():
            skipped_missing += 1
            continue
        row = base_orders[mask].iloc[0]
        entry = merged.get(soid, {"fields": {}, "by": "", "at": ""})
        cur = entry.get("fields", {})
        for col, new in fields.items():
            if col not in EDITABLE_ORDER_FIELDS:
                continue
            new = clean_number(new)
            if new is None:
                continue
            orig = float(row[col]) if pd.notna(row[col]) else 0.0
            if abs(new - orig) <= 1e-6:                 # back to original -> revert
                if col in cur:
                    changes.append({"OrderID": soid, "field": col,
                                    "old": cur[col], "new": orig, "action": "revert"})
                    del cur[col]
                continue
            applied = cur.get(col, orig)                # value currently in effect
            if abs(new - float(applied)) <= 1e-6:
                continue                                # no real change
            changes.append({"OrderID": soid, "field": col,
                            "old": applied, "new": new, "action": "edit"})
            cur[col] = new
        if cur:
            entry["fields"] = cur
            entry["by"] = by
            entry["at"] = at
            merged[soid] = entry
        elif soid in merged:
            del merged[soid]                            # all fields reverted
    return merged, {"changes": changes, "skipped_perm": skipped_perm,
                    "skipped_missing": skipped_missing}


def apply_order_overrides(orders: pd.DataFrame, overrides: dict, cfg: Config) -> pd.DataFrame:
    """Overlay saved manual edits onto the per-order frame and recompute the
    derived columns. The credit-memo effect baked into NetSales is preserved:
    original NetSales = TotalInvoice − CreditMemo − SalesTax, so the implied
    credit memo survives an edit to the invoice or tax amount."""
    o = orders.copy()
    o["Edited"] = False
    if not overrides:
        return o
    credit_memo = o["TotalInvoice"] - o["SalesTax"] - o["NetSales"]
    ids = _norm_id(o["OrderID"])
    for oid, entry in overrides.items():
        mask = ids == str(oid)
        if not mask.any():
            continue
        for col, val in entry.get("fields", {}).items():
            if col in EDITABLE_ORDER_FIELDS:
                o.loc[mask, col] = float(val)
        o.loc[mask, "Edited"] = True
    o["NetSales"] = o["TotalInvoice"] - o["SalesTax"] - credit_memo
    if cfg.income_basis == "billed":
        o["Income"] = o["NetSales"]
    o["OperationalCost"] = o["TemplateCost"] + o["InstallCost"] + o["FabricationCost"]
    o["Overhead"] = o["SqFtBilled"] * cfg.overhead_rate
    o["TotalCost"] = o["MaterialCost"] + o["OperationalCost"] + o["Overhead"]
    o["Profit"] = o["Income"] - o["TotalCost"]
    o["Margin"] = np.where(o["Income"] != 0, o["Profit"] / o["Income"], np.nan)
    return o

# --------------------------------------------------------------------------
# Configuration (defaults; rep->shop/manager can be overridden by a map file)
# --------------------------------------------------------------------------

# Sales rep -> Shop (authoritative list from Juan, 2026-07-08 — no mapping
# file needed; still overridable in-app or via an optional mapping upload)
DEFAULT_SHOP_MAP = {
    "April Glenn": "Blue Ridge",
    "Joy Buckner": "Blue Ridge",
    "Carlos Lombera": "Jasper",
    "Cassandra Yoke": "Jasper",
    "Courtney Enix": "Jasper",
    "Robert Steel": "Jasper",
    "Sylvia Berrios": "Jasper",
    "Teresa Lozano": "Jasper",
    "Cynthia Donette": "Kennesaw",
    "Dwayne Parrott": "Kennesaw",
    "Elias Rodriguez": "Kennesaw",
    "Erin Behnke": "Kennesaw",
    "Esther Leal": "Kennesaw",
    "Gabriel Hernandez": "Kennesaw",
    "Karrie Newell": "Kennesaw",
    "Kate Hunt": "Kennesaw",
    "Lorena Soto": "Kennesaw",
    "Luciano Lombera": "Kennesaw",
    "Maryori Velasquez": "Kennesaw",
    "Mel Kavanagh": "Kennesaw",
    "Miguel Lombera": "Kennesaw",
    "Phalyssa Brown": "Kennesaw",
    "Susan Farantatos": "Kennesaw",
}

# Shop -> Manager (authoritative list from Juan, 2026-07-08); each rep inherits
# the manager of their shop.
DEFAULT_SHOP_MANAGERS = {
    "Kennesaw": "Luciano Lombera",
    "Blue Ridge": "Manuel Mogollon",
    "Jasper": "Carlos Lombera",
}
DEFAULT_MANAGER_MAP = {rep: DEFAULT_SHOP_MANAGERS[shop]
                       for rep, shop in DEFAULT_SHOP_MAP.items()}

# Catalogs first word that means "stone slab material" (per Juan's authoritative
# list of stone Catalogs values, 2026-06-17). Compared case-insensitively.
STONE_CATALOGS = {
    "granite", "marble", "porcelain", "quartz", "quartzite", "soapstone",
}

UNASSIGNED = "(Unassigned)"
UNKNOWN = "Unknown"


@dataclass
class Config:
    shop_map: dict = field(default_factory=lambda: dict(DEFAULT_SHOP_MAP))
    manager_map: dict = field(default_factory=lambda: dict(DEFAULT_MANAGER_MAP))  # rep -> manager
    overhead_rate: float = 0.0
    income_basis: str = "billed"                      # "paid" | "billed"


# --------------------------------------------------------------------------
# File identification & loading
# --------------------------------------------------------------------------

def _read_any(src, name: str) -> pd.DataFrame:
    is_xlsx = str(name).lower().endswith((".xlsx", ".xlsm", ".xls"))
    if hasattr(src, "read"):
        data = src.read()
        if is_xlsx:
            return pd.read_excel(io.BytesIO(data), dtype=str)
        return pd.read_csv(io.BytesIO(data), dtype=str, encoding="utf-8-sig", on_bad_lines="skip")
    if is_xlsx:
        return pd.read_excel(src, dtype=str)
    return pd.read_csv(src, dtype=str, encoding="utf-8-sig", on_bad_lines="skip")


_SIGNATURES = {
    "sales_by_sku": {"SKU", "Order ID", "Quantity", "Extended"},
    "invoices": {"TotalAmount", "AmountPaid", "OrderID", "InvoiceDate"},
    "sales_person": {"SalesPersonName", "OrderID", "OrderDate"},
    "allocations": {"ProductSKU", "ExtendedCost", "AllocationQty", "OrderID"},
    "customers": {"CustomerNumber", "Type", "SalesPersonName"},
    "products": {"SKU", "Title", "Catalogs"},
}

# Lookup tables — safe to stack across files (deduped by key downstream).
_MERGE_ROLES = {"sales_person", "customers", "products"}

# Possible column names in a rep -> shop/manager mapping file
_REP_COLS = ["sales rep", "salespersonname", "sales person", "project manager", "rep", "name"]
_SHOP_COLS = ["shop", "branch", "store", "location"]
_MGR_COLS = ["manager", "mgr"]


def _is_rep_map(df: pd.DataFrame) -> bool:
    low = {str(c).strip().lower() for c in df.columns}
    has_mgr = any(c in low for c in _MGR_COLS)
    has_rep = any(c in low for c in _REP_COLS)
    return has_mgr and has_rep


def identify(df: pd.DataFrame) -> str | None:
    if _is_rep_map(df):
        return "rep_map"
    cols = set(df.columns.astype(str))
    best, best_score = None, 0
    for key, needed in _SIGNATURES.items():
        if needed.issubset(cols) and len(needed) > best_score:
            best, best_score = key, len(needed)
    return best


def load_files(sources: list[tuple]) -> tuple[dict, list]:
    collected: dict = {}
    opcost: dict = {}        # role -> {'map': Series, 'diag': {...}}
    notes = []
    for src, name in sources:
        try:
            raw = src.read() if hasattr(src, "read") else open(src, "rb").read()
        except Exception as exc:  # noqa: BLE001
            notes.append(f"{name} (could not read: {exc})")
            continue
        # Operational-cost schedule workbooks (multi-sheet) take a dedicated path.
        try:
            op = parse_opcost_workbook(raw, name)
        except Exception as exc:  # noqa: BLE001
            op = None
            notes.append(f"{name} (op-cost parse failed: {exc})")
        if op is not None:
            role, payload = op
            if role in opcost:                       # 2nd file of same kind -> merge
                opcost[role]["map"] = opcost[role]["map"].add(payload["map"], fill_value=0)
                opcost[role]["diag"]["rows"] += payload["diag"]["rows"]
                opcost[role]["diag"]["sheets"] += payload["diag"]["sheets"]
                opcost[role]["diag"]["total_in_file"] += payload["diag"]["total_in_file"]
                opcost[role]["diag"]["orders_in_file"] = int(opcost[role]["map"].size)
            else:
                opcost[role] = payload
            d = payload["diag"]
            notes.append(f"{d['label']} cost: {d['rows']:,} rows from "
                         f"{len(d['sheets'])} crew tab(s)")
            continue
        try:
            df = _read_any(io.BytesIO(raw), name)
        except Exception as exc:  # noqa: BLE001
            notes.append(f"{name} (could not read: {exc})")
            continue
        role = identify(df)
        if not role:
            notes.append(name)
            continue
        collected.setdefault(role, []).append((name, df))

    found: dict = {}
    for role, items in collected.items():
        if len(items) == 1:
            found[role] = items[0][1]
        elif role in _MERGE_ROLES:
            found[role] = pd.concat([d for _, d in items], ignore_index=True)
            notes.append(f"merged {len(items)} '{role}' files")
        else:
            name, df = max(items, key=lambda it: len(it[1]))
            found[role] = df
            notes.append(f"{role}: used '{name}', ignored others (avoid double-count)")
    found.update(opcost)
    return found, notes


def parse_rep_map(df: pd.DataFrame) -> tuple[dict, dict]:
    """Return (rep->shop, rep->manager) from a mapping file."""
    low = {str(c).strip().lower(): c for c in df.columns}
    rep_c = next((low[c] for c in _REP_COLS if c in low), None)
    shop_c = next((low[c] for c in _SHOP_COLS if c in low), None)
    mgr_c = next((low[c] for c in _MGR_COLS if c in low), None)
    shop_map, mgr_map = {}, {}
    if rep_c is None:
        return shop_map, mgr_map
    for _, r in df.iterrows():
        rep = str(r[rep_c]).strip()
        if not rep or rep.lower() in ("nan", "none"):
            continue
        if shop_c and pd.notna(r[shop_c]) and str(r[shop_c]).strip():
            shop_map[rep] = str(r[shop_c]).strip()
        if mgr_c and pd.notna(r[mgr_c]) and str(r[mgr_c]).strip():
            mgr_map[rep] = str(r[mgr_c]).strip()
    return shop_map, mgr_map


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def _num(s) -> pd.Series:
    return pd.to_numeric(
        pd.Series(s).astype(str).str.replace(",", "", regex=False).str.replace("$", "", regex=False),
        errors="coerce",
    )


def _to_dt(s) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")


def _norm_id(s) -> pd.Series:
    """Normalize an ID column for joining across files: string, trimmed, and with
    a trailing '.0' removed (Excel/pandas read numeric IDs as floats -> '75961.0',
    which otherwise fails to match the same ID read as an int/text '75961')."""
    return (pd.Series(s).astype(str).str.strip()
            .str.replace(r"\.0+$", "", regex=True))


def _blankish(v) -> bool:
    return (not isinstance(v, str)) or v.strip() == "" or v.strip().lower() in ("nan", "none")


# --------------------------------------------------------------------------
# Operational-cost schedule workbooks (Template / Install / Fabrication)
# --------------------------------------------------------------------------
# These three are multi-sheet workbooks with one tab per crew member; the
# per-order cost is summed across every qualifying tab (and across repeat visits
# to the same order). They don't fit the single-sheet column-signature path, so
# they get a dedicated reader. Aggregate/summary tabs are skipped so nothing is
# double-counted.
_OPCOST_SPECS = {
    "template_cost": {
        "label": "Template",
        "order_aliases": ["order no.", "order no", "order #", "order#"],
        # Per Order = Temple $ + mileage + other fees = all-in paid to templater.
        "value_aliases": ["per order", "temple $", "template $"],
        "guard_aliases": ["per order", "temple $", "template $"],
    },
    "install_cost": {
        "label": "Install",
        "order_aliases": ["order #", "order#", "order no.", "order no"],
        "value_aliases": ["total"],
        # "SQF Paid" only exists on the real per-installer tabs, not the overview.
        "guard_aliases": ["sqf paid"],
    },
    "fabrication_cost": {
        "label": "Fabrication",
        "order_aliases": ["order"],
        "value_aliases": ["total amount"],
        "guard_aliases": ["total amount"],
        "year_sheets_only": True,        # only the "2026"-style production tab
    },
}

# Tabs that aggregate/duplicate the per-person tabs -> skip to avoid double count.
_OPCOST_SKIP_SHEET = re.compile(
    r"^\s*sheet\s*\d+\s*$|total\s*payment|consolidate|production\s*list", re.I)
_YEAR_SHEET = re.compile(r"^\s*20\d\d\s*$")


def _order_key(v) -> str | None:
    """An order reference -> the numeric OrderID used across the ERP exports.
    'SOU-76926' -> '76926', 'K4P-74213' -> '74213', '76926' -> '76926'."""
    nums = re.findall(r"\d+", str(v))
    return nums[-1] if nums else None


def _excel_engine(name: str):
    return "pyxlsb" if str(name).lower().endswith(".xlsb") else None


def _pick_col(df: pd.DataFrame, aliases):
    low = {str(c).strip().lower(): c for c in df.columns}
    for a in aliases:
        if a in low:
            return low[a]
    return None


def _find_header_row(book, sheet, engine, order_aliases) -> int | None:
    """Find the row that holds the column headers (these workbooks often carry a
    banner row, e.g. '2026', above the real header)."""
    raw = pd.read_excel(book, sheet, header=None, nrows=8, engine=engine)
    for i in range(len(raw)):
        vals = {str(x).strip().lower() for x in raw.iloc[i].tolist()}
        if any(a in vals for a in order_aliases):
            return i
    return None


def _classify_opcost(book, engine) -> str | None:
    """Decide whether a workbook is the Template / Install / Fabrication schedule
    by sniffing the column names present on any of its sheets."""
    cols = set()
    for sheet in book.sheet_names[:30]:
        try:
            raw = pd.read_excel(book, sheet, header=None, nrows=8, engine=engine)
        except Exception:  # noqa: BLE001
            continue
        for _, row in raw.iterrows():
            cols |= {str(x).strip().lower() for x in row.tolist()}

    def has(*opts):
        return any(o in cols for o in opts)

    # Most specific signatures first.
    if has("temple $", "template $") or (has("order no.", "order no") and "per order" in cols):
        return "template_cost"
    if "sqf paid" in cols and has("order #", "order#") and "total" in cols:
        return "install_cost"
    if "total amount" in cols and "order" in cols:
        return "fabrication_cost"
    return None


def parse_opcost_workbook(raw: bytes, name: str):
    """If `raw` is one of the operational-cost schedule workbooks, return
    (role, payload) where payload = {'map': Series(order_key -> cost),
    'diag': {...}}.  Returns None for anything else."""
    if not str(name).lower().endswith((".xlsx", ".xlsm", ".xls", ".xlsb")):
        return None
    engine = _excel_engine(name)
    try:
        book = pd.ExcelFile(io.BytesIO(raw), engine=engine)
    except Exception:  # noqa: BLE001
        return None
    role = _classify_opcost(book, engine)
    if role is None:
        return None
    spec = _OPCOST_SPECS[role]

    per_order: dict = {}
    sheets_used, rows_used, total_in_file = [], 0, 0.0
    for sheet in book.sheet_names:
        if _OPCOST_SKIP_SHEET.search(str(sheet)):
            continue
        if spec.get("year_sheets_only") and not _YEAR_SHEET.match(str(sheet)):
            continue
        h = _find_header_row(book, sheet, engine, spec["order_aliases"])
        if h is None:
            continue
        df = pd.read_excel(book, sheet, header=h, engine=engine)
        df.columns = [str(c).strip() for c in df.columns]
        oc = _pick_col(df, spec["order_aliases"])
        vc = _pick_col(df, spec["value_aliases"])
        if oc is None or vc is None or _pick_col(df, spec["guard_aliases"]) is None:
            continue
        keys = df[oc].map(_order_key)
        vals = _num(df[vc])
        keep = keys.notna() & vals.notna()
        for k, v in zip(keys[keep], vals[keep]):
            per_order[k] = per_order.get(k, 0.0) + float(v)
        sheets_used.append(str(sheet))
        rows_used += int(keep.sum())
        total_in_file += float(vals[keep].sum())

    if not per_order:
        return None
    series = pd.Series(per_order, dtype=float)
    diag = {"label": spec["label"], "sheets": sheets_used, "rows": rows_used,
            "orders_in_file": int(series.size), "total_in_file": total_in_file}
    return role, {"map": series, "diag": diag}


# --------------------------------------------------------------------------
# Lookups
# --------------------------------------------------------------------------

def _sku_prefix(sku: str) -> str:
    """Leading alphabetic run of a SKU, capped at 4 chars (e.g. QRTZTAMALEA -> QRTZ).

    Used to infer stone-vs-not for sale SKUs that are missing from the Products
    master, by matching the prefix against prefixes that are unambiguously stone
    in the master."""
    m = re.match(r"[A-Za-z]+", str(sku).strip())
    return m.group(0)[:4].upper() if m else ""


def build_maps(data: dict, cfg: Config):
    sps = data["sales_person"].copy()
    sps["OrderID"] = _norm_id(sps["OrderID"])
    sps = sps[sps["OrderID"].ne("")]
    sps["_blank"] = sps["SalesPersonName"].fillna("").astype(str).str.strip().eq("")
    sps = sps.sort_values("_blank").drop_duplicates("OrderID", keep="first")
    order2rep = sps.set_index("OrderID")["SalesPersonName"].astype(str).str.strip()
    order2cust = _norm_id(sps.set_index("OrderID")["CustomerNumber"])

    cust2type, cust2rep, cust2name = {}, {}, {}
    if "customers" in data:
        c = data["customers"].copy()
        c["CustomerNumber"] = _norm_id(c["CustomerNumber"])
        # Drop rows with a blank/NaN customer number — otherwise a junk row keyed on
        # "" pollutes the lookups and every order with no customer number resolves to
        # that one bogus name/type.
        c = c[c["CustomerNumber"].ne("") & ~c["CustomerNumber"].str.lower().isin(["nan", "none"])]
        c = c.drop_duplicates("CustomerNumber")
        def clean_type(v):
            return UNKNOWN if _blankish(v) else str(v).strip()
        cust2type = {k: clean_type(v) for k, v in zip(c["CustomerNumber"], c["Type"])}
        cust2rep = c.set_index("CustomerNumber")["SalesPersonName"].astype(str).str.strip().to_dict()
        if "CustomerFullName" in c:
            cust2name = dict(zip(c["CustomerNumber"], c["CustomerFullName"].astype(str)))

    sku2cat, sku2mat, sku2stone, sku2title = {}, {}, {}, {}
    stone_prefixes: set = set()
    if "products" in data:
        p = data["products"].copy()
        p["SKU"] = p["SKU"].astype(str).str.strip()
        p = p.drop_duplicates("SKU")
        sku2title = {s: str(t).strip() for s, t in zip(p["SKU"], p["Title"].fillna(""))
                     if str(t).strip()}
        for sku, cat in zip(p["SKU"], p["Catalogs"].fillna("")):
            head = str(cat).split()[0] if str(cat).strip() else "Other"
            sku2cat[sku] = str(cat)
            sku2mat[sku] = head if head else "Other"
            sku2stone[sku] = head.lower() in STONE_CATALOGS
        # Learn SKU prefixes that are *unambiguously* stone in the master, so a
        # stone SKU that is missing from Products (e.g. QRTZTAMALEA) is still
        # counted toward SqFt Billed instead of being silently dropped. A prefix
        # qualifies only if every product carrying it is stone and at least two
        # products carry it (guards against one-off / typo SKUs).
        pref_stone, pref_total = {}, {}
        for sku, is_stone in sku2stone.items():
            pf = _sku_prefix(sku)
            if not pf:
                continue
            pref_total[pf] = pref_total.get(pf, 0) + 1
            pref_stone[pf] = pref_stone.get(pf, 0) + (1 if is_stone else 0)
        stone_prefixes = {pf for pf, tot in pref_total.items()
                          if tot >= 2 and pref_stone[pf] == tot}
    return (order2rep, order2cust, cust2type, cust2rep, cust2name,
            sku2mat, sku2stone, stone_prefixes, sku2title)


# --------------------------------------------------------------------------
# Core computation
# --------------------------------------------------------------------------

@dataclass
class Report:
    orders: pd.DataFrame
    rep_detail: pd.DataFrame          # one row per (shop, manager, rep)
    by_material: pd.DataFrame
    by_color: pd.DataFrame            # stone colors (Products Title), by revenue
    by_customer_type: pd.DataFrame
    by_service_group: pd.DataFrame
    periods: list
    meta: dict


VALUE_COLS = ["SqFtBilled", "SqFtAllocated", "Income", "MaterialCost",
              "TemplateCost", "InstallCost", "FabricationCost",
              "OperationalCost", "Overhead", "TotalCost", "Profit"]


def compute(data: dict, cfg: Config, period: str | None = None) -> Report:
    (order2rep, order2cust, cust2type, cust2rep, cust2name,
     sku2mat, sku2stone, stone_prefixes, sku2title) = build_maps(data, cfg)

    # ---- Invoices -> Income ----
    inv = data["invoices"].copy()
    inv["OrderID"] = _norm_id(inv["OrderID"])
    inv["InvoiceDate"] = _to_dt(inv["InvoiceDate"])
    inv["Period"] = inv["InvoiceDate"].dt.strftime("%Y-%m")
    status = inv.get("InvoiceStatus", pd.Series("", index=inv.index)).astype(str).str.lower()
    inv = inv[~status.str.contains("cancel", na=False)]
    inv["TotalInvoice"] = _num(inv["TotalAmount"]).fillna(0)
    inv["SalesTaxAmt"] = _num(inv.get("SalesTax", 0)).fillna(0)
    inv["Billed"] = (inv["TotalInvoice"]
                     - _num(inv.get("CreditMemoTotal", 0)).fillna(0)
                     - inv["SalesTaxAmt"])
    inv["Paid"] = _num(inv["AmountPaid"]).fillna(0)
    inv["Income"] = inv["Paid"] if cfg.income_basis == "paid" else inv["Billed"]

    periods = sorted(p for p in inv["Period"].dropna().unique())
    if period:
        inv = inv[inv["Period"] == period]
    month_orders = set(inv["OrderID"].unique())
    income_by_order = inv.groupby("OrderID")["Income"].sum()
    # Net Margin fields: gross invoice, tax and net sales (= Total − CreditMemo − Tax)
    # broken out per order, plus the human order/ref identifiers from the invoice.
    total_invoice_by_order = inv.groupby("OrderID")["TotalInvoice"].sum()
    salestax_by_order = inv.groupby("OrderID")["SalesTaxAmt"].sum()
    netsales_by_order = inv.groupby("OrderID")["Billed"].sum()

    def _first_nonblank(col):
        if col not in inv.columns:
            return pd.Series(dtype=str)
        t = inv[["OrderID", col]].copy()
        t[col] = t[col].astype(str).str.strip()
        t = t[~t[col].str.lower().isin(["", "nan", "none"])]
        return t.drop_duplicates("OrderID").set_index("OrderID")[col]

    ordernum_by_order = _first_nonblank("OrderNumber")
    refnum_by_order = _first_nonblank("RefNumber")
    invname_by_order = _first_nonblank("CustomerFullName")
    # The invoice carries a CustomerNumber for every order; it's the most reliable
    # source (the Sales Person Summary export often doesn't cover the period's orders).
    invcust_by_order = _norm_id(_first_nonblank("CustomerNumber")) \
        if "CustomerNumber" in inv.columns else pd.Series(dtype=str)
    invcust_map = invcust_by_order.to_dict()

    def cust_of(order):
        """Order -> CustomerNumber, preferring the Sales Person Summary, then invoice."""
        c = order2cust.get(order)
        if isinstance(c, str) and not _blankish(c):
            return c
        c = invcust_map.get(order)
        return c if isinstance(c, str) and not _blankish(c) else None

    def attribute(order):
        rep = order2rep.get(order)
        if rep and not _blankish(rep):
            return rep
        rep = cust2rep.get(cust_of(order))
        return rep if rep and not _blankish(rep) else UNASSIGNED

    # ---- Sales By SKU -> SqFt Billed, revenue, material mix ----
    sbs = data["sales_by_sku"].copy()
    sbs["OrderID"] = _norm_id(sbs["Order ID"])
    sbs["SKU"] = sbs["SKU"].astype(str).str.strip()
    sbs["Quantity"] = _num(sbs["Quantity"]).fillna(0)
    sbs["Extended"] = _num(sbs["Extended"]).fillna(0)
    # Material category = first word of the Products "Catalogs" field, joined by
    # SKU. Two distinct "unknown" cases are kept separate so they can be audited:
    #   "Other"            = SKU IS in Products but its Catalogs field is blank
    #   "(Not in Products)" = SKU not found in the Products export at all
    sbs["Material"] = sbs["SKU"].map(sku2mat).fillna("(Not in Products)")
    # Stone vs not, from the Products "Catalogs" join (same join used for Material).
    # SqFt Billed must count ONLY stone slab lines — NOT sinks, edges, services or
    # other accessories, whose Quantity is a unit count, not square footage.
    sbs["InProducts"] = sbs["SKU"].isin(sku2stone)        # matched to the catalog at all
    # Stone if the catalog join says so, OR (for SKUs missing/blank in the
    # master) if the SKU prefix is unambiguously stone in the master. This keeps
    # orphan stone SKUs like QRTZTAMALEA in SqFt Billed instead of dropping them.
    matched_stone = sbs["SKU"].map(sku2stone).fillna(False)
    prefix_stone = sbs["SKU"].map(lambda s: _sku_prefix(s) in stone_prefixes)
    sbs["IsStone"] = matched_stone | prefix_stone
    sbs_m = sbs[sbs["OrderID"].isin(month_orders)].copy()
    sbs_stone = sbs_m[sbs_m["IsStone"]]
    sqft_by_order = sbs_stone.groupby("OrderID")["Quantity"].sum()
    rev_by_order = sbs_m.groupby("OrderID")["Extended"].sum()
    # Diagnostics so the stone filter can be sanity-checked against the raw total.
    sqft_stone_total = float(sbs_stone["Quantity"].sum())
    sqft_all_total = float(sbs_m["Quantity"].sum())
    sqft_unmatched_qty = float(sbs_m.loc[~sbs_m["InProducts"], "Quantity"].sum())

    # ---- Allocations -> Material cost + SqFt allocated ----
    matcost_by_order = pd.Series(dtype=float)
    sqftalloc_by_order = pd.Series(dtype=float)
    al_rows = al_matched_rows = al_matched_orders = 0
    al_linecost_all = 0.0          # total material cost in the file, before order-matching
    al_distinct_orders = 0         # distinct non-blank OrderIDs present in the file
    al_id_sample: list = []        # a few allocation OrderIDs (to eyeball format)
    inv_id_sample: list = []       # a few of this period's invoice OrderIDs
    if "allocations" in data:
        al = data["allocations"].copy()
        al["OrderID"] = _norm_id(al["OrderID"])
        al["ProductSKU"] = al["ProductSKU"].astype(str).str.strip()
        for c in ["ExtendedCost", "AllocationQty", "AllocationMeasure", "UnitCost"]:
            al[c] = _num(al.get(c, 0)).fillna(0)
        # ExtendedCost is the actual extended material cost on every real material
        # line (slabs); only sinks/accessories leave it 0, where AllocationQty x
        # UnitCost is the correct per-unit price. This does NOT depend on joining the
        # allocation's raw-material ProductSKU to the Products catalog (those SKUs are
        # a different namespace than the bundled sale SKUs, so that join misses).
        al["LineCost"] = np.where(al["ExtendedCost"] != 0,
                                  al["ExtendedCost"],
                                  al["AllocationQty"] * al["UnitCost"])
        # SqFt allocated = the line's allocated measure (sqft for slabs; 0 for
        # sinks/accessories, which carry no AllocationMeasure).
        al["LineSqFt"] = al["AllocationMeasure"]
        al_m = al[al["OrderID"].isin(month_orders)]
        matcost_by_order = al_m.groupby("OrderID")["LineCost"].sum()
        sqftalloc_by_order = al_m.groupby("OrderID")["LineSqFt"].sum()
        al_rows = len(al)
        al_matched_rows = len(al_m)
        al_matched_orders = al_m["OrderID"].nunique()
        al_linecost_all = float(al["LineCost"].sum())
        nonblank = al["OrderID"][al["OrderID"].ne("")]
        al_distinct_orders = int(nonblank.nunique())
        al_id_sample = list(nonblank.drop_duplicates().head(10))
        inv_id_sample = sorted(month_orders)[:10]

    # ---- Operational cost = Template + Install + Fabrication, joined by order ----
    def _opmap(role):
        return data[role]["map"] if role in data else pd.Series(dtype=float)
    template_by_order = _opmap("template_cost")
    install_by_order = _opmap("install_cost")
    fab_by_order = _opmap("fabrication_cost")

    # ---- Per-order frame ----
    o = pd.DataFrame({"OrderID": sorted(month_orders)})
    o["Rep"] = o["OrderID"].map(attribute)
    o["Shop"] = o["Rep"].map(cfg.shop_map).fillna(UNASSIGNED)
    o["Manager"] = o["Rep"].map(cfg.manager_map).fillna(UNASSIGNED)
    o["OrderNumber"] = o["OrderID"].map(ordernum_by_order)
    o["RefNumber"] = o["OrderID"].map(refnum_by_order)
    o["CustomerNumber"] = o["OrderID"].map(cust_of)
    o["CustomerType"] = o["CustomerNumber"].map(cust2type).fillna(UNKNOWN)
    o["Customer"] = o["CustomerNumber"].map(cust2name)
    # Fall back to the invoice's customer name when the customer master has none.
    need_name = o["Customer"].isna() | (o["Customer"].astype(str).str.strip() == "")
    o.loc[need_name, "Customer"] = o.loc[need_name, "OrderID"].map(invname_by_order)
    o["Income"] = o["OrderID"].map(income_by_order).fillna(0)
    o["TotalInvoice"] = o["OrderID"].map(total_invoice_by_order).fillna(0)
    o["SalesTax"] = o["OrderID"].map(salestax_by_order).fillna(0)
    o["NetSales"] = o["OrderID"].map(netsales_by_order).fillna(0)
    o["SqFtBilled"] = o["OrderID"].map(sqft_by_order).fillna(0)
    o["Revenue"] = o["OrderID"].map(rev_by_order).fillna(0)
    o["MaterialCost"] = o["OrderID"].map(matcost_by_order).fillna(0)
    o["SqFtAllocated"] = o["OrderID"].map(sqftalloc_by_order).fillna(0)
    o["TemplateCost"] = o["OrderID"].map(template_by_order).fillna(0)
    o["InstallCost"] = o["OrderID"].map(install_by_order).fillna(0)
    o["FabricationCost"] = o["OrderID"].map(fab_by_order).fillna(0)
    o["OperationalCost"] = o["TemplateCost"] + o["InstallCost"] + o["FabricationCost"]
    o["Overhead"] = o["SqFtBilled"] * cfg.overhead_rate
    o["TotalCost"] = o["MaterialCost"] + o["OperationalCost"] + o["Overhead"]
    o["Profit"] = o["Income"] - o["TotalCost"]
    o["Margin"] = np.where(o["Income"] != 0, o["Profit"] / o["Income"], np.nan)

    rep_detail = (o.groupby(["Shop", "Manager", "Rep"], as_index=False)[VALUE_COLS].sum())
    rep_detail["Margin"] = np.where(rep_detail["Income"] != 0,
                                    rep_detail["Profit"] / rep_detail["Income"], np.nan)
    rep_detail = rep_detail[rep_detail[["SqFtBilled", "Income", "MaterialCost", "OperationalCost"]]
                            .abs().sum(axis=1) > 0]

    by_material = (sbs_m.groupby("Material", as_index=False)
                   .agg(Quantity=("Quantity", "sum"), Revenue=("Extended", "sum"))
                   .sort_values("Revenue", ascending=False))
    by_material["AvgPrice"] = np.where(by_material["Quantity"] != 0,
                                       by_material["Revenue"] / by_material["Quantity"], np.nan)
    # Drop the "(Not in Products)" bucket everywhere: it is almost entirely the
    # ERP's generic COUNTERTOPS roll-up lines, which restate each job's total in
    # one lump next to the itemised lines — keeping it would double-count revenue
    # (the bucket is roughly equal to the period's whole Job Income).
    by_material = (by_material[by_material["Material"] != "(Not in Products)"]
                   .reset_index(drop=True))

    # Countertop colors = stone slab lines only, named by the Products "Title"
    # (falls back to the raw SKU for stone SKUs missing from the catalog).
    # Titles carry a service-bundle suffix after ":" ("3CM Taj Mahal Polished:
    # Material, Fabrication & Installation") — strip it so one color = one row.
    colors = sbs_stone.copy()
    colors["Color"] = colors["SKU"].map(
        lambda s: sku2title.get(s, "").split(":")[0].strip() or None)
    colors["Color"] = colors["Color"].where(
        colors["Color"].notna() & colors["Color"].astype(str).str.strip().ne(""),
        colors["SKU"])
    by_color = (colors.groupby(["Color", "Material"], as_index=False)
                .agg(SqFt=("Quantity", "sum"), Revenue=("Extended", "sum"),
                     Orders=("OrderID", "nunique"))
                .sort_values("Revenue", ascending=False).reset_index(drop=True))
    by_color["AvgPrice"] = np.where(by_color["SqFt"] != 0,
                                    by_color["Revenue"] / by_color["SqFt"], np.nan)

    by_ct = o.groupby("CustomerType", as_index=False).agg(
        Income=("Income", "sum"), Profit=("Profit", "sum"), Orders=("OrderID", "count"))
    by_ct = by_ct.sort_values("Income", ascending=False)

    # ---- Operational cost breakdown + match diagnostics ----
    by_service_group = pd.DataFrame(
        [("Template", float(o["TemplateCost"].sum())),
         ("Install", float(o["InstallCost"].sum())),
         ("Fabrication", float(o["FabricationCost"].sum()))],
        columns=["Group", "Amount"])
    by_service_group = (by_service_group[by_service_group["Amount"] != 0]
                        .sort_values("Amount", ascending=False).reset_index(drop=True))

    op_diag = {}
    for role, key, costcol in [("template_cost", "template", "TemplateCost"),
                               ("install_cost", "install", "InstallCost"),
                               ("fabrication_cost", "fabrication", "FabricationCost")]:
        d = data[role]["diag"] if role in data else {}
        op_diag[key] = {
            "label": _OPCOST_SPECS[role]["label"],
            "loaded": role in data,
            "sheets": d.get("sheets", []),
            "rows": d.get("rows", 0),
            "orders_in_file": d.get("orders_in_file", 0),
            "total_in_file": d.get("total_in_file", 0.0),
            "matched_orders": int((o[costcol] != 0).sum()),
            "matched_total": float(o[costcol].sum()),
        }
    has_opcost = any(v["loaded"] for v in op_diag.values())

    meta = {
        "period": period or "All periods",
        "income_basis": cfg.income_basis,
        "overhead_rate": cfg.overhead_rate,
        "n_orders": len(o),
        "has_allocations": "allocations" in data,
        "alloc_rows": al_rows,
        "alloc_matched_rows": al_matched_rows,
        "alloc_matched_orders": al_matched_orders,
        "alloc_linecost_all": al_linecost_all,
        "alloc_distinct_orders": al_distinct_orders,
        "alloc_id_sample": al_id_sample,
        "inv_id_sample": inv_id_sample,
        "material_total": float(o["MaterialCost"].sum()),
        "sqft_billed_total": float(o["SqFtBilled"].sum()),
        "sqft_stone_total": sqft_stone_total,
        "sqft_all_total": sqft_all_total,
        "sqft_nonstone_qty": sqft_all_total - sqft_stone_total,
        "sqft_unmatched_qty": sqft_unmatched_qty,
        "has_opcost": has_opcost,
        "op_diag": op_diag,
        "op_total": float(o["OperationalCost"].sum()),
        "has_manager_map": bool(cfg.manager_map),
        "income_total": float(o["Income"].sum()),
        "profit_total": float(o["Profit"].sum()),
    }
    return Report(o, rep_detail, by_material, by_color, by_ct, by_service_group, periods, meta)


def build_summary(rep_detail: pd.DataFrame, group_col: str) -> pd.DataFrame:
    """Hierarchical summary: group_col subtotal rows + rep rows + grand total."""
    cols = VALUE_COLS
    rows = []
    for g in sorted(rep_detail[group_col].unique()):
        sub = rep_detail[rep_detail[group_col] == g]
        gr = {c: sub[c].sum() for c in cols}
        gr.update({"Level": group_col, "Name": g})
        gr["Margin"] = gr["Profit"] / gr["Income"] if gr["Income"] else np.nan
        rows.append(gr)
        for _, r in sub.sort_values("Income", ascending=False).iterrows():
            d = {c: r[c] for c in cols}
            d.update({"Level": "Rep", "Name": r["Rep"], "Margin": r["Margin"]})
            rows.append(d)
    total = {c: rep_detail[c].sum() for c in cols}
    total.update({"Level": "Total", "Name": "Grand Total"})
    total["Margin"] = total["Profit"] / total["Income"] if total["Income"] else np.nan
    rows.append(total)
    out = pd.DataFrame(rows)
    return out[["Level", "Name"] + cols + ["Margin"]]
