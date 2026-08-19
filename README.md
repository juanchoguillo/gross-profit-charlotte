# MC Granite — Dynamic Gross Profit Dashboard

A local web app that turns the ERP export files into an interactive gross-profit
report, grouped by **Shop** or by **Manager**, down to each sales rep.

## How to run

**Easiest (macOS):** double-click **`run.command`** (first launch sets up, ~1 min).

**Terminal:**
```bash
cd "MC_Granite_Dev/gross-profit"
python3 -m venv .venv                            # first time only
./.venv/bin/pip install -r requirements.txt      # first time only
./.venv/bin/streamlit run app.py
```
Opens at http://localhost:8501. Files are auto-detected by their columns, so
names can vary.

## Input files

| File | Required | Provides |
|---|---|---|
| Sales By SKU | ✅ | SqFt billed & revenue per order |
| Invoice List | ✅ | Income (Amount Paid, non-cancelled) + invoice month |
| Sales Person Summary | ✅ | Order → sales rep |
| Inventory Allocation Details | ✅ | Material cost & SqFt allocated |
| Customers (ReportAdHoc) | ✅ | Customer **Type** |
| Products | ✅ | SKU → material (`Catalogs`) |
| **2026 Template Schedule** (`.xlsx`) | optional | **Template cost** per order (one tab per templater) |
| **2026 Install Schedule** (`.xlsb`) | optional | **Install cost** per order (one tab per installer) |
| **2026 Production Install LOG** (`.xlsx`) | optional | **Fabrication cost** per order (`2026` tab) |
| **Vendor bill export** (`.xlsx`) | optional | **Additional cost** per order (e.g. *Adicional Cost - Hydroshield*) |
| **Shop/Manager mapping** | optional | rep → shop & manager |

## Calculation

| Column | How |
|---|---|
| **Income** | Invoice `AmountPaid`, non-cancelled (toggle: Billed = Total − Tax) |
| **SqFt Billed** | Sales-By-SKU `Quantity` for the month's invoiced orders, **stone lines only** (sinks/edges/services/accessories excluded) |
| **Material Cost** | stone → `ExtendedCost`; non-stone → `AllocationQty × UnitCost` |
| **SqFt Allocated** | stone `AllocationMeasure` |
| **Install Cost** | Install Schedule amount per order; orders with none share the **Install total cost $** (sidebar, default 0) by their SqFt |
| **Fabrication Cost** | Production LOG amount per order; orders with none fall back to `SqFt Billed × fab fixed cost rate` (sidebar, default 0) |
| **Plumbing Cost** | typed per order in the **📑 Orders** tab (no export carries it) |
| **Additional Costs** | vendor bill exports, matched by the order number in each bill's Memo; also typeable per order in the **📑 Orders** tab — the catch-all for anything the other cost columns don't cover |
| **Operational Cost** | Template + Install + Fabrication + Plumbing + Additional, per order |
| **Overhead** | `SqFt Billed × fixed cost rate` (default 0, editable) |
| **Total Cost** | Material + Operational + Overhead |
| **Profit / Margin** | Income − Total Cost ; Profit ÷ Income |

Stone vs non-stone and material category come from the Products **Catalogs**
field (Granite/Quartz/Quartzite/Marble/Porcelain/Soapstone = stone).
Customer type comes from the customer master **Type** field.

A SKU that *is* in Products but has a **blank Catalogs** field lands in the
**Other** category — in practice labor, edge, cutout, removal and service lines,
not stone. The 🪨 Material tab opens that bucket up ("What's inside *Other*"):
one row per uncategorised SKU, plus a line-by-line list of the orders each one
was billed on, so they can be found and given a `Catalogs` value in the ERP.
Both land in the full Excel export as the **Other Breakdown** and
**Other — Detail** sheets.

## Operational cost — Template, Install, Fabrication, Plumbing & Additional

Template, Install, Fabrication and Additional come from four workbooks, joined to
each order:

| File | Tabs read | Order column | Cost column |
|---|---|---|---|
| **Template Schedule** (`.xlsx` / `.xlsm`) | `Templater_<name>` (e.g. `Templater_Frankie`) — older files: one tab per templater (Ricardo, Caio, Michael) | `Order Id` (older: `Order No.`) | `Total Final Paid` (older: `Per Order`) |
| **Install Schedule** (`.xlsm` / `.xlsb`) | `Installer_<name>` (e.g. `Installer_EliteStone`) — older files: one tab per installer (Oscar, Flavio, …) | `Order Id` (older: `Order #`) | `Total Final Paid` (older: `Total`) |
| **Production Install LOG** (`.xlsx`) | the `2026` tab | `ORDER` | `Total Amount` (summed — multiple lines per order) |
| **Vendor bill export** (`.xlsx`) | every tab | `Memo` (order number read out of the free text) | `Debit` |

**Tab naming decides the file.** A workbook with `Installer_…` tabs is the install
schedule, one with `Templater_…` tabs is the template schedule — so both can be
exported straight from the ERP with the raw `Invoice List` / `Sales Person
Summary` tabs still attached; those are ignored, not read.

**Vendor bills carry no order column.** The QuickBooks-style bill export
(`Type · Date · Num · Name · Memo · Class · Split · Debit`) is recognised by those
columns, and the order number is read out of the free-text `Memo` —
`sealing service  Bizops 30006` → order `30006`, `Bizops JEN 29609` → `29609`. The
last 4–7-digit run in the memo wins. A memo that names the customer instead of the
job (`Bizops Michele Blair`) has no order number, so **that line is left out** of
the report rather than misfiled — the 🛠️ Operational tab's "cost in file" vs.
"applied" columns show the gap. Drop in several bill exports (one vendor each) and
they all add into **Additional Costs**.

**`Total Final Paid`, not `Total`.** It is the amount that actually left the bank
(it applies the `descuentos` adjustments). `Total` is the pre-adjustment figure.

- The order reference (e.g. `SOU-76926`) is matched on its **numeric part**
  (`76926`), which is the OrderID used across the other exports.
- Each workbook's per-crew tabs are read and summed; **summary/aggregate tabs
  (`Total Payments`, `Sheet1`, `INSTALLS-2026`, …) are skipped** so nothing is
  double-counted. Repeat visits to the same order are summed.
- Inside a crew tab, rows that group the list rather than bill an order are
  ignored: the **`TOTAL` footer** (it repeats the tab's own sum), **month
  banners** (`2026-06-01`), **`INVOICE #…` banners**, and the **blank-`Order Id`
  invoice subtotal / `MILEAGE` rows**. These carry roughly as much money as the
  order rows themselves, so counting them would nearly double the cost.
- A row whose cost cannot be tied to an order (blank `Order Id`) is left out of
  the report — the 🛠️ Operational tab's "cost in file" reflects only the
  assignable rows.
- Operational cost attaches to an order **when it's invoiced** in the selected
  period (the same way material cost is counted). The **🛠️ Operational** tab
  shows a per-type breakdown plus match diagnostics (cost in file vs. applied).

### Fab fixed cost $/SqFt — when there's no Production LOG

Set **Fab fixed cost $/SqFt (stone)** in the sidebar and every order with **no**
Production LOG amount gets `Fab Cost = Sq Ft (Stone) × rate`. Orders the LOG does
cover keep the LOG amount, so nothing is double-counted — with no LOG loaded at
all, every order is rate-based. It is a **direct** cost, so it reduces Gross
Profit (unlike the *Fixed cost $/SqFt*, which is deducted after Gross Profit).
The **🛠️ Operational** tab reports how many orders the rate filled and for how
much; the file-diagnostics table below it stays LOG-only.

### Install total cost $ — when the Install Schedule doesn't cover an order

Some install work is billed as one lump for the month rather than order by order.
Type that lump into **Install total cost $ (spread by SqFt)** in the sidebar and
the app:

1. finds every order in the selected period with **no** Install Schedule cost,
2. adds up their **Sq Ft (Stone)**,
3. divides the total by that sqft → a `$/SqFt` install rate,
4. sets `Install Cost = Sq Ft (Stone) × that rate` on each of those orders.

So the whole amount you type lands on the report, split in proportion to how big
each uncovered job is. Orders the Install Schedule *does* cover keep their own
amount and are left out of the split entirely — they neither draw from the pot
nor dilute the rate. A hand-typed Install Cost counts as covered the same way, so
the remaining orders re-share what is left. The sidebar shows the derived rate
(`$50,000.00 ÷ 4,605.83 sqft = $10.86/SqFt on 98 orders`) and the 🛠️ Operational
tab repeats it with the split total. Like the fab rate, it is a **direct** cost.

Note the split is per **reporting period** — the total you type is the total for
the month you are looking at.

### Plumbing cost & Additional costs

No export carries **Plumbing Cost**, so it is typed per order: open the
**📑 Orders** tab, turn on **✏️ Edit mode**, type the amount and **💾 Save edits**
(stored in `order_overrides.json`, applied for every user, source files never
touched). **Additional Costs** comes from the vendor bill exports and can be typed
the same way — a typed amount replaces the file's for that order. Both round-trip
through the Excel exports too: edit the column and re-upload the workbook.
*Additional Costs* is the catch-all for anything the other columns don't cover.

Both are **direct** costs: they feed Total Direct Cost and reduce Gross Profit.

## Shop / Manager mapping — edit in-app

Open the **⚙️ Shops & Managers** tab. It lists every sales rep found in your data
(shops pre-filled from the known list). Type each rep's **Shop** and **Manager**,
add rows for new reps, and click **💾 Save** — it persists to `config.json` and
loads automatically next time. Edits update the report immediately.

You can still *bulk-import* a mapping by dropping in a spreadsheet with
`Sales Rep / Shop / Manager` columns (rep column may also be `SalesPersonName`
or `Project Manager`); it seeds the editor.

## ⚠️ Match the periods
Export all files for the **same date range**. If the invoice month and the
allocation / schedule dates don't overlap, costs read low — the app shows a banner.

## Files
- `app.py` — dashboard · `gp_core.py` — engine · `report_xlsx.py` — Excel
  export · `requirements.txt` · `run.command`
- [`ORDERS_COLUMNS.md`](ORDERS_COLUMNS.md) — what every column of the **📑 Orders**
  tab means and how it is calculated
