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
| **Shop/Manager mapping** | optional | rep → shop & manager |

## Calculation

| Column | How |
|---|---|
| **Income** | Invoice `AmountPaid`, non-cancelled (toggle: Billed = Total − Tax) |
| **SqFt Billed** | Sales-By-SKU `Quantity` for the month's invoiced orders, **stone lines only** (sinks/edges/services/accessories excluded) |
| **Material Cost** | stone → `ExtendedCost`; non-stone → `AllocationQty × UnitCost` |
| **SqFt Allocated** | stone `AllocationMeasure` |
| **Operational Cost** | Template + Install + Fabrication cost, joined per order |
| **Overhead** | `SqFt Billed × overhead rate` (default 0, editable) |
| **Total Cost** | Material + Operational + Overhead |
| **Profit / Margin** | Income − Total Cost ; Profit ÷ Income |

Stone vs non-stone and material category come from the Products **Catalogs**
field (Granite/Quartz/Quartzite/Marble/Porcelain/Soapstone = stone).
Customer type comes from the customer master **Type** field.

## Operational cost — Template, Install & Fabrication

Operational Cost is the sum of three schedule workbooks, joined to each order:

| File | Tabs read | Order column | Cost column |
|---|---|---|---|
| **Template Schedule** (`.xlsx`) | one per templater (e.g. Ricardo, Caio, Michael) | `Order No.` | `Per Order` (all-in: Temple $ + mileage + other fees) |
| **Install Schedule** (`.xlsb`) | one per installer (Oscar, Flavio, …) | `Order #` | `Total` |
| **Production Install LOG** (`.xlsx`) | the `2026` tab | `ORDER` | `Total Amount` (summed — multiple lines per order) |

- The order reference (e.g. `SOU-76926`) is matched on its **numeric part**
  (`76926`), which is the OrderID used across the other exports.
- Each workbook's per-crew tabs are read and summed; **summary/aggregate tabs
  (`Total Payments`, `Sheet1`, `INSTALLS-2026`, …) are skipped** so nothing is
  double-counted. Repeat visits to the same order are summed.
- Operational cost attaches to an order **when it's invoiced** in the selected
  period (the same way material cost is counted). The **🛠️ Operational** tab
  shows a per-type breakdown plus match diagnostics (cost in file vs. applied).

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
- `app.py` — dashboard · `gp_core.py` — engine · `requirements.txt` · `run.command`
