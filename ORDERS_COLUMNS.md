# 📑 Orders tab — column reference

Every column of the **Orders** tab (and of the *Orders — Net Margin* sheet in both
Excel downloads), what it means, why it matters and how it is calculated.

One row = **one order**, in the month that order was **invoiced**.

---

## Identity — who and what

| Column | What it is | Where it comes from |
|---|---|---|
| **Order ID** | The internal ERP order number. This is the **join key** — every other file is matched to the order on this. | Invoice List `OrderID`, normalised (a trailing `.0` stripped so `75961.0` and `75961` match) |
| **Order #** | The human-facing order number. | Invoice List `OrderNumber` |
| **Ref #** | Customer's own reference / PO. | Invoice List `RefNumber` |
| **Cust #** | Customer account number. | Sales Person Summary, falling back to the invoice |
| **Customer Name** | Who the job is for. | Customer master `CustomerFullName`, falling back to the invoice |
| **Branch** | The shop. **Drives the whole Shop grouping** and what a shop-scoped login can see. | The rep → shop mapping (⚙️ Shops & Managers tab), not the ERP |
| **Sales Person** | The rep credited with the order. | Sales Person Summary `SalesPersonName`; if blank, the customer's assigned rep; else `(Unassigned)` |
| **Customer Type** | Builder / retail / trade, etc. Feeds the 👥 Customer Type tab. | Customer master `Type`, else `Unknown` |

---

## Volume — the denominator for everything per-sqft

| Column | What it is | How it's calculated |
|---|---|---|
| **Sq Ft (Stone)** | Square footage **billed**. The most important non-money column: Fixed Cost, the fab rate, the install spread and Contribution/SqFt are all driven off it. | Sales By SKU `Quantity`, **stone lines only** — sinks, edges, services and accessories are excluded because their Quantity is a unit count, not sqft |
| **Sq Ft Allocated** | Square footage of slab actually **pulled from inventory**. | Allocation `AllocationMeasure`, stone lines |

> **Why both:** Allocated ≫ Billed means you cut more slab than you sold — waste,
> remakes, or a bad allocation. It is your yield check.

---

## Revenue

| Column | What it is | How it's calculated |
|---|---|---|
| **Total Invoice** | Gross invoice amount. | Invoice List `TotalAmount`, cancelled invoices excluded |
| **Sales Tax** | Tax collected — not revenue, it is passed through. | Invoice List `SalesTax` |
| **Net Sales** | **The real top line.** Every margin on this row is a percentage of this. | `Total Invoice − Credit Memos − Sales Tax` |

---

## Direct costs — the five operational buckets + material

| Column | Importance | How it's calculated |
|---|---|---|
| **Template Cost** | What you paid the templater. | Template Schedule, `Total Final Paid` (the post-*descuentos* figure — what actually left the bank) |
| **Install Cost** | Usually the largest operational line. | Install Schedule `Total Final Paid`; orders the schedule misses get their share of the sidebar **Install total cost $**, split by Sq Ft (Stone) |
| **Fab Cost** | In-house fabrication labour. | Production LOG `Total Amount` (summed — an order can have several lines); orders with none fall back to `Sq Ft (Stone) × fab fixed cost rate` |
| **Plumbing Cost** | No export carries it. | Typed per order (✏️ Edit mode) |
| **Additional Costs** | The catch-all — sealing, freight, rework, anything else. | Vendor bill exports, matched on the order number inside each bill's `Memo`; also typeable by hand |
| **Material Cost** | The slab itself — normally your single biggest cost. | Allocation `ExtendedCost` per line; where that is 0 (sinks/accessories) `AllocationQty × UnitCost` |

> **Timing rule:** every one of these attaches to the order **in the month the
> order is invoiced**, not the month the cost was incurred. That is what keeps
> cost and revenue on the same row.

---

## Profit

| Column | What it is | How it's calculated |
|---|---|---|
| **Total Direct Cost** | Everything the job itself consumed. | `Material + Template + Install + Fab + Plumbing + Additional` |
| **Gross Profit ($)** | **The headline number.** What the job earned before any overhead. | `Net Sales − Total Direct Cost` |
| **GP Margin %** | Gross Profit as a share of the sale — lets you compare a $4k job to a $40k one. | `Gross Profit ÷ Net Sales` |
| **Fixed Cost** | Overhead allocated to this job. Deducted *after* Gross Profit — it is not a job cost, it is this job's share of rent / admin / equipment. | `Sq Ft (Stone) × Fixed cost $/SqFt` (sidebar; the example report used $28.36) |
| **Net Profit ($)** | What the job left after carrying its share of the business. | `Gross Profit − Fixed Cost` |
| **Net Margin %** | The bottom line as a percentage. | `Net Profit ÷ Net Sales` |
| **Contribution / Sq Ft** | **The best cross-job comparison in the table.** Strips out job size entirely — tells you which work is actually worth taking. | `Gross Profit ÷ Sq Ft (Stone)` |

---

## The chain, end to end

```
        Total Invoice
      − Credit Memos
      − Sales Tax
      ─────────────────
      = NET SALES
      − Material Cost
      − Template + Install + Fab + Plumbing + Additional   (= operational)
      ─────────────────
      = GROSS PROFIT           GP Margin %  = GP ÷ Net Sales
      − Fixed Cost             (Sq Ft (Stone) × fixed rate)
      ─────────────────
      = NET PROFIT             Net Margin % = Net Profit ÷ Net Sales

      Contribution / Sq Ft = Gross Profit ÷ Sq Ft (Stone)
```

---

## Two things worth knowing

**Editable vs derived.** In ✏️ Edit mode only these 10 columns can be typed:

> Total Invoice · Sales Tax · Sq Ft (Stone) · Sq Ft Allocated · Material Cost ·
> Template Cost · Install Cost · Fab Cost · Plumbing Cost · Additional Costs

Everything else recalculates from them — so if you fix **Sq Ft (Stone)**, the Fixed
Cost, the fab-rate fill, the install split and Contribution/SqFt all move with it.
Typing the original value back clears the override. Edits are saved to
`order_overrides.json`, apply for every user, and never touch the source files.

**Net Sales, not Income.** The sidebar's *Income basis* toggle (Billed vs Amount
Paid) drives the **Summary** tab. The Orders tab always computes margin off
**Net Sales** (billed), so a job that is invoiced but not yet collected still shows
its true margin instead of reading as a 100% loss.

---

*See [README.md](README.md) for the input files, the cost-file formats and the
sidebar rate settings.*
