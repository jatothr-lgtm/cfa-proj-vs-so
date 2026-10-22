# Projection vs SO

Did we project what customers actually ordered? Projection against sales
orders per CFA, by customer and by item, split into over- and under-projected.

A static page: no build, no dependencies, no backend. The Excel is read in the
browser and nothing is uploaded or stored. It also ships inside the
[CFA Suite](https://github.com/jatothr-lgtm/cfa-suite) as `proj-vs-so/`.

## Input

One workbook with the `projection`, `so` and `dispatches` tabs plus
`Normalisation` — or the three as separate files. Sheets are recognised by
their columns, not their names, and the header row may sit anywhere in the
first 15 rows.

| Tab | Warehouse | Customer Group | Kg | Item Group |
|---|---|---|---|---|
| projection | `Wareouse` | `Customer Group` | `Total KGs` | `Item Group` |
| so | `Origin` | `Customer Group` | `Stock Qty in KGS` | `NEW MIS ITEM GROUP` |
| dispatches | `From Warehouse` | `Sales Channel` | `Stock Qty In Kg` | `New Mis Item Group` |

Customer, Item Parent (`Parent Item` in so), Item Name and Item Code complete
each tab.

## The rules

```
Month proj = projection › Total KGs
SO days    = unique Sales Order Dates in the so tab
MTD proj   = Month proj ÷ days in the projection month × SO days   (per-day projection × SO days)
SO         = so › Stock Qty in KGS            where Customer Group ≠ Category A
           + dispatches › Stock Qty In Kg     where Sales Channel  = Category A
                                                and Invoice Status ≠ Return
Gap        = MTD proj − SO
Over       = Gap > 0      Under = Gap < 0      (per customer, or per item)
```

- **CFAs.** Only the warehouses on the `Normalisation` tab are compared; its
  `CFA (XXX)` rows say which spelling in each tab is the same CFA. Without that
  tab the built-in GGN / BLR / MUM mapping is used. Everything else is left out
  and listed on the Data tab with its kg.
- **Category A is one customer.** The projection books it to a single customer
  (GT Retail) while dispatches carry each distributor, so every Category A
  customer on both sides is added into one row.
- **Items are matched on Item Code.** The tabs label the same SKU's parent
  differently (projection *Prasadam Roasted Makhana* vs so *Roasted &
  Flavoured Makhana - Achaari*); matching on labels would split one SKU into a
  false over row and a false under row. Labels come from the projection first,
  then so, then dispatches.
- **SO is compared with the MTD projection.** Both columns are shown: Month
  proj (Total KGs) and MTD proj — the per-day projection (Total KGs ÷ calendar
  days in the month) times the number of unique Sales Order Dates in the so
  tab. Gap, MTD ÷ SO and Over / Under all use MTD proj, so a part-month SO dump
  is compared with the matching share of the month. The days are counted over the **whole**
  so tab — every origin, before the Category A and warehouse rules — so every
  CFA and customer, Category A included, is scaled by the same factor. With the
  month filter on, only dates in the projection month count. The banner and the
  MTD projection tile show the day count (e.g. 29 of 30 days). With no usable
  SO dates, MTD proj falls back to the month and a warning says so. The export
  carries both columns.
- **The projection's grand-total row is dropped** — summed in, it would double
  every figure.
- **Month.** By default so and dispatches rows outside the projection's month
  are left out, and a tab with no rows in that month is called out in red
  rather than quietly reading as zero. Untick *Only the projection month* to
  compare the files as loaded.

## On screen

- **Period, in short**, top right on every tab: *September 2026 · 29 / 30 SO
  days*, with a thin progress line. The MTD projection tile repeats it as a
  meter.
- **Tiles:** Month projection · MTD projection · SO · Net gap · Over · Under.
- **MTD projection vs SO by group** — paired bars per Customer Group (or Item
  Group), the gap in the right column.
- **Largest gaps** — diverging bars from zero, over-projected right (blue),
  under-projected left (red), the top 12 customers or items.
- Charts follow every filter (CFA, mode, search, month) and carry hover / focus
  tooltips; the table below holds every value.

Chart colours were checked with a colour-blind-safety validator in light and
dark. Motion follows the reference animation repos (HyperFrames, Claude
Animation Skill, ClaudeAnimationBase): long-tail ease-out
`cubic-bezier(0.16,1,0.3,1)` on arrivals, ~50 ms stagger, bars grow from the
baseline, figures count up in turn, nothing bounces or keeps moving, and it all
switches off under *reduce motion*.

Over / Under is decided at the lowest level — each customer, each item — so a
group's subtotal in the Over view adds only its over-projected rows.

**Customer tab drill-down.** Each customer opens into Item Group (NEW MIS ITEM
GROUP) › Item Parent › Item Name, with the same columns and its own Over / Under
status on every item. The customer row is still what the Over / Under tabs,
tiles and charts classify, so opening a customer never changes their counts. A
search that matches items (not the customer) narrows each customer to those
items and its totals to theirs, so the rows on screen always add up. The export
adds a **Customer x Item** sheet with every customer-item pair.

## Files

`calc.js` holds every rule and no DOM; `charts.js` draws the SVG charts; `app.js` is loading, views and export.
`styles.css` carries the CFA Suite's design tokens, so the two read as one system.

## Running it

```bash
python -m http.server 8790
```

then open <http://localhost:8790>. `app.js` is an ES module, so opening
`index.html` straight from disk will not work — serve the folder.

## Deploying

Vercel, from the GitHub repo:

1. Vercel → **Add New → Project** → import this repository.
2. Framework preset **Other**. Leave Build Command and Install Command empty;
   Output Directory is the repo root (`.`).
3. Deploy. No environment variables are needed — there is no backend.

`vercel.json` turns on clean URLs and sends `no-cache` for the HTML, CSS and JS
so a redeploy is picked up on the next page load.

**Commit author matters on Vercel Hobby.** A push-triggered deploy is blocked
unless the commit author is the account owner, so commit as
`jatoth.r@farmley.com`.

## Data hygiene

`.gitignore` blocks `*.xlsx` / `*.xls` / `*.csv` — real Farmley extracts never
go in this repo.
