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
SO days    = distinct Sales Order Dates in the so tab
Projection = projection › Total KGs ÷ days in the projection month × SO days
SO         = so › Stock Qty in KGS            where Customer Group ≠ Category A
           + dispatches › Stock Qty In Kg     where Sales Channel  = Category A
                                                and Invoice Status ≠ Return
Gap        = Projection − SO
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
- **The projection is pro-rated to the SO days.** A per-day projection
  (Total KGs ÷ calendar days in the month) is multiplied by the number of
  distinct Sales Order Dates in the so tab, so a part-month SO dump is compared
  with the matching share of the month. The days are counted over the **whole**
  so tab — every origin, before the Category A and warehouse rules — so every
  CFA and customer, Category A included, is scaled by the same factor. With the
  month filter on, only dates in the projection month count. The banner and the
  Projection tile show the factor (e.g. 29 / 30 days); untick *Pro-rate
  projection to SO days* to compare the full month. The export carries both.
- **The projection's grand-total row is dropped** — summed in, it would double
  every figure.
- **Month.** By default so and dispatches rows outside the projection's month
  are left out, and a tab with no rows in that month is called out in red
  rather than quietly reading as zero. Untick *Only the projection month* to
  compare the files as loaded.

Over / Under is decided at the lowest level — each customer, each item — so a
group's subtotal in the Over view adds only its over-projected rows.

## Files

`calc.js` holds every rule and no DOM; `app.js` is loading, views and export.
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
