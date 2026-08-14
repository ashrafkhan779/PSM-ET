# Finance Control — Sales & Financial BI Dashboard

A single-file web dashboard for a UAE trading business. It reads a cleaned sales dataset and adds
editable finance modules (expenses, banks, assets, cash flow) — all in the browser, no backend.

---

## Files

| File | What it is |
|---|---|
| `index.html` | The whole application. |
| `data.json` | The dataset the dashboard loads on start (sales + any saved manual data). |
| `Dashboard_Master.xlsx` | **The fill-in template for every tab** — fill it, upload it, done. |
| `convert.py` | Regenerates a sales-only `data.json` from an Excel export (optional). |
| `README.md` | This file. |

Keep `index.html` and `data.json` in the **same folder**.

---

## Running it

The app loads `data.json` with `fetch`, which browsers block on `file://`, so **serve the folder**:

```bash
python -m http.server        # then open http://localhost:8000
```

Or push the folder to **GitHub Pages** and use the live URL. Charts need an internet connection
(Chart.js loads from a CDN); all tables and figures work offline.

---

## The three buttons (top bar)

- **⭱ Upload Excel** — upload a filled `Dashboard_Master.xlsx`. Every tab updates from it:
  sales, banks, operating expenses, corporate tax, assets + depreciation/interest/returns,
  liabilities + payments/returns, cash flow, and upcoming expenses.
- **⬇ data.json** — downloads **all dashboard data** (sales **and** every manual tab) as a single
  `data.json`. Upload that file to your GitHub repo (replace the existing `data.json`) and the
  **live link updates for everyone** — nothing else to send.
- **⬇ Sample Excel** — downloads the master template (all tabs) reflecting your current data, so
  you can fill/adjust and re-upload.

### The workflow

1. Open **`Dashboard_Master.xlsx`**, fill the sheets you need (see the *How to use* sheet inside).
2. **Upload Excel** → the dashboard updates every tab.
3. **Download data.json** → commit/replace it in your GitHub repo → the live link shows everything.

Your data is also saved in the browser, so it persists across reloads on that device without any
of the above.

---

## The master template — sheets

`Dashboard_Master.xlsx` (and the **Sample Excel** download) contain one sheet per tab:

| Sheet | Fills | Key columns |
|---|---|---|
| **Input Data** | Sales | Invoice#, Invoice Date, Customer Name, Quarter, Month, Years, Mode, Invoice Amount, Cost, Sales Person |
| **Banks** | Bank Management | Account, Type (`bank`/`petty`/`other`), Balance, Note |
| **Operating Expenses** | Expenses matrix | Year, Expense Head, Jan … Dec |
| **Corporate Tax** | P&L tax | Year, Corporate Tax |
| **Assets** | Fixed Assets | Asset, Category, Acquired, Cost, Note |
| **Asset Movements** | Depreciation/interest/returns | Asset, Type (`depreciation`/`interest`/`return`), Date, Amount, Note |
| **Liabilities** | Liabilities | Liability, Category, Date, Amount, Note |
| **Liability Movements** | Payments/returns | Liability, Type (`payment`/`return`), Date, Amount, Note |
| **Cash Flow** | Receivables/Payables | Category (Local/Oversea Receivable/Payable), Amount, Note |
| **Upcoming Expenses** | Cash Flow → upcoming | Month (YYYY-MM), Date, Description, Amount, Note |

Keep the **sheet names and header rows** exactly as provided. Profit and Gross % on sales are
computed automatically by the app.

---

## Tabs

**Analytics:** Dashboard · Customer Analysis · Salesperson Performance · Quarterly Comparison ·
New Customer Onboard · Dormant Customers.
**Finance:** Profit & Loss · Fixed Assets & Liabilities · Operational Expenses · Bank Management ·
Cash Flow Control.

---

## Notes

- No figures are invented — every number comes from the data you provide.
- Nothing is uploaded anywhere; the app runs entirely in your browser.
- A `data.json` produced by this dashboard bundles sales **and** manual data together, so replacing
  it in your repo is all that's needed to update the live link.
