# PSM Procurement & Shipment Dashboard

An interactive, self-contained dashboard for tracking the PSM–ET purchase book:
overview KPIs, PO tracking with a lifecycle timeline, product & material analytics,
and an exception register for overdue lines. Opens in any browser — no install.

Files:
- `index.html` — the dashboard (this is all you need to view it)
- `data.json` — the data it shows (regenerate this when your working file changes)
- `convert.py` — optional script to rebuild `data.json` from Excel

---

## 1. Publish a public link with GitHub Pages

1. Sign in at github.com → **New repository** → name it e.g. `psm-dashboard` → set **Public** → **Create**.
2. On the repo page → **Add file → Upload files** → drag in `index.html` and `data.json` (and `README.md`) → **Commit changes**.
3. **Settings → Pages** → under *Build and deployment*, Source = **Deploy from a branch**, Branch = **main** / **/(root)** → **Save**.
4. Wait ~1 minute. Your link appears at the top of the Pages settings:
   **`https://<your-username>.github.io/psm-dashboard/`**
5. Share that link with management. It always shows whatever `data.json` is currently in the repo.

> The dashboard loads `data.json` from the same folder. On the live link it uses the committed
> `data.json`; opened locally by double-click it falls back to the data baked into `index.html`.

---

## 2. Update the data (do this daily)

You maintain your Excel working file as usual. To push an update, pick **one** route.

### Option A — no code (recommended)
1. Open the dashboard (the live link, or `index.html` locally).
2. Click **Update data** (top right) → choose your `PSM_WORKING_FILE.xlsx`.
   The dashboard cleans and re-renders instantly — you can sanity-check the numbers.
3. Click **Download data.json**.
4. In your GitHub repo → **Add file → Upload files** → drop the new `data.json` → **Commit**.
   The live link refreshes within a minute.

*If you only need it for yourself, stop after step 2 — no upload needed.*

### Option B — scripted (for automation)
```bash
pip install pandas openpyxl        # first time only
python convert.py PSM_WORKING_FILE.xlsx
git add data.json && git commit -m "data update" && git push
```

Both routes apply the same cleaning: they parse delays out of the *Time Frame* field,
split the mixed *STATUS* column into a workflow stage plus a revised ship date, correct
the delivered-line quantities, and trim stray spaces in *Material #*.

---

## 3. Before you share publicly — a note on the data

A **public** GitHub Pages link is readable by anyone who has the URL, and `data.json`
contains supplier names, prices and PO values. If that is acceptable for this audience, you're done.
If it is **not**, don't use public Pages for it — host behind access control instead
(e.g. Netlify or Cloudflare Pages with password / SSO, or an internal server). The dashboard
file works the same way on any of them; only `data.json` needs to sit next to `index.html`.
