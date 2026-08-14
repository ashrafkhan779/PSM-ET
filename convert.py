#!/usr/bin/env python3
"""
convert.py  —  Clean the raw sales workbook and emit data.json for the dashboard.

Usage:
    python convert.py                              # uses the default filename below
    python convert.py "My Sales File.xlsx"         # custom input
    python convert.py "in.xlsx" "Input Data" out.json

What it does:
    * Reads the 'Input Data' sheet.
    * Trims + merges duplicate customer names that differ only by case/spacing.
    * Repairs dates stored as raw Excel serial numbers.
    * Recomputes Profit (= Invoice Amount - Cost) and Gross % (zero-safe).
    * Dictionary-encodes Customer / Mode / Salesperson to keep the JSON compact.
    * Writes data.json next to this script.

Dependencies:  pandas, numpy, openpyxl   ->   pip install pandas numpy openpyxl
"""

import sys, json, datetime as dt
import pandas as pd
import numpy as np

INPUT   = sys.argv[1] if len(sys.argv) > 1 else "Sales_Report_2026_-_WORKING_FILE.xlsx"
SHEET   = sys.argv[2] if len(sys.argv) > 2 else "Input Data"
OUTPUT  = sys.argv[3] if len(sys.argv) > 3 else "data.json"

MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]


def parse_date(v):
    """Return a datetime for either a normal date string or an Excel serial number."""
    s = str(v).strip()
    d = pd.to_datetime(s, errors="coerce")
    if pd.isna(d):
        try:
            n = float(s)
            d = dt.datetime(1899, 12, 30) + dt.timedelta(days=n)  # Excel epoch
        except Exception:
            return None
    return d


def main():
    print(f"Reading {INPUT} :: sheet '{SHEET}' ...")
    df = pd.read_excel(INPUT, sheet_name=SHEET, header=0)
    df.columns = [c.strip() for c in df.columns]

    required = ["Invoice#", "Invoice Date", "Customer Name", "Quarter", "Month",
                "Years", "Mode", "Invoice Amount", "Cost", "Sales Person"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise SystemExit(f"Missing expected columns: {missing}\nFound: {list(df.columns)}")

    # 1) clean + merge customer names (case/space variants -> most common spelling)
    df["Customer Name"] = df["Customer Name"].astype(str).str.strip()
    key = df["Customer Name"].str.upper()
    canon = {k: g.value_counts().index[0] for k, g in df.groupby(key)["Customer Name"]}
    df["Customer Name"] = key.map(canon)

    # 2) repair dates -> ISO
    dates = df["Invoice Date"].apply(parse_date)
    bad = dates.isna().sum()
    print(f"  dates repaired / unparseable remaining: {bad}")
    df["ISODate"] = dates.dt.strftime("%Y-%m-%d")

    # 3) recompute profit & gross %
    df["Profit"]   = (df["Invoice Amount"] - df["Cost"]).round(2)
    df["GrossPct"] = np.where(df["Invoice Amount"] != 0,
                              df["Profit"] / df["Invoice Amount"], 0.0).round(6)

    # 4) dictionary-encode categoricals
    def enc(col):
        vals = sorted(df[col].dropna().unique().tolist())
        idx = {v: i for i, v in enumerate(vals)}
        return vals, df[col].map(idx).astype(int).tolist()

    cust_vals, cust_idx = enc("Customer Name")
    mode_vals, mode_idx = enc("Mode")
    sp_vals,   sp_idx   = enc("Sales Person")
    mmap = {m: i for i, m in enumerate(MONTH_ORDER)}

    rows = []
    for i, r in df.iterrows():
        rows.append([
            r["Invoice#"],
            r["ISODate"],
            cust_idx[i],
            mode_idx[i],
            sp_idx[i],
            int(r["Years"]),
            mmap.get(str(r["Month"]).strip(), None),
            str(r["Quarter"]).strip(),
            round(float(r["Invoice Amount"]), 2),
            round(float(r["Cost"]), 2),
            round(float(r["Profit"]), 2),
        ])

    payload = {
        "schema": ["inv", "date", "c", "m", "s", "yr", "mo", "q", "rev", "cost", "profit"],
        "customers":  cust_vals,
        "modes":      mode_vals,
        "salespeople": sp_vals,
        "months":     MONTH_ORDER,
        "rows":       rows,
        "meta": {
            "source": INPUT,
            "generated": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "record_count": len(rows),
            "date_min": df["ISODate"].min(),
            "date_max": df["ISODate"].max(),
            "total_revenue": round(float(df["Invoice Amount"].sum()), 2),
            "total_cost":    round(float(df["Cost"].sum()), 2),
            "total_profit":  round(float(df["Profit"].sum()), 2),
        },
    }

    with open(OUTPUT, "w", encoding="utf-8") as f:
        json.dump(payload, f, separators=(",", ":"))

    m = payload["meta"]
    print(f"Wrote {OUTPUT}: {m['record_count']} rows, "
          f"{len(cust_vals)} customers, {len(mode_vals)} modes, {len(sp_vals)} reps")
    print(f"  revenue {m['total_revenue']:,.0f} | profit {m['total_profit']:,.0f} "
          f"| {m['date_min']} -> {m['date_max']}")


if __name__ == "__main__":
    main()
