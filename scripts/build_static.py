"""
build_static.py
----------------
Converts the raw "Stock Position Report" Excel export into public/data.json,
using the exact same mechanism as the store's own Sales Dashboard SOH sheet:

  - iPhone      = Category "DEVICES"     + SubCategory "PHONE"
  - iPad        = Category "DEVICES"     + SubCategory "TABLETS"
  - Mac         = Category "DEVICES"     + SubCategory "LAPTOPS"
  - Apple Watch = Category "DEVICES"     + SubCategory "WATCH"
  - AirPods     = Category "ACCESSORIES" + SubCategory "EARPHONES" or "HEADPHONES"

Stock quantity = MAX(number of matching rows for that Article code, sum of the
raw Qty column for that Article code) — some lines are tracked one-row-per
physical/serialized unit (row count = real stock), others as a single row
with a Qty total (e.g. AirPods); taking the larger of the two avoids
under-reporting either pattern.
"""
import json
import pathlib
import re
from datetime import datetime

import pandas as pd

SRC_PATH = pathlib.Path("downloaded/stockpositionreportdetail.xls")
OUT_PATH = pathlib.Path("data.json")

OFFICIAL_MAP = [
    ("iPhone", "DEVICES", "PHONE"),
    ("iPad", "DEVICES", "TABLETS"),
    ("Mac", "DEVICES", "LAPTOPS"),
    ("Apple Watch", "DEVICES", "WATCH"),
    ("AirPods", "ACCESSORIES", "EARPHONES"),
    ("AirPods", "ACCESSORIES", "HEADPHONES"),
]


def extract_report_date(df: pd.DataFrame):
    """Looks for the 'Date :' and 'Time :' fields in the report's title block (row 0),
    same convention the store's export always uses, and combines them — the "Time :"
    cell's own date part is meaningless (Excel time-only serials land on an arbitrary
    epoch date), only its hour/minute/second matter."""
    date_found = None
    time_found = None
    for i in range(min(5, len(df))):
        row = df.iloc[i]
        for j, cell in enumerate(row):
            if isinstance(cell, str) and re.match(r"^date\s*:?$", cell.strip(), re.I):
                nxt = row.iloc[j + 1] if j + 1 < len(row) else None
                if isinstance(nxt, (pd.Timestamp, datetime)):
                    date_found = nxt.to_pydatetime() if hasattr(nxt, "to_pydatetime") else nxt
            if isinstance(cell, str) and re.match(r"^time\s*:?$", cell.strip(), re.I):
                nxt = row.iloc[j + 1] if j + 1 < len(row) else None
                if isinstance(nxt, (pd.Timestamp, datetime)):
                    time_found = nxt.to_pydatetime() if hasattr(nxt, "to_pydatetime") else nxt
            if date_found is None and isinstance(cell, (pd.Timestamp, datetime)):
                date_found = cell.to_pydatetime() if hasattr(cell, "to_pydatetime") else cell
        if date_found is None:
            joined = " ".join(str(c) for c in row if pd.notna(c))
            m = re.search(r"for\s+(\d{2})-(\d{2})-(\d{4})", joined, re.I)
            if m:
                date_found = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    if date_found is None:
        return None
    if time_found is not None:
        return datetime(date_found.year, date_found.month, date_found.day,
                         time_found.hour, time_found.minute, time_found.second)
    return date_found


def build():
    df = pd.read_excel(SRC_PATH, sheet_name=0, header=None)
    df.columns = ["brand", "dept", "category", "subcat", "skuname",
                  "barcode", "price", "qty", "value"] + \
                 [f"extra{i}" for i in range(max(0, df.shape[1] - 9))]

    report_date = extract_report_date(df)

    d = df[df["brand"].astype(str).str.strip().str.upper() == "APPLE"].copy()
    d = d[d["skuname"].notna()]
    d["kode"] = d["skuname"].astype(str).str.split(" / ").str[0].str.strip()
    d["nama"] = d["skuname"].astype(str).str.split(" / ").str[1:].apply(
        lambda x: " / ".join(x).strip() if isinstance(x, list) else ""
    )
    d["category"] = d["category"].astype(str).str.strip().str.upper()
    d["subcat"] = d["subcat"].astype(str).str.strip().str.upper()

    # COUNTIF equivalent: how many rows use each Article code, across the whole sheet.
    all_codes = df["skuname"].astype(str).str.split(" / ").str[0].str.strip()
    row_counts = all_codes.value_counts()

    records = []
    seen = set()
    for label, cat, sub in OFFICIAL_MAP:
        subset = d[(d["category"] == cat) & (d["subcat"] == sub)]
        for kode, nama in subset[["kode", "nama"]].drop_duplicates().itertuples(index=False):
            if not kode or kode in seen:
                continue
            seen.add(kode)
            baris = int(row_counts.get(kode, 0))
            qty_kolom = int(d.loc[d["kode"] == kode, "qty"].sum())
            stok = max(baris, qty_kolom)
            records.append({
                "kategori": label,
                "kode": kode,
                "nama": nama,
                "stok": stok,
                "baris": baris,
                "qtyKolom": qty_kolom,
            })

    records.sort(key=lambda r: (r["kategori"], r["nama"]))

    payload = {
        "generatedAt": datetime.utcnow().isoformat() + "Z",
        "sourceDate": report_date.isoformat() if report_date else None,
        "items": records,
    }

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {len(records)} SKUs to {OUT_PATH} (report date: {report_date})")


if __name__ == "__main__":
    build()
