import json
from pathlib import Path
import pandas as pd

# Project-root relative path to raw
EXCEL_PATH = Path(__file__).resolve().parents[1] / 'data' / 'raw' / "20250617  HoFF TXA buffer-1.xlsx"
SHEET_NAME = "Table All Cycles"
CHANNEL_LABEL = " Raw Data (625-30/680-30)"

raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None, engine="openpyxl")
metadata = {}
for i in range(0, 10):
    val = raw.iat[i, 0] if i < len(raw) else None
    if isinstance(val, str) and ":" in val:
        k, v = val.split(":", 1)
        metadata[k.strip()] = v.strip()
    elif isinstance(val, str) and val.strip():
        metadata.setdefault("Info", []).append(val.strip())

wide = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=[9,10], engine="openpyxl")
wide = wide.dropna(axis=1, how="all")

def find_col(df, top_label):
    for c in df.columns:
        if isinstance(c, tuple) and c[0] == top_label:
            return c
        if not isinstance(c, tuple) and c == top_label:
            return c
    return None

well_col = find_col(wide, "Well")
content_col = find_col(wide, "Content")

if isinstance(wide.columns, pd.MultiIndex):
    time_cols = [c for c in wide.columns if c[0] == CHANNEL_LABEL]
else:
    time_cols = []

time_labels = [str(c[1]).strip() for c in time_cols]

wide_clean = wide[[well_col, content_col] + time_cols].dropna(subset=[well_col]).copy()
wide_clean.columns = ["Well", "Content"] + time_labels

wells = []
for _, row in wide_clean.iterrows():
    values = [None if pd.isna(row[t]) else float(row[t]) for t in time_labels]
    wells.append({
        "well": str(row["Well"]),
        "content": None if pd.isna(row["Content"]) else str(row["Content"]),
        "values": values
    })

all_vals = [v for w in wells for v in w["values"] if v is not None]
y_min = min(all_vals) if all_vals else 0.0
y_max = max(all_vals) if all_vals else 1.0

payload = {
    "metadata": metadata,
    "sheet": SHEET_NAME,
    "channel": CHANNEL_LABEL.strip(),
    "timeLabels": time_labels,
    "wells": wells,
    "yMin": y_min,
    "yMax": y_max,
}

print(json.dumps(payload, ensure_ascii=False))
