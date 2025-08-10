import json
from pathlib import Path
import pandas as pd

# Project-root relative path
p = Path(__file__).resolve().parents[1] / 'data' / 'raw' / '20250617  HoFF TXA buffer-1.xlsx'
# Read with two header rows (row 10 and 11 are headers, 1-based)
df = pd.read_excel(p, sheet_name='Table All Cycles', header=[9,10], engine='openpyxl')
# Drop completely empty columns
df = df.dropna(axis=1, how='all')

# Identify time columns under the measurement channel
channel_label = ' Raw Data (625-30/680-30)'
if isinstance(df.columns, pd.MultiIndex):
    time_cols = [c for c in df.columns if c[0] == channel_label]
    time_labels = [str(c[1]).strip() for c in time_cols]
    # Extract Well and Content/Time columns robustly
    well_col = None
    for c in df.columns:
        if (isinstance(c, tuple) and c[0] == 'Well') or (not isinstance(c, tuple) and c == 'Well'):
            well_col = c
            break
    content_time_col = None
    for c in df.columns:
        if isinstance(c, tuple) and c[0] == 'Content' and str(c[1]).strip().lower() == 'time':
            content_time_col = c
            break
else:
    time_cols = []
    time_labels = []
    well_col = 'Well'
    content_time_col = 'Time'

# Build a compact summary
wells = df[well_col].dropna().astype(str)
summary = {
    'sheet': 'Table All Cycles',
    'rows': int(df.shape[0]),
    'cols': int(df.shape[1]),
    'channel_label': channel_label.strip(),
    'num_timepoints': len(time_cols),
    'first_time_labels': time_labels[:10],
    'last_time_labels': time_labels[-5:],
    'num_wells': int(wells.shape[0]),
    'first_wells': wells.head(10).tolist(),
}
print(json.dumps(summary, indent=2))
