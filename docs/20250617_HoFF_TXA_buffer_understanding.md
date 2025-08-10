### Workbook understanding: `20250617  HoFF TXA buffer-1.xlsx`

- **sheet**: `Table All Cycles`
- **metadata (rows 1–9)**: user, test ID, test name `HoFF 647 60m`, date/time, modality `Fluorescence (FI)`
- **header**: two-row header at rows 10–11 (1-based). Columns: `Well`, `Content`, then 60 timepoints under channel `Raw Data (625-30/680-30)`
- **timepoints**: `0 min` … `59 min` (60 total)
- **data region**: rows 10–31, up to column 62
- **wells**: 20 wells present (e.g., `A01, A02, A05, A06, B01, B02, B05, B06, C01, C02, …`)
- **shape (parsed)**: 20 wells × 60 timepoints

### Load with pandas
```python
import pandas as pd
from pathlib import Path

p = Path('20250617  HoFF TXA buffer-1.xlsx')
# Two-row header at rows 10–11 (1-based) => header=[9,10]
df = pd.read_excel(p, sheet_name='Table All Cycles', header=[9,10], engine='openpyxl').dropna(axis=1, how='all')

channel = ' Raw Data (625-30/680-30)'
time_cols = [c for c in df.columns if isinstance(c, tuple) and c[0] == channel]
well_col = next(c for c in df.columns if (isinstance(c, tuple) and c[0]=='Well') or c=='Well')
content_col = next(c for c in df.columns if (isinstance(c, tuple) and c[0]=='Content') or c=='Content')

wide = df[[well_col, content_col] + time_cols].dropna(subset=[well_col])
wide.columns = ['Well','Content'] + [str(c[1]).strip() for c in time_cols]

# Long format (one row per well-time)
long = wide.melt(id_vars=['Well','Content'], var_name='Time', value_name='FI')
```

### Notes
- The top-level header for time series includes a leading space in the label (`" Raw Data (625-30/680-30)"`).
- Time labels are strings like `"0 min"` … `"59 min"`.

