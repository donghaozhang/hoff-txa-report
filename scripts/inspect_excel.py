import json
from pathlib import Path
import pandas as pd

# Project-root relative path to raw data
path = Path(__file__).resolve().parents[1] / 'data' / 'raw' / '20250617  HoFF TXA buffer-1.xlsx'
xl = pd.ExcelFile(path)
print('SHEETS_START')
print('|'.join(xl.sheet_names))
print('SHEETS_END')

for sheet in xl.sheet_names:
    print('SHEET_START')
    print(sheet)
    df = pd.read_excel(path, sheet_name=sheet, header=None, engine='openpyxl')
    nrows, ncols = df.shape
    print('SHAPE_START')
    print(f"{nrows}|{ncols}")
    print('SHAPE_END')

    # Non-empty count per row (first 80 rows)
    ne = df.notna().sum(axis=1).tolist()
    first_rows = [(i+1, int(ne[i])) for i in range(min(80, len(ne)))]
    print('ROW_NONEMPTY_START')
    print(json.dumps(first_rows))
    print('ROW_NONEMPTY_END')

    # Try to find a header row containing the string 'Well'
    header_row_idx = None
    for i in range(len(df)):
        row_vals = df.iloc[i].astype(str).tolist()
        if any(val.strip().lower() == 'well' for val in row_vals):
            header_row_idx = i
            break
    print('HEADER_ROW_START')
    print('' if header_row_idx is None else str(header_row_idx+1))
    print('HEADER_ROW_END')

    # If header found, show that row and next 5 rows (first 20 columns)
    if header_row_idx is not None:
        sl = df.iloc[header_row_idx: header_row_idx+6, :20]
        print('HEADER_SLICE_START')
        print(sl.to_json(orient='split', index=True, date_format='iso'))
        print('HEADER_SLICE_END')

        # Infer used range below header: last row & last col with any data
        below = df.iloc[header_row_idx:, :]
        non_empty_rows = below.dropna(how='all')
        non_empty_cols = below.dropna(axis=1, how='all')
        last_row_idx = non_empty_rows.index[-1] if not non_empty_rows.empty else header_row_idx
        last_col_idx = non_empty_cols.columns[-1] if not non_empty_cols.empty else 0
        print('USED_RANGE_START')
        print(json.dumps({
            'start_row': int(header_row_idx+1),
            'end_row': int(last_row_idx + 1),
            'end_col': int(last_col_idx + 1)
        }))
        print('USED_RANGE_END')

        # Show a 10x15 sample area starting 1 row after header
        r0 = header_row_idx+1
        r1 = min(r0+10, nrows)
        c1 = min(15, ncols)
        sample = df.iloc[r0:r1, :c1]
        print('DATA_SAMPLE_START')
        print(sample.to_json(orient='split', index=True, date_format='iso'))
        print('DATA_SAMPLE_END')
    else:
        # Fallback: show top-left 20x10 slice
        sample = df.iloc[:20, :10]
        print('DATA_SAMPLE_START')
        print(sample.to_json(orient='split', index=True, date_format='iso'))
        print('DATA_SAMPLE_END')

    print('SHEET_END')
