import json
from pathlib import Path
import pandas as pd

# Project-root relative paths
ROOT = Path(__file__).resolve().parents[1]
EXCEL_PATH = ROOT / 'data' / 'raw' / "20250617  HoFF TXA buffer-1.xlsx"
SHEET_NAME = "Table All Cycles"
CHANNEL_LABEL = " Raw Data (625-30/680-30)"

# 1) Extract metadata from the first column of the top rows
raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None, engine="openpyxl")
metadata = {}
for i in range(0, 10):
    val = raw.iat[i, 0] if i < len(raw) else None
    if isinstance(val, str) and ":" in val:
        k, v = val.split(":", 1)
        metadata[k.strip()] = v.strip()
    elif isinstance(val, str) and val.strip():
        # Single string like modality
        metadata.setdefault("Info", []).append(val.strip())

# 2) Load with two-row header and identify time columns, wells, content
wide = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=[9,10], engine="openpyxl")
wide = wide.dropna(axis=1, how="all")

# Find columns
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

# Keep only rows with a well id
wide_clean = wide[[well_col, content_col] + time_cols].dropna(subset=[well_col]).copy()
wide_clean.columns = ["Well", "Content"] + time_labels

# Build wells list
wells = []
for _, row in wide_clean.iterrows():
    values = [None if pd.isna(row[t]) else float(row[t]) for t in time_labels]
    wells.append({
        "well": str(row["Well"]),
        "content": None if pd.isna(row["Content"]) else str(row["Content"]),
        "values": values
    })

# Global y-axis min/max across all values
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

# 3) Generate a self-contained HTML
html = f"""
<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>HoFF TXA Buffer Report</title>
  <style>
    body {{ font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; margin: 16px; }}
    header {{ margin-bottom: 16px; }}
    .summary {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px; margin-bottom: 16px; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 12px; background: #fafafa; }}
    .row {{ display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin: 8px 0; }}
    select, button {{ padding: 6px 10px; font-size: 14px; }}
    #chart {{ border: 1px solid #eee; background: #fff; border-radius: 8px; }}
    .table-wrap {{ overflow: auto; border: 1px solid #eee; border-radius: 8px; margin-top: 16px; }}
    table {{ border-collapse: collapse; font-size: 12px; min-width: max(900px, 100%); }}
    th, td {{ padding: 6px 8px; border: 1px solid #eee; white-space: nowrap; }}
    th {{ position: sticky; top: 0; background: #f7f7f7; z-index: 1; }}
    .muted {{ color: #666; }}
  </style>
</head>
<body>
  <header>
    <h2>HoFF TXA Buffer Report</h2>
    <div class=\"summary\">
      <div class=\"card\"><b>Sheet</b><div id=\"sheetName\"></div></div>
      <div class=\"card\"><b>Channel</b><div id=\"channel\"></div></div>
      <div class=\"card\"><b>Wells</b><div id=\"numWells\"></div></div>
      <div class=\"card\"><b>Timepoints</b><div id=\"numTimes\"></div></div>
    </div>
    <div class=\"card\">
      <b>Metadata</b>
      <div id=\"meta\" class=\"muted\"></div>
    </div>
  </header>

  <section>
    <div class=\"row\">
      <label for=\"wellSelect\"><b>Select well</b></label>
      <select id=\"wellSelect\"></select>
      <span class=\"muted\" id=\"wellContent\"></span>
    </div>
    <svg id=\"chart\" width=\"960\" height=\"360\"></svg>
  </section>

  <section class=\"table-wrap\">
    <table id=\"dataTable\"></table>
  </section>

  <script>
    const DATA = {json.dumps(payload)};

    // Populate summary
    document.getElementById('sheetName').textContent = DATA.sheet;
    document.getElementById('channel').textContent = DATA.channel;
    document.getElementById('numWells').textContent = DATA.wells.length;
    document.getElementById('numTimes').textContent = DATA.timeLabels.length;

    // Metadata
    const metaDiv = document.getElementById('meta');
    const md = DATA.metadata || {};
    const lines = [];
    for (const k of Object.keys(md)) {{
      const v = md[k];
      if (Array.isArray(v)) {{
        for (const item of v) lines.push(item);
      }} else {{
        lines.push(`${{k}}: ${{v}}`);
      }}
    }}
    metaDiv.textContent = lines.join('  |  ');

    // Build well selector
    const wellSelect = document.getElementById('wellSelect');
    for (const w of DATA.wells) {{
      const opt = document.createElement('option');
      opt.value = w.well; opt.textContent = w.well; wellSelect.appendChild(opt);
    }}

    const wellContent = document.getElementById('wellContent');
    const svg = document.getElementById('chart');

    function renderChart(wellId) {{
      const w = DATA.wells.find(x => x.well === wellId) || DATA.wells[0];
      wellContent.textContent = w && w.content ? `Content: ${'{'}w.content{'}'}` : '';
      const width = svg.viewBox.baseVal && svg.viewBox.baseVal.width ? svg.viewBox.baseVal.width : svg.getAttribute('width');
      const height = svg.viewBox.baseVal && svg.viewBox.baseVal.height ? svg.viewBox.baseVal.height : svg.getAttribute('height');
      const W = +width, H = +height;

      // Clear
      while (svg.firstChild) svg.removeChild(svg.firstChild);

      // Padding and axes
      const pad = {{ left: 50, right: 20, top: 20, bottom: 40 }};
      const innerW = W - pad.left - pad.right;
      const innerH = H - pad.top - pad.bottom;

      // Scales
      const n = DATA.timeLabels.length;
      const x = i => pad.left + (i/(Math.max(1,n-1))) * innerW;
      const y = v => pad.top + innerH - ( (v - DATA.yMin) / (Math.max(1, DATA.yMax - DATA.yMin)) ) * innerH;

      // Axes lines
      const axis = document.createElementNS('http://www.w3.org/2000/svg','path');
      axis.setAttribute('d', `M${'{'}pad.left{'}'},${'{'}pad.top{'}'} v${'{'}innerH{'}'} h${'{'}innerW{'}'}`);
      axis.setAttribute('stroke', '#aaa');
      axis.setAttribute('fill', 'none');
      axis.setAttribute('stroke-width', '1');
      svg.appendChild(axis);

      // Y ticks (5)
      for (let i=0; i<=5; i++) {{
        const tv = DATA.yMin + (i/5)*(DATA.yMax - DATA.yMin);
        const ty = y(tv);
        const grid = document.createElementNS('http://www.w3.org/2000/svg','line');
        grid.setAttribute('x1', pad.left);
        grid.setAttribute('x2', pad.left + innerW);
        grid.setAttribute('y1', ty);
        grid.setAttribute('y2', ty);
        grid.setAttribute('stroke', '#eee');
        svg.appendChild(grid);

        const lbl = document.createElementNS('http://www.w3.org/2000/svg','text');
        lbl.setAttribute('x', pad.left - 6);
        lbl.setAttribute('y', ty + 4);
        lbl.setAttribute('text-anchor', 'end');
        lbl.setAttribute('font-size', '11');
        lbl.textContent = tv.toFixed(0);
        svg.appendChild(lbl);
      }}

      // X ticks (every 10 min)
      for (let i=0; i<n; i+=10) {{
        const tx = x(i);
        const lbl = document.createElementNS('http://www.w3.org/2000/svg','text');
        lbl.setAttribute('x', tx);
        lbl.setAttribute('y', pad.top + innerH + 20);
        lbl.setAttribute('text-anchor', 'middle');
        lbl.setAttribute('font-size', '11');
        lbl.textContent = DATA.timeLabels[i];
        svg.appendChild(lbl);
      }}

      if (!w) return;

      // Line path
      let d = '';
      for (let i=0; i<n; i++) {{
        const v = w.values[i];
        if (v == null) continue;
        const px = x(i), py = y(v);
        d += (d ? ' L' : 'M') + px + ',' + py;
      }}
      const path = document.createElementNS('http://www.w3.org/2000/svg','path');
      path.setAttribute('d', d);
      path.setAttribute('fill', 'none');
      path.setAttribute('stroke', '#1976d2');
      path.setAttribute('stroke-width', '2');
      svg.appendChild(path);

      // Points (hover)
      for (let i=0; i<n; i++) {{
        const v = w.values[i];
        if (v == null) continue;
        const cx = x(i), cy = y(v);
        const c = document.createElementNS('http://www.w3.org/2000/svg','circle');
        c.setAttribute('cx', cx);
        c.setAttribute('cy', cy);
        c.setAttribute('r', 3);
        c.setAttribute('fill', '#1976d2');
        c.setAttribute('opacity', '0.8');
        c.addEventListener('mouseenter', () => {{ c.setAttribute('r','5'); }});
        c.addEventListener('mouseleave', () => {{ c.setAttribute('r','3'); }});
        svg.appendChild(c);
      }}
    }}

    wellSelect.addEventListener('change', (e) => renderChart(e.target.value));
    // Initial render
    if (DATA.wells.length) {{
      wellSelect.value = DATA.wells[0].well;
      renderChart(DATA.wells[0].well);
    }}

    // Build wide table
    const table = document.getElementById('dataTable');
    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
    const thWell = document.createElement('th'); thWell.textContent = 'Well'; trh.appendChild(thWell);
    const thContent = document.createElement('th'); thContent.textContent = 'Content'; trh.appendChild(thContent);
    for (const t of DATA.timeLabels) {{
      const th = document.createElement('th'); th.textContent = t; trh.appendChild(th);
    }}
    thead.appendChild(trh); table.appendChild(thead);

    const tbody = document.createElement('tbody');
    for (const w of DATA.wells) {{
      const tr = document.createElement('tr');
      const tdW = document.createElement('td'); tdW.textContent = w.well; tr.appendChild(tdW);
      const tdC = document.createElement('td'); tdC.textContent = w.content || ''; tr.appendChild(tdC);
      for (const v of w.values) {{
        const td = document.createElement('td'); td.textContent = (v==null? '' : v); tr.appendChild(td);
      }}
      tbody.appendChild(tr);
    }}
    table.appendChild(tbody);
  </script>
</body>
</html>
"""

out_html = ROOT / 'reports' / 'hoff_txa_report.html'
out_html.write_text(html, encoding="utf-8")
print(f"Wrote {out_html}")
