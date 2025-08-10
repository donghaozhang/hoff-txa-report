## HoFF TXA Buffer Report

Self-contained scripts to parse a CLARIOstar export and generate an interactive HTML report with well-level time courses and correlation analysis.

### Project structure
- `data/raw/`: original Excel exports (excluded from Git)
- `data/processed/`: derived artifacts (e.g., JSON payloads)
- `scripts/`: Python utilities to inspect, parse, and build reports
- `reports/`: generated HTML reports
- `docs/`: notes and understanding

### Quickstart
1) Install dependencies
```bash
pip install -r requirements.txt
```

2) Place your workbook in `data/raw/` (e.g., `20250617  HoFF TXA buffer-1.xlsx`).

3) Generate the report
```bash
python scripts/write_html_direct.py
```
Open `reports/hoff_txa_report.html` in a browser.

### Scripts
- `scripts/inspect_excel.py`: explore sheet structure and header rows
- `scripts/analyze_excel.py`: summarize timepoints/wells
- `scripts/write_html_direct.py`: build a single HTML report (recommended)
- `scripts/build_html_report.py`: alternative generator

### Notes
- Raw data files are ignored by Git. Do not commit sensitive lab data.

### License
MIT

