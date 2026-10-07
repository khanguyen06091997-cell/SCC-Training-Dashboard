# Nguồn build SCC Training Dashboard

- `head.html` + `body.html` + `app.js` + logo + `sample_v9.json` → `build_page.py` → `scc-training-dashboard.html` (= `index.html` ở gốc repo).
- `export_pdf.py <html> <SCC Training Record.xlsx> <out.pdf>`: mở dashboard bằng Playwright, nạp file Excel, bấm nút Xuất PDF.
- Build: `cd _src && python3 build_page.py && cp scc-training-dashboard.html ../index.html`
