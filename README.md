# CPA IT — Chromebook tools

Static tools hosted on GitHub Pages, backed by a Google Apps Script + Google Sheet.

## Live URLs
- **Inventory scanner** — scan Chromebook QR/serials, export CSV/templates
  https://kylewandek-sketch.github.io/cpa-it/
- **Help desk ticket** — scan + describe an issue, logs to the Sheet and emails
  https://kylewandek-sketch.github.io/cpa-it/help-desk/
- **IT ticket dashboard** — manage/sort/filter/edit tickets (admin token required)
  https://kylewandek-sketch.github.io/cpa-it/dashboard/

## Backend
- `apps-script.gs` — deploy in the ticket Google Sheet (Extensions ▸ Apps Script) as a Web app.
  Handles form submissions (doPost) and the dashboard's read/edit endpoints (doGet, JSONP).
- The Google Sheet bound to that script is the single source of truth (persistence).

## Layout
- `index.html` — inventory scanner (+ `templates.json` for shared export templates)
- `help-desk/index.html` — ticket submission page
- `dashboard/index.html` — management dashboard
