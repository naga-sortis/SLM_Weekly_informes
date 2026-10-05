# SLM Weekly Informes

Static web page that turns the weekly Oceane / BRISE ticket export (Excel) into the
**Informe SLM-OSS – Gestores Propietarios** (Word) and an Excel file with the extracted fields.

Everything runs **in the browser**: the Excel file is never uploaded anywhere. All libraries are bundled in `vendor/`, so the page needs no CDN.

## How to use

1. Open the page and drop the ticket export (`.xlsx`, `.xls` or `.csv`).
   The sheet with the `Ticket ID` header row is found automatically. If there are several, the largest one is used, and you can switch sheets.
2. Pick the **report month**. This selects its ISO weeks (the weeks whose Monday falls in that month, e.g. April 2024 → weeks 14–18). You can also set the week range by hand (up to 8 weeks).
3. Optionally fill in the cover, revision table and the “Incidencias destacadas / Otros trabajos” bullet points.
4. **Download Word report** or **Download extracted data (Excel)**.

## Extracted fields

| # | Field | Source column |
|---|-------|---------------|
| 1 | Ticket ID | `Ticket ID` (duplicates ignored) |
| 2 | Third party reference → vendor | `Third party reference`: `STA-` Ericsson, `H-` Huawei, `1-` Nokia |
| 3 | Creation date | `Creation date` |
| 4–6 | Creation week / month / year | `Creation week`, `Creation month`, `Creation year` |
| 7 | Processing priority and Failure / OT | `Processing priority`, `Ticket type` (`Failure` = Incidencia, any other type = OT) |
| 8 | Status | `Status` |
| 9 | Initiator group ID | `Initiator - Group ID` |
| 10 | Initiator group abbreviation name | `Initiator - Group abbreviation name` |
| 11 | Current action → gestor / problema / técnico | `Current action` (`GESTOR - PROBLEMA - TECNICO`) |

## Report rules

* **Escalada**: the third party reference belongs to a vendor (Ericsson, Huawei or Nokia).
* **Devuelta**: `Current action` is `DEVUELTO`.
* **Resuelta**: the status is not `Current` and the `Restoration date` falls before the end of the week. Past weeks therefore show the situation as it was then.
* **Backlog**: cases created before the week that were not returned and were still unresolved when the week started.
* **Abiertas por el SLM**: `Initiator - Group ID` is one of the SLM groups (default `XSP00025, XSP00027`, which you can change on the page).
* **Casos por gestor**: new incidencias per week, grouped by the first part of `Current action`. Technician statuses such as `JR - Trabajando` or `GV - ACCESO …` have no gestor.
* Weeks are ISO weeks calculated from the creation date.

## Hosting on GitHub Pages

Option A (recommended): **Settings → Pages → Source: GitHub Actions**. The workflow in `.github/workflows/pages.yml` publishes the site on every push to `main`. You can also run it manually from the Actions tab.

Option B: **Settings → Pages → Source: Deploy from a branch**, then choose the branch and `/ (root)`.

To run it locally, run `python3 -m http.server` in this folder and open <http://localhost:8000>.

## Project layout

```
index.html               page
assets/css/styles.css    styles (light/dark)
assets/js/core.js        parsing, field extraction and report calculations (no DOM; testable in Node)
assets/js/charts.js      Chart.js configs and PNG rendering for Word
assets/js/docx-report.js Word document builder
assets/js/app.js         UI
vendor/                  SheetJS 0.20.3 (Apache-2.0), Chart.js 4 (MIT), docx 9 (MIT)
```
