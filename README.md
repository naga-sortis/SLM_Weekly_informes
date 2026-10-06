# SLM Weekly Informes

Static web page that turns the weekly Oceane / BRISE ticket export (Excel) into the
**Informe SLM-OSS – Gestores Propietarios** (Word) and an Excel file with the extracted fields.

Everything runs **in the browser**: the Excel file is never uploaded anywhere. All libraries are bundled in `vendor/`, so the page needs no CDN.

## How to use

1. Open the page and drop the ticket export (`.xlsx`, `.xls` or `.csv`).
   The sheet with the `Ticket ID` header row is found automatically. If there are several, the largest one is used, and you can switch sheets.
2. Pick the **report week**. By default it's the latest week in the file. The week number, its dates (Monday–Sunday) and its month are used everywhere, on the page and in Word. A week that spans two months is labelled e.g. “Septiembre / Octubre 2026”. The tables also show the previous weeks (5 by default, 1–8).
3. Check the **consistency checks** panel. If any check fails, the Word export is blocked.
4. Optionally fill in the cover, revision table and the “Incidencias destacadas / Otros trabajos” bullet points.
5. **Download Word report** or **Download extracted data (Excel)**.

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
* **All tickets count, whoever opened them.** The SLM groups (XSP00025 / XSP00027) are treated like any other initiator. The row **Abiertas por el SLM** only shows how many of the new cases those groups opened.
* **Casos por gestor**: every new incidencia of each week is counted once, under the gestor from the first part of `Current action` (`GESTOR - PROBLEMA - TÉCNICO`). Spelling variants are merged (NFM-T / NFMT / NFM -T, ENM 3 / ENM3…). Incidencias with no gestor in the action (`JR - Trabajando`, empty, `DEVUELTO`) appear as *Sin gestor identificado*, so the weekly totals always equal “Nuevos durante la semana”. There is an optional switch, off by default, that takes the gestor from the description instead; the report states when it's used.
* **Cut-off**: nothing created or resolved after the end of the report week is counted, including in the monthly charts. A past week therefore always gives the same figures.
* Week, month and year are recalculated from the creation date (ISO weeks), and the file's own columns are cross-checked against them.

## Month / Year view

The **Month / Year** tab filters all incidencias + OTs created in a calendar month or a whole year. It works independently of the report week. It shows:

* a summary table: opened, resolved, pending at the close, returned, escalated, opened by the SLM, resolved during the period, total pending;
* charts per month (for a year) or per week of the month (weeks are trimmed to the month's days): cases opened and resolved by priority, status, and escalations by vendor;
* cases by gestor (top 15 chart plus the full table, with incidencias and OTs).

With “Include in the Word report” ticked, this becomes **section 4** of the Word document. It has its own consistency checks (independent recount from the creation date, opened = resolved + pending + returned, charts and gestor table adding up). If any check fails, the export is blocked.

## Tickets without gestor

On the *Cases by gestor* and *Month / Year* tabs, a **Sin gestor identificado** panel lists every ticket counted in that row: its Ticket ID, date, type, status and Current action exactly as written in the Excel. Each ticket also shows the **reason** it has no gestor:

* action empty;
* technician / work status (e.g. `JR - Trabajando`);
* `DEVUELTO`;
* `Cerrado`;
* format not recognised.

The panel also shows any gestor mentioned in the description, as a hint. You can filter by reason, search, and export the list to Excel. Only “format not recognised” would point to a reading problem in the tool, and the panel says explicitly whether there are any.

## Numbers on the charts

Every bar shows its total. Under each chart's legend there is a table with the number of cases per series (OTs, Inc. P1–P4, statuses, vendors…), *Total incidencias* / *Total OTs* subtotals and a *Total* row. The same tables appear under each chart in the Word report.

## Consistency checks

Every report is verified before it can be exported:

* Nuevos = Resueltas + Sin resolver (no esc.) + Sin resolver (esc.) + Devueltas, for every week and both types.
* Backlog = Resueltos + Sin resolver, and each week's backlog = the previous week's pending (backlog + new).
* “Nuevos” equals the tickets actually created in each week.
* 2.2 totals per week = “Nuevos durante la semana” (Incidencias), and each gestor appears only once.
* 2.3 TOTAL = the number of cases listed in 2.4, and every listed case was really pending at the cut-off.
* The monthly charts cover every case of the month exactly once, and the report month contains the report week.

Unit tests (synthetic data only) run with `node --test tests/core.test.js`. They also run in the GitHub Actions workflow before every deployment.

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
