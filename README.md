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
| 2 | Third party reference → vendor | `Third party reference`: `STA-` or `CSR` Ericsson, `H-` Huawei, `1-` Nokia |
| 3 | Creation date | `Creation date` |
| 4–6 | Creation week / month / year | `Creation week`, `Creation month`, `Creation year` |
| 7 | Processing priority and Failure / OT | `Processing priority`, `Ticket type` (`Failure` = Incidencia, any other type = OT; a technician / work status in `Current action`, e.g. `JR - Trabajando`, = OT whatever the type) |
| 8 | Status | `Status` |
| 9 | Initiator group ID | `Initiator - Group ID` |
| 10 | Initiator group abbreviation name | `Initiator - Group abbreviation name` |
| 11 | Current action → gestor / problema / técnico | `Current action` (`GESTOR - PROBLEMA - TECNICO`) |

## Report rules

* **Escalada**: the third party reference belongs to a vendor: `STA-` or `CSR` → Ericsson, `H-` → Huawei, `1-` → Nokia.
* **Devuelta**: `Current action` is `DEVUELTO`.
* **Resuelta**: the status is not `Current` and the `Restoration date` falls before the end of the week. Past weeks therefore show the situation as it was then.
* **Backlog**: cases created before the week that were not returned and were still unresolved when the week started.
* **All tickets count, whoever opened them.** The SLM groups (XSP00025 / XSP00027) are treated like any other initiator. The row **Abiertas por el SLM** only shows how many of the new cases those groups opened.
* **Casos por gestor**: every new incidencia of each week is counted once, under the gestor from the first part of `Current action` (`GESTOR - PROBLEMA - TÉCNICO`). Spelling variants are merged (NFM-T / NFMT / NFM -T, ENM 3 / ENM3…). Incidencias with no gestor in the action (`JR - Trabajando`, empty, `DEVUELTO`) appear as *Sin gestor identificado*, so the weekly totals always equal “Nuevos durante la semana”. There is an optional switch, off by default, that takes the gestor from the description instead; the report states when it's used.
* **Cut-off**: nothing created or resolved after the end of the report week is counted, including in the monthly charts. A past week therefore always gives the same figures.
* Week, month and year are recalculated from the creation date (ISO weeks), and the file's own columns are cross-checked against them.

## Trend charts (section 3)

*Trend charts* in step 2 offers either the **last 3 / 6 / 7 / 12 months** up to the report week (default: 7), or **one calendar year**. The years run from the latest to the oldest in the file, up to the report week's year. With a year selected, the charts, their number tables and the Word report show only that year: January–December, or January to the report month for the current year.

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

## All details & who dealt with each ticket

* **Every column of every ticket sheet** is read, including Acknowledgement date, Initiator user, Identifier 1–4, Calculated duration, Short label, Final nature, Restoration group / name / **user**, and Closure group / name / user. Rows that repeat a ticket are merged rather than dropped. If they differ (e.g. several closure groups), every value is kept and the difference is flagged.
* **Click any Ticket ID** (data table, unresolved cases, “Sin gestor”, Sortis team) to open its detail window. It shows a *who dealt with it* timeline (opened by → acknowledged → restored by → closed by), what the tool interpreted (gestor, vendor, time to restore…), and every column exactly as written in the Excel.
* **Extracted data** tab: a column chooser covering all columns, plus filters by restoration group and restoration user.
  * **Export ALL data** downloads every ticket (ignoring the filters) with every column. It adds sheets to cross-check the analysis:
    * *Informe semanal*;
    * *Original rows (all sheets)*: every ticket of every sheet exactly as written in the Excel, the sheets it appears in, and any other values found in repeated rows;
    * *Checks*: every consistency check with OK / FAILED;
    * *About this export*: the file, the sheets read, counts, settings, the rules applied and any warnings.
  * **Export filtered** downloads only the rows currently shown.

## Sortis team (XSP00025 / XSP00027)

The **Sortis team** tab uses **all ticket sheets of the file**. It takes the tickets whose *Restoration group ID* is one of the SLM / Sortis group IDs, and the engineer is the *Restoration user name*. Filters: year / month (by restoration date or creation date) and queue. It shows:

* totals per queue;
* a table per engineer: incidencias / OTs per queue, median time to restore, tickets also closed by the same person, first and last case;
* charts per engineer and their evolution over time;
* the list of tickets, filterable by engineer, with the detail window and an Excel export;
* all restoration groups in the period, to see who else dealt with the cases.

It has its own consistency checks. Tickets without a restoration date are counted under “All years” and reported, so none are lost.

## Choosing what goes into Word

In *Word report details → Sections to include*, tick the sections you want. Section numbering and the table of contents adapt automatically. The default selection produces exactly the same report as before. The Month / Year and Sortis team sections can be added from there or from their tabs.

## Consistency checks

Every report is verified before it can be exported:

* Nuevos = Resueltas + Sin resolver (no esc.) + Sin resolver (esc.) + Devueltas, for every week and both types.
* Backlog = Resueltos + Sin resolver, and each week's backlog = the previous week's pending (backlog + new).
* “Nuevos” equals the tickets actually created in each week.
* 2.2 totals per week = “Nuevos durante la semana” (Incidencias), and each gestor appears only once.
* 2.3 TOTAL = the number of cases listed in 2.4, and every listed case was really pending at the cut-off.
* The monthly charts cover every case of the month exactly once, and the report month contains the report week.

Unit tests (synthetic data only) run with `npm test` (Node.js ≥ 22.13). They also run in the GitHub Actions workflow before every deployment.

## Source sheets tab

The first tab analyses every sheet in the workbook (Incidencias, OTs, Asignados… whatever is present) before anything else: header row, rows and unique tickets, duplicated rows, invalid dates, date range, Incidencias vs OTs, types, statuses, years, restoration groups, missing expected columns and warnings. It also shows how many tickets each pair of sheets has in common, and lets you browse, search and export the raw rows of any sheet.

## Manual review of “Sin gestor identificado”

In the *Sin gestor identificado* panel, each case has **Review: Incidencia / OT** buttons (also available in bulk for the filtered list and in the ticket dialog). A reviewed case:

* is counted under the chosen category in every section (weekly, Month/Year, Sortis team, Word report) — the automatic category is kept and shown as “Auto”;
* keeps your **name and the date** of the decision (column “Counted as”, ticket dialog, Excel export and Word section 2.4);
* can be undone at any time (“Undo” restores the automatic rule).

On the server the reviews are shared by the whole team and every change is logged. On GitHub Pages / file mode they are saved in this browser only.

## Running on a server (team database)

The same page can run on a server with a shared database of every weekly Excel. It needs **Node.js ≥ 22.13** and no other package (SQLite is built into Node).

```
npm start                       # http://<server-ip>:8080/
```

| Variable | Default | Meaning |
|---|---|---|
| `PORT` | `8080` | port |
| `HOST` | `0.0.0.0` | interface to bind (`127.0.0.1` behind a reverse proxy) |
| `SLM_DB` | `data/slm-informes.sqlite` | database file |
| `SLM_BASIC_AUTH` | — | `user:password` to require a login |
| `SLM_MAX_UPLOAD_MB` | `300` | largest upload accepted |

Weekly workflow:

1. Upload the new weekly Excel. The report is computed in the browser as before, and a **preview** shows what would change in the database (new tickets, updated tickets, tickets not in this file).
2. Type your name and click **Update database**. New rows are added, changed rows are updated (each changed field is logged with old/new value), and tickets missing from the new file are kept.
3. **Use database** builds the report from the full history; **Import history** lists every import; **Backup** downloads a copy of the database.

Notes for production:

* Only the page files are served (never the database, server code or tests); security headers and a strict Content-Security-Policy are set.
* Without `SLM_BASIC_AUTH` there is no login: restrict access with the firewall/VPN. For HTTPS put a reverse proxy (nginx, IIS, Apache) in front and bind `HOST=127.0.0.1`.
* Back up the `data/` folder (or use the Backup button) regularly. The database contains personal data from the tickets — keep it on the server only; it is excluded from git.
* Run it as a service, e.g. systemd:

```
[Unit]
Description=SLM Weekly Informes
After=network.target

[Service]
WorkingDirectory=/opt/slm-weekly-informes
ExecStart=/usr/bin/npm start
Environment=PORT=8080 SLM_DB=/var/lib/slm/slm-informes.sqlite
Restart=on-failure
User=slm

[Install]
WantedBy=multi-user.target
```

On Windows Server use NSSM or the Task Scheduler (“At startup”) to run `npm start` in the project folder.

When the page is opened from GitHub Pages or a plain static server, it detects that there is no database (the browser console shows one expected 404 for `api/health`) and works exactly as before, entirely in the browser.

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
server/server.js         optional Node server (page + team database API)
server/db.js             SQLite database: imports, change log, manual classifications
tests/                   unit tests (synthetic data)
vendor/                  SheetJS 0.20.3 (Apache-2.0), Chart.js 4 (MIT), docx 9 (MIT)
```
