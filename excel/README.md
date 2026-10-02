# Hermes for Excel

The Excel add-in for [Hermes for Office](../README.md) — a task-pane chat that reads your sheet and edits it on approval, plus `=HERMES.*` custom functions that call the agent from a cell.

Two surfaces, one add-in (Office.js, shared runtime):

- **Task pane** — chat about the active sheet; it proposes an action plan (edit cells, format, rename tab, create tables/charts, add sheets); you review a preview and click **Apply**. Nothing changes until you approve.
- **Custom functions** — `=HERMES.CLASSIFY`, `=HERMES.EXTRACT`, `=HERMES.SUMMARIZE`, `=HERMES.FORMULA_HELP`. Read-only, cached, drag-fillable.

## Prerequisites

- The **shared backend running** (Hermes API Server + Caddy) — see the [root README](../README.md#shared-backend-set-up-once-used-by-every-app).
- Node.js 18+.
- Excel (desktop, Windows or Mac) with any Microsoft 365 account.

## Run

```
npm install
npm start
```
This trusts a dev cert, serves the add-in on `https://localhost:3000`, sideloads it, and opens Excel. Then: **Home tab → Show Taskpane**.

## Usage

**Task pane:** open your data tab and type e.g. *"clean this up: fix casing, format the P&L column as currency, rename the tab to 'Acquisitions 2021', and add a chart of P&L by collection."* Review the proposed actions → **Apply**.

**Custom functions:**
```
=HERMES.CLASSIFY(A2, "lead quality: hot/warm/cold")
=HERMES.EXTRACT(A2, "company name")
=HERMES.SUMMARIZE(A1:D20)
=HERMES.FORMULA_HELP("year-over-year growth")
```
Text arguments must be quoted; the first arg can be a cell reference. Drag-fill down a column.

## Gotchas

- **CORS allowed headers** are `Authorization, Content-Type, Idempotency-Key`. Don't send other custom headers from the client or the browser preflight fails ("Failed to fetch").
- **Shared-runtime caching:** after editing add-in code, the pane may keep running the old bundle. Right-click the pane → Reload; if it persists, quit Excel and `npm start` again.
- **Custom functions** use a dotted namespace: `=HERMES.CLASSIFY`, not an underscore.
