# Retirement modeler

A year-by-year retirement plan for a couple, kept in an ordinary spreadsheet, checked by a Python copy of the
same model, stress-tested against 98 years of market history. It comes with a pipeline that turns downloaded
bank, card and investment statements into a budget and account balances. It is built to be run **with
Claude** (Claude Code): you bring the documents and the questions, Claude does the bookkeeping, edits the
workbook, runs the tools and keeps a log. See [HOW_TO_USE_WITH_CLAUDE.md](HOW_TO_USE_WITH_CLAUDE.md).

- **United States** (federal + a state of your choice, Social Security, Medicare/IRMAA, 401(k)/IRA/Roth/HSA,
  RMDs, Roth conversions) or **Canada** (federal + British Columbia, CPP/OAS with the clawback,
  RRSP/RRIF/TFSA, pension splitting, RRSP meltdown).
- Kids with their own college and after-college costs, workplace pensions, a home and up to four other
  properties (sell, rent out or move in, by scenario), one-off events (severance, inheritance, a wedding),
  long-term care, inflation per cost type.
- Every number in the workbook is a live formula: it works in Excel, LibreOffice and Google Sheets.
- Monte Carlo: every historical starting year since 1928 plus random block-bootstrap runs, each with lifetimes
  drawn from a life table, so the headline is "the money outlasts you both in X% of lifetimes".

Not financial advice. The tax rules are simplified ([MODEL.md](MODEL.md) lists what's left out); check important
decisions with a professional.

## Try it (5 minutes)

Needs Python 3.9+ and, recommended, [LibreOffice](https://www.libreoffice.org) (free; used to recalculate the
workbook so the scripts can read it; Mac: `brew install --cask libreoffice`).

```
git clone <this repo> retirement-modeler && cd retirement-modeler
./plan demo us        # or: ./plan demo ca
```

The demo copies the example household to `demo-us/`, reads its sample statements, merges the budget and
balances into its workbook, checks every formula against the Python model and runs the Monte Carlo. Open
`demo-us/plan.xlsx` to look around: start on the `Model - Example` tab (yellow cells are inputs; results at the
top right), then `monte_carlo`, which compares the scenarios.

Just want the spreadsheet? `examples/us/plan.xlsx` and `examples/ca/plan.xlsx` are ready to import into Google
Sheets (File > Import > Upload) or open in Excel. They're filled in with fictional households.

## Start your own plan

```
./plan new us ~/retirement-plan      # or ca; keep it OUTSIDE this repo (it will hold your financial data)
cd ~/retirement-plan && claude       # then: "Read CLAUDE.md and PLANNING.md and let's set up my plan"
```

`new` creates a plan folder with a copy of the example workbook to overwrite, starter categorization rules,
a `statements/` folder, a `PLANNING.md` log and a `CLAUDE.md` that tells Claude how to work on it. A synced
folder (Dropbox, iCloud Drive, OneDrive) lets you work from any machine.

## Commands

| | |
| --- | --- |
| `./plan new us\|ca <folder>` | Start a plan folder from the example household |
| `./plan statements <folder> [--update-workbook] [--replace-budget]` | Statements -> categorized ledger, 12-month budget, balances; optionally merged into the workbook |
| `./plan inputs <folder> [tab]` | List a scenario's inputs (keys, cells, values) |
| `./plan set <folder> "<tab>" key=value ...` | Change inputs (backup first; prints old -> new) |
| `./plan duplicate <folder> "<tab>" "<name>"` | Copy a scenario tab |
| `./plan whatif <folder> --set key=value [--vs --set ...] [--mc]` | Try changes without saving them; `--mc` adds a quick Monte Carlo |
| `./plan check <folder>` | Recalculate and compare every formula with the Python model |
| `./plan montecarlo <folder>` | Replay every scenario through history; writes the `monte_carlo` tab |
| `./plan demo us\|ca` | The end-to-end demo above |

The first run creates a Python environment in `~/.venvs/retirement-modeler` (just `openpyxl`). On Windows, use
WSL, or: `py -m venv .venv`, `.venv\Scripts\pip install openpyxl`, then `.venv\Scripts\python modeler\cli.py
<command> ...`.

## How it fits together

```
statements/ (CSV, OFX/QFX; PDFs via Claude)
   └─ pipeline.py ── rules/*.csv ──> ledger/ (transactions, categorized, needs_review, budget_summary, balances)
                                        └─ --update-workbook ──> plan.xlsx  budget + assets tabs
plan.xlsx  (the plan: Model tabs = scenarios, shared tabs = facts, tax_tables = the law)
   ├─ check_workbook.py: recalculates it (LibreOffice) and compares every cell with model.py
   ├─ whatif.py:          reruns a scenario with changes, in memory
   └─ monte_carlo.py:     reruns every scenario through history and lifetimes -> monte_carlo tab
PLANNING.md  (the log Claude keeps: decisions, every change with old values, open items)
```

The workbook is the source of truth. The Python model (`modeler/model.py`) is an independent copy of the Model
tab's formulas: `check` proves the two agree, which is what makes the Monte Carlo trustworthy, and it catches a
mistyped formula the moment it happens.

| Folder | |
| --- | --- |
| `modeler/` | The code: `model.py` (the model), `build_workbook.py` (writes a new workbook's formulas), `workbook.py` (reads/edits a workbook), `check_workbook.py`, `monte_carlo.py`, `whatif.py`, `pipeline.py` + `importers.py` + `update_workbook.py` (statements), `tax_tables.py` (law used to seed new workbooks), `cli.py` |
| `examples/us`, `examples/ca` | Fictional households: `inputs/` (what `new` starts from), `plan.xlsx`, `statements/` (synthetic exports in real bank formats), `rules/` |
| `templates/` | What `new` puts in a plan folder: `CLAUDE.md`, `PLANNING.md`, starter `rules/` |
| `data/` | Historical returns (Damodaran, NYU Stern, 1928-2025) and the SSA 2023 period life table |

## Data sources

- Returns and inflation: Aswath Damodaran, "Historical Returns on Stocks, Bonds and Bills" (NYU Stern, January
  2026): S&P 500 total return, 10-year US Treasury, 3-month T-bill, US CPI. Canadians can replace the columns
  with Canadian series (keep the headers).
- Mortality: US Social Security Administration period life table for 2023 (2026 Trustees Report).
- Tax law: 2026 IRS amounts (Rev. Proc. 2025-32, One Big Beautiful Bill Act); CRA 2026 indexation; BC Budget
  2026. Sources are noted on each workbook's `tax_tables` tab.

The sample statements are synthetic: the people, merchants' amounts and account numbers are made up; only the
file layouts follow real exports.
