# Retirement plan

This folder is a retirement plan built with retirement-modeler ({REPO}). Start every session by reading
`PLANNING.md` (where things stand, decisions, the change log, open items). The method is described in
`{REPO}/HOW_TO_USE_WITH_CLAUDE.md`; the model itself in `{REPO}/MODEL.md`.

## The pieces

- `plan.xlsx` - the plan, and the source of truth. Model tabs (`Model - ...`) are scenarios; their yellow cells
  are that scenario's inputs. Shared facts are on the other tabs (household, budget, kids_budget, other_budget,
  properties, healthcare, assets, tax_tables). `monte_carlo` holds the Monte Carlo settings and results.
- `statements/` - downloaded bank, card and investment statements, one subfolder per account. `rules/` - how
  the pipeline reads and categorizes them. `ledger/` - pipeline output (never edit by hand).
- `inputs/` - only used once, to create plan.xlsx. Don't rebuild from it.
- `PLANNING.md` - the plan's memory. Keep it current.

## Commands (run from anywhere)

```
{REPO}/plan inputs .                        # a Model tab's input keys, cells and values
{REPO}/plan set . "Model - X" key=value ... # change inputs (backs up first; prints old -> new)
{REPO}/plan duplicate . "Model - X" "Name"  # new scenario tab
{REPO}/plan whatif . --tab "Model - X" --set key=value [--vs --set ...] [--mc]   # try changes, nothing saved
{REPO}/plan check .                         # formulas vs the Python model, cell by cell
{REPO}/plan montecarlo .                    # history + lifetimes; writes the monte_carlo tab
{REPO}/plan statements . [--update-workbook] [--replace-budget]
```

## Rules

- **The workbook is the source of truth.** Change it in place - with `plan set`, or with openpyxl for anything
  else (load without `data_only`, change cells, save; formulas recalculate when opened). Never regenerate it.
  The person may have it open in Excel/LibreOffice: the scripts refuse to write while a lock file exists; ask
  them to close it. Every script that writes makes a backup in `backups/` first.
- **Scenario changes** ("assume we sell the house in 2040 and rent", "what if I retire at 60"): first list the
  related assumptions you'd also change - costs that stop or start, sale price and selling costs, housing
  afterwards, taxes when moving, healthcare when coverage changes, one-off costs - and confirm them with the
  person. Then make all the changes, run `plan check .`, and record in PLANNING.md (dated) what changed, with the
  old values, and the headline results before and after.
- To explore without changing the plan, use `plan whatif`, or write a read-only script in `analysis/` that
  uses `workbook.read_plan()` and `model.run()` from `{REPO}/modeler` (see whatif.py for the pattern).
- After any change to formulas: `plan check .` must say OK. A MISMATCH in the Monte Carlo's "Check vs Model
  tab" column means the workbook has something the Python model doesn't read (e.g. a hand-edited formula in the
  year table) - explain it to the person rather than ignoring it. Change inputs, not the year table's formulas;
  the yellow "Return override" and "Household spending multiplier" columns cover most per-year needs.
- Statements: run the pipeline, then go through `ledger/needs_review.csv` with the person and add rules to
  `rules/payee_rules.csv` (specific rules - their kids, their employer - at the top). Transfers between their
  own accounts and card payments must end up as `Transfer`. Convert PDF statements yourself: read them and write
  `statements/converted/<same name>.csv` with columns date,account,description,amount,source_category,person,
  source_file (amount negative = money out); spot-check totals against the statement.
- Be explicit about uncertainty: the model simplifies tax (see MODEL.md, "Not modeled"). Flag decisions where
  that matters and suggest checking with an accountant or planner. Don't present results as advice.
- Privacy: this folder holds personal financial data. Never copy it into the retirement-modeler repo, a
  GitHub issue, or anywhere public.
