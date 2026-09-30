# Retirement plan: log

This file is the plan's memory. Claude reads it at the start of a session and adds to it at the end, so work
carries over between sessions and machines. Keep it in the plan folder, next to plan.xlsx.

## Where things stand

(One paragraph: which scenario is the working plan, its headline results, what's being worked on.)

## Decisions and conclusions

(Dated, most recent first. What was decided and why, with the numbers that supported it.)

## Changes to the workbook

(Dated, most recent first. Every input changed, with old -> new values, so any change can be undone.)

- YYYY-MM-DD: created the plan from the US/CA example (`./plan new`).

## Open items

- [ ] Replace the example inputs with ours: household, assets, properties, healthcare, kids_budget, other_budget.
- [ ] Download 12 months of statements into statements/ and run `./plan statements <folder> --update-workbook --replace-budget`.
- [ ] Set the stock/bond/cash mix on the monte_carlo tab.
