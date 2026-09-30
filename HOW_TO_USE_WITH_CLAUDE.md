# How to build and run a retirement plan with Claude

This project grew out of one household's planning, done over a couple of weeks of conversations with Claude
Code. The tools matter less than the method. You bring documents, questions and decisions; Claude does the
bookkeeping, keeps the spreadsheet and its Python twin in step, runs the tests and writes everything down.
This guide describes that method.

## The idea in five lines

1. **The plan is a spreadsheet you can read.** Each scenario is a tab. Every number is a formula you can click
   on, so nothing is a black box, and a partner who never opens a terminal can still follow it.
2. **A Python copy of the same model checks the spreadsheet** cell by cell. The same copy then replays the plan
   through every stretch of market history and thousands of possible lifetimes.
3. **Statements, not guesses, set the budget.** Twelve months of real transactions, categorized, become the
   spending baseline; you then decide what changes in retirement.
4. **Every change is logged with its old value** in `PLANNING.md`. The log is how the next session (or the next
   machine) picks up where the last one stopped, and how you undo things.
5. **Scenarios change together.** "Sell the house in 2040" also means different housing costs afterward, selling
   costs, maybe a move and new taxes. Claude lists those knock-on assumptions and confirms them before it
   changes anything.

## Setup (once)

1. Install [Claude Code](https://docs.claude.com/en/docs/claude-code), Python 3.9+ and LibreOffice (see
   README).
2. Clone this repository and run `./plan demo us` (or `ca`) to see the whole thing work on a made-up household.
3. Create your plan folder **outside the repository**: `./plan new us ~/retirement-plan`. Put it in Dropbox,
   iCloud or OneDrive if you want to work from more than one computer. It will hold your statements and balances;
   treat it like your tax returns.
4. Open Claude Code in the plan folder: `cd ~/retirement-plan && claude`. The folder's `CLAUDE.md` tells Claude
   the rules (the workbook is the source of truth, log every change, check after every change, how to handle
   scenarios). Start each session with:

   > Read CLAUDE.md and PLANNING.md and let's continue.

## What to gather

You don't need all of this on day one; the plan improves as you add it.

| For | Documents |
| --- | --- |
| Spending | 12 months of statements for every account that pays for things: checking, every credit card, Venmo/PayPal. OFX/QFX or CSV downloads work best; PDFs are fine, Claude converts them |
| Balances | Each investment and retirement account's balance, and cost basis (US) / book value (Canada) for taxable accounts. A positions/holdings CSV download has both |
| Government pensions | US: your Social Security statement (ssa.gov/myaccount), for both of you. Canada: CPP Statement of Contributions (My Service Canada Account) and years lived in Canada after 18 (for OAS) |
| Work | Pay, bonus, stock vesting, retirement contributions and match (a pay stub or total-rewards statement). Any workplace pension estimate. Severance terms, if that's on the table |
| Taxes | Last year's return: shows your withholding, what you owed in April, rental income and depreciation |
| Housing | Home value (an estimate is fine), purchase price plus improvements, mortgage statement (balance, rate, payment), property tax bill, condo/HOA/strata fees, insurance |
| Other property | Rent, costs, depreciation / CCA schedule (from the tax return), purchase price |
| Kids | What you expect to pay per kid and per year: support, tuition, help after college. The 529/RESP balances |
| Healthcare | US: marketplace quotes for your ages and state (healthcare.gov); Medicare costs. Canada: a quote for retiree extended health and dental |

## Session by session

### 1. Replace the example household with yours

> Let's replace the example household with ours. Go tab by tab (household, assets, properties, healthcare,
> kids_budget, other_budget, then the Model tab's inputs). Ask me for what you need, tell me which documents
> would answer it, and log what you change.

Claude reads your documents, writes the values into the workbook, then runs `./plan check`. Good habits:
- Ask where each number came from, and have Claude put the source in the notes column.
- Where you only have an estimate, say so; Claude marks it as an estimate and lists it under open items.
- Give each kid their own lines on `kids_budget`: college years, a year or two of help afterward, a wedding.
  Each line has its own amount and years.

### 2. Turn statements into a budget

Drop the downloads into `statements/`, one folder per account (`statements/chase-sapphire/`,
`statements/checking/`...), then:

> Run the statement pipeline. Convert the PDFs first. Then walk me through needs_review.csv and add rules.

Claude runs `./plan statements .`, reads the per-file summary (money in and out for each file should look
right; if they're swapped, it adds a line to `rules/statement_formats.csv`), and goes through the uncategorized
payees with you, adding rules to `rules/payee_rules.csv`. Things to get right:
- **Transfers and card payments** must be categorized `Transfer`, or the same dollar counts twice (once
  leaving checking, once as the card's purchases). The starter rules catch the common ones; check the totals.
- **Kids' spending** (a kid's card, Venmo to a kid) goes to `Kids` with one subcategory per kid. It is modeled
  on the kids_budget tab, not in the household budget.
- **Housing, healthcare and taxes** are recognized but modeled elsewhere (properties, healthcare, the tax
  calculation), so they don't count twice.
- **Accounts with less than 12 months** of data are scaled up and flagged; better to download the full year.

When it looks right, merge it (the first time, replacing the example budget completely):

> Update the workbook with the budget and balances, replacing the example budget.

That runs `./plan statements . --update-workbook --replace-budget`. Then decide what changes in retirement
in the `retirement_budget` column: commuting and work clothes go down, travel may go up. Claude can suggest
changes, but the numbers are yours to choose.

### 3. Build the scenarios you're choosing between

> Make a scenario where we sell the house in 2040 and rent.

Claude answers first with the related assumptions, for example: rent afterward (how much?), selling costs,
the sale price (the formula's appreciation or your estimate?), moving costs as a one-off, renovations stopping,
and whether property tax and insurance stop. You confirm or correct them. Then it duplicates the tab
(`./plan duplicate`), changes the inputs (`./plan set`), runs `./plan check`, and logs the change with the old
values and the before/after results. Each Model tab shows its own results at the top; the Monte Carlo
(next section) compares all of them side by side.

Typical scenarios: retire now vs. in a year or two; one partner keeps working; claim Social Security/CPP at
62/67/70; sell or keep a rental; downsize or move somewhere with lower taxes; bigger or smaller budget; long-term
care early and long; a lost decade of returns (type returns into the yellow "Return override" column).

### 4. Stress-test

> Run the Monte Carlo. What do the failures have in common?

`./plan montecarlo .` replays every scenario through each starting year since 1928 and 1,000 random stitchings
of history, each with lifetimes drawn from the life table. How to read it:
- **"Money outlasts you both"** is the headline: the share of lifetime runs where you never run short while
  either of you is alive.
- **"Both alive to the end"** is the harsh version: you both live to the end of the table.
- **Investments left at the end**: the 10th percentile is your cushion in a bad but plausible world.
- Set your real **stock/bond/cash mix** and **extra years of life expectancy** on the `monte_carlo` tab first.
  Healthy, well-off people live longer than the population table, so more years is the cautious direction.
- Ask what the failing runs have in common. Usually it's a bad first decade combined with a long life and never
  cutting spending. That points to what to watch.

For close calls, compare options directly with exact failure counts:

> Is working one more year worth more than cutting spending 10%?

`./plan whatif . --set p1_retire=2031 --vs --set spend=0.9 --mc` answers that without changing the workbook.
For anything more elaborate, Claude writes a small read-only script in `analysis/` using the same model.

### 5. Keep it current

- **Each quarter or year**: download statements, rerun the pipeline, update balances. Ask Claude to compare
  actual spending with the plan.
- **Each January**: move `start_year` on the household tab forward and refresh the balances. Update the
  `tax_tables` tab for the new year's brackets (ask Claude to look them up and cite sources). Add last year's
  returns to `historical_returns`.
- **Life changes** (a layoff offer, an inheritance, a diagnosis, a move): tell Claude. It updates the
  affected inputs across every scenario, logs them, and reruns everything.
- Keep `PLANNING.md`'s "Where things stand" paragraph current: it's the first thing the next session reads.

## Good prompts

- "Before changing anything, list every assumption this touches and what you'd set it to."
- "Where did this number come from? Put the source in the notes."
- "Check the model: run ./plan check and tell me if anything doesn't match."
- "Which three inputs move the result most? Show me with whatif."
- "What would have to be true for us to run out of money?"
- "Write the case for and against retiring this year, for my partner, in plain language."
- "What's not modeled here that matters for us?" (Then read MODEL.md's list together.)
- "Log what we did today in PLANNING.md."

## Using Google Sheets instead of Excel

Import `plan.xlsx` into Google Sheets (File > Import > Upload) and it works as is: every formula is compatible.
The scripts read the local `.xlsx`, though. To run the checks or the Monte Carlo on a Google Sheet you've
edited, download it (File > Download > Microsoft Excel) over `plan.xlsx` in the plan folder first. Pick one
place to edit so the two don't drift apart. (Working directly against Google Sheets through its API is
possible and a natural next step, but it needs a Google Cloud OAuth client, so it isn't set up here.)

## Privacy

- Your plan folder holds account balances and every transaction you've downloaded. Keep it out of this
  repository (`.gitignore` blocks the obvious places, but don't rely on that) and out of public places.
- Claude reads these files to work on them. Use it under an account and data settings you're comfortable with.
- Redact account numbers from PDFs you don't need to keep; the pipeline only uses the last four digits.

## Limits worth knowing

The model is a planning tool, not tax software. The main simplifications are listed in [MODEL.md](MODEL.md):
couples only, no ACA premium subsidies, simplified state tax, rentals without mortgages, and in Canada no GIS
or probate. Treat results as ranges, not predictions. When a decision hinges on a detail the model
simplifies, Claude should say so, and you should check it with an accountant or planner.
