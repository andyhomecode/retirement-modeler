# What the model does

One row per year from `start_year` until the younger person is 100. All inputs are in today's dollars (the
prices of the year before `start_year`) and grow with their own inflation rate: general inflation, healthcare,
college, long-term care, each property's appreciation. Tax brackets are for `TAX_YEAR` (2026) and are indexed
where the law indexes them.

## Each year

1. **Work.** Each person is paid until their retirement year, minus retirement contributions (401(k)/403(b) or
   RRSP/pension) and payroll tax (Social Security + Medicare, or CPP/CPP2 + EI). A flat withholding rate is
   taken from pay; the difference from the real tax is settled the next April.
2. **Costs.** The household budget (`budget` tab, `household` lines, times that year's spending multiplier),
   housing (mortgage from its balance, rate and payment; carrying costs; after selling, the new housing cost),
   kids (`kids_budget`: each line has its own years; college lines are paid from 529s/RESPs first), other
   costs (`other_budget`), healthcare, long-term care and one-off expenses.
3. **Income.** Take-home pay, government pensions, workplace pensions, property income net of cash costs,
   dividends, required withdrawals, sale proceeds, one-off income.
4. **The gap** is covered from the accounts in this order:
   - US: taxable, then pre-tax (split between the two of you by balance), then Roth, then HSA.
   - Canada: non-registered, then RRSP/RRIF, then TFSA.
   A surplus is reinvested: in the US into the taxable account; in Canada into TFSAs up to the available room,
   then non-registered.
5. **Tax** is computed on what actually happened that year and paid the following April.
6. **Investments** grow at the year's return (a flat rate, the per-year override, or history in the Monte
   Carlo). Taxable accounts track cost basis, so sales realize the right share of gains.

## United States

- Federal, married filing jointly: ordinary brackets; qualified dividends and capital gains stacked on top at
  0/15/20%; 3.8% net investment income tax; taxable Social Security (the 50%/85% formula); standard deduction
  plus the 65+ addition, and the 2025-2028 senior deduction ($6,000 per person 65+, phased out above $150,000);
  itemized deductions (mortgage interest, state and property tax up to the SALT cap with its phase-down,
  medical costs above 7.5% of AGI, charity), whichever is larger; 25% on depreciation recapture.
- State: one set of brackets (the `tax_tables` tab; the examples use New York), a standard deduction, optional
  indexing, optional Social Security exemption, a per-person exclusion for retirement-account and pension
  income, and a flat local tax. Moving: from the move year, a flat rate on income except Social Security (0%
  for no-income-tax states).
- Social Security: each person's benefit at full retirement age (67) from their SSA statement; reduced for
  early claims, 8%/yr more for delayed ones; a spouse gets the larger of their own and half the other's (once
  both claim); an optional reduction for zero-earning years and an across-the-board cut from a chosen year.
- Healthcare: an employer plan while either works; then unsubsidized marketplace premiums and out-of-pocket for
  adults under 65 and kids still on the plan; Medicare B + D + Medigap and out-of-pocket from 65; IRMAA from
  income two years earlier; the HSA pays eligible costs.
- Required minimum distributions from 73 or 75 (by birth year) with the Uniform Lifetime Table.
- Roth conversions: an amount per year between two years, taxed as income.
- Home sale: $500,000 of gain excluded.

## Canada (British Columbia)

- Each spouse is taxed separately: federal and provincial brackets; basic personal amount (the federal one
  shrinks at high income); age amount (reduced above the threshold); pension amount; Canada employment amount;
  CPP/EI credits; medical expense credit (claimed by the lower-income spouse, including long-term care);
  eligible dividends grossed up with the dividend tax credit; half of capital gains taxable; the BC tax
  reduction. BC's pause on indexing (2027-2030) is built in.
- Pension income splitting: workplace pension income (any age) and RRIF income (65+) are split to even out your
  net incomes, up to half.
- CPP: each person's amount at 65 from their statement, -0.6%/month earlier, +0.7%/month later. OAS: the full
  pension prorated by years in Canada after 18, +0.6%/month deferred to 70, +10% from 75, and the 15% clawback
  above the threshold (on each person's net income).
- RRSPs become RRIFs: minimum withdrawals from the year you turn 72. An optional extra RRSP/RRIF withdrawal per
  year (an "RRSP meltdown") between two years.
- TFSA: new room each January for each spouse; withdrawals give the room back the next year.
- Healthcare: the provincial plan covers the basics; the model adds a private extended health and dental plan
  and out-of-pocket costs after retirement, plus long-term care.
- Home: principal residence, so no tax on the sale. Rentals: capital cost allowance, and recapture on sale.

## Properties

Row 1 of the `properties` tab is your home; rows 2-5 are other properties. Each Model tab decides, per
property: sell in a year (price = value grown at its appreciation, or a price you type), change use from a year
(new rent and costs, e.g. rent out the cottage, or move in), and, for the home, housing costs after selling.
Rental income is taxed after depreciation/CCA; losses aren't deducted.

## Monte Carlo

Each run replaces the flat return and inflation with a sequence from history (your stock/bond/cash mix of each
year's S&P 500, 10-year Treasury and T-bill returns; that year's CPI). Healthcare, college, long-term care and
property values keep their spread over inflation. Starting points are every year from 1928, wrapping around,
plus random runs built from 10-year blocks.

Lifetimes come from the SSA life table, with each person's death rates scaled so their life expectancy is the
table's plus the extra years you set. After the first death, household and other spending drop to the
survivor share. The survivor keeps the larger Social Security benefit (or their own CPP plus a survivor
pension), plus the survivor share of any workplace pension. The deceased's pre-tax accounts roll over, and
their healthcare stops. Long-term care happens in each person's final years. A run fails if the accounts run
out while either of you is alive.

## Not modeled (the main simplifications)

- Couples only; after a death, taxes stay joint (US) and the survivor's tax isn't recomputed as a single filer.
- US: ACA premium tax credits (the model assumes full-price marketplace premiums, which is cautious for low-income
  early-retirement years); state itemized deductions and state-specific credits; progressive local taxes (flat
  here); 10-year-rule inherited IRAs; Social Security earnings test; estate tax.
- Canada: GIS and other income-tested benefits; the enhanced CPP contributions' deduction (treated as a credit);
  the non-eligible dividend credit; provincial health premiums and PharmaCare deductibles; tax at death (the
  deemed sale of RRIFs and capital property) and probate.
- Other properties are mortgage-free (put loan payments in their cash costs); property sales are taxed in the
  year of sale with no installment or replacement relief.
- Rental and business losses aren't deducted; there's no carry-forward.
- Investment returns are the same for every account (no asset location), and fees come out of your chosen
  return.
- Buying a new home: its price can go in as a one-off expense, but its value isn't added to net worth.
- Historical returns are US series. Canadians can substitute Canadian ones (keep the column headers).
