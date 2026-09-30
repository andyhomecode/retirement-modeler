Put downloaded statements here, one subfolder per account (the folder name becomes the account name), e.g.

  statements/chase-sapphire/Chase4821_Activity20260101_20261231.CSV
  statements/checking/checking_2026.qfx
  statements/brokerage/Portfolio_Positions_Dec-31-2026.csv

Best formats, in order: OFX/QFX/QBO ("Quicken" or "Money" download), then CSV. Download 12 months (a calendar
year is easiest) for every account that pays for things: checking, every credit card, Venmo/PayPal if you use them.
For balances, download a positions/holdings CSV from each brokerage and retirement account.

PDF statements: ask Claude to read them and write CSVs in statements/converted/ with the columns
  date,account,description,amount,source_category,person,source_file
(date as YYYY-MM-DD, amount negative for money out). The pipeline skips a PDF once converted/ has a CSV of the
same name.

This folder is ignored by git. Don't commit statements anywhere public.
