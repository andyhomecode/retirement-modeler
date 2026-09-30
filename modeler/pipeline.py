#!/usr/bin/env python3
"""Turn downloaded statements into the plan's budget and account balances.

    pipeline.py <plan_folder> [--months 12 | --year 2026] [--update-workbook]

Reads everything under <plan_folder>/statements/ (put each account's exports in its own subfolder; the folder name
becomes the account name), then:
  1. imports every transaction (OFX/QFX and CSV; PDFs are listed for conversion) -> ledger/transactions.csv
  2. categorizes them with rules/payee_rules.csv, then the bank's own categories via rules/source_categories.csv
     -> ledger/categorized.csv, and ledger/needs_review.csv for whatever is left (biggest dollar amounts first)
  3. totals 12 months of spending by category -> ledger/budget_summary.csv (accounts with less than 12 months of
     data are scaled up to a year, and flagged)
  4. reads balances (bank ledger balances, investment positions) -> ledger/balances.csv
With --update-workbook it then merges steps 3 and 4 into <plan_folder>/plan.xlsx: the budget tab's actual column
and the assets tab's values (after saving a backup). Your retirement_budget column and buckets are kept.

Money in the ledger: negative = out. Card payments and transfers between your own accounts must be categorized as
Transfer (the starter rules catch the common ones) or they'd count twice.
"""
import argparse
import csv
import fnmatch
import re
import sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

import importers as IM

EXCLUDED_LINES = {"transfer", "income", "property"}   # not spending the budget should count


def read_rules(path):
    if not Path(path).exists():
        return []
    with open(path, newline="", encoding="utf-8-sig") as f:
        return [r for r in csv.DictReader(f) if any((v or "").strip() for v in r.values())
                and not (list(r.values())[0] or "").startswith("#")]


def write_csv(path, fields, rows):
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


# ---------------------------------------------------------------------------------------------- 1. import

def import_all(plan):
    stmts = plan / "statements"
    formats = read_rules(plan / "rules" / "statement_formats.csv")
    txns, balances, report, pdfs = [], [], [], []
    converted = {p.stem.lower() for p in (stmts / "converted").glob("*.csv")} if (stmts / "converted").exists() else set()
    for path in sorted(stmts.rglob("*")):
        if not path.is_file() or path.name.startswith("."):
            continue
        rel = path.relative_to(stmts)
        folder = rel.parts[0] if len(rel.parts) > 1 and rel.parts[0] != "converted" else None
        ext = path.suffix.lower()
        override = next((f for f in formats if fnmatch.fnmatch(str(rel).lower(), (f.get("file_pattern") or "").lower())), None)
        label = (override or {}).get("account") or folder
        try:
            if ext in (".ofx", ".qfx", ".qbo"):
                rows = IM.ofx_transactions(path)
                if label:
                    for r in rows:
                        r["account"] = label
                txns += rows
                for acct, v, b, asof in IM.ofx_balances(path):
                    balances.append(dict(source=str(rel), account=label or acct, value=v, cost_basis=b, as_of=asof))
                report.append((str(rel), "OFX", rows))
            elif ext == ".csv":
                if IM.is_positions_csv(path):
                    for acct, v, b, asof in IM.positions_balances(path):
                        balances.append(dict(source=str(rel), account=acct or label or path.stem, value=v, cost_basis=b, as_of=asof))
                    report.append((str(rel), "positions", []))
                    continue
                ov = None
                if override:
                    headerless = (override.get("headerless") or "").strip().lower() in ("1", "yes", "true", "y")
                    conv = (lambda v: int(v) if v not in (None, "") else None) if headerless else (lambda v: v or None)
                    ov = dict(date=conv(override.get("date_col")), desc=conv(override.get("description_col")),
                              amount=conv(override.get("amount_col")), debit=conv(override.get("debit_col")),
                              credit=conv(override.get("credit_col")), headerless=headerless,
                              sign=override.get("sign") or None, date_format=override.get("date_format") or None)
                rows, prof, bal = IM.csv_transactions(path, label, ov)
                txns += rows
                if bal and rows:
                    balances.append(dict(source=str(rel), account=rows[-1]["account"], value=bal[1], cost_basis=None, as_of=bal[0]))
                report.append((str(rel), prof, rows))
            elif ext == ".pdf":
                if path.stem.lower() not in converted:
                    pdfs.append(str(rel))
        except Exception as e:  # noqa: BLE001 - report and keep going
            report.append((str(rel), f"ERROR: {e}", []))
    # the same transaction in two overlapping exports counts once
    seen, deduped = defaultdict(lambda: defaultdict(int)), []
    for r in txns:
        seen[(r["account"], r["date"], r["description"], r["amount"])][r["source_file"]] += 1
    kept = defaultdict(int)
    for r in txns:
        key = (r["account"], r["date"], r["description"], r["amount"])
        if kept[key] < max(seen[key].values()):
            kept[key] += 1
            deduped.append(r)
    return deduped, balances, report, pdfs, len(txns) - len(deduped)


# ---------------------------------------------------------------------------------------------- 2. categorize

def normalize_payee(desc):
    d = re.sub(r"\b\d{4,}\b|#\d+|\*\w+", "", desc)
    d = re.sub(r"\s+", " ", d).strip().upper()
    return d[:40]


def categorize(txns, plan):
    rules = read_rules(plan / "rules" / "payee_rules.csv")
    compiled = [(re.compile(r["pattern"], re.I), r["category"], r.get("subcategory") or "General",
                 re.compile(r["account"], re.I) if r.get("account") else None,
                 re.compile(r["person"], re.I) if r.get("person") else None) for r in rules]
    src = [(re.compile(r["source_category"], re.I), r["category"], r.get("subcategory") or "General")
           for r in read_rules(plan / "rules" / "source_categories.csv")]
    review = defaultdict(lambda: dict(total=0.0, n=0, accounts=set(), sample=""))
    for t in txns:
        t["category"] = t["subcategory"] = t["matched_by"] = ""
        for pat, cat, sub, acct, person in compiled:
            if pat.search(t["description"]) and (not acct or acct.search(t["account"])) and (not person or person.search(t.get("person", ""))):
                t["category"], t["subcategory"], t["matched_by"] = cat, sub, "payee_rules"
                break
        if not t["category"] and t.get("source_category"):
            for pat, cat, sub in src:
                if pat.search(t["source_category"]):
                    t["category"], t["subcategory"], t["matched_by"] = cat, sub, "source_categories"
                    break
        if not t["category"]:
            g = review[normalize_payee(t["description"])]
            g["total"] += t["amount"]
            g["n"] += 1
            g["accounts"].add(t["account"])
            g["sample"] = t["description"]
    return sorted(review.items(), key=lambda kv: -abs(kv[1]["total"]))


# ---------------------------------------------------------------------------------------------- 3. budget

def month_span(d0, d1):
    return (d1 - d0).days / 30.4375


def summarize(txns, plan, months=12, year=None):
    cats = {r["category"]: r for r in read_rules(plan / "rules" / "categories.csv")}
    if year:
        lo, hi = date(year, 1, 1), date(year, 12, 31)
    else:
        hi = max(date.fromisoformat(t["date"]) for t in txns)
        if (hi + timedelta(days=7)).month != hi.month:      # data runs to (nearly) the end of a month: whole months
            hi = (hi.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
            m = hi.year * 12 + hi.month - months
            lo = date(m // 12, m % 12 + 1, 1)
        else:
            lo = hi - timedelta(days=round(months * 30.4375) - 1)
    window = month_span(lo, hi + timedelta(days=1))
    cover = {}
    for t in txns:
        d = date.fromisoformat(t["date"])
        a = cover.setdefault(t["account"], [d, d])
        a[0], a[1] = min(a[0], d), max(a[1], d)
    scale, coverage = {}, []
    for acct, (d0, d1) in cover.items():
        c0, c1 = max(d0, lo), min(d1, hi)
        m = max(0.0, month_span(c0, c1 + timedelta(days=1))) if c1 >= c0 else 0.0
        scale[acct] = (window / m) if 0 < m < window - 0.75 else 1.0
        coverage.append((acct, d0, d1, m))
    tot, n, scaled = defaultdict(float), defaultdict(int), defaultdict(set)
    for t in txns:
        d = date.fromisoformat(t["date"])
        if not (lo <= d <= hi) or not t["category"]:
            continue
        line = (cats.get(t["category"]) or {}).get("model_line", "household")
        if line in EXCLUDED_LINES:
            continue
        key = (t["category"], t["subcategory"])
        tot[key] += -t["amount"] * scale[t["account"]]
        n[key] += 1
        if scale[t["account"]] != 1.0:
            scaled[key].add(t["account"])
    rows = []
    for (cat, sub), amt in tot.items():
        c = cats.get(cat) or {}
        rows.append(dict(model_line=c.get("model_line", "household"), expense_group=c.get("expense_group", ""), category=cat,
                         subcategory=sub, annual_actual=round(amt, 2), source_period=f"{lo:%Y-%m}..{hi:%Y-%m}",
                         n_transactions=n[(cat, sub)],
                         notes=("annualized: " + ", ".join(sorted(scaled[(cat, sub)])) + " cover less than the period") if scaled[(cat, sub)] else ""))
    order = {"household": 0, "housing": 1, "healthcare": 2, "kids": 3, "taxes": 4}
    rows.sort(key=lambda r: (order.get(r["model_line"], 9), r["expense_group"], -r["annual_actual"]))
    return rows, (lo, hi), coverage


# ---------------------------------------------------------------------------------------------- 4. balances

def map_balances(balances, plan):
    amap = read_rules(plan / "rules" / "account_map.csv")
    out = []
    for b in balances:
        hit = next((m for m in amap if re.search(m["match"], b["account"], re.I) or re.search(m["match"], b["source"], re.I)), None)
        out.append(dict(b, account=(hit or {}).get("account_name") or b["account"], bucket=(hit or {}).get("bucket", ""),
                        as_of=b["as_of"].isoformat() if b["as_of"] else ""))
    return out


# ---------------------------------------------------------------------------------------------- main

def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("plan_folder")
    ap.add_argument("--months", type=int, default=12, help="budget = this many months ending at the latest transaction")
    ap.add_argument("--year", type=int, help="budget = this calendar year instead")
    ap.add_argument("--update-workbook", action="store_true", help="merge the budget and balances into plan.xlsx")
    ap.add_argument("--replace-budget", action="store_true",
                    help="with --update-workbook: replace every line of the budget tab (first import into a new plan)")
    a = ap.parse_args(argv)
    plan = Path(a.plan_folder)
    if not (plan / "statements").exists():
        sys.exit(f"no statements/ folder in {plan}")
    ledger = plan / "ledger"
    ledger.mkdir(exist_ok=True)

    txns, balances, report, pdfs, dupes = import_all(plan)
    print("1. Imported")
    for rel, prof, rows in report:
        if prof == "positions":
            print(f"   {rel}: holdings (balances only)")
            continue
        out = sum(-r["amount"] for r in rows if r["amount"] < 0)
        inn = sum(r["amount"] for r in rows if r["amount"] > 0)
        span = f"{min(r['date'] for r in rows)}..{max(r['date'] for r in rows)}" if rows else ""
        print(f"   {rel}: {len(rows)} transactions {span}  out ${out:,.0f}  in ${inn:,.0f}  [{prof}]")
    if dupes:
        print(f"   {dupes} duplicate transactions from overlapping exports dropped")
    if pdfs:
        print("   PDFs not converted yet (ask Claude to write them as CSVs in statements/converted/):")
        for p in pdfs:
            print(f"     {p}")
    if not txns:
        sys.exit("no transactions found")
    write_csv(ledger / "transactions.csv", IM.STANDARD_FIELDS, txns)

    review = categorize(txns, plan)
    write_csv(ledger / "categorized.csv", IM.STANDARD_FIELDS + ["category", "subcategory", "matched_by"], txns)
    write_csv(ledger / "needs_review.csv", ["payee", "total", "count", "accounts", "sample_description"],
              [dict(payee=k, total=round(g["total"], 2), count=g["n"], accounts="|".join(sorted(g["accounts"])),
                    sample_description=g["sample"]) for k, g in review])
    absolute = sum(abs(t["amount"]) for t in txns)
    unc = sum(abs(t["amount"]) for t in txns if not t["category"])
    print(f"2. Categorized: {1 - unc / absolute:.1%} of dollars; {sum(g['n'] for _, g in review)} transactions in "
          f"{len(review)} payee groups left -> ledger/needs_review.csv")
    for k, g in review[:8]:
        print(f"     {g['total']:>12,.2f}  {g['n']:>3}x  {k}")

    rows, (lo, hi), coverage = summarize(txns, plan, a.months, a.year)
    fields = ["model_line", "expense_group", "category", "subcategory", "annual_actual", "source_period", "n_transactions", "notes"]
    write_csv(ledger / "budget_summary.csv", fields, rows)
    print(f"3. Budget {lo}..{hi} -> ledger/budget_summary.csv")
    for acct, d0, d1, m in sorted(coverage):
        flag = "" if m >= (month_span(lo, hi + timedelta(days=1)) - 0.75) else "  <- partial: scaled up to the full period"
        print(f"     {acct:<28} {d0}..{d1}  {m:4.1f} months in the period{flag}")
    by_line = defaultdict(float)
    for r in rows:
        by_line[r["model_line"]] += r["annual_actual"]
    print("     " + "   ".join(f"{k}: ${v:,.0f}" for k, v in sorted(by_line.items())))
    kids = [r for r in rows if r["model_line"] == "kids"]
    if kids:
        print("     kids (for the kids_budget tab): " + "; ".join(f"{r['subcategory']} ${r['annual_actual']:,.0f}" for r in kids))

    bal = map_balances(balances, plan)
    write_csv(ledger / "balances.csv", ["account", "bucket", "value", "cost_basis", "as_of", "source"], bal)
    print(f"4. Balances -> ledger/balances.csv")
    for b in bal:
        print(f"     {b['account']:<34} ${b['value']:>12,.0f}  {b['bucket'] or '(not mapped - ignored; map it in rules/account_map.csv if it is an asset)'}")

    if a.update_workbook:
        import update_workbook as UW
        UW.update(plan / "plan.xlsx", rows, bal, a.replace_budget)


if __name__ == "__main__":
    main()
