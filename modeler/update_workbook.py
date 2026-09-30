"""Merge pipeline output into an existing plan workbook (never rebuilds it).

budget tab:  each category/subcategory's annual_actual, source_period and n_transactions are replaced; your
             retirement_budget column is kept. New categories are added at the bottom with retirement_budget =
             the actual (for household lines). Lines with no spending this time are left alone and listed.
assets tab:  accounts are matched by name (rules/account_map.csv sets the names and buckets). value and as_of (and
             cost_basis when the export has one) are updated. A balance for an account that isn't on the tab is added
             only if account_map.csv gives it a bucket (so credit-card balances never become assets).
A backup of the workbook goes to backups/ first. Close the workbook in Excel/LibreOffice before running.
"""
import shutil
import sys
import time
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill

INPUT = PatternFill("solid", fgColor="FFF2CC")
NEW = PatternFill("solid", fgColor="FCE4D6")


def backup(path):
    p = Path(path)
    for lock in (p.with_name(f".~lock.{p.name}#"), p.with_name(f"~${p.name}")):
        if lock.exists():
            sys.exit(f"{p.name} looks open in Excel/LibreOffice ({lock.name}); close it and rerun")
    d = p.parent / "backups"
    d.mkdir(exist_ok=True)
    dest = d / f"{p.stem} {time.strftime('%Y-%m-%d %H%M%S')}.xlsx"
    shutil.copy2(p, dest)
    return dest


def update_budget(ws, rows, replace=False):
    head = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(head) if h}
    if replace:                                  # start the tab over from these rows
        for r in range(2, ws.max_row + 1):
            for k in col.values():
                ws.cell(r, k).value = None
                ws.cell(r, k).fill = PatternFill()
    existing, last = {}, 1
    for r in range(2, ws.max_row + 1):
        cat, sub = ws.cell(r, col["category"]).value, ws.cell(r, col["subcategory"]).value
        if cat:
            existing[(str(cat), str(sub or ""))] = r
            last = r
    changed, added, seen = 0, [], set()
    for row in rows:
        key = (row["category"], row["subcategory"])
        seen.add(key)
        r = existing.get(key)
        if r is None:
            last += 1
            r = last
            added.append(f"{row['category']} / {row['subcategory']} ${row['annual_actual']:,.0f}")
            for k in ("model_line", "expense_group", "category", "subcategory"):
                ws.cell(r, col[k], row[k])
            if row["model_line"] == "household":
                ws.cell(r, col["retirement_budget"], row["annual_actual"]).number_format = "#,##0"
            for k in col.values():
                ws.cell(r, k).fill = NEW
            ws.cell(r, col["model_line"]).fill = ws.cell(r, col["retirement_budget"]).fill = INPUT
        else:
            changed += 1
        ws.cell(r, col["annual_actual"], row["annual_actual"]).number_format = "#,##0"
        ws.cell(r, col["source_period"], row["source_period"])
        ws.cell(r, col["n_transactions"], row["n_transactions"])
        if row.get("notes"):
            ws.cell(r, col["notes"], row["notes"])
    untouched = [f"{c} / {s}" for (c, s) in existing if (c, s) not in seen]
    return changed, added, untouched


def update_assets(ws, balances):
    existing, last = {}, 1
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if name:
            existing[str(name).strip().lower()] = r
            last = r
    changed, added = [], []
    for b in balances:
        r = existing.get(b["account"].strip().lower())
        if r is None and not b["bucket"]:
            continue                             # unmapped (e.g. a credit card's balance): not an asset
        if r is None:
            last += 1
            r = last
            ws.cell(r, 1, b["account"]).fill = INPUT
            ws.cell(r, 2, b["bucket"] or None).fill = INPUT
            ws.cell(r, 6, "added by the statement pipeline")
            added.append(b["account"])
        else:
            changed.append(b["account"])
        ws.cell(r, 3, round(b["value"], 2)).number_format = "#,##0"
        ws.cell(r, 4, b["as_of"] or None)
        if b.get("cost_basis") is not None:
            ws.cell(r, 5, round(b["cost_basis"], 2)).number_format = "#,##0"
    return changed, added


def update(path, budget_rows, balances, replace_budget=False):
    path = Path(path)
    if not path.exists():
        sys.exit(f"{path} not found")
    dest = backup(path)
    wb = openpyxl.load_workbook(path)
    changed, added, untouched = update_budget(wb["budget"], budget_rows, replace_budget)
    print(f"5. {path.name}: budget tab - {changed} lines updated, {len(added)} added" + (": " + "; ".join(added) if added else ""))
    if untouched:
        print(f"     no spending this period (left as they were): {'; '.join(untouched)}")
    ch, ad = update_assets(wb["assets"], balances)
    print(f"   assets tab - {len(ch)} accounts updated, {len(ad)} added" + (": " + ", ".join(ad) if ad else ""))
    wb.save(path)
    print(f"   backup: {dest}")
