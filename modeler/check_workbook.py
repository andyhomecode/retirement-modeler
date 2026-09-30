#!/usr/bin/env python3
"""Check a plan workbook's formulas against model.py, cell by cell.

Recalculates the workbook (LibreOffice), reads every Model tab's inputs, reruns the scenario in Python and compares
every column the Python model computes, for every year, plus the results block. Any difference over $1 (or 0.01%)
is reported. Run it after editing formulas, after changing model.py, or when the Monte Carlo reports a MISMATCH.

Usage: check_workbook.py <plan.xlsx> [--verbose]
"""
import sys

import model as M
from workbook import column_letters, compare_headline, headline, read_plan


def check(path, verbose=False):
    plan = read_plan(path)
    if not plan["recalculated"]:
        print("note: LibreOffice not found - checking the values last saved in the file")
    ok = True
    for name, (S, data, table, results) in plan["models"].items():
        res = M.run(data, S)
        letters = column_letters(data["country"])
        bad, checked = [], 0
        for i, R in enumerate(res):
            for key, val in R.items():
                if key not in table or not isinstance(val, (int, float)) or isinstance(val, bool):
                    continue
                sheet = table[key][i]
                sheet = 0.0 if sheet in (None, "") else sheet
                if not isinstance(sheet, (int, float)):
                    bad.append((R["year"], key, letters[key], sheet, val))
                    continue
                checked += 1
                if abs(sheet - val) > max(1.0, 1e-4 * abs(val)):
                    bad.append((R["year"], key, letters[key], sheet, val))
        head = compare_headline(results, headline(res, data))
        if bad or head:
            ok = False
            print(f"{name}: {len(bad)} cell differences in {checked} checked" + (f"; results block: {'; '.join(head)}" if head else ""))
            seen = set()
            for y, key, col, sheet, val in bad:
                if verbose or key not in seen:
                    print(f"  {y} {key} (column {col}): sheet {sheet!r:>14}  model {val:14,.2f}")
                    seen.add(key)
        else:
            print(f"{name}: OK ({checked} cells over {len(res)} years match; results block matches)")
    return ok


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    sys.exit(0 if check(sys.argv[1], "--verbose" in sys.argv) else 1)


if __name__ == "__main__":
    main()
