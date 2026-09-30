#!/usr/bin/env python3
"""Try changes to a scenario without touching the workbook, and compare with it as it stands.

    whatif.py <plan.xlsx> --tab "Model - Example" --set p1_retire=2031 --set spend=0.9 [--mc]

--set key=value changes a Model-tab input (the keys `./plan inputs` lists), with two extras:
    spend=0.9        household spending multiplier for every year (spend=0.9@2030 = from 2030 on)
    return=0.04      flat return for every year (same as return_default)
Several --set flags make one variant; separate variants with --vs:
    whatif.py plan.xlsx --set p1_retire=2031 --vs --set spend=0.9
--mc adds a quick Monte Carlo (every historical start with lifetimes drawn as on the monte_carlo tab) for each
variant, with exact failure counts - enough to rank close alternatives. Nothing is written anywhere.
"""
import argparse
import random
import statistics
import sys

import model as M
from monte_carlo import adjusted_survival, death_year, history_rows, life_table, paths, read_settings
from workbook import headline, parse_value, read_plan


def apply(S, data, sets):
    S = dict(S)
    for key, raw in sets:
        if key == "spend":
            v, _, frm = raw.partition("@")
            frm = int(frm) if frm else data["start"]
            S["spend_mult"] = dict(S.get("spend_mult") or {})
            for y in range(frm, data["end"] + 1):
                S["spend_mult"][y] = S["spend_mult"].get(y, 1.0) * parse_value(v)
        elif key == "return":
            S["return_default"] = parse_value(raw)
            S["returns"] = {}
        elif key.startswith("prop") and "_" in key:
            i, field = int(key[4:key.index("_")]) - 1, key[key.index("_") + 1:]
            S["props"] = [dict(p) for p in S["props"]]
            S["props"][i][field] = parse_value(raw)
        elif key in S:
            S[key] = parse_value(raw)
        else:
            sys.exit(f"unknown input {key!r}")
    return S


def quick_mc(plan, S, data):
    wb = plan["wb"]
    st = read_settings(wb["monte_carlo"])
    hist = history_rows(wb["historical_returns"])
    years = list(range(data["start"], data["end"] + 1))
    q = life_table(wb["life_table"])
    people = data["people"]
    surv = [adjusted_survival({a: v[people[i]["sex"]] for a, v in q.items()}, data["start"] - people[i]["born"], st["extra"][i])
            for i in (0, 1)]
    rng = random.Random(st["seed"])
    flat = {k: v for k, v in S.items() if k != "returns"}
    fails, ends = 0, []
    runs = [p for p in paths(hist, years, st["mix"], 0, st["block"], st["seed"])]
    for _, _, ret, infl in runs:
        for _ in range(st["draws"]):
            d1 = death_year(surv[0], people[0]["born"], rng, data["end"])
            d2 = death_year(surv[1], people[1]["born"], rng, data["end"])
            res = M.run(data, flat, returns=ret, inflation_path=infl, p1_death_year=d1, p2_death_year=d2,
                        survivor_spending=st["survivor"])
            fails += any(r["shortfall"] > 1 for r in res)
            ends.append(res[-1]["investments_today"])
    ends.sort()
    return fails, len(ends), ends[len(ends) // 20], statistics.median(ends)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--tab", default=None, help="Model tab (default: the first one)")
    ap.add_argument("--mc", action="store_true")
    ap.add_argument("rest", nargs=argparse.REMAINDER)
    a = ap.parse_args(argv)
    variants, cur = [], []
    toks = a.rest
    i = 0
    while i < len(toks):
        if toks[i] == "--vs":
            variants.append(cur)
            cur = []
        elif toks[i] == "--set":
            k, _, v = toks[i + 1].partition("=")
            cur.append((k.strip(), v))
            i += 1
        elif toks[i] == "--mc":
            a.mc = True
        i += 1
    variants.append(cur)
    plan = read_plan(a.workbook)
    tab = a.tab or next(iter(plan["models"]))
    S0, data, _, _ = plan["models"][tab]
    born = data["people"][0]["born"]
    name1 = data["people"][0]["name"]
    rows = [("as is", S0)] + [(" + ".join(f"{k}={v}" for k, v in sets) or "as is", apply(S0, data, sets)) for sets in variants if sets]
    print(f"{tab}, flat rates; today's dollars")
    base = None
    for label, S in rows:
        res = M.run(data, S)
        h = headline(res, data)
        by = {r["year"]: r for r in res}
        inv = {age: by[born + age]["investments_today"] for age in (70, 80, 90) if born + age in by}
        line = (f"  {label:<44} investments at {name1} 70/80/90: " + " / ".join(f"${v / 1e6:,.2f}M" for v in inv.values())
                + f"   first shortfall: {h['short']}   lifetime tax: ${h['tax'] / 1e3:,.0f}k")
        if base:
            line += "   (at 90: " + f"{inv.get(90, 0) - base.get(90, 0):+,.0f})"
        print(line)
        base = base or inv
        if a.mc:
            f, n, p5, med = quick_mc(plan, S, data)
            print(f"  {'':<44} Monte Carlo: {f} of {n} lifetimes run short; investments left at the end: 5th percentile "
                  f"${p5 / 1e6:,.2f}M, median ${med / 1e6:,.2f}M")


if __name__ == "__main__":
    main()
