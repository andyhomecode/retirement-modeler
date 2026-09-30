"""The retirement model in Python: an independent copy of the Model tab's formulas.

The workbook's Model tabs are the plan; this file computes the same thing so that
  - check_workbook.py can compare the two cell by cell (a formula typo shows up as a mismatch), and
  - monte_carlo.py can rerun a scenario thousands of times with historical returns and random lifetimes.
Keep it in step with build_workbook.py: change one, change the other, then run check_workbook.py.

run(data, scenario) takes what workbook.read_plan() returns: `data` = the shared tabs (household, budgets,
properties, healthcare, assets, tax tables), `scenario` = one Model tab's yellow inputs. Keyword overrides replace
scenario inputs (the Monte Carlo passes returns, inflation paths and death years this way).

Both countries share the structure; they differ in pay deductions, healthcare, government pensions, how accounts
are drawn down and taxed. Account names are generic:
  taxable  US taxable brokerage + cash        | CA non-registered + cash
  pre1/2   US 401(k)/IRA (pre-tax), per person | CA RRSP/RRIF, per person
  free     US Roth IRAs                         | CA TFSAs
  hsa      US HSA                               | (none)
  college  US 529s                              | CA RESPs (kids' money; not counted in your investments)
"""
import math

from tax_tables import rmd_start_age

ONEOFF_TYPES = ["Wages: person 1", "Wages: person 2", "Other taxable income", "Tax-free money in", "Expense"]


def prog(x, brackets, idx=1.0):
    """Tax on x with progressive brackets [(over, rate)], thresholds scaled by idx."""
    tax = 0.0
    for i, (lo, rate) in enumerate(brackets):
        lo *= idx
        hi = brackets[i + 1][0] * idx if i + 1 < len(brackets) else float("inf")
        if x > lo:
            tax += (min(x, hi) - lo) * rate
    return tax


def clamp(x, lo=0.0, hi=1.0):
    return max(lo, min(hi, x))


def ss_factor(claim_age, fra):
    """US Social Security: benefit at claim_age as a share of the full-retirement-age benefit."""
    m = round((claim_age - fra) * 12)
    if m >= 0:
        return 1 + 0.08 / 12 * min(m, 36)          # delayed credits stop at 70
    m = -m
    return 1 - (5 / 9 / 100) * min(m, 36) - (5 / 12 / 100) * max(0, m - 36)


def cpp_factor(age):
    """CPP: -0.6%/month before 65, +0.7%/month after (60-70)."""
    m = round((age - 65) * 12)
    return 1 + (0.007 * min(m, 60) if m >= 0 else 0.006 * max(m, -60))


def oas_factor(age):
    """OAS: +0.6%/month deferred past 65 (to 70)."""
    return 1 + 0.006 * clamp(round((age - 65) * 12), 0, 60)


def run(d, S, **over):
    S = dict(S, **over)
    T, hc, P = d["tax"], d["hc"], d["props"]
    us = d["country"] == "US"
    start, end = d["start"], d["end"]
    base = start - 1                                    # "today's dollars" = prices of the year before the plan
    infl = S["inflation"]
    path = S.get("inflation_path")
    tax_year = T["TAX_YEAR"]
    b1, b2 = d["people"][0]["born"], d["people"][1]["born"]
    deaths = [S.get("p1_death_year"), S.get("p2_death_year")]
    survivor_share = S.get("survivor_spending", 1.0)

    levels = {}

    def level(rate, y):
        """Price level vs `base` for something growing at `rate` (on a path it keeps its spread over inflation)."""
        if not path:
            return (1 + rate) ** (y - base)
        if rate not in levels:
            lvl, levels[rate] = 1.0, {base: 1.0}
            for t in range(start, end + 1):
                lvl *= 1 + path[t] + rate - infl
                levels[rate][t] = lvl
        return levels[rate][y]

    def pension_level(y0, y):
        """A pension of 1 (today's $) starting in y0, in year y's dollars: inflation to y0, then the indexed share."""
        k = S["pension_indexing"]
        if not path:
            return (1 + infl) ** (y0 - base) * (1 + k * infl) ** (y - y0)
        y0 = math.ceil(y0)
        lvl = level(infl, y0) if y0 >= start else 1.0
        for t in range(max(y0, base) + 1, y + 1):
            lvl *= 1 + k * path[t]
        return lvl

    def tidx(y):
        """Tax-law index: price level vs TAX_YEAR (flat inflation before the plan starts)."""
        return level(infl, y) * (1 + infl) ** (base - tax_year) if y >= start else (1 + infl) ** (y - tax_year)

    def pidx(y):
        """Provincial index, frozen during the pause years."""
        a, b = T.get("PROV_INDEX_PAUSE_FROM"), T.get("PROV_INDEX_PAUSE_TO")
        if not a or y < a:
            return tidx(y)
        return tidx(y) / (tidx(min(y, b)) / tidx(a - 1))

    sale_factor = S.get("sale_price_factor") or [1.0] * len(P)
    dec = S["props"]
    home = P[0]
    home_sell = dec[0]["sell"]

    # ---- mortgage schedule and property sales (the Model tab's decision block)
    mort_by_year, bal = {}, home["mortgage_balance"]
    for y in range(start, end + 1):
        mort_by_year[y] = bal
        own = home_sell is None or y < home_sell
        interest = bal * home["mortgage_rate"] if own else 0.0
        pay = min(home["mortgage_payment"], bal + interest) if own else 0.0
        bal = bal + interest - pay if own else 0.0
    sale = []
    for i, p in enumerate(P):
        s = dec[i]["sell"]
        if s is None or not p["name"]:
            sale.append(dict(year=None, proceeds=0.0, gain=0.0, recap=0.0))
            continue
        grow = level(p["appreciation"], s) if start <= s <= end else (1 + p["appreciation"]) ** (s - base)
        price = p["value_now"] * grow * sale_factor[i]
        net = price * (1 - p["selling_cost_pct"])
        if i == 0:
            gain = max(0.0, net - p["cost_basis"] - T["HOME_SALE_EXCLUSION"]) if us else 0.0
            sale.append(dict(year=s, proceeds=net - mort_by_year.get(s, 0.0), gain=gain, recap=0.0))
        else:
            taken = p["accum_depreciation"] + p["depreciation"] * max(0, min(s, p["depreciation_end_year"] or 0) - start + 1)
            if us:
                gain = max(0.0, net - (p["cost_basis"] - taken))
                recap = min(gain, taken)
            else:
                gain = max(0.0, net - p["cost_basis"])
                recap = min(taken, max(0.0, net - (p["cost_basis"] - taken)))
            sale.append(dict(year=s, proceeds=net, gain=gain, recap=recap))

    st = d["start_balances"]
    taxable, basis, pre, free, hsa, college = st["taxable"], st["basis"], [st["pre1"], st["pre2"]], st["free"], st["hsa"], st["college"]
    tfsa_room = d.get("tfsa_room0", 0.0)
    mort = home["mortgage_balance"]
    prev, agi_hist, out = None, {}, []
    for y in range(start, end + 1):
        alive = [deaths[i] is None or y < deaths[i] for i in (0, 1)]
        if not any(alive):
            break
        for i in (0, 1):                               # a spouse's pre-tax accounts roll over to the survivor
            if not alive[i] and pre[i]:
                pre[1 - i] += pre[i]
                pre[i] = 0.0
        solo = survivor_share if alive[0] != alive[1] else 1.0
        if not us and y > start:                       # new TFSA room each January for each living spouse
            tfsa_room += T["TFSA_LIMIT"] * tidx(y) * (alive[0] + alive[1])
        age = [y - b1, y - b2]
        R = {"year": y, "age1": age[0], "age2": age[1], "alive1": alive[0], "alive2": alive[1]}
        ret = S["returns"].get(y, S["return_default"]) if S.get("returns") else S["return_default"]
        f = level(infl, y)
        ti = tidx(y)
        medf, colf, ltcf = level(S["med_infl"], y), level(S["col_infl"], y), level(S["ltc_infl"], y)
        R.update(ret=ret, infl_f=f, tax_f=ti)

        # ---- work and one-off events
        working = [1 if (y < S["p1_retire"] and alive[0]) else 0, 1 if (y < S["p2_retire"] and alive[1]) else 0]
        any_work = 1 if (working[0] or working[1]) else 0
        one = {t: 0.0 for t in ONEOFF_TYPES}
        for e in S["oneoffs"]:
            if e["year"] == y and e["type"] in one:
                one[e["type"]] += e["amount"] * f
        pay = [working[i] * S[f"p{i + 1}_pay"] * f + one[f"Wages: person {i + 1}"] for i in (0, 1)]
        contrib = [working[i] * S[f"p{i + 1}_contrib"] * f for i in (0, 1)]
        match = [working[i] * S[f"p{i + 1}_match"] * f for i in (0, 1)]
        hsa_c = any_work * S.get("hsa_contrib", 0.0) * f if us else 0.0
        gross = pay[0] + pay[1]
        taxable_wages = gross - contrib[0] - contrib[1] - hsa_c
        if us:
            wb = T["SS_WAGE_BASE"] * ti
            payroll_i = [0.062 * min(pay[i], wb) for i in (0, 1)]
            payroll = sum(payroll_i) + 0.0145 * gross + 0.009 * max(0.0, gross - T["ADDITIONAL_MEDICARE_THRESHOLD"])
        else:
            ympe, yampe = T["YMPE"] * ti, T["YAMPE"] * ti
            payroll_i = [T["CPP_RATE"] * max(0.0, min(pay[i], ympe) - T["CPP_EXEMPTION"])
                         + T["CPP2_RATE"] * max(0.0, min(pay[i], yampe) - ympe)
                         + T["EI_RATE"] * min(pay[i], T["EI_MAX_INSURABLE"] * ti) for i in (0, 1)]
            payroll = sum(payroll_i)
        withholding = S["withholding_rate"] * taxable_wages
        net_pay = gross - contrib[0] - contrib[1] - hsa_c - payroll - withholding
        R.update(pay1=pay[0], pay2=pay[1], gross_wages=gross, taxable_wages=taxable_wages, payroll_tax=payroll,
                 withholding=withholding, net_pay=net_pay)

        # ---- housing
        own = 1 if (home_sell is None or y < home_sell) else 0
        m_int = own * mort * home["mortgage_rate"]
        m_pay = own * min(home["mortgage_payment"], mort + m_int)
        mort_end = own * (mort + m_int - m_pay)
        carrying = own * home["carrying_costs"] * f
        new_housing = (1 - own) * dec[0]["costs"] * f
        housing = m_pay + carrying + new_housing
        R.update(mort_bal_start=mort, mort_interest=m_int, mort_payment=m_pay, mort_bal_end=mort_end,
                 home_costs=carrying, new_housing=new_housing, housing_total=housing)

        # ---- living
        mult = S["spend_mult"].get(y, 1.0) if S.get("spend_mult") else 1.0
        household = d["household"] * f * mult * solo
        other = sum(o["amount"] for o in d["other_lines"] if o["start"] <= y <= o["end"]) * f * solo
        kids_par = college_cost = 0.0
        for k in d["kids_lines"]:
            if k["start"] <= y <= k["end"]:
                amt = k["amount"] * (colf if k["inflation"] == "college" else f)
                if k["paid_from"] == "college_funds":
                    college_cost += amt
                else:
                    kids_par += amt
        from_funds = min(college_cost, college * (1 + ret))
        col_short = college_cost - from_funds
        college_end = college * (1 + ret) - from_funds
        R.update(household=household, other_budget=other, kids_parents=kids_par, college_cost=college_cost,
                 college_shortfall=col_short, college_end=college_end, oneoff_expense=one["Expense"])

        # ---- healthcare
        n_alive = alive[0] + alive[1]
        emp = any_work * hc["employer_plan_cost"] * medf
        ltc_n = 0
        for i in (0, 1):
            a0, yrs = S[f"p{i + 1}_ltc_age"], S[f"p{i + 1}_ltc_years"]
            if alive[i] and (y >= deaths[i] - yrs if deaths[i] is not None else a0 <= age[i] < a0 + yrs):
                ltc_n += 1
        ltc = ltc_n * S["ltc_cost"] * ltcf
        if us:
            adults = (1 - any_work) * sum(1 for i in (0, 1) if alive[i] and age[i] < 65)
            kids_cov = sum(1 for k in d["kids"] if k["health_through"] and y <= k["health_through"])
            premium = (adults * hc["marketplace_adult_monthly"] + (kids_cov if adults > 0 else 0)
                       * hc["marketplace_child_monthly"]) * 12 * medf
            mkt_oop = (adults + kids_cov) * hc["marketplace_oop_per_person"] * medf if adults > 0 else 0.0
            mn = sum(1 for i in (0, 1) if alive[i] and age[i] >= 65)
            bd = (hc["medicare_part_b_monthly"] + hc["medicare_part_d_monthly"]) * 12 * medf
            mprem = mn * (hc["medicare_part_b_monthly"] + hc["medicare_part_d_monthly"] + hc["medigap_monthly"]) * 12 * medf
            moop = mn * hc["medicare_oop_per_person"] * medf
            irmaa = 0.0
            if y >= start + 2:
                steps, prevc = 0.0, 0.0
                for thr, b, dd in T["irmaa"]:
                    if agi_hist[y - 2] > thr * ti:
                        steps += (b + dd) - prevc
                    prevc = b + dd
                irmaa = mn * 12 * f * steps
            health = premium + mkt_oop + emp + mprem + moop + irmaa + ltc
            eligible = mkt_oop + mn * bd + moop + irmaa + ltc
            hsa_used = 0.0 if any_work else min(hsa * (1 + ret), eligible)
            hsa_end = hsa * (1 + ret) - hsa_used + hsa_c
            R.update(adults_marketplace=adults, kids_covered=kids_cov, marketplace_premium=premium,
                     marketplace_oop=mkt_oop, medicare_premiums=mprem, medicare_oop=moop, irmaa=irmaa)
        else:
            retired_n = (1 - any_work) * n_alive
            premium = retired_n * hc["private_plan_monthly"] * 12 * medf
            oop = retired_n * hc["oop_per_person"] * medf
            health = emp + premium + oop + ltc
            hsa_used, hsa_end = 0.0, 0.0
            R.update(private_premium=premium, health_oop=oop)
        R.update(employer_health=emp, ltc_people=ltc_n, ltc_cost=ltc, health_total=health, hsa_used=hsa_used)

        # ---- other properties (rentals, second homes)
        rent_cash = dep = rental_value = 0.0
        for i, p in enumerate(P[1:], start=1):
            s = dec[i]["sell"]
            if not p["name"] or (s is not None and y >= s):
                continue
            ch = dec[i]["change"]
            if ch is not None and y >= ch:
                rent_cash += (dec[i]["rent"] - dec[i]["costs"]) * f
            else:
                rent_cash += (p["rent"] - p["cash_costs"]) * f
            dep += p["depreciation"] if y <= (p["depreciation_end_year"] or 0) else 0.0
            rental_value += p["value_now"] * level(p["appreciation"], y)
        rental_taxable = max(0.0, rent_cash - dep)
        sale_proceeds = sum(s["proceeds"] for s in sale if s["year"] == y)
        sale_gain = sum(s["gain"] for s in sale if s["year"] == y)
        sale_recap = sum(s["recap"] for s in sale if s["year"] == y)
        home_value = own * home["value_now"] * level(home["appreciation"], y)
        R.update(rental_cash=rent_cash, rental_dep=dep, rental_taxable=rental_taxable, sale_proceeds=sale_proceeds,
                 sale_gain=sale_gain, sale_recapture=sale_recap, home_value=home_value, rental_value=rental_value)

        # ---- government pensions
        people = d["people"]
        if us:
            fra = T["SS_FULL_RETIREMENT_AGE"]
            cut = (1 - S["ss_cut_pct"]) if y >= S["ss_cut_year"] else 1.0
            claimed = [y >= people[i]["born"] + S[f"p{i + 1}_claim_age"] for i in (0, 1)]
            own_b = [people[i]["benefit"] * 12 * (1 - S["ss_early_cut"]) * f * cut for i in (0, 1)]
            gov = []
            for i in (0, 1):
                j = 1 - i
                mine = own_b[i] * ss_factor(S[f"p{i + 1}_claim_age"], fra)
                spousal = 0.5 * own_b[j] * min(1.0, ss_factor(S[f"p{i + 1}_claim_age"], fra) if S[f"p{i + 1}_claim_age"] >= fra
                                               else 1 - (25 / 36 / 100) * min(36, (fra - S[f"p{i + 1}_claim_age"]) * 12)
                                               - (5 / 12 / 100) * max(0, (fra - S[f"p{i + 1}_claim_age"]) * 12 - 36))
                gov.append(0.0 if not claimed[i] else (max(mine, spousal) if claimed[j] else mine))
            if alive[0] != alive[1]:                   # the survivor keeps the larger of the two benefits
                k = 0 if alive[0] else 1
                dec_k = own_b[1 - k] * ss_factor(S[f"p{2 - k}_claim_age"], fra)
                gov = [0.0, 0.0]
                gov[k] = max(own_b[k] * ss_factor(S[f"p{k + 1}_claim_age"], fra), dec_k) if claimed[k] else 0.0
            R.update(ss1=gov[0], ss2=gov[1])
            oas = [0.0, 0.0]
            cpp = gov
        else:
            cpp, oas = [], []
            for i in (0, 1):
                started = y >= people[i]["born"] + S[f"p{i + 1}_cpp_age"]
                cpp.append(people[i]["benefit"] * 12 * cpp_factor(S[f"p{i + 1}_cpp_age"]) * f if started else 0.0)
                o_started = y >= people[i]["born"] + S[f"p{i + 1}_oas_age"]
                oas.append(T["OAS_MONTHLY"] * 12 * ti * min(1.0, people[i]["oas_years"] / 40)
                           * oas_factor(S[f"p{i + 1}_oas_age"]) * (1 + T["OAS_75_BOOST"] * (age[i] >= 75))
                           if o_started else 0.0)
            if alive[0] != alive[1]:
                k = 0 if alive[0] else 1
                surv = T["CPP_SURVIVOR_SHARE"] * people[1 - k]["benefit"] * 12 * f
                cpp[k] += max(0.0, min(surv, T["CPP_MAX_MONTHLY"] * 12 * ti - cpp[k]))
                cpp[1 - k] = oas[1 - k] = 0.0
            R.update(cpp1=cpp[0], cpp2=cpp[1], oas1=oas[0], oas2=oas[1])
            gov = [cpp[0] + oas[0], cpp[1] + oas[1]]
        gov_total = gov[0] + gov[1]
        pens = [0.0, 0.0]
        for i in (0, 1):
            y0 = max(people[i]["born"] + S[f"p{i + 1}_pension_age"], base)
            if S[f"p{i + 1}_pension"] and y >= math.ceil(y0):
                pens[i] = S[f"p{i + 1}_pension"] * pension_level(y0, y)
        if alive[0] != alive[1]:                       # a share continues to the survivor
            k = 0 if alive[0] else 1
            pens[k] += S["pension_survivor"] * pens[1 - k]
            pens[1 - k] = 0.0
        R.update(gov_total=gov_total, pension1=pens[0], pension2=pens[1])

        # ---- investments: required withdrawals, cash needed, where it comes from
        div = S["dividend_yield"] * taxable
        grown = [pre[i] * (1 + ret) for i in (0, 1)]
        req = [0.0, 0.0]
        for i in (0, 1):
            if not alive[i]:
                continue
            if us and age[i] >= rmd_start_age(people[i]["born"]):
                req[i] = max(0.0, pre[i]) / T["uniform"][age[i]]
            if not us and age[i] >= T["RRIF_START_AGE"]:
                req[i] = max(0.0, pre[i]) * T["rrif"][min(age[i] - 1, 120)]
        conv_on = S["conv_from"] is not None and S["conv_to"] is not None and S["conv_from"] <= y <= S["conv_to"]
        conv_want = S["conv_amount"] * f if conv_on else 0.0
        avail = [max(0.0, grown[i] - req[i]) for i in (0, 1)]
        melt = 0.0 if us else min(conv_want, avail[0] + avail[1])  # CA: extra RRSP/RRIF withdrawal, taken as cash
        melt_i = [melt * avail[i] / (avail[0] + avail[1]) if avail[0] + avail[1] > 1 else 0.0 for i in (0, 1)]
        avail = [avail[i] - melt_i[i] for i in (0, 1)]
        true_up = S["tax_due_first"] if prev is None else prev["income_tax"] - prev["withholding"]
        living = housing + household + other + kids_par + col_short + health + one["Expense"]
        cash_in = (net_pay + gov_total + pens[0] + pens[1] + rent_cash + div + req[0] + req[1] + melt + sale_proceeds + hsa_used
                   + one["Other taxable income"] + one["Tax-free money in"])
        need = living + true_up - cash_in
        sell = max(0.0, min(need, taxable * (1 + ret - S["dividend_yield"])))
        from_pre = max(0.0, min(need - sell, avail[0] + avail[1]))
        w = [from_pre * avail[i] / (avail[0] + avail[1]) if avail[0] + avail[1] > 1 else 0.0 for i in (0, 1)]
        free_grown = free * (1 + ret)
        from_free = max(0.0, min(need - sell - from_pre, free_grown))
        from_hsa = max(0.0, min(need - sell - from_pre - from_free, hsa_end)) if us else 0.0
        hsa_end -= from_hsa
        surplus = max(0.0, -need)
        shortfall = max(0.0, need - sell - from_pre - from_free - from_hsa)
        gain_frac = max(0.0, 1 - basis / taxable) if taxable > 0 else 0.0
        realized = sell * gain_frac
        # US: Roth conversion moves pre-tax money to the Roth (no cash). CA: surplus fills TFSA room first.
        left = [avail[i] - w[i] for i in (0, 1)]
        conv = min(conv_want, left[0] + left[1]) if us else 0.0
        conv_i = [conv * left[i] / (left[0] + left[1]) if left[0] + left[1] > 1 else 0.0 for i in (0, 1)]
        if us:
            to_free, to_taxable = conv, surplus
        else:
            room = tfsa_room
            to_free = min(surplus, max(0.0, room))
            to_taxable = surplus - to_free
            tfsa_room = room - to_free + from_free    # withdrawals give the room back the next year
            R.update(tfsa_room=room)
        taxable_end = taxable * (1 + ret - S["dividend_yield"]) - sell + to_taxable
        basis_end = basis - (sell - realized) + to_taxable
        pre_end = [grown[i] - req[i] - melt_i[i] - w[i] - conv_i[i] + contrib[i] + match[i] for i in (0, 1)]
        free_end = free_grown - from_free + to_free
        R.update(dividends=div, req1=req[0], req2=req[1], extra_pre=melt, conversion=conv, tax_true_up=true_up,
                 living_costs=living, cash_in=cash_in, need=need, sell_taxable=sell, from_pre=from_pre, w1=w[0],
                 w2=w[1], from_free=from_free, from_hsa=from_hsa, surplus=surplus, to_free=to_free,
                 shortfall=shortfall, realized_gain=realized, taxable_end=taxable_end, basis_end=basis_end,
                 pre1_end=pre_end[0], pre2_end=pre_end[1], free_end=free_end, hsa_end=hsa_end)

        # ---- income tax
        if us:
            ordinary = (taxable_wages + from_pre + req[0] + req[1] + conv + from_hsa + rental_taxable
                        + one["Other taxable income"] + pens[0] + pens[1])
            pref = div + realized + sale_gain
            ss = gov_total
            prov_inc = ordinary + pref + 0.5 * ss
            if prov_inc <= T["SS_TAXABLE_BASE"]:
                ss_tax = 0.0
            elif prov_inc <= T["SS_TAXABLE_ADJUSTED_BASE"]:
                ss_tax = min(0.5 * ss, 0.5 * (prov_inc - T["SS_TAXABLE_BASE"]))
            else:
                ss_tax = min(0.85 * ss, 0.85 * (prov_inc - T["SS_TAXABLE_ADJUSTED_BASE"])
                             + min(0.5 * ss, T["SS_TAXABLE_ADJUSTED_BASE"] - T["SS_TAXABLE_BASE"]))
            agi = ordinary + pref + ss_tax
            agi_hist[y] = agi
            n65 = sum(1 for i in (0, 1) if alive[i] and age[i] >= 65)
            std = (T["FED_STD"] + T["FED_65_ADD"] * n65) * ti
            senior = (n65 * max(0.0, T["SENIOR_DED"] - T["SENIOR_PHASEOUT_RATE"] * max(0.0, agi - T["SENIOR_PHASEOUT_START"]))
                      if T["SENIOR_DED_FROM"] <= y <= T["SENIOR_DED_TO"] else 0.0)
            salt_years = dict(T["salt"])
            if y in salt_years:
                salt_cap = max(T["SALT_CAP_AFTER"], salt_years[y] - T["SALT_PHASEDOWN_RATE"]
                               * max(0.0, agi - T["SALT_PHASEDOWN_START"] * 1.01 ** (y - tax_year)))
            else:
                salt_cap = T["SALT_CAP_AFTER"]
            state_paid = 0.0 if prev is None else prev["state_tax"] + prev["local_tax"] + prev["other_state_tax"]
            med_ded = max(0.0, health - emp - T["MEDICAL_FLOOR"] * agi)
            prop_tax = own * home["property_tax"] * f
            itemized = m_int + min(salt_cap, state_paid + prop_tax) + med_ded + S["charity"] * f
            ded = max(std, itemized)
            ord_ti = max(0.0, ordinary + ss_tax - ded - senior)
            pref_ti = max(0.0, pref - max(0.0, ded + senior - ordinary - ss_tax))
            fed = (prog(ord_ti, T["fed"], ti) + prog(ord_ti + pref_ti, T["ltcg"], ti) - prog(ord_ti, T["ltcg"], ti)
                   + T["UNRECAPTURED_1250_EXTRA"] * sale_recap
                   + T["NIIT_RATE"] * min(div + realized + rental_taxable + sale_gain, max(0.0, agi - T["NIIT_THRESHOLD"])))
            moved = 1 if (S["move_year"] is not None and y >= S["move_year"]) else 0
            sidx = ti if T["STATE_INDEXED"] else 1.0
            excl = sum(min(w[i] + req[i] + conv_i[i] + pens[i], T["STATE_RET_EXCLUSION"]) for i in (0, 1)
                       if age[i] + 0.5 >= T["STATE_RET_EXCLUSION_AGE"])
            st_agi = agi - T["STATE_SS_EXEMPT"] * ss_tax - excl
            st_ti = max(0.0, st_agi - T["STATE_STD"] * sidx)
            state = (1 - moved) * prog(st_ti, T["state"], sidx)
            local = (1 - moved) * T["LOCAL_RATE"] * st_ti
            oth = moved * S["other_state_rate"] * max(0.0, agi - ss_tax)
            income_tax = fed + state + local + oth
            R.update(ordinary=ordinary, pref=pref, ss_taxable=ss_tax, agi=agi, std_deduction=std, senior_deduction=senior,
                     salt_cap=salt_cap, medical_deduction=med_ded, itemized=itemized, deduction=ded, ord_ti=ord_ti,
                     pref_ti=pref_ti, fed_tax=fed, state_taxable=st_ti, state_tax=state, local_tax=local,
                     other_state_tax=oth, income_tax=income_tax)
        else:
            js = d["joint_share"]
            share = [js, 1 - js] if alive[0] and alive[1] else [1.0 * alive[0], 1.0 * alive[1]]
            eds = S["eligible_div_share"]
            gu = T["ELIGIBLE_GROSSUP"]
            incl = T["CG_INCLUSION"]
            rrif = [req[i] + melt_i[i] + w[i] for i in (0, 1)]
            elig = [pens[i] + (rrif[i] if age[i] >= 65 else 0.0) for i in (0, 1)]
            gross_div = [div * share[i] * eds * (1 + gu) for i in (0, 1)]
            pre_split = [pay[i] + cpp[i] + oas[i] + pens[i] + rrif[i] + gross_div[i] + div * share[i] * (1 - eds)
                         + (realized + sale_gain) * share[i] * incl + (sale_recap + rental_taxable
                                                                        + one["Other taxable income"]) * share[i]
                         - contrib[i] for i in (0, 1)]
            hi = 0 if pre_split[0] >= pre_split[1] else 1
            lo = 1 - hi
            t = min(0.5 * elig[hi], max(0.0, (pre_split[hi] - pre_split[lo]) / 2)) if alive[0] and alive[1] else 0.0
            net = list(pre_split)
            net[hi] -= t
            net[lo] += t
            pension = [elig[i] for i in (0, 1)]
            pension[hi] -= t
            pension[lo] += t
            med_exp = health - emp
            claimant = 0 if (net[0] <= net[1] and alive[0]) or not alive[1] else 1
            tax, fed_t, prov_t, claw = [], [], [], []
            for i in (0, 1):
                if not alive[i]:
                    tax.append(0.0); fed_t.append(0.0); prov_t.append(0.0); claw.append(0.0)
                    continue
                pi = pidx(y)
                cb = min(oas[i], T["OAS_CLAWBACK_RATE"] * max(0.0, net[i] - T["OAS_CLAWBACK_THRESHOLD"] * ti))
                taxable_inc = max(0.0, net[i] - cb)
                bpa = (T["FED_BPA_MAX"] - (T["FED_BPA_MAX"] - T["FED_BPA_MIN"])
                       * clamp((net[i] - T["FED_BPA_PHASE_START"] * ti) / ((T["FED_BPA_PHASE_END"] - T["FED_BPA_PHASE_START"]) * ti))) * ti
                is65 = age[i] >= 65
                f_age = is65 * max(0.0, T["FED_AGE_AMOUNT"] * ti - T["AGE_REDUCTION_RATE"] * max(0.0, net[i] - T["FED_AGE_THRESHOLD"] * ti))
                p_age = is65 * max(0.0, T["PROV_AGE_AMOUNT"] * pi - T["AGE_REDUCTION_RATE"] * max(0.0, net[i] - T["PROV_AGE_THRESHOLD"] * pi))
                f_pen = min(T["FED_PENSION_AMOUNT"], pension[i])
                p_pen = min(T["PROV_PENSION_AMOUNT"], pension[i])
                f_emp = min(T["FED_EMPLOYMENT_AMOUNT"] * ti, pay[i])
                f_med = (i == claimant) * max(0.0, med_exp - min(0.03 * net[i], T["FED_MEDICAL_THRESHOLD"] * ti))
                p_med = (i == claimant) * max(0.0, med_exp - min(0.03 * net[i], T["PROV_MEDICAL_THRESHOLD"] * pi))
                fed = max(0.0, prog(taxable_inc, T["fed"], ti) - T["FED_CREDIT_RATE"] * (bpa + f_age + f_pen + f_emp + payroll_i[i] + f_med)
                          - T["FED_ELIGIBLE_DTC"] * gross_div[i])
                prov = max(0.0, prog(taxable_inc, T["prov"], pi)
                           - T["PROV_CREDIT_RATE"] * (T["PROV_BPA"] * pi + p_age + p_pen + payroll_i[i] + p_med)
                           - T["PROV_ELIGIBLE_DTC"] * gross_div[i]
                           - max(0.0, T["PROV_REDUCTION_MAX"] * pi - T["PROV_REDUCTION_RATE"] * max(0.0, net[i] - T["PROV_REDUCTION_THRESHOLD"] * pi)))
                fed_t.append(fed); prov_t.append(prov); claw.append(cb)
                tax.append(fed + prov + cb)
            income_tax = tax[0] + tax[1]
            R.update(net_income1=net[0], net_income2=net[1], split=t, oas_clawback1=claw[0], oas_clawback2=claw[1],
                     fed_tax1=fed_t[0], fed_tax2=fed_t[1], prov_tax1=prov_t[0], prov_tax2=prov_t[1],
                     income_tax=income_tax)

        investments = taxable_end + pre_end[0] + pre_end[1] + free_end + hsa_end
        net_worth = investments + home_value - mort_end + rental_value
        R.update(investments_end=investments, net_worth_end=net_worth, investments_today=investments / f,
                 net_worth_today=net_worth / f)
        out.append(R)
        prev = R
        taxable, basis, pre, free, hsa, college, mort = taxable_end, basis_end, pre_end, free_end, hsa_end, college_end, mort_end
    return out
