#!/usr/bin/env python3
"""Generate the synthetic sample statements in examples/us/statements and examples/ca/statements.

Everything here is made up (people, accounts, amounts); only the file layouts copy what banks really export:
Chase card CSV, Amex CSV, Apple Card CSV, OFX for a checking account, a brokerage positions CSV, RBC chequing CSV,
a TD Visa CSV (no header row) and a holdings CSV with book values. Rerun to regenerate (same seed, same files).
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
rng = random.Random(2026)
YEAR = 2026


def days(step=1, start=date(YEAR, 1, 1), end=date(YEAR, 12, 31)):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=step)


def monthly(day):
    return [date(YEAR, m, day) for m in range(1, 13)]


def amt(lo, hi):
    return round(rng.uniform(lo, hi), 2)


def spread(n, merchants, lo, hi):
    """n random purchases over the year from a list of merchants."""
    out = []
    for _ in range(n):
        d = date(YEAR, 1, 1) + timedelta(days=rng.randrange(365))
        out.append((d, rng.choice(merchants), amt(lo, hi)))
    return out


def write(path, header, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        if header:
            w.writerow(header)
        w.writerows(rows)


# ============================================================================================== US
def us():
    out = HERE / "us" / "statements"
    # ---- Chase Sapphire (Chase CSV: purchases negative, issuer categories)
    p = []
    for d, m, a in spread(95, [("WHOLE FOODS MARKET #10234", "Groceries"), ("PRICE CHOPPER #188", "Groceries"),
                               ("TRADER JOE S #551", "Groceries")], 45, 210):
        p.append((d, m[0], m[1], -a))
    for d, m, a in spread(70, [("DOORDASH*THAI GARDEN", "Food & Drink"), ("STARBUCKS STORE 11923", "Food & Drink"),
                               ("OSTERIA RESTAURANT", "Food & Drink"), ("CHIPOTLE 2231", "Food & Drink")], 8, 120):
        p.append((d, m[0], m[1], -a))
    for d, m, a in spread(40, [("SHELL OIL 57444", "Gas"), ("SUNOCO 0823", "Gas"), ("EZPASS NY REPLENISH", "Travel")], 30, 75):
        p.append((d, m[0], m[1], -a))
    for d, m, a in spread(8, [("DELTA AIR LINES", "Travel"), ("MARRIOTT BOSTON", "Travel"), ("AIRBNB * HMQ2ZK", "Travel")], 250, 1400):
        p.append((d, m[0], m[1], -a))
    for d, m, a in spread(45, [("AMAZON MKTPL*2K8LM", "Shopping"), ("TARGET 00012", "Shopping"), ("BEST BUY 00419", "Shopping")], 12, 260):
        p.append((d, m[0], m[1], -a))
    for d, m, a in spread(10, [("THE HOME DEPOT #1234", "Home"), ("PETSMART # 0612", "Shopping")], 25, 180):
        p.append((d, m[0], m[1], -a))
    for d in monthly(3):
        p.append((d, "NETFLIX.COM", "Bills & Utilities", -22.99))
        p.append((d + timedelta(days=4), "VERIZON WRLS P1234-01", "Bills & Utilities", -148.20))
    p.append((date(YEAR, 5, 14), "LOCAL BOOKSHOP", "Shopping", -64.10))
    p.append((date(YEAR, 6, 2), "OSTERIA RESTAURANT", "Food & Drink", 18.00))   # a refund
    rows = []
    for d, desc, cat, a in sorted(p):
        rows.append([d.strftime("%m/%d/%Y"), (d + timedelta(days=1)).strftime("%m/%d/%Y"), desc, cat,
                     "Sale" if a < 0 else "Return", f"{a:.2f}", ""])
    for d in monthly(22):
        rows.append([d.strftime("%m/%d/%Y"), d.strftime("%m/%d/%Y"), "Payment Thank You-Mobile", "", "Payment", "2400.00", ""])
    rows.sort(key=lambda r: (r[0][6:], r[0][:5]), reverse=True)
    write(out / "chase-sapphire" / f"Chase4821_Activity{YEAR}0101_{YEAR}1231_{YEAR + 1}0103.CSV",
          ["Transaction Date", "Post Date", "Description", "Category", "Type", "Amount", "Memo"], rows)

    # ---- Amex (positive = charge)
    p = spread(12, ["ALASKA AIRLINES SEATTLE", "HILTON HOTELS", "HERTZ RENT-A-CAR"], 180, 900) + \
        spread(25, ["BLUE HILL TAVERN", "SUSHI ZEN", "LA BELLE BISTRO"], 60, 240) + \
        spread(6, ["WINE CELLAR DIRECT"], 40, 150)
    rows = [[d.strftime("%m/%d/%Y"), m, f"{a:.2f}", "ALEX RIVERA", "-31005", "", "", ""] for d, m, a in sorted(p)]
    rows += [[d.strftime("%m/%d/%Y"), "AUTOPAY PAYMENT - THANK YOU", "-900.00", "ALEX RIVERA", "-31005", "", "", ""] for d in monthly(15)]
    write(out / "amex-gold" / "activity.csv",
          ["Date", "Description", "Amount", "Card Member", "Account #", "Extended Details", "Appears On Your Statement As", "Category"],
          rows)

    # ---- Apple Card, Family Sharing: Riley's purchases are support
    rows = []
    for who, n, merchants in (("Jordan", 60, [("Apple Services", "Other"), ("Uber Eats", "Restaurants"), ("Lululemon", "Shopping")]),
                              ("Riley", 120, [("Campus Market", "Grocery"), ("Sweetgreen", "Restaurants"), ("Uber", "Transportation"),
                                              ("Urban Outfitters", "Shopping")])):
        for d, (m, cat), a in spread(n, merchants, 6, 95):
            rows.append([d.strftime("%m/%d/%Y"), (d + timedelta(days=1)).strftime("%m/%d/%Y"), f"{m.upper()} 1 INFINITE LOOP",
                         m, cat, "Purchase", f"{a:.2f}", who])
    for d in monthly(28):
        rows.append([d.strftime("%m/%d/%Y"), d.strftime("%m/%d/%Y"), "ACH DEPOSIT INTERNET TRANSFER FROM ACCOUNT ENDING IN 7710",
                     "ACH Deposit", "Payment", "Payment", "-850.00", "Jordan"])
    rows.sort(key=lambda r: (r[0][6:], r[0][:5]), reverse=True)
    write(out / "apple-card" / f"Apple Card Transactions - {YEAR}.csv",
          ["Transaction Date", "Clearing Date", "Description", "Merchant", "Category", "Type", "Amount (USD)", "Purchased By"], rows)

    # ---- Checking account as OFX
    tx = []
    for d in days(14, date(YEAR, 1, 9)):
        tx.append((d, "CREDIT", 5410.18, "NORTHWIND LABS PAYROLL", "DIRECT DEP"))
        tx.append((d, "CREDIT", 2860.44, "CITY SCHOOL DIST PAYROLL", "DIRECT DEP"))
    for d in monthly(1):
        tx.append((d, "DEBIT", -1845.00, "SARATOGA SAVINGS MORTGAGE", "MTG PMT"))
        tx.append((d + timedelta(days=5), "DEBIT", -195.00 - rng.randrange(0, 90), "NATIONAL GRID", "WEB PMT"))
        tx.append((d + timedelta(days=6), "DEBIT", -2400.00, "CHASE CREDIT CRD AUTOPAY", "PPD"))
        tx.append((d + timedelta(days=8), "DEBIT", -900.00, "AMEX EPAYMENT", "ACH PMT"))
        tx.append((d + timedelta(days=9), "DEBIT", -850.00, "APPLECARD GSBANK PAYMENT", "WEB"))
        tx.append((d + timedelta(days=10), "DEBIT", -1500.00, "VENMO PAYMENT RILEY RIVERA", "rent help"))
        tx.append((d + timedelta(days=12), "DEBIT", -2500.00, "FIDELITY MONEYLINE", "INVESTMENT"))
        tx.append((d + timedelta(days=14), "DEBIT", -350.00, "GEICO AUTO", "INS PREM"))
        tx.append((d + timedelta(days=16), "DEBIT", -200.00, "ATM WITHDRAWAL 0042", ""))
        tx.append((d + timedelta(days=18), "DEBIT", -45.00 - rng.randrange(0, 60), "SARATOGA DENTAL GROUP", ""))
    for m in (1, 4, 7, 10):
        tx.append((date(YEAR, m, 15), "DEBIT", -2000.00, "US TREASURY TAX PMT", "1040-ES"))
    for m in (1, 9):
        tx.append((date(YEAR, m, 20), "DEBIT", -5600.00, "CITY OF SARATOGA SPRINGS PROPERTY TAX", ""))
    tx.append((date(YEAR, 3, 3), "DEBIT", -2400.00, "ERIE INSURANCE HOMEOWNERS", "annual"))
    tx.append((date(YEAR, 8, 18), "DEBIT", -12600.00, "STATE UNIVERSITY BURSAR", "tuition"))
    tx.sort()
    bal = 18250.00 + sum(t[2] for t in tx)
    trn = "\n".join(
        f"<STMTTRN>\n<TRNTYPE>{t}\n<DTPOSTED>{d:%Y%m%d}120000\n<TRNAMT>{a:.2f}\n<FITID>{d:%Y%m%d}{i:04d}\n<NAME>{n}\n<MEMO>{m}\n</STMTTRN>"
        for i, (d, t, a, n, m) in enumerate(tx))
    ofx = f"""OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
<SIGNONMSGSRSV1><SONRS><STATUS><CODE>0<SEVERITY>INFO</STATUS><DTSERVER>{YEAR + 1}0103120000<LANGUAGE>ENG</SONRS></SIGNONMSGSRSV1>
<BANKMSGSRSV1><STMTTRNRS><TRNUID>1<STATUS><CODE>0<SEVERITY>INFO</STATUS>
<STMTRS><CURDEF>USD
<BANKACCTFROM><BANKID>000000000<ACCTID>000077107710<ACCTTYPE>CHECKING</BANKACCTFROM>
<BANKTRANLIST><DTSTART>{YEAR}0101<DTEND>{YEAR}1231
{trn}
</BANKTRANLIST>
<LEDGERBAL><BALAMT>{bal:.2f}<DTASOF>{YEAR}1231120000</LEDGERBAL>
</STMTRS></STMTTRNRS></BANKMSGSRSV1>
</OFX>
"""
    (out / "checking").mkdir(parents=True, exist_ok=True)
    (out / "checking" / f"checking_{YEAR}.qfx").write_text(ofx)

    # ---- Brokerage positions (a typical "positions" CSV export)
    write(out / "brokerage" / f"Portfolio_Positions_Dec-31-{YEAR}.csv",
          ["Account Number", "Account Name", "Symbol", "Description", "Quantity", "Last Price", "Current Value", "Cost Basis Total", "Type"],
          [["X00000001", "Joint Brokerage", "VTI", "VANGUARD TOTAL STOCK MKT ETF", "1850", "331.20", "612720.00", "402100.00", "Cash"],
           ["X00000001", "Joint Brokerage", "VXUS", "VANGUARD TOTAL INTL STOCK ETF", "2200", "70.15", "154330.00", "121800.00", "Cash"],
           ["X00000001", "Joint Brokerage", "BND", "VANGUARD TOTAL BOND MARKET ETF", "1400", "74.10", "103740.00", "104900.00", "Cash"],
           ["X00000001", "Joint Brokerage", "SPAXX**", "HELD IN MONEY MARKET", "", "", "11890.55", "", "Cash"],
           ["Z00000002", "Alex Rollover IRA", "FXAIX", "FIDELITY 500 INDEX FUND", "940", "225.40", "211876.00", "", "Cash"],
           ["Z00000003", "Alex Roth IRA", "FZROX", "FIDELITY ZERO TOTAL MARKET", "4700", "20.55", "96585.00", "", "Cash"],
           [], ["Date downloaded 01/03/2027 9:14 AM ET"]])


# ============================================================================================== Canada
def ca():
    out = HERE / "ca" / "statements"
    # ---- RBC chequing CSV (signed amounts, two description columns)
    rows = []
    add = lambda d, d1, d2, a: rows.append(["Chequing", "01234-5006789", d.strftime("%m/%d/%Y"), "", d1, d2, f"{a:.2f}", ""])  # noqa: E731
    for d in days(14, date(YEAR, 1, 2)):
        add(d, "PAYROLL DEPOSIT", "CASCADIA ENGINEERING", 4210.55)
        add(d, "PAYROLL DEPOSIT", "VANCOUVER COASTAL HLTH", 3120.80)
    for d in monthly(1):
        add(d, "MORTGAGE PAYMENT", "RBC MORTGAGE 00456", -3033.33)
        add(d + timedelta(days=3), "BC HYDRO", "BILL PAYMENT", -95.00 - rng.randrange(0, 60))
        add(d + timedelta(days=4), "FORTISBC GAS", "BILL PAYMENT", -70.00 - rng.randrange(0, 60))
        add(d + timedelta(days=5), "TELUS", "BILL PAYMENT", -185.40)
        add(d + timedelta(days=7), "TD VISA", "PAYMENT - VISA", -2100.00)
        add(d + timedelta(days=9), "E-TRANSFER SENT", "MORGAN LEE", -1000.00)
        add(d + timedelta(days=11), "WEALTHSIMPLE", "TFSA CONTRIBUTION", -500.00)
        add(d + timedelta(days=13), "ICBC", "AUTOPLAN PREAUTH", -168.00)
        add(d + timedelta(days=15), "ATM WITHDRAWAL", "RBC ABM 4412", -140.00)
    for m in (7,):
        add(date(YEAR, m, 2), "CITY OF VANCOUVER", "PROPERTY TAX", -6300.00)
    for m in (3, 6, 9, 12):
        add(date(YEAR, m, 20), "STRATA", "BURNABY CONDO RENTAL STRATA", -1150.00)
    for d in monthly(1):
        add(d, "E-TRANSFER RECEIVED", "BURNABY CONDO TENANT RENT", 2700.00)
    add(date(YEAR, 9, 5), "UBC TUITION", "BILL PAYMENT", -4750.00)
    rows.sort(key=lambda r: (r[2][6:], r[2][:5]))
    write(out / "rbc-chequing" / f"csv{YEAR}1231.csv",
          ["Account Type", "Account Number", "Transaction Date", "Cheque Number", "Description 1", "Description 2", "CAD$", "USD$"], rows)

    # ---- TD Visa CSV: no header; date, description, debit, credit, balance
    p = spread(100, ["SAVE ON FOODS #2231", "COSTCO WHOLESALE W548", "T&T SUPERMARKET #009"], 40, 260) + \
        spread(60, ["TIM HORTONS #5012", "SKIPTHEDISHES", "MIKU RESTAURANT", "JJ BEAN COFFEE"], 6, 140) + \
        spread(30, ["PETRO-CANADA 4491", "ESSO CIRCLE K", "IMPARK00123"], 25, 90) + \
        spread(8, ["AIR CANADA 0142", "WESTJET AIR", "FAIRMONT HOTEL"], 200, 1300) + \
        spread(30, ["AMAZON.CA", "CANADIAN TIRE #321", "MEC VANCOUVER", "WINNERS #778"], 15, 250) + \
        spread(12, ["SHOPPERS DRUG MART #22", "LONDON DRUGS 88"], 12, 80) + \
        spread(6, ["GROUSE MOUNTAIN", "CINEPLEX ODEON"], 30, 140)
    for d in monthly(6):
        p.append((d, "NETFLIX.COM", 18.99))
        p.append((d, "SPOTIFY P1A2B3", 12.99))
    rows, balance = [], 0.0
    items = [(d, m, a, "") for d, m, a in p] + [(d, "PAYMENT - THANK YOU", "", 2100.00) for d in monthly(8)]
    for d, m, a, c in sorted(items, key=lambda x: x[0]):
        balance += (a or 0) - (c or 0)
        rows.append([d.strftime("%m/%d/%Y"), m, f"{a:.2f}" if a != "" else "", f"{c:.2f}" if c != "" else "", f"{balance:.2f}"])
    write(out / "td-visa" / f"accountactivity_{YEAR}.csv", None, rows)

    # ---- Holdings export with book values (Canadian brokers say "book value" for cost)
    write(out / "wealthsimple" / f"holdings-report-{YEAR}-12-31.csv",
          ["Account Name", "Account Type", "Symbol", "Name", "Quantity", "Market Price", "Market Value (CAD)", "Book Value (CAD)"],
          [["Joint non-registered", "Non-registered", "XEQT", "iShares Core Equity ETF Portfolio", "9800", "35.10", "343980.00", "241000.00"],
           ["Joint non-registered", "Non-registered", "ZAG", "BMO Aggregate Bond Index ETF", "5200", "14.20", "73840.00", "76200.00"],
           ["Sam TFSA", "TFSA", "VEQT", "Vanguard All-Equity ETF Portfolio", "2600", "46.25", "120250.00", "88000.00"],
           ["Taylor TFSA", "TFSA", "VEQT", "Vanguard All-Equity ETF Portfolio", "2100", "46.25", "97125.00", "70100.00"]])


if __name__ == "__main__":
    us()
    ca()
    print("wrote examples/us/statements and examples/ca/statements")
