"""Read bank, credit-card and investment statement exports into one standard format.

Transactions: date (YYYY-MM-DD), account, description, amount (negative = money out), source_category (the bank's
own category, if any), person (cardholder, if any), source_file.

Supported:
  - OFX / QFX / QBO (what "Download for Quicken / Money" gives you) - bank, card and investment accounts
  - CSV exports: known layouts (Chase, Amex, Apple Card, Capital One, Citi, Discover, Bank of America, RBC, TD,
    and this project's own standard CSV) are recognized by their header; anything else is auto-detected (a date,
    a description and an amount or debit/credit columns). Check the per-file summary the pipeline prints: if a
    file's money in and out look swapped, add a line for it to rules/statement_formats.csv.
  - PDFs aren't parsed. Ask Claude to read them and write the standard CSV into statements/converted/.
Positions (for balances): OFX investment statements, and CSV "positions" / "holdings" exports that have a market
value column (and ideally a cost basis / book value column).
"""
import csv
import io
import re
from datetime import date, datetime
from pathlib import Path

STANDARD_FIELDS = ["date", "account", "description", "amount", "source_category", "person", "source_file"]
DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y/%m/%d", "%d-%b-%Y", "%b %d, %Y", "%d %b %Y", "%Y%m%d"]


def parse_date(s, fmt=None):
    s = str(s).strip()
    if fmt:
        return datetime.strptime(s, fmt).date()
    for f in DATE_FORMATS:
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    raise ValueError(f"unrecognized date {s!r}")


def parse_amount(s):
    s = str(s or "").strip().replace("$", "").replace(",", "").replace("CAD", "").replace("USD", "").strip()
    if s in ("", "-"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    v = float(s)
    return -v if neg else v


# ---------------------------------------------------------------------------------------------- OFX

def _ofx_blocks(text, tag):
    """Contents of every <TAG>...</TAG> block (works for SGML OFX 1.x and XML OFX 2.x)."""
    return re.findall(rf"<{tag}>(.*?)</{tag}>", text, flags=re.S | re.I)


def _ofx_field(block, tag):
    m = re.search(rf"<{tag}>([^<\r\n]*)", block, flags=re.I)
    return m.group(1).strip() if m else ""


def _ofx_date(s):
    return datetime.strptime(s[:8], "%Y%m%d").date()


def ofx_transactions(path):
    text = Path(path).read_text(errors="replace")
    rows = []
    for kind in ("STMTRS", "CCSTMTRS"):
        for stmt in _ofx_blocks(text, kind):
            acct = _ofx_field(stmt, "ACCTID")
            label = f"{'card' if kind == 'CCSTMTRS' else 'bank'}_{acct[-4:] or 'unknown'}"
            for t in _ofx_blocks(stmt, "STMTTRN"):
                name, memo = _ofx_field(t, "NAME"), _ofx_field(t, "MEMO")
                desc = name if not memo or memo in name else f"{name} {memo}"
                rows.append(dict(date=_ofx_date(_ofx_field(t, "DTPOSTED")).isoformat(), account=label,
                                 description=desc.replace("&amp;", "&"), amount=float(_ofx_field(t, "TRNAMT")),
                                 source_category="", person="", source_file=Path(path).name))
    return rows


def ofx_balances(path):
    """[(account label, value, cost basis or None, as_of)] - bank ledger balances and investment positions."""
    text = Path(path).read_text(errors="replace")
    out = []
    for stmt in _ofx_blocks(text, "STMTRS"):
        acct = _ofx_field(stmt, "ACCTID")
        bal = _ofx_blocks(stmt, "LEDGERBAL")
        if bal:
            out.append((f"bank_{acct[-4:]}", float(_ofx_field(bal[0], "BALAMT")), None, _ofx_date(_ofx_field(bal[0], "DTASOF"))))
    for stmt in _ofx_blocks(text, "INVSTMTRS"):
        acct = _ofx_field(stmt, "ACCTID")
        total = sum(float(_ofx_field(p, "MKTVAL") or 0) for p in _ofx_blocks(stmt, "INVPOS"))
        cash = _ofx_field(stmt, "AVAILCASH")
        total += float(cash) if cash else 0.0
        asof = _ofx_field(stmt, "DTASOF") or _ofx_field(stmt, "DTEND")
        out.append((f"invest_{acct[-4:]}", total, None, _ofx_date(asof) if asof else None))
    return out


# ---------------------------------------------------------------------------------------------- CSV

# Known layouts, recognized by their header columns. sign: +1 = the file's amounts already use negative = money out;
# -1 = positive means a purchase (flip). "skip" drops rows whose column matches (card payments are counted once,
# as money leaving the bank account).
PROFILES = [
    dict(name="standard", need={"date", "account", "description", "amount"}, date="date", desc="description",
         amount="amount", sign=1, account_col="account", category="source_category", person="person"),
    dict(name="chase_card", need={"Transaction Date", "Post Date", "Description", "Category", "Type", "Amount"},
         date="Transaction Date", desc="Description", amount="Amount", sign=1, category="Category",
         skip=("Type", r"^Payment$"), account="chase_card"),
    dict(name="chase_bank", need={"Details", "Posting Date", "Description", "Amount", "Type"},
         date="Posting Date", desc="Description", amount="Amount", sign=1, balance="Balance", account="chase_bank"),
    dict(name="apple_card", need={"Transaction Date", "Clearing Date", "Description", "Merchant", "Category", "Type", "Amount (USD)"},
         date="Transaction Date", desc="Description", amount="Amount (USD)", sign=-1, category="Category",
         person="Purchased By", skip=("Type", r"^(Payment|Debit)$"), account="apple_card"),
    dict(name="amex", need={"Date", "Description", "Amount"}, also={"Card Member", "Account #", "Extended Details", "Appears On Your Statement As"},
         date="Date", desc="Description", amount="Amount", sign=-1, category="Category", person="Card Member",
         skip=("Description", r"(AUTOPAY PAYMENT|PAYMENT - THANK YOU|ONLINE PAYMENT)"), account="amex"),
    dict(name="capital_one", need={"Transaction Date", "Posted Date", "Card No.", "Description", "Category", "Debit", "Credit"},
         date="Transaction Date", desc="Description", debit="Debit", credit="Credit", category="Category",
         skip=("Description", r"(CAPITAL ONE (MOBILE|ONLINE) PYMT|PAYMENT/CREDIT)"), account="capital_one"),
    dict(name="citi_card", need={"Status", "Date", "Description", "Debit", "Credit"},
         date="Date", desc="Description", debit="Debit", credit="Credit", skip=("Description", r"(ONLINE PAYMENT|AUTOPAY)"),
         account="citi_card"),
    dict(name="discover", need={"Trans. Date", "Post Date", "Description", "Amount", "Category"},
         date="Trans. Date", desc="Description", amount="Amount", sign=-1, category="Category",
         skip=("Category", r"^Payments and Credits$"), account="discover"),
    dict(name="bofa_bank", need={"Date", "Description", "Amount", "Running Bal."},
         date="Date", desc="Description", amount="Amount", sign=1, balance="Running Bal.", account="bofa_bank"),
    dict(name="rbc", need={"Account Type", "Account Number", "Transaction Date", "Description 1", "Description 2", "CAD$"},
         date="Transaction Date", desc=("Description 1", "Description 2"), amount="CAD$", sign=1, account_col="Account Number"),
]
HEADERLESS = dict(name="headerless (date, description, debit, credit[, balance]) - e.g. TD", date=0, desc=1, debit=2, credit=3,
                  balance=4)

DATE_HINTS = ["transaction date", "trans. date", "trans date", "date", "posting date", "posted date", "post date"]
DESC_HINTS = ["description", "payee", "merchant", "name", "details", "memo", "transaction"]
AMOUNT_HINTS = ["amount", "cad$", "amount (usd)", "amount (cad)", "transaction amount"]
DEBIT_HINTS = ["debit", "withdrawal", "withdrawals", "money out", "paid out"]
CREDIT_HINTS = ["credit", "deposit", "deposits", "money in", "paid in"]


def _find(head, hints):
    low = [h.strip().lower() for h in head]
    for hint in hints:
        for i, h in enumerate(low):
            if h == hint:
                return head[i]
    for hint in hints:
        for i, h in enumerate(low):
            if hint in h:
                return head[i]
    return None


def _read_rows(path):
    raw = Path(path).read_bytes().decode("utf-8-sig", errors="replace")
    return list(csv.reader(io.StringIO(raw)))


def _header_row(rows):
    """Index of the header row: the first row (within the first 20) that names a date column."""
    for i, r in enumerate(rows[:20]):
        cells = [c.strip().lower() for c in r]
        if any("date" in c for c in cells) and len([c for c in cells if c]) >= 3:
            return i
    return None


def detect(path, override=None):
    """(profile dict, header row index) for a CSV file."""
    rows = _read_rows(path)
    if override and override.get("date") not in (None, ""):
        return dict(override, name="rules/statement_formats.csv"), (None if override.get("headerless") else _header_row(rows))
    h = _header_row(rows)
    if h is None:
        if rows and len(rows[0]) >= 4:
            try:
                parse_date(rows[0][0])
                return HEADERLESS, None
            except ValueError:
                pass
        return None, None
    head = [c.strip() for c in rows[h]]
    hs = set(head)
    for p in PROFILES:
        if p["need"] <= hs and (not p.get("also") or p["also"] & hs):
            return p, h
    auto = dict(name="auto-detected", date=_find(head, DATE_HINTS), desc=_find(head, DESC_HINTS),
                amount=_find(head, AMOUNT_HINTS), debit=_find(head, DEBIT_HINTS), credit=_find(head, CREDIT_HINTS),
                category=_find(head, ["category"]), balance=_find(head, ["balance", "running balance"]), sign=1)
    if auto["date"] and auto["desc"] and (auto["amount"] or (auto["debit"] and auto["credit"])):
        if auto["amount"]:
            auto["debit"] = auto["credit"] = None
        return auto, h
    return None, h


def csv_transactions(path, account_label, override=None):
    """(rows, profile name, last running balance or None)"""
    prof, h = detect(path, override)
    if prof is None:
        raise ValueError("couldn't recognize the columns - add a line for this file to rules/statement_formats.csv")
    rows = _read_rows(path)
    headerless = h is None
    head = None if headerless else [c.strip() for c in rows[h]]
    body = rows if headerless else rows[h + 1:]
    col = (lambda r, k: r[k].strip() if isinstance(k, int) and k < len(r) else "") if headerless else \
        (lambda r, k: r[head.index(k)].strip() if k in head and head.index(k) < len(r) else "")
    out, balance = [], None
    skip = prof.get("skip")
    fmt = (override or {}).get("date_format") or None
    sign = int(float((override or {}).get("sign") or prof.get("sign", 1)))
    for r in body:
        if not any(c.strip() for c in r):
            continue
        d = col(r, prof["date"])
        if not d:
            continue
        try:
            dt = parse_date(d, fmt)
        except ValueError:
            continue   # a totals line or a note
        if skip and re.search(skip[1], col(r, skip[0]), flags=re.I):
            continue
        desc = " ".join(col(r, k) for k in prof["desc"]).strip() if isinstance(prof["desc"], tuple) else col(r, prof["desc"])
        if prof.get("amount") not in (None, ""):
            amt = parse_amount(col(r, prof["amount"]))
            if amt is None:
                continue
            amt *= sign
        else:
            deb, cre = parse_amount(col(r, prof["debit"])), parse_amount(col(r, prof["credit"]))
            if deb is None and cre is None:
                continue
            amt = abs(cre or 0.0) - abs(deb or 0.0)
        acct = account_label or (col(r, prof["account_col"]) if prof.get("account_col") else "") or prof.get("account") or Path(path).stem
        if prof.get("balance") is not None:
            b = parse_amount(col(r, prof["balance"]))
            if b is not None:
                balance = (dt, b)
        out.append(dict(date=dt.isoformat(), account=acct, description=re.sub(r"\s+", " ", desc),
                        amount=round(amt, 2), source_category=col(r, prof["category"]) if prof.get("category") else "",
                        person=col(r, prof["person"]) if prof.get("person") else "", source_file=Path(path).name))
    return out, prof["name"], balance


# ---------------------------------------------------------------------------------------------- positions CSVs

VALUE_HINTS = ["current value", "market value", "total value", "market value (cad)", "value", "marketvalue"]
BASIS_HINTS = ["cost basis total", "cost basis", "book value", "book cost", "total cost", "adjusted cost base"]


def is_positions_csv(path):
    rows = _read_rows(path)
    for r in rows[:15]:
        low = [c.strip().lower() for c in r]
        if any(h in low for h in ("symbol", "ticker")) and any(v in " ".join(low) for v in ("value", "market")):
            return True
    return False


def positions_balances(path):
    """[(account label, value, cost basis or None, None)] from a holdings export (grouped by account column if any)."""
    rows = _read_rows(path)
    h = next(i for i, r in enumerate(rows[:15]) if any(c.strip().lower() in ("symbol", "ticker") for c in r))
    head = [c.strip() for c in rows[h]]
    vcol, bcol = _find(head, VALUE_HINTS), _find(head, BASIS_HINTS)
    acol = _find(head, ["account number", "account name", "account"])
    totals = {}
    for r in rows[h + 1:]:
        if len(r) < len(head) // 2:
            break
        cells = dict(zip(head, (c.strip() for c in r)))
        if not cells.get(vcol) or re.search(r"total|pending", " ".join(r[:2]), flags=re.I):
            continue
        try:
            v = parse_amount(cells[vcol])
        except ValueError:
            continue
        if v is None:
            continue
        b = None
        if bcol:
            try:
                b = parse_amount(cells.get(bcol, ""))
            except ValueError:
                b = None
        key = cells.get(acol) if acol else None
        tv, tb = totals.get(key, (0.0, 0.0))
        totals[key] = (tv + v, tb + (b if b is not None else v))
    return [(k, v, b, None) for k, (v, b) in totals.items()]


def today():
    return date.today()
