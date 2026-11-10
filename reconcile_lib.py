"""
reconcile_lib.py - matching engine for the daily Smart Mart / Smart Fashions
sales-vs-payments reconciliation.

KEY DESIGN: no payment source is pre-assigned to a shop. Every actual payment
(bank UPI credit, PhonePe txn, card-machine swipe, direct transfer) goes into
ONE shared pool and ALL shops' bills are matched against it together. So a new
card machine (for either shop) is just one more file - and cross-shop routing
(a Mart sale paid via the Fashions QR) is discovered automatically.
"""
from __future__ import annotations
import csv
import itertools
import os
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import openpyxl

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

AMOUNT_TOLERANCE = 15     # rupees of rounding / typo noise allowed in matching (max)
AMOUNT_TOLERANCE_PCT = 0.10   # ...but never more than 10% of the bill, so small bills must match closely


def tolerance_for(card, amount_tol=AMOUNT_TOLERANCE):
    return min(amount_tol, max(1.0, AMOUNT_TOLERANCE_PCT * card))
TIME_TOLERANCE_MIN = 240  # minutes, for non-exact amounts (tills batch-enter bills)
ANOMALY_TOLERANCE = 15    # digital recorded above bill Net by more than this = flagged
COMBO_AMOUNT_TOL = 0.05   # rupees: combined / split payments must add up exactly (limits coincidences)
COMBO_BILL_WINDOW_MIN = 60    # bills paid together are billed within this many minutes of each other
COMBO_PAY_WINDOW_MIN = 240    # a payment with a real clock time must be this close to its bills
COMBO_MIN_PART = 20       # rupees: smaller payments/bills (e.g. Rs.5 test payments) are never merged into combos
SPLIT_SEP = " + "         # joins source names when one bill was paid through several sources


@dataclass
class Bill:
    shop: str
    billno: str
    time: datetime
    net: float
    card: float
    system: str = ""


@dataclass
class PaymentRecord:
    amount: float
    order_key: object          # datetime if the source has real times, else int sequence
    time_str: str
    source: str
    ref: str
    shop_hint: Optional[str] = None
    parts: Optional[list] = None   # [(source, amount)] when one bill was paid through several payments
    suspects: Optional[list] = None   # unassigned payment only: cash-keyed bills that could be its bill


@dataclass
class MatchResult:
    bill: Bill
    payment: Optional[PaymentRecord]
    amount_diff: Optional[float]
    combo: Optional[str] = None    # "combined" (one payment, several bills) | "split" (one bill, several payments)
    combo_note: str = ""
    used: Optional[list] = None    # the ORIGINAL payment record(s) this bill consumed (for payment -> bill lookups)


def portions(m):
    """[(source, amount)] actually credited to this bill."""
    if not m.payment:
        return []
    return m.payment.parts or [(m.payment.source, m.payment.amount)]


# ------------------------------------------------------------------ helpers
def _to_number(val) -> Optional[float]:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace(",", "").replace("Rs.", "").replace("\u20b9", "").strip()
    try:
        return float(s) if s not in ("", "-") else None
    except ValueError:
        return None


def _parse_datetime(val) -> Optional[datetime]:
    if val is None:
        return None
    if isinstance(val, datetime):
        return val
    s = str(val).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M",
                "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def _parse_time_on_date(val, day: Optional[datetime]) -> Optional[datetime]:
    """Card-machine slips often give only HH:MM[:SS]; attach the business date."""
    full = _parse_datetime(val)
    if full:
        return full
    if val is None or day is None:
        return None
    s = str(val).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            t = datetime.strptime(s, fmt)
            return day.replace(hour=t.hour, minute=t.minute, second=t.second)
        except ValueError:
            pass
    return None


def _pick(cols: dict, candidates):
    for cand in candidates:
        for lower, original in cols.items():
            if cand in lower:
                return original
    return None


# ------------------------------------------------------------------ loaders
def _find_pos_header(rows):
    """rows: first few rows (tuples). -> (header_index, {field: col}) or (None, {})"""
    for r, row in enumerate(rows):
        lower = [str(v or "").strip().lower() for v in row]
        if any("bill" in x for x in lower) and "net" in lower and "card" in lower:
            cmap = {}
            for i, x in enumerate(lower):
                if "bill" in x and "billno" not in cmap:
                    cmap["billno"] = i
                elif x == "net":
                    cmap["net"] = i
                elif x == "card":
                    cmap["card"] = i
                elif x == "time":
                    cmap["time"] = i
                elif x in ("system", "till", "counter"):
                    cmap["system"] = i
            if {"billno", "net", "card", "time"} <= set(cmap):
                return r, cmap
    return None, {}


def load_pos_bills(path: str, shop: str) -> list[Bill]:
    """Shop POS export (.xlsx). Header row is located by text, so column order
    may drift between exports. Read-only streaming with the declared sheet size
    ignored: some billing exports declare ~1M rows x 16K columns, which makes a
    normal openpyxl load/iteration run for minutes."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            ws.reset_dimensions()
            rows = ws.iter_rows(values_only=True)
            head = []
            for row in rows:
                head.append(row)
                if len(head) >= 15:
                    break
            hr, cmap = _find_pos_header(head)
            if hr is None:
                continue
            bills, empty = [], 0

            def consume(row):
                nonlocal empty
                if not row or all(v is None for v in row):
                    empty += 1
                    return
                empty = 0
                if len(row) <= max(cmap.values()):
                    return
                billno = row[cmap["billno"]]
                if billno is None or str(billno).strip().upper() in ("", "TOTAL", "GRAND TOTAL"):
                    return
                net = _to_number(row[cmap["net"]])
                t = _parse_datetime(row[cmap["time"]])
                if net is None or t is None:
                    return
                card = _to_number(row[cmap["card"]]) or 0.0
                system = str(row[cmap["system"]] or "") if "system" in cmap else ""
                bills.append(Bill(shop, str(billno).strip(), t, net, card, system))
            for row in head[hr + 1:]:
                consume(row)
            for row in rows:
                consume(row)
                if empty >= 500:   # long run of blank rows = end of data
                    break
            return bills
    finally:
        wb.close()
    raise ValueError(f"Could not find Bill No / Net / Card / Time columns in {path}")


def _bank_page_rows(page):
    """[(deposit, ref_or_None, date_str)] for one statement page."""
    out = []
    for table in page.extract_tables():
        if not table:
            continue
        header = [str(c or "").strip().lower() for c in table[0]]

        def col(*names):
            for n in names:
                for i, h in enumerate(header):
                    if n in h:
                        return i
        dep_c, id_c, date_c = col("deposit", "credit"), col("tran id", "reference"), col("date")
        if dep_c is None:
            continue
        for row in table[1:]:
            if not row or len(row) <= dep_c:
                continue
            if "total" in " ".join(str(c or "") for c in row).lower():
                continue
            dep = _to_number(row[dep_c])
            if not dep:
                continue
            ref = str(row[id_c]).strip() if id_c is not None and row[id_c] else None
            date_s = str(row[date_c]).strip() if date_c is not None and row[date_c] else ""
            out.append((dep, ref, date_s))
    return out


def load_federal_bank_pdf(path: str, source_label: str = "Federal Bank UPI",
                          shop_hint: Optional[str] = None) -> list[PaymentRecord]:
    """DEPOSIT rows of a bank statement PDF. Withdrawals (charges, transfers
    out) are never sale receipts and are ignored; 'Grand Total' rows skipped.
"""
    if pdfplumber is None:
        raise RuntimeError("pdfplumber is required: pip install pdfplumber")
    import re
    with pdfplumber.open(path) as pdf:
        n = len(pdf.pages)
        first_text = (pdf.pages[0].extract_text() or "") if n else ""
        m_name = re.search(r"Name\s+([A-Za-z][A-Za-z &.]+?)(?:\s+Branch|\n|$)", first_text)
        holder = shop_hint or (m_name.group(1).strip() if m_name else None)
        by_page = {}
        for i, page in enumerate(pdf.pages):
            t0 = time.time()
            by_page[i] = _bank_page_rows(page)
            page.flush_cache()
            print(f"[bank pdf] {os.path.basename(path)} page {i + 1}/{n}: {time.time() - t0:.1f}s, "
                  f"{len(by_page[i])} deposits", flush=True)
    records, seq = [], 0
    for i in sorted(by_page):
        for dep, ref, date_s in by_page[i]:
            seq += 1
            records.append(PaymentRecord(dep, seq, date_s, source_label, ref or f"row{seq}", holder))
    return records


def load_phonepe_csv(path: str, source_label: str = "PhonePe") -> list[PaymentRecord]:
    """PhonePe / EDC settlement CSV. Only COMPLETED rows count."""
    records = []
    with open(path, encoding="utf-8", errors="replace") as f:
        reader = csv.DictReader(f)
        cols = {c.strip().lower(): c for c in (reader.fieldnames or [])}
        amt_c = _pick(cols, ["total transaction amount", "amount"])
        time_c = _pick(cols, ["transaction date", "date"])
        stat_c = _pick(cols, ["transaction status", "status"])
        ref_c = _pick(cols, ["phonepe reference id", "reference id", "transaction id"])
        store_c = _pick(cols, ["store name", "shop", "outlet"])
        for row in reader:
            if stat_c and str(row.get(stat_c, "")).strip().upper() not in ("COMPLETED", "SUCCESS", ""):
                continue
            amt = _to_number(row.get(amt_c)) if amt_c else None
            if not amt:
                continue
            t = _parse_datetime(row.get(time_c)) if time_c else None
            records.append(PaymentRecord(amt, t or 0, row.get(time_c, "") if time_c else "",
                                         source_label, row.get(ref_c, "") if ref_c else "",
                                         (row.get(store_c) or "").strip() or None if store_c else None))
    return records


def load_card_machine_csv(path: str, machine_label: str, shop_hint: Optional[str] = None,
                          business_date: Optional[datetime] = None) -> list[PaymentRecord]:
    """Generic card-machine list: columns time, amount, ref (card_type/card_no optional).
    Works for ANY machine, for ANY shop - they all land in the same pool."""
    records = []
    with open(path, encoding="utf-8", errors="replace") as f:
        reader = csv.DictReader(f)
        cols = {c.strip().lower(): c for c in (reader.fieldnames or [])}
        amt_c = _pick(cols, ["amount"])
        time_c = _pick(cols, ["time", "date"])
        ref_c = _pick(cols, ["ref", "auth"])
        for row in reader:
            amt = _to_number(row.get(amt_c)) if amt_c else None
            if not amt:
                continue
            raw_t = row.get(time_c) if time_c else None
            t = _parse_time_on_date(raw_t, business_date)
            records.append(PaymentRecord(amt, t or 0, str(raw_t or ""), machine_label,
                                         row.get(ref_c, "") if ref_c else "", shop_hint or None))
    return records


def load_manual_transfers_csv(path: str, source_label: str = "Direct Bank Transfer",
                              business_date: Optional[datetime] = None) -> list[PaymentRecord]:
    """Customer transfers straight to the owner's account: time, amount, ref/notes, shop."""
    records = []
    with open(path, encoding="utf-8", errors="replace") as f:
        reader = csv.DictReader(f)
        cols = {c.strip().lower(): c for c in (reader.fieldnames or [])}
        amt_c, time_c = _pick(cols, ["amount"]), _pick(cols, ["time", "date"])
        ref_c, shop_c = _pick(cols, ["ref", "notes"]), _pick(cols, ["shop"])
        for row in reader:
            amt = _to_number(row.get(amt_c)) if amt_c else None
            if not amt:
                continue
            raw_t = row.get(time_c) if time_c else None
            t = _parse_time_on_date(raw_t, business_date)
            records.append(PaymentRecord(amt, t or 0, str(raw_t or ""), source_label,
                                         row.get(ref_c, "") if ref_c else "",
                                         row.get(shop_c) if shop_c else None))
    return records


# ------------------------------------------------------------------ matching
CROSS_SHOP_PENALTY = 30  # score units (~minutes): nudges ties toward a source's home shop
TIMELESS_EXACT_PENALTY_MIN = 60   # an exact-amount bank row (no clock time) counts as "an hour away", so a
                                  # PhonePe/card payment minutes from the bill wins over it (and vice versa)


def match_bills_to_pool(bills, pool, amount_tol=AMOUNT_TOLERANCE, time_tol_min=TIME_TOLERANCE_MIN, home=None):
    """Unified greedy matcher. Every (bill, payment) pair within tolerance is
    scored; best pairs are assigned first. Exact amounts need no time window
    (a day-unique amount is strong evidence by itself); inexact amounts must
    also be within time_tol_min. Payments without a real clock time (bank
    statement rows) are matched on amount alone - their sequence index is
    never compared with a real timestamp."""
    pool = list(pool)
    used = [False] * len(pool)
    cands = []
    for bi, b in enumerate(bills):
        bt = b.time.timestamp()
        if b.card <= 0:   # cash-only bill: nothing digital to match (else it would soak up tiny payments)
            continue
        for pi, p in enumerate(pool):
            da = abs(p.amount - b.card)
            if da > tolerance_for(b.card, amount_tol):
                continue
            if isinstance(p.order_key, datetime):
                dt = abs(p.order_key.timestamp() - bt) / 60.0
                if da > 0.01 and dt > time_tol_min:
                    continue
            else:
                dt = TIMELESS_EXACT_PENALTY_MIN if da <= 0.01 else time_tol_min
            pen = CROSS_SHOP_PENALTY if home and home.get(p.source) and home[p.source] != b.shop else 0
            cands.append((da * 100000 + dt + pen, bi, pi))
    cands.sort(key=lambda x: x[0])
    match, done = {}, set()
    for _, bi, pi in cands:
        if bi in done or used[pi]:
            continue
        match[bi], used[pi] = pi, True
        done.add(bi)
    results = []
    for bi, b in enumerate(bills):
        if bi in match:
            p = pool[match[bi]]
            results.append(MatchResult(b, p, round(p.amount - b.card, 2), used=[p]))
        else:
            results.append(MatchResult(b, None, None))
    return results, [pool[i] for i in range(len(pool)) if not used[i]]


def _home_from_hints(pool, shops):
    """Source -> home shop, from the hint text on its records (PhonePe store
    name, bank account holder, card-machine hint), if it names a known shop."""
    from collections import Counter
    norm = {x.strip().lower(): x for x in shops}
    votes = {}
    for p in pool:
        if p.shop_hint and p.shop_hint.strip().lower() in norm:
            votes.setdefault(p.source, Counter())[norm[p.shop_hint.strip().lower()]] += 1
    return {src: c.most_common(1)[0][0] for src, c in votes.items()}


def _has_clock(p):
    return isinstance(p.order_key, datetime)


def _minutes(a, b):
    return abs(a - b).total_seconds() / 60.0


def _combine_bills_into_payments(results, remaining):
    """One payment covering several bills (customer bought from both shops, or
    two bills, and paid once). Uses only still-unmatched bills and unused payments.
    All candidate (payment, bills) groups are scored by time closeness and the
    best are taken first, so two equal payments are not mixed up."""
    unmatched = sorted((i for i, m in enumerate(results) if m.payment is None and m.bill.card > 0),
                       key=lambda i: results[i].bill.time)
    unmatched = [i for i in unmatched if results[i].bill.card >= COMBO_MIN_PART]
    cands = []
    for pi, p in enumerate(remaining):
        if p.amount < COMBO_MIN_PART:
            continue
        pool_b = [i for i in unmatched
                  if results[i].bill.card < p.amount + COMBO_AMOUNT_TOL
                  and (not _has_clock(p) or _minutes(p.order_key, results[i].bill.time) <= COMBO_PAY_WINDOW_MIN)]
        for size in (2, 3):
            if len(pool_b) < size or (size == 3 and len(pool_b) > 40) or len(pool_b) > 300:
                continue
            for combo in itertools.combinations(pool_b, size):
                if abs(sum(results[i].bill.card for i in combo) - p.amount) > COMBO_AMOUNT_TOL:
                    continue
                times = [results[i].bill.time for i in combo]
                spread = _minutes(max(times), min(times))
                if spread > COMBO_BILL_WINDOW_MIN:
                    continue
                gap = sum(_minutes(p.order_key, t) for t in times) if _has_clock(p) else 0.0
                cands.append((size, spread + gap, pi, combo))
    cands.sort(key=lambda c: (c[0], c[1]))
    used_bills, used_pay = set(), set()
    for size, _, pi, combo in cands:
        if pi in used_pay or any(i in used_bills for i in combo):
            continue
        p = remaining[pi]
        used_pay.add(pi)
        used_bills.update(combo)
        desc = ", ".join(f"{results[i].bill.shop} {results[i].bill.billno} (Rs.{results[i].bill.card:,.0f})" for i in combo)
        note = (f"One {p.source} payment of Rs.{p.amount:,.0f} (ref {p.ref or '-'}) covers {len(combo)} bills paid together: {desc}. "
                "Inferred from the amounts adding up exactly - confirm if unsure.")
        for i in combo:
            b = results[i].bill
            piece = PaymentRecord(b.card, p.order_key, p.time_str, p.source, p.ref, p.shop_hint)
            results[i] = MatchResult(b, piece, 0.0, "combined", note, [p])
    return [p for i, p in enumerate(remaining) if i not in used_pay]


def _claimable_alone(p, results):
    """True if some bill's recorded digital amount already fits this payment on its own
    (within the normal tolerance): such a payment is a bill's payment, not a spare 'part'."""
    for m in results:
        b = m.bill
        if b.card > 0 and abs(p.amount - b.card) <= tolerance_for(b.card) and (
                not _has_clock(p) or abs(p.order_key.timestamp() - b.time.timestamp()) / 60.0 <= TIME_TOLERANCE_MIN):
            return True
    return False


def _split_bills_across_payments(results, remaining):
    """One bill paid through several payments (part Federal UPI, part PhonePe, part card...).
    Never uses tiny payments (test transfers) or payments that fit another bill on their own."""
    todo = [i for i, m in enumerate(results) if m.payment is None and m.bill.card >= COMBO_MIN_PART]
    spare = {pi for pi, p in enumerate(remaining)
             if p.amount >= COMBO_MIN_PART and not _claimable_alone(p, results)}
    cands = []
    for i in todo:
        b = results[i].bill
        pool_p = [pi for pi in spare
                  if remaining[pi].amount < b.card
                  and (not _has_clock(remaining[pi]) or _minutes(remaining[pi].order_key, b.time) <= COMBO_PAY_WINDOW_MIN)]
        pool_p.sort()
        for size in (2, 3):
            if len(pool_p) < size or (size == 3 and len(pool_p) > 40) or len(pool_p) > 300:
                continue
            for combo in itertools.combinations(pool_p, size):
                if abs(sum(remaining[pi].amount for pi in combo) - b.card) > COMBO_AMOUNT_TOL:
                    continue
                gap = sum(_minutes(remaining[pi].order_key, b.time) for pi in combo if _has_clock(remaining[pi]))
                cands.append((size, gap, i, combo))
    cands.sort(key=lambda c: (c[0], c[1]))
    used, done = set(), set()
    for size, _, i, combo in cands:
        if i in done or any(pi in used for pi in combo):
            continue
        b = results[i].bill
        ps = [remaining[pi] for pi in combo]
        used.update(combo)
        done.add(i)
        sources = list(dict.fromkeys(p.source for p in ps))
        total = sum(p.amount for p in ps)
        clocks = [p.order_key for p in ps if _has_clock(p)]
        synth = PaymentRecord(total, min(clocks) if clocks else ps[0].order_key, ps[0].time_str,
                              SPLIT_SEP.join(sources), SPLIT_SEP.join(str(p.ref or "-") for p in ps),
                              ps[0].shop_hint, [(p.source, p.amount) for p in ps])
        desc = " + ".join(f"{p.source} Rs.{p.amount:,.0f}" for p in ps)
        note = (f"Bill of Rs.{b.card:,.0f} paid in {len(ps)} parts: {desc}. "
                "Inferred from the amounts adding up exactly - confirm if unsure.")
        results[i] = MatchResult(b, synth, round(total - b.card, 2), "split", note, ps)
    return [p for pi, p in enumerate(remaining) if pi not in used]


def _fits(p, b, tol=None):
    """Would payment p be accepted for bill b by the normal one-to-one rule?"""
    tol = tolerance_for(b.card) if tol is None else tol
    da = abs(p.amount - b.card)
    if b.card <= 0 or da > tol:
        return False
    return not _has_clock(p) or da <= 0.01 or _minutes(p.order_key, b.time) <= TIME_TOLERANCE_MIN


def _rebalance_for_splits(results, remaining):
    """A still-unmatched bill U equals (spare payment s) + (payment p that was loosely matched to
    another bill B, e.g. 165 given to a 164 bill). If B can take another unused payment q instead,
    free p and settle U as s + p. Only swaps approximate matches; never touches exact ones."""
    changed = True
    while changed:
        changed = False
        for ui, um in enumerate(results):
            U = um.bill
            if um.payment is not None or U.card < COMBO_MIN_PART:
                continue
            for si, s in enumerate(remaining):
                if s.amount < COMBO_MIN_PART or s.amount >= U.card:
                    continue
                if _has_clock(s) and _minutes(s.order_key, U.time) > COMBO_PAY_WINDOW_MIN:
                    continue
                need = U.card - s.amount
                for bi, bm in enumerate(results):
                    if bm.payment is None or bm.combo or bm.payment.parts or bi == ui:
                        continue
                    p = bm.payment
                    if abs(p.amount - need) > COMBO_AMOUNT_TOL or abs(p.amount - bm.bill.card) <= 0.01:
                        continue            # wrong amount, or an exact match we must not disturb
                    if _has_clock(p) and _minutes(p.order_key, U.time) > COMBO_PAY_WINDOW_MIN:
                        continue
                    qi = next((k for k, q in enumerate(remaining)
                               if k != si and q.amount >= COMBO_MIN_PART and _fits(q, bm.bill)), None)
                    if qi is None:
                        continue
                    q = remaining[qi]
                    synth = PaymentRecord(U.card, min([x.order_key for x in (s, p) if _has_clock(x)] or [s.order_key]),
                                          s.time_str, SPLIT_SEP.join(dict.fromkeys([s.source, p.source])),
                                          SPLIT_SEP.join([str(s.ref or "-"), str(p.ref or "-")]), s.shop_hint,
                                          [(s.source, s.amount), (p.source, p.amount)])
                    note = (f"Bill of Rs.{U.card:,.0f} paid in 2 parts: {s.source} Rs.{s.amount:,.0f} + {p.source} Rs.{p.amount:,.0f}. "
                            f"Rs.{p.amount:,.0f} was first tied to bill {bm.bill.billno}, which fits Rs.{q.amount:,.0f} better, so it was moved here. "
                            "Inferred from the amounts adding up exactly - confirm if unsure.")
                    results[ui] = MatchResult(U, synth, round(synth.amount - U.card, 2), "split", note, [s, p])
                    results[bi] = MatchResult(bm.bill, q, round(q.amount - bm.bill.card, 2), used=[q])
                    remaining = [x for k, x in enumerate(remaining) if k not in (si, qi)]
                    changed = True
                    break
                if changed:
                    break
            if changed:
                break
    return remaining


def reconcile_all_shops(bills_by_shop, pool, combos=True):
    """ONE joint pass over all shops' bills against the shared pool. Sources
    are lightly nudged toward their home shop (from hints, else from a first
    unbiased pass) so that identical small amounts that appear on both shops
    are not swapped arbitrarily - but a clear amount/time fit still wins, so
    genuine cross-shop routing is still found."""
    from collections import defaultdict
    combined = [(shop, b) for shop, bills in bills_by_shop.items() for b in bills]
    just = [b for _, b in combined]
    home = _home_from_hints(pool, list(bills_by_shop))
    if len(bills_by_shop) > 1:
        first, _ = match_bills_to_pool(just, pool, home=home)
        tot = defaultdict(lambda: defaultdict(float))
        for (shop, _), m in zip(combined, first):
            if m.payment:
                tot[m.payment.source][shop] += m.payment.amount
        for src, d in tot.items():
            home.setdefault(src, max(d, key=d.get))
    results, remaining = match_bills_to_pool(just, pool, home=home)
    if combos:
        remaining = _combine_bills_into_payments(results, remaining)
        remaining = _split_bills_across_payments(results, remaining)
        remaining = _rebalance_for_splits(results, remaining)
    by_shop = {shop: [] for shop in bills_by_shop}
    for (shop, _), m in zip(combined, results):
        by_shop[shop].append(m)
    return by_shop, remaining


def find_cash_suspects(unused, results_by_shop, max_show=4):
    """For each unassigned payment (>= Rs.20): bills that were rung up as CASH (card = 0) whose bill
    amount equals the payment. A customer who paid digitally but was billed as cash looks exactly
    like this. Payments with a clock time must be within 2 h of the bill."""
    cash = [m.bill for rs in results_by_shop.values() for m in rs if m.bill.card <= 0 and m.bill.net > 0]
    for p in unused:
        p.suspects = None
        if p.amount < COMBO_MIN_PART:
            continue
        found = []
        for b in cash:
            if abs(b.net - p.amount) > 0.5:
                continue
            gap = _minutes(p.order_key, b.time) if _has_clock(p) else None
            if gap is not None and gap > 120:
                continue
            found.append((gap if gap is not None else 9999, b, gap))
        found.sort(key=lambda x: (x[0], x[1].time))
        if found:
            p.suspects = [dict(shop=b.shop, bill=b.billno, time=b.time.strftime("%H:%M"), net=b.net,
                               gap_min=None if g is None else round(g), total=len(found))
                          for _, b, g in found[:max_show]]
    return unused
