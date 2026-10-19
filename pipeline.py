"""pipeline.py - one function that runs the whole reconciliation; used by both
the command-line tool (reconcile.py) and the web app (app.py)."""
from __future__ import annotations
from datetime import datetime

from reconcile_lib import (load_pos_bills, load_federal_bank_pdf, load_phonepe_csv,
                           load_card_machine_csv, load_manual_transfers_csv,
                           reconcile_all_shops, ANOMALY_TOLERANCE, portions)
from report import build_workbook, classify, source_shop_majority


def _unique(label, taken):
    base, n = label, 2
    while label in taken:
        label, n = f"{base} ({n})", n + 1
    return label


def run_reconciliation(date_label, pos_files, bank_pdfs=(), phonepe_files=(), card_machines=(),
                       transfer_files=(), out_path="reconciliation.xlsx", log=print):
    """
    pos_files:     {"Smart Mart": path.xlsx, "Smart Fashions": path.xlsx}
    bank_pdfs:     [(label, path.pdf[, shop_hint]), ...]   <- one per bank account
    phonepe_files: [(label, path.csv), ...]
    card_machines: [(label, path.csv, shop_hint_or_None), ...]   <- as many as you have
    transfer_files:[(label, path.csv), ...]
    """
    try:
        day = datetime.strptime(date_label, "%Y-%m-%d")
    except ValueError:
        day = None

    bills_by_shop = {}
    for shop, path in pos_files.items():
        bills_by_shop[shop] = load_pos_bills(path, shop)
        log(f"{shop}: {len(bills_by_shop[shop])} bills, net Rs.{sum(b.net for b in bills_by_shop[shop]):,.0f}")

    sources = {}
    for label, path, *rest in bank_pdfs:  # (label, path[, shop_hint]) - one entry per bank account
        label = _unique(label, sources)
        sources[label] = load_federal_bank_pdf(path, label, rest[0] if rest else None)
        log(f"{label}: {len(sources[label])} deposits, Rs.{sum(p.amount for p in sources[label]):,.0f}")
    for label, path in phonepe_files:
        label = _unique(label, sources)
        sources[label] = load_phonepe_csv(path, label)
        log(f"{label}: {len(sources[label])} transactions, Rs.{sum(p.amount for p in sources[label]):,.0f}")
    for label, path, hint in card_machines:
        label = _unique(label, sources)
        sources[label] = load_card_machine_csv(path, label, hint, day)
        log(f"{label}: {len(sources[label])} swipes, Rs.{sum(p.amount for p in sources[label]):,.0f}")
    for label, path in transfer_files:
        label = _unique(label, sources)
        sources[label] = load_manual_transfers_csv(path, label, day)
        log(f"{label}: {len(sources[label])} transfers, Rs.{sum(p.amount for p in sources[label]):,.0f}")

    if not bills_by_shop:
        raise ValueError("Upload at least one shop's billing export.")
    if not sources:
        raise ValueError("Upload at least one payment source (bank PDF, PhonePe CSV, card machine...).")

    pool = [p for recs in sources.values() for p in recs]
    results_by_shop, unused = reconcile_all_shops(bills_by_shop, pool)
    shops = list(bills_by_shop)
    build_workbook(date_label, shops, results_by_shop, unused, sources, out_path)

    majority = source_shop_majority(results_by_shop)
    summary = {"date": date_label, "shops": {}, "unassigned": unused, "issues": [], "sources": {}}
    for s in shops:
        res = results_by_shop[s]
        net = sum(m.bill.net for m in res)
        dig = sum(m.payment.amount for m in res if m.payment)
        summary["shops"][s] = dict(bills=len(res), net=net, recorded_card=sum(m.bill.card for m in res),
                                   confirmed_digital=dig, implied_cash=net - dig)
        for m in res:
            status, conf, note, _ = classify(m, majority)
            if status != "Matched":
                summary["issues"].append(dict(shop=s, bill=m.bill.billno, time=m.bill.time.strftime("%H:%M"),
                                              net=m.bill.net, recorded_digital=m.bill.card,
                                              source=m.payment.source if m.payment else "",
                                              status=status, confidence=conf, note=note))
    for n, recs in sources.items():
        per = {s: sum(a for m in results_by_shop[s] for src, a in portions(m) if src == n) for s in shops}
        summary["sources"][n] = dict(count=len(recs), total=sum(p.amount for p in recs), per_shop=per)
    summary["total_net"] = sum(v["net"] for v in summary["shops"].values())
    summary["total_digital"] = sum(v["confirmed_digital"] for v in summary["shops"].values())
    summary["implied_cash"] = summary["total_net"] - summary["total_digital"]
    return summary
