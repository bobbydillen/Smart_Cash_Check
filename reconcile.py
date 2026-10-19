#!/usr/bin/env python3
"""Command-line version.  Example:

python reconcile.py --date 2026-08-25 \\
  --pos "Smart Mart=mart.xlsx" --pos "Smart Fashions=fashions.xlsx" \\
  --bank-pdf federal.pdf --phonepe phonepe.csv \\
  --card-machine "ICICI Machine 1=icici1.csv" \\
  --card-machine "HDFC Machine (Mart)=hdfc.csv" \\
  --out report.xlsx

Each --card-machine is "Label=file.csv" (optionally "Label=file.csv=Shop Hint").
Add as many as you have - a new machine for either shop is just one more flag."""
import argparse
from pipeline import run_reconciliation


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--date", required=True)
    ap.add_argument("--pos", action="append", required=True, metavar="Shop=file.xlsx")
    ap.add_argument("--bank-pdf", action="append", default=[])
    ap.add_argument("--phonepe", action="append", default=[])
    ap.add_argument("--card-machine", action="append", default=[], metavar="Label=file.csv[=Shop]")
    ap.add_argument("--transfers", action="append", default=[])
    ap.add_argument("--out", default=None)
    a = ap.parse_args()
    cms = []
    for spec in a.card_machine:
        p = spec.split("=")
        cms.append((p[0], p[1], p[2] if len(p) > 2 else None))
    s = run_reconciliation(
        a.date, dict(x.split("=", 1) for x in a.pos),
        [("Bank UPI", p) for p in a.bank_pdf], [("PhonePe", p) for p in a.phonepe], cms,
        [("Direct Transfer", p) for p in a.transfers], a.out or f"Reconciliation_{a.date}.xlsx")
    print(f"\nNet sales Rs.{s['total_net']:,.0f} | digital Rs.{s['total_digital']:,.0f} | implied cash Rs.{s['implied_cash']:,.0f}")
    print(f"{len(s['issues'])} items to review; {len(s['unassigned'])} unassigned payments")


if __name__ == "__main__":
    main()
