# Daily Sales vs Payments Reconciliation

Upload the day's files -> get a transaction-level reconciliation (Excel + on-screen summary).

## Start the app
1. Install Python 3.10+ (python.org). On Windows tick "Add Python to PATH".
2. Double-click **start_web.bat** (Windows) or run `python app.py`. First run installs the libraries.
   A browser tab opens at http://127.0.0.1:5000.
3. Upload files, press **Run reconciliation**, then **Download Excel report**.
(The older Streamlit version is still there: `start_app.bat`. To host online see **DEPLOY.md**.)

## Daily files
| What | Format |
|---|---|
| Smart Mart sales | Excel export from your billing software |
| Smart Fashions sales | Excel export |
| Bank statement | Federal Bank PDF (each bank layout needs its own reader) |
| PhonePe / EDC settlement | CSV |
| **Card machines** | Just type the amounts from the slip (one per line). Time/ref are optional; CSV upload also works. Any number, either shop. |
| Direct transfers | Type into the optional table (time, amount, note) |

## Cards photos
Read the slip's transactions (time, amount, auth code) into the card-machine table.
Tip: send the photo to Claude and ask it to produce the CSV - then upload it.

## How it decides
* Every payment (bank credit, PhonePe, card swipe) goes into ONE pool; every shop's bills are matched
  against it together. Nothing is pre-assigned to a shop, so a new card machine just works, and
  cross-shop payments (a Mart sale paid on the Fashions QR) are found automatically.
* Split payments (part digital, part cash) are NOT flagged: only the recorded card/digital amount is checked.
* Combined payments: one payment covering several bills (e.g. a customer buying from both shops) and one bill paid through
  several payments (e.g. Federal + PhonePe) are found when the amounts add up exactly; shown as "Matched - Combined / Split"
  with Medium confidence so you can verify.
* Flags: **Unmatched** (no payment found), **Excess digital** (digital recorded above the bill - typo or
  under-billing), **Cross-shop source** (informational), **Unassigned payments** (money that fits no bill).
* "Implied cash" = net sales - digital matched. It is NOT a cash count.

## Limits (be aware)
* Bank statements have only a date, so bank credits are matched on amount alone; identical small
  amounts on both shops can be paired interchangeably (totals are right, a specific reference may differ).
* Amounts within Rs.15 and, for non-exact amounts, 4 hours are treated as the same payment
  (change in `reconcile_lib.py` top constants).
* Optional password: set environment variable `APP_PASSWORD` before starting. Keep the app on a PC you
  control - it handles bank statements. Do not expose it to the internet without a login.

## Files
`app.py` Flask web app (+ `web/index.html`) - `streamlit_app.py` old Streamlit page - `pipeline.py` runs everything - `reconcile_lib.py` loaders + matcher -
`report.py` Excel builder - `reconcile.py` command-line version - `templates/` sample CSVs.
