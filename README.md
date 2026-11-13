# Debt Payoff Planner

Standalone desktop app for planning debt payoff with per-period extras, snowball additional payments, income and DTI on Overview, and year archives.

## Requirements
- Python 3.8 or newer
- Tkinter (included with most Python installs)
- openpyxl (only needed for Excel export)

### Install openpyxl (optional, for Excel export)
```
pip install openpyxl
```

On Linux you may also need:
```
sudo apt install python3-tk
```

## How to Run
1. Save `debt_payoff_app.py` somewhere convenient.
2. Open a terminal / command prompt in that folder.
3. Run:
   ```
   python debt_payoff_app.py
   ```
   (or `python3` on some systems)

Close and run that command again after you change the Python file. The last plan you saved or opened is reopened automatically. If none is open, you are asked to open a JSON file or create a new plan. The tutorial runs once on first launch; you can skip it on the first slide.

## Files
- `debt_payoff_app.py` — the program
- `DebtPayoff_Sample.json` — sample plan (File → Open or File → Load Sample Data)
- `DebtPayoff_Blank.json` — empty starter you can share with others

## Tabs
1. **Overview** — yearly and monthly income, front-end and back-end DTI, monthly and per-pay minimums, remaining balances by debt type, and an estimated debt-free date per type.
2. **Pay Schedule** — running balances each pay period. Double-click a cell to assign extras. Paid-off cells show a check. The Additional Per Pay box on the tab strip is the snowball recommendation (starting additional plus freed minimums).
3. **Debts** — add, edit, delete, and update balances. Fields include Debt Type and Payment Frequency. Payoff Rank is the order extra goes; drag a row or use Move Up / Move Down. A new debt is last.
4. **Archives** — completed pay periods, by year. Archived rows stay as they were, including debts you later delete.
5. **Settings** — pay frequency, First Pay Date, Balances As Of, Starting Additional Per Pay, projection length, and yearly net salary, spouse salary, and other income. A completed field is applied when you leave it. There is no Apply button.

## File menu
- New Blank — clean slate
- Load Sample Data — restore the sample debts
- Open / Save / Save As — portable JSON
- Export Schedule / Archives to Excel
- Undo / Redo
- Exit

## How the numbers work
- Interest is: balance × (annual rate ÷ pay periods per year).
- Each debt’s minimum is taken on the pay period that contains its due date. Payment Frequency on the debt is how often that minimum repeats.
- Balances As Of is the payday your Debts-tab current balances belong to. First Pay Date is only how far back the grid starts; the program may replace it with the first payday of that year.
- Period numbers restart at 1 each calendar year.
- Extra on the Pay Schedule is typed per debt and period. Overview’s debt-free dates ignore those typed extras. They use minimums, interest, Additional Per Pay, and Payoff Rank, and they add a paid-off debt’s minimum into Additional Per Pay on later paydays.
- Front-end DTI is mortgage minimums only (N/A if there is no mortgage). Back-end DTI is all active monthly minimums. Both need income; otherwise they show a dash.
- A paid-off debt stays on Debts until you delete it. It drops off the live pay schedule when you archive and when you open the plan. Give it a balance again and it returns on the current pay period, last in Payoff Rank.

In-app notes are also on Settings. Help → How to Use is a short recap. Help → Tutorial walks through the screens with pointed boxes.
