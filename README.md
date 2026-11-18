# Debt Payoff Planner

Desktop planner for payoff dates, extra payments, and interest. Version 2.03.

Your saved plan stays on your computer. It is not included in this repository.

## Install

### Windows

Download and run [DebtPayoffPlannerSetup-2.03.exe](installer/Version%202.x/DebtPayoffPlannerSetup-2.03.exe).

It installs Debt Payoff Planner for Windows and replaces an older copy of this program. Your last saved plan is kept.

### Debian and Ubuntu

```sh
git clone https://github.com/nlaero/DebtPlanner.git
cd DebtPlanner
sh packaging/debian/install.sh
```

Or, without cloning the whole project:

```sh
curl -L -o debt-payoff-planner.deb https://github.com/nlaero/DebtPlanner/raw/main/packaging/debian/debt-payoff-planner_2.03-1_all.deb
sudo apt install ./debt-payoff-planner.deb
```

Start it from the application menu, or run `debt-payoff-planner`. Plans are saved in `~/.config/debt-payoff-planner`.

### Flatpak

```sh
git clone https://github.com/nlaero/DebtPlanner.git
cd DebtPlanner
sh packaging/flatpak/install.sh
```

The first run downloads the Flatpak runtime, so the computer needs a network connection. Debian and Ubuntu also need `flatpak` and `flatpak-builder`:

```sh
sudo apt install flatpak flatpak-builder
```

Start it with `flatpak run io.github.nlaero.DebtPayoffPlanner`.

### From the Python source

Use this when you are changing the program. Python 3.9 or newer is required, plus Tk and CustomTkinter. Excel export also needs openpyxl.

```sh
pip install customtkinter openpyxl
python debt_payoff_app.py
```

On Linux, install Tk as well: `sudo apt install python3-tk python3-pil`.

## How to use

The first time the program opens, choose **View Sample Plan** or **Create New Plan**. The sample is view only. Create your own plan, or use **File → Save As**, when you want to edit.

Help → How to Use shows this same short guide inside the program.

1. **Settings.** Set how often you are paid, the payday your balances belong to, and the extra you can add each payday. Yearly take-home pay is optional. It feeds the income and debt-to-income figures on Overview.
2. **Debts.** Add each balance, rate, minimum payment, and due day. Current Balance is what remains on the payday now in effect, after that day's minimum, interest, and any extra entered on the Pay Schedule. Rank 1 gets the extra payment first. Drag a row, or use Move Up and Move Down, to change that order. When a debt is paid off, its minimum joins the extra for the debts still open.
3. **Pay Schedule.** Each row is one payday. Double-click a row to record an extra for that day. Archive Completed Pay Periods files past paydays. Those rows stay under Archives and do not change later.
4. **Overview, Charts, and Interest.** Overview is the snapshot: balances, payoff dates, and debt-to-income. Charts and Interest compare extra payments in rank order with minimum payments only.
5. **Saving.** File → Save stores the plan and opens it next time.

Interest each payday is the balance times the annual rate, divided by the number of pay periods in a year.

## Tabs

- **Overview** — balances, payoff dates, income, and debt-to-income.
- **Pay Schedule** — one row per payday, with extra payments.
- **Debts** — add, edit, reorder, and update balances.
- **Charts** — balances over time. Click a name in the key to show or hide it.
- **Interest** — interest with minimum payments only, compared with extra payments.
- **Archives** — completed pay periods, kept as they were.
- **Settings** — pay frequency, dates, extra per payday, and income.

## File menu

- **New Blank** — start an empty plan.
- **Open, Save, Save As** — plans are JSON files.
- **Export Schedule to Excel** and **Export Archives to Excel**.
- **Undo** and **Redo**.
- **Exit**.
