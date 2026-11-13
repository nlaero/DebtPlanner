# Install Debt Payoff Planner

These are folders in one repository, not separate branches. Clone once, then use the command for your system.

## Debian and Ubuntu

```sh
git clone https://github.com/nlaero/DebtPlanner.git
cd DebtPlanner
sh packaging/debian/install.sh
```

Plans are saved in `~/.config/debt-payoff-planner`.

## Flatpak

```sh
git clone https://github.com/nlaero/DebtPlanner.git
cd DebtPlanner
sh packaging/flatpak/install.sh
```

The first run downloads the Flatpak runtime. Plans are saved inside the Flatpak data folder.

## Windows

Download `installer/Version 2.x/DebtPayoffPlannerSetup-2.03.exe` from this repository and run it.
