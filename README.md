# AI Buildout Portfolio Recomp Cockpit

Local-first browser app for portfolio x-ray, AI buildout exposure analysis, strategy-template comparison, and simulated recomp planning.

## Run

On a new PC, install Node.js first if needed, then double-click:

```text
setup-portfolio-cockpit.bat
```

The setup script checks for Node.js, installs dependencies, and verifies the app builds. This app needs Node.js 20.19 or newer, or Node.js 22.12 or newer.

After setup, double-click `start-portfolio-cockpit.bat`, or run:

```powershell
cmd /c npm.cmd run dev -- --host 127.0.0.1 --port 3002
```

Then open:

```text
http://127.0.0.1:3002
```

## Local Files

Brokerage exports are intentionally ignored by git:

```text
*.csv
*.xlsx
*.xls
```

Use the Import & Snapshots tab to select one or more account files. The browser cannot directly browse folders without user file selection, so snapshots are created from files you pick in the upload control.
