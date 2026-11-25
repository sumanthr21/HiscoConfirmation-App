# HISCO Confirmation - Daily Report app

Adds each agent's EOD rows for one day to the team report
(`Hisco_OS_Confirmation ... .xlsx`) and saves the result back to a private GitHub repo.

## What it updates
| EOD task            | Report sheet   |
|---------------------|----------------|
| Confirmation        | CONF           |
| Tracking            | Tracking       |
| Certificates        | Certificate    |
| Moved to FAB Conf   | Moved to FAB   |

It fills Month, Date, Week and Name itself, keeps the % price formulas on CONF and Tracking,
and sets the report date (cell B6) on the month sheet (for example Nov'26).
Daily, Weekly and the month sheet are formula-driven, so they update when Excel opens the file.

It does **not** touch: New Tasks (open-order reports, Safety Documents), Daily Allocation vs Completed,
Price Issue (pivot - right-click Refresh in Excel).

## Secrets (Streamlit -> Advanced settings -> Secrets)
```
APP_PASSWORD = "choose-a-strong-password"

[GITHUB]
token = "github_pat_..."
repo = "yourname/hisco-confirmation-data"
path = "report.xlsx"
branch = "main"
```

## Daily use
1. Select all agents' EOD workbooks together.
2. Check the date and the agent shown for each file.
3. Click **Update report**, read the messages, wait for "Saved back to GitHub".
4. Download the report and share it.

## Safety
* A day cannot be loaded twice: if the report already has rows for that agent, sheet and date,
  nothing is changed.
* Charts, comments and pivot tables are kept (the file is edited at XML level, not re-created).

## New month
Before the first load of a new month, add the new month sheet (copy the previous one, rename it,
for example Dec'26) and its Daily and Weekly blocks, and put that report in the private repo.
