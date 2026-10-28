# AI Usage Budget Tracker

A Python script that tracks Claude usage across several machines. It pulls per-period usage
from [`ccusage`](https://github.com/ryoppippi/ccusage) and keeps one Excel workbook, shared
through Git, with daily, weekly and monthly token and cost totals per machine, plus a
comparison of API-equivalent cost against a $200/month subscription.

## Quick start

```bash
pip install -r requirements.txt
python track_usage.py
```

Requires Python 3.7+ and `bunx` with access to `bunx ccusage`.

## How it works

Each run:

1. Identifies the machine as `<hostname>-<hash>` (a short hash of hostname + MAC address).
2. Runs `bunx ccusage daily|weekly|monthly --json`.
3. Inserts or updates one row per period for this machine in `usage_tracker.xlsx`,
   leaving other machines' rows alone, and recalculates the `TOTALS` row.
4. Rebuilds the **Machine Totals** sheet from the Daily sheet and prints a summary.

Commit the workbook after each run and pull before running on another machine, so every
machine's rows end up in the same file:

```bash
git pull
python track_usage.py
git add usage_tracker.xlsx
git commit -m "Update usage totals"
git push
```

## Workbook layout

| Sheet | Columns |
|---|---|
| Daily, Weekly | Date, Machine, Tokens, Cost ($) |
| Monthly | Date, Machine, Tokens, API Cost ($), Subscription ($), Savings ($) |
| Machine Totals | Machine, Tokens, Cost ($) |

Each time sheet ends with a `TOTALS` row. On the Monthly sheet, the $200 subscription for a
month is split evenly across that month's machine rows, and Savings is API cost minus that
share, so the totals row compares total API cost with $200 × number of months. Change
`MONTHLY_SUBSCRIPTION` in `track_usage.py` if your plan differs.

`dashboard.html` is a Chart.js dashboard (cost by machine, machine breakdown) with a snapshot
of the tracked data embedded in the file; the script does not regenerate it.

## Troubleshooting

**`bunx` command not found**: make sure `bunx` is installed and on your PATH (`bunx --version`).

**`ccusage` not found**: check that `bunx ccusage --help` works.

**"Cannot save - usage_tracker.xlsx is open in another program"**: close the file in Excel and
run again.

**Git merge conflicts on `usage_tracker.xlsx`**: accept either version, run the script again on
each machine so its rows are rewritten, then commit the result.

## License

This project is provided as-is for personal use.
