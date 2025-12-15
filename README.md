# AI Usage Budget Tracker

A Python script to track Claude AI usage across **both** Windows and macOS. Automatically fetches total usage data from `bunx ccusage` and maintains an Excel file with overall daily, weekly, and monthly totals - tracked separately per OS and combined.

## Features

- Cross-platform support (Windows & macOS)
- Single Excel file shared via GitHub
- Three sheets: Daily, Weekly, Monthly
- **Separate tracking** for Windows and macOS
- **Automatic combined totals** (Windows + macOS)
- Simple, clean output with just the numbers you need
- Updates automatically each time you run it

## Prerequisites

- Python 3.7 or higher
- `bunx` installed and configured
- Access to `bunx ccusage` command

## Installation

1. Clone this repository or download the files

2. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the Script

Simply run:
```bash
python track_usage.py
```

The script will:
1. Fetch overall totals from `bunx ccusage daily/weekly/monthly`
2. Update the current OS row in the Excel file
3. Recalculate the combined totals (Windows + macOS)
4. Display a summary

### Example Output

```
============================================================
AI Usage Budget Tracker
============================================================

Operating System: Windows

Fetching usage totals from ccusage...

Fetching daily totals from ccusage...
✓ Fetched daily totals
Fetching weekly totals from ccusage...
✓ Fetched weekly totals
Fetching monthly totals from ccusage...
✓ Fetched monthly totals

Updating Excel file...

✓ Updated Daily sheet (Windows + Combined)
✓ Updated Weekly sheet (Windows + Combined)
✓ Updated Monthly sheet (Windows + Combined)

✓ Saved to usage_tracker.xlsx

Last updated: 2025-12-14 14:23:45

============================================================
Current Totals (All Time):
============================================================

Daily Totals:
------------------------------------------------------------
  Windows:
    Total Tokens: 580,255,768
    Total Cost: $400.45

  macOS:
    Total Tokens: 600,255,767
    Total Cost: $418.46

  **Combined:**
    Total Tokens: 1,180,511,535
    Total Cost: $818.91

Weekly Totals:
------------------------------------------------------------
  Windows:
    Total Tokens: 350,120,000
    Total Cost: $245.32

  macOS:
    Total Tokens: 400,114,123
    Total Cost: $278.13

  **Combined:**
    Total Tokens: 750,234,123
    Total Cost: $523.45

Monthly Totals:
------------------------------------------------------------
  Windows:
    Total Tokens: 450,228,394
    Total Cost: $314.41

  macOS:
    Total Tokens: 500,228,395
    Total Cost: $349.41

  **Combined:**
    Total Tokens: 950,456,789
    Total Cost: $663.82

============================================================
Done! You can now commit and push to GitHub:
  git add usage_tracker.xlsx
  git commit -m "Update usage totals from Windows"
  git push
============================================================
```

## Excel File Structure

Each sheet (Daily, Weekly, Monthly) contains **three rows**:

### Row Structure:
1. **Windows** - Usage from Windows machine
2. **macOS** - Usage from macOS machine
3. **Combined** - Sum of Windows + macOS (highlighted in yellow)

### Columns (same for all sheets):
| OS | Input Tokens | Output Tokens | Cache Creation Tokens | Cache Read Tokens | Total Tokens | Cost ($) | Last Updated |

### Example Daily Sheet:
| OS | Input Tokens | Output Tokens | Cache Creation Tokens | Cache Read Tokens | Total Tokens | Cost ($) | Last Updated |
|----|--------------|---------------|-----------------------|-------------------|--------------|----------|--------------|
| Windows | 750,000 | 710,000 | 42,500,000 | 546,000,000 | 590,000,000 | 412.34 | 2025-12-14 10:30:15 |
| macOS | 749,984 | 710,937 | 42,669,823 | 546,420,791 | 590,511,535 | 406.57 | 2025-12-14 14:23:45 |
| **Combined** | **1,499,984** | **1,420,937** | **85,169,823** | **1,092,420,791** | **1,180,511,535** | **818.91** | 2025-12-14 14:23:45 |

## How It Works

### Cross-Platform Tracking

1. **On Windows**:
   - Script fetches Windows usage from `bunx ccusage`
   - Updates the "Windows" row
   - Recalculates "Combined" row (Windows + macOS)

2. **On macOS**:
   - Script fetches macOS usage from `bunx ccusage`
   - Updates the "macOS" row
   - Recalculates "Combined" row (Windows + macOS)

3. **Combined Totals**:
   - Automatically calculated by summing Windows + macOS rows
   - Shows your total usage across both systems
   - Highlighted in yellow for easy visibility

### GitHub Workflow

Since both OS share the same Excel file via GitHub:

1. Run script on Windows → Updates Windows row
2. Commit and push to GitHub
3. Pull on macOS → Get latest Windows data
4. Run script on macOS → Updates macOS row, Combined is recalculated with both
5. Commit and push back to GitHub

## GitHub Workflow

### Initial Setup

```bash
# Initialize git if not already done
git init

# Add files
git add track_usage.py requirements.txt README.md .gitignore usage_tracker.xlsx

# Commit
git commit -m "Initial commit: AI usage tracker"

# Push to GitHub
git remote add origin YOUR_GITHUB_REPO_URL
git push -u origin main
```

### Regular Updates

**On Windows:**
```bash
# Update Windows usage
python track_usage.py

# Commit and push
git add usage_tracker.xlsx
git commit -m "Update usage totals from Windows"
git push
```

**On macOS:**
```bash
# Pull latest data (includes Windows updates)
git pull

# Update macOS usage
python track_usage.py

# Commit and push
git add usage_tracker.xlsx
git commit -m "Update usage totals from macOS"
git push
```

**Back on Windows:**
```bash
# Pull latest data (includes macOS updates)
git pull

# Run to update Windows again
python track_usage.py

# Commit and push
git add usage_tracker.xlsx
git commit -m "Update usage totals from Windows"
git push
```

The Combined row always shows the sum of both systems!

## What Gets Tracked

The script fetches the "totals" section from each ccusage command:

- **Daily totals**: Aggregate of all your usage on that OS
- **Weekly totals**: Aggregate of all your weekly usage on that OS
- **Monthly totals**: Aggregate of all your monthly usage on that OS
- **Combined**: Sum of Windows + macOS for complete picture

## Troubleshooting

### "bunx command not found"

Make sure `bunx` is installed and in your PATH. Test by running:
```bash
bunx --version
```

### "ccusage not found"

Make sure you have access to the `ccusage` command via bunx:
```bash
bunx ccusage --help
```

### Permission errors on Excel file

If the Excel file is open in another program, close it before running the script.

### Git merge conflicts

If you get merge conflicts on `usage_tracker.xlsx`:
1. Accept one version (doesn't matter which)
2. Run the script on each OS again
3. Both rows will be updated correctly
4. Commit the resolved version

### One OS showing zeros

This is normal if you haven't run the script on that OS yet. Just run it on the other machine and the row will populate.

## Notes

- The Excel file is small and safe to commit to GitHub
- Each OS has its own row, preserving separate tracking
- Combined row is always the sum of both OS rows
- Each run updates the current OS and recalculates Combined
- Simple and clean - per-OS breakdown plus total!

## License

This project is provided as-is for personal use.
