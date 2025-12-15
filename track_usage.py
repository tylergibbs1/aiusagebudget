#!/usr/bin/env python3
"""
AI Usage Budget Tracker
Tracks Claude AI usage across Windows and macOS using bunx ccusage.
Updates a single Excel file with daily, weekly, and monthly TOTALS.
Tracks both OS separately and shows combined totals.
"""

import subprocess
import json
import platform
import sys
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Constants
EXCEL_FILE = "usage_tracker.xlsx"
SHEET_NAMES = ["Daily", "Weekly", "Monthly"]

# Column headers - now includes OS label column
HEADERS = [
    "OS", "Input Tokens", "Output Tokens", "Cache Creation Tokens",
    "Cache Read Tokens", "Total Tokens", "Cost ($)", "Last Updated"
]

# Row identifiers
OS_ROWS = {
    "Windows": 2,
    "macOS": 3,
    "Combined": 4
}


def detect_os():
    """Detect the current operating system."""
    system = platform.system()
    if system == "Darwin":
        return "macOS"
    elif system == "Windows":
        return "Windows"
    else:
        return system


def fetch_usage_data(period):
    """
    Fetch usage data from bunx ccusage for the given period.

    Args:
        period: One of 'daily', 'weekly', or 'monthly'

    Returns:
        Totals dict or None if error
    """
    try:
        print(f"Fetching {period} totals from ccusage...")
        result = subprocess.run(
            ['bunx', 'ccusage', period, '--json'],
            capture_output=True,
            text=True,
            check=True,
            encoding='utf-8'
        )

        data = json.loads(result.stdout)

        # Extract the totals
        totals = data.get('totals', {})

        if totals:
            print(f"✓ Fetched {period} totals")
        else:
            print(f"✗ No totals found for {period}")

        return totals

    except subprocess.CalledProcessError as e:
        print(f"✗ Error running bunx ccusage {period}: {e}")
        print(f"  Output: {e.stderr}")
        return None
    except json.JSONDecodeError as e:
        print(f"✗ Error parsing JSON for {period}: {e}")
        return None
    except Exception as e:
        print(f"✗ Unexpected error fetching {period} data: {e}")
        return None


def load_or_create_workbook():
    """Load existing workbook or create a new one with all sheets."""
    excel_path = Path(EXCEL_FILE)

    if excel_path.exists():
        print(f"Loading existing workbook: {EXCEL_FILE}")
        try:
            wb = load_workbook(EXCEL_FILE)
            return wb
        except Exception as e:
            print(f"✗ Error loading workbook: {e}")
            print("Creating new workbook instead...")

    print(f"Creating new workbook: {EXCEL_FILE}")
    wb = Workbook()

    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Create all three sheets with headers
    for sheet_name in SHEET_NAMES:
        ws = wb.create_sheet(sheet_name)

        # Write headers
        for col_idx, header in enumerate(HEADERS, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Initialize the OS label column
        ws.cell(row=2, column=1, value="Windows")
        ws.cell(row=3, column=1, value="macOS")
        ws.cell(row=4, column=1, value="Combined")
        ws.cell(row=4, column=1).font = Font(bold=True)

        # Freeze the header row
        ws.freeze_panes = "A2"

    return wb


def get_row_values(sheet, row_num):
    """Get numeric values from a row (excluding OS label and timestamp)."""
    values = []
    # Columns 2-7 are the numeric columns (tokens and cost)
    for col_idx in range(2, 8):
        cell_value = sheet.cell(row=row_num, column=col_idx).value
        values.append(cell_value if cell_value else 0)
    return values


def update_sheet_with_totals(sheet, totals, os_name):
    """Update sheet with totals for the current OS and recalculate combined."""
    if not totals:
        return False

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Determine which row to update based on OS
    row_num = OS_ROWS.get(os_name)
    if not row_num:
        print(f"✗ Unknown OS: {os_name}")
        return False

    # Row data for current OS (starting from column 2, after OS label)
    row_data = [
        totals.get('inputTokens', 0),
        totals.get('outputTokens', 0),
        totals.get('cacheCreationTokens', 0),
        totals.get('cacheReadTokens', 0),
        totals.get('totalTokens', 0),
        round(totals.get('totalCost', 0), 2),
        timestamp
    ]

    # Update the OS-specific row (columns 2-8)
    for col_idx, value in enumerate(row_data, 2):
        sheet.cell(row=row_num, column=col_idx, value=value)

    # Recalculate combined totals (sum of Windows and macOS)
    windows_values = get_row_values(sheet, OS_ROWS["Windows"])
    macos_values = get_row_values(sheet, OS_ROWS["macOS"])

    combined_values = [w + m for w, m in zip(windows_values, macos_values)]
    combined_values.append(timestamp)  # Add timestamp

    # Update combined row (columns 2-8)
    for col_idx, value in enumerate(combined_values, 2):
        cell = sheet.cell(row=OS_ROWS["Combined"], column=col_idx, value=value)
        # Round the cost column
        if col_idx == 7:  # Cost column
            cell.value = round(value, 2)

    return True


def format_sheet(sheet):
    """Apply formatting to the sheet."""
    # Auto-adjust column widths
    for column in sheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)

        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass

        adjusted_width = min(max_length + 2, 50)  # Cap at 50
        sheet.column_dimensions[column_letter].width = adjusted_width

    # Right-align numeric columns (columns 2-7: tokens and cost)
    for row in sheet.iter_rows(min_row=2):
        for idx, cell in enumerate(row):
            # Right-align numeric columns (tokens, cost)
            if 1 <= idx <= 6:  # Columns 2-7 (0-indexed in row array)
                cell.alignment = Alignment(horizontal="right", vertical="center")

    # Add light yellow background to Combined row
    for col_idx in range(1, 9):
        cell = sheet.cell(row=OS_ROWS["Combined"], column=col_idx)
        cell.fill = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
        if col_idx == 1:  # OS label column
            cell.font = Font(bold=True)


def display_sheet_totals(sheet, sheet_name):
    """Display totals from a sheet."""
    print(f"\n{sheet_name} Totals:")
    print("-" * 60)

    for os_label, row_num in OS_ROWS.items():
        total_tokens = sheet.cell(row=row_num, column=6).value or 0
        total_cost = sheet.cell(row=row_num, column=7).value or 0

        indent = "  " if os_label != "Combined" else ""
        bold = "**" if os_label == "Combined" else ""

        print(f"{indent}{bold}{os_label}:{bold}")
        print(f"{indent}  Total Tokens: {total_tokens:,}")
        print(f"{indent}  Total Cost: ${total_cost:.2f}")


def main():
    """Main execution flow."""
    print("\n" + "="*60)
    print("AI Usage Budget Tracker")
    print("="*60 + "\n")

    # Detect OS
    os_name = detect_os()
    print(f"Operating System: {os_name}\n")

    # Fetch data from all three periods
    print("Fetching usage totals from ccusage...\n")
    daily_totals = fetch_usage_data('daily')
    weekly_totals = fetch_usage_data('weekly')
    monthly_totals = fetch_usage_data('monthly')

    # Check if we got any data
    if daily_totals is None and weekly_totals is None and monthly_totals is None:
        print("\n✗ Failed to fetch any data. Please check your bunx ccusage installation.")
        sys.exit(1)

    # Load or create workbook
    print("\nUpdating Excel file...\n")
    wb = load_or_create_workbook()

    # Update each sheet
    updated_sheets = []

    if daily_totals is not None:
        daily_sheet = wb["Daily"]
        if update_sheet_with_totals(daily_sheet, daily_totals, os_name):
            format_sheet(daily_sheet)
            updated_sheets.append("Daily")
            print(f"✓ Updated Daily sheet ({os_name} + Combined)")

    if weekly_totals is not None:
        weekly_sheet = wb["Weekly"]
        if update_sheet_with_totals(weekly_sheet, weekly_totals, os_name):
            format_sheet(weekly_sheet)
            updated_sheets.append("Weekly")
            print(f"✓ Updated Weekly sheet ({os_name} + Combined)")

    if monthly_totals is not None:
        monthly_sheet = wb["Monthly"]
        if update_sheet_with_totals(monthly_sheet, monthly_totals, os_name):
            format_sheet(monthly_sheet)
            updated_sheets.append("Monthly")
            print(f"✓ Updated Monthly sheet ({os_name} + Combined)")

    if not updated_sheets:
        print("\n✗ No sheets were updated.")
        sys.exit(1)

    # Save workbook
    try:
        wb.save(EXCEL_FILE)
        print(f"\n✓ Saved to {EXCEL_FILE}")

        # Print summary
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"\nLast updated: {timestamp}")

        # Display totals from each sheet
        print("\n" + "="*60)
        print("Current Totals (All Time):")
        print("="*60)

        if daily_totals:
            display_sheet_totals(wb["Daily"], "Daily")

        if weekly_totals:
            display_sheet_totals(wb["Weekly"], "Weekly")

        if monthly_totals:
            display_sheet_totals(wb["Monthly"], "Monthly")

        print("\n" + "="*60)
        print("Done! You can now commit and push to GitHub:")
        print(f"  git add {EXCEL_FILE}")
        print(f'  git commit -m "Update usage totals from {os_name}"')
        print("  git push")
        print("="*60 + "\n")

    except Exception as e:
        print(f"\n✗ Error saving workbook: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
