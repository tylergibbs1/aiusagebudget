#!/usr/bin/env python3
"""
AI Usage Budget Tracker
Tracks Claude AI usage across multiple machines using bunx ccusage.
Simple format: Daily and Monthly views with totals.
"""

import subprocess
import json
import hashlib
import platform
import sys
import uuid

from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Constants
EXCEL_FILE = "usage_tracker.xlsx"

# Simple column headers
HEADERS = ["Date", "Machine", "Tokens", "Cost ($)"]

# Totals row label
TOTALS_LABEL = "TOTALS"


def get_machine_id():
    """Get a deterministic machine identifier based on hostname + MAC address."""
    hostname = platform.node()
    mac = uuid.getnode()  # Gets MAC address as integer

    # Create deterministic hash from hostname + MAC
    machine_str = f"{hostname}-{mac}"
    hash_bytes = hashlib.sha256(machine_str.encode()).hexdigest()[:8]

    return f"{hostname}-{hash_bytes}"


def fetch_usage_data(period):
    """Fetch usage data from bunx ccusage for the given period."""
    try:
        print(f"Fetching {period} data...")
        result = subprocess.run(
            ['bunx', 'ccusage', period, '--json'],
            capture_output=True,
            text=True,
            check=True,
            encoding='utf-8'
        )

        data = json.loads(result.stdout)
        entries = data.get(period, [])

        if entries:
            print(f"  [OK] {len(entries)} entries")
        else:
            print(f"  [X] No entries found")

        return entries

    except subprocess.CalledProcessError as e:
        print(f"  [X] Error: {e.stderr}")
        return None
    except json.JSONDecodeError as e:
        print(f"  [X] JSON parse error: {e}")
        return None
    except Exception as e:
        print(f"  [X] Error: {e}")
        return None


def load_or_create_workbook():
    """Load existing workbook or create a new one."""
    excel_path = Path(EXCEL_FILE)

    if excel_path.exists():
        print(f"Loading: {EXCEL_FILE}")
        try:
            return load_workbook(EXCEL_FILE)
        except Exception as e:
            print(f"  [X] Error loading: {e}")
            print("  Creating new workbook...")

    print(f"Creating: {EXCEL_FILE}")
    wb = Workbook()

    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Create Daily, Weekly, and Monthly sheets
    for sheet_name in ["Daily", "Weekly", "Monthly"]:
        ws = wb.create_sheet(sheet_name)
        setup_sheet_headers(ws)

    return wb


def setup_sheet_headers(ws):
    """Set up headers for a sheet."""
    for col_idx, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"


def find_existing_row(sheet, date_str, machine_id):
    """Find row with matching date and machine."""
    for row_num in range(2, sheet.max_row + 1):
        cell_date = sheet.cell(row=row_num, column=1).value
        cell_machine = sheet.cell(row=row_num, column=2).value

        if cell_date == TOTALS_LABEL:
            continue

        if str(cell_date) == str(date_str) and cell_machine == machine_id:
            return row_num
    return None


def find_totals_row(sheet):
    """Find the TOTALS row."""
    for row_num in range(2, sheet.max_row + 1):
        if sheet.cell(row=row_num, column=1).value == TOTALS_LABEL:
            return row_num
    return None


def get_insert_position(sheet, date_str):
    """Find position to insert new row (sorted by date)."""
    totals_row = find_totals_row(sheet)
    max_row = totals_row if totals_row else sheet.max_row + 1

    for row_num in range(2, max_row):
        cell_date = sheet.cell(row=row_num, column=1).value
        if cell_date and str(cell_date) > str(date_str):
            return row_num

    return max_row


def update_sheet(sheet, entries, machine_id):
    """Update sheet with entries."""
    if not entries:
        return False

    rows_added = 0
    rows_updated = 0

    for entry in entries:
        date_str = entry.get('date') or entry.get('week') or entry.get('month', 'Unknown')

        existing_row = find_existing_row(sheet, date_str, machine_id)

        if existing_row:
            row_num = existing_row
            rows_updated += 1
        else:
            insert_pos = get_insert_position(sheet, date_str)
            sheet.insert_rows(insert_pos)
            row_num = insert_pos
            rows_added += 1

        # Write simplified data
        sheet.cell(row=row_num, column=1, value=date_str)
        sheet.cell(row=row_num, column=2, value=machine_id)
        sheet.cell(row=row_num, column=3, value=entry.get('totalTokens', 0))
        sheet.cell(row=row_num, column=4, value=round(entry.get('totalCost', 0), 2))

    update_totals_row(sheet)
    print(f"  Added {rows_added}, updated {rows_updated} rows")
    return True


def update_totals_row(sheet):
    """Calculate and update the totals row."""
    totals_row = find_totals_row(sheet)

    if not totals_row:
        totals_row = sheet.max_row + 1

    total_tokens = 0
    total_cost = 0.0

    for row_num in range(2, sheet.max_row + 1):
        if sheet.cell(row=row_num, column=1).value == TOTALS_LABEL:
            continue

        total_tokens += sheet.cell(row=row_num, column=3).value or 0
        total_cost += sheet.cell(row=row_num, column=4).value or 0

    sheet.cell(row=totals_row, column=1, value=TOTALS_LABEL)
    sheet.cell(row=totals_row, column=2, value="")
    sheet.cell(row=totals_row, column=3, value=total_tokens)
    sheet.cell(row=totals_row, column=4, value=round(total_cost, 2))


def format_sheet(sheet):
    """Apply formatting to the sheet."""
    # Column widths
    widths = [12, 28, 15, 12]
    for idx, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(idx)].width = width

    # Format data rows
    for row_num in range(2, sheet.max_row + 1):
        date_val = sheet.cell(row=row_num, column=1).value

        # Right-align numeric columns
        sheet.cell(row=row_num, column=3).alignment = Alignment(horizontal="right")
        sheet.cell(row=row_num, column=4).alignment = Alignment(horizontal="right")

        # Number formats
        sheet.cell(row=row_num, column=3).number_format = '#,##0'
        sheet.cell(row=row_num, column=4).number_format = '$#,##0.00'

        # Style TOTALS row
        if date_val == TOTALS_LABEL:
            for col_idx in range(1, 5):
                cell = sheet.cell(row=row_num, column=col_idx)
                cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                cell.font = Font(bold=True)
                cell.border = Border(top=Side(style='thin'))


def display_summary(wb):
    """Display summary."""
    print("\n" + "=" * 45)
    print("SUMMARY")
    print("=" * 45)

    for sheet_name in ["Daily", "Weekly", "Monthly"]:
        if sheet_name not in wb.sheetnames:
            continue

        sheet = wb[sheet_name]
        totals_row = find_totals_row(sheet)

        if totals_row:
            tokens = sheet.cell(row=totals_row, column=3).value or 0
            cost = sheet.cell(row=totals_row, column=4).value or 0
            entry_count = totals_row - 2

            print(f"\n{sheet_name}: {entry_count} entries")
            print(f"  Tokens: {tokens:,}")
            print(f"  Cost:   ${cost:.2f}")

    print("\n" + "=" * 45)


def main():
    """Main execution flow."""
    print("\n" + "=" * 45)
    print("AI Usage Budget Tracker")
    print("=" * 45)

    machine_id = get_machine_id()
    print(f"Machine: {machine_id}\n")

    # Fetch data
    daily_entries = fetch_usage_data('daily')
    weekly_entries = fetch_usage_data('weekly')
    monthly_entries = fetch_usage_data('monthly')

    if not daily_entries and not weekly_entries and not monthly_entries:
        print("\n[X] No data fetched. Check bunx ccusage.")
        sys.exit(1)

    # Update workbook
    print("\nUpdating Excel...")
    wb = load_or_create_workbook()

    if daily_entries:
        print("Daily sheet:")
        if "Daily" not in wb.sheetnames:
            ws = wb.create_sheet("Daily")
            setup_sheet_headers(ws)
        update_sheet(wb["Daily"], daily_entries, machine_id)
        format_sheet(wb["Daily"])

    if weekly_entries:
        print("Weekly sheet:")
        if "Weekly" not in wb.sheetnames:
            ws = wb.create_sheet("Weekly")
            setup_sheet_headers(ws)
        update_sheet(wb["Weekly"], weekly_entries, machine_id)
        format_sheet(wb["Weekly"])

    if monthly_entries:
        print("Monthly sheet:")
        if "Monthly" not in wb.sheetnames:
            ws = wb.create_sheet("Monthly")
            setup_sheet_headers(ws)
        update_sheet(wb["Monthly"], monthly_entries, machine_id)
        format_sheet(wb["Monthly"])

    # Save
    try:
        wb.save(EXCEL_FILE)
        print(f"\n[OK] Saved to {EXCEL_FILE}")
        display_summary(wb)
    except Exception as e:
        print(f"\n[X] Error saving: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
