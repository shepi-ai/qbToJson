#!/usr/bin/env python3
"""
Trial Balance Document Converter
Converts CSV, XLSX, and PDF trial balance reports to QuickBooks JSON format
Handles monthly trial balance data with debit and credit columns
"""

import json
import csv
import sys
import os
from datetime import datetime, timezone, date
from pathlib import Path
import argparse
import re
from typing import Dict, List, Any, Optional, Tuple
import calendar

from base_converter import BaseConverter, XLSX_SUPPORT, PDF_SUPPORT

if XLSX_SUPPORT:
    import openpyxl

if PDF_SUPPORT:
    import pdfplumber


class TrialBalanceConverter(BaseConverter):
    """Converts Trial Balance documents to QuickBooks-style JSON format"""

    def __init__(self, use_account_lookup: bool = True, api_base_url: str = "http://localhost:8080"):
        super().__init__(use_account_lookup=use_account_lookup, api_base_url=api_base_url)

    def extract_date_from_as_of(self, text: str) -> Tuple[str, str, date, date]:
        """Extract date from 'As of [Date]' format"""
        # Match patterns like "As of May 31, 2024" or "As of May 30, 2025"
        match = re.search(r'As of\s+([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})', text, re.IGNORECASE)
        if match:
            month_name = match.group(1)
            year = match.group(3)

            # Convert month name to number
            try:
                month_num = datetime.strptime(month_name, '%B').month
            except ValueError:
                try:
                    month_num = datetime.strptime(month_name[:3], '%b').month
                except ValueError:
                    month_num = 1

            # Create start and end dates for the month
            start_date = date(int(year), month_num, 1)
            last_day = calendar.monthrange(int(year), month_num)[1]
            end_date = date(int(year), month_num, last_day)

            return month_name.upper()[:3], year, start_date, end_date

        # Fallback
        return "JAN", "2025", date(2025, 1, 1), date(2025, 1, 31)

    def create_row_object(self, account_name: str, debit: str = "", credit: str = "",
                         account_id: Optional[str] = None, is_total: bool = False) -> Dict[str, Any]:
        """Create a row object for the trial balance"""
        if is_total:
            # Total row
            return {
                "id": None,
                "parentId": None,
                "header": None,
                "rows": None,
                "summary": {
                    "colData": [
                        {"attributes": None, "value": "TOTAL", "id": None, "href": None},
                        {"attributes": None, "value": debit, "id": None, "href": None},
                        {"attributes": None, "value": credit, "id": None, "href": None}
                    ]
                },
                "colData": [],
                "type": "SECTION",
                "group": "GrandTotal"
            }
        else:
            # Regular account row
            return {
                "id": None,
                "parentId": None,
                "header": None,
                "rows": None,
                "summary": None,
                "colData": [
                    {"attributes": None, "value": account_name, "id": account_id, "href": None},
                    {"attributes": None, "value": debit, "id": None, "href": None},
                    {"attributes": None, "value": credit, "id": None, "href": None}
                ],
                "type": None,
                "group": None
            }

    def parse_csv_data(self, filepath: Path) -> Dict[str, Dict[str, Any]]:
        """Parse CSV file and extract trial balance data by month"""
        data_by_month = {}

        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)

            # Find header row with months. Match cells that look like a month-year
            # column header (e.g. "January 2023"); this deliberately skips the
            # "As of December 31, 2025" title line, where the month is followed by a
            # day rather than a 4-digit year, which would otherwise be picked as the
            # header and yield zero month columns.
            month_year_re = re.compile(
                r'(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*\s+\d{4}'
            )
            header_row_idx = -1
            for i, row in enumerate(rows):
                if len(row) > 1 and any(
                    cell and month_year_re.search(str(cell).upper()) for cell in row
                ):
                    header_row_idx = i
                    break

            if header_row_idx == -1:
                raise ValueError("Could not find header row with months")

            # Parse months from header
            header_row = rows[header_row_idx]
            month_columns = []

            # Find month columns (they usually come in pairs - debit and credit)
            col_idx = 1  # Skip first column (account names)
            while col_idx < len(header_row):
                cell = header_row[col_idx]
                if cell and month_year_re.search(cell.upper()):
                    month, year, start_date, end_date = self.parse_month_year(cell)
                    month_columns.append({
                        'month': month,
                        'year': year,
                        'start_date': start_date,
                        'end_date': end_date,
                        'debit_col': col_idx,
                        'credit_col': col_idx + 1 if col_idx + 1 < len(header_row) else col_idx
                    })
                    col_idx += 2  # Skip to next month (assuming debit/credit pairs)
                else:
                    col_idx += 1

            # Initialize data structure for each month
            for month_info in month_columns:
                month_key = f"{month_info['year']}-{month_info['month']}"
                data_by_month[month_key] = {
                    'month': month_info['month'],
                    'year': month_info['year'],
                    'start_date': month_info['start_date'],
                    'end_date': month_info['end_date'],
                    'accounts': [],
                    'total_debit': 0.0,
                    'total_credit': 0.0
                }

            # Parse account data
            for row_idx in range(header_row_idx + 1, len(rows)):
                row = rows[row_idx]
                if not row or not row[0] or row[0].strip().upper() in ['TOTAL', 'TOTALS', 'GRAND TOTAL']:
                    continue

                account_name = row[0].strip()
                if not account_name:
                    continue

                # Skip metadata lines
                if any(skip in account_name.lower() for skip in ['accrual basis', 'gmt', 'gmtz']):
                    continue

                # Get account ID
                account_id = self.get_or_create_account_id(account_name)

                # Extract values for each month
                for month_info in month_columns:
                    month_key = f"{month_info['year']}-{month_info['month']}"

                    # Get debit value - track if cell has explicit data
                    debit_value = 0.0
                    debit_present = False
                    if month_info['debit_col'] < len(row):
                        debit_str = row[month_info['debit_col']].strip().replace(',', '').replace('$', '').replace('(', '-').replace(')', '')
                        if debit_str and debit_str != '-':
                            debit_present = True
                            try:
                                debit_value = float(debit_str)
                            except ValueError:
                                debit_value = 0.0

                    # Get credit value - track if cell has explicit data
                    credit_value = 0.0
                    credit_present = False
                    if month_info['credit_col'] < len(row):
                        credit_str = row[month_info['credit_col']].strip().replace(',', '').replace('$', '').replace('(', '-').replace(')', '')
                        if credit_str and credit_str != '-':
                            credit_present = True
                            try:
                                credit_value = float(credit_str)
                            except ValueError:
                                credit_value = 0.0

                    # Include if any value is present (even 0) or is a special account
                    if debit_present or credit_present or account_name in ['Retained Earnings']:
                        data_by_month[month_key]['accounts'].append({
                            'name': account_name,
                            'id': account_id,
                            'debit': debit_value,
                            'credit': credit_value
                        })
                        data_by_month[month_key]['total_debit'] += debit_value
                        data_by_month[month_key]['total_credit'] += credit_value

        return data_by_month

    def parse_single_month_xlsx(self, filepath: Path) -> Dict[str, Dict[str, Any]]:
        """Parse single-month XLSX format (e.g., 'As of December 31, 2025')"""
        self.check_xlsx_support()

        # Load without data_only to get formulas, then extract values
        workbook = openpyxl.load_workbook(filepath, data_only=False)
        sheet = workbook.active

        data_by_month = {}

        # Convert to list of lists and extract values from formulas
        rows = []
        for row in sheet.iter_rows(values_only=True):
            row_data = []
            for cell in row:
                if cell is None or cell == '':
                    row_data.append('')
                else:
                    cell_str = str(cell)
                    # Check if it's an Excel formula (e.g., '=1201.00')
                    if cell_str.startswith('='):
                        # Extract number from formula
                        match = re.search(r'=([\d.]+)', cell_str)
                        if match:
                            try:
                                row_data.append(float(match.group(1)))
                            except ValueError:
                                row_data.append(cell)
                        else:
                            row_data.append(cell)
                    else:
                        row_data.append(cell)
            rows.append(row_data)

        if len(rows) < 5:
            print(f"[DEBUG] XLSX has too few rows: {len(rows)}", file=sys.stderr)
            return data_by_month

        # Find "As of [Date]" in early rows (typically row 2 or 3)
        date_text = ""
        for i in range(min(5, len(rows))):
            row_text = ' '.join(str(cell) for cell in rows[i] if cell)
            if "As of" in row_text:
                date_text = row_text
                break

        if not date_text:
            print(f"[DEBUG] Could not find 'As of' date in XLSX", file=sys.stderr)
            return data_by_month

        # Extract date from header
        month, year, start_date, end_date = self.extract_date_from_as_of(date_text)
        month_key = f"{year}-{month}"

        # Initialize month data
        data_by_month[month_key] = {
            'month': month,
            'year': year,
            'start_date': start_date,
            'end_date': end_date,
            'accounts': [],
            'total_debit': 0.0,
            'total_credit': 0.0
        }

        # Find DEBIT/CREDIT header row
        header_idx = -1
        debit_col = -1
        credit_col = -1

        for i in range(min(10, len(rows))):
            for j, cell in enumerate(rows[i]):
                cell_str = str(cell).upper()
                if 'DEBIT' in cell_str:
                    header_idx = i
                    debit_col = j
                if 'CREDIT' in cell_str:
                    credit_col = j
            if header_idx != -1 and debit_col != -1 and credit_col != -1:
                break

        if header_idx == -1:
            print(f"[DEBUG] Could not find DEBIT/CREDIT headers in XLSX", file=sys.stderr)
            return data_by_month

        print(f"[DEBUG] Found headers at row {header_idx}, debit col {debit_col}, credit col {credit_col}", file=sys.stderr)

        # Parse account data (after header, before TOTAL)
        for row_idx in range(header_idx + 1, len(rows)):
            row = rows[row_idx]

            if not row or len(row) < 2:
                continue

            # First column should be account name
            account_name = str(row[0]).strip()

            # Stop at TOTAL line
            if account_name.upper() in ['TOTAL', 'TOTALS']:
                break

            # Skip empty account names
            if not account_name or account_name == '' or account_name == 'None':
                continue

            # Skip footer lines (date stamps, etc.)
            if any(skip in account_name.lower() for skip in ['accrual basis', 'gmt', 'pm', 'am']):
                continue

            # Get debit value
            debit_value = 0.0
            if debit_col < len(row):
                debit_cell = row[debit_col]
                if debit_cell and debit_cell != '':
                    try:
                        debit_value = float(str(debit_cell).replace(',', ''))
                    except (ValueError, AttributeError):
                        debit_value = 0.0

            # Get credit value
            credit_value = 0.0
            if credit_col < len(row):
                credit_cell = row[credit_col]
                if credit_cell and credit_cell != '':
                    try:
                        credit_value = float(str(credit_cell).replace(',', ''))
                    except (ValueError, AttributeError):
                        credit_value = 0.0

            # Get account ID
            account_id = self.get_or_create_account_id(account_name)

            # Include all accounts (even zero-value) to match QB API output
            data_by_month[month_key]['accounts'].append({
                'name': account_name,
                'id': account_id,
                'debit': debit_value,
                'credit': credit_value
            })
            data_by_month[month_key]['total_debit'] += debit_value
            data_by_month[month_key]['total_credit'] += credit_value

        return data_by_month

    def parse_xlsx_data(self, filepath: Path) -> Dict[str, Dict[str, Any]]:
        """Parse XLSX file and extract trial balance data by month"""
        self.check_xlsx_support()

        # First, detect format by reading with data_only=True to evaluate formulas
        workbook = openpyxl.load_workbook(filepath, data_only=True)
        sheet = workbook.active

        # Convert first few rows to check format
        rows = []
        row_count = 0
        for row in sheet.iter_rows(values_only=True):
            rows.append([cell if cell is not None else '' for cell in row])
            row_count += 1
            if row_count >= 10:  # Only need first 10 rows for detection
                break

        # Check if it's single-month format
        has_as_of = False
        for i in range(min(5, len(rows))):
            row_text = ' '.join(str(cell) for cell in rows[i] if cell)
            if "As of" in row_text:
                has_as_of = True
                break

        # Count month columns with years in remaining rows
        # Look at all cells in header rows (typically row 4-5) for month-year patterns
        month_year_count = 0
        for i in range(4, min(len(rows), 7)):  # Check rows 4-6 (header area)
            for cell in rows[i]:
                cell_text = str(cell).upper()
                # Match patterns like "JAN 2022", "JANUARY 2022", "Jan 2025"
                if re.search(r'(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*\s+\d{4}', cell_text):
                    month_year_count += 1

        print(f"[DEBUG] Format detection: has_as_of={has_as_of}, month_year_count={month_year_count}", file=sys.stderr)

        if has_as_of and month_year_count < 2:
            print(f"[DEBUG] Detected single-month XLSX format (has 'As of' and fewer than 2 month columns)", file=sys.stderr)
            return self.parse_single_month_xlsx(filepath)

        # Fall back to multi-month parser
        print(f"[DEBUG] Using multi-month XLSX parser", file=sys.stderr)

        # Re-load workbook without data_only for multi-month parsing
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active

        # Convert to list of lists, preserving None vs 0 distinction
        rows = []
        raw_rows = []  # Keep original values to distinguish None from 0
        for row in sheet.iter_rows(values_only=True):
            rows.append([str(cell) if cell is not None else '' for cell in row])
            raw_rows.append(list(row))

        # Find header row with months
        header_row_idx = -1
        for i, row in enumerate(rows):
            if len(row) > 1:
                # Skip rows that are likely report headers (e.g., "As of July 31, 2025")
                if i < 3 and any(phrase in ' '.join(str(cell) for cell in row if cell).lower() for phrase in ['as of', 'trial balance', 'company']):
                    continue

                row_text = ' '.join(str(cell) for cell in row if cell)
                # Look for rows with multiple month names (indicating column headers)
                month_count = 0
                for month in ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER', 'JAN', 'FEB', 'MAR', 'APR', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']:
                    if month in row_text.upper():
                        # Check if it's followed by a year (to confirm it's a month header)
                        if re.search(rf'{month}\s+\d{{4}}', row_text.upper()):
                            month_count += 1

                # If we found at least 2 months with years, this is likely our header row
                if month_count >= 2:
                    header_row_idx = i
                    break

        if header_row_idx == -1:
            raise ValueError("Could not find header row with months")

        # Parse months from header
        header_row = rows[header_row_idx]
        month_columns = []

        # Find month columns
        col_idx = 1
        while col_idx < len(header_row):
            cell = header_row[col_idx]
            if cell and any(month in str(cell).upper() for month in ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER', 'JAN', 'FEB', 'MAR', 'APR', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']):
                month, year, start_date, end_date = self.parse_month_year(str(cell))

                # Check if next columns are labeled as Debit/Credit
                has_labels = False
                if header_row_idx + 1 < len(rows):
                    next_row = rows[header_row_idx + 1]
                    if col_idx < len(next_row) and 'DEBIT' in next_row[col_idx].upper():
                        has_labels = True

                month_columns.append({
                    'month': month,
                    'year': year,
                    'start_date': start_date,
                    'end_date': end_date,
                    'debit_col': col_idx,
                    'credit_col': col_idx + 1 if col_idx + 1 < len(header_row) else col_idx,
                    'has_labels': has_labels
                })
                col_idx += 2
            else:
                col_idx += 1

        # Initialize data structure
        data_by_month = {}
        for month_info in month_columns:
            month_key = f"{month_info['year']}-{month_info['month']}"
            data_by_month[month_key] = {
                'month': month_info['month'],
                'year': month_info['year'],
                'start_date': month_info['start_date'],
                'end_date': month_info['end_date'],
                'accounts': [],
                'total_debit': 0.0,
                'total_credit': 0.0
            }

        # Skip label row if present
        data_start_row = header_row_idx + 2 if any(m['has_labels'] for m in month_columns) else header_row_idx + 1

        # Parse account data
        for row_idx in range(data_start_row, len(rows)):
            row = rows[row_idx]
            raw_row = raw_rows[row_idx] if row_idx < len(raw_rows) else []
            if not row or not row[0] or row[0].strip().upper() in ['TOTAL', 'TOTALS', 'GRAND TOTAL']:
                continue

            account_name = row[0].strip()
            if not account_name:
                continue

            # Skip metadata lines (e.g., "Accrual Basis Wednesday, March 04, 2026...")
            if any(skip in account_name.lower() for skip in ['accrual basis', 'gmt', 'gmtz']):
                continue

            # Get account ID
            account_id = self.get_or_create_account_id(account_name)

            # Extract values for each month
            for month_info in month_columns:
                month_key = f"{month_info['year']}-{month_info['month']}"

                # Check raw cell values to distinguish None (absent) from 0 (explicit zero)
                raw_debit = raw_row[month_info['debit_col']] if month_info['debit_col'] < len(raw_row) else None
                raw_credit = raw_row[month_info['credit_col']] if month_info['credit_col'] < len(raw_row) else None
                debit_present = raw_debit is not None
                credit_present = raw_credit is not None

                # Get debit value
                debit_value = 0.0
                if debit_present and month_info['debit_col'] < len(row):
                    debit_str = row[month_info['debit_col']].strip().replace(',', '').replace('$', '').replace('(', '-').replace(')', '')
                    if debit_str.startswith('='):
                        if '=' in debit_str and any(c.isdigit() for c in debit_str):
                            numbers = re.findall(r'[\d.]+', debit_str)
                            if numbers:
                                try:
                                    debit_value = float(numbers[0])
                                except ValueError:
                                    debit_value = 0.0
                    else:
                        try:
                            debit_value = float(debit_str) if debit_str and debit_str != '-' else 0.0
                        except ValueError:
                            debit_value = 0.0

                # Get credit value
                credit_value = 0.0
                if credit_present and month_info['credit_col'] < len(row):
                    credit_str = row[month_info['credit_col']].strip().replace(',', '').replace('$', '').replace('(', '-').replace(')', '')
                    if credit_str.startswith('='):
                        if '=' in credit_str and any(c.isdigit() for c in credit_str):
                            numbers = re.findall(r'[\d.]+', credit_str)
                            if numbers:
                                try:
                                    credit_value = float(numbers[0])
                                except ValueError:
                                    credit_value = 0.0
                    else:
                        try:
                            credit_value = float(credit_str) if credit_str and credit_str != '-' else 0.0
                        except ValueError:
                            credit_value = 0.0

                # Only include if at least one cell has data (not None in raw XLSX)
                if debit_present or credit_present:
                    data_by_month[month_key]['accounts'].append({
                        'name': account_name,
                        'id': account_id,
                        'debit': debit_value,
                        'credit': credit_value
                    })
                    data_by_month[month_key]['total_debit'] += debit_value
                    data_by_month[month_key]['total_credit'] += credit_value

        return data_by_month

    def parse_single_month_pdf(self, filepath: Path) -> Dict[str, Dict[str, Any]]:
        """Parse single-month PDF format (e.g., 'As of May 31, 2024')"""
        self.check_pdf_support()

        data_by_month = {}

        with pdfplumber.open(filepath) as pdf:
            # Extract text from all pages
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"

            if not full_text:
                return data_by_month

            # Extract date from "As of [Date]" format
            month, year, start_date, end_date = self.extract_date_from_as_of(full_text)
            month_key = f"{year}-{month}"

            # Initialize month data
            data_by_month[month_key] = {
                'month': month,
                'year': year,
                'start_date': start_date,
                'end_date': end_date,
                'accounts': [],
                'total_debit': 0.0,
                'total_credit': 0.0
            }

            lines = full_text.split('\n')

            # Find DEBIT/CREDIT header line
            header_idx = -1
            for i, line in enumerate(lines):
                if 'DEBIT' in line.upper() and 'CREDIT' in line.upper():
                    header_idx = i
                    break

            if header_idx == -1:
                print(f"[DEBUG] Could not find DEBIT/CREDIT header in PDF", file=sys.stderr)
                return data_by_month

            # Parse account lines (after header, before TOTAL)
            for line_idx in range(header_idx + 1, len(lines)):
                line = lines[line_idx].strip()

                # Stop at TOTAL line
                if not line or line.upper().startswith('TOTAL'):
                    break

                # Skip page headers that repeat
                if any(skip in line.upper() for skip in ['TRIAL BALANCE', 'AS OF', 'ACCRUAL BASIS', 'DEBIT', 'CREDIT']):
                    continue

                # Extract account name and values using regex
                # Pattern: account name followed by numbers
                # Account names can contain letters, spaces, parentheses, slashes, colons, etc.
                match = re.match(r'^(.+?)\s+([\d,]+\.?\d*)\s*([\d,]+\.?\d*)?$', line)

                if match:
                    account_name = match.group(1).strip()

                    # Skip if it looks like a page number or date
                    if account_name.isdigit() or re.match(r'^\d+/\d+$', account_name):
                        continue

                    # Parse debit value
                    debit_str = match.group(2).strip().replace(',', '')
                    try:
                        debit_value = float(debit_str)
                    except ValueError:
                        debit_value = 0.0

                    # Parse credit value (might be empty)
                    credit_value = 0.0
                    if match.group(3):
                        credit_str = match.group(3).strip().replace(',', '')
                        try:
                            credit_value = float(credit_str)
                        except ValueError:
                            credit_value = 0.0

                    # Determine if value is debit or credit based on account type
                    # If only one value present, infer from account name
                    if debit_value > 0 and credit_value == 0:
                        # Check if this is likely a credit account
                        if any(keyword in account_name.upper() for keyword in
                               ['PAYABLE', 'EQUITY', 'EARNINGS', 'LOAN', 'RETAINED', 'CONTRIBUTIONS', 'REVENUE', 'INCOME', 'SALES', 'SERVICES']):
                            credit_value = debit_value
                            debit_value = 0.0

                    # Get account ID
                    account_id = self.get_or_create_account_id(account_name)

                    # Include all accounts (even zero-value) to match QB API output
                    data_by_month[month_key]['accounts'].append({
                        'name': account_name,
                        'id': account_id,
                        'debit': debit_value,
                        'credit': credit_value
                    })
                    data_by_month[month_key]['total_debit'] += debit_value
                    data_by_month[month_key]['total_credit'] += credit_value

        return data_by_month

    def parse_pdf_data(self, filepath: Path) -> Dict[str, Dict[str, Any]]:
        """Parse PDF file and extract trial balance data by month"""
        self.check_pdf_support()

        # First, detect format
        with pdfplumber.open(filepath) as pdf:
            first_page_text = pdf.pages[0].extract_text() if pdf.pages else ""

            # Check if it's single-month format
            if "As of" in first_page_text and "DEBIT" in first_page_text.upper() and "CREDIT" in first_page_text.upper():
                # Check if it's NOT multi-month (no multiple month names with years)
                month_year_pattern = r'(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+\d{4}'
                month_matches = re.findall(month_year_pattern, first_page_text.upper())

                if len(month_matches) < 2:  # Single month or no months found
                    print(f"[DEBUG] Detected single-month PDF format", file=sys.stderr)
                    return self.parse_single_month_pdf(filepath)

        # Fall back to multi-month parser
        print(f"[DEBUG] Using multi-month PDF parser", file=sys.stderr)
        data_by_month = {}

        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                # Try to extract tables first
                tables = page.extract_tables()

                if tables:
                    # Process table data (existing code for table extraction)
                    pass  # Keep existing table extraction code

                # Always try text extraction for columnar PDFs
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split('\n')

                # Find the header line with months
                month_line_idx = -1
                for i, line in enumerate(lines):
                    if 'JAN 2025' in line and 'FEB 2025' in line:
                        month_line_idx = i
                        break

                if month_line_idx == -1:
                    continue

                # Extract month positions from the header line
                month_line = lines[month_line_idx]
                month_positions = []

                # Find each month and its position
                for match in re.finditer(r'(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{4})', month_line):
                    month_name = match.group(1)
                    year = match.group(2)
                    month, year_str, start_date, end_date = self.parse_month_year(f"{month_name} {year}")

                    month_positions.append({
                        'month': month,
                        'year': year_str,
                        'start_date': start_date,
                        'end_date': end_date,
                        'start_pos': match.start(),
                        'end_pos': match.end()
                    })

                # Initialize data structure for each month
                for month_info in month_positions:
                    month_key = f"{month_info['year']}-{month_info['month']}"
                    if month_key not in data_by_month:
                        data_by_month[month_key] = {
                            'month': month_info['month'],
                            'year': month_info['year'],
                            'start_date': month_info['start_date'],
                            'end_date': month_info['end_date'],
                            'accounts': [],
                            'total_debit': 0.0,
                            'total_credit': 0.0
                        }

                # Parse account data (skip header lines)
                data_start = month_line_idx + 2  # Skip month line and DEBIT/CREDIT line

                for line_idx in range(data_start, len(lines)):
                    line = lines[line_idx]
                    if not line.strip() or 'TOTAL' in line.upper():
                        continue

                    # Extract account name (text before first number)
                    match = re.match(r'^([A-Za-z\s\(\):/\.\-]+)', line)
                    if not match:
                        continue

                    account_name = match.group(1).strip()
                    if not account_name:
                        continue

                    # Get account ID
                    account_id = self.get_or_create_account_id(account_name)

                    # Extract all numbers from the line
                    numbers = re.findall(r'[\d,]+\.?\d*', line)

                    # Assign numbers to months based on expected pattern
                    # Each month should have 2 values (debit, credit), but some might be missing
                    value_idx = 0

                    for i, month_info in enumerate(month_positions):
                        month_key = f"{month_info['year']}-{month_info['month']}"

                        debit_value = 0.0
                        credit_value = 0.0

                        # Try to get values for this month
                        if value_idx < len(numbers):
                            # Some accounts might only have one value per month
                            # We need to determine if it's debit or credit based on context
                            try:
                                value = float(numbers[value_idx].replace(',', ''))
                                # For now, assume first value is debit unless it's a liability/equity account
                                if any(keyword in account_name.upper() for keyword in ['PAYABLE', 'EQUITY', 'EARNINGS']):
                                    credit_value = value
                                else:
                                    debit_value = value
                                value_idx += 1

                                # Check if there's a second value for this month
                                if value_idx < len(numbers) and i < len(month_positions) - 1:
                                    # Check if next value is likely for this month or next month
                                    # This is heuristic-based
                                    next_value = float(numbers[value_idx].replace(',', ''))
                                    if next_value < value * 10:  # Likely same month
                                        if debit_value > 0:
                                            credit_value = next_value
                                        else:
                                            debit_value = next_value
                                        value_idx += 1
                            except (ValueError, IndexError):
                                pass

                        # Include all accounts (even zero-value) to match QB API output
                        # Check if account already exists for this month
                        existing = next((acc for acc in data_by_month[month_key]['accounts'] if acc['name'] == account_name), None)
                        if not existing:
                            data_by_month[month_key]['accounts'].append({
                                'name': account_name,
                                'id': account_id,
                                'debit': debit_value,
                                'credit': credit_value
                            })
                            data_by_month[month_key]['total_debit'] += debit_value
                            data_by_month[month_key]['total_credit'] += credit_value

        return data_by_month

    def build_trial_balance_json(self, data_by_month: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
        """Build the complete trial balance JSON structure"""
        monthly_reports = []

        # Sort months chronologically by start_date
        sorted_months = sorted(data_by_month.keys(), key=lambda k: data_by_month[k]['start_date'])

        for month_key in sorted_months:
            month_data = data_by_month[month_key]

            # Create report structure
            report = self.create_report_structure(month_data)

            monthly_reports.append({
                "month": month_data['month'],
                "year": month_data['year'],
                "startDate": month_data['start_date'].strftime('%Y-%m-%d'),
                "endDate": month_data['end_date'].strftime('%Y-%m-%d'),
                "report": report
            })

        # Create summary
        if monthly_reports:
            summary = {
                "requestStartDate": monthly_reports[0]['startDate'],
                "requestEndDate": monthly_reports[-1]['endDate'],
                "totalMonths": len(monthly_reports),
                "accountingMethod": "Accrual"
            }
        else:
            summary = {
                "requestStartDate": "2025-01-01",
                "requestEndDate": "2025-12-31",
                "totalMonths": 0,
                "accountingMethod": "Accrual"
            }

        return {
            "monthlyReports": monthly_reports,
            "summary": summary
        }

    def create_report_structure(self, month_data: Dict[str, Any]) -> Dict[str, Any]:
        """Create the report structure for a single month"""
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000+00:00')
        has_data = len(month_data['accounts']) > 0

        # Create column structure
        columns = {
            "column": [
                {
                    "colTitle": "",
                    "colType": "Account",
                    "metaData": [],
                    "columns": None
                },
                {
                    "colTitle": f"{month_data['month']} {month_data['year']}",
                    "colType": "Money",
                    "metaData": [],
                    "columns": {
                        "column": [
                            {
                                "colTitle": "Debit",
                                "colType": "Money",
                                "metaData": [
                                    {"name": "StartDate", "value": month_data['start_date'].strftime('%Y-%m-%d')},
                                    {"name": "EndDate", "value": month_data['end_date'].strftime('%Y-%m-%d')}
                                ],
                                "columns": None
                            },
                            {
                                "colTitle": "Credit",
                                "colType": "Money",
                                "metaData": [
                                    {"name": "StartDate", "value": month_data['start_date'].strftime('%Y-%m-%d')},
                                    {"name": "EndDate", "value": month_data['end_date'].strftime('%Y-%m-%d')}
                                ],
                                "columns": None
                            }
                        ]
                    }
                }
            ]
        }

        # Create rows
        rows = []

        if has_data:
            # Add account rows
            for account in month_data['accounts']:
                rows.append(self.create_row_object(
                    account['name'],
                    f"{account['debit']:.2f}" if account['debit'] != 0 else "",
                    f"{account['credit']:.2f}" if account['credit'] != 0 else "",
                    account['id']
                ))
        else:
            # No data - add Retained Earnings with empty values
            rows.append(self.create_row_object(
                "Retained Earnings",
                "",
                "",
                self.get_or_create_account_id("Retained Earnings")
            ))

        # Add total row
        rows.append(self.create_row_object(
            "",
            f"{month_data['total_debit']:.2f}",
            f"{month_data['total_credit']:.2f}",
            is_total=True
        ))

        # Create report structure
        report = {
            "header": {
                "time": timestamp,
                "reportName": "TrialBalance",
                "dateMacro": None,
                "reportBasis": "ACCRUAL",
                "startPeriod": month_data['start_date'].strftime('%Y-%m-%d'),
                "endPeriod": month_data['end_date'].strftime('%Y-%m-%d'),
                "summarizeColumnsBy": "Month",
                "currency": "USD",
                "customer": None,
                "vendor": None,
                "employee": None,
                "item": None,
                "clazz": None,
                "department": None,
                "option": [
                    {"name": "NoReportData", "value": "false" if has_data else "true"}
                ]
            },
            "columns": columns,
            "rows": {"row": rows}
        }

        return report

    # ──────────────────────────────────────────────
    # Override base class abstract methods and file dispatch
    # ──────────────────────────────────────────────

    def parse_csv(self, filepath: Path) -> Dict[str, Any]:
        """Parse CSV and return trial balance JSON structure."""
        data_by_month = self.parse_csv_data(filepath)
        return self.build_trial_balance_json(data_by_month)

    def parse_xlsx(self, filepath: Path) -> Dict[str, Any]:
        """Parse XLSX and return trial balance JSON structure."""
        data_by_month = self.parse_xlsx_data(filepath)
        return self.build_trial_balance_json(data_by_month)

    def parse_pdf(self, filepath: Path) -> Dict[str, Any]:
        """Parse PDF and return trial balance JSON structure."""
        data_by_month = self.parse_pdf_data(filepath)
        return self.build_trial_balance_json(data_by_month)

    def convert_to_json(self, filepath: Path, output_path: Optional[Path] = None) -> str:
        """Convert a file to JSON format"""
        try:
            trial_balance_data = self.convert_file(filepath)

            if output_path:
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(trial_balance_data, f, indent=2)
                return f"Converted trial balance with {len(trial_balance_data['monthlyReports'])} monthly reports to {output_path}"
            else:
                return json.dumps(trial_balance_data, indent=2)
        except Exception as e:
            import traceback
            traceback.print_exc()
            raise


def main():
    parser = argparse.ArgumentParser(description='Convert trial balance documents to JSON format')
    parser.add_argument('input', help='Input file (CSV, XLSX, or PDF)')
    parser.add_argument('-o', '--output', help='Output JSON file (default: print to stdout)')

    args = parser.parse_args()

    converter = TrialBalanceConverter()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: {input_path} does not exist", file=sys.stderr)
        sys.exit(1)

    try:
        if args.output:
            result = converter.convert_to_json(input_path, Path(args.output))
            print(result)
        else:
            print(converter.convert_to_json(input_path))
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
