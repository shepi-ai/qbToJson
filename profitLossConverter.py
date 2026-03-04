#!/usr/bin/env python3
"""
Profit and Loss Document Converter
Converts CSV, XLSX, and PDF profit and loss reports to QuickBooks JSON format
Generic converter that works with any P&L structure
"""

import csv
import sys
import os
from datetime import datetime, timezone, date
from pathlib import Path
import argparse
import re
from typing import Dict, List, Any, Optional, Tuple

from base_converter import BaseConverter, XLSX_SUPPORT, PDF_SUPPORT


class ProfitLossConverter(BaseConverter):
    """Converts Profit and Loss documents to QuickBooks-style JSON format"""

    def __init__(self, use_account_lookup: bool = True, api_base_url: str = "http://localhost:8080"):
        super().__init__(use_account_lookup=use_account_lookup, api_base_url=api_base_url)

    def create_row_object(self, name: str, value: Optional[str] = None,
                         account_id: Optional[str] = None, row_type: str = "DATA",
                         group: Optional[str] = None, is_section: bool = False,
                         sub_rows: Optional[List] = None, is_summary: bool = False) -> Dict[str, Any]:
        """Create a row object for the profit and loss report"""
        row = {
            "id": None,
            "parentId": None,
            "header": None,
            "rows": None,
            "summary": None,
            "colData": [],
            "type": row_type if row_type else None,
            "group": group
        }

        if is_summary:
            # Summary row (like Total Income, Gross Profit, etc.)
            row["summary"] = {
                "colData": [
                    {"attributes": None, "value": name, "id": None, "href": None},
                    {"attributes": None, "value": value if value else "", "id": None, "href": None}
                ]
            }
            row["type"] = "SECTION"
        elif is_section:
            # Section header
            row["header"] = {
                "colData": [
                    {"attributes": None, "value": name, "id": None, "href": None},
                    {"attributes": None, "value": value if value else "", "id": None, "href": None}
                ]
            }
            if sub_rows:
                row["rows"] = {"row": sub_rows}
            row["type"] = "SECTION"
        else:
            # Data row
            row["colData"] = [
                {"attributes": None, "value": name, "id": account_id, "href": None},
                {"attributes": None, "value": value if value else "", "id": None, "href": None}
            ]
            row["type"] = "DATA"

        return row

    def detect_hierarchy_level(self, row: List[str], row_idx: int, all_rows: List[List[str]]) -> str:
        """Detect if a row is a section header, group header, or data row"""
        account_name = row[0].strip()

        # Check if it's a total row
        if account_name.lower().startswith('total for '):
            return 'total'

        # Check if it's a calculated row (contains specific keywords)
        calc_keywords = ['gross profit', 'net income', 'net operating income', 'net other income']
        if any(keyword in account_name.lower() for keyword in calc_keywords):
            return 'calculated'

        # Look ahead to see if there's a "Total for" this account
        for future_idx in range(row_idx + 1, min(row_idx + 50, len(all_rows))):
            future_row = all_rows[future_idx]
            if future_row and future_row[0]:
                future_name = future_row[0].strip()
                if future_name.lower() == f"total for {account_name.lower()}":
                    return 'group'
                # If we hit another major section, stop looking
                if not future_name.startswith(' ') and future_name and not future_name.lower().startswith('total'):
                    break

        # Check if all value columns are empty (might be a section header)
        has_values = False
        for i in range(1, len(row)):
            if row[i].strip() and row[i].strip() not in ['0.00', '0', '-']:
                # Check if it's a number
                try:
                    float(row[i].strip().replace(',', '').replace('$', '').replace('(', '-').replace(')', ''))
                    has_values = True
                    break
                except:
                    pass

        if not has_values and row_idx < len(all_rows) - 1:
            # Check if next non-empty row is indented or a sub-item
            next_row = all_rows[row_idx + 1] if row_idx + 1 < len(all_rows) else None
            if next_row and next_row[0] and not next_row[0].strip().lower().startswith('total'):
                return 'section'

        return 'data'

    def parse_csv_hierarchy(self, filepath: Path) -> Tuple[List[str], Dict[str, Dict[str, Any]]]:
        """Parse CSV file and extract hierarchical profit and loss data"""
        months = []
        data_by_month = {}

        with open(filepath, 'r', encoding='utf-8') as f:
            # Use csv reader to handle quoted fields properly
            reader = csv.reader(f)
            rows = list(reader)

            # Find the header row with months
            header_row_idx = -1
            for i, row in enumerate(rows):
                if len(row) > 1:  # Must have at least 2 columns
                    # Check each cell for month names
                    for cell in row[1:]:  # Skip first column
                        if cell and any(month in str(cell) for month in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']):
                            header_row_idx = i
                            break
                    if header_row_idx != -1:
                        break

            if header_row_idx == -1:
                raise ValueError("Could not find header row with months")

            # Parse header to get months
            header_row = rows[header_row_idx]
            month_columns = []
            for i, part in enumerate(header_row[1:], 1):  # Skip first column
                if part.strip() and not any(skip in part.lower() for skip in ['total', 'ytd', 'year to date']):
                    month_str, start_date, end_date = self.parse_month_column(part.strip())
                    months.append(month_str)
                    month_columns.append({
                        'index': i,
                        'month': month_str,
                        'start_date': start_date,
                        'end_date': end_date,
                        'header': part.strip()
                    })

            # Initialize data structure for each month
            for month_info in month_columns:
                data_by_month[month_info['month']] = {
                    'start_date': month_info['start_date'],
                    'end_date': month_info['end_date'],
                    'sections': []  # Will store hierarchical data
                }

            # Parse data rows and build hierarchy
            data_rows = rows[header_row_idx + 1:]
            self.parse_rows_recursive(data_rows, 0, month_columns, data_by_month)

        return months, data_by_month

    def parse_rows_recursive(self, rows: List[List[str]], start_idx: int,
                           month_columns: List[Dict], data_by_month: Dict) -> int:
        """Recursively parse rows and build hierarchical structure"""
        idx = start_idx

        while idx < len(rows):
            row = rows[idx]

            # Skip empty rows or rows with accounting basis info
            if not row or not row[0] or 'Accrual Basis' in row[0] or 'Cash Basis' in row[0]:
                idx += 1
                continue

            account_name = row[0].strip()
            if not account_name:
                idx += 1
                continue

            # Detect the type of row
            row_type = self.detect_hierarchy_level(row, idx, rows)

            if row_type == 'total':
                # End of current section/group
                return idx + 1

            elif row_type == 'calculated':
                # Add calculated row to all months
                for month_info in month_columns:
                    month = month_info['month']
                    value = 0.0
                    if month_info['index'] < len(row):
                        value_str = row[month_info['index']].strip().replace(',', '').replace('$', '').replace('(', '-').replace(')', '')
                        try:
                            value = float(value_str) if value_str else 0.0
                        except:
                            value = 0.0

                    # Determine group based on account name
                    group = None
                    if 'gross profit' in account_name.lower():
                        group = 'GrossProfit'
                    elif 'net operating income' in account_name.lower():
                        group = 'NetOperatingIncome'
                    elif 'net other income' in account_name.lower():
                        group = 'NetOtherIncome'
                    elif 'net income' in account_name.lower():
                        group = 'NetIncome'

                    data_by_month[month]['sections'].append({
                        'type': 'calculated',
                        'name': account_name,
                        'value': value,
                        'group': group
                    })
                idx += 1

            elif row_type == 'section':
                # This is a section header
                section_data = {
                    'type': 'section',
                    'name': account_name,
                    'items': [],
                    'total': 0.0
                }

                # Determine group based on common patterns
                group = None
                name_lower = account_name.lower()
                if 'income' in name_lower and 'other' not in name_lower:
                    group = 'Income'
                elif 'cost of goods' in name_lower or 'cogs' in name_lower:
                    group = 'COGS'
                elif 'expense' in name_lower and 'other' not in name_lower:
                    group = 'Expenses'
                elif 'other income' in name_lower:
                    group = 'OtherIncome'
                elif 'other expense' in name_lower:
                    group = 'OtherExpenses'

                section_data['group'] = group

                # Parse subsection
                idx += 1
                idx, items = self.parse_section_items(rows, idx, month_columns)
                section_data['items'] = items

                # Add section to all months
                for month_info in month_columns:
                    month = month_info['month']
                    month_section = section_data.copy()
                    month_section['items'] = [item[month] for item in items if month in item]
                    # Calculate total
                    total = sum(self.calculate_item_total(item) for item in month_section['items'])
                    month_section['total'] = total
                    data_by_month[month]['sections'].append(month_section)

            elif row_type == 'group':
                # This is a group within a section
                idx += 1  # Will be handled within section parsing

            else:  # data row
                # This is a standalone data row
                for month_info in month_columns:
                    month = month_info['month']
                    value = 0.0
                    if month_info['index'] < len(row):
                        value_str = row[month_info['index']].strip().replace(',', '').replace('$', '').replace('(', '-').replace(')', '')
                        try:
                            value = float(value_str) if value_str else 0.0
                        except:
                            value = 0.0

                    if value != 0.0:
                        data_by_month[month]['sections'].append({
                            'type': 'data',
                            'name': account_name,
                            'value': value,
                            'id': self.get_or_create_account_id(account_name)
                        })
                idx += 1

        return idx

    def parse_section_items(self, rows: List[List[str]], start_idx: int,
                          month_columns: List[Dict]) -> Tuple[int, List[Dict]]:
        """Parse items within a section"""
        items = []
        idx = start_idx

        while idx < len(rows):
            row = rows[idx]

            if not row or not row[0]:
                idx += 1
                continue

            account_name = row[0].strip()
            if not account_name:
                idx += 1
                continue

            # Check if we've hit the end of this section
            if account_name.lower().startswith('total for '):
                return idx + 1, items

            # Check if this is a new major section
            row_type = self.detect_hierarchy_level(row, idx, rows)
            if row_type in ['section', 'calculated']:
                return idx, items

            if row_type == 'group':
                # Parse group and its items
                group_data = {}
                group_name = account_name
                idx += 1
                idx, group_items = self.parse_group_items(rows, idx, month_columns, group_name)

                # Store group data for each month
                for month_info in month_columns:
                    month = month_info['month']
                    month_items = [item[month] for item in group_items if month in item]
                    group_data[month] = {
                        'type': 'group',
                        'name': group_name,
                        'items': month_items,
                        'id': self.get_or_create_account_id(group_name)
                    }

                items.append(group_data)
            else:
                # Regular data item
                item_data = {}
                for month_info in month_columns:
                    month = month_info['month']
                    value = 0.0
                    if month_info['index'] < len(row):
                        value_str = row[month_info['index']].strip().replace(',', '').replace('$', '').replace('(', '-').replace(')', '')
                        try:
                            value = float(value_str) if value_str else 0.0
                        except:
                            value = 0.0

                    item_data[month] = {
                        'type': 'data',
                        'name': account_name,
                        'value': value,
                        'id': self.get_or_create_account_id(account_name)
                    }

                items.append(item_data)
                idx += 1

        return idx, items

    def parse_group_items(self, rows: List[List[str]], start_idx: int,
                         month_columns: List[Dict], group_name: str) -> Tuple[int, List[Dict]]:
        """Parse items within a group"""
        items = []
        idx = start_idx

        while idx < len(rows):
            row = rows[idx]

            if not row or not row[0]:
                idx += 1
                continue

            account_name = row[0].strip()
            if not account_name:
                idx += 1
                continue

            # Check if we've hit the end of this group
            if account_name.lower() == f"total for {group_name.lower()}":
                return idx + 1, items

            # Regular data item in group
            item_data = {}
            for month_info in month_columns:
                month = month_info['month']
                value = 0.0
                if month_info['index'] < len(row):
                    value_str = row[month_info['index']].strip().replace(',', '').replace('$', '').replace('(', '-').replace(')', '')
                    try:
                        value = float(value_str) if value_str else 0.0
                    except:
                        value = 0.0

                item_data[month] = {
                    'type': 'data',
                    'name': account_name,
                    'value': value,
                    'id': self.get_or_create_account_id(account_name)
                }

            items.append(item_data)
            idx += 1

        return idx, items

    def calculate_item_total(self, item: Dict) -> float:
        """Calculate total value for an item"""
        if item['type'] == 'data':
            return item.get('value', 0.0)
        elif item['type'] == 'group':
            return sum(self.calculate_item_total(sub_item) for sub_item in item.get('items', []))
        return 0.0

    def build_profit_loss_json(self, months: List[str], data_by_month: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Build the complete profit and loss JSON structure"""
        result = []

        for month in months:
            month_data = data_by_month[month]

            # Check if there's any data for this month
            has_data = len(month_data['sections']) > 0

            # Create the month object
            month_obj = {
                "month": month,
                "endDate": month_data['end_date'].strftime('%Y-%m-%d'),
                "report": self.create_report_structure(month_data, has_data),
                "startDate": month_data['start_date'].strftime('%Y-%m-%d')
            }

            result.append(month_obj)

        return result

    def create_report_structure(self, month_data: Dict[str, Any], has_data: bool) -> Dict[str, Any]:
        """Create the report structure for a single month"""
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000+00:00')

        report = {
            "header": {
                "time": timestamp,
                "reportName": "ProfitAndLoss",
                "dateMacro": None,
                "reportBasis": "ACCRUAL",
                "startPeriod": month_data['start_date'].strftime('%Y-%m-%d'),
                "endPeriod": month_data['end_date'].strftime('%Y-%m-%d'),
                "summarizeColumnsBy": "Total",
                "currency": "USD",
                "customer": None,
                "vendor": None,
                "employee": None,
                "item": None,
                "clazz": None,
                "department": None,
                "option": [
                    {"name": "AccountingStandard", "value": "GAAP"},
                    {"name": "NoReportData", "value": "false" if has_data else "true"}
                ]
            },
            "columns": {
                "column": [
                    {
                        "colTitle": "",
                        "colType": "Account",
                        "metaData": [{"name": "ColKey", "value": "account"}],
                        "columns": None
                    }
                ]
            },
            "rows": {"row": []}
        }

        if has_data:
            # Add the Total column
            report["columns"]["column"].append({
                "colTitle": "Total",
                "colType": "Money",
                "metaData": [{"name": "ColKey", "value": "total"}],
                "columns": None
            })

            # Build rows from sections
            rows = []
            for section in month_data['sections']:
                if section['type'] == 'section':
                    rows.append(self.build_section_row(section))
                elif section['type'] == 'calculated':
                    rows.append(self.create_row_object(
                        section['name'],
                        f"{section['value']:.2f}" if section['value'] != 0 else "0.00",
                        is_summary=True,
                        group=section.get('group')
                    ))
                elif section['type'] == 'data':
                    rows.append(self.create_row_object(
                        section['name'],
                        f"{section['value']:.2f}",
                        account_id=section.get('id')
                    ))

            report["rows"]["row"] = rows

        return report

    def build_section_row(self, section: Dict) -> Dict:
        """Build a section row with all its items"""
        sub_rows = []

        for item in section.get('items', []):
            if item['type'] == 'group':
                # Build group with sub-items
                group_rows = []
                for sub_item in item.get('items', []):
                    group_rows.append(self.create_row_object(
                        sub_item['name'],
                        f"{sub_item['value']:.2f}",
                        account_id=sub_item.get('id')
                    ))

                group_row = self.create_row_object(
                    item['name'],
                    "",
                    account_id=item.get('id'),
                    is_section=True,
                    sub_rows=group_rows
                )

                # Add group summary
                group_total = sum(sub['value'] for sub in item.get('items', []))
                group_row["summary"] = {
                    "colData": [
                        {"attributes": None, "value": f"Total {item['name']}", "id": None, "href": None},
                        {"attributes": None, "value": f"{group_total:.2f}", "id": None, "href": None}
                    ]
                }

                sub_rows.append(group_row)
            else:
                # Regular data item
                sub_rows.append(self.create_row_object(
                    item['name'],
                    f"{item['value']:.2f}",
                    account_id=item.get('id')
                ))

        # Create section row
        section_row = self.create_row_object(
            section['name'],
            "",
            is_section=True,
            sub_rows=sub_rows,
            group=section.get('group')
        )

        # Add section summary
        section_row["summary"] = {
            "colData": [
                {"attributes": None, "value": f"Total {section['name']}", "id": None, "href": None},
                {"attributes": None, "value": f"{section['total']:.2f}", "id": None, "href": None}
            ]
        }

        return section_row

    def parse_csv(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse CSV file and convert to profit and loss JSON"""
        months, data_by_month = self.parse_csv_hierarchy(filepath)
        return self.build_profit_loss_json(months, data_by_month)

    def parse_xlsx(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse XLSX file and convert to profit and loss JSON"""
        import openpyxl

        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active

        # Convert to list of lists for easier processing
        rows = []
        for row in sheet.iter_rows(values_only=True):
            # Convert row to list, handling None values
            row_data = []
            for cell in row:
                if cell is None:
                    row_data.append('')
                else:
                    row_data.append(str(cell))
            rows.append(row_data)

        # Find the header row with months
        header_row_idx = -1
        for i, row in enumerate(rows):
            if len(row) > 1:  # Must have at least 2 columns
                # Check each cell for month names
                for cell in row[1:]:  # Skip first column
                    if cell and any(month in str(cell) for month in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']):
                        header_row_idx = i
                        break
                if header_row_idx != -1:
                    break

        if header_row_idx == -1:
            raise ValueError("Could not find header row with months")

        # Parse header to get months
        header_row = rows[header_row_idx]
        month_columns = []
        months = []
        for i, part in enumerate(header_row[1:], 1):  # Skip first column
            if part.strip() and not any(skip in part.lower() for skip in ['total', 'ytd', 'year to date']):
                month_str, start_date, end_date = self.parse_month_column(part.strip())
                months.append(month_str)
                month_columns.append({
                    'index': i,
                    'month': month_str,
                    'start_date': start_date,
                    'end_date': end_date,
                    'header': part.strip()
                })

        # Initialize data structure for each month
        data_by_month = {}
        for month_info in month_columns:
            data_by_month[month_info['month']] = {
                'start_date': month_info['start_date'],
                'end_date': month_info['end_date'],
                'sections': []  # Will store hierarchical data
            }

        # Parse data rows and build hierarchy
        data_rows = rows[header_row_idx + 1:]
        self.parse_rows_recursive(data_rows, 0, month_columns, data_by_month)

        return self.build_profit_loss_json(months, data_by_month)

    def parse_pdf(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse PDF file and convert to profit and loss JSON"""
        import pdfplumber

        with pdfplumber.open(filepath) as pdf:
            # Extract text from all pages
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"

            # Split into lines
            lines = all_text.split('\n')

            # Find header line with months
            header_idx = -1
            year_line_idx = -1
            for i, line in enumerate(lines):
                # Skip lines that are date ranges (e.g., "January 1-July 31, 2025")
                if '-' in line and any(month in line.upper() for month in ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY']):
                    continue

                # Look for line that contains month names (case insensitive)
                if any(month in line.upper() for month in ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY']):
                    # Verify it's likely a header by checking for multiple months
                    month_count = sum(1 for month in ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY'] if month in line.upper())
                    if month_count >= 3:  # Require at least 3 months to be sure it's the header
                        header_idx = i
                        # Check if next line has years
                        if i + 1 < len(lines) and '2025' in lines[i + 1]:
                            year_line_idx = i + 1
                        break

            if header_idx == -1:
                raise ValueError("Could not find header row with months in PDF")

            # Parse months from header
            header_line = lines[header_idx]
            year_line = lines[year_line_idx] if year_line_idx != -1 else ""
            months = []
            month_columns = []

            # Extract month names and positions
            # Find all month patterns (case insensitive)
            # Include all 12 months to be comprehensive
            all_months = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY',
                         'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_pattern = r'(?i)(' + '|'.join(all_months) + ')'
            matches = list(re.finditer(month_pattern, header_line, re.IGNORECASE))

            # Extract years from year line if available
            year_pattern = r'\d{4}'
            year_matches = list(re.finditer(year_pattern, year_line)) if year_line else []

            for i, match in enumerate(matches):
                month_name = match.group()
                # Try to find corresponding year
                year = "2025"  # Default
                if i < len(year_matches):
                    year = year_matches[i].group()

                month_text = f"{month_name} {year}"
                month_str, start_date, end_date = self.parse_month_column(month_text)
                months.append(month_str)
                month_columns.append({
                    'text': month_text,
                    'month': month_str,
                    'start_date': start_date,
                    'end_date': end_date,
                    'start_pos': match.start(),
                    'end_pos': match.end(),
                    'index': i + 1  # Column index for data extraction
                })

            # Initialize data structure
            data_by_month = {}
            for month_info in month_columns:
                data_by_month[month_info['month']] = {
                    'start_date': month_info['start_date'],
                    'end_date': month_info['end_date'],
                    'sections': []
                }

            # Convert lines to row format for parsing
            rows = []
            start_line = header_idx + 2 if year_line_idx != -1 else header_idx + 1

            for line in lines[start_line:]:
                if not line.strip() or 'Page' in line:
                    rows.append([''])
                    continue

                # Extract account name (first part before numbers)
                # Find where numbers start
                number_match = re.search(r'[\-\$\d,\.]+', line)
                if number_match:
                    account_name = line[:number_match.start()].strip()
                    values_part = line[number_match.start():]
                else:
                    account_name = line.strip()
                    values_part = ""

                row_data = [account_name]

                # Extract all numbers from the values part
                if values_part:
                    numbers = re.findall(r'[\-\$]?[\d,]+\.?\d*', values_part)
                    # Remove dollar signs and clean up
                    cleaned_numbers = []
                    for num in numbers:
                        cleaned = num.replace('$', '').replace(',', '')
                        if cleaned and cleaned != '.':
                            cleaned_numbers.append(cleaned)

                    # Add numbers as columns
                    for num in cleaned_numbers[:len(month_columns)]:  # Only take as many as we have months
                        row_data.append(num)

                # Pad with empty strings if needed
                while len(row_data) < len(month_columns) + 1:
                    row_data.append('')

                rows.append(row_data)

            # Parse data rows and build hierarchy
            self.parse_rows_recursive(rows, 0, month_columns, data_by_month)

            return self.build_profit_loss_json(months, data_by_month)


def main():
    parser = argparse.ArgumentParser(description='Convert profit and loss documents to JSON format')
    parser.add_argument('input', help='Input file (CSV, XLSX, or PDF)')
    parser.add_argument('-o', '--output', help='Output JSON file (default: print to stdout)')

    args = parser.parse_args()

    converter = ProfitLossConverter()

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
