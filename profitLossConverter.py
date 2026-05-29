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
            full_months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
            abbr_months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            
            for i, row in enumerate(rows):
                if len(row) > 1:  # Must have at least 2 columns
                    # Check each cell for month names (full or abbreviated)
                    for cell in row[1:]:  # Skip first column
                        cell_str = str(cell).strip()
                        if cell_str and (any(month in cell_str for month in full_months) or any(month in cell_str for month in abbr_months)):
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
                idx, items = self.parse_section_items(rows, idx, month_columns, section_name=account_name)
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
                          month_columns: List[Dict], section_name: str = "") -> Tuple[int, List[Dict]]:
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
                total_name = account_name[len('total for '):].strip().lower()
                if not section_name or total_name == section_name.lower():
                    return idx + 1, items
                # Otherwise it's a sub-group total — skip it
                idx += 1
                continue

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

        workbook = openpyxl.load_workbook(filepath, data_only=True)
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
        full_months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        abbr_months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        
        for i, row in enumerate(rows):
            if len(row) > 1:  # Must have at least 2 columns
                # Check each cell for month names (full or abbreviated)
                for cell in row[1:]:  # Skip first column
                    cell_str = str(cell).strip()
                    if cell_str and (any(month in cell_str for month in full_months) or any(month in cell_str for month in abbr_months)):
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

    # ──────────────────────────────────────────────────────────────────
    # PDF parsing (wide multi-month "Profit and Loss by Month" reports)
    #
    # QuickBooks exports these as border-less PDFs where the 36 months do
    # not fit on a single page width. The report is therefore split into
    # horizontal PAGE-GROUPS: each group covers a subset of months for ALL
    # accounts, and the SAME accounts repeat in every group. We recover the
    # column layout from word x-positions (amounts are right-aligned under
    # each month header), stitch the months back together by matching
    # accounts across groups, then feed a reconstructed wide table through
    # the exact same hierarchy/build pipeline used by the CSV/XLSX parsers
    # so the output shape is identical.
    # ──────────────────────────────────────────────────────────────────

    _PDF_MONTHS = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY',
                   'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
    _PDF_AMOUNT_RE = re.compile(r'^-?\$?[\d,]+\.\d{2}$')

    @classmethod
    def _pdf_is_amount(cls, text: str) -> bool:
        return bool(cls._PDF_AMOUNT_RE.match(text.strip()))

    @staticmethod
    def _pdf_group_lines(words: List[Dict[str, Any]], tol: float = 3.0) -> List[List[Dict[str, Any]]]:
        """Cluster extracted words into visual rows by vertical position."""
        lines: List[List[Dict[str, Any]]] = []
        for w in sorted(words, key=lambda x: (round(x['top']), x['x0'])):
            placed = False
            for line in lines:
                if abs(line[0]['top'] - w['top']) <= tol:
                    line.append(w)
                    placed = True
                    break
            if not placed:
                lines.append([w])
        for line in lines:
            line.sort(key=lambda x: x['x0'])
        lines.sort(key=lambda ln: ln[0]['top'])
        return lines

    def _pdf_find_header(self, lines: List[List[Dict[str, Any]]]):
        """Locate the 'DISTRIBUTION ACCOUNT | <months>' header line.

        Returns (header_top, month_names, edges) where edges is the list of
        right-edge x positions used to snap amounts to a column. A trailing
        'TOTAL' column (present only on the last page-group) IS included as an
        extra edge so amounts snap correctly; month_names is padded with a
        None entry for it so the caller can discard that column's values.
        """
        for line in lines:
            joined = ' '.join(w['text'] for w in line).upper()
            # The header label can wrap so that "DISTRIBUTION" sits on the
            # month line while "ACCOUNT" drops to the year sub-line. Anchor on
            # "DISTRIBUTION" plus the presence of month names.
            if 'DISTRIBUTION' not in joined:
                continue
            if not any(m in joined for m in self._PDF_MONTHS):
                continue
            month_names: List[str] = []
            edges: List[float] = []
            for w in line:
                t = w['text'].upper()
                if t in self._PDF_MONTHS:
                    month_names.append(t)
                    edges.append(w['x1'])
                elif t == 'TOTAL':
                    # Keep an edge for the Total column so amounts align, but
                    # mark it so its values are dropped (CSV/XLSX omit Total).
                    month_names.append(None)
                    edges.append(w['x1'])
            if month_names:
                return line[0]['top'], month_names, edges
        return None, None, None

    def parse_pdf(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse a QuickBooks 'Profit and Loss by Month' PDF.

        Stitches month columns across horizontal page-groups, then reuses the
        CSV hierarchy pipeline so the JSON shape matches parse_csv/parse_xlsx.
        """
        import pdfplumber

        # Per-page extracted info: list of (month_names, edges, data_lines)
        # where data_lines is a list of (x0, name_words, amount_words).
        pages_info = []

        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                if not words:
                    continue

                lines = self._pdf_group_lines(words)
                header_top, month_names, edges = self._pdf_find_header(lines)
                if not month_names:
                    continue

                data_lines = []
                for line in lines:
                    top = line[0]['top']
                    # Skip header and everything above it (title / company /
                    # date-range / the header line and its year sub-line).
                    if top <= header_top + 12:
                        continue

                    amount_words = [w for w in line if self._pdf_is_amount(w['text'])]
                    name_words = [w for w in line if not self._pdf_is_amount(w['text'])]
                    name = ' '.join(w['text'] for w in name_words).strip()

                    # Report furniture: footers, page numbers (e.g. "1/15"),
                    # the "Accrual Basis ... GMTZ" stamp.
                    if not name and not amount_words:
                        continue
                    upper = name.upper()
                    if 'GMTZ' in upper or 'ACCRUAL BASIS' in upper or 'CASH BASIS' in upper:
                        continue
                    if re.fullmatch(r'\d+/\d+', name):
                        continue
                    if not name and amount_words:
                        # Stray amounts with no account name; ignore.
                        continue

                    x0 = name_words[0]['x0'] if name_words else (amount_words[0]['x0'] if amount_words else 0.0)
                    data_lines.append({
                        'top': top,
                        'x0': x0,
                        'name': name,
                        'amount_words': amount_words,
                    })

                pages_info.append({
                    'month_names': month_names,
                    'edges': edges,
                    'data_lines': data_lines,
                })

        if not pages_info:
            raise ValueError("Could not find header row with months in PDF")

        # ── Group consecutive pages that share the same month sequence ──
        page_groups = []
        for info in pages_info:
            if page_groups and page_groups[-1]['month_names'] == info['month_names']:
                page_groups[-1]['data_lines'].extend(info['data_lines'])
            else:
                page_groups.append({
                    'month_names': list(info['month_names']),
                    'edges': info['edges'],
                    'data_lines': list(info['data_lines']),
                })

        # ── Assign a calendar (month, year) to every group column ──
        # Months run sequentially across groups starting from the report's
        # first month. Determine the start year from the first group's first
        # month relative to the report; default to chronological inference.
        start_year = self._pdf_detect_start_year(filepath, page_groups[0]['month_names'][0])

        # Build the full ordered list of (iso_month, header_text) and a per
        # group mapping of column-index -> iso_month.
        cur_year = start_year
        prev_month_num = None
        ordered_months: List[str] = []
        month_meta: List[Tuple[str, date, date]] = []  # (iso, start, end) in order
        for group in page_groups:
            # col_iso is parallel to edges; the Total column (month_names entry
            # of None) maps to None so its snapped values are discarded.
            group['col_iso'] = []
            for mname in group['month_names']:
                if mname is None:
                    group['col_iso'].append(None)
                    continue
                month_num = self._PDF_MONTHS.index(mname) + 1
                if prev_month_num is not None and month_num <= prev_month_num:
                    cur_year += 1
                prev_month_num = month_num
                iso, sd, ed = self.parse_month_column(f"{mname} {cur_year}")
                group['col_iso'].append(iso)
                if iso not in ordered_months:
                    ordered_months.append(iso)
                    month_meta.append((iso, sd, ed))

        # ── Stitch months across groups by POSITIONAL alignment ──
        # Every page-group repeats the SAME accounts in the SAME reading order
        # (account names may legitimately repeat within a report, so we cannot
        # de-duplicate by name). We therefore merge wrapped names within each
        # group, then align the resulting line lists row-by-row across groups.
        # The group with the most lines defines the canonical row skeleton.
        merged_groups = []
        for group in page_groups:
            merged = self._pdf_merge_wrapped(group['data_lines'])
            merged_groups.append({
                'edges': group['edges'],
                'col_iso': group['col_iso'],
                'lines': merged,
            })

        skeleton_idx = max(range(len(merged_groups)),
                           key=lambda i: len(merged_groups[i]['lines']))
        skeleton = merged_groups[skeleton_idx]['lines']
        n_rows = len(skeleton)

        # Per-row, per-month values.
        row_values: List[Dict[str, float]] = [dict() for _ in range(n_rows)]

        for mg in merged_groups:
            lines = mg['lines']
            edges = mg['edges']
            col_iso = mg['col_iso']
            offset = self._pdf_align_offset(skeleton, lines)
            for li, entry in enumerate(lines):
                ri = li + offset
                if ri < 0 or ri >= n_rows:
                    continue
                for w in entry['amount_words']:
                    x1 = w['x1']
                    best = min(range(len(edges)), key=lambda i: abs(edges[i] - x1))
                    iso = col_iso[best]
                    if iso is None:  # Total column — not part of the output
                        continue
                    row_values[ri][iso] = self.parse_amount(w['text'])

        # ── Build a synthetic wide table identical in shape to the CSV ──
        month_columns = []
        months = []
        data_by_month = {}
        for idx, (iso, sd, ed) in enumerate(month_meta, start=1):
            months.append(iso)
            month_columns.append({
                'index': idx,
                'month': iso,
                'start_date': sd,
                'end_date': ed,
                'header': iso,
            })
            data_by_month[iso] = {
                'start_date': sd,
                'end_date': ed,
                'sections': [],
            }

        rows: List[List[str]] = []
        for ri, entry in enumerate(skeleton):
            row = [entry['name']]
            vals = row_values[ri]
            for iso in months:
                row.append(f"{vals[iso]:.2f}" if iso in vals else '')
            rows.append(row)

        # Reuse the exact CSV hierarchy pipeline.
        self.parse_rows_recursive(rows, 0, month_columns, data_by_month)

        return self.build_profit_loss_json(months, data_by_month)

    @staticmethod
    def _pdf_align_offset(skeleton: List[Dict[str, Any]], lines: List[Dict[str, Any]]) -> int:
        """Find the row offset that best aligns `lines` onto `skeleton`.

        Page-groups repeat the same accounts in the same order, so the
        alignment is normally offset 0. We still search a small window of
        offsets and pick the one that maximises account-name matches, which
        gracefully handles any group that drops leading/trailing furniture.
        """
        sk_names = [e['name'] for e in skeleton]
        ln_names = [e['name'] for e in lines]
        best_off, best_score = 0, -1
        for off in range(-3, 4):
            score = 0
            for i, nm in enumerate(ln_names):
                j = i + off
                if 0 <= j < len(sk_names) and sk_names[j] == nm:
                    score += 1
            if score > best_score:
                best_score, best_off = score, off
        return best_off

    @staticmethod
    def _pdf_merge_wrapped(data_lines: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Merge wrapped account-name continuation lines.

        QB wraps long account names onto a second line that carries no
        amounts, sits tightly under the previous line (vertical gap notably
        smaller than a normal row), and at the same or deeper indent. Such a
        line is appended to the previous account's name. All other name-only
        lines (section headers, all-zero data rows) are kept standalone.
        """
        WRAP_GAP = 11.5
        merged: List[Dict[str, Any]] = []
        for ln in data_lines:
            gap = ln['top'] - merged[-1]['top'] if merged else 0
            if (merged and not ln['amount_words']
                    and 0 < gap <= WRAP_GAP
                    and ln['x0'] >= merged[-1]['x0'] - 1
                    and not ln['name'].lower().startswith('total ')):
                merged[-1]['name'] = (merged[-1]['name'] + ' ' + ln['name']).strip()
                merged[-1]['top'] = ln['top']
            else:
                merged.append(dict(ln))
        return merged

    def _pdf_detect_start_year(self, filepath: Path, first_month_name: str) -> int:
        """Determine the starting calendar year for the first month-group.

        Reads the report date-range line (e.g. "January 1, 2023-December 31,
        2025") to find the earliest year. Falls back to the latest 4-digit
        year minus a best-effort offset, then to the current year.
        """
        import pdfplumber
        try:
            with pdfplumber.open(filepath) as pdf:
                text = pdf.pages[0].extract_text() or ""
        except Exception:
            text = ""
        years = [int(y) for y in re.findall(r'\b(19|20)\d{2}\b', text)]
        # re.findall with the group above returns the prefix; redo properly:
        years = [int(y) for y in re.findall(r'\b(?:19|20)\d{2}\b', text)]
        if years:
            return min(years)
        return datetime.now().year


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
