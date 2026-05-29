#!/usr/bin/env python3
"""
Cash Flow Statement Converter
Converts CSV, XLSX, and PDF cash flow reports to QuickBooks JSON format
Specifically designed for Cash Flow reports with monthly data
"""

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


class CashFlowConverter(BaseConverter):
    """Converts Cash Flow Statement documents to QuickBooks-style JSON format"""

    def __init__(self, use_account_lookup: bool = True, api_base_url: str = "http://localhost:8080"):
        super().__init__(use_account_lookup=use_account_lookup, api_base_url=api_base_url)

    def create_row_object(self, name: str, value: Optional[str] = None,
                         account_id: Optional[str] = None, row_type: str = "DATA",
                         group: Optional[str] = None, is_section: bool = False,
                         sub_rows: Optional[List] = None, is_summary: bool = False) -> Dict[str, Any]:
        """Create a row object for the cash flow statement"""
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

        if is_section:
            # Section header
            row["header"] = {
                "colData": [
                    {"attributes": None, "value": name, "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None}
                ]
            }
            if sub_rows:
                row["rows"] = {"row": sub_rows}
            row["colData"] = []
            row["type"] = "SECTION"
        elif is_summary:
            # Summary row
            row["summary"] = {
                "colData": [
                    {"attributes": None, "value": name, "id": None, "href": None},
                    {"attributes": None, "value": value if value else "", "id": None, "href": None}
                ]
            }
            row["colData"] = []
            row["type"] = "SECTION" if row_type == "SECTION" else None
        else:
            # Data row
            row["colData"] = [
                {"attributes": None, "value": name, "id": account_id, "href": None},
                {"attributes": None, "value": value if value else "", "id": None, "href": None}
            ]
            row["type"] = "DATA"

        return row

    def parse_csv_hierarchy(self, filepath: Path) -> Tuple[List[str], Dict[str, Dict[str, Any]]]:
        """Parse CSV file and extract hierarchical cash flow data"""
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
                if len(row) > 1 and (
                    'Full name' in row[0] or
                    sum(1 for cell in row[1:] if cell and (any(month in cell for month in full_months) or any(month in cell for month in abbr_months))) >= 2
                ):
                    header_row_idx = i
                    break

            if header_row_idx == -1:
                raise ValueError("Could not find header row with months")

            # Parse header to get months
            header_row = rows[header_row_idx]
            month_columns = []
            for i, part in enumerate(header_row[1:], 1):  # Skip first column
                if part.strip() and part.strip() != 'Total':
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
                    'operating': {
                        'net_income': None,
                        'adjustments': {},
                        'total_adjustments': 0.0,
                        'net_cash': 0.0
                    },
                    'investing': {
                        'items': {},
                        'net_cash': 0.0
                    },
                    'financing': {
                        'items': {},
                        'net_cash': 0.0
                    },
                    'net_increase': 0.0,
                    'beginning_cash': 0.0,
                    'ending_cash': 0.0
                }

            # Parse data rows
            current_section = None
            in_adjustments = False

            # Track running cash balance
            running_cash = 0.0

            for row_idx, row in enumerate(rows[header_row_idx + 1:]):
                if not row or not row[0] or 'Accrual Basis' in row[0]:
                    continue

                line_item = row[0].strip()

                if not line_item:
                    continue

                # Determine section
                if 'OPERATING ACTIVITIES' in line_item:
                    current_section = 'operating'
                    in_adjustments = False
                    continue
                elif 'INVESTING ACTIVITIES' in line_item:
                    current_section = 'investing'
                    in_adjustments = False
                    continue
                elif 'FINANCING ACTIVITIES' in line_item:
                    current_section = 'financing'
                    in_adjustments = False
                    continue
                elif 'Adjustments to reconcile' in line_item and not line_item.startswith('Total for'):
                    in_adjustments = True
                    continue
                elif 'Total for Adjustments' in line_item:
                    # Process total adjustments row
                    for month_info in month_columns:
                        if month_info['index'] < len(row):
                            value_str = row[month_info['index']].strip().replace(',', '').replace('$', '')
                            try:
                                value = float(value_str) if value_str else 0.0
                                data_by_month[month_info['month']]['operating']['total_adjustments'] = value
                            except ValueError:
                                pass
                    continue
                elif line_item.startswith('Net cash provided by'):
                    # Process net cash rows
                    for month_info in month_columns:
                        if month_info['index'] < len(row):
                            value_str = row[month_info['index']].strip().replace(',', '').replace('$', '')
                            try:
                                value = float(value_str) if value_str else 0.0
                                month = month_info['month']
                                if current_section == 'operating':
                                    data_by_month[month]['operating']['net_cash'] = value
                                elif current_section == 'investing':
                                    data_by_month[month]['investing']['net_cash'] = value
                                elif current_section == 'financing':
                                    data_by_month[month]['financing']['net_cash'] = value
                            except ValueError:
                                pass
                    continue
                elif 'NET CASH INCREASE FOR PERIOD' in line_item or 'Net cash increase for period' in line_item:
                    # Process net increase row
                    for month_idx, month_info in enumerate(month_columns):
                        if month_info['index'] < len(row):
                            value_str = row[month_info['index']].strip().replace(',', '').replace('$', '')
                            try:
                                value = float(value_str) if value_str else 0.0
                                month = month_info['month']
                                data_by_month[month]['net_increase'] = value

                                # Calculate cash positions
                                if month_idx == 0:
                                    # First month
                                    data_by_month[month]['beginning_cash'] = 0.0
                                    data_by_month[month]['ending_cash'] = value
                                    running_cash = value
                                else:
                                    # Subsequent months
                                    data_by_month[month]['beginning_cash'] = running_cash
                                    data_by_month[month]['ending_cash'] = running_cash + value
                                    running_cash = running_cash + value
                            except ValueError:
                                pass
                    continue

                # Process regular line items
                if current_section:
                    for month_info in month_columns:
                        if month_info['index'] < len(row):
                            value_str = row[month_info['index']].strip().replace(',', '').replace('$', '')
                            if value_str:  # Only process non-empty values
                                try:
                                    value = float(value_str)
                                except ValueError:
                                    continue

                                month = month_info['month']

                                if current_section == 'operating':
                                    if line_item == 'Net Income':
                                        data_by_month[month]['operating']['net_income'] = value
                                    elif in_adjustments:
                                        account_id = self.get_or_create_account_id(line_item)
                                        data_by_month[month]['operating']['adjustments'][line_item] = {
                                            'value': value,
                                            'id': account_id
                                        }
                                elif current_section == 'investing':
                                    account_id = self.get_or_create_account_id(line_item)
                                    data_by_month[month]['investing']['items'][line_item] = {
                                        'value': value,
                                        'id': account_id
                                    }
                                elif current_section == 'financing':
                                    account_id = self.get_or_create_account_id(line_item)
                                    data_by_month[month]['financing']['items'][line_item] = {
                                        'value': value,
                                        'id': account_id
                                    }

        return months, data_by_month

    def build_cash_flow_json(self, months: List[str], data_by_month: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Build the complete cash flow JSON structure"""
        result = []

        for month in months:
            month_data = data_by_month[month]

            # Check if there's any data for this month
            has_data = (
                month_data['operating']['net_income'] is not None or
                month_data['operating']['adjustments'] or
                month_data['investing']['items'] or
                month_data['financing']['items']
            )

            # Create the month object
            month_obj = {
                "month": month,
                "endDate": month_data['end_date'].strftime('%Y-%m-%d'),
                "startDate": month_data['start_date'].strftime('%Y-%m-%d'),
                "report": self.create_report_structure(month_data, has_data)
            }

            result.append(month_obj)

        return result

    def create_report_structure(self, month_data: Dict[str, Any], has_data: bool) -> Dict[str, Any]:
        """Create the report structure for a single month"""
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000+00:00')

        report = {
            "header": {
                "time": timestamp,
                "reportName": "CashFlow",
                "dateMacro": None,
                "reportBasis": None,
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
                    {"name": "NoReportData", "value": "false" if has_data else "true"}
                ]
            },
            "columns": {
                "column": [
                    {
                        "colTitle": "",
                        "colType": "Account",
                        "metaData": [],
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

            # Build the rows structure
            rows = []

            # OPERATING ACTIVITIES section
            operating_rows = self.build_operating_section(month_data['operating'])
            if operating_rows:
                rows.append(operating_rows)

            # INVESTING ACTIVITIES section
            if month_data['investing']['items']:
                investing_rows = self.build_investing_section(month_data['investing'])
                if investing_rows:
                    rows.append(investing_rows)

            # FINANCING ACTIVITIES section
            if month_data['financing']['items']:
                financing_rows = self.build_financing_section(month_data['financing'])
                if financing_rows:
                    rows.append(financing_rows)

            # NET CASH INCREASE FOR PERIOD
            net_increase_value = month_data['net_increase']
            net_increase_row = self.create_row_object(
                "Net cash increase for period",
                f"{net_increase_value:.2f}",
                is_summary=True,
                group="CashIncrease"
            )
            rows.append(net_increase_row)

            # Cash at beginning/end of period
            if month_data['beginning_cash'] != 0.0:
                beginning_cash_row = self.create_row_object(
                    "Cash at beginning of period",
                    f"{month_data['beginning_cash']:.2f}",
                    group="BeginningCash"
                )
                rows.append(beginning_cash_row)

            ending_cash_value = month_data['ending_cash']
            ending_cash_row = self.create_row_object(
                "Cash at end of period",
                f"{ending_cash_value:.2f}",
                is_summary=True,
                group="EndingCash"
            )
            rows.append(ending_cash_row)

            report["rows"]["row"] = rows
        else:
            # No data - create minimal structure
            rows = [
                self.create_row_object("OPERATING ACTIVITIES", is_section=True, group="OperatingActivities", sub_rows=[
                    self.create_row_object("Net Income", None, None, group="NetIncome"),
                    self.create_row_object("Adjustments to reconcile Net Income to Net Cash provided by operations:",
                                         None, None, group="OperatingAdjustments")
                ]),
                self.create_row_object("Net cash provided by operating activities", None, None, group="OperatingActivities"),
                self.create_row_object("Net cash increase for period", None, None, group="CashIncrease"),
                self.create_row_object("Cash at end of period", None, None, group="EndingCash")
            ]

            report["rows"]["row"] = rows

        return report

    def build_operating_section(self, operating_data: Dict[str, Any]) -> Dict[str, Any]:
        """Build the OPERATING ACTIVITIES section"""
        operating_rows = []

        # Net Income
        net_income_value = operating_data['net_income']
        if net_income_value is not None:
            operating_rows.append(self.create_row_object(
                "Net Income",
                f"{net_income_value:.2f}",
                group="NetIncome"
            ))

        # Adjustments
        if operating_data['adjustments']:
            adjustment_rows = []
            for account_name, account_data in operating_data['adjustments'].items():
                adjustment_rows.append(self.create_row_object(
                    account_name,
                    f"{account_data['value']:.2f}",
                    account_data['id']
                ))

            adjustments_section = self.create_row_object(
                "Adjustments to reconcile Net Income to Net Cash provided by operations:",
                is_section=True,
                sub_rows=adjustment_rows,
                group="OperatingAdjustments"
            )

            # Add summary for adjustments
            total_adjustments = operating_data['total_adjustments']
            adjustments_section["summary"] = {
                "colData": [
                    {"attributes": None, "value": "Total Adjustments to reconcile Net Income to Net Cash provided by operations:",
                     "id": None, "href": None},
                    {"attributes": None, "value": f"{total_adjustments:.2f}", "id": None, "href": None}
                ]
            }

            operating_rows.append(adjustments_section)
        else:
            # Empty adjustments header
            operating_rows.append(self.create_row_object(
                "Adjustments to reconcile Net Income to Net Cash provided by operations:",
                group="OperatingAdjustments"
            ))

        # Create main OPERATING ACTIVITIES section
        operating_section = self.create_row_object(
            "OPERATING ACTIVITIES",
            is_section=True,
            sub_rows=operating_rows,
            group="OperatingActivities"
        )

        # Add summary for net cash provided by operating activities
        net_cash = operating_data['net_cash']
        operating_section["summary"] = {
            "colData": [
                {"attributes": None, "value": "Net cash provided by operating activities",
                 "id": None, "href": None},
                {"attributes": None, "value": f"{net_cash:.2f}", "id": None, "href": None}
            ]
        }

        return operating_section

    def build_investing_section(self, investing_data: Dict[str, Any]) -> Dict[str, Any]:
        """Build the INVESTING ACTIVITIES section"""
        investing_rows = []

        # Add items
        for account_name, account_data in investing_data['items'].items():
            investing_rows.append(self.create_row_object(
                account_name,
                f"{account_data['value']:.2f}",
                account_data['id']
            ))

        # Create INVESTING ACTIVITIES section
        investing_section = self.create_row_object(
            "INVESTING ACTIVITIES",
            is_section=True,
            sub_rows=investing_rows,
            group="InvestingActivities"
        )

        # Add summary
        net_cash = investing_data['net_cash']
        investing_section["summary"] = {
            "colData": [
                {"attributes": None, "value": "Net cash provided by investing activities",
                 "id": None, "href": None},
                {"attributes": None, "value": f"{net_cash:.2f}", "id": None, "href": None}
            ]
        }

        return investing_section

    def build_financing_section(self, financing_data: Dict[str, Any]) -> Dict[str, Any]:
        """Build the FINANCING ACTIVITIES section"""
        financing_rows = []

        # Add items
        for account_name, account_data in financing_data['items'].items():
            financing_rows.append(self.create_row_object(
                account_name,
                f"{account_data['value']:.2f}",
                account_data['id']
            ))

        # Create FINANCING ACTIVITIES section
        financing_section = self.create_row_object(
            "FINANCING ACTIVITIES",
            is_section=True,
            sub_rows=financing_rows,
            group="FinancingActivities"
        )

        # Add summary
        net_cash = financing_data['net_cash']
        financing_section["summary"] = {
            "colData": [
                {"attributes": None, "value": "Net cash provided by financing activities",
                 "id": None, "href": None},
                {"attributes": None, "value": f"{net_cash:.2f}", "id": None, "href": None}
            ]
        }

        return financing_section

    def parse_csv(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse CSV file and convert to cash flow JSON"""
        months, data_by_month = self.parse_csv_hierarchy(filepath)
        return self.build_cash_flow_json(months, data_by_month)

    def parse_xlsx(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse XLSX file and convert to cash flow JSON"""
        if not XLSX_SUPPORT:
            raise ImportError("openpyxl is required for XLSX support. Install with: pip install openpyxl")

        import openpyxl
        workbook = openpyxl.load_workbook(filepath, data_only=True)
        sheet = workbook.active

        # Convert to list of lists for easier processing
        rows = []
        for row in sheet.iter_rows(values_only=True):
            rows.append(list(row))

        # Find the header row with months
        header_row_idx = -1
        full_months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        abbr_months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        
        for i, row in enumerate(rows):
            if row and len(row) > 1 and (
                (row[0] and 'Full name' in str(row[0])) or
                sum(1 for cell in row[1:] if cell and (any(month in str(cell) for month in full_months) or any(month in str(cell) for month in abbr_months))) >= 2
            ):
                header_row_idx = i
                break

        if header_row_idx == -1:
            raise ValueError("Could not find header row with months")

        # Convert rows to CSV-like format and reuse CSV parser logic
        # This is a bit of a hack, but it keeps the logic consistent
        temp_rows = []
        for row in rows:
            temp_row = []
            for cell in row:
                if cell is None:
                    temp_row.append('')
                else:
                    temp_row.append(str(cell))
            temp_rows.append(temp_row)

        # Process using the same logic as CSV
        months = []
        month_columns = []
        header_row = temp_rows[header_row_idx]

        for i, part in enumerate(header_row[1:], 1):  # Skip first column
            if part.strip() and part.strip() != 'Total':
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
                'operating': {
                    'net_income': None,
                    'adjustments': {},
                    'total_adjustments': 0.0,
                    'net_cash': 0.0
                },
                'investing': {
                    'items': {},
                    'net_cash': 0.0
                },
                'financing': {
                    'items': {},
                    'net_cash': 0.0
                },
                'net_increase': 0.0,
                'beginning_cash': 0.0,
                'ending_cash': 0.0
            }

        # Parse data rows (reuse logic from CSV parser)
        current_section = None
        in_adjustments = False
        running_cash = 0.0

        for row_idx, row in enumerate(temp_rows[header_row_idx + 1:]):
            if not row or not row[0] or 'Accrual Basis' in row[0]:
                continue

            line_item = row[0].strip()

            if not line_item:
                continue

            # Determine section
            if 'OPERATING ACTIVITIES' in line_item:
                current_section = 'operating'
                in_adjustments = False
                continue
            elif 'INVESTING ACTIVITIES' in line_item:
                current_section = 'investing'
                in_adjustments = False
                continue
            elif 'FINANCING ACTIVITIES' in line_item:
                current_section = 'financing'
                in_adjustments = False
                continue
            elif 'Adjustments to reconcile' in line_item and not line_item.startswith('Total for'):
                in_adjustments = True
                continue
            elif 'Total for Adjustments' in line_item:
                # Process total adjustments row
                for month_info in month_columns:
                    if month_info['index'] < len(row):
                        value_str = row[month_info['index']].strip().replace(',', '').replace('$', '')
                        try:
                            value = float(value_str) if value_str else 0.0
                            data_by_month[month_info['month']]['operating']['total_adjustments'] = value
                        except ValueError:
                            pass
                continue
            elif line_item.startswith('Net cash provided by'):
                # Process net cash rows
                for month_info in month_columns:
                    if month_info['index'] < len(row):
                        value_str = row[month_info['index']].strip().replace(',', '').replace('$', '')
                        try:
                            value = float(value_str) if value_str else 0.0
                            month = month_info['month']
                            if current_section == 'operating':
                                data_by_month[month]['operating']['net_cash'] = value
                            elif current_section == 'investing':
                                data_by_month[month]['investing']['net_cash'] = value
                            elif current_section == 'financing':
                                data_by_month[month]['financing']['net_cash'] = value
                        except ValueError:
                            pass
                continue
            elif 'NET CASH INCREASE FOR PERIOD' in line_item or 'Net cash increase for period' in line_item:
                # Process net increase row
                for month_idx, month_info in enumerate(month_columns):
                    if month_info['index'] < len(row):
                        value_str = row[month_info['index']].strip().replace(',', '').replace('$', '')
                        try:
                            value = float(value_str) if value_str else 0.0
                            month = month_info['month']
                            data_by_month[month]['net_increase'] = value

                            # Calculate cash positions
                            if month_idx == 0:
                                # First month
                                data_by_month[month]['beginning_cash'] = 0.0
                                data_by_month[month]['ending_cash'] = value
                                running_cash = value
                            else:
                                # Subsequent months
                                data_by_month[month]['beginning_cash'] = running_cash
                                data_by_month[month]['ending_cash'] = running_cash + value
                                running_cash = running_cash + value
                        except ValueError:
                            pass
                continue

            # Process regular line items
            if current_section:
                for month_info in month_columns:
                    if month_info['index'] < len(row):
                        value_str = row[month_info['index']].strip().replace(',', '').replace('$', '')
                        if value_str:  # Only process non-empty values
                            try:
                                value = float(value_str)
                            except ValueError:
                                continue

                            month = month_info['month']

                            if current_section == 'operating':
                                if line_item == 'Net Income':
                                    data_by_month[month]['operating']['net_income'] = value
                                elif in_adjustments:
                                    account_id = self.get_or_create_account_id(line_item)
                                    data_by_month[month]['operating']['adjustments'][line_item] = {
                                        'value': value,
                                        'id': account_id
                                    }
                            elif current_section == 'investing':
                                account_id = self.get_or_create_account_id(line_item)
                                data_by_month[month]['investing']['items'][line_item] = {
                                    'value': value,
                                    'id': account_id
                                }
                            elif current_section == 'financing':
                                account_id = self.get_or_create_account_id(line_item)
                                data_by_month[month]['financing']['items'][line_item] = {
                                    'value': value,
                                    'id': account_id
                                }

        return self.build_cash_flow_json(months, data_by_month)

    # ──────────────────────────────────────────────────────────────────────
    # PDF parsing helpers
    #
    # QB exports the multi-month Statement of Cash Flows as a border-less,
    # wide report that is split across pages HORIZONTALLY: each page covers a
    # subset of the months but repeats EVERY line item. Plain text extraction
    # collapses empty month columns, so values land in the wrong month. We
    # instead work from word positions: every amount is right-aligned under its
    # month column, so we snap each amount to the column whose right edge it is
    # closest to. Months are then stitched together across pages by month key.
    # ──────────────────────────────────────────────────────────────────────

    _PDF_MONTH_NAMES = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE',
                        'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']

    # An amount: optional sign / $ / parens, thousands-separated, two decimals.
    _PDF_AMOUNT_RE = re.compile(r'^-?\$?\(?-?[\d,]+\.\d{2}\)?\$?$')

    @classmethod
    def _pdf_is_amount(cls, text: str) -> bool:
        return bool(cls._PDF_AMOUNT_RE.match(text.strip()))

    @staticmethod
    def _pdf_group_lines(words: List[Dict[str, Any]], tol: float = 3.0) -> List[List[Dict[str, Any]]]:
        """Cluster extracted words into visual lines by vertical position."""
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

    def _pdf_build_columns(self, lines: List[List[Dict[str, Any]]]) -> Tuple[Optional[List[Dict[str, Any]]], Optional[float], int]:
        """Locate the "FULL NAME ... <months>" header and return month columns.

        QB right-aligns every amount under its month's year token, so we use the
        right edge (x1) of the year token as that column's anchor. Years may be
        inline ("MAY 2023") on the header line or wrapped onto the line below
        ("JANUARY" on the header, "2023" beneath it); we collect 4-digit year
        tokens from both lines, sort by x1, and pair them with the month-name
        words (sorted by x0) in left-to-right order. The optional grand-total
        column ("TOTAL") is returned separately so its amounts can be ignored.
        """
        header_idx = -1
        for i, line in enumerate(lines):
            joined = ' '.join(w['text'] for w in line).upper()
            if 'FULL' in joined and 'NAME' in joined:
                header_idx = i
                break
        if header_idx == -1:
            return None, None, -1

        header = lines[header_idx]
        year_line = lines[header_idx + 1] if header_idx + 1 < len(lines) else []

        month_words = [w for w in header if w['text'].upper() in self._PDF_MONTH_NAMES]
        month_words.sort(key=lambda w: w['x0'])

        year_words = [w for w in header if re.fullmatch(r'\d{4}', w['text'])]
        year_words += [w for w in year_line if re.fullmatch(r'\d{4}', w['text'])]
        year_words.sort(key=lambda w: w['x1'])

        columns: List[Dict[str, Any]] = []
        for month_word, year_word in zip(month_words, year_words):
            month_num = self._PDF_MONTH_NAMES.index(month_word['text'].upper()) + 1
            year = int(year_word['text'])
            month_str, start_date, end_date = self.parse_month_column(f"{month_word['text']} {year}")
            columns.append({
                'month': month_str,
                'start_date': start_date,
                'end_date': end_date,
                'x1': year_word['x1'],
            })

        total_words = [w for w in header if w['text'].upper() == 'TOTAL']
        total_x1 = total_words[0]['x1'] if total_words else None

        return columns, total_x1, header_idx

    def parse_pdf(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse PDF file and convert to cash flow JSON.

        Stitches the horizontally-split, multi-page QB report back together by
        snapping each right-aligned amount to its month column and keying all
        data by month so the same line item collected across pages merges.
        """
        if not PDF_SUPPORT:
            raise ImportError("pdfplumber is required for PDF support. Install with: pip install pdfplumber")

        import pdfplumber

        data_by_month: Dict[str, Dict[str, Any]] = {}
        # Preserve chronological discovery order across pages for cash roll-forward.
        month_order: List[str] = []

        def ensure_month(col: Dict[str, Any]) -> str:
            month = col['month']
            if month not in data_by_month:
                data_by_month[month] = {
                    'start_date': col['start_date'],
                    'end_date': col['end_date'],
                    'operating': {
                        'net_income': None,
                        'adjustments': {},
                        'total_adjustments': 0.0,
                        'net_cash': 0.0,
                    },
                    'investing': {'items': {}, 'net_cash': 0.0},
                    'financing': {'items': {}, 'net_cash': 0.0},
                    'net_increase': 0.0,
                    'beginning_cash': 0.0,
                    'ending_cash': 0.0,
                }
            return month

        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                if not words:
                    continue

                lines = self._pdf_group_lines(words)
                columns, total_x1, header_idx = self._pdf_build_columns(lines)
                if not columns:
                    continue

                anchors = [c['x1'] for c in columns]
                for col in columns:
                    month = ensure_month(col)
                    if month not in month_order:
                        month_order.append(month)

                def snap(amount_word: Dict[str, Any]) -> Optional[int]:
                    """Return the index of the month column this amount belongs to.

                    Snaps to the nearest column right edge; returns None when the
                    amount is closest to the grand-total column (which we drop).
                    """
                    x1 = amount_word['x1']
                    best = min(range(len(anchors)), key=lambda i: abs(anchors[i] - x1))
                    if total_x1 is not None and abs(total_x1 - x1) < abs(anchors[best] - x1):
                        return None
                    return best

                current_section = None
                in_adjustments = False

                # Start right after the header. A wrapped year sub-header line
                # ("2023 2023 ...") is skipped by the furniture filter below;
                # we do not blindly skip a fixed number of lines because some
                # pages carry their years inline on the header itself.
                for line in lines[header_idx + 1:]:
                    amount_words = [w for w in line if self._pdf_is_amount(w['text'])]
                    name_words = [w for w in line if not self._pdf_is_amount(w['text'])]
                    line_item = ' '.join(w['text'] for w in name_words).strip()

                    # Skip report furniture (footer with timestamp / GMTZ, year
                    # sub-header rows, blank lines).
                    if not line_item and not amount_words:
                        continue
                    if 'GMTZ' in line_item.upper():
                        continue
                    if re.fullmatch(r'(\d{4} ?)+', line_item):
                        continue

                    # Section transitions (these rows carry no amounts).
                    if 'OPERATING ACTIVITIES' in line_item:
                        current_section = 'operating'
                        in_adjustments = False
                        continue
                    elif 'INVESTING ACTIVITIES' in line_item:
                        current_section = 'investing'
                        in_adjustments = False
                        continue
                    elif 'FINANCING ACTIVITIES' in line_item:
                        current_section = 'financing'
                        in_adjustments = False
                        continue

                    # The adjustments header may share a line with amounts only
                    # for its "Total for ..." variant; the plain header has none.
                    is_total_adjustments = line_item.startswith('Total for Adjustments') or 'Total Adjustments' in line_item
                    if 'Adjustments to reconcile' in line_item and not is_total_adjustments:
                        in_adjustments = True
                        # Header line itself has no amounts; nothing to place.
                        if not amount_words:
                            continue

                    if not amount_words:
                        # A wrapped name/label tail with no values; nothing to place.
                        continue

                    is_net_cash = line_item.startswith('Net cash provided by')
                    is_net_increase = 'NET CASH INCREASE' in line_item.upper() or 'Net cash increase' in line_item

                    for amount_word in amount_words:
                        col_idx = snap(amount_word)
                        if col_idx is None:
                            continue
                        month = columns[col_idx]['month']
                        value = self.parse_amount(amount_word['text'])

                        if is_net_cash:
                            if current_section == 'operating':
                                data_by_month[month]['operating']['net_cash'] = value
                            elif current_section == 'investing':
                                data_by_month[month]['investing']['net_cash'] = value
                            elif current_section == 'financing':
                                data_by_month[month]['financing']['net_cash'] = value
                        elif is_net_increase:
                            data_by_month[month]['net_increase'] = value
                        elif is_total_adjustments:
                            data_by_month[month]['operating']['total_adjustments'] = value
                        elif current_section == 'operating':
                            if line_item == 'Net Income':
                                data_by_month[month]['operating']['net_income'] = value
                            elif in_adjustments:
                                account_id = self.get_or_create_account_id(line_item)
                                data_by_month[month]['operating']['adjustments'][line_item] = {
                                    'value': value,
                                    'id': account_id,
                                }
                        elif current_section == 'investing':
                            account_id = self.get_or_create_account_id(line_item)
                            data_by_month[month]['investing']['items'][line_item] = {
                                'value': value,
                                'id': account_id,
                            }
                        elif current_section == 'financing':
                            account_id = self.get_or_create_account_id(line_item)
                            data_by_month[month]['financing']['items'][line_item] = {
                                'value': value,
                                'id': account_id,
                            }

        # Roll cash balances forward in chronological month order, mirroring the
        # CSV/XLSX parsers (first month begins at 0; each subsequent month opens
        # at the prior month's ending balance).
        months = sorted(data_by_month.keys())
        running_cash = 0.0
        for idx, month in enumerate(months):
            increase = data_by_month[month]['net_increase']
            if idx == 0:
                data_by_month[month]['beginning_cash'] = 0.0
            else:
                data_by_month[month]['beginning_cash'] = running_cash
            running_cash = data_by_month[month]['beginning_cash'] + increase
            data_by_month[month]['ending_cash'] = running_cash

        return self.build_cash_flow_json(months, data_by_month)


def main():
    parser = argparse.ArgumentParser(description='Convert cash flow statement documents to JSON format')
    parser.add_argument('input', help='Input file (CSV, XLSX, or PDF)')
    parser.add_argument('-o', '--output', help='Output JSON file (default: print to stdout)')

    args = parser.parse_args()

    converter = CashFlowConverter()

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
