#!/usr/bin/env python3
"""
Balance Sheet Document Converter
Converts CSV, XLSX, and PDF balance sheet reports to QuickBooks JSON format
Specifically designed for Balance Sheet reports with monthly data
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

# Conditional imports for type checking in parse methods
if XLSX_SUPPORT:
    import openpyxl

if PDF_SUPPORT:
    import pdfplumber


class BalanceSheetConverter(BaseConverter):
    """Converts Balance Sheet documents to QuickBooks-style JSON format"""

    def __init__(self, use_account_lookup: bool = True, api_base_url: str = "http://localhost:8080"):
        super().__init__(use_account_lookup=use_account_lookup, api_base_url=api_base_url)

    def create_row_object(self, name: str, value: Optional[str] = None,
                         account_id: Optional[str] = None, row_type: str = "DATA",
                         group: Optional[str] = None, is_section: bool = False,
                         sub_rows: Optional[List] = None) -> Dict[str, Any]:
        """Create a row object for the balance sheet"""
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
        else:
            # Data row
            row["colData"] = [
                {"attributes": None, "value": name, "id": account_id, "href": None},
                {"attributes": None, "value": value if value else "", "id": None, "href": None}
            ]
            row["type"] = "DATA"

        return row

    def parse_csv_hierarchy(self, filepath: Path) -> Tuple[List[str], Dict[str, Dict[str, Any]]]:
        """Parse CSV file and extract hierarchical balance sheet data"""
        months = []
        data_by_month = {}

        with open(filepath, 'r', encoding='utf-8') as f:
            # Use csv reader to handle quoted fields properly
            reader = csv.reader(f)
            rows = list(reader)

            # Find the header row with months
            header_row_idx = -1
            for i, row in enumerate(rows):
                if len(row) > 0 and ('Distribution account' in row[0] or any(month in ' '.join(row) for month in ['January', 'February', 'March'])):
                    header_row_idx = i
                    break

            if header_row_idx == -1:
                raise ValueError("Could not find header row with months")

            # Parse header to get months
            header_row = rows[header_row_idx]
            month_columns = []
            for i, part in enumerate(header_row[1:], 1):  # Skip first column
                if part.strip():
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
                    'assets': {},
                    'liabilities': {},
                    'equity': {}
                }

            # Parse data rows
            current_section = None
            current_subsection = None
            current_group = None

            for row in rows[header_row_idx + 1:]:
                if not row or not row[0] or 'Accrual Basis' in row[0]:
                    continue

                account_name = row[0].strip()

                if not account_name:
                    continue

                # Determine hierarchy level based on account name
                if account_name in ['Assets', 'Liabilities and Equity']:
                    current_section = account_name.lower().replace(' and ', '_').replace(' ', '_')
                    continue
                elif account_name in ['Current Assets', 'Fixed Assets', 'Other Assets',
                                    'Liabilities', 'Equity', 'Current Liabilities',
                                    'Long-term Liabilities']:
                    current_subsection = account_name
                    continue
                elif account_name.startswith('Total for '):
                    # Skip total rows for now, we'll calculate them
                    continue
                elif any(account_name == cat for cat in ['Bank Accounts', 'Accounts Receivable',
                                                         'Other Current Assets', 'Accounts Payable',
                                                         'Credit Cards', 'Other Current Liabilities',
                                                         'Truck']):
                    current_group = account_name
                    continue

                # This is an actual account line
                for month_info in month_columns:
                    if month_info['index'] < len(row):
                        value_str = row[month_info['index']].strip().replace(',', '').replace('$', '')
                        try:
                            value = float(value_str) if value_str else 0.0
                        except ValueError:
                            value = 0.0

                        if value != 0.0 or account_name in ['Retained Earnings', 'Net Income']:
                            # Store the account data
                            month = month_info['month']
                            if current_section == 'assets':
                                section_data = data_by_month[month]['assets']
                            elif current_section == 'liabilities_equity':
                                if current_subsection == 'Equity':
                                    section_data = data_by_month[month]['equity']
                                else:
                                    section_data = data_by_month[month]['liabilities']
                            else:
                                continue

                            if current_subsection not in section_data:
                                section_data[current_subsection] = {}
                            if current_group and current_group not in section_data[current_subsection]:
                                section_data[current_subsection][current_group] = {}

                            # Get account ID from lookup or generate one
                            account_id = self.get_or_create_account_id(account_name)

                            if current_group:
                                section_data[current_subsection][current_group][account_name] = {
                                    'value': value,
                                    'id': account_id
                                }
                            else:
                                section_data[current_subsection][account_name] = {
                                    'value': value,
                                    'id': account_id
                                }

        return months, data_by_month

    def parse_csv(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse CSV file and convert to balance sheet JSON"""
        months, data_by_month = self.parse_csv_hierarchy(filepath)
        return self.build_balance_sheet_json(months, data_by_month)

    def build_balance_sheet_json(self, months: List[str], data_by_month: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Build the complete balance sheet JSON structure"""
        result = []

        for month in months:
            month_data = data_by_month[month]

            # Check if there's any data for this month
            has_data = any(
                month_data['assets'] or
                month_data['liabilities'] or
                month_data['equity']
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
                "reportName": "BalanceSheet",
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

            # Build the rows structure
            rows = []

            # ASSETS section
            assets_rows = self.build_assets_section(month_data['assets'])
            if assets_rows:
                rows.append(assets_rows)

            # LIABILITIES AND EQUITY section
            liabilities_equity_rows = self.build_liabilities_equity_section(
                month_data['liabilities'],
                month_data['equity']
            )
            if liabilities_equity_rows:
                rows.append(liabilities_equity_rows)

            report["rows"]["row"] = rows
        else:
            # No data - create minimal structure
            rows = [
                self.create_row_object("ASSETS", is_section=True, group="TotalAssets"),
                self.create_row_object("LIABILITIES AND EQUITY", is_section=True,
                                     group="TotalLiabilitiesAndEquity", sub_rows=[
                    self.create_row_object("Liabilities", is_section=True, group="Liabilities"),
                    self.create_row_object("Equity", is_section=True, group="Equity", sub_rows=[
                        self.create_row_object("Retained Earnings", None, None),
                        self.create_row_object("Net Income", None, None, group="NetIncome")
                    ]),
                    self.create_row_object("Total Equity", None, None, group="Equity")
                ])
            ]

            # Add the total row
            rows.append(self.create_row_object("TOTAL LIABILITIES AND EQUITY", None, None,
                                             row_type=None, group="TotalLiabilitiesAndEquity"))

            report["rows"]["row"] = rows

        return report

    def build_assets_section(self, assets_data: Dict[str, Any]) -> Dict[str, Any]:
        """Build the ASSETS section of the balance sheet"""
        assets_rows = []
        total_assets = 0.0

        # Current Assets
        if 'Current Assets' in assets_data:
            current_assets_rows = []
            current_assets_total = 0.0

            # Bank Accounts
            if 'Bank Accounts' in assets_data['Current Assets']:
                bank_rows = []
                bank_total = 0.0
                for account, data in assets_data['Current Assets']['Bank Accounts'].items():
                    bank_rows.append(self.create_row_object(
                        account, f"{data['value']:.2f}", data['id']
                    ))
                    bank_total += data['value']

                if bank_rows:
                    bank_section = self.create_row_object(
                        "Bank Accounts", is_section=True, sub_rows=bank_rows
                    )
                    bank_section["summary"] = {
                        "colData": [
                            {"attributes": None, "value": "Total Bank Accounts", "id": None, "href": None},
                            {"attributes": None, "value": f"{bank_total:.2f}", "id": None, "href": None}
                        ]
                    }
                    bank_section["group"] = "BankAccounts"
                    current_assets_rows.append(bank_section)
                    current_assets_total += bank_total

            # Accounts Receivable
            if 'Accounts Receivable' in assets_data['Current Assets']:
                ar_rows = []
                ar_total = 0.0
                for account, data in assets_data['Current Assets']['Accounts Receivable'].items():
                    ar_rows.append(self.create_row_object(
                        account, f"{data['value']:.2f}", data['id']
                    ))
                    ar_total += data['value']

                if ar_rows:
                    ar_section = self.create_row_object(
                        "Accounts Receivable", is_section=True, sub_rows=ar_rows
                    )
                    ar_section["summary"] = {
                        "colData": [
                            {"attributes": None, "value": "Total Accounts Receivable", "id": None, "href": None},
                            {"attributes": None, "value": f"{ar_total:.2f}", "id": None, "href": None}
                        ]
                    }
                    ar_section["group"] = "AR"
                    current_assets_rows.append(ar_section)
                    current_assets_total += ar_total

            # Other Current Assets
            if 'Other Current Assets' in assets_data['Current Assets']:
                other_rows = []
                other_total = 0.0
                for account, data in assets_data['Current Assets']['Other Current Assets'].items():
                    other_rows.append(self.create_row_object(
                        account, f"{data['value']:.2f}", data['id']
                    ))
                    other_total += data['value']

                if other_rows:
                    other_section = self.create_row_object(
                        "Other Current Assets", is_section=True, sub_rows=other_rows
                    )
                    other_section["summary"] = {
                        "colData": [
                            {"attributes": None, "value": "Total Other Current Assets", "id": None, "href": None},
                            {"attributes": None, "value": f"{other_total:.2f}", "id": None, "href": None}
                        ]
                    }
                    other_section["group"] = "OtherCurrentAssets"
                    current_assets_rows.append(other_section)
                    current_assets_total += other_total

            if current_assets_rows:
                current_assets_section = self.create_row_object(
                    "Current Assets", is_section=True, sub_rows=current_assets_rows
                )
                current_assets_section["summary"] = {
                    "colData": [
                        {"attributes": None, "value": "Total Current Assets", "id": None, "href": None},
                        {"attributes": None, "value": f"{current_assets_total:.2f}", "id": None, "href": None}
                    ]
                }
                current_assets_section["group"] = "CurrentAssets"
                assets_rows.append(current_assets_section)
                total_assets += current_assets_total

        # Fixed Assets
        if 'Fixed Assets' in assets_data:
            fixed_assets_rows = []
            fixed_assets_total = 0.0

            for group_name, group_data in assets_data['Fixed Assets'].items():
                if isinstance(group_data, dict) and group_name == 'Truck':
                    truck_rows = []
                    truck_total = 0.0
                    for account, data in group_data.items():
                        truck_rows.append(self.create_row_object(
                            account, f"{data['value']:.2f}", data['id']
                        ))
                        truck_total += data['value']

                    if truck_rows:
                        truck_section = self.create_row_object(
                            "Truck", is_section=True, sub_rows=truck_rows
                        )
                        # Don't hardcode IDs
                        truck_section["summary"] = {
                            "colData": [
                                {"attributes": None, "value": "Total Truck", "id": None, "href": None},
                                {"attributes": None, "value": f"{truck_total:.2f}", "id": None, "href": None}
                            ]
                        }
                        fixed_assets_rows.append(truck_section)
                        fixed_assets_total += truck_total

            if fixed_assets_rows:
                fixed_assets_section = self.create_row_object(
                    "Fixed Assets", is_section=True, sub_rows=fixed_assets_rows
                )
                fixed_assets_section["summary"] = {
                    "colData": [
                        {"attributes": None, "value": "Total Fixed Assets", "id": None, "href": None},
                        {"attributes": None, "value": f"{fixed_assets_total:.2f}", "id": None, "href": None}
                    ]
                }
                fixed_assets_section["group"] = "FixedAssets"
                assets_rows.append(fixed_assets_section)
                total_assets += fixed_assets_total

        # Create main ASSETS section
        assets_section = self.create_row_object(
            "ASSETS", is_section=True, sub_rows=assets_rows
        )
        assets_section["summary"] = {
            "colData": [
                {"attributes": None, "value": "TOTAL ASSETS", "id": None, "href": None},
                {"attributes": None, "value": f"{total_assets:.2f}", "id": None, "href": None}
            ]
        }
        assets_section["group"] = "TotalAssets"

        return assets_section

    def build_liabilities_equity_section(self, liabilities_data: Dict[str, Any],
                                       equity_data: Dict[str, Any]) -> Dict[str, Any]:
        """Build the LIABILITIES AND EQUITY section"""
        le_rows = []

        # Liabilities section
        liabilities_rows = []
        total_liabilities = 0.0

        # Current Liabilities
        if 'Current Liabilities' in liabilities_data:
            current_liab_rows = []
            current_liab_total = 0.0

            # Accounts Payable
            if 'Accounts Payable' in liabilities_data['Current Liabilities']:
                ap_rows = []
                ap_total = 0.0
                for account, data in liabilities_data['Current Liabilities']['Accounts Payable'].items():
                    ap_rows.append(self.create_row_object(
                        account, f"{data['value']:.2f}", data['id']
                    ))
                    ap_total += data['value']

                if ap_rows:
                    ap_section = self.create_row_object(
                        "Accounts Payable", is_section=True, sub_rows=ap_rows
                    )
                    ap_section["summary"] = {
                        "colData": [
                            {"attributes": None, "value": "Total Accounts Payable", "id": None, "href": None},
                            {"attributes": None, "value": f"{ap_total:.2f}", "id": None, "href": None}
                        ]
                    }
                    ap_section["group"] = "AP"
                    current_liab_rows.append(ap_section)
                    current_liab_total += ap_total

            # Credit Cards
            if 'Credit Cards' in liabilities_data['Current Liabilities']:
                cc_rows = []
                cc_total = 0.0
                for account, data in liabilities_data['Current Liabilities']['Credit Cards'].items():
                    cc_rows.append(self.create_row_object(
                        account, f"{data['value']:.2f}", data['id']
                    ))
                    cc_total += data['value']

                if cc_rows:
                    cc_section = self.create_row_object(
                        "Credit Cards", is_section=True, sub_rows=cc_rows
                    )
                    cc_section["summary"] = {
                        "colData": [
                            {"attributes": None, "value": "Total Credit Cards", "id": None, "href": None},
                            {"attributes": None, "value": f"{cc_total:.2f}", "id": None, "href": None}
                        ]
                    }
                    cc_section["group"] = "CreditCards"
                    current_liab_rows.append(cc_section)
                    current_liab_total += cc_total

            # Other Current Liabilities
            if 'Other Current Liabilities' in liabilities_data['Current Liabilities']:
                other_rows = []
                other_total = 0.0
                for account, data in liabilities_data['Current Liabilities']['Other Current Liabilities'].items():
                    other_rows.append(self.create_row_object(
                        account, f"{data['value']:.2f}", data['id']
                    ))
                    other_total += data['value']

                if other_rows:
                    other_section = self.create_row_object(
                        "Other Current Liabilities", is_section=True, sub_rows=other_rows
                    )
                    other_section["summary"] = {
                        "colData": [
                            {"attributes": None, "value": "Total Other Current Liabilities", "id": None, "href": None},
                            {"attributes": None, "value": f"{other_total:.2f}", "id": None, "href": None}
                        ]
                    }
                    other_section["group"] = "OtherCurrentLiabilities"
                    current_liab_rows.append(other_section)
                    current_liab_total += other_total

            if current_liab_rows:
                current_liab_section = self.create_row_object(
                    "Current Liabilities", is_section=True, sub_rows=current_liab_rows
                )
                current_liab_section["summary"] = {
                    "colData": [
                        {"attributes": None, "value": "Total Current Liabilities", "id": None, "href": None},
                        {"attributes": None, "value": f"{current_liab_total:.2f}", "id": None, "href": None}
                    ]
                }
                current_liab_section["group"] = "CurrentLiabilities"
                liabilities_rows.append(current_liab_section)
                total_liabilities += current_liab_total

        # Long-Term Liabilities
        if 'Long-term Liabilities' in liabilities_data:
            lt_rows = []
            lt_total = 0.0
            lt_items = liabilities_data['Long-term Liabilities']

            # Handle both direct accounts and grouped accounts
            for account, data in lt_items.items():
                if isinstance(data, dict) and 'value' in data:
                    lt_rows.append(self.create_row_object(
                        account, f"{data['value']:.2f}", data['id']
                    ))
                    lt_total += data['value']

            if lt_rows:
                lt_section = self.create_row_object(
                    "Long-Term Liabilities", is_section=True, sub_rows=lt_rows
                )
                lt_section["summary"] = {
                    "colData": [
                        {"attributes": None, "value": "Total Long-Term Liabilities", "id": None, "href": None},
                        {"attributes": None, "value": f"{lt_total:.2f}", "id": None, "href": None}
                    ]
                }
                lt_section["group"] = "LongTermLiabilities"
                liabilities_rows.append(lt_section)
                total_liabilities += lt_total

        if liabilities_rows:
            liabilities_section = self.create_row_object(
                "Liabilities", is_section=True, sub_rows=liabilities_rows
            )
            liabilities_section["summary"] = {
                "colData": [
                    {"attributes": None, "value": "Total Liabilities", "id": None, "href": None},
                    {"attributes": None, "value": f"{total_liabilities:.2f}", "id": None, "href": None}
                ]
            }
            liabilities_section["group"] = "Liabilities"
            le_rows.append(liabilities_section)

        # Equity section
        equity_rows = []
        total_equity = 0.0

        if 'Equity' in equity_data:
            equity_items = equity_data['Equity']

            # Handle both direct accounts and grouped accounts
            for key, value in equity_items.items():
                if isinstance(value, dict) and 'value' in value:
                    # Direct account with value
                    if key != 'Net Income':  # Net Income is handled separately
                        equity_rows.append(self.create_row_object(
                            key, f"{value['value']:.2f}" if value['value'] != 0 or key == 'Retained Earnings' else f"{value['value']:.2f}",
                            value['id']
                        ))
                    total_equity += value['value']

        # Always add Net Income row
        net_income_row = self.create_row_object("Net Income", None, None, group="NetIncome")
        if 'Equity' in equity_data and 'Net Income' in equity_data['Equity']:
            net_income_data = equity_data['Equity']['Net Income']
            if isinstance(net_income_data, dict) and 'value' in net_income_data:
                net_income_value = net_income_data['value']
                net_income_row["colData"][1]["value"] = f"{net_income_value:.2f}"
                total_equity += net_income_value
        equity_rows.append(net_income_row)

        if equity_rows:
            equity_section = self.create_row_object(
                "Equity", is_section=True, sub_rows=equity_rows
            )
            equity_section["summary"] = {
                "colData": [
                    {"attributes": None, "value": "Total Equity", "id": None, "href": None},
                    {"attributes": None, "value": f"{total_equity:.2f}", "id": None, "href": None}
                ]
            }
            equity_section["group"] = "Equity"
            le_rows.append(equity_section)

        # Create main LIABILITIES AND EQUITY section
        le_section = self.create_row_object(
            "LIABILITIES AND EQUITY", is_section=True, sub_rows=le_rows
        )
        le_section["summary"] = {
            "colData": [
                {"attributes": None, "value": "TOTAL LIABILITIES AND EQUITY", "id": None, "href": None},
                {"attributes": None, "value": f"{total_liabilities + total_equity:.2f}", "id": None, "href": None}
            ]
        }
        le_section["group"] = "TotalLiabilitiesAndEquity"

        return le_section

    def parse_xlsx(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse XLSX file and convert to balance sheet JSON"""
        if not XLSX_SUPPORT:
            raise ImportError("openpyxl is required for XLSX support. Install with: pip install openpyxl")

        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active

        # Convert to list of lists for easier processing
        rows = []
        for row in sheet.iter_rows(values_only=True):
            rows.append(list(row))

        # Find the header row with months
        header_row_idx = -1
        for i, row in enumerate(rows):
            if row and row[0] and ('Distribution account' in str(row[0]) or
                                  any(month in ' '.join(str(cell) for cell in row if cell)
                                      for month in ['January', 'February', 'March'])):
                header_row_idx = i
                break

        if header_row_idx == -1:
            raise ValueError("Could not find header row with months")

        # Parse using the same logic as CSV
        months = []
        month_columns = []
        header_row = rows[header_row_idx]

        for i, cell in enumerate(header_row[1:], 1):  # Skip first column
            if cell:
                month_str, start_date, end_date = self.parse_month_column(str(cell).strip())
                months.append(month_str)
                month_columns.append({
                    'index': i,
                    'month': month_str,
                    'start_date': start_date,
                    'end_date': end_date,
                    'header': str(cell).strip()
                })

        # Initialize data structure
        data_by_month = {}
        for month_info in month_columns:
            data_by_month[month_info['month']] = {
                'start_date': month_info['start_date'],
                'end_date': month_info['end_date'],
                'assets': {},
                'liabilities': {},
                'equity': {}
            }

        # Parse data rows (reuse the same logic from CSV parser)
        current_section = None
        current_subsection = None
        current_group = None

        for row in rows[header_row_idx + 1:]:
            if not row or not row[0] or 'Accrual Basis' in str(row[0]):
                continue

            account_name = str(row[0]).strip()

            if not account_name:
                continue

            # Determine hierarchy level based on account name
            if account_name in ['Assets', 'Liabilities and Equity']:
                current_section = account_name.lower().replace(' and ', '_').replace(' ', '_')
                continue
            elif account_name in ['Current Assets', 'Fixed Assets', 'Other Assets',
                                'Liabilities', 'Equity', 'Current Liabilities',
                                'Long-term Liabilities']:
                current_subsection = account_name
                continue
            elif account_name.startswith('Total for '):
                continue
            elif any(account_name == cat for cat in ['Bank Accounts', 'Accounts Receivable',
                                                     'Other Current Assets', 'Accounts Payable',
                                                     'Credit Cards', 'Other Current Liabilities',
                                                     'Truck']):
                current_group = account_name
                continue

            # Process account values
            for month_info in month_columns:
                if month_info['index'] < len(row) and row[month_info['index']] is not None:
                    value_str = str(row[month_info['index']]).strip().replace(',', '').replace('$', '')
                    try:
                        value = float(value_str) if value_str else 0.0
                    except ValueError:
                        value = 0.0

                    if value != 0.0 or account_name in ['Retained Earnings', 'Net Income']:
                        # Store the account data
                        month = month_info['month']
                        if current_section == 'assets':
                            section_data = data_by_month[month]['assets']
                        elif current_section == 'liabilities_equity':
                            if current_subsection == 'Equity':
                                section_data = data_by_month[month]['equity']
                            else:
                                section_data = data_by_month[month]['liabilities']
                        else:
                            continue

                        if current_subsection not in section_data:
                            section_data[current_subsection] = {}
                        if current_group and current_group not in section_data[current_subsection]:
                            section_data[current_subsection][current_group] = {}

                        # Get account ID from lookup or generate one
                        account_id = self.get_or_create_account_id(account_name)

                        if current_group:
                            section_data[current_subsection][current_group][account_name] = {
                                'value': value,
                                'id': account_id
                            }
                        else:
                            section_data[current_subsection][account_name] = {
                                'value': value,
                                'id': account_id
                            }

        return self.build_balance_sheet_json(months, data_by_month)

    def parse_pdf(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse PDF file and convert to balance sheet JSON"""
        if not PDF_SUPPORT:
            raise ImportError("pdfplumber is required for PDF support. Install with: pip install pdfplumber")

        with pdfplumber.open(filepath) as pdf:
            # Extract text from all pages
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"

            # Split into lines for processing
            lines = all_text.split('\n')

            # Find header line with months
            header_idx = -1
            for i, line in enumerate(lines):
                # Look for line that contains month names (case insensitive)
                line_upper = line.upper()
                if any(month in line_upper for month in ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY']):
                    # Verify it's likely a header by checking for multiple months
                    month_count = sum(1 for month in ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY'] if month in line_upper)
                    if month_count >= 2:
                        header_idx = i
                        break

            if header_idx == -1:
                raise ValueError("Could not find header row with months in PDF")

            # Parse months from header
            header_line = lines[header_idx]
            months = []
            month_columns = []

            # Extract month names and positions
            # Find all month patterns (case insensitive)
            month_pattern = r'(?i)(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\s+\d{4}|[A-Z]{3}\s+\d+\s*-\s*[A-Z]{3}\s+\d+\s+\d{4}'
            matches = list(re.finditer(month_pattern, header_line, re.IGNORECASE))

            for i, match in enumerate(matches):
                month_text = match.group()
                month_str, start_date, end_date = self.parse_month_column(month_text)
                months.append(month_str)
                month_columns.append({
                    'text': month_text,
                    'month': month_str,
                    'start_date': start_date,
                    'end_date': end_date,
                    'start_pos': match.start(),
                    'end_pos': match.end()
                })

            # Initialize data structure
            data_by_month = {}
            for month_info in month_columns:
                data_by_month[month_info['month']] = {
                    'start_date': month_info['start_date'],
                    'end_date': month_info['end_date'],
                    'assets': {},
                    'liabilities': {},
                    'equity': {}
                }

            # Parse data lines
            current_section = None
            current_subsection = None
            current_group = None

            for line_idx in range(header_idx + 1, len(lines)):
                line = lines[line_idx].strip()

                if not line or 'Page' in line:
                    continue

                # Extract account name (usually the first part before numbers)
                # Find where numbers start
                number_match = re.search(r'[\d,\.\-\$\s]+$', line)
                if number_match:
                    account_name = line[:number_match.start()].strip()
                    values_part = number_match.group()
                else:
                    account_name = line
                    values_part = ""

                if not account_name:
                    continue

                # Determine hierarchy
                if account_name in ['ASSETS', 'Assets']:
                    current_section = 'assets'
                    continue
                elif account_name in ['LIABILITIES AND EQUITY', 'Liabilities and Equity']:
                    current_section = 'liabilities_equity'
                    continue
                elif account_name in ['Current Assets', 'Fixed Assets', 'Other Assets']:
                    current_subsection = account_name
                    continue
                elif account_name in ['Liabilities', 'Current Liabilities', 'Long-term Liabilities', 'Long-Term Liabilities']:
                    current_subsection = account_name.replace('-term', '-term')
                    continue
                elif account_name == 'Equity':
                    current_subsection = 'Equity'
                    continue
                elif account_name.startswith('Total ') or account_name == 'TOTAL':
                    continue
                elif account_name in ['Bank Accounts', 'Accounts Receivable', 'Other Current Assets',
                                    'Accounts Payable', 'Credit Cards', 'Other Current Liabilities']:
                    current_group = account_name
                    continue

                # Parse values for each month
                if values_part and current_section:
                    # Extract all numbers from the values part
                    numbers = re.findall(r'[\-\$]?[\d,]+\.?\d*', values_part)

                    # Try to match numbers to months based on position
                    for i, month_info in enumerate(month_columns):
                        if i < len(numbers):
                            value_str = numbers[i].replace('$', '').replace(',', '')
                            try:
                                value = float(value_str)
                            except ValueError:
                                value = 0.0

                            if value != 0.0 or account_name in ['Retained Earnings', 'Net Income']:
                                month = month_info['month']

                                if current_section == 'assets':
                                    section_data = data_by_month[month]['assets']
                                elif current_section == 'liabilities_equity':
                                    if current_subsection == 'Equity':
                                        section_data = data_by_month[month]['equity']
                                    else:
                                        section_data = data_by_month[month]['liabilities']
                                else:
                                    continue

                                if current_subsection not in section_data:
                                    section_data[current_subsection] = {}
                                if current_group and current_group not in section_data[current_subsection]:
                                    section_data[current_subsection][current_group] = {}

                                account_id = self.get_or_create_account_id(account_name)

                                if current_group:
                                    section_data[current_subsection][current_group][account_name] = {
                                        'value': value,
                                        'id': account_id
                                    }
                                else:
                                    section_data[current_subsection][account_name] = {
                                        'value': value,
                                        'id': account_id
                                    }

        return self.build_balance_sheet_json(months, data_by_month)

    def convert_file(self, filepath: Path) -> List[Dict[str, Any]]:
        """Convert a file to balance sheet JSON based on its extension"""
        filepath = Path(filepath)
        ext = filepath.suffix.lower()

        if ext == '.csv':
            return self.parse_csv(filepath)
        elif ext == '.xlsx':
            self.check_xlsx_support()
            return self.parse_xlsx(filepath)
        elif ext == '.pdf':
            self.check_pdf_support()
            return self.parse_pdf(filepath)
        else:
            raise ValueError(f"Unsupported file format: {ext}")


def main():
    parser = argparse.ArgumentParser(description='Convert balance sheet documents to JSON format')
    parser.add_argument('input', help='Input file (CSV, XLSX, or PDF)')
    parser.add_argument('-o', '--output', help='Output JSON file (default: print to stdout)')

    args = parser.parse_args()

    converter = BalanceSheetConverter()

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
