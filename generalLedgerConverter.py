#!/usr/bin/env python3
"""
General Ledger Converter
Converts CSV, XLSX, and PDF general ledger reports to QuickBooks JSON format
Specifically designed for General Ledger reports with transaction-level detail
"""

import json
import csv
import sys
import os
from datetime import datetime, timezone, date
from pathlib import Path
import argparse
import re
import calendar
from typing import Dict, List, Any, Optional, Tuple

from base_converter import BaseConverter, XLSX_SUPPORT, PDF_SUPPORT

try:
    from gl_llm_fallback import GeneralLedgerLLMFallback
    GL_LLM_FALLBACK_AVAILABLE = True
except ImportError:
    GL_LLM_FALLBACK_AVAILABLE = False

if XLSX_SUPPORT:
    import openpyxl

if PDF_SUPPORT:
    import pdfplumber


class GeneralLedgerConverter(BaseConverter):
    """Converts General Ledger documents to QuickBooks-style JSON format"""

    def __init__(self, use_account_lookup: bool = False, api_base_url: str = "http://localhost:8080"):
        super().__init__(use_account_lookup=False, api_base_url=api_base_url)
        # Account lookups disabled for performance - generated IDs work fine
        print("ℹ️  Using generated account IDs (account lookup disabled for performance)", file=sys.stderr)
        self.llm_fallback = GeneralLedgerLLMFallback() if GL_LLM_FALLBACK_AVAILABLE else None
        self.force_llm_pdf = os.getenv('GL_PDF_FORCE_LLM', 'false').lower() == 'true'

    def _can_use_llm_pdf_fallback(self) -> bool:
        return bool(self.llm_fallback and self.llm_fallback.is_available())

    def _normalize_llm_date(self, value: Any) -> str:
        raw = str(value or '').strip()
        if not raw:
            return ''
        for fmt in ('%m/%d/%Y', '%m-%d-%Y', '%Y-%m-%d', '%m/%d/%y'):
            try:
                return datetime.strptime(raw, fmt).strftime('%m/%d/%Y')
            except ValueError:
                continue
        return raw

    def _parse_pdf_with_llm_fallback(self, full_text: str) -> Dict[str, Any]:
        if not self._can_use_llm_pdf_fallback():
            raise ValueError("LLM PDF fallback is not enabled/configured")

        print("🤖 Using LLM fallback for General Ledger PDF parsing", file=sys.stderr)
        rows, llm_meta = self.llm_fallback.extract_rows(full_text)

        if not rows:
            issues = llm_meta.get('issues', []) if isinstance(llm_meta, dict) else []
            raise ValueError(f"LLM extraction returned no rows. Issues: {issues[:5]}")
        
        # DEBUG: Log reconciliation status
        print(f"📊 LLM Extraction Results:", file=sys.stderr)
        print(f"   Rows extracted: {len(rows)}", file=sys.stderr)
        print(f"   Reconciled: {llm_meta.get('reconciled', 'N/A')}", file=sys.stderr)
        print(f"   Confidence: {llm_meta.get('confidence', 'N/A')}", file=sys.stderr)
        if 'totals' in llm_meta:
            totals = llm_meta['totals']
            print(f"   Debit total: ${totals.get('total_debit', 0):.2f}", file=sys.stderr)
            print(f"   Credit total: ${totals.get('total_credit', 0):.2f}", file=sys.stderr)
            print(f"   Delta: ${totals.get('delta', 0):.2f}", file=sys.stderr)
        if llm_meta.get('issues'):
            print(f"   Issues: {llm_meta['issues'][:3]}", file=sys.stderr)

        accounts_data: Dict[str, Any] = {}
        totals_by_account: Dict[str, float] = {}
        parsed_dates: List[date] = []

        for row in rows:
            account_name = (row.get('account') or '').strip() or 'Uncategorized'
            if account_name not in accounts_data:
                accounts_data[account_name] = {
                    'id': self.get_account_id(account_name),
                    'transactions': [],
                    'total': "0.00"
                }
                totals_by_account[account_name] = 0.0

            debit = float(row.get('debit', 0) or 0)
            credit = float(row.get('credit', 0) or 0)
            amount_value = debit - credit
            running_balance = row.get('running_balance')
            norm_date = self._normalize_llm_date(row.get('date'))

            if norm_date:
                for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%m-%d-%Y'):
                    try:
                        parsed_dates.append(datetime.strptime(norm_date, fmt).date())
                        break
                    except ValueError:
                        continue

            tx = {
                'date': norm_date,
                'type': '',
                'num': str(row.get('doc_num', '') or ''),
                'name': '',
                'memo': str(row.get('memo', '') or ''),
                'split_account': '',
                'amount': f"{amount_value:.2f}" if amount_value != 0 else '',
                'balance': f"{float(running_balance):.2f}" if running_balance not in (None, '') else ''
            }

            accounts_data[account_name]['transactions'].append(tx)
            totals_by_account[account_name] += amount_value

        for account_name, total in totals_by_account.items():
            accounts_data[account_name]['total'] = f"{total:.2f}"

        period_info = None
        if parsed_dates:
            start_date = min(parsed_dates)
            end_date = max(parsed_dates)
            period_info = (
                f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}",
                start_date,
                end_date,
            )

        return {
            'period_info': period_info,
            'accounts': accounts_data,
            'llm_meta': llm_meta,
        }

    def get_account_id(self, account_name: str) -> str:
        """Get account ID from lookup service or generate one"""
        return self.get_or_create_account_id(account_name)

    def parse_date_range(self, header_text: str) -> Optional[Tuple[str, date, date]]:
        """
        Parse date range from header text with multiple pattern support.

        Extends base implementation with additional GL-specific patterns:
        - ISO format dates (Pattern 5)
        - Abbreviated months with comma (Pattern 6)
        - Enhanced regex for en-dash/em-dash support
        - Support for single-digit day/month in numeric dates

        Returns: (period_string, start_date, end_date) or None if no match
        """
        # Month name mappings
        months_full = {
            'January': 1, 'February': 2, 'March': 3, 'April': 4,
            'May': 5, 'June': 6, 'July': 7, 'August': 8,
            'September': 9, 'October': 10, 'November': 11, 'December': 12
        }
        months_abbr = {
            'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
            'Jul': 7, 'Aug': 8, 'Sep': 9, 'Sept': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
        }

        # Pattern 1: "April 1-30, 2024" (day range within single month)
        match = re.search(r'(\w+)\s+(\d+)-(\d+),?\s*(\d{4})', header_text, re.IGNORECASE)
        if match:
            month_name = match.group(1)
            start_day = int(match.group(2))
            end_day = int(match.group(3))
            year = int(match.group(4))

            month_num = months_full.get(month_name.capitalize()) or months_abbr.get(month_name.capitalize())
            if month_num:
                try:
                    start_date = date(year, month_num, start_day)
                    end_date = date(year, month_num, end_day)
                    period = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
                    print(f"📅 Parsed header date (Pattern 1): {period}", file=sys.stderr)
                    return period, start_date, end_date
                except ValueError as e:
                    print(f"⚠️  Date parsing error (Pattern 1): {e}", file=sys.stderr)

        # Pattern 2: "January 1 - September 8, 2025" (cross-month range)
        match = re.search(r'(\w+)\s+(\d+)\s*[-–—]\s*(\w+)\s+(\d+),?\s*(\d{4})', header_text, re.IGNORECASE)
        if match:
            start_month_name = match.group(1)
            start_day = int(match.group(2))
            end_month_name = match.group(3)
            end_day = int(match.group(4))
            year = int(match.group(5))

            start_month_num = months_full.get(start_month_name.capitalize()) or months_abbr.get(start_month_name.capitalize())
            end_month_num = months_full.get(end_month_name.capitalize()) or months_abbr.get(end_month_name.capitalize())

            if start_month_num and end_month_num:
                try:
                    start_date = date(year, start_month_num, start_day)
                    end_date = date(year, end_month_num, end_day)
                    period = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
                    print(f"📅 Parsed header date (Pattern 2): {period}", file=sys.stderr)
                    return period, start_date, end_date
                except ValueError as e:
                    print(f"⚠️  Date parsing error (Pattern 2): {e}", file=sys.stderr)

        # Pattern 3: "01/01/2024 - 01/31/2024" or "1/1/2024 to 12/31/2024" (numeric dates)
        match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})\s*[-–—to]+\s*(\d{1,2})/(\d{1,2})/(\d{4})', header_text, re.IGNORECASE)
        if match:
            try:
                start_date = date(int(match.group(3)), int(match.group(1)), int(match.group(2)))
                end_date = date(int(match.group(6)), int(match.group(4)), int(match.group(5)))
                period = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
                print(f"📅 Parsed header date (Pattern 3): {period}", file=sys.stderr)
                return period, start_date, end_date
            except ValueError as e:
                print(f"⚠️  Date parsing error (Pattern 3): {e}", file=sys.stderr)

        # Pattern 4: "January 2024" (full month - infer 1st to last day)
        match = re.search(r'(?:Period:?\s*)?(\w+)\s+(\d{4})', header_text, re.IGNORECASE)
        if match:
            month_name = match.group(1)
            year = int(match.group(2))

            month_num = months_full.get(month_name.capitalize()) or months_abbr.get(month_name.capitalize())
            if month_num:
                try:
                    start_date = date(year, month_num, 1)
                    last_day = calendar.monthrange(year, month_num)[1]
                    end_date = date(year, month_num, last_day)
                    period = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
                    print(f"📅 Parsed header date (Pattern 4 - full month): {period}", file=sys.stderr)
                    return period, start_date, end_date
                except ValueError as e:
                    print(f"⚠️  Date parsing error (Pattern 4): {e}", file=sys.stderr)

        # Pattern 5: "2024-01-01 to 2024-01-31" (ISO format)
        match = re.search(r'(\d{4})-(\d{2})-(\d{2})\s+to\s+(\d{4})-(\d{2})-(\d{2})', header_text, re.IGNORECASE)
        if match:
            try:
                start_date = date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
                end_date = date(int(match.group(4)), int(match.group(5)), int(match.group(6)))
                period = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
                print(f"📅 Parsed header date (Pattern 5): {period}", file=sys.stderr)
                return period, start_date, end_date
            except ValueError as e:
                print(f"⚠️  Date parsing error (Pattern 5): {e}", file=sys.stderr)

        # Pattern 6: "Jan 1, 2024 - Jan 31, 2024" (abbreviated months with comma)
        match = re.search(r'(\w{3,})\s+(\d+),\s*(\d{4})\s*[-–—]\s*(\w{3,})\s+(\d+),\s*(\d{4})', header_text, re.IGNORECASE)
        if match:
            start_month_name = match.group(1)
            start_day = int(match.group(2))
            start_year = int(match.group(3))
            end_month_name = match.group(4)
            end_day = int(match.group(5))
            end_year = int(match.group(6))

            start_month_num = months_full.get(start_month_name.capitalize()) or months_abbr.get(start_month_name.capitalize())
            end_month_num = months_full.get(end_month_name.capitalize()) or months_abbr.get(end_month_name.capitalize())

            if start_month_num and end_month_num:
                try:
                    start_date = date(start_year, start_month_num, start_day)
                    end_date = date(end_year, end_month_num, end_day)
                    period = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
                    print(f"📅 Parsed header date (Pattern 6): {period}", file=sys.stderr)
                    return period, start_date, end_date
                except ValueError as e:
                    print(f"⚠️  Date parsing error (Pattern 6): {e}", file=sys.stderr)

        # No match found - will use transaction dates as fallback
        print(f"⚠️  No header date pattern matched for: {header_text[:100]}", file=sys.stderr)
        return None

    def extract_transaction_date_range(self, accounts_data: Dict[str, Any]) -> Tuple[Optional[date], Optional[date]]:
        """Extract min/max dates from actual transaction data"""
        all_dates = []

        for account_info in accounts_data.values():
            for transaction in account_info.get('transactions', []):
                tx_date = transaction.get('date', '').strip()
                if tx_date:
                    # Try common date formats
                    for date_format in ['%m/%d/%Y', '%Y-%m-%d', '%m-%d-%Y']:
                        try:
                            parsed = datetime.strptime(tx_date, date_format).date()
                            all_dates.append(parsed)
                            break
                        except ValueError:
                            continue

        if all_dates:
            return min(all_dates), max(all_dates)

        # Return None if no dates found
        return None, None

    def validate_date_ranges(self, header_dates: Optional[Tuple], transaction_dates: Tuple[Optional[date], Optional[date]]) -> Dict[str, Any]:
        """Compare header vs transaction dates and determine which to use"""
        warnings = []
        use_transaction_dates = False
        today = date.today()

        tx_start, tx_end = transaction_dates

        # If no transaction dates found, we have a problem
        if tx_start is None or tx_end is None:
            if header_dates:
                # Validate header dates before trusting them
                header_start, header_end = header_dates[1], header_dates[2]
                
                # RED FLAG: end date is today (likely a parsing fallback - reject)
                if header_end == today:
                    print(f"❌ Header end date equals today ({today}), likely a parsing fallback - REJECTING", file=sys.stderr)
                    return {
                        'use_transaction_dates': False,
                        'warnings': [f'Header end date equals today ({today}), likely invalid - both header and transaction parsing failed'],
                        'start_date': None,
                        'end_date': None
                    }
                
                # RED FLAG: start >= end (invalid range)
                if header_start >= header_end:
                    print(f"❌ Invalid date range: start ({header_start}) >= end ({header_end}) - REJECTING", file=sys.stderr)
                    return {
                        'use_transaction_dates': False,
                        'warnings': [f'Invalid date range: start ({header_start}) >= end ({header_end})'],
                        'start_date': None,
                        'end_date': None
                    }
                
                # Header dates look valid
                print(f"✅ Using header dates: {header_start} to {header_end} (no transaction dates found)", file=sys.stderr)
                return {
                    'use_transaction_dates': False,
                    'warnings': ['No transaction dates found, using header dates'],
                    'start_date': header_start,
                    'end_date': header_end
                }
            else:
                # DO NOT fall back to today's date - return None instead
                print("❌ CRITICAL: Could not determine dates from header or transactions", file=sys.stderr)
                return {
                    'use_transaction_dates': False,
                    'warnings': ['Could not determine dates - both header parsing and transaction scan failed'],
                    'start_date': None,
                    'end_date': None
                }

        # If no header dates, use transaction dates
        if header_dates is None:
            return {
                'use_transaction_dates': True,
                'warnings': ['No header dates found, using transaction date range'],
                'start_date': tx_start,
                'end_date': tx_end
            }

        header_start, header_end = header_dates[1], header_dates[2]

        # Check if header dates are suspiciously wide (more than 9 months)
        if (header_end - header_start).days > 274:  # 9 months
            warnings.append(f"Header date range unusually wide: {header_start} to {header_end} ({(header_end - header_start).days} days)")
            use_transaction_dates = True

        # Check if transaction dates differ significantly from header (more than 30 days)
        if abs((tx_start - header_start).days) > 30:
            warnings.append(f"Transaction start date ({tx_start}) differs significantly from header ({header_start})")
            use_transaction_dates = True

        if abs((tx_end - header_end).days) > 30:
            warnings.append(f"Transaction end date ({tx_end}) differs significantly from header ({header_end})")
            use_transaction_dates = True

        # Determine which dates to use
        if use_transaction_dates:
            return {
                'use_transaction_dates': True,
                'warnings': warnings,
                'start_date': tx_start,
                'end_date': tx_end
            }
        else:
            return {
                'use_transaction_dates': False,
                'warnings': warnings if warnings else ['Header dates match transaction dates'],
                'start_date': header_start,
                'end_date': header_end
            }

    def create_transaction_row(self, transaction_data: Dict[str, Any]) -> Dict[str, Any]:
        """Create a transaction row object matching QuickBooks API format"""
        tx_type = transaction_data.get('type', '')
        date_str = transaction_data.get('date', '')

        # Convert date from MM/DD/YYYY to YYYY-MM-DD (ISO format matching QB API)
        if date_str and '/' in date_str:
            parts = date_str.split('/')
            if len(parts) == 3:
                date_str = f"{parts[2]}-{parts[0].zfill(2)}-{parts[1].zfill(2)}"

        # Beginning Balance: QB API puts it in colData[0] (Date field) with balance in colData[7]
        if tx_type == 'Beginning Balance':
            return {
                "id": None,
                "parentId": None,
                "header": None,
                "rows": None,
                "summary": None,
                "colData": [
                    {"attributes": None, "value": "Beginning Balance", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": transaction_data.get('balance', ''), "id": None, "href": None}
                ],
                "type": "DATA",
                "group": None
            }

        return {
            "id": None,
            "parentId": None,
            "header": None,
            "rows": None,
            "summary": None,
            "colData": [
                {"attributes": None, "value": date_str, "id": None, "href": None},
                {"attributes": None, "value": tx_type, "id": None, "href": None},
                {"attributes": None, "value": transaction_data.get('num', ''), "id": None, "href": None},
                {"attributes": None, "value": transaction_data.get('name', ''), "id": None, "href": None},
                {"attributes": None, "value": transaction_data.get('memo', ''), "id": None, "href": None},
                {"attributes": None, "value": transaction_data.get('split_account', ''), "id": None, "href": None},
                {"attributes": None, "value": transaction_data.get('amount', ''), "id": None, "href": None},
                {"attributes": None, "value": transaction_data.get('balance', ''), "id": None, "href": None}
            ],
            "type": "DATA",
            "group": None
        }

    def create_account_section(self, account_name: str, account_id: str,
                             transactions: List[Dict[str, Any]],
                             total_amount: str) -> Dict[str, Any]:
        """Create an account section with its transactions"""
        # Create header for the account
        section = {
            "id": None,
            "parentId": None,
            "header": {
                "colData": [
                    {"attributes": None, "value": account_name, "id": account_id, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None}
                ]
            },
            "rows": {"row": transactions},
            "summary": {
                "colData": [
                    {"attributes": None, "value": f"Total for {account_name}", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None},
                    {"attributes": None, "value": total_amount, "id": None, "href": None},
                    {"attributes": None, "value": "", "id": None, "href": None}
                ]
            },
            "colData": [],
            "type": "SECTION",
            "group": None
        }

        return section

    def parse_csv(self, filepath: Path) -> Dict[str, Any]:
        """Parse CSV file and extract general ledger data"""
        accounts_data = {}
        period_info = None

        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)

            # Find the header with date range (usually in first few rows)
            for i, row in enumerate(rows[:5]):
                if row and any('January' in str(cell) or '-' in str(cell) for cell in row):
                    # Extract date range
                    header_text = ' '.join(row)
                    period_info = self.parse_date_range(header_text)
                    break

            # Find the column headers row
            header_row_idx = -1
            for i, row in enumerate(rows):
                if row and len(row) > 5:
                    # Look for transaction headers
                    row_text = ' '.join(str(cell).lower() for cell in row if cell)
                    if 'transaction date' in row_text or 'transaction type' in row_text:
                        header_row_idx = i
                        break

            if header_row_idx == -1:
                # Try alternative: look for "Distribution account" pattern
                for i, row in enumerate(rows):
                    if row and row[0] and 'Distribution account' in str(row[0]):
                        header_row_idx = i
                        break

            if header_row_idx == -1:
                raise ValueError("Could not find transaction header row")

            # Parse data rows
            current_account = None
            current_account_id = None
            current_transactions = []
            current_total = 0.0

            for row_idx in range(header_row_idx + 1, len(rows)):
                row = rows[row_idx]

                if not row or all(not cell for cell in row):
                    continue

                # Check if this is an account header
                first_cell = str(row[0]).strip() if row else ''

                # Skip grand total rows
                if 'TOTAL' in first_cell.upper() and current_account is None:
                    continue

                # Check if this is a new account section
                if first_cell and len(row) > 1 and not any(row[1:]):
                    # This looks like an account header (only first cell has content)
                    # Save previous account data if exists
                    if current_account and current_transactions:
                        accounts_data[current_account] = {
                            'id': current_account_id,
                            'transactions': current_transactions,
                            'total': f"{current_total:.2f}"
                        }

                    # Start new account
                    current_account = first_cell
                    current_account_id = self.get_account_id(current_account)
                    current_transactions = []
                    current_total = 0.0
                    continue

                # Check if this is a total row for current account
                if current_account and first_cell.startswith(f"Total for {current_account}"):
                    # Extract total from the amount column (usually column 6)
                    if len(row) > 6:
                        total_str = str(row[6]).strip().replace(',', '').replace('$', '')
                        if total_str:
                            try:
                                current_total = float(total_str)
                            except ValueError:
                                pass
                    continue

                # This should be a transaction row
                if current_account and len(row) >= 8:
                    # Parse transaction data
                    # Expected columns: Date, Type, Num, Name, Memo, Split, Amount, Balance
                    transaction = {
                        'date': str(row[1]).strip() if len(row) > 1 else '',
                        'type': str(row[2]).strip() if len(row) > 2 else '',
                        'num': str(row[3]).strip() if len(row) > 3 else '',
                        'name': str(row[4]).strip() if len(row) > 4 else '',
                        'memo': str(row[5]).strip() if len(row) > 5 else '',
                        'split_account': str(row[6]).strip() if len(row) > 6 else '',
                        'amount': str(row[7]).strip() if len(row) > 7 else '',
                        'balance': str(row[8]).strip() if len(row) > 8 else ''
                    }

                    # Only add if it's a valid transaction (has at least a date)
                    if transaction['date']:
                        current_transactions.append(transaction)

            # Save last account
            if current_account and current_transactions:
                accounts_data[current_account] = {
                    'id': current_account_id,
                    'transactions': current_transactions,
                    'total': f"{current_total:.2f}"
                }

        return {
            'period_info': period_info,
            'accounts': accounts_data
        }

    def parse_xlsx(self, filepath: Path) -> Dict[str, Any]:
        """Parse XLSX file and convert to general ledger JSON"""
        self.check_xlsx_support()

        workbook = openpyxl.load_workbook(filepath, data_only=True)
        sheet = workbook.active

        # Convert to list of lists for easier processing
        rows = []
        for row in sheet.iter_rows(values_only=True):
            rows.append(list(row))

        # Find date range in header
        period_info = None
        for row in rows[:5]:
            if row:
                header_text = ' '.join(str(cell) for cell in row if cell)
                if 'January' in header_text or '-' in header_text:
                    period_info = self.parse_date_range(header_text)
                    break

        # Find the column headers row
        header_row_idx = -1
        for i, row in enumerate(rows):
            if row and len(row) > 5:
                row_text = ' '.join(str(cell).lower() for cell in row if cell)
                if 'transaction date' in row_text or 'transaction type' in row_text or 'distribution account' in row_text:
                    header_row_idx = i
                    break

        if header_row_idx == -1:
            raise ValueError("Could not find transaction header row")

        # Process data similar to CSV
        accounts_data = {}
        current_account = None
        current_account_id = None
        current_transactions = []
        current_total = 0.0

        for row_idx in range(header_row_idx + 1, len(rows)):
            row = rows[row_idx]

            if not row or all(cell is None for cell in row):
                continue

            # Convert None values to empty strings
            row = [str(cell) if cell is not None else '' for cell in row]

            first_cell = row[0].strip() if row else ''

            # Skip grand total rows
            if 'TOTAL' in first_cell.upper() and current_account is None:
                continue

            # Check if this is a new account section
            if first_cell and len(row) > 1 and all(not cell.strip() for cell in row[1:]):
                # Save previous account
                if current_account and current_transactions:
                    accounts_data[current_account] = {
                        'id': current_account_id,
                        'transactions': current_transactions,
                        'total': f"{current_total:.2f}"
                    }

                # Start new account
                current_account = first_cell
                current_account_id = self.get_account_id(current_account)
                current_transactions = []
                current_total = 0.0
                continue

            # Check if this is a total row
            if current_account and first_cell.startswith(f"Total for {current_account}"):
                if len(row) > 6:
                    total_str = row[6].strip().replace(',', '').replace('$', '')
                    if total_str:
                        try:
                            current_total = float(total_str)
                        except ValueError:
                            pass
                continue

            # Transaction row
            if current_account and len(row) >= 8:
                transaction = {
                    'date': row[1].strip() if len(row) > 1 else '',
                    'type': row[2].strip() if len(row) > 2 else '',
                    'num': row[3].strip() if len(row) > 3 else '',
                    'name': row[4].strip() if len(row) > 4 else '',
                    'memo': row[5].strip() if len(row) > 5 else '',
                    'split_account': row[6].strip() if len(row) > 6 else '',
                    'amount': row[7].strip() if len(row) > 7 else '',
                    'balance': row[8].strip() if len(row) > 8 else ''
                }

                if transaction['date']:
                    current_transactions.append(transaction)

        # Save last account
        if current_account and current_transactions:
            accounts_data[current_account] = {
                'id': current_account_id,
                'transactions': current_transactions,
                'total': f"{current_total:.2f}"
            }

        return {
            'period_info': period_info,
            'accounts': accounts_data
        }

    def _preprocess_pdf_lines(self, lines: List[str]) -> List[str]:
        """
        Preprocess PDF lines to handle wrapped lines and formatting issues.
        Merges continuation lines like '(Check)' back to the previous transaction.
        """
        processed = []
        i = 0
        
        # Page header/footer patterns to skip
        skip_patterns = [
            r'^General Ledger$',
            r'^Sandbox Company',
            r'^\w+\s+\d+-\d+,\s*\d{4}$',  # Date ranges like "January 1-31, 2023"
            r'^\w+\s+\d+\s*-\s*\w+\s+\d+,\s*\d{4}$',  # Date ranges like "January 1 - March 31, 2025"
            r'^Accrual Basis.*GMTZ.*\d+/\d+$',
            r'^DATE TYPE$',  # The wrapped second line of the header
        ]
        # Track if we've seen the DISTRIBUTION ACCOUNT header (keep first, skip repeats)
        seen_dist_header = False
        
        while i < len(lines):
            line = lines[i].strip()
            
            # Skip empty lines
            if not line:
                i += 1
                continue
            
            # Skip page headers/footers (but keep the column header line)
            skip_line = False
            for pattern in skip_patterns:
                if re.match(pattern, line):
                    skip_line = True
                    break
            
            if skip_line:
                i += 1
                continue

            # Skip repeated DISTRIBUTION ACCOUNT column headers (keep the first one)
            if re.match(r'^DISTRIBUTION ACCOUNT\s+TRANSACTION', line):
                if seen_dist_header:
                    i += 1
                    continue
                seen_dist_header = True

            # Check if this line is a wrapped continuation
            # Continuations are typically: vendor names, "(Type)" wrappers, or isolated words
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                
                # Skip if next line is empty or a skip pattern
                if not next_line:
                    processed.append(line)
                    i += 1
                    continue
                
                # Check for skip patterns in next line
                skip_next = False
                for pattern in skip_patterns:
                    if re.match(pattern, next_line):
                        skip_next = True
                        break
                
                if skip_next:
                    processed.append(line)
                    i += 1
                    continue
                
                # If next line has no date and no amount, it might be a continuation
                has_next_date = bool(re.search(r'\d{1,2}/\d{1,2}/\d{4}', next_line))
                has_next_amount = bool(re.search(r'-?[\d,]+\.\d{2}', next_line))
                
                # Merge if next line is:
                # 1. Type wrapper like "(Check)" or "(Type)"
                # 2. Short text (< 40 chars) with no date/amount (likely wrapped name/vendor)
                # 3. But NOT if next line is an account header or "Beginning Balance"
                
                # Check if next line looks like a real account header or key marker
                # Account keywords that indicate the next line is an account header, not a continuation
                # Only match if the keyword starts the line (to avoid false positives like "Income"
                # being a continuation of "Unapplied Cash Payment")
                account_start_keywords = ['Checking', 'Savings', 'Accounts Receivable', 'Accounts Payable',
                                         'Undeposited', 'Inventory', 'Prepaid', 'Equipment Rental',
                                         'Retained', 'Opening Balance', 'Cost of', 'Fixed Asset']

                is_likely_account = any(next_line.startswith(keyword) for keyword in account_start_keywords)
                is_beginning_balance = 'Beginning Balance' in next_line
                is_total_line = next_line.startswith('Total for ')
                
                # Don't merge if next line is a structural element
                if is_likely_account or is_beginning_balance or is_total_line:
                    processed.append(line)
                    i += 1
                    continue

                # Don't merge if current line is a "Total for" line - the next line is likely an account header
                if line.startswith('Total for '):
                    processed.append(line)
                    i += 1
                    continue

                # Don't merge if next_line looks like an account header:
                # check if the line AFTER next_line starts with next_line (transaction pattern)
                # or is "Beginning Balance" or is another short standalone line
                if not has_next_date and not has_next_amount and len(next_line) < 40:
                    is_account_header = False
                    if i + 2 < len(lines):
                        line_after = lines[i + 2].strip()
                        # Next line is an account header if the line after starts with it + date
                        if line_after.startswith(next_line + ' ') and re.search(r'\d{1,2}/\d{1,2}/\d{4}', line_after):
                            is_account_header = True
                        # Or if line after is "Beginning Balance"
                        elif line_after.startswith('Beginning Balance'):
                            is_account_header = True
                        # Or if line after is also a short standalone (could be sub-account header after parent)
                        elif (not re.search(r'\d{1,2}/\d{1,2}/\d{4}', line_after) and
                              not re.search(r'-?[\d,]+\.\d{2}', line_after) and
                              not line_after.startswith('Total') and
                              line_after and len(line_after) < 40):
                            # Check one more line to see if it's a transaction for the sub-account
                            if i + 3 < len(lines):
                                line_after2 = lines[i + 3].strip()
                                if line_after2.startswith(line_after + ' ') and re.search(r'\d{1,2}/\d{1,2}/\d{4}', line_after2):
                                    is_account_header = True
                    if is_account_header:
                        processed.append(line)
                        i += 1
                        continue

                # Merge if next line is type wrapper or short continuation
                if (next_line.startswith('(') and next_line.endswith(')')) or \
                   (not has_next_date and not has_next_amount and len(next_line) < 40):
                    # Merge the lines
                    line = line + ' ' + next_line
                    i += 2
                    processed.append(line)
                    continue
            
            processed.append(line)
            i += 1
        
        return processed

    def _parse_gl_transaction_line(self, line: str, account_name: str) -> Optional[Dict[str, Any]]:
        """
        Parse a General Ledger transaction line that starts with account name.
        
        Expected format:
        "Checking 01/01/2025 Bill Payment (Check) McClure-Thiel Accounts Payable (A/P) -2,295.63 391,198.80"
        
        Returns dict with transaction fields or None if not a valid transaction.
        """
        # Must start with account name (or a truncated version for long wrapped names)
        # PDF wraps long account names so "Fountains and Garden Lighting" becomes
        # "Fountains and Garden 01/10/2023 ..." with "Lighting" at end of line
        if line.startswith(account_name):
            line_without_account = line[len(account_name):].strip()
        else:
            # Try matching truncated account name: find where date starts
            date_in_line = re.search(r'\d{1,2}/\d{1,2}/\d{4}', line)
            if date_in_line:
                prefix = line[:date_in_line.start()].strip()
                # Check if account name starts with this prefix (truncated match)
                if account_name.startswith(prefix) and len(prefix) >= len(account_name) * 0.5:
                    line_without_account = line[date_in_line.start():].strip()
                else:
                    return None
            else:
                return None

        # Must have a date
        date_match = re.match(r'^(\d{1,2}/\d{1,2}/\d{4})\s+(.+)$', line_without_account)
        if not date_match:
            return None
        
        tx_date = date_match.group(1)
        remainder = date_match.group(2)
        
        # Handle "(Check)" or "(Type)" that may appear after numbers due to line wrapping
        # Pattern: "... -259.78 150,841.20 (Check) Vendor Name Continuation"
        check_suffix = ''
        trailing_extra = ''
        check_match = re.search(r'(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s+\((\w+)\)\s*(.*?)\s*$', remainder)
        if check_match:
            # Extract the wrapped type modifier and any trailing vendor name
            check_suffix = f" ({check_match.group(3)})"  # e.g., " (Check)"
            trailing_extra = check_match.group(4).strip()  # e.g., "Stoltenberg"
            # Trim remainder to end at the balance number
            remainder = remainder[:check_match.start()] + f" {check_match.group(1)} {check_match.group(2)}"

        # Extract amount and balance (last two numbers on the line)
        numbers = re.findall(r'-?[\d,]+\.\d{2}', remainder)

        if len(numbers) < 2:
            return None

        amount = numbers[-2].replace(',', '')
        balance = numbers[-1].replace(',', '')

        # Remove trailing amount+balance to get middle content
        trailing_match = re.search(r'\s+-?[\d,]+\.\d{2}\s+-?[\d,]+\.\d{2}\s*$', remainder)
        if trailing_match:
            middle_content = remainder[:trailing_match.start()].strip()
        else:
            amount_pos = remainder.find(numbers[-2])
            middle_content = remainder[:amount_pos].strip()

        # Append any trailing vendor name continuation
        if trailing_extra:
            middle_content = middle_content + ' ' + trailing_extra
        
        # Parse middle content for: Type, Num, Name, Memo, Split
        # Pattern: "Transaction Type (Check) Name Memo/Description Split Account"
        
        transaction = {
            'date': tx_date,
            'type': '',
            'num': '',
            'name': '',
            'memo': '',
            'split_account': '',
            'amount': amount,
            'balance': balance
        }
        
        # Extract transaction type (first word/phrase before names)
        # Order matters: longer/more specific patterns first
        type_patterns = [
            'Bill Payment \\(Check\\)',
            'Bill Payment',
            'Credit Card Credit',
            'Sales Tax Payment',
            'Sales Receipt',
            'Vendor Credit',
            'Cash Expense',
            'Journal Entry',
            'Deposit',
            'Transfer',
            'Expense',
            'Invoice',
            'Payment',
            'Check',
            'Refund',
            'Bill',
        ]
        
        for pattern in type_patterns:
            match = re.search(pattern, middle_content, re.IGNORECASE)
            if match:
                tx_type_str = match.group(0)
                # Append check suffix if present (e.g., "Bill Payment" + " (Check)")
                if check_suffix and tx_type_str.lower() in ('bill payment', 'payment'):
                    tx_type_str = tx_type_str + check_suffix
                transaction['type'] = tx_type_str
                # Remove type from middle content
                middle_content = middle_content[:match.start()] + middle_content[match.end():]
                middle_content = middle_content.strip()
                break
        
        # What remains should be: Name, Memo, Split Account
        # Try to identify split account (often contains ":" or recognizable account names)
        # Split account is usually at the end before the numbers
        
        # Common split account patterns
        split_patterns = [
            r'(Accounts Payable \(A/P\))',
            r'(Accounts Receivable \(A/R\))',
            r'(Opening Balance Equity)',
            r'([A-Za-z\s]+:[A-Za-z\s:]+)',  # Account with colons like "Utilities:Telephone"
            r'(Savings|Checking)',
            r'(-Split-)',
        ]
        
        for pattern in split_patterns:
            match = re.search(pattern, middle_content)
            if match:
                transaction['split_account'] = match.group(1).strip()
                # Remove from middle content
                middle_content = middle_content[:match.start()] + middle_content[match.end():]
                middle_content = middle_content.strip()
                break
        
        # What's left is Name and Memo - hard to distinguish without better structure
        # For now, treat it all as name
        if middle_content:
            # Clean up multiple spaces from line merging
            transaction['name'] = re.sub(r'\s{2,}', ' ', middle_content).strip()
        
        return transaction

    def _build_parent_child_map(self, lines: List[str]) -> Dict[str, List[str]]:
        """Pre-scan lines to build parent->children map using 'Total for X with sub-accounts' markers.

        Handles multi-level nesting (e.g., Landscaping Services -> Labor -> Installation).
        Returns dict mapping parent account name -> list of DIRECT child account names.
        """
        # First, find all parent accounts from "Total for X with sub-accounts" lines
        parent_accounts = set()
        for line in lines:
            stripped = line.strip()
            match = re.match(r'^Total for (.+?) with sub-accounts', stripped)
            if match:
                parent_accounts.add(match.group(1))

        if not parent_accounts:
            return {}

        # Use a stack to handle nested parents.
        # When we see a parent header, push it. When we see "Total for X with sub-accounts", pop X.
        # Direct children of a parent are standalone lines that appear at the current nesting level.
        parent_children: Dict[str, List[str]] = {p: [] for p in parent_accounts}
        parent_stack: List[str] = []  # stack of open parent contexts

        for line in lines:
            stripped = line.strip()

            # Check if this line ends a parent section
            match = re.match(r'^Total for (.+?) with sub-accounts', stripped)
            if match:
                closing_parent = match.group(1)
                # Pop the stack back to this parent
                while parent_stack and parent_stack[-1] != closing_parent:
                    parent_stack.pop()
                if parent_stack:
                    parent_stack.pop()
                continue

            # Skip total lines, transaction lines, beginning balance, etc.
            if stripped.startswith('Total for '):
                continue

            has_date = bool(re.search(r'\d{1,2}/\d{1,2}/\d{4}', stripped))
            has_numbers = bool(re.search(r'[\d,]+\.\d{2}', stripped))

            # Only consider standalone text lines (potential account headers)
            if has_date or has_numbers or not stripped or stripped.startswith('Beginning Balance') or ':' in stripped:
                continue

            # Check if this is a parent account header
            if stripped in parent_accounts:
                # This is both a child of the current parent (if any) and a new parent context
                if parent_stack:
                    immediate_parent = parent_stack[-1]
                    if stripped not in parent_children[immediate_parent]:
                        parent_children[immediate_parent].append(stripped)
                parent_stack.append(stripped)
                continue

            # If we're inside a parent context, this is a direct child (sub-account)
            if parent_stack:
                immediate_parent = parent_stack[-1]
                if stripped not in parent_children[immediate_parent]:
                    parent_children[immediate_parent].append(stripped)

        # Log what we found
        for parent, children in parent_children.items():
            if children:
                print(f"📂 Parent account '{parent}' has sub-accounts: {children}", file=sys.stderr)

        return parent_children

    def parse_pdf(self, filepath: Path) -> Dict[str, Any]:
        """Parse PDF file and convert to general ledger JSON"""
        self.check_pdf_support()

        accounts_data = {}
        period_info = None

        with pdfplumber.open(filepath) as pdf:
            # Extract text from all pages
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"

            raw_lines = all_text.split('\n')

            # Force LLM mode for GL PDFs when explicitly enabled
            if self.force_llm_pdf and self._can_use_llm_pdf_fallback():
                print("🤖 GL_PDF_FORCE_LLM=true: bypassing deterministic PDF parser", file=sys.stderr)
                return self._parse_pdf_with_llm_fallback(all_text)

            # Find date range in header (first few lines)
            for line in raw_lines[:10]:
                if 'January' in line or 'February' in line or 'March' in line or '-' in line:
                    period_info = self.parse_date_range(line)
                    if period_info:
                        break

            # Preprocess lines (merge wrapped lines)
            lines = self._preprocess_pdf_lines(raw_lines)

            print(f"📄 Preprocessed {len(raw_lines)} raw lines into {len(lines)} lines", file=sys.stderr)

            # Pre-scan for parent-child account hierarchy
            parent_child_map = self._build_parent_child_map(lines)
            # Build reverse lookup: child_name -> parent_name
            child_to_parent = {}
            for parent, children in parent_child_map.items():
                for child in children:
                    child_to_parent[child] = parent

            # Find header line - look for "DISTRIBUTION ACCOUNT" pattern
            header_idx = -1
            for i, line in enumerate(lines):
                line_upper = line.upper()
                if ('DISTRIBUTION ACCOUNT' in line_upper or
                    'TRANSACTION DATE' in line_upper or
                    'TRANSACTION TYPE' in line_upper):
                    header_idx = i
                    print(f"📋 Found header at line {i}: {line[:80]}", file=sys.stderr)
                    break

            if header_idx == -1:
                print("⚠️  Could not find standard GL header, attempting LLM fallback", file=sys.stderr)
                if self._can_use_llm_pdf_fallback():
                    return self._parse_pdf_with_llm_fallback(all_text)
                raise ValueError("Could not find header row in PDF and LLM fallback not available")

            # Parse transactions starting after header
            current_account = None
            current_account_id = None
            current_transactions = []
            current_total = 0.0
            # Track parent account context using a stack for multi-level nesting
            # Each entry: parent account name
            parent_stack: List[str] = []

            def _get_current_parent() -> Optional[str]:
                """Get the immediate parent from the stack, or None."""
                return parent_stack[-1] if parent_stack else None

            def _save_current_account():
                """Save current account data if it has transactions."""
                nonlocal current_account, current_account_id, current_transactions, current_total
                if current_account and current_transactions:
                    parent = _get_current_parent()
                    # Don't set parent to self
                    if parent == current_account:
                        parent = parent_stack[-2] if len(parent_stack) >= 2 else None
                    accounts_data[current_account] = {
                        'id': current_account_id,
                        'transactions': current_transactions,
                        'total': f"{current_total:.2f}",
                        'parent': parent
                    }
                current_account = None
                current_transactions = []
                current_total = 0.0

            line_idx = header_idx + 1
            while line_idx < len(lines):
                line = lines[line_idx].strip()

                if not line or 'Page' in line:
                    line_idx += 1
                    continue

                # Check for "Beginning Balance" line
                if line.startswith('Beginning Balance'):
                    balance_match = re.search(r'([\d,]+\.\d{2})', line)
                    if balance_match and current_account:
                        beginning_balance = balance_match.group(1).replace(',', '')
                        current_transactions.append({
                            'date': '',
                            'type': 'Beginning Balance',
                            'num': '',
                            'name': '',
                            'memo': '',
                            'split_account': '',
                            'amount': '',
                            'balance': beginning_balance
                        })
                    line_idx += 1
                    continue

                # Check for "Total for [Account]" line
                if line.startswith('Total for '):
                    # Check if this is a "with sub-accounts" parent total
                    if 'with sub-accounts' in line:
                        # Save current sub-account if pending
                        _save_current_account()
                        # Extract parent name and total
                        parent_total_match = re.match(r'Total for (.+?) with sub-accounts\s*(.*)', line)
                        if parent_total_match:
                            parent_name = parent_total_match.group(1)
                            total_str = parent_total_match.group(2)
                            parent_total = 0.0
                            total_match = re.search(r'[-\$\d,]+\.\d{2}', total_str)
                            if total_match:
                                parent_total = float(total_match.group(0).replace(',', '').replace('$', ''))
                            # Determine this parent's own parent
                            # Pop the stack back to this parent
                            while parent_stack and parent_stack[-1] != parent_name:
                                parent_stack.pop()
                            if parent_stack:
                                parent_stack.pop()  # Remove this parent from stack
                            parent_of_parent = _get_current_parent()
                            # Store/update parent account
                            if parent_name not in accounts_data:
                                accounts_data[parent_name] = {
                                    'id': self.get_account_id(parent_name),
                                    'transactions': [],
                                    'total': f"{parent_total:.2f}",
                                    'parent': parent_of_parent,
                                    'is_parent': True
                                }
                            else:
                                accounts_data[parent_name]['total'] = f"{parent_total:.2f}"
                                accounts_data[parent_name]['is_parent'] = True
                                # Set correct parent (could be a grandparent)
                                accounts_data[parent_name]['parent'] = parent_of_parent
                        line_idx += 1
                        continue

                    # Regular "Total for X" line (sub-account or standalone account)
                    if current_account:
                        total_match = re.search(r'([-\$\d,]+\.\d{2})', line)
                        if total_match:
                            current_total = float(total_match.group(1).replace(',', '').replace('$', ''))
                        _save_current_account()
                    line_idx += 1
                    continue

                # Check if this is a standalone account header (no date, no numbers)
                has_date = bool(re.search(r'\d{1,2}/\d{1,2}/\d{4}', line))
                has_numbers = bool(re.search(r'[\d,]+\.\d{2}', line))

                if not has_date and not has_numbers and not line.startswith('Total'):
                    # Skip if line contains colon
                    if ':' in line:
                        line_idx += 1
                        continue

                    # Check if this is a parent account header
                    if line in parent_child_map:
                        _save_current_account()
                        parent_stack.append(line)
                        current_account = line
                        current_account_id = self.get_account_id(line)
                        current_transactions = []
                        current_total = 0.0
                        print(f"🏦 Parent account section: {line} (depth={len(parent_stack)})", file=sys.stderr)
                        line_idx += 1
                        continue

                    # Check if this is a known sub-account AND we're inside its parent's context
                    if line in child_to_parent and _get_current_parent() is not None:
                        expected_parent = child_to_parent[line]
                        # Only treat as sub-account if we're inside the right parent
                        if expected_parent in parent_stack:
                            _save_current_account()
                            current_account = line
                            current_account_id = self.get_account_id(line)
                            current_transactions = []
                            current_total = 0.0
                            print(f"  📎 Sub-account: {line} (under {_get_current_parent()})", file=sys.stderr)
                            line_idx += 1
                            continue

                    # Validate if this is a real account header by looking ahead
                    is_real_account = False
                    if line_idx + 1 < len(lines):
                        next_line = lines[line_idx + 1].strip()

                        if next_line.startswith('Beginning Balance'):
                            is_real_account = True
                        else:
                            # Check if next line starts with account name (or truncated version for wrapped names)
                            name_matches = next_line.startswith(line + ' ')
                            if not name_matches and re.search(r'\d{1,2}/\d{1,2}/\d{4}', next_line):
                                # Try truncated match: find where date starts in next line
                                date_pos = re.search(r'\d{1,2}/\d{1,2}/\d{4}', next_line)
                                if date_pos:
                                    prefix = next_line[:date_pos.start()].strip()
                                    if line.startswith(prefix) and len(prefix) >= len(line) * 0.5:
                                        name_matches = True

                            if name_matches and re.search(r'\d{1,2}/\d{1,2}/\d{4}', next_line):
                                # Look for "Total for [Account]" footer
                                total_footer_found = False
                                search_limit = min(line_idx + 500, len(lines))
                                for check_idx in range(line_idx + 2, search_limit):
                                    if lines[check_idx].strip().startswith(f'Total for {line}'):
                                        total_footer_found = True
                                        break
                                    check_line = lines[check_idx].strip()
                                    if (not re.search(r'\d{1,2}/\d{1,2}/\d{4}', check_line) and
                                        not re.search(r'[\d,]+\.\d{2}', check_line) and
                                        check_line and
                                        check_line != 'Beginning Balance' and
                                        not check_line.startswith('Total')):
                                        break

                                if total_footer_found:
                                    is_real_account = True

                    if is_real_account:
                        _save_current_account()
                        current_account = line
                        current_account_id = self.get_account_id(current_account)
                        current_transactions = []
                        current_total = 0.0
                        print(f"🏦 New account section: {current_account}", file=sys.stderr)

                    line_idx += 1
                    continue

                # Try to parse as transaction line
                if current_account and has_date:
                    transaction = self._parse_gl_transaction_line(line, current_account)
                    if transaction:
                        current_transactions.append(transaction)
                        if len(current_transactions) <= 3:
                            print(f"  ✅ Transaction: {transaction['date']} {transaction['type']} {transaction['amount']}", file=sys.stderr)

                line_idx += 1

            # Save last account
            if current_account and current_transactions:
                _save_current_account()

        if not accounts_data and self._can_use_llm_pdf_fallback():
            print("⚠️  No data extracted, attempting LLM fallback", file=sys.stderr)
            return self._parse_pdf_with_llm_fallback(all_text)

        print(f"📊 Extracted {len(accounts_data)} accounts with transactions", file=sys.stderr)

        return {
            'period_info': period_info,
            'accounts': accounts_data
        }

    def build_json_structure(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Build the complete general ledger JSON structure"""
        period_info = data.get('period_info')
        accounts_data = data.get('accounts', {})
        llm_meta = data.get('llm_meta')

        # Extract transaction date range from actual data
        transaction_dates = self.extract_transaction_date_range(accounts_data)

        # Validate and determine which dates to use
        date_validation = self.validate_date_ranges(period_info, transaction_dates)

        # Log warnings
        for warning in date_validation['warnings']:
            print(f"⚠️  GL Date Validation: {warning}", file=sys.stderr)

        # Use the validated dates
        start_date = date_validation['start_date']
        end_date = date_validation['end_date']

        # If dates are None, we have a critical issue - fail fast
        if start_date is None or end_date is None:
            print("❌ CRITICAL: Could not determine valid date range for General Ledger", file=sys.stderr)
            for warning in date_validation['warnings']:
                print(f"  - {warning}", file=sys.stderr)
            
            # Raise an error - better to fail than return wrong data
            raise ValueError(
                "Could not determine valid date range for General Ledger. "
                "Both header parsing and transaction date extraction failed. "
                "Please check that the file has a date range in the header (e.g., 'January 1, 2025-December 31, 2025') "
                "and/or contains transaction records with valid dates."
            )

        if date_validation['use_transaction_dates']:
            print(f"✅ Using transaction dates: {start_date} to {end_date}", file=sys.stderr)

        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000+00:00')

        # Build the header
        result = {
            "header": {
                "time": timestamp,
                "reportName": "GeneralLedger",
                "dateMacro": None,
                "reportBasis": "ACCRUAL",
                "startPeriod": start_date.strftime('%Y-%m-%d'),
                "endPeriod": end_date.strftime('%Y-%m-%d'),
                "summarizeColumnsBy": None,
                "currency": "USD",
                "customer": None,
                "vendor": None,
                "employee": None,
                "item": None,
                "clazz": None,
                "department": None,
                "option": [
                    {"name": "NoReportData", "value": "false" if accounts_data else "true"}
                ]
            },
            "columns": {
                "column": [
                    {"colTitle": "Date", "colType": "Date", "metaData": [{"name": "ColKey", "value": "tx_date"}], "columns": None},
                    {"colTitle": "Transaction Type", "colType": "String", "metaData": [{"name": "ColKey", "value": "txn_type"}], "columns": None},
                    {"colTitle": "Num", "colType": "String", "metaData": [{"name": "ColKey", "value": "doc_num"}], "columns": None},
                    {"colTitle": "Name", "colType": "String", "metaData": [{"name": "ColKey", "value": "name"}], "columns": None},
                    {"colTitle": "Memo/Description", "colType": "String", "metaData": [{"name": "ColKey", "value": "memo"}], "columns": None},
                    {"colTitle": "Split", "colType": "String", "metaData": [{"name": "ColKey", "value": "split_acc"}], "columns": None},
                    {"colTitle": "Amount", "colType": "Money", "metaData": [{"name": "ColKey", "value": "subt_nat_amount"}], "columns": None},
                    {"colTitle": "Balance", "colType": "Money", "metaData": [{"name": "ColKey", "value": "rbal_nat_amount"}], "columns": None}
                ]
            },
            "rows": {"row": []}
        }

        # Build rows - nest sub-accounts under parents (supports multi-level nesting)
        # First, build parent -> [children] map from the parsed data
        parent_accounts = {}  # parent_name -> [child_names]
        child_accounts = set()
        for account_name, account_info in accounts_data.items():
            parent = account_info.get('parent')
            if parent and parent != account_name:
                child_accounts.add(account_name)
                if parent not in parent_accounts:
                    parent_accounts[parent] = []
                parent_accounts[parent].append(account_name)

        def _build_account_section(account_name: str) -> Dict[str, Any]:
            """Recursively build an account section, nesting children if any."""
            account_info = accounts_data[account_name]

            if account_name in parent_accounts:
                # Parent account - build nested structure
                children = parent_accounts[account_name]
                inner_rows = []

                # Parent's own direct transactions (in a headerless SECTION)
                if account_info['transactions']:
                    direct_txns = [self.create_transaction_row(t) for t in account_info['transactions']]
                    inner_rows.append({
                        "id": None, "parentId": None, "header": None,
                        "rows": {"row": direct_txns}, "summary": None,
                        "colData": [], "type": "SECTION", "group": None
                    })

                # Recursively build child sections
                for child_name in children:
                    if child_name in accounts_data:
                        inner_rows.append(_build_account_section(child_name))

                is_parent = account_info.get('is_parent', False)
                summary_text = f"Total for {account_name} with sub-accounts" if is_parent else f"Total for {account_name}"

                return {
                    "id": None, "parentId": None,
                    "header": {
                        "colData": [
                            {"attributes": None, "value": account_name, "id": account_info['id'], "href": None},
                            *[{"attributes": None, "value": "", "id": None, "href": None} for _ in range(7)]
                        ]
                    },
                    "rows": {"row": inner_rows},
                    "summary": {
                        "colData": [
                            {"attributes": None, "value": summary_text, "id": None, "href": None},
                            *[{"attributes": None, "value": "", "id": None, "href": None} for _ in range(5)],
                            {"attributes": None, "value": account_info['total'], "id": None, "href": None},
                            {"attributes": None, "value": "", "id": None, "href": None}
                        ]
                    },
                    "colData": [], "type": "SECTION", "group": None
                }
            else:
                # Leaf account - flat section
                txn_rows = [self.create_transaction_row(t) for t in account_info['transactions']]
                return self.create_account_section(
                    account_name, account_info['id'], txn_rows, account_info['total']
                )

        rows = []
        for account_name in accounts_data:
            if account_name not in child_accounts:
                rows.append(_build_account_section(account_name))

        result["rows"]["row"] = rows

        if llm_meta:
            result["extractionMeta"] = llm_meta

        return result

    def convert_file(self, filepath: Path) -> Dict[str, Any]:
        """Convert a file to general ledger JSON based on its extension"""
        filepath = Path(filepath)
        ext = filepath.suffix.lower()

        if ext == '.csv':
            data = self.parse_csv(filepath)
        elif ext == '.xlsx':
            data = self.parse_xlsx(filepath)
        elif ext == '.pdf':
            data = self.parse_pdf(filepath)
        else:
            raise ValueError(f"Unsupported file format: {ext}")

        return self.build_json_structure(data)

    def convert_to_json(self, filepath: Path, output_path: Optional[Path] = None) -> str:
        """Convert a file to JSON format"""
        try:
            general_ledger = self.convert_file(filepath)

            if output_path:
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(general_ledger, f, indent=2)
                return f"Converted general ledger to {output_path}"
            else:
                return json.dumps(general_ledger, indent=2)
        except Exception as e:
            import traceback
            traceback.print_exc()
            raise


def main():
    parser = argparse.ArgumentParser(description='Convert general ledger documents to JSON format')
    parser.add_argument('input', help='Input file (CSV, XLSX, or PDF)')
    parser.add_argument('-o', '--output', help='Output JSON file (default: print to stdout)')
    parser.add_argument('--no-lookup', action='store_true', help='Disable account lookup service')

    args = parser.parse_args()

    converter = GeneralLedgerConverter(use_account_lookup=not args.no_lookup)

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
