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
from typing import Dict, List, Any, Optional, Tuple

# Import account lookup client
try:
    from account_lookup_client import get_account_lookup_client
    ACCOUNT_LOOKUP_AVAILABLE = True
except ImportError:
    ACCOUNT_LOOKUP_AVAILABLE = False

# Try to import optional dependencies
try:
    import openpyxl
    XLSX_SUPPORT = True
except ImportError:
    XLSX_SUPPORT = False

try:
    import pdfplumber
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False


class GeneralLedgerConverter:
    """Converts General Ledger documents to QuickBooks-style JSON format"""
    
    def __init__(self, use_account_lookup: bool = False, api_base_url: str = "http://localhost:8080"):
        self.account_id_counter = 1
        # Disable account lookup by default - it causes performance issues
        self.use_account_lookup = False
        self.account_lookup_client = None
        
        # Account lookups disabled for performance - generated IDs work fine
        print("ℹ️  Using generated account IDs (account lookup disabled for performance)", file=sys.stderr)
        
    def get_account_id(self, account_name: str) -> str:
        """Get account ID from lookup service or generate one"""
        if self.use_account_lookup and self.account_lookup_client:
            # Try to look up the account ID
            account_id = self.account_lookup_client.lookup_account_id(account_name)
            if account_id:
                return account_id
        
        # Fallback to generating an ID
        return self.generate_account_id()
        
    def generate_account_id(self) -> str:
        """Generate a unique account ID"""
        id_str = str(self.account_id_counter)
        self.account_id_counter += 1
        return id_str
    
    def parse_date_range(self, header_text: str) -> Optional[Tuple[str, date, date]]:
        """
        Parse date range from header text with multiple pattern support
        Returns: (period_string, start_date, end_date) or None if no match
        """
        import calendar
        
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
        
        tx_start, tx_end = transaction_dates
        
        # If no transaction dates found, we have a problem
        if tx_start is None or tx_end is None:
            if header_dates:
                return {
                    'use_transaction_dates': False,
                    'warnings': ['No transaction dates found, using header dates'],
                    'start_date': header_dates[1],
                    'end_date': header_dates[2]
                }
            else:
                # Fallback to current year Jan 1
                return {
                    'use_transaction_dates': False,
                    'warnings': ['No dates found in header or transactions, using fallback'],
                    'start_date': date.today().replace(month=1, day=1),
                    'end_date': date.today()
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
        """Create a transaction row object"""
        return {
            "id": None,
            "parentId": None,
            "header": None,
            "rows": None,
            "summary": None,
            "colData": [
                {"attributes": None, "value": transaction_data.get('date', ''), "id": None, "href": None},
                {"attributes": None, "value": transaction_data.get('type', ''), "id": None, "href": None},
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
        if not XLSX_SUPPORT:
            raise ImportError("openpyxl is required for XLSX support. Install with: pip install openpyxl")
        
        workbook = openpyxl.load_workbook(filepath)
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
    
    def parse_pdf(self, filepath: Path) -> Dict[str, Any]:
        """Parse PDF file and convert to general ledger JSON"""
        if not PDF_SUPPORT:
            raise ImportError("pdfplumber is required for PDF support. Install with: pip install pdfplumber")
        
        accounts_data = {}
        period_info = None
        
        with pdfplumber.open(filepath) as pdf:
            # Extract text from all pages
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
            
            lines = all_text.split('\n')
            
            # Find date range in header
            for line in lines[:10]:
                if 'January' in line or '-' in line:
                    period_info = self.parse_date_range(line)
                    break
            
            # Find header line
            header_idx = -1
            for i, line in enumerate(lines):
                line_lower = line.lower()
                if ('transaction date' in line_lower or 'transaction type' in line_lower or 
                    'distribution account' in line_lower):
                    header_idx = i
                    break
            
            if header_idx == -1:
                raise ValueError("Could not find header row in PDF")
            
            # Parse data
            current_account = None
            current_account_id = None
            current_transactions = []
            current_total = 0.0
            
            for line_idx in range(header_idx + 1, len(lines)):
                line = lines[line_idx].strip()
                
                if not line or 'Page' in line:
                    continue
                
                # Check if this is an account header (no numbers, just account name)
                if not re.search(r'\d{1,2}/\d{1,2}/\d{4}', line) and not re.search(r'[\d,]+\.\d{2}', line):
                    # Skip total rows
                    if 'TOTAL' in line.upper() and current_account is None:
                        continue
                    
                    # Check if this is a total row for current account
                    if current_account and line.startswith(f"Total for {current_account}"):
                        # Try to extract total from the line
                        total_match = re.search(r'[\-\$]?([\d,]+\.?\d*)', line)
                        if total_match:
                            total_str = total_match.group(1).replace(',', '')
                            try:
                                current_total = float(total_str)
                            except ValueError:
                                pass
                        continue
                    
                    # This might be a new account
                    if not line.startswith('Total'):
                        # Save previous account
                        if current_account and current_transactions:
                            accounts_data[current_account] = {
                                'id': current_account_id,
                                'transactions': current_transactions,
                                'total': f"{current_total:.2f}"
                            }
                        
                        # Start new account
                        current_account = line
                        current_account_id = self.get_account_id(current_account)
                        current_transactions = []
                        current_total = 0.0
                    continue
                
                # Try to parse as transaction line
                if current_account:
                    # Look for date pattern at the start
                    date_match = re.match(r'(\d{1,2}/\d{1,2}/\d{4})', line)
                    if date_match:
                        # This is likely a transaction line
                        # Try to extract fields
                        parts = line.split()
                        
                        transaction = {
                            'date': date_match.group(1),
                            'type': '',
                            'num': '',
                            'name': '',
                            'memo': '',
                            'split_account': '',
                            'amount': '',
                            'balance': ''
                        }
                        
                        # Extract amount and balance (usually the last two numeric values)
                        numbers = re.findall(r'[\-\$]?([\d,]+\.?\d*)', line)
                        if len(numbers) >= 2:
                            transaction['amount'] = numbers[-2].replace(',', '')
                            transaction['balance'] = numbers[-1].replace(',', '')
                        
                        # Try to extract other fields
                        # This is challenging with PDF as formatting can vary
                        # Basic approach: after date, look for transaction type keywords
                        remaining = line[date_match.end():].strip()
                        
                        # Common transaction types
                        type_patterns = ['Invoice', 'Payment', 'Bill', 'Check', 'Deposit', 'Transfer']
                        for pattern in type_patterns:
                            if pattern in remaining:
                                transaction['type'] = pattern
                                break
                        
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
    
    def build_json_structure(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Build the complete general ledger JSON structure"""
        period_info = data.get('period_info')
        accounts_data = data.get('accounts', {})
        
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
        
        if date_validation['use_transaction_dates']:
            print(f"✅ Using transaction dates: {start_date} to {end_date}", file=sys.stderr)
        
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000+00:00')
        
        # Build the header
        result = {
            "header": {
                "time": timestamp,
                "reportName": "GeneralLedger",
                "dateMacro": None,
                "reportBasis": None,
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
                    {"colTitle": "Date", "colType": "tx_date", "metaData": None, "columns": None},
                    {"colTitle": "Transaction Type", "colType": "txn_type", "metaData": None, "columns": None},
                    {"colTitle": "Num", "colType": "doc_num", "metaData": None, "columns": None},
                    {"colTitle": "Name", "colType": "name", "metaData": None, "columns": None},
                    {"colTitle": "Memo/Description", "colType": "memo", "metaData": None, "columns": None},
                    {"colTitle": "Split", "colType": "split_acc", "metaData": None, "columns": None},
                    {"colTitle": "Amount", "colType": "amount", "metaData": None, "columns": None},
                    {"colTitle": "Balance", "colType": "balance", "metaData": None, "columns": None}
                ]
            },
            "rows": {"row": []}
        }
        
        # Build rows for each account
        rows = []
        for account_name, account_info in accounts_data.items():
            # Convert transactions to row objects
            transaction_rows = []
            for trans in account_info['transactions']:
                transaction_rows.append(self.create_transaction_row(trans))
            
            # Create account section
            account_section = self.create_account_section(
                account_name,
                account_info['id'],
                transaction_rows,
                account_info['total']
            )
            
            rows.append(account_section)
        
        result["rows"]["row"] = rows
        
        return result
    
    def convert_file(self, filepath: Path) -> Dict[str, Any]:
        """Convert a file to general ledger JSON based on its extension"""
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
