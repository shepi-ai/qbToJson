#!/usr/bin/env python3
"""
Chart of Accounts Document Converter
Converts CSV, XLSX, and PDF account lists to QuickBooks JSON format
Specifically designed for Chart of Accounts reports
"""

import csv
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
import argparse
from typing import Dict, List, Any, Optional

from base_converter import BaseConverter, XLSX_SUPPORT, PDF_SUPPORT

if XLSX_SUPPORT:
    import openpyxl
if PDF_SUPPORT:
    import pdfplumber


class AccountsConverter(BaseConverter):
    """Converts Chart of Accounts documents to QuickBooks-style JSON format"""

    def __init__(self, **kwargs):
        super().__init__(use_account_lookup=False)
        self.account_id_counter = 200  # Starting ID for generated accounts

    def get_classification_from_type(self, type_str: str) -> str:
        """Determine classification based on type string"""
        type_lower = type_str.lower()
        if 'accounts payable' in type_lower or type_lower == 'a/p':
            return 'LIABILITY'
        elif 'accounts receivable' in type_lower or type_lower == 'a/r':
            return 'ASSET'
        elif type_lower == 'bank' or type_lower == 'checking' or type_lower == 'savings':
            return 'ASSET'
        elif 'credit card' in type_lower:
            return 'LIABILITY'
        elif 'cost of goods sold' in type_lower or type_lower == 'cogs':
            return 'EXPENSE'
        elif 'equity' in type_lower:
            return 'EQUITY'
        elif 'expense' in type_lower:
            return 'EXPENSE'
        elif 'income' in type_lower:
            return 'REVENUE'
        elif 'asset' in type_lower:
            return 'ASSET'
        elif 'liabilit' in type_lower:
            return 'LIABILITY'
        else:
            return 'EXPENSE'

    def get_account_type_from_type(self, type_str: str) -> str:
        """Determine account type based on type string"""
        type_lower = type_str.lower()
        if 'accounts payable' in type_lower or type_lower == 'a/p':
            return 'ACCOUNTS_PAYABLE'
        elif 'accounts receivable' in type_lower or type_lower == 'a/r':
            return 'ACCOUNTS_RECEIVABLE'
        elif type_lower == 'bank' or type_lower == 'checking' or type_lower == 'savings':
            return 'BANK'
        elif 'credit card' in type_lower:
            return 'CREDIT_CARD'
        elif 'cost of goods sold' in type_lower or type_lower == 'cogs':
            return 'COST_OF_GOODS_SOLD'
        elif 'equity' in type_lower:
            return 'EQUITY'
        elif 'other expense' in type_lower:
            return 'OTHER_EXPENSE'
        elif 'expense' in type_lower:
            return 'EXPENSE'
        elif 'other income' in type_lower:
            return 'OTHER_INCOME'
        elif 'income' in type_lower:
            return 'INCOME'
        elif 'other current asset' in type_lower:
            return 'OTHER_CURRENT_ASSET'
        elif 'fixed asset' in type_lower:
            return 'FIXED_ASSET'
        elif 'asset' in type_lower:
            return 'OTHER_CURRENT_ASSET'
        elif 'other current liabilit' in type_lower:
            return 'OTHER_CURRENT_LIABILITY'
        elif 'current liabilit' in type_lower:
            return 'CURRENT_LIABILITY'
        elif 'long term liabilit' in type_lower:
            return 'LONG_TERM_LIABILITY'
        elif 'liabilit' in type_lower:
            return 'OTHER_CURRENT_LIABILITY'
        else:
            return 'EXPENSE'

    def create_account_object(self, name: str, type_str: str, detail_type: str,
                              description: Optional[str] = None,
                              balance: float = 0.0,
                              parent_id: Optional[str] = None) -> Dict[str, Any]:
        """Create a QuickBooks-style account object"""
        classification = self.get_classification_from_type(type_str)
        account_type = self.get_account_type_from_type(type_str)

        account_subtype = detail_type.strip()
        account_subtype = account_subtype.replace('/', '')
        account_subtype = account_subtype.replace('&', 'And')
        account_subtype = account_subtype.replace(' ', '')

        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000+00:00')

        # Detect sub-account from colon-separated name
        is_sub_account = ':' in name
        if is_sub_account:
            fully_qualified_name = name
            leaf_name = name.split(':')[-1].strip()
            parent_ref = {"value": parent_id, "name": None, "type": None} if parent_id else None
        else:
            fully_qualified_name = name
            leaf_name = name
            parent_ref = None

        return {
            "id": self.generate_account_id(),
            "syncToken": "0",
            "metaData": {
                "createdByRef": None,
                "createTime": timestamp,
                "lastModifiedByRef": None,
                "lastUpdatedTime": timestamp,
                "lastChangedInQB": None,
                "synchronized": None
            },
            "customField": [],
            "attachableRef": [],
            "domain": "QBO",
            "status": None,
            "sparse": False,
            "name": leaf_name,
            "subAccount": is_sub_account,
            "parentRef": parent_ref,
            "description": description if description else None,
            "fullyQualifiedName": fully_qualified_name,
            "accountAlias": None,
            "txnLocationType": None,
            "active": True,
            "classification": classification,
            "accountType": account_type,
            "accountSubType": account_subtype,
            "accountPurposes": [],
            "acctNum": None,
            "acctNumExtn": None,
            "bankNum": None,
            "openingBalance": None,
            "openingBalanceDate": None,
            "currentBalance": balance,
            "currentBalanceWithSubAccounts": balance,
            "currencyRef": {
                "value": "USD",
                "name": "United States Dollar",
                "type": None
            },
            "taxAccount": None,
            "taxCodeRef": None,
            "onlineBankingEnabled": None,
            "journalCodeRef": None,
            "accountEx": None,
            "finame": None
        }

    def parse_csv(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse CSV file and convert to account objects"""
        accounts = []
        # Track parent account IDs for sub-account references
        parent_ids = {}  # Maps parent name to its generated ID

        with open(filepath, 'r', encoding='utf-8') as f:
            all_lines = f.readlines()

        print(f"[CSV Parser] Total lines in file: {len(all_lines)}")

        header_line_idx = -1
        for idx, line in enumerate(all_lines):
            line_lower = line.lower()
            if 'full name' in line_lower or ('name' in line_lower and 'type' in line_lower):
                header_line_idx = idx
                print(f"[CSV Parser] Found header at line {idx}: {line.strip()}")
                break

        if header_line_idx == -1:
            print(f"[CSV Parser] WARNING: No header row found. First 5 lines:")
            for i, line in enumerate(all_lines[:5]):
                print(f"  Line {i}: {line.strip()}")
            return accounts

        from io import StringIO
        csv_content = ''.join(all_lines[header_line_idx:])
        reader = csv.DictReader(StringIO(csv_content))

        print(f"[CSV Parser] Column headers: {reader.fieldnames}")

        row_count = 0
        skipped_count = 0

        for row in reader:
            row_count += 1

            full_name = (row.get('Full name') or row.get('Name') or
                         row.get('FULL NAME') or row.get('NAME') or '').strip()

            if not full_name:
                skipped_count += 1
                continue

            if full_name == 'TOTAL':
                skipped_count += 1
                continue

            if any(keyword in full_name.lower() for keyword in ['basis', 'gmtz', 'accrual', 'cash']):
                skipped_count += 1
                continue

            type_str = (row.get('Type') or row.get('TYPE') or
                        row.get('Account Type') or '').strip()
            detail_type = (row.get('Detail type') or row.get('Detail Type') or
                           row.get('DETAIL TYPE') or row.get('Sub Type') or '').strip()
            description = (row.get('Description') or row.get('DESCRIPTION') or '').strip()

            balance_str = (row.get('Total balance') or row.get('Balance') or
                           row.get('TOTAL BALANCE') or row.get('Current Balance') or '0')
            balance_str = str(balance_str).replace('$', '').replace(',', '').strip()
            try:
                balance = float(balance_str) if balance_str else 0.0
            except ValueError:
                balance = 0.0

            # Determine parent ID for sub-accounts
            parent_id = None
            if ':' in full_name:
                parent_name = full_name.rsplit(':', 1)[0]
                parent_id = parent_ids.get(parent_name)

            account = self.create_account_object(
                name=full_name,
                type_str=type_str,
                detail_type=detail_type or 'Other',
                description=description,
                balance=balance,
                parent_id=parent_id
            )
            accounts.append(account)

            # Track this account's ID for potential sub-accounts
            parent_ids[full_name] = account['id']

        print(f"[CSV Parser] Processed {row_count} rows, created {len(accounts)} accounts, skipped {skipped_count} rows")

        return accounts

    def parse_xlsx(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse XLSX file and convert to account objects"""
        self.check_xlsx_support()

        accounts = []
        parent_ids = {}
        workbook = openpyxl.load_workbook(filepath, data_only=True)
        sheet = workbook.active

        header_row = None
        for idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
            if row and 'Full name' in str(row):
                header_row = idx
                break

        if not header_row:
            raise ValueError("Could not find header row in XLSX file")

        headers = list(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
        col_map = {header: idx for idx, header in enumerate(headers) if header}

        for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
            if not row or not row[col_map.get('Full name', 0)]:
                continue

            name = str(row[col_map.get('Full name', 0)])
            if name == 'TOTAL':
                continue

            if any(keyword in name.lower() for keyword in ['basis', 'gmtz', 'accrual', 'cash']):
                continue

            balance_str = str(row[col_map.get('Total balance', 4)] or '0')
            balance_str = balance_str.replace('$', '').replace(',', '')
            try:
                balance = float(balance_str) if balance_str else 0.0
            except ValueError:
                balance = 0.0

            # Determine parent ID for sub-accounts
            parent_id = None
            if ':' in name:
                parent_name = name.rsplit(':', 1)[0]
                parent_id = parent_ids.get(parent_name)

            account = self.create_account_object(
                name=name,
                type_str=row[col_map.get('Type', 1)] or '',
                detail_type=row[col_map.get('Detail type', 2)] or '',
                description=row[col_map.get('Description', 3)],
                balance=balance,
                parent_id=parent_id
            )
            accounts.append(account)
            parent_ids[name] = account['id']

        return accounts

    def parse_pdf(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse PDF file and convert to account objects"""
        self.check_pdf_support()

        accounts = []

        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split('\n')

                header_idx = -1
                for i, line in enumerate(lines):
                    if 'FULL NAME' in line and 'TYPE' in line and 'DETAIL TYPE' in line:
                        header_idx = i
                        break

                if header_idx == -1:
                    continue

                for i in range(header_idx + 1, len(lines)):
                    line = lines[i].strip()

                    if not line or line.startswith('TOTAL'):
                        break

                    if any(keyword in line.lower() for keyword in ['basis', 'gmtz', 'accrual', 'cash']):
                        continue

                    type_patterns = ['Equity', 'Expenses', 'Income', 'Other Current Assets',
                                     'Other Expense', 'Other Income']

                    type_start = -1
                    found_type = None
                    for pattern in type_patterns:
                        idx = line.find(pattern)
                        if idx > 0:
                            type_start = idx
                            found_type = pattern
                            break

                    if type_start > 0:
                        name = line[:type_start].strip()
                        rest = line[type_start + len(found_type):].strip()

                        detail_type = rest
                        detail_type = re.sub(r'\s*\$?[\d,]+\.?\d*\s*$', '', detail_type).strip()

                        if not detail_type:
                            if 'Expense' in found_type:
                                detail_type = 'Other Miscellaneous Service Cost'
                            elif 'Income' in found_type:
                                detail_type = 'Service/Fee Income'
                            elif 'Asset' in found_type:
                                detail_type = 'Other Current Assets'
                            else:
                                detail_type = 'Other'

                        account = self.create_account_object(
                            name=name,
                            type_str=found_type,
                            detail_type=detail_type,
                            balance=0.0
                        )
                        accounts.append(account)

        return accounts


def main():
    parser = argparse.ArgumentParser(description='Convert account documents to JSON format')
    parser.add_argument('input', help='Input file (CSV, XLSX, or PDF)')
    parser.add_argument('-o', '--output', help='Output JSON file (default: print to stdout)')
    parser.add_argument('--batch', action='store_true', help='Process all files in a directory')

    args = parser.parse_args()

    converter = AccountsConverter()

    if args.batch:
        input_path = Path(args.input)
        if not input_path.is_dir():
            print(f"Error: {input_path} is not a directory", file=sys.stderr)
            sys.exit(1)

        output_dir = Path(args.output) if args.output else input_path.parent / 'converted'
        output_dir.mkdir(exist_ok=True)

        for file in input_path.glob('*'):
            if file.suffix.lower() in ['.csv', '.xlsx', '.pdf']:
                try:
                    output_file = output_dir / f"{file.stem}.json"
                    result = converter.convert_to_json(file, output_file)
                    print(result)
                except Exception as e:
                    print(f"Error processing {file}: {e}", file=sys.stderr)
    else:
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
