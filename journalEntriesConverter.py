import json
import csv
import re
from datetime import datetime, timedelta
import openpyxl
from openpyxl.utils import get_column_letter
import pdfplumber
from account_lookup_client import AccountLookupClient

class JournalEntriesConverter:
    def __init__(self, use_account_lookup=True, api_base_url=None):
        """
        Initialize converter with optional account lookup functionality.
        
        :param use_account_lookup: Boolean to enable/disable account lookup
        :param api_base_url: Optional API base URL for account lookup service
        """
        self.use_account_lookup = use_account_lookup
        if use_account_lookup:
            self.lookup_client = AccountLookupClient(api_base_url=api_base_url)
        self.account_cache = {}
        self.transactions = []
        
    def convert(self, file_path):
        """
        Convert journal entries file to QuickBooks JSON format.
        Automatically detects file type and uses appropriate parser.
        
        :param file_path: Path to the input file
        :return: List of journal entry dictionaries
        """
        # Convert Path object to string if needed
        file_path_str = str(file_path)
        
        if file_path_str.lower().endswith('.csv'):
            return self.parse_csv(file_path)
        elif file_path_str.lower().endswith('.xlsx'):
            return self.parse_xlsx(file_path)
        elif file_path_str.lower().endswith('.pdf'):
            return self.parse_pdf(file_path)
        else:
            raise ValueError("Unsupported file format. Please use CSV, XLSX, or PDF.")
    
    def convert_file(self, file_path):
        """Alias for convert() to match API convention"""
        return self.convert(file_path)
    
    def parse_csv(self, file_path):
        """Parse CSV format journal entries"""
        with open(file_path, 'r', encoding='utf-8-sig') as file:
            # Read all lines
            lines = file.readlines()
            
            # Skip header lines
            start_idx = 0
            for i, line in enumerate(lines):
                if 'Transaction date' in line or 'TRANSACTION DATE' in line:
                    start_idx = i
                    break
            
            # Process data
            current_transaction = None
            current_id = None
            
            for i in range(start_idx + 1, len(lines)):
                line = lines[i].strip()
                if not line or line.startswith('Total for') or line.startswith('TOTAL'):
                    if current_transaction:
                        self.transactions.append(current_transaction)
                        current_transaction = None
                        current_id = None
                    continue
                
                # Split line by comma, handling quoted values
                reader = csv.reader([line])
                parts = list(reader)[0]
                
                if len(parts) < 8:
                    continue
                
                # Check if this is a transaction ID line (first column has a number)
                if parts[0].strip().isdigit():
                    if current_transaction:
                        self.transactions.append(current_transaction)
                    current_id = parts[0].strip()
                    current_transaction = {
                        'id': current_id,
                        'lines': []
                    }
                    continue
                
                # Parse transaction line
                if current_transaction:
                    date_str = parts[1].strip()
                    trans_type = parts[2].strip()
                    num = parts[3].strip()
                    name = parts[4].strip()
                    memo = parts[5].strip()
                    account = parts[6].strip()
                    debit = parts[7].strip().replace(',', '').replace('$', '')
                    credit = parts[8].strip().replace(',', '').replace('$', '') if len(parts) > 8 else ''
                    
                    if date_str and trans_type:
                        # Update transaction header info
                        if not current_transaction.get('date'):
                            current_transaction['date'] = date_str
                            current_transaction['type'] = trans_type
                            current_transaction['num'] = num
                            current_transaction['name'] = name
                            current_transaction['memo'] = memo
                        
                        # Add line item
                        line_item = {
                            'account': account,
                            'description': memo,
                            'debit': float(debit) if debit else 0.0,
                            'credit': float(credit) if credit else 0.0
                        }
                        
                        if line_item['debit'] > 0 or line_item['credit'] > 0:
                            current_transaction['lines'].append(line_item)
            
            # Don't forget the last transaction
            if current_transaction:
                self.transactions.append(current_transaction)
            
        return self.build_json_structure()
    
    def parse_xlsx(self, file_path):
        """Parse XLSX format journal entries"""
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        
        # Find header row
        header_row = None
        for row in range(1, min(20, sheet.max_row + 1)):
            cell_value = str(sheet.cell(row=row, column=2).value or '').upper()
            if 'TRANSACTION DATE' in cell_value:
                header_row = row
                break
        
        if not header_row:
            raise ValueError("Could not find header row in XLSX file")
        
        # Process data
        current_transaction = None
        current_id = None
        
        for row in range(header_row + 1, sheet.max_row + 1):
            # Get values from row
            id_val = sheet.cell(row=row, column=1).value
            date_val = sheet.cell(row=row, column=2).value
            type_val = sheet.cell(row=row, column=3).value
            num_val = sheet.cell(row=row, column=4).value
            name_val = sheet.cell(row=row, column=5).value
            memo_val = sheet.cell(row=row, column=6).value
            account_val = sheet.cell(row=row, column=7).value
            debit_val = sheet.cell(row=row, column=8).value
            credit_val = sheet.cell(row=row, column=9).value
            
            # Check for transaction ID
            if id_val and str(id_val).strip().isdigit():
                if current_transaction:
                    self.transactions.append(current_transaction)
                current_id = str(id_val).strip()
                current_transaction = {
                    'id': current_id,
                    'lines': []
                }
                continue
            
            # Check for total line
            if account_val and str(account_val).startswith('Total for'):
                if current_transaction:
                    self.transactions.append(current_transaction)
                    current_transaction = None
                    current_id = None
                continue
            
            # Parse transaction line
            if current_transaction and date_val:
                # Update transaction header info
                if not current_transaction.get('date'):
                    if isinstance(date_val, datetime):
                        current_transaction['date'] = date_val.strftime('%m/%d/%Y')
                    else:
                        current_transaction['date'] = str(date_val)
                    current_transaction['type'] = str(type_val or '')
                    current_transaction['num'] = str(num_val or '')
                    current_transaction['name'] = str(name_val or '')
                    current_transaction['memo'] = str(memo_val or '')
                
                # Add line item
                if account_val:
                    debit = float(debit_val) if debit_val else 0.0
                    credit = float(credit_val) if credit_val else 0.0
                    
                    if debit > 0 or credit > 0:
                        line_item = {
                            'account': str(account_val),
                            'description': str(memo_val or ''),
                            'debit': debit,
                            'credit': credit
                        }
                        current_transaction['lines'].append(line_item)
        
        # Don't forget the last transaction
        if current_transaction:
            self.transactions.append(current_transaction)
        
        return self.build_json_structure()
    
    def parse_pdf(self, file_path):
        """Parse PDF format journal entries - QuickBooks Journal report format"""
        all_text = []
        
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text.append(text)
        
        full_text = '\n'.join(all_text)
        lines = full_text.split('\n')
        
        # Find header line
        start_idx = 0
        for i, line in enumerate(lines):
            if 'TRANSACTION DATE' in line.upper():
                start_idx = i + 1
                break
        
        # Process data
        current_transaction = None
        current_id = None
        
        for i in range(start_idx, len(lines)):
            line_orig = lines[i]
            line = line_orig.strip()
            
            # Skip empty lines and page footers/headers
            if not line or 'Accrual Basis' in line or 'Journal' in line and len(line) < 50:
                continue
            
            # Check for "Total for" line - marks end of transaction
            if line.startswith('Total for'):
                if current_transaction and current_transaction.get('lines'):
                    self.transactions.append(current_transaction)
                current_transaction = None
                current_id = None
                continue
            
            # Check if this is a transaction ID line (just a number at start of line)
            if line.isdigit() and len(line) <= 4:
                if current_transaction and current_transaction.get('lines'):
                    self.transactions.append(current_transaction)
                current_id = line
                current_transaction = {
                    'id': current_id,
                    'lines': []
                }
                continue
            
            # Parse transaction detail line
            if current_transaction:
                # Use 2+ spaces as delimiter (QuickBooks PDFs use spacing to separate columns)
                parts = re.split(r'\s{2,}', line)
                
                # Filter out empty parts
                parts = [p.strip() for p in parts if p.strip()]
                
                if len(parts) < 2:
                    continue
                
                # Check for date at start (transaction header line)
                date_match = re.match(r'^(\d{2}/\d{2}/\d{4})$', parts[0])
                
                if date_match:
                    # This is a main transaction line with date
                    date_str = date_match.group(1)
                    
                    # Set transaction header info if not set
                    if not current_transaction.get('date'):
                        current_transaction['date'] = date_str
                        if len(parts) > 1:
                            current_transaction['type'] = parts[1]
                        if len(parts) > 2 and parts[2].isdigit():
                            current_transaction['num'] = parts[2]
                        if len(parts) > 3 and not parts[3].replace(',', '').replace('.', '').isdigit():
                            current_transaction['name'] = parts[3]
                    
                    # Extract account and amounts from this line
                    # Format: DATE TYPE NUM NAME MEMO ACCOUNT DEBIT CREDIT
                    # Last 1 or 2 items are amounts, item before that is account
                    debit = 0.0
                    credit = 0.0
                    account = ''
                    memo = ''
                    
                    # Check last two parts for amounts
                    if len(parts) >= 2:
                        # Try to parse last part as amount
                        last_part = parts[-1].replace(',', '')
                        if re.match(r'^\d+\.?\d*$', last_part):
                            # Last part is an amount - could be debit or credit
                            # Check if second-to-last is also an amount
                            if len(parts) >= 3:
                                second_last = parts[-2].replace(',', '')
                                if re.match(r'^\d+\.?\d*$', second_last):
                                    # Both are amounts: debit and credit
                                    debit = float(parts[-2].replace(',', ''))
                                    credit = float(parts[-1].replace(',', ''))
                                    # Everything before last 2 items (after date/type/num/name)
                                    if len(parts) > 6:
                                        memo_account_parts = parts[4:-2]
                                    elif len(parts) > 4:
                                        memo_account_parts = parts[4:-2]
                                    else:
                                        memo_account_parts = []
                                    
                                    # Last part before amounts is account, rest is memo
                                    if memo_account_parts:
                                        account = memo_account_parts[-1]
                                        if len(memo_account_parts) > 1:
                                            memo = ' '.join(memo_account_parts[:-1])
                                else:
                                    # Only last part is amount - it's a credit
                                    credit = float(parts[-1].replace(',', ''))
                                    # Second-to-last is account
                                    account = parts[-2]
                                    if len(parts) > 5:
                                        memo = ' '.join(parts[4:-2])
                            else:
                                # Single amount
                                credit = float(parts[-1].replace(',', ''))
                                if len(parts) > 1:
                                    account = parts[-2] if len(parts) > 1 else ''
                    
                    if account and (debit > 0 or credit > 0):
                        line_item = {
                            'account': account,
                            'description': memo,
                            'debit': debit,
                            'credit': credit
                        }
                        current_transaction['lines'].append(line_item)
                
                else:
                    # This is a continuation line (no date) - indented transaction detail
                    # Format: spaces + TYPE NUM NAME MEMO ACCOUNT DEBIT CREDIT
                    # Or just: spaces + MEMO ACCOUNT DEBIT CREDIT
                    
                    debit = 0.0
                    credit = 0.0
                    account = ''
                    memo = ''
                    
                    # Check last two parts for amounts
                    if len(parts) >= 2:
                        last_part = parts[-1].replace(',', '')
                        if re.match(r'^\d+\.?\d*$', last_part):
                            # Last part is an amount
                            if len(parts) >= 3:
                                second_last = parts[-2].replace(',', '')
                                if re.match(r'^\d+\.?\d*$', second_last):
                                    # Both are amounts
                                    debit = float(parts[-2].replace(',', ''))
                                    credit = float(parts[-1].replace(',', ''))
                                    # Everything before amounts
                                    if len(parts) > 2:
                                        memo_account_parts = parts[:-2]
                                        account = memo_account_parts[-1]
                                        if len(memo_account_parts) > 1:
                                            memo = ' '.join(memo_account_parts[:-1])
                                else:
                                    # Only last is amount (credit)
                                    credit = float(parts[-1].replace(',', ''))
                                    account = parts[-2]
                                    if len(parts) > 2:
                                        memo = ' '.join(parts[:-2])
                            else:
                                credit = float(parts[-1].replace(',', ''))
                                account = parts[0] if len(parts) > 0 else ''
                    
                    if account and (debit > 0 or credit > 0):
                        line_item = {
                            'account': account,
                            'description': memo,
                            'debit': debit,
                            'credit': credit
                        }
                        current_transaction['lines'].append(line_item)
        
        # Don't forget the last transaction
        if current_transaction and current_transaction.get('lines'):
            self.transactions.append(current_transaction)
        
        return self.build_json_structure()
    
    def lookup_account_id(self, account_name):
        """Look up account ID by name"""
        if not self.use_account_lookup or not account_name:
            return None
        
        # Check cache first
        if account_name in self.account_cache:
            return self.account_cache[account_name]
        
        # Lookup via API
        account_id = self.lookup_client.lookup_account(account_name)
        if account_id:
            self.account_cache[account_name] = account_id
        
        return account_id
    
    def build_json_structure(self):
        """Build QuickBooks-compatible JSON structure from parsed transactions"""
        journal_entries = []
        
        for trans in self.transactions:
            if not trans.get('lines'):
                continue
            
            # Parse date
            date_obj = datetime.strptime(trans['date'], '%m/%d/%Y')
            
            # Filter for actual Journal Entry type transactions if needed
            # For now, we'll include all transaction types as they appear in the report
            
            entry = {
                "id": trans['id'],
                "syncToken": "0",
                "metaData": {
                    "createTime": date_obj.strftime('%Y-%m-%dT%H:%M:%S.000+00:00'),
                    "lastUpdatedTime": date_obj.strftime('%Y-%m-%dT%H:%M:%S.000+00:00')
                },
                "customField": [],
                "attachableRef": [],
                "domain": "QBO",
                "sparse": False,
                "docNumber": trans.get('num'),
                "txnDate": date_obj.strftime('%Y-%m-%dT00:00:00.000+00:00'),
                "currencyRef": {
                    "value": "USD",
                    "name": "United States Dollar"
                },
                "privateNote": trans.get('memo', ''),
                "linkedTxn": [],
                "line": [],
                "txnTaxDetail": {
                    "txnTaxCodeRef": None,
                    "totalTax": None,
                    "taxLine": [],
                    "useAutomatedSalesTax": None
                },
                "adjustment": False,
                "globalTaxCalculation": None
            }
            
            # Add transaction type info in a custom field or private note
            if trans.get('type'):
                entry['privateNote'] = f"{trans['type']}: {entry['privateNote']}" if entry['privateNote'] else trans['type']
            
            # Add name to private note if present
            if trans.get('name'):
                if entry['privateNote']:
                    entry['privateNote'] += f" - {trans['name']}"
                else:
                    entry['privateNote'] = trans['name']
            
            # Process line items
            line_num = 0
            for line in trans['lines']:
                # Skip lines with no amounts
                if line['debit'] == 0 and line['credit'] == 0:
                    continue
                
                # Determine posting type and amount
                if line['debit'] > 0:
                    posting_type = "DEBIT"
                    amount = line['debit']
                else:
                    posting_type = "CREDIT"
                    amount = line['credit']
                
                # Look up account ID
                account_id = self.lookup_account_id(line['account'])
                
                line_item = {
                    "id": str(line_num),
                    "description": line['description'],
                    "amount": amount,
                    "linkedTxn": [],
                    "detailType": "JOURNAL_ENTRY_LINE_DETAIL",
                    "journalEntryLineDetail": {
                        "postingType": posting_type,
                        "accountRef": {
                            "value": account_id or str(100 + line_num),  # Default ID if lookup fails
                            "name": line['account']
                        }
                    },
                    "customField": []
                }
                
                entry['line'].append(line_item)
                line_num += 1
            
            # Calculate total amount (should be 0 for balanced entries)
            total_debits = sum(l['debit'] for l in trans['lines'])
            total_credits = sum(l['credit'] for l in trans['lines'])
            entry['totalAmt'] = abs(total_debits - total_credits)
            
            journal_entries.append(entry)
        
        return journal_entries
    
    def save_to_file(self, journal_entries, output_file):
        """Save journal entries to JSON file"""
        with open(output_file, 'w') as f:
            json.dump(journal_entries, f, indent=2)


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='Convert journal entries to QuickBooks JSON format')
    parser.add_argument('input_file', help='Input file path (CSV, XLSX, or PDF)')
    parser.add_argument('-o', '--output', help='Output JSON file path')
    parser.add_argument('--no-lookup', action='store_true', help='Disable account lookup')
    parser.add_argument('--api-url', help='API base URL for account lookup')
    
    args = parser.parse_args()
    
    # Create converter
    converter = JournalEntriesConverter(
        use_account_lookup=not args.no_lookup,
        api_base_url=args.api_url
    )
    
    try:
        # Convert file
        entries = converter.convert(args.input_file)
        
        if args.output:
            # Save to file
            converter.save_to_file(entries, args.output)
            print(f"Successfully converted {len(entries)} journal entries to {args.output}")
        else:
            # Print to stdout
            print(json.dumps(entries, indent=2))
            
    except Exception as e:
        print(f"Error: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())
