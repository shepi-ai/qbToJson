#!/usr/bin/env python3
"""
Customer Concentration (Sales by Customer Summary) Converter
Converts CSV, XLSX, and PDF Sales by Customer Summary reports to simplified JSON array format
"""

import csv
import re
import sys
from pathlib import Path
import argparse
from typing import Dict, List, Any, Optional

from base_converter import BaseConverter

# Import guards from base (openpyxl, pdfplumber)
from base_converter import XLSX_SUPPORT, PDF_SUPPORT
if XLSX_SUPPORT:
    import openpyxl
if PDF_SUPPORT:
    import pdfplumber


class CustomerConcentrationConverter(BaseConverter):
    """Converts Sales by Customer Summary reports to simplified JSON array format"""

    def __init__(self, **kwargs):
        # Concentration converters don't use account lookup
        super().__init__(use_account_lookup=False)

    def calculate_percentages(self, customers: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Calculate percentage of total for each customer"""
        grand_total = sum(c['revenue'] for c in customers)

        for customer in customers:
            if grand_total > 0:
                customer['percentage'] = (customer['revenue'] / grand_total) * 100
            else:
                customer['percentage'] = 0.0

        customers.sort(key=lambda x: x['revenue'], reverse=True)
        return customers

    def parse_csv(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse CSV file and convert to simplified JSON array format"""
        customer_map = {}

        with open(filepath, 'r', encoding='utf-8') as f:
            # Skip header lines
            for _ in range(4):
                f.readline()

            reader = csv.DictReader(f)
            current_parent = None

            for row in reader:
                customer_name = row.get('Customer', '').strip()

                # Skip empty rows or TOTAL row
                if not customer_name or customer_name.upper() == 'TOTAL':
                    break

                # Check if this is a "Total for" row (sub-customer total)
                if customer_name.startswith('Total for '):
                    parent_name = customer_name.replace('Total for ', '')
                    if parent_name in customer_map:
                        total = self.parse_amount(row.get('Total', '0'))
                        customer_map[parent_name]['revenue'] = total
                    current_parent = None
                    continue

                total = self.parse_amount(row.get('Total', '0'))

                if total == 0.0:
                    current_parent = customer_name
                    customer_map[customer_name] = {
                        'customerName': customer_name,
                        'revenue': 0.0,
                        'percentage': 0.0
                    }
                    continue

                if current_parent and current_parent in customer_map:
                    customer_map[current_parent]['revenue'] += total
                else:
                    if customer_name not in customer_map:
                        customer_map[customer_name] = {
                            'customerName': customer_name,
                            'revenue': total,
                            'percentage': 0.0
                        }

        return self.calculate_percentages(list(customer_map.values()))

    def parse_xlsx(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse XLSX file and convert to simplified JSON array format"""
        self.check_xlsx_support()

        customer_map = {}

        workbook = openpyxl.load_workbook(filepath, data_only=True)
        sheet = workbook.active

        # Find header row
        header_row = None
        for idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
            if row and any('Customer' in str(cell) for cell in row if cell):
                header_row = idx
                break

        if not header_row:
            raise ValueError("Could not find header row in XLSX file")

        headers = list(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
        col_map = {str(header).strip(): idx for idx, header in enumerate(headers) if header}

        current_parent = None

        for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
            if not row or not row[col_map.get('Customer', 0)]:
                continue

            customer_name = str(row[col_map.get('Customer', 0)]).strip()

            if customer_name.upper() == 'CUSTOMER':
                continue

            if 'Sandbox Company' in customer_name or 'Sales by Customer' in customer_name:
                continue

            if any(month in customer_name for month in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']):
                continue

            if customer_name.upper() == 'TOTAL':
                break

            if customer_name.startswith('Total for '):
                parent_name = customer_name.replace('Total for ', '')
                if parent_name in customer_map:
                    total = self.parse_amount(str(row[col_map.get('Total', 1)] or '0'))
                    customer_map[parent_name]['revenue'] = total
                current_parent = None
                continue

            total = self.parse_amount(str(row[col_map.get('Total', 1)] or '0'))

            if total == 0.0:
                current_parent = customer_name
                customer_map[customer_name] = {
                    'customerName': customer_name,
                    'revenue': 0.0,
                    'percentage': 0.0
                }
                continue

            if current_parent and current_parent in customer_map:
                customer_map[current_parent]['revenue'] += total
            else:
                if customer_name not in customer_map:
                    customer_map[customer_name] = {
                        'customerName': customer_name,
                        'revenue': total,
                        'percentage': 0.0
                    }

        return self.calculate_percentages(list(customer_map.values()))

    def parse_pdf(self, filepath: Path) -> List[Dict[str, Any]]:
        """Parse PDF file and convert to simplified JSON array format"""
        self.check_pdf_support()

        customer_map = {}

        print("[CUSTOMER-CONC-PARSER] Starting PDF parse")

        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split('\n')
                print(f"[CUSTOMER-CONC-PARSER] Page has {len(lines)} lines")

                header_idx = -1
                for i, line in enumerate(lines):
                    if 'CUSTOMER' in line.upper() and 'TOTAL' in line.upper():
                        header_idx = i
                        print(f"[CUSTOMER-CONC-PARSER] Found header at line {i}")
                        break

                if header_idx == -1:
                    continue

                current_parent = None

                for line in lines[header_idx + 1:]:
                    original_line = line
                    is_indented = line and line[0] == ' '

                    line = line.strip()
                    if not line:
                        continue

                    if line.upper().startswith('TOTAL') and not line.startswith('Total for'):
                        print(f"[CUSTOMER-CONC-PARSER] Found TOTAL line, stopping parse")
                        break

                    if line.startswith('Total for '):
                        parent_name = line.replace('Total for ', '').split()[0:3]
                        parent_name = ' '.join(parent_name)
                        amounts = re.findall(r'[\d,]+\.\d{2}', line)
                        if amounts and parent_name in customer_map:
                            total = self.parse_amount(amounts[-1])
                            customer_map[parent_name]['revenue'] = total
                            print(f"[CUSTOMER-CONC-PARSER] Updated parent '{parent_name}' total: ${total}")
                        current_parent = None
                        continue

                    amounts = re.findall(r'-?[\d,]+\.\d{2}', line)

                    if not amounts:
                        if not is_indented:
                            current_parent = line.strip()
                            customer_map[current_parent] = {
                                'customerName': current_parent,
                                'revenue': 0.0,
                                'percentage': 0.0
                            }
                            print(f"[CUSTOMER-CONC-PARSER] Starting parent customer: {current_parent}")
                        continue

                    customer_line = line
                    for amt in amounts:
                        customer_line = customer_line.replace(amt, '', 1)
                    customer_name = customer_line.strip()

                    total = self.parse_amount(amounts[-1])

                    if current_parent:
                        customer_map[current_parent]['revenue'] += total
                        print(f"[CUSTOMER-CONC-PARSER] Added sub-customer to '{current_parent}': {customer_name} (${total})")
                    else:
                        if customer_name and customer_name not in customer_map:
                            customer_map[customer_name] = {
                                'customerName': customer_name,
                                'revenue': total,
                                'percentage': 0.0
                            }
                            print(f"[CUSTOMER-CONC-PARSER] Added customer: {customer_name} (${total})")

        print(f"[CUSTOMER-CONC-PARSER] Parsed {len(customer_map)} customers")

        return self.calculate_percentages(list(customer_map.values()))


def main():
    parser = argparse.ArgumentParser(description='Convert Sales by Customer Summary reports to JSON format')
    parser.add_argument('input', help='Input file (CSV, XLSX, or PDF)')
    parser.add_argument('-o', '--output', help='Output JSON file (default: print to stdout)')
    parser.add_argument('--batch', action='store_true', help='Process all files in a directory')

    args = parser.parse_args()

    converter = CustomerConcentrationConverter()

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
                    output_file = output_dir / f"{file.stem}_customer_concentration.json"
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
