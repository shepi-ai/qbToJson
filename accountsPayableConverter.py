#!/usr/bin/env python3
"""
Accounts Payable (A/P Aging Summary) Converter
Converts CSV, XLSX, and PDF A/P Aging Summary reports to QuickBooks JSON format
"""

import json
import csv
import sys
import os
from datetime import datetime, timezone
from pathlib import Path
import argparse
import re
from typing import Dict, List, Any, Optional

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


class AccountsPayableConverter:
    """Converts A/P Aging Summary reports to QuickBooks-style JSON format"""
    
    def __init__(self):
        self.report_date = datetime.now(timezone.utc).strftime('%Y-%m-%d')
    
    def create_header(self, report_date: str = None) -> Dict[str, Any]:
        """Create QuickBooks-style header for AgedPayables report"""
        if report_date is None:
            report_date = self.report_date
        
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000+00:00')
        
        return {
            "time": timestamp,
            "reportName": "AgedPayables",
            "dateMacro": "today",
            "reportBasis": None,
            "startPeriod": report_date,
            "endPeriod": report_date,
            "summarizeColumnsBy": "Total",
            "currency": "USD",
            "customer": None,
            "vendor": None,
            "employee": None,
            "item": None,
            "clazz": None,
            "department": None,
            "option": [
                {"name": "report_date", "value": report_date},
                {"name": "NoReportData", "value": "false"}
            ]
        }
    
    def create_columns(self) -> Dict[str, List[Dict[str, Any]]]:
        """Create column definitions for A/P Aging report"""
        return {
            "column": [
                {
                    "colTitle": "",
                    "colType": "Vendor",
                    "metaData": [],
                    "columns": None
                },
                {
                    "colTitle": "Current",
                    "colType": "Money",
                    "metaData": [{"name": "ColKey", "value": "current"}],
                    "columns": None
                },
                {
                    "colTitle": "1 - 30",
                    "colType": "Money",
                    "metaData": [{"name": "ColKey", "value": "0"}],
                    "columns": None
                },
                {
                    "colTitle": "31 - 60",
                    "colType": "Money",
                    "metaData": [{"name": "ColKey", "value": "1"}],
                    "columns": None
                },
                {
                    "colTitle": "61 - 90",
                    "colType": "Money",
                    "metaData": [{"name": "ColKey", "value": "2"}],
                    "columns": None
                },
                {
                    "colTitle": "91 and over",
                    "colType": "Money",
                    "metaData": [{"name": "ColKey", "value": "3"}],
                    "columns": None
                },
                {
                    "colTitle": "Total",
                    "colType": "Money",
                    "metaData": [{"name": "ColKey", "value": "total"}],
                    "columns": None
                }
            ]
        }
    
    def create_vendor_row(self, vendor_name: str, vendor_id: str, 
                         current: float, days_1_30: float, days_31_60: float,
                         days_61_90: float, days_91_over: float, total: float) -> Dict[str, Any]:
        """Create a vendor row with aging data"""
        return {
            "id": None,
            "parentId": None,
            "header": None,
            "rows": None,
            "summary": None,
            "colData": [
                {"attributes": None, "value": vendor_name, "id": vendor_id, "href": None},
                {"attributes": None, "value": str(current) if current else "", "id": None, "href": None},
                {"attributes": None, "value": str(days_1_30) if days_1_30 else "", "id": None, "href": None},
                {"attributes": None, "value": str(days_31_60) if days_31_60 else "", "id": None, "href": None},
                {"attributes": None, "value": str(days_61_90) if days_61_90 else "", "id": None, "href": None},
                {"attributes": None, "value": str(days_91_over) if days_91_over else "", "id": None, "href": None},
                {"attributes": None, "value": str(total), "id": None, "href": None}
            ],
            "type": None,
            "group": None
        }
    
    def create_total_row(self, current: float, days_1_30: float, days_31_60: float,
                        days_61_90: float, days_91_over: float, total: float) -> Dict[str, Any]:
        """Create the grand total summary row"""
        return {
            "id": None,
            "parentId": None,
            "header": None,
            "rows": None,
            "summary": {
                "colData": [
                    {"attributes": None, "value": "TOTAL", "id": None, "href": None},
                    {"attributes": None, "value": str(current), "id": None, "href": None},
                    {"attributes": None, "value": str(days_1_30), "id": None, "href": None},
                    {"attributes": None, "value": str(days_31_60), "id": None, "href": None},
                    {"attributes": None, "value": str(days_61_90), "id": None, "href": None},
                    {"attributes": None, "value": str(days_91_over), "id": None, "href": None},
                    {"attributes": None, "value": str(total), "id": None, "href": None}
                ]
            },
            "colData": [],
            "type": "SECTION",
            "group": "GrandTotal"
        }
    
    def parse_amount(self, value: str) -> float:
        """Parse monetary amount from string"""
        if not value or value.strip() == '':
            return 0.0
        # Remove currency symbols, commas, and whitespace
        clean_value = value.replace('$', '').replace(',', '').replace('"', '').strip()
        try:
            return float(clean_value)
        except ValueError:
            return 0.0
    
    def parse_csv(self, filepath: Path) -> Dict[str, Any]:
        """Parse CSV file and convert to QuickBooks JSON format"""
        vendors = []
        totals = {
            'current': 0.0,
            '1_30': 0.0,
            '31_60': 0.0,
            '61_90': 0.0,
            '91_over': 0.0,
            'total': 0.0
        }
        vendor_id = 1
        
        with open(filepath, 'r', encoding='utf-8') as f:
            # Skip header lines
            for _ in range(4):
                f.readline()
            
            reader = csv.DictReader(f)
            for row in reader:
                vendor_name = row.get('Vendor', '').strip()
                
                # Skip empty rows or TOTAL row
                if not vendor_name or vendor_name.upper() == 'TOTAL':
                    # If it's TOTAL, capture the totals
                    if vendor_name.upper() == 'TOTAL':
                        totals['current'] = self.parse_amount(row.get('CURRENT', '0'))
                        totals['1_30'] = self.parse_amount(row.get('1 - 30', '0'))
                        totals['31_60'] = self.parse_amount(row.get('31 - 50', '0'))  # Note: CSV uses 31-50
                        totals['61_90'] = self.parse_amount(row.get('51 - 60', '0'))  # Note: CSV uses 51-60
                        totals['91_over'] = self.parse_amount(row.get('91 AND OVER', '0'))
                        totals['total'] = self.parse_amount(row.get('Total', '0'))
                    break
                
                # Parse aging bucket amounts
                current = self.parse_amount(row.get('CURRENT', '0'))
                days_1_30 = self.parse_amount(row.get('1 - 30', '0'))
                days_31_60 = self.parse_amount(row.get('31 - 50', '0'))  # Note: CSV format uses different ranges
                days_61_90 = self.parse_amount(row.get('51 - 60', '0'))
                days_91_over = self.parse_amount(row.get('91 AND OVER', '0'))
                total = self.parse_amount(row.get('Total', '0'))
                
                vendor_row = self.create_vendor_row(
                    vendor_name, str(vendor_id),
                    current, days_1_30, days_31_60, days_61_90, days_91_over, total
                )
                vendors.append(vendor_row)
                vendor_id += 1
        
        # Add grand total row
        total_row = self.create_total_row(
            totals['current'], totals['1_30'], totals['31_60'],
            totals['61_90'], totals['91_over'], totals['total']
        )
        vendors.append(total_row)
        
        return {
            "header": self.create_header(),
            "columns": self.create_columns(),
            "rows": {"row": vendors}
        }
    
    def parse_xlsx(self, filepath: Path) -> Dict[str, Any]:
        """Parse XLSX file and convert to QuickBooks JSON format"""
        if not XLSX_SUPPORT:
            raise ImportError("openpyxl is required for XLSX support. Install with: pip install openpyxl")
        
        vendors = []
        totals = {
            'current': 0.0,
            '1_30': 0.0,
            '31_60': 0.0,
            '61_90': 0.0,
            '91_over': 0.0,
            'total': 0.0
        }
        vendor_id = 1
        
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active
        
        # Find header row
        header_row = None
        for idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
            if row and any('Vendor' in str(cell) for cell in row if cell):
                header_row = idx
                break
        
        if not header_row:
            raise ValueError("Could not find header row in XLSX file")
        
        # Get column indices
        headers = list(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
        col_map = {str(header).strip(): idx for idx, header in enumerate(headers) if header}
        
        # Parse data rows
        for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
            if not row or not row[col_map.get('Vendor', 0)]:
                continue
            
            vendor_name = str(row[col_map.get('Vendor', 0)]).strip()
            
            if vendor_name.upper() == 'TOTAL':
                totals['current'] = self.parse_amount(str(row[col_map.get('CURRENT', 1)] or '0'))
                totals['1_30'] = self.parse_amount(str(row[col_map.get('1 - 30', 2)] or '0'))
                totals['31_60'] = self.parse_amount(str(row[col_map.get('31 - 50', 3)] or '0'))
                totals['61_90'] = self.parse_amount(str(row[col_map.get('51 - 60', 4)] or '0'))
                totals['91_over'] = self.parse_amount(str(row[col_map.get('91 AND OVER', 6)] or '0'))
                totals['total'] = self.parse_amount(str(row[col_map.get('Total', 7)] or '0'))
                break
            
            current = self.parse_amount(str(row[col_map.get('CURRENT', 1)] or '0'))
            days_1_30 = self.parse_amount(str(row[col_map.get('1 - 30', 2)] or '0'))
            days_31_60 = self.parse_amount(str(row[col_map.get('31 - 50', 3)] or '0'))
            days_61_90 = self.parse_amount(str(row[col_map.get('51 - 60', 4)] or '0'))
            days_91_over = self.parse_amount(str(row[col_map.get('91 AND OVER', 6)] or '0'))
            total = self.parse_amount(str(row[col_map.get('Total', 7)] or '0'))
            
            vendor_row = self.create_vendor_row(
                vendor_name, str(vendor_id),
                current, days_1_30, days_31_60, days_61_90, days_91_over, total
            )
            vendors.append(vendor_row)
            vendor_id += 1
        
        total_row = self.create_total_row(
            totals['current'], totals['1_30'], totals['31_60'],
            totals['61_90'], totals['91_over'], totals['total']
        )
        vendors.append(total_row)
        
        return {
            "header": self.create_header(),
            "columns": self.create_columns(),
            "rows": {"row": vendors}
        }
    
    def parse_pdf(self, filepath: Path) -> Dict[str, Any]:
        """Parse PDF file and convert to QuickBooks JSON format"""
        if not PDF_SUPPORT:
            raise ImportError("pdfplumber is required for PDF support. Install with: pip install pdfplumber")
        
        vendors = []
        totals = {
            'current': 0.0,
            '1_30': 0.0,
            '31_60': 0.0,
            '61_90': 0.0,
            '91_over': 0.0,
            'total': 0.0
        }
        vendor_id = 1
        
        print("[AP-PARSER] Starting PDF parse")
        
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                # Extract text instead of tables (AP PDFs don't have table structure)
                text = page.extract_text()
                if not text:
                    continue
                
                lines = text.split('\n')
                print(f"[AP-PARSER] Page has {len(lines)} lines")
                
                # Find header row
                header_idx = -1
                for i, line in enumerate(lines):
                    if 'VENDOR' in line.upper() and ('CURRENT' in line.upper() or 'TOTAL' in line.upper()):
                        header_idx = i
                        print(f"[AP-PARSER] Found header at line {i}")
                        break
                
                if header_idx == -1:
                    continue
                
                # Parse data lines
                for line in lines[header_idx + 1:]:
                    line = line.strip()
                    if not line:
                        continue
                    
                    # Check for TOTAL line
                    if line.upper().startswith('TOTAL'):
                        # Extract amounts from end of line
                        amounts = re.findall(r'[\d,]+\.\d{2}', line)
                        if len(amounts) >= 6:
                            totals['current'] = self.parse_amount(amounts[0])
                            totals['1_30'] = self.parse_amount(amounts[1])
                            totals['31_60'] = self.parse_amount(amounts[2])
                            totals['61_90'] = self.parse_amount(amounts[3])
                            totals['91_over'] = self.parse_amount(amounts[4])
                            totals['total'] = self.parse_amount(amounts[5])
                        print(f"[AP-PARSER] Found TOTAL row: {totals}")
                        break
                    
                    # Extract amounts from line (up to 6 decimal numbers)
                    amounts = re.findall(r'[\d,]+\.\d{2}', line)
                    
                    if not amounts:
                        continue
                    
                    # Remove amounts from line to get vendor name
                    vendor_line = line
                    for amt in amounts:
                        vendor_line = vendor_line.replace(amt, '', 1)
                    vendor_name = vendor_line.strip()
                    
                    # Parse amounts based on count
                    if len(amounts) == 6:
                        # Full row: current, 1-30, 31-50, 51-60, 61-90, 91+, total
                        current = self.parse_amount(amounts[0])
                        days_1_30 = self.parse_amount(amounts[1])
                        days_31_60 = self.parse_amount(amounts[2])
                        days_61_90 = self.parse_amount(amounts[3])
                        days_91_over = self.parse_amount(amounts[4])
                        total = self.parse_amount(amounts[5])
                    elif len(amounts) == 1:
                        # Just one bucket + total (most common in AP aging)
                        current = 0.0
                        days_1_30 = 0.0
                        days_31_60 = 0.0
                        days_61_90 = 0.0
                        days_91_over = 0.0
                        total = self.parse_amount(amounts[0])
                    elif len(amounts) == 2:
                        # One bucket + total
                        current = 0.0
                        days_1_30 = self.parse_amount(amounts[0])
                        days_31_60 = 0.0
                        days_61_90 = 0.0
                        days_91_over = 0.0
                        total = self.parse_amount(amounts[1])
                    else:
                        # Skip unusual patterns
                        print(f"[AP-PARSER] Skipping line with {len(amounts)} amounts: {line[:50]}")
                        continue
                    
                    vendor_row = self.create_vendor_row(
                        vendor_name, str(vendor_id),
                        current, days_1_30, days_31_60, days_61_90, days_91_over, total
                    )
                    vendors.append(vendor_row)
                    vendor_id += 1
                    print(f"[AP-PARSER] Added vendor: {vendor_name} (${total})")
        
        print(f"[AP-PARSER] Parsed {len(vendors)} vendor rows")
        
        total_row = self.create_total_row(
            totals['current'], totals['1_30'], totals['31_60'],
            totals['61_90'], totals['91_over'], totals['total']
        )
        vendors.append(total_row)
        
        return {
            "header": self.create_header(),
            "columns": self.create_columns(),
            "rows": {"row": vendors}
        }
    
    def convert_file(self, filepath: Path) -> Dict[str, Any]:
        """Convert a file to QuickBooks JSON format based on its extension"""
        ext = filepath.suffix.lower()
        
        if ext == '.csv':
            return self.parse_csv(filepath)
        elif ext == '.xlsx':
            return self.parse_xlsx(filepath)
        elif ext == '.pdf':
            return self.parse_pdf(filepath)
        else:
            raise ValueError(f"Unsupported file format: {ext}")
    
    def convert_to_json(self, filepath: Path, output_path: Optional[Path] = None) -> str:
        """Convert a file to JSON format"""
        result = self.convert_file(filepath)
        
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, indent=2)
            vendor_count = len(result['rows']['row']) - 1  # Exclude total row
            return f"Converted {vendor_count} vendors to {output_path}"
        else:
            return json.dumps(result, indent=2)


def main():
    parser = argparse.ArgumentParser(description='Convert A/P Aging Summary reports to JSON format')
    parser.add_argument('input', help='Input file (CSV, XLSX, or PDF)')
    parser.add_argument('-o', '--output', help='Output JSON file (default: print to stdout)')
    parser.add_argument('--batch', action='store_true', help='Process all files in a directory')
    
    args = parser.parse_args()
    
    converter = AccountsPayableConverter()
    
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
                    output_file = output_dir / f"{file.stem}_ap_aging.json"
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
