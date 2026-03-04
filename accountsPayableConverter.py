#!/usr/bin/env python3
"""
Accounts Payable (A/P Aging Summary) Converter
Converts CSV, XLSX, and PDF A/P Aging Summary reports to QuickBooks JSON format
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


class AccountsPayableConverter(BaseConverter):
    """Converts A/P Aging Summary reports to QuickBooks-style JSON format"""

    def __init__(self, **kwargs):
        super().__init__(use_account_lookup=False)
        self.report_date = datetime.now(timezone.utc).strftime('%Y-%m-%d')

    def create_header(self, report_date: str = None) -> Dict[str, Any]:
        """Create QuickBooks-style header for AgedPayables report"""
        if report_date is None:
            report_date = self.report_date
        return self.make_qb_header(
            "AgedPayables",
            start_period=report_date,
            end_period=report_date,
            options=[
                {"name": "report_date", "value": report_date},
                {"name": "NoReportData", "value": "false"}
            ]
        )

    def create_columns(self) -> Dict[str, List[Dict[str, Any]]]:
        """Create column definitions for A/P Aging report"""
        return {
            "column": [
                {"colTitle": "", "colType": "Vendor", "metaData": [], "columns": None},
                {"colTitle": "Current", "colType": "Money", "metaData": [{"name": "ColKey", "value": "current"}], "columns": None},
                {"colTitle": "1 - 30", "colType": "Money", "metaData": [{"name": "ColKey", "value": "0"}], "columns": None},
                {"colTitle": "31 - 60", "colType": "Money", "metaData": [{"name": "ColKey", "value": "1"}], "columns": None},
                {"colTitle": "61 - 90", "colType": "Money", "metaData": [{"name": "ColKey", "value": "2"}], "columns": None},
                {"colTitle": "91 and over", "colType": "Money", "metaData": [{"name": "ColKey", "value": "3"}], "columns": None},
                {"colTitle": "Total", "colType": "Money", "metaData": [{"name": "ColKey", "value": "total"}], "columns": None},
            ]
        }

    def create_vendor_row(self, vendor_name: str, vendor_id: str,
                          current: float, days_1_30: float, days_31_60: float,
                          days_61_90: float, days_91_over: float, total: float) -> Dict[str, Any]:
        """Create a vendor row with aging data"""
        return {
            "id": None, "parentId": None, "header": None, "rows": None, "summary": None,
            "colData": [
                self.make_coldata_cell(vendor_name, vendor_id),
                self.make_coldata_cell(str(current) if current else ""),
                self.make_coldata_cell(str(days_1_30) if days_1_30 else ""),
                self.make_coldata_cell(str(days_31_60) if days_31_60 else ""),
                self.make_coldata_cell(str(days_61_90) if days_61_90 else ""),
                self.make_coldata_cell(str(days_91_over) if days_91_over else ""),
                self.make_coldata_cell(str(total)),
            ],
            "type": None, "group": None
        }

    def create_total_row(self, current: float, days_1_30: float, days_31_60: float,
                         days_61_90: float, days_91_over: float, total: float) -> Dict[str, Any]:
        """Create the grand total summary row"""
        return {
            "id": None, "parentId": None, "header": None, "rows": None,
            "summary": {
                "colData": [
                    self.make_coldata_cell("TOTAL"),
                    self.make_coldata_cell(str(current)),
                    self.make_coldata_cell(str(days_1_30)),
                    self.make_coldata_cell(str(days_31_60)),
                    self.make_coldata_cell(str(days_61_90)),
                    self.make_coldata_cell(str(days_91_over)),
                    self.make_coldata_cell(str(total)),
                ]
            },
            "colData": [],
            "type": "SECTION", "group": "GrandTotal"
        }

    def parse_csv(self, filepath: Path) -> Dict[str, Any]:
        """Parse CSV file and convert to QuickBooks JSON format"""
        vendors = []
        totals = {'current': 0.0, '1_30': 0.0, '31_60': 0.0, '61_90': 0.0, '91_over': 0.0, 'total': 0.0}
        vendor_id = 1

        with open(filepath, 'r', encoding='utf-8') as f:
            for _ in range(4):
                f.readline()

            reader = csv.DictReader(f)
            for row in reader:
                vendor_name = row.get('Vendor', '').strip()

                if not vendor_name or vendor_name.upper() == 'TOTAL':
                    if vendor_name and vendor_name.upper() == 'TOTAL':
                        totals['current'] = self.parse_amount(row.get('CURRENT', '0'))
                        totals['1_30'] = self.parse_amount(row.get('1 - 30', '0'))
                        totals['31_60'] = self.parse_amount(row.get('31 - 50', '0'))
                        totals['61_90'] = self.parse_amount(row.get('51 - 60', '0'))
                        totals['91_over'] = self.parse_amount(row.get('91 AND OVER', '0'))
                        totals['total'] = self.parse_amount(row.get('Total', '0'))
                    break

                current = self.parse_amount(row.get('CURRENT', '0'))
                days_1_30 = self.parse_amount(row.get('1 - 30', '0'))
                days_31_60 = self.parse_amount(row.get('31 - 50', '0'))
                days_61_90 = self.parse_amount(row.get('51 - 60', '0'))
                days_91_over = self.parse_amount(row.get('91 AND OVER', '0'))
                total = self.parse_amount(row.get('Total', '0'))

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

    def parse_xlsx(self, filepath: Path) -> Dict[str, Any]:
        """Parse XLSX file and convert to QuickBooks JSON format"""
        self.check_xlsx_support()

        vendors = []
        totals = {'current': 0.0, '1_30': 0.0, '31_60': 0.0, '61_90': 0.0, '91_over': 0.0, 'total': 0.0}
        vendor_id = 1

        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active

        header_row = None
        for idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
            if row and any('Vendor' in str(cell) for cell in row if cell):
                header_row = idx
                break

        if not header_row:
            raise ValueError("Could not find header row in XLSX file")

        headers = list(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
        col_map = {str(header).strip(): idx for idx, header in enumerate(headers) if header}

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
        self.check_pdf_support()

        vendors = []
        totals = {'current': 0.0, '1_30': 0.0, '31_60': 0.0, '61_90': 0.0, '91_over': 0.0, 'total': 0.0}
        vendor_id = 1

        print("[AP-PARSER] Starting PDF parse")

        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split('\n')
                print(f"[AP-PARSER] Page has {len(lines)} lines")

                header_idx = -1
                for i, line in enumerate(lines):
                    if 'VENDOR' in line.upper() and ('CURRENT' in line.upper() or 'TOTAL' in line.upper()):
                        header_idx = i
                        print(f"[AP-PARSER] Found header at line {i}")
                        break

                if header_idx == -1:
                    continue

                for line in lines[header_idx + 1:]:
                    line = line.strip()
                    if not line:
                        continue

                    if line.upper().startswith('TOTAL'):
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

                    amounts = re.findall(r'[\d,]+\.\d{2}', line)

                    if not amounts:
                        continue

                    vendor_line = line
                    for amt in amounts:
                        vendor_line = vendor_line.replace(amt, '', 1)
                    vendor_name = vendor_line.strip()

                    if len(amounts) == 6:
                        current = self.parse_amount(amounts[0])
                        days_1_30 = self.parse_amount(amounts[1])
                        days_31_60 = self.parse_amount(amounts[2])
                        days_61_90 = self.parse_amount(amounts[3])
                        days_91_over = self.parse_amount(amounts[4])
                        total = self.parse_amount(amounts[5])
                    elif len(amounts) == 1:
                        current = 0.0
                        days_1_30 = 0.0
                        days_31_60 = 0.0
                        days_61_90 = 0.0
                        days_91_over = 0.0
                        total = self.parse_amount(amounts[0])
                    elif len(amounts) == 2:
                        current = 0.0
                        days_1_30 = self.parse_amount(amounts[0])
                        days_31_60 = 0.0
                        days_61_90 = 0.0
                        days_91_over = 0.0
                        total = self.parse_amount(amounts[1])
                    else:
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

    def convert_to_json(self, filepath: Path, output_path: Optional[Path] = None) -> str:
        """Convert a file to JSON format"""
        import json
        result = self.convert_file(filepath)

        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, indent=2)
            vendor_count = len(result['rows']['row']) - 1
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
