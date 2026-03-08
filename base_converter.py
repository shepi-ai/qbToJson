#!/usr/bin/env python3
"""
Base Converter - Shared infrastructure for all file-to-JSON converters.

Provides:
- File format dispatch (CSV, XLSX, PDF)
- Amount parsing (handles $, commas, parens-as-negatives)
- Account ID generation with optional API lookup
- QB report envelope helpers (colData cells, headers)
- Import guards for optional dependencies (openpyxl, pdfplumber)
"""

import json
import sys
import re
import calendar
from abc import ABC, abstractmethod
from datetime import datetime, timezone, date
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple, Union

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


class BaseConverter(ABC):
    """
    Abstract base class for all file-to-JSON converters.

    Subclasses must implement:
        parse_csv(filepath) -> converted data
        parse_xlsx(filepath) -> converted data
        parse_pdf(filepath) -> converted data
    """

    def __init__(self, use_account_lookup: bool = True, api_base_url: str = "http://localhost:8080"):
        self.account_id_counter = 1
        self.account_id_map: Dict[str, str] = {}
        self.use_account_lookup = use_account_lookup and ACCOUNT_LOOKUP_AVAILABLE
        self.account_lookup_client = None

        if self.use_account_lookup:
            try:
                self.account_lookup_client = get_account_lookup_client(api_base_url)
                if not self.account_lookup_client.is_api_available():
                    print("Warning: Account lookup API is not available. Using generated IDs.", file=sys.stderr)
                    self.use_account_lookup = False
            except Exception as e:
                print(f"Warning: Could not initialize account lookup client: {e}. Using generated IDs.", file=sys.stderr)
                self.use_account_lookup = False

    # ──────────────────────────────────────────────
    # Amount parsing
    # ──────────────────────────────────────────────

    @staticmethod
    def parse_amount(value) -> float:
        """
        Parse a monetary amount from a string or number.

        Handles: $1,234.56  "1,234.56"  (1,234.56) → negative  empty/None → 0.0
        """
        if value is None:
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)

        s = str(value).strip()
        if not s:
            return 0.0

        # Detect parentheses-as-negative: (1,234.56) → -1234.56
        is_negative = s.startswith('(') and s.endswith(')')
        if is_negative:
            s = s[1:-1]

        # Strip $, commas, quotes, whitespace
        s = s.replace('$', '').replace(',', '').replace('"', '').strip()

        try:
            result = float(s)
            return -result if is_negative else result
        except ValueError:
            return 0.0

    # ──────────────────────────────────────────────
    # File dispatch
    # ──────────────────────────────────────────────

    def convert_file(self, filepath: Path) -> Union[List, Dict]:
        """Dispatch to parse_csv / parse_xlsx / parse_pdf based on extension."""
        filepath = Path(filepath)
        ext = filepath.suffix.lower()

        try:
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
        except (ValueError, KeyError) as e:
            # Try to detect if this is a non-QB accounting system
            detected_dialect = self._detect_file_dialect(filepath, ext)
            if detected_dialect:
                raise ValueError(
                    f"This file appears to be from {detected_dialect.name}. "
                    f"Currently only QuickBooks Online formats are supported."
                ) from e
            # Otherwise, re-raise original error
            raise

    def _detect_file_dialect(self, filepath: Path, ext: str):
        """Extract headers from file and detect dialect."""
        from dialects.registry import DialectRegistry
        
        headers = []
        try:
            if ext == '.csv':
                with open(filepath, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    for _ in range(5):  # Check first 5 rows for headers
                        row = next(reader, None)
                        if row:
                            headers.extend(row)
            elif ext == '.xlsx':
                if not XLSX_SUPPORT:
                    return None
                workbook = openpyxl.load_workbook(filepath)
                sheet = workbook.active
                for row in sheet.iter_rows(max_row=5, values_only=True):
                    headers.extend([str(cell) for cell in row if cell])
        except:
            return None
        
        return DialectRegistry.detect(headers)

    def convert_to_json(self, filepath: Path, output_path: Optional[Path] = None) -> str:
        """Convert file to JSON, optionally writing to disk."""
        result = self.convert_file(filepath)

        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, indent=2)
            return f"Converted to {output_path}"
        else:
            return json.dumps(result, indent=2)

    # ──────────────────────────────────────────────
    # Abstract parse methods
    # ──────────────────────────────────────────────

    @abstractmethod
    def parse_csv(self, filepath: Path):
        """Parse a CSV file. Return type depends on the converter."""
        ...

    @abstractmethod
    def parse_xlsx(self, filepath: Path):
        """Parse an XLSX file. Return type depends on the converter."""
        ...

    @abstractmethod
    def parse_pdf(self, filepath: Path):
        """Parse a PDF file. Return type depends on the converter."""
        ...

    # ──────────────────────────────────────────────
    # Import guards
    # ──────────────────────────────────────────────

    @staticmethod
    def check_xlsx_support():
        if not XLSX_SUPPORT:
            raise ImportError("openpyxl is required for XLSX support. Install with: pip install openpyxl")

    @staticmethod
    def check_pdf_support():
        if not PDF_SUPPORT:
            raise ImportError("pdfplumber is required for PDF support. Install with: pip install pdfplumber")

    # ──────────────────────────────────────────────
    # Account ID management
    # ──────────────────────────────────────────────

    def get_or_create_account_id(self, account_name: str) -> str:
        """
        Get a stable account ID — from local map, API lookup, or generated.

        IDs are cached in self.account_id_map for consistency within a conversion.
        """
        if account_name in self.account_id_map:
            return self.account_id_map[account_name]

        # Try API lookup
        if self.use_account_lookup and self.account_lookup_client:
            account_id = self.account_lookup_client.lookup_account_id(account_name)
            if account_id:
                self.account_id_map[account_name] = account_id
                return account_id

        # Fallback: generate sequential ID
        account_id = str(self.account_id_counter)
        self.account_id_counter += 1
        self.account_id_map[account_name] = account_id
        return account_id

    def generate_account_id(self) -> str:
        """Generate a sequential account ID (no caching)."""
        id_str = str(self.account_id_counter)
        self.account_id_counter += 1
        return id_str

    # ──────────────────────────────────────────────
    # QB report envelope helpers
    # ──────────────────────────────────────────────

    @staticmethod
    def make_coldata_cell(value, id=None) -> Dict[str, Any]:
        """Build a single colData cell dict."""
        return {
            "attributes": None,
            "value": value if value is not None else "",
            "id": id,
            "href": None
        }

    @staticmethod
    def make_qb_header(report_name: str, start_period: str = None, end_period: str = None,
                       summarize_columns_by: str = "Total", currency: str = "USD",
                       options: List[Dict] = None) -> Dict[str, Any]:
        """Build a QB-style report header dict."""
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000+00:00')
        return {
            "time": timestamp,
            "reportName": report_name,
            "dateMacro": "today",
            "reportBasis": None,
            "startPeriod": start_period,
            "endPeriod": end_period,
            "summarizeColumnsBy": summarize_columns_by,
            "currency": currency,
            "customer": None,
            "vendor": None,
            "employee": None,
            "item": None,
            "clazz": None,
            "department": None,
            "option": options or []
        }

    # ──────────────────────────────────────────────
    # Date / month parsing utilities
    # ──────────────────────────────────────────────

    @staticmethod
    def parse_month_year(text: str) -> Tuple[str, str, date, date]:
        """
        Parse month and year from text like "January 2025" or "Jan 2025".

        Returns: (month_abbr_upper, year_str, start_date, end_date)
                 e.g. ("JAN", "2025", date(2025,1,1), date(2025,1,31))
        """
        match = re.search(r'(\w+)\s+(\d{4})', text)
        if match:
            month_name = match.group(1)
            year = match.group(2)

            try:
                month_num = datetime.strptime(month_name, '%B').month
            except ValueError:
                try:
                    month_num = datetime.strptime(month_name[:3], '%b').month
                except ValueError:
                    month_num = 1

            start_date = date(int(year), month_num, 1)
            last_day = calendar.monthrange(int(year), month_num)[1]
            end_date = date(int(year), month_num, last_day)

            return month_name.upper()[:3], year, start_date, end_date

        return "JAN", "2025", date(2025, 1, 1), date(2025, 1, 31)

    @staticmethod
    def parse_month_column(column_header: str) -> Tuple[str, date, date]:
        """
        Parse a month column header to ISO month key + date range.

        Handles: 
        - "January 2025" (full month name)
        - "Jan 2025" (abbreviated month name)
        - "Jul 1 - Jul 27 2025" (date range)
        
        Returns: (iso_month "2025-01", start_date, end_date)
        """
        if ' - ' in column_header:
            parts = column_header.split(' - ')
            end_part = parts[1].strip()
            match = re.search(r'(\w+)\s+\d+\s+(\d{4})', end_part)
            if match:
                month_name = match.group(1)
                year = int(match.group(2))
                month_num = datetime.strptime(month_name[:3], '%b').month

                start_match = re.search(r'(\w+)\s+(\d+)', parts[0])
                start_day = int(start_match.group(2)) if start_match else 1
                start_date = date(year, month_num, start_day)

                end_match = re.search(r'(\w+)\s+(\d+)\s+(\d{4})', parts[1])
                end_day = int(end_match.group(2)) if end_match else calendar.monthrange(year, month_num)[1]
                end_date = date(year, month_num, end_day)

                return f"{year}-{month_num:02d}", start_date, end_date
        else:
            match = re.search(r'(\w+)\s+(\d{4})', column_header)
            if match:
                month_name = match.group(1)
                year = int(match.group(2))
                
                # Try full month name first (%B = January, February, ...)
                try:
                    month_num = datetime.strptime(month_name, '%B').month
                except ValueError:
                    # Fall back to abbreviated month name (%b = Jan, Feb, ...)
                    try:
                        month_num = datetime.strptime(month_name[:3], '%b').month
                    except ValueError:
                        # If both fail, default to January
                        month_num = 1
                
                start_date = date(year, month_num, 1)
                last_day = calendar.monthrange(year, month_num)[1]
                end_date = date(year, month_num, last_day)
                return f"{year}-{month_num:02d}", start_date, end_date

        return "2025-01", date(2025, 1, 1), date(2025, 1, 31)

    @staticmethod
    def parse_date_range(header_text: str) -> Optional[Tuple[str, date, date]]:
        """
        Parse date range from header text. Supports multiple formats:
        - "April 1-30, 2024"
        - "January 1 - September 8, 2025"
        - "01/01/2024 - 01/31/2024"
        - "January 2024"

        Returns: (period_string, start_date, end_date) or None
        """
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
                    return period, start_date, end_date
                except ValueError:
                    pass

        # Pattern 2: "January 1 - September 8, 2025" (cross-month range)
        match = re.search(
            r'(\w+)\s+(\d+)\s*-\s*(\w+)\s+(\d+),?\s*(\d{4})',
            header_text, re.IGNORECASE
        )
        if match:
            start_month_name = match.group(1)
            start_day = int(match.group(2))
            end_month_name = match.group(3)
            end_day = int(match.group(4))
            year = int(match.group(5))
            start_month = months_full.get(start_month_name.capitalize()) or months_abbr.get(start_month_name.capitalize())
            end_month = months_full.get(end_month_name.capitalize()) or months_abbr.get(end_month_name.capitalize())
            if start_month and end_month:
                try:
                    start_date = date(year, start_month, start_day)
                    end_date = date(year, end_month, end_day)
                    period = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
                    return period, start_date, end_date
                except ValueError:
                    pass

        # Pattern 3: "01/01/2024 - 01/31/2024" (numeric dates)
        match = re.search(r'(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', header_text)
        if match:
            try:
                start_date = datetime.strptime(match.group(1), '%m/%d/%Y').date()
                end_date = datetime.strptime(match.group(2), '%m/%d/%Y').date()
                period = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
                return period, start_date, end_date
            except ValueError:
                pass

        # Pattern 4: "January 2024" (full month)
        match = re.search(r'(\w+)\s+(\d{4})', header_text)
        if match:
            month_name = match.group(1)
            year = int(match.group(2))
            month_num = months_full.get(month_name.capitalize()) or months_abbr.get(month_name.capitalize())
            if month_num:
                start_date = date(year, month_num, 1)
                last_day = calendar.monthrange(year, month_num)[1]
                end_date = date(year, month_num, last_day)
                period = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
                return period, start_date, end_date

        return None
