"""
QuickBooks Dialect — maps QuickBooks Online report column headers
to standardized field names.
"""

from typing import Dict, List, Optional, Any
from dialects.base_dialect import BaseDialect
from dialects.registry import DialectRegistry


@DialectRegistry.register
class QuickBooksDialect(BaseDialect):
    """Dialect for QuickBooks Online exported reports."""

    name = "QuickBooks"

    # QB header keywords per report type
    _REPORT_SIGNATURES = {
        "trial_balance": ["debit", "credit"],
        "balance_sheet": ["total"],
        "profit_loss": ["total"],
        "cash_flow": ["total"],
        "general_ledger": ["date", "transaction type", "num", "name", "memo"],
        "aging_ar": ["current", "1 - 30", "31 - 60", "61 - 90", "91 and over"],
        "aging_ap": ["current", "1 - 30", "31 - 60", "61 - 90", "91 and over"],
        "accounts": ["full name", "type", "detail type"],
        "customer_concentration": ["customer", "total"],
        "vendor_concentration": ["vendor", "total"],
        "journal_entries": ["transaction date", "transaction type"],
    }

    # Standard QB column name → canonical field name
    _COLUMN_MAP = {
        # Trial Balance
        "debit": "debit",
        "credit": "credit",

        # Accounts
        "full name": "account_name",
        "name": "account_name",
        "type": "account_type",
        "detail type": "account_sub_type",
        "description": "description",
        "total balance": "balance",
        "balance": "balance",

        # Aging
        "customer": "name",
        "vendor": "name",
        "current": "current",
        "1 - 30": "days_1_30",
        "31 - 60": "days_31_60",
        "61 - 90": "days_61_90",
        "91 and over": "days_91_plus",
        "total": "total",

        # General Ledger
        "date": "date",
        "transaction type": "transaction_type",
        "num": "num",
        "memo": "memo",
        "split": "split",

        # Journal Entries
        "transaction date": "date",
        "account": "account",
    }

    @classmethod
    def detect(cls, headers: List[str]) -> bool:
        """
        QuickBooks is our default dialect — returns True for any header set
        that contains at least one known QB column name.
        """
        known = set(cls._COLUMN_MAP.keys())
        return any(h in known for h in headers)

    def map_row(self, raw_row: Dict[str, Any], headers: List[str]) -> Dict[str, Any]:
        """Map a QB row to standardized field names."""
        result = {}
        for original_key, value in raw_row.items():
            normalized_key = self._COLUMN_MAP.get(str(original_key).lower().strip())
            if normalized_key:
                result[normalized_key] = value
            else:
                # Pass through unmapped columns with original key
                result[original_key] = value
        return result

    def get_report_type_hint(self, headers: List[str]) -> Optional[str]:
        """Determine report type from QB headers."""
        lower_headers = set(h.lower().strip() for h in headers if h)

        # Check for aging reports first (most specific headers)
        aging_cols = {"current", "1 - 30", "31 - 60"}
        if aging_cols.issubset(lower_headers):
            if "customer" in lower_headers:
                return "aging_ar"
            if "vendor" in lower_headers:
                return "aging_ap"
            return "aging"

        # Accounts / CoA
        if "full name" in lower_headers and "type" in lower_headers:
            return "accounts"

        # GL
        if "transaction type" in lower_headers and "memo" in lower_headers:
            if "transaction date" in lower_headers:
                return "journal_entries"
            return "general_ledger"

        # Concentration
        if "customer" in lower_headers and "total" in lower_headers:
            return "customer_concentration"
        if "vendor" in lower_headers and "total" in lower_headers:
            return "vendor_concentration"

        # Trial Balance (has debit + credit)
        if "debit" in lower_headers and "credit" in lower_headers:
            return "trial_balance"

        return None
