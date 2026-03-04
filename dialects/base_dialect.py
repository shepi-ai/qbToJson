"""
Base Dialect — abstract interface for mapping raw column headers and rows
from a specific accounting system to standardized field names.
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Any


class BaseDialect(ABC):
    """
    A dialect knows how to:
    1. Detect whether a set of headers belongs to its accounting system
    2. Map raw rows (keyed by original header names) to standard field names
    """

    # Human-readable name for this dialect (e.g. "QuickBooks", "Xero")
    name: str = "Unknown"

    @classmethod
    @abstractmethod
    def detect(cls, headers: List[str]) -> bool:
        """
        Return True if *headers* (lowercased column names from a file)
        match this accounting system's expected format.
        """
        ...

    @abstractmethod
    def map_row(self, raw_row: Dict[str, Any], headers: List[str]) -> Dict[str, Any]:
        """
        Normalize a single data row from original column names to
        standard field names used by the canonical schemas.

        Args:
            raw_row:  dict keyed by original header names (or column indices)
            headers:  the original header list (for positional lookups)

        Returns:
            dict with standardized keys (e.g. "account_name", "debit", "credit")
        """
        ...

    @abstractmethod
    def get_report_type_hint(self, headers: List[str]) -> Optional[str]:
        """
        Return a hint about what report type these headers represent.
        e.g. "trial_balance", "balance_sheet", "aging_ar"
        Returns None if the dialect can't determine the report type.
        """
        ...
