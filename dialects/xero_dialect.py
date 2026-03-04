"""
Xero Dialect — stub for future implementation.

When Xero export samples are available, implement detect() and map_row()
to normalize Xero column headers to canonical field names.
"""

from typing import Dict, List, Optional, Any
from dialects.base_dialect import BaseDialect
from dialects.registry import DialectRegistry


@DialectRegistry.register
class XeroDialect(BaseDialect):
    """Dialect for Xero accounting exports."""

    name = "Xero"

    # Xero-specific header patterns (to be refined with real samples)
    _XERO_MARKERS = {"account code", "account name", "ytd"}

    @classmethod
    def detect(cls, headers: List[str]) -> bool:
        lower = set(headers)
        # Require at least 2 Xero-specific markers
        return len(cls._XERO_MARKERS.intersection(lower)) >= 2

    def map_row(self, raw_row: Dict[str, Any], headers: List[str]) -> Dict[str, Any]:
        raise NotImplementedError("Xero dialect not yet implemented — need sample files")

    def get_report_type_hint(self, headers: List[str]) -> Optional[str]:
        return None
