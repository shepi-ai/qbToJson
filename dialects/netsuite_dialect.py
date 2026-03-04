"""
NetSuite Dialect — stub for future implementation.
"""

from typing import Dict, List, Optional, Any
from dialects.base_dialect import BaseDialect
from dialects.registry import DialectRegistry


@DialectRegistry.register
class NetSuiteDialect(BaseDialect):
    name = "NetSuite"

    _NETSUITE_MARKERS = {"internal id", "account", "debit amount", "credit amount"}

    @classmethod
    def detect(cls, headers: List[str]) -> bool:
        return len(cls._NETSUITE_MARKERS.intersection(set(headers))) >= 2

    def map_row(self, raw_row: Dict[str, Any], headers: List[str]) -> Dict[str, Any]:
        raise NotImplementedError("NetSuite dialect not yet implemented — need sample files")

    def get_report_type_hint(self, headers: List[str]) -> Optional[str]:
        return None
