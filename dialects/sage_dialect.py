"""
Sage Dialect — stub for future implementation.
"""

from typing import Dict, List, Optional, Any
from dialects.base_dialect import BaseDialect
from dialects.registry import DialectRegistry


@DialectRegistry.register
class SageDialect(BaseDialect):
    name = "Sage"

    _SAGE_MARKERS = {"nominal code", "nominal name", "period"}

    @classmethod
    def detect(cls, headers: List[str]) -> bool:
        return len(cls._SAGE_MARKERS.intersection(set(headers))) >= 2

    def map_row(self, raw_row: Dict[str, Any], headers: List[str]) -> Dict[str, Any]:
        raise NotImplementedError("Sage dialect not yet implemented — need sample files")

    def get_report_type_hint(self, headers: List[str]) -> Optional[str]:
        return None
