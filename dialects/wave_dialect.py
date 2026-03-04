"""
Wave Dialect — stub for future implementation.
"""

from typing import Dict, List, Optional, Any
from dialects.base_dialect import BaseDialect
from dialects.registry import DialectRegistry


@DialectRegistry.register
class WaveDialect(BaseDialect):
    name = "Wave"

    _WAVE_MARKERS = {"account number", "account name", "net movement"}

    @classmethod
    def detect(cls, headers: List[str]) -> bool:
        return len(cls._WAVE_MARKERS.intersection(set(headers))) >= 2

    def map_row(self, raw_row: Dict[str, Any], headers: List[str]) -> Dict[str, Any]:
        raise NotImplementedError("Wave dialect not yet implemented — need sample files")

    def get_report_type_hint(self, headers: List[str]) -> Optional[str]:
        return None
