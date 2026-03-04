"""
Dialect Registry — auto-detect the accounting system from file headers.
"""

from typing import List, Optional, Type
from dialects.base_dialect import BaseDialect


class DialectRegistry:
    """Registry of all known dialects.  Call detect() with headers to find the right one."""

    _dialects: List[Type[BaseDialect]] = []

    @classmethod
    def register(cls, dialect_cls: Type[BaseDialect]):
        """Register a dialect class."""
        if dialect_cls not in cls._dialects:
            cls._dialects.append(dialect_cls)
        return dialect_cls

    @classmethod
    def detect(cls, headers: List[str]) -> Optional[BaseDialect]:
        """
        Try each registered dialect's detect() classmethod against *headers*.
        Returns an instance of the first matching dialect, or None.
        """
        lower_headers = [str(h).lower().strip() for h in headers if h]
        for dialect_cls in cls._dialects:
            if dialect_cls.detect(lower_headers):
                return dialect_cls()
        return None

    @classmethod
    def get_all(cls) -> List[Type[BaseDialect]]:
        """Return all registered dialect classes."""
        return list(cls._dialects)
