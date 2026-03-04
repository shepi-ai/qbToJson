"""
Canonical Schemas — source-agnostic dataclasses for all report types.

These are the intermediate representation between raw file data (from any
accounting system) and the final JSON output (QB envelope, flat array, etc.).

Every converter maps parsed rows into one of these dataclasses.  Renderers
then convert them to the desired output format.
"""

from dataclasses import dataclass, field
from typing import List, Optional


# ──────────────────────────────────────────────
# Chart of Accounts
# ──────────────────────────────────────────────

@dataclass
class CanonicalAccount:
    """A single account from the Chart of Accounts."""
    name: str
    classification: str          # ASSET, LIABILITY, EQUITY, REVENUE, EXPENSE
    account_type: str            # BANK, ACCOUNTS_RECEIVABLE, INCOME, EXPENSE, etc.
    account_sub_type: str = ""
    account_id: Optional[str] = None
    description: Optional[str] = None
    balance: float = 0.0
    active: bool = True


# ──────────────────────────────────────────────
# Trial Balance
# ──────────────────────────────────────────────

@dataclass
class CanonicalTrialBalanceLine:
    """A single row in a trial balance report."""
    account_name: str
    debit: float = 0.0
    credit: float = 0.0
    account_id: Optional[str] = None


# ──────────────────────────────────────────────
# Balance Sheet / Profit & Loss / Cash Flow
# ──────────────────────────────────────────────

@dataclass
class CanonicalBalanceLine:
    """A single line item in BS, P&L, or Cash Flow statements."""
    name: str
    value: float = 0.0
    account_id: Optional[str] = None
    section: Optional[str] = None       # e.g. "ASSETS", "Revenue", "Operating Activities"
    is_section_header: bool = False
    is_total: bool = False
    depth: int = 0                      # Indentation level for hierarchy


# ──────────────────────────────────────────────
# General Ledger
# ──────────────────────────────────────────────

@dataclass
class CanonicalTransaction:
    """A single transaction line from the General Ledger."""
    date: str
    transaction_type: str = ""
    num: str = ""
    name: str = ""
    memo: str = ""
    debit: float = 0.0
    credit: float = 0.0
    balance: float = 0.0
    account_name: str = ""
    account_id: Optional[str] = None
    split: str = ""


# ──────────────────────────────────────────────
# AR / AP Aging
# ──────────────────────────────────────────────

@dataclass
class CanonicalAgingLine:
    """A single row in an aging report (AR or AP)."""
    name: str                           # Customer or Vendor name
    current: float = 0.0
    days_1_30: float = 0.0
    days_31_60: float = 0.0
    days_61_90: float = 0.0
    days_91_plus: float = 0.0
    total: float = 0.0
    entity_id: Optional[str] = None


# ──────────────────────────────────────────────
# Customer / Vendor Concentration
# ──────────────────────────────────────────────

@dataclass
class CanonicalConcentrationLine:
    """A single customer or vendor concentration row."""
    name: str
    amount: float = 0.0
    percentage: float = 0.0


# ──────────────────────────────────────────────
# Journal Entries
# ──────────────────────────────────────────────

@dataclass
class CanonicalJournalLine:
    """A single line within a journal entry."""
    account: str
    description: str = ""
    debit: float = 0.0
    credit: float = 0.0


@dataclass
class CanonicalJournalEntry:
    """A complete journal entry with header + line items."""
    entry_id: str = ""
    date: str = ""
    entry_type: str = ""
    num: str = ""
    name: str = ""
    memo: str = ""
    lines: List[CanonicalJournalLine] = field(default_factory=list)


# ──────────────────────────────────────────────
# Fixed Assets (new)
# ──────────────────────────────────────────────

@dataclass
class CanonicalFixedAsset:
    """A single fixed asset from the Fixed Asset Register."""
    asset_name: str
    asset_id: Optional[str] = None
    category: str = ""                  # e.g. "Furniture", "Equipment", "Vehicles"
    acquisition_date: str = ""
    acquisition_cost: float = 0.0
    accumulated_depreciation: float = 0.0
    net_book_value: float = 0.0
    useful_life_years: Optional[float] = None
    depreciation_method: str = ""       # e.g. "Straight-line", "Declining balance"
    disposal_date: Optional[str] = None
    disposal_amount: float = 0.0
    status: str = "Active"              # Active, Disposed, Fully Depreciated


# ──────────────────────────────────────────────
# Depreciation Schedule (new)
# ──────────────────────────────────────────────

@dataclass
class CanonicalDepreciationScheduleLine:
    """A single line in a depreciation schedule."""
    asset_name: str
    asset_id: Optional[str] = None
    period: str = ""                    # e.g. "2025-01" or "2025"
    beginning_balance: float = 0.0
    depreciation_expense: float = 0.0
    ending_balance: float = 0.0
    category: str = ""
    depreciation_method: str = ""
