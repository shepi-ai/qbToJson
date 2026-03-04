"""
Flat Array Renderer — converts canonical data into simple List[Dict] output.

Used by converters that don't need the QB envelope structure:
- Customer/Vendor Concentration
- Journal Entries
- Fixed Asset Register
- Depreciation Schedule
"""

from typing import Dict, List, Any
from canonical.schemas import (
    CanonicalAccount,
    CanonicalConcentrationLine,
    CanonicalJournalEntry,
    CanonicalFixedAsset,
    CanonicalDepreciationScheduleLine,
)
from datetime import datetime, timezone


class FlatArrayRenderer:
    """Builds simple List[Dict] output from canonical data."""

    # ──────────────────────────────────────────────
    # Customer / Vendor Concentration
    # ──────────────────────────────────────────────

    @staticmethod
    def render_customer_concentration(lines: List[CanonicalConcentrationLine]) -> List[Dict[str, Any]]:
        return [
            {
                "customerName": line.name,
                "revenue": line.amount,
                "percentage": line.percentage,
            }
            for line in lines
        ]

    @staticmethod
    def render_vendor_concentration(lines: List[CanonicalConcentrationLine]) -> List[Dict[str, Any]]:
        return [
            {
                "vendorName": line.name,
                "payments": line.amount,
                "percentage": line.percentage,
            }
            for line in lines
        ]

    # ──────────────────────────────────────────────
    # Chart of Accounts
    # ──────────────────────────────────────────────

    @staticmethod
    def render_accounts(accounts: List[CanonicalAccount]) -> List[Dict[str, Any]]:
        """Render accounts into QB-style account objects."""
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000+00:00')
        result = []
        for acct in accounts:
            # Clean sub type for QB format
            sub_type = acct.account_sub_type.strip()
            sub_type = sub_type.replace('/', '').replace('&', 'And').replace(' ', '')

            result.append({
                "id": acct.account_id,
                "syncToken": "0",
                "metaData": {
                    "createdByRef": None,
                    "createTime": timestamp,
                    "lastModifiedByRef": None,
                    "lastUpdatedTime": timestamp,
                    "lastChangedInQB": None,
                    "synchronized": None
                },
                "customField": [],
                "attachableRef": [],
                "domain": "QBO",
                "status": None,
                "sparse": False,
                "name": acct.name,
                "subAccount": False,
                "parentRef": None,
                "description": acct.description,
                "fullyQualifiedName": acct.name,
                "accountAlias": None,
                "txnLocationType": None,
                "active": acct.active,
                "classification": acct.classification,
                "accountType": acct.account_type,
                "accountSubType": sub_type,
                "accountPurposes": [],
                "acctNum": None,
                "acctNumExtn": None,
                "bankNum": None,
                "openingBalance": None,
                "openingBalanceDate": None,
                "currentBalance": acct.balance,
                "currentBalanceWithSubAccounts": acct.balance,
                "currencyRef": {
                    "value": "USD",
                    "name": "United States Dollar",
                    "type": None
                },
                "taxAccount": None,
                "taxCodeRef": None,
                "onlineBankingEnabled": None,
                "journalCodeRef": None,
                "accountEx": None,
                "finame": None
            })
        return result

    # ──────────────────────────────────────────────
    # Journal Entries
    # ──────────────────────────────────────────────

    @staticmethod
    def render_journal_entries(entries: List[CanonicalJournalEntry]) -> List[Dict[str, Any]]:
        result = []
        for entry in entries:
            lines = []
            for jl in entry.lines:
                lines.append({
                    "account": jl.account,
                    "description": jl.description,
                    "debit": jl.debit,
                    "credit": jl.credit,
                })
            result.append({
                "id": entry.entry_id,
                "date": entry.date,
                "type": entry.entry_type,
                "num": entry.num,
                "name": entry.name,
                "memo": entry.memo,
                "lines": lines,
            })
        return result

    # ──────────────────────────────────────────────
    # Fixed Asset Register (new)
    # ──────────────────────────────────────────────

    @staticmethod
    def render_fixed_assets(assets: List[CanonicalFixedAsset]) -> List[Dict[str, Any]]:
        return [
            {
                "assetName": a.asset_name,
                "assetId": a.asset_id,
                "category": a.category,
                "acquisitionDate": a.acquisition_date,
                "acquisitionCost": a.acquisition_cost,
                "accumulatedDepreciation": a.accumulated_depreciation,
                "netBookValue": a.net_book_value,
                "usefulLifeYears": a.useful_life_years,
                "depreciationMethod": a.depreciation_method,
                "disposalDate": a.disposal_date,
                "disposalAmount": a.disposal_amount,
                "status": a.status,
            }
            for a in assets
        ]

    # ──────────────────────────────────────────────
    # Depreciation Schedule (new)
    # ──────────────────────────────────────────────

    @staticmethod
    def render_depreciation_schedule(lines: List[CanonicalDepreciationScheduleLine]) -> List[Dict[str, Any]]:
        return [
            {
                "assetName": line.asset_name,
                "assetId": line.asset_id,
                "period": line.period,
                "beginningBalance": line.beginning_balance,
                "depreciationExpense": line.depreciation_expense,
                "endingBalance": line.ending_balance,
                "category": line.category,
                "depreciationMethod": line.depreciation_method,
            }
            for line in lines
        ]
