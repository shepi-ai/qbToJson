"""
QB Envelope Renderer — converts canonical data into QuickBooks-style
report envelopes with {header, columns, rows} structure.
"""

from datetime import datetime, timezone, date
from typing import Dict, List, Any, Optional
from canonical.schemas import (
    CanonicalTrialBalanceLine,
    CanonicalBalanceLine,
    CanonicalTransaction,
    CanonicalAgingLine,
)


class QBEnvelopeRenderer:
    """Builds QuickBooks-style JSON report envelopes from canonical data."""

    # ──────────────────────────────────────────────
    # Cell / row helpers
    # ──────────────────────────────────────────────

    @staticmethod
    def cell(value, id=None) -> Dict[str, Any]:
        return {
            "attributes": None,
            "value": value if value is not None else "",
            "id": id,
            "href": None
        }

    @staticmethod
    def header(report_name: str, start_period: str = None, end_period: str = None,
               options: List[Dict] = None) -> Dict[str, Any]:
        ts = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000+00:00')
        return {
            "time": ts,
            "reportName": report_name,
            "dateMacro": "today",
            "reportBasis": None,
            "startPeriod": start_period,
            "endPeriod": end_period,
            "summarizeColumnsBy": "Total",
            "currency": "USD",
            "customer": None,
            "vendor": None,
            "employee": None,
            "item": None,
            "clazz": None,
            "department": None,
            "option": options or []
        }

    # ──────────────────────────────────────────────
    # Trial Balance
    # ──────────────────────────────────────────────

    @classmethod
    def render_trial_balance(cls, lines: List[CanonicalTrialBalanceLine],
                             month_abbr: str, year: str,
                             start_date: date, end_date: date) -> Dict[str, Any]:
        """Render TB lines into a QB report envelope for a single month."""
        rows = []
        for line in lines:
            debit_str = f"{line.debit:.2f}" if line.debit else ""
            credit_str = f"{line.credit:.2f}" if line.credit else ""
            rows.append({
                "colData": [
                    cls.cell(line.account_name, line.account_id),
                    cls.cell(debit_str),
                    cls.cell(credit_str),
                ]
            })

        return {
            "header": cls.header(
                "TrialBalance",
                start_period=start_date.isoformat(),
                end_period=end_date.isoformat(),
                options=[
                    {"name": "accounting_method", "value": "Accrual"},
                    {"name": "NoReportData", "value": "false"}
                ]
            ),
            "columns": {
                "column": [
                    {"colTitle": "", "colType": "Account", "metaData": [], "columns": None},
                    {"colTitle": "Debit", "colType": "Money", "metaData": [], "columns": None},
                    {"colTitle": "Credit", "colType": "Money", "metaData": [], "columns": None},
                ]
            },
            "rows": {"row": rows}
        }

    # ──────────────────────────────────────────────
    # Balance Sheet / P&L / Cash Flow
    # ──────────────────────────────────────────────

    @classmethod
    def render_financial_statement(cls, lines: List[CanonicalBalanceLine],
                                   report_name: str,
                                   start_date: date, end_date: date) -> Dict[str, Any]:
        """Render BS/PL/CF lines into a QB report with SECTION/DATA structure."""
        # Group by section
        sections: Dict[str, List] = {}
        current_section = None

        for line in lines:
            if line.is_section_header:
                current_section = line.name
                sections[current_section] = []
            elif current_section:
                sections[current_section].append(line)
            else:
                sections.setdefault("_default", []).append(line)

        rows = []
        for section_name, section_lines in sections.items():
            data_rows = []
            for sl in section_lines:
                value_str = f"{sl.value:.2f}" if sl.value else ""
                data_rows.append({
                    "colData": [
                        cls.cell(sl.name, sl.account_id),
                        cls.cell(value_str),
                    ],
                    "type": "DATA"
                })

            if section_name != "_default":
                rows.append({
                    "type": "SECTION",
                    "header": {"colData": [cls.cell(section_name)]},
                    "rows": {"row": data_rows},
                    "summary": None
                })
            else:
                rows.extend(data_rows)

        return {
            "header": cls.header(report_name,
                                 start_period=start_date.isoformat(),
                                 end_period=end_date.isoformat()),
            "columns": {
                "column": [
                    {"colTitle": "", "colType": "Account", "metaData": [], "columns": None},
                    {"colTitle": "Total", "colType": "Money", "metaData": [], "columns": None},
                ]
            },
            "rows": {"row": rows}
        }

    # ──────────────────────────────────────────────
    # Aging (AR / AP)
    # ──────────────────────────────────────────────

    @classmethod
    def render_aging(cls, lines: List[CanonicalAgingLine],
                     report_name: str, entity_type: str = "Customer",
                     report_date: str = None) -> Dict[str, Any]:
        """Render aging lines into a QB AgedReceivables/AgedPayables envelope."""
        if report_date is None:
            report_date = datetime.now(timezone.utc).strftime('%Y-%m-%d')

        rows = []
        for line in lines:
            rows.append({
                "colData": [
                    cls.cell(line.name, line.entity_id),
                    cls.cell(f"{line.current:.2f}" if line.current else ""),
                    cls.cell(f"{line.days_1_30:.2f}" if line.days_1_30 else ""),
                    cls.cell(f"{line.days_31_60:.2f}" if line.days_31_60 else ""),
                    cls.cell(f"{line.days_61_90:.2f}" if line.days_61_90 else ""),
                    cls.cell(f"{line.days_91_plus:.2f}" if line.days_91_plus else ""),
                    cls.cell(f"{line.total:.2f}" if line.total else ""),
                ]
            })

        return {
            "header": cls.header(report_name,
                                 start_period=report_date,
                                 end_period=report_date,
                                 options=[
                                     {"name": "report_date", "value": report_date},
                                     {"name": "NoReportData", "value": "false"}
                                 ]),
            "columns": {
                "column": [
                    {"colTitle": "", "colType": entity_type, "metaData": [], "columns": None},
                    {"colTitle": "Current", "colType": "Money",
                     "metaData": [{"name": "ColKey", "value": "current"}], "columns": None},
                    {"colTitle": "1 - 30", "colType": "Money",
                     "metaData": [{"name": "ColKey", "value": "0"}], "columns": None},
                    {"colTitle": "31 - 60", "colType": "Money",
                     "metaData": [{"name": "ColKey", "value": "1"}], "columns": None},
                    {"colTitle": "61 - 90", "colType": "Money",
                     "metaData": [{"name": "ColKey", "value": "2"}], "columns": None},
                    {"colTitle": "91 and over", "colType": "Money",
                     "metaData": [{"name": "ColKey", "value": "3"}], "columns": None},
                    {"colTitle": "Total", "colType": "Money",
                     "metaData": [{"name": "ColKey", "value": "total"}], "columns": None},
                ]
            },
            "rows": {"row": rows}
        }

    # ──────────────────────────────────────────────
    # General Ledger
    # ──────────────────────────────────────────────

    @classmethod
    def render_general_ledger(cls, transactions: List[CanonicalTransaction],
                              start_date: date = None, end_date: date = None) -> Dict[str, Any]:
        """Render GL transactions into a QB report envelope."""
        rows = []
        for txn in transactions:
            debit_str = f"{txn.debit:.2f}" if txn.debit else ""
            credit_str = f"{txn.credit:.2f}" if txn.credit else ""
            balance_str = f"{txn.balance:.2f}" if txn.balance else ""
            rows.append({
                "colData": [
                    cls.cell(txn.date),
                    cls.cell(txn.transaction_type),
                    cls.cell(txn.num),
                    cls.cell(txn.name),
                    cls.cell(txn.memo),
                    cls.cell(txn.split),
                    cls.cell(debit_str),
                    cls.cell(credit_str),
                    cls.cell(balance_str),
                ]
            })

        start_str = start_date.isoformat() if start_date else None
        end_str = end_date.isoformat() if end_date else None

        return {
            "header": cls.header("GeneralLedger",
                                 start_period=start_str,
                                 end_period=end_str),
            "columns": {
                "column": [
                    {"colTitle": "Date", "colType": "Date", "metaData": [], "columns": None},
                    {"colTitle": "Transaction Type", "colType": "String", "metaData": [], "columns": None},
                    {"colTitle": "Num", "colType": "String", "metaData": [], "columns": None},
                    {"colTitle": "Name", "colType": "String", "metaData": [], "columns": None},
                    {"colTitle": "Memo", "colType": "String", "metaData": [], "columns": None},
                    {"colTitle": "Split", "colType": "String", "metaData": [], "columns": None},
                    {"colTitle": "Debit", "colType": "Money", "metaData": [], "columns": None},
                    {"colTitle": "Credit", "colType": "Money", "metaData": [], "columns": None},
                    {"colTitle": "Balance", "colType": "Money", "metaData": [], "columns": None},
                ]
            },
            "rows": {"row": rows}
        }
