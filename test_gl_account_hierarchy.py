#!/usr/bin/env python3
"""
Regression tests for General Ledger sections whose leaf account names collide.

QuickBooks prints only an account's leaf name as a GL section header, so a
revenue account "Landscaping Services:Job Materials" and an expense account
"Job Expenses:Job Materials" both render as a section called "Job Materials".
The parser keyed sections by that header, so the later section silently
overwrote the earlier one and took every one of its rows with it.

On the Sandbox Company accrual GL that dropped 655 of 8,346 rows across 7
accounts -- the entire income side of Job Materials, Plants and Soil,
Sprinklers and Drip Systems, Decks and Patios, Installation, Maintenance and
Repair and Equipment Rental -- which left ~$914k of customer deposits with no
reciprocal ledger row to reconcile against.

The export does carry the hierarchy: an account closes its own rows with
"Total for X" and, when it has children, closes its subtree with
"Total for X with sub-accounts".
"""
import tempfile
import csv
from pathlib import Path

import openpyxl
import pytest

from generalLedgerConverter import GeneralLedgerConverter


HEADER = ['', 'Distribution account', 'Transaction date', 'Transaction type',
          'Num', 'Name', 'Memo/Description', 'Split account', 'Amount', 'Balance']

# Mirrors the real Sandbox Company structure: a revenue tree and an expense
# tree that share the leaf names "Job Materials" and "Plants and Soil", plus a
# top-level expense account whose leaf collides with a nested revenue account.
ROWS = [
    ['Landscaping Services'],
    ['', 'Landscaping Services', '01/05/2023', 'Invoice', '1', 'Cust A', '', 'Accounts Receivable (A/R)', '500.00', '500.00'],
    ['Total for Landscaping Services', '', '', '', '', '', '', '', '500.00'],

    ['Job Materials'],
    ['', 'Job Materials', '01/14/2023', 'Deposit', '', '', '', 'Checking', '2293.04', '2293.04'],
    ['Total for Job Materials', '', '', '', '', '', '', '', '2293.04'],

    ['Plants and Soil'],
    ['', 'Plants and Soil', '05/12/2023', 'Deposit', '', '', '', 'Savings', '7737.56', '7737.56'],
    ['Total for Plants and Soil', '', '', '', '', '', '', '', '7737.56'],

    ['Total for Job Materials with sub-accounts', '', '', '', '', '', '', '', '10030.60'],
    ['Total for Landscaping Services with sub-accounts', '', '', '', '', '', '', '', '10530.60'],

    ['Maintenance and Repair'],
    ['', 'Maintenance and Repair', '02/08/2024', 'Expense', '', '', '', 'Checking', '-55.86', '-55.86'],
    ['Total for Maintenance and Repair', '', '', '', '', '', '', '', '-55.86'],

    ['Job Expenses'],
    ['', 'Job Expenses', '01/19/2023', 'Expense', '', '', '', 'Savings', '-46.42', '-46.42'],
    ['Total for Job Expenses', '', '', '', '', '', '', '', '-46.42'],

    ['Job Materials'],
    ['', 'Job Materials', '02/17/2023', 'Bill', '', 'Hirthe-West', '', 'Accounts Payable (A/P)', '-316.47', '-316.47'],
    ['Total for Job Materials', '', '', '', '', '', '', '', '-316.47'],

    ['Plants and Soil'],
    ['', 'Plants and Soil', '02/07/2023', 'Expense', '', '', '', 'Savings', '-93.16', '-93.16'],
    ['Total for Plants and Soil', '', '', '', '', '', '', '', '-93.16'],

    ['Total for Job Materials with sub-accounts', '', '', '', '', '', '', '', '-409.63'],
    ['Total for Job Expenses with sub-accounts', '', '', '', '', '', '', '', '-456.05'],
]

PREAMBLE = [['Sandbox Company'], ['General Ledger'],
            ['January 1, 2023-December 31, 2025'], []]


def _write_csv(tmpdir: Path) -> Path:
    path = tmpdir / 'gl.csv'
    with open(path, 'w', newline='', encoding='utf-8') as handle:
        writer = csv.writer(handle)
        for row in PREAMBLE:
            writer.writerow(row)
        writer.writerow(HEADER)
        for row in ROWS:
            writer.writerow(row)
    return path


def _write_xlsx(tmpdir: Path) -> Path:
    path = tmpdir / 'gl.xlsx'
    book = openpyxl.Workbook()
    sheet = book.active
    for row in PREAMBLE:
        sheet.append(row)
    sheet.append(HEADER)
    for row in ROWS:
        sheet.append(row)
    book.save(path)
    return path


EXPECTED = {
    'Landscaping Services': 1,
    'Landscaping Services:Job Materials': 1,
    'Landscaping Services:Job Materials:Plants and Soil': 1,
    'Maintenance and Repair': 1,
    'Job Expenses': 1,
    'Job Expenses:Job Materials': 1,
    'Job Expenses:Job Materials:Plants and Soil': 1,
}


@pytest.fixture(params=['csv', 'xlsx'])
def parsed(request):
    converter = GeneralLedgerConverter()
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        if request.param == 'csv':
            return converter.parse_csv(_write_csv(tmpdir))
        return converter.parse_xlsx(_write_xlsx(tmpdir))


def test_colliding_leaf_names_are_kept_apart(parsed):
    """Both a revenue and an expense 'Job Materials' survive, fully qualified."""
    assert set(parsed['accounts']) == set(EXPECTED)


def test_no_rows_are_dropped(parsed):
    """Every transaction row in the export reaches the output."""
    total = sum(len(v['transactions']) for v in parsed['accounts'].values())
    assert total == sum(EXPECTED.values())


def test_rows_land_in_the_right_account(parsed):
    """The revenue deposit and the expense bill do not end up in one section."""
    revenue = parsed['accounts']['Landscaping Services:Job Materials']['transactions']
    expense = parsed['accounts']['Job Expenses:Job Materials']['transactions']
    assert [t['type'] for t in revenue] == ['Deposit']
    assert [t['type'] for t in expense] == ['Bill']
    assert revenue[0]['amount'] == '2293.04'
    assert expense[0]['amount'] == '-316.47'


def test_subtree_total_is_not_read_as_the_accounts_own_total(parsed):
    """'Total for X with sub-accounts' must not overwrite X's own total."""
    assert parsed['accounts']['Landscaping Services:Job Materials']['total'] == '2293.04'
    assert parsed['accounts']['Job Expenses:Job Materials']['total'] == '-316.47'


def test_top_level_account_is_not_nested(parsed):
    """A bare expense account keeps its own name, not a parent's path."""
    assert 'Maintenance and Repair' in parsed['accounts']
    assert parsed['accounts']['Maintenance and Repair']['transactions'][0]['amount'] == '-55.86'


def test_parent_detection_uses_subtree_markers_only():
    """Leaves never open a nesting level; only accounts with subtree totals do."""
    cells = [r[0] for r in ROWS]
    parents = GeneralLedgerConverter.gl_parent_accounts(cells)
    assert parents == {'Landscaping Services', 'Job Materials', 'Job Expenses'}
    assert 'Plants and Soil' not in parents
    assert 'Maintenance and Repair' not in parents


def test_flat_export_without_subaccounts_is_unchanged():
    """A GL with no hierarchy keys sections exactly as before."""
    flat = [
        ['Checking'],
        ['', 'Checking', '01/06/2023', 'Expense', '', '', '', 'Depreciation', '-645.16', '-645.16'],
        ['Total for Checking', '', '', '', '', '', '', '', '-645.16'],
        ['Utilities'],
        ['', 'Utilities', '01/27/2023', 'Expense', '', '', '', 'Checking', '-40.40', '-40.40'],
        ['Total for Utilities', '', '', '', '', '', '', '', '-40.40'],
    ]
    converter = GeneralLedgerConverter()
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / 'flat.csv'
        with open(path, 'w', newline='', encoding='utf-8') as handle:
            writer = csv.writer(handle)
            for row in PREAMBLE:
                writer.writerow(row)
            writer.writerow(HEADER)
            for row in flat:
                writer.writerow(row)
        parsed = converter.parse_csv(path)
    assert set(parsed['accounts']) == {'Checking', 'Utilities'}
    assert parsed['accounts']['Checking']['total'] == '-645.16'


def test_unresolvable_collision_merges_instead_of_dropping():
    """If a format still collides, rows are merged, never silently discarded."""
    converter = GeneralLedgerConverter()
    accounts = {}
    converter.gl_store_account(accounts, 'Job Materials', 'id-1', [{'amount': '1.00'}], 1.0)
    converter.gl_store_account(accounts, 'Job Materials', 'id-2', [{'amount': '2.00'}], 2.0)
    assert len(accounts['Job Materials']['transactions']) == 2
    assert accounts['Job Materials']['total'] == '3.00'
