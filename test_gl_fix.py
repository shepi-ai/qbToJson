#!/usr/bin/env python3
"""Test GL parser fix for false account sections"""
from generalLedgerConverter import GeneralLedgerConverter

converter = GeneralLedgerConverter()
result = converter.convert_file('sampleReports/2023GL/GeneralLedger-01.23.pdf')

# Count accounts
accounts = result.get('rows', {}).get('row', [])
print(f'\n=== TEST RESULTS - January 2023 PDF ===')
print(f'Total accounts extracted: {len(accounts)}')
print(f'Expected: 7-10 accounts (without false sub-account sections)')

# List account names
account_names = []
for account_row in accounts:
    if account_row.get('type') == 'SECTION':
        header_data = account_row.get('header', {}).get('colData', [])
        if header_data:
            account_name = header_data[0].get('value', '')
            account_names.append(account_name)

print(f'\nAccounts found:')
false_accounts = ['Workers Compensation', 'Equipment Rental', 'Accounting', 'Bookkeeper',
                  'Installation', 'Labor', 'Fuel', 'Fountains and Garden Lighting',
                  'Plants and Soil', 'Telephone', 'Permits']

for name in account_names:
    if name in false_accounts:
        print(f'  X  {name} (FALSE - should not be separate account)')
    else:
        print(f'  OK {name}')

false_found = [name for name in account_names if name in false_accounts]
if false_found:
    print(f'\nSTATUS: FAILED - Still finding {len(false_found)} false accounts')
    print(f'False accounts: {false_found}')
else:
    print(f'\nSTATUS: SUCCESS - No false account sections!')
