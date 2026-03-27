#!/usr/bin/env python3
"""
Compare batch processing GL output with expected 2023GL.json
"""
import json
from datetime import datetime
from collections import defaultdict

def extract_gl_stats(data):
    """Extract key statistics from GL JSON structure"""
    stats = {
        'accounts': [],
        'total_transactions': 0,
        'date_range': {'min': None, 'max': None},
        'account_balances': {},
        'transaction_types': defaultdict(int)
    }
    
    # Get header info
    header = data.get('header', {})
    stats['start_period'] = header.get('startPeriod')
    stats['end_period'] = header.get('endPeriod')
    stats['report_basis'] = header.get('reportBasis')
    
    # Process rows (accounts)
    rows = data.get('rows', {}).get('row', [])
    
    for account_row in rows:
        if account_row.get('type') != 'SECTION':
            continue
            
        # Get account name from header
        header_data = account_row.get('header', {}).get('colData', [])
        if header_data:
            account_name = header_data[0].get('value', '')
            account_id = header_data[0].get('id', '')
            
            if not account_name:
                continue
                
            # Count transactions in this account
            transaction_rows = account_row.get('rows', {}).get('row', [])
            transaction_count = 0
            beginning_balance = None
            ending_balance = None
            
            for trans_row in transaction_rows:
                col_data = trans_row.get('colData', [])
                if not col_data:
                    continue
                    
                # Check if it's a transaction or balance line
                trans_type = col_data[1].get('value', '') if len(col_data) > 1 else ''
                
                if trans_type == 'Beginning Balance':
                    # Get beginning balance
                    if len(col_data) > 7:
                        beginning_balance = col_data[7].get('value', '')
                elif trans_type:  # Has a transaction type
                    transaction_count += 1
                    stats['total_transactions'] += 1
                    stats['transaction_types'][trans_type] += 1
                    
                    # Track date range
                    if len(col_data) > 0:
                        date_str = col_data[0].get('value', '')
                        if date_str and date_str != 'Beginning Balance':
                            try:
                                # Handle different date formats
                                for fmt in ['%Y-%m-%d', '%m/%d/%Y']:
                                    try:
                                        trans_date = datetime.strptime(date_str, fmt).date()
                                        if stats['date_range']['min'] is None or trans_date < stats['date_range']['min']:
                                            stats['date_range']['min'] = trans_date
                                        if stats['date_range']['max'] is None or trans_date > stats['date_range']['max']:
                                            stats['date_range']['max'] = trans_date
                                        break
                                    except ValueError:
                                        continue
                            except:
                                pass
                    
                    # Get ending balance (last transaction balance)
                    if len(col_data) > 7:
                        balance_val = col_data[7].get('value', '')
                        if balance_val:
                            ending_balance = balance_val
            
            stats['accounts'].append({
                'name': account_name,
                'id': account_id,
                'transaction_count': transaction_count,
                'beginning_balance': beginning_balance,
                'ending_balance': ending_balance
            })
    
    return stats

def compare_gl_files(file1_path, file2_path):
    """Compare two GL JSON files"""
    print("=" * 80)
    print("GENERAL LEDGER COMPARISON")
    print("=" * 80)
    
    # Load files
    with open(file1_path, 'r') as f:
        batch_data = json.load(f)
    
    with open(file2_path, 'r') as f:
        expected_data = json.load(f)
    
    # Extract statistics
    print("\n📊 Extracting statistics from batch output...")
    batch_stats = extract_gl_stats(batch_data)
    
    print("📊 Extracting statistics from expected output...")
    expected_stats = extract_gl_stats(expected_data)
    
    # Compare headers
    print("\n" + "=" * 80)
    print("HEADER COMPARISON")
    print("=" * 80)
    
    print(f"\nDate Range:")
    print(f"  Batch:    {batch_stats['start_period']} to {batch_stats['end_period']}")
    print(f"  Expected: {expected_stats['start_period']} to {expected_stats['end_period']}")
    print(f"  Match: {'✅' if batch_stats['start_period'] == expected_stats['start_period'] and batch_stats['end_period'] == expected_stats['end_period'] else '❌'}")
    
    print(f"\nReport Basis:")
    print(f"  Batch:    {batch_stats['report_basis']}")
    print(f"  Expected: {expected_stats['report_basis']}")
    print(f"  Match: {'✅' if batch_stats['report_basis'] == expected_stats['report_basis'] else '❌'}")
    
    # Compare accounts
    print("\n" + "=" * 80)
    print("ACCOUNT COMPARISON")
    print("=" * 80)
    
    print(f"\nTotal Accounts:")
    print(f"  Batch:    {len(batch_stats['accounts'])}")
    print(f"  Expected: {len(expected_stats['accounts'])}")
    print(f"  Match: {'✅' if len(batch_stats['accounts']) == len(expected_stats['accounts']) else '❌'}")
    
    # Create account name maps
    batch_accounts = {acc['name']: acc for acc in batch_stats['accounts']}
    expected_accounts = {acc['name']: acc for acc in expected_stats['accounts']}
    
    all_account_names = sorted(set(batch_accounts.keys()) | set(expected_accounts.keys()))
    
    print(f"\nAccount Details (showing first 10):")
    for i, acc_name in enumerate(all_account_names[:10]):
        if acc_name in batch_accounts and acc_name in expected_accounts:
            batch_acc = batch_accounts[acc_name]
            expected_acc = expected_accounts[acc_name]
            
            trans_match = batch_acc['transaction_count'] == expected_acc['transaction_count']
            print(f"\n  {acc_name}:")
            print(f"    Transactions: Batch={batch_acc['transaction_count']}, Expected={expected_acc['transaction_count']} {'✅' if trans_match else '❌'}")
            print(f"    End Balance:  Batch={batch_acc['ending_balance']}, Expected={expected_acc['ending_balance']}")
        elif acc_name in batch_accounts:
            print(f"\n  ⚠️  {acc_name}: Only in batch output")
        else:
            print(f"\n  ⚠️  {acc_name}: Only in expected output")
    
    if len(all_account_names) > 10:
        print(f"\n  ... and {len(all_account_names) - 10} more accounts")
    
    # Compare transactions
    print("\n" + "=" * 80)
    print("TRANSACTION COMPARISON")
    print("=" * 80)
    
    print(f"\nTotal Transactions:")
    print(f"  Batch:    {batch_stats['total_transactions']}")
    print(f"  Expected: {expected_stats['total_transactions']}")
    print(f"  Match: {'✅' if batch_stats['total_transactions'] == expected_stats['total_transactions'] else '❌'}")
    
    print(f"\nTransaction Date Range (from data):")
    print(f"  Batch:    {batch_stats['date_range']['min']} to {batch_stats['date_range']['max']}")
    print(f"  Expected: {expected_stats['date_range']['min']} to {expected_stats['date_range']['max']}")
    
    # Transaction types
    print(f"\nTransaction Types (top 10):")
    all_types = sorted(set(batch_stats['transaction_types'].keys()) | set(expected_stats['transaction_types'].keys()))
    
    for trans_type in sorted(all_types, key=lambda t: batch_stats['transaction_types'].get(t, 0) + expected_stats['transaction_types'].get(t, 0), reverse=True)[:10]:
        batch_count = batch_stats['transaction_types'][trans_type]
        expected_count = expected_stats['transaction_types'][trans_type]
        match = '✅' if batch_count == expected_count else '❌'
        print(f"  {trans_type:30s} Batch={batch_count:4d}, Expected={expected_count:4d} {match}")
    
    # Summary
    print("\n" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    
    issues = []
    
    if batch_stats['start_period'] != expected_stats['start_period'] or batch_stats['end_period'] != expected_stats['end_period']:
        issues.append("❌ Date range mismatch")
    else:
        print("✅ Date ranges match")
    
    if len(batch_stats['accounts']) != len(expected_stats['accounts']):
        issues.append(f"❌ Account count mismatch: {len(batch_stats['accounts'])} vs {len(expected_stats['accounts'])}")
    else:
        print("✅ Account counts match")
    
    if batch_stats['total_transactions'] != expected_stats['total_transactions']:
        diff = batch_stats['total_transactions'] - expected_stats['total_transactions']
        issues.append(f"❌ Transaction count mismatch: {batch_stats['total_transactions']} vs {expected_stats['total_transactions']} (diff: {diff:+d})")
    else:
        print("✅ Transaction counts match")
    
    # Check for missing accounts
    batch_only = set(batch_accounts.keys()) - set(expected_accounts.keys())
    expected_only = set(expected_accounts.keys()) - set(batch_accounts.keys())
    
    if batch_only:
        issues.append(f"⚠️  Accounts only in batch: {', '.join(sorted(batch_only))}")
    
    if expected_only:
        issues.append(f"⚠️  Accounts only in expected: {', '.join(sorted(expected_only))}")
    
    if not batch_only and not expected_only:
        print("✅ Same accounts in both files")
    
    if issues:
        print("\n⚠️  ISSUES FOUND:")
        for issue in issues:
            print(f"  {issue}")
    else:
        print("\n🎉 FILES MATCH PERFECTLY!")
    
    # Detailed account-by-account comparison
    print("\n" + "=" * 80)
    print("DETAILED ACCOUNT COMPARISON")
    print("=" * 80)
    
    print("\nAccounts with transaction count mismatches:")
    mismatch_found = False
    for acc_name in sorted(set(batch_accounts.keys()) & set(expected_accounts.keys())):
        batch_acc = batch_accounts[acc_name]
        expected_acc = expected_accounts[acc_name]
        
        if batch_acc['transaction_count'] != expected_acc['transaction_count']:
            mismatch_found = True
            diff = batch_acc['transaction_count'] - expected_acc['transaction_count']
            print(f"  {acc_name}: Batch={batch_acc['transaction_count']}, Expected={expected_acc['transaction_count']} (diff: {diff:+d})")
    
    if not mismatch_found:
        print("  ✅ All accounts have matching transaction counts!")
    
    print("\n" + "=" * 80)

if __name__ == '__main__':
    compare_gl_files(
        'test_2023gl_batch.json',
        'sampleJson/2023GL.json'
    )
