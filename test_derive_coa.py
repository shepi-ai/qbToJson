#!/usr/bin/env python3
"""
Test script for deriving Chart of Accounts from General Ledger
"""

import json
from pathlib import Path
from accountsInferenceConverter import convert_general_ledger_to_coa

def test_derive_coa_from_sample_gl():
    """Test deriving CoA from the sample General Ledger JSON"""
    
    # Load sample GL data
    gl_path = Path('sampleJson/generalledger.json')
    
    if not gl_path.exists():
        print(f"❌ Sample file not found: {gl_path}")
        return False
    
    print(f"Loading sample GL data from {gl_path}...")
    with open(gl_path, 'r') as f:
        gl_data = json.load(f)
    
    # Derive Chart of Accounts
    print("\nDeriving Chart of Accounts...")
    derived_coa = convert_general_ledger_to_coa(gl_data)
    
    print(f"✅ Successfully derived {len(derived_coa)} accounts")
    
    # Display sample accounts
    print("\nSample derived accounts:")
    print("-" * 80)
    
    for account in derived_coa[:5]:  # Show first 5 accounts
        print(f"ID: {account['id']:4s} | {account['name']:40s} | {account['classification']:10s} | {account['accountType']}")
    
    if len(derived_coa) > 5:
        print(f"... and {len(derived_coa) - 5} more accounts")
    
    print("-" * 80)
    
    # Verify structure matches qbToJson format
    print("\n✅ Verifying structure matches qbToJson format...")
    required_fields = [
        'id', 'name', 'classification', 'accountType', 'fullyQualifiedName',
        'active', 'currentBalance', 'currencyRef', 'metaData'
    ]
    
    first_account = derived_coa[0]
    missing_fields = [field for field in required_fields if field not in first_account]
    
    if missing_fields:
        print(f"❌ Missing required fields: {missing_fields}")
        return False
    else:
        print("✅ All required fields present")
    
    # Check classifications
    print("\nClassification breakdown:")
    classifications = {}
    for account in derived_coa:
        cls = account['classification']
        classifications[cls] = classifications.get(cls, 0) + 1
    
    for cls, count in sorted(classifications.items()):
        print(f"  {cls:10s}: {count:3d} accounts")
    
    # Save derived CoA to output file
    output_path = Path('sampleJson/derived_chart_of_accounts.json')
    with open(output_path, 'w') as f:
        json.dump(derived_coa, f, indent=2)
    
    print(f"\n✅ Derived CoA saved to {output_path}")
    
    return True

if __name__ == '__main__':
    print("=" * 80)
    print("Testing Chart of Accounts Derivation from General Ledger")
    print("=" * 80)
    print()
    
    success = test_derive_coa_from_sample_gl()
    
    print()
    print("=" * 80)
    if success:
        print("✅ All tests passed!")
    else:
        print("❌ Tests failed")
    print("=" * 80)
