# General Ledger Batch Processing Comparison Report
**Date**: March 25, 2026  
**Test**: Batch processing of 12 monthly 2023 GL PDFs vs expected 2023GL.json

## Executive Summary

✅ **Date range matches**: 2023-01-01 to 2023-12-31  
✅ **Report basis matches**: ACCRUAL  
❌ **Account count**: 69 (batch) vs 48 (expected) - **21 extra accounts**  
❌ **Transaction count**: 1,757 (batch) vs 1,888 (expected) - **131 missing transactions**

## Root Cause Analysis

### Issue #1: False Account Sections (21 extra accounts)

**Problem**: Sub-account names from SPLIT column are being misidentified as new account headers.

**Example from PDF**:
```
Savings 01/25/2023 Expense Legal & Professional Fees:Bookkeeper -28.84 453,746.45
```

The SPLIT column contains `Legal & Professional Fees:Bookkeeper`, but during text extraction/preprocessing, this hierarchical name can wrap or appear alone, and the parser treats **"Bookkeeper"** as a new account section.

**False accounts created** (should be sub-accounts in SPLIT column):
- Accounting (part of Legal & Professional Fees:Accounting)
- Bookkeeper (part of Legal & Professional Fees:Bookkeeper)  
- Workers Compensation (part of Insurance:Workers Compensation)
- Building Repairs (part of Maintenance and Repair:Building Repairs)
- Fuel (part of Automobile:Fuel)
- Gas and Electric (part of Utilities:Gas and Electric)
- Telephone (part of Utilities:Telephone)
- Job Materials, Plants and Soil, Sprinklers, etc. (parts of hierarchical Landscaping Services accounts)

### Issue #2: Transaction Type Truncation

**Problem**: Transaction type "Bill Payment (Check)" is split across lines in PDF extraction.

**PDF Format**:
```
Checking 01/01/2023 Bill Payment Yundt, Torp and Accounts Payable (A/P) -259.78 150,841.20
(Check) Stoltenberg
```

**Current Parser Output**: 
- Extracts as `"Payment"` or `"Bill"` (truncated)

**Expected Output**: 
- Should extract as `"Bill Payment (Check)"` (complete type)

**Impact**:
- Transaction type mismatch: "Payment" (453 in batch) vs "Bill Payment (Check)" (258 in expected)
- Also affects: "deposit" (148 in batch) vs "Deposit" (106 in expected)

### Issue #3: Missing Transactions (131 total)

**Breakdown by account**:
- Accounts Receivable: -268 transactions
- Services: -10 transactions  
- Design income: -5 transactions
- Pest Control Services: -4 transactions
- And 24 other accounts with minor differences

**Likely Causes**:
1. Lines not being properly merged during preprocessing
2. Transactions rejected due to parsing failures
3. Continuation lines causing data loss

## Detailed Comparison Results

### Header Comparison
| Field | Batch | Expected | Match |
|-------|-------|----------|-------|
| Start Period | 2023-01-01 | 2023-01-01 | ✅ |
| End Period | 2023-12-31 | 2023-12-31 | ✅ |
| Report Basis | ACCRUAL | ACCRUAL | ✅ |

### Account Counts
| Metric | Batch | Expected | Difference |
|--------|-------|----------|------------|
| Total Accounts | 69 | 48 | +21 |
| Checking Trans | 192 | 192 | ✅ 0 |
| Savings Trans | 88 | 88 | ✅ 0 |
| A/R Trans | 182 | 450 | ❌ -268 |
| A/P Trans | 296 | 309 | ❌ -13 |

### Transaction Type Comparison
| Type | Batch | Expected | Difference |
|------|-------|----------|------------|
| Payment | 453 | 390 | +63 |
| Bill Payment (Check) | 0 | 258 | -258 |
| Invoice | 299 | 383 | -84 |
| Expense | 234 | 152 | +82 |
| Sales Receipt | 129 | 175 | -46 |
| deposit (lowercase) | 148 | 0 | +148 |
| Deposit | 103 | 106 | -3 |
| Journal Entry | 92 | 74 | +18 |
| Transfer | 48 | 48 | ✅ 0 |

## Recommended Fixes

### Fix #1: Improve Account Header Detection

**Current logic**: Looks ahead for "Beginning Balance" or account-prefixed transaction
**Problem**: Sub-account names can match this pattern

**Solution**: Add additional validation checks:
1. Real account headers should NOT contain colons (`:`) - those are hierarchical split accounts
2. Real account headers should match known QuickBooks account types or be in the account list
3. Look for "Total for [Account Name]" footer to confirm account section

**Code changes needed** in `parse_pdf()`:
```python
# Enhanced account header detection
if line and not re.search(r'\d{1,2}/\d{1,2}/\d{4}', line):
    is_real_account = False
    
    # Skip if line contains colon (hierarchical sub-account name)
    if ':' in line:
        is_real_account = False
    # Look ahead for "Beginning Balance"
    elif line_idx + 1 < len(lines):
        next_line = lines[line_idx + 1].strip()
        if next_line.startswith('Beginning Balance'):
            is_real_account = True
        elif next_line.startswith(line + ' ') and re.search(r'\d{1,2}/\d{1,2}/\d{4}', next_line):
            # Extra validation: check if we can find a "Total for" footer
            is_real_account = True
```

### Fix #2: Capture Complete Transaction Types

**Problem**: `_parse_gl_transaction_line()` regex patterns don't capture multi-line types

**Solution**: Update transaction type regex patterns:
```python
# Current pattern misses "(Check)" suffix
transaction_type_pattern = r'(Bill|Bill Payment|Payment|Deposit|Expense|Transfer|Invoice|Sales Receipt|Credit Memo|Journal Entry|Check)'

# Should be:
transaction_type_pattern = r'(Bill Payment \(Check\)|Bill|Payment|Deposit|Expense|Transfer|Invoice|Sales Receipt|Credit Memo|Journal Entry|Check)'
```

**Better approach**: Since preprocessing should merge "(Check)" with "Bill Payment", verify preprocessor is working correctly.

### Fix #3: Enhanced Line Preprocessing

**Verify** `_preprocess_pdf_lines()` properly merges:
- "(Check)" continuations with "Bill Payment"
- "(Type)" patterns with preceding transaction types
- Wrapped vendor/customer names

## Test Results Summary

### ✅ What Works Well
1. Date range extraction (2023-01-01 to 2023-12-31)
2. Header information (report basis, currency)
3. Checking account: 192 transactions match perfectly
4. Savings account: 88 transactions match perfectly
5. Transfer transactions: 48 match perfectly
6. Overall structure and JSON format

### ❌ What Needs Fixing
1. Account structure detection (21 false account sections)
2. Transaction type extraction (truncation issues)
3. Missing 131 transactions (primarily in A/R, Services, Design income)
4. Balance discrepancies due to missing/miscategorized transactions

## Impact Assessment

**Severity**: High  
**Data Integrity**: Medium - Structure is correct but ~7% of transactions missing and accounts over-segmented

**Recommendation**: Fix account header detection and transaction type extraction before using in production. The parser successfully processes all 12 monthly files, but accuracy needs improvement.

## Next Steps

1. **Fix account header detection** - Add colon check and better validation
2. **Fix transaction type extraction** - Ensure full types captured including "(Check)"
3. **Debug A/R transactions** - Investigate why 268 transactions missing from A/R
4. **Re-run batch test** - Verify fixes produce matching output
5. **Add regression tests** - Ensure monthly GL parsing maintains accuracy
