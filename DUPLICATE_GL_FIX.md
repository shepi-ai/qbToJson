# Duplicate General Ledger Records Fix
**Date**: March 31, 2026  
**Issue**: Duplicate GL records in `processed_data` table causing inflated transaction counts

## Problem Summary

### Symptoms
- 7 GL records for 3 periods (should be 3 records)
- 2 records with `source_document_id` (legitimate)
- 5 records with `source_document_id = NULL` (orphans)
- All records have same `created_at` timestamp
- Transaction counts inflated (216K instead of ~85K)

### Root Cause
**Dual-write pattern**: Two code paths both saving to `processed_data`:

1. **Edge Function** (`process-quickbooks-file`)
   - Calls qbToJson `/api/convert-from-storage/general-ledger`
   - Receives converted data
   - Saves to `processed_data` with `source_document_id` ✅

2. **qbToJson Storage Helper** (`convert_from_storage_helper()`)
   - Also called `db_client.save_converted_data()`
   - Inserted to `processed_data` with `source_document_id` often NULL ❌

Both writes went through `db-proxy`, creating duplicates.

## Solution Implemented

### Fix: Remove DB Save from qbToJson Storage Endpoints

**File**: `api_server.py`  
**Function**: `convert_from_storage_helper()` (line 854)

**Changed:**
```python
# OLD CODE (caused duplicates):
save_success = db_client.save_converted_data(
    project_id=project_id,
    data_type=data_type,
    data=result,
    source_document_id=source_document_id,
    filename=file_path.split('/')[-1]
)
return result, save_success
```

**To:**
```python
# NEW CODE (prevents duplicates):
# NOTE: Database save removed to prevent duplicates
# Edge function (process-quickbooks-file) handles all DB writes
# This endpoint is now a pure converter - returns data only

return result, False  # False = not saved (edge function will save)
```

### Result
- qbToJson is now a **pure converter** - returns JSON only
- Edge function is **sole writer** to `processed_data`
- No duplicate inserts
- No orphaned NULL records
- Clean separation of concerns

## Database Cleanup

### Step 1: Find Orphaned Records

```sql
-- Find all orphaned GL records (source_document_id IS NULL)
SELECT 
    id,
    project_id,
    data_type,
    created_at,
    record_count,
    source_type
FROM processed_data
WHERE data_type = 'general_ledger'
  AND source_document_id IS NULL
ORDER BY project_id, created_at;
```

### Step 2: Count Orphans by Project

```sql
-- Count orphans per project
SELECT 
    project_id,
    COUNT(*) as orphan_count,
    SUM(record_count) as total_orphan_records
FROM processed_data
WHERE data_type = 'general_ledger'
  AND source_document_id IS NULL
GROUP BY project_id
ORDER BY orphan_count DESC;
```

### Step 3: Delete Orphaned Records

```sql
-- DELETE orphaned GL records
-- ⚠️ BACKUP DATABASE FIRST!
DELETE FROM processed_data
WHERE data_type = 'general_ledger'
  AND source_document_id IS NULL;

-- Verify deletion
SELECT COUNT(*) as remaining_orphans
FROM processed_data  
WHERE data_type = 'general_ledger'
  AND source_document_id IS NULL;
-- Should return 0
```

### Step 4: Audit All Document Types

```sql
-- Find orphans across ALL document types
SELECT 
    data_type,
    COUNT(*) as orphan_count,
    SUM(record_count) as total_records
FROM processed_data
WHERE source_document_id IS NULL
GROUP BY data_type
ORDER BY orphan_count DESC;
```

## Verification After Fix

### Test 1: Single Upload
```bash
# Upload GL file for project
# Expected: 1 record in processed_data with source_document_id set
```

### Test 2: Reprocessing
```bash
# Re-upload same GL file
# Expected: Still 1 record (edge function should handle dedup)
```

### Test 3: Transaction Counts
```bash
# Check canonical_transactions count
# Expected: Count should match actual GL transactions (~85K, not 216K)
```

## Future Prevention

### Option A: Database Constraint (Recommended)
```sql
-- Add unique constraint to prevent duplicates at DB level
CREATE UNIQUE INDEX idx_processed_data_unique_source
ON processed_data (
    project_id, 
    data_type, 
    COALESCE(source_document_id, '00000000-0000-0000-0000-000000000000')
);
```

### Option B: Edge Function Dedup
Already implemented in edge function - delete existing record before insert.

## Impact Assessment

**Before Fix:**
- Duplicate GL records causing 2-3x data inflation
- Inflated transaction counts downstream
- Orphaned records with NULL source_document_id

**After Fix:**
- Single source of truth for DB writes (edge function)
- No duplicate inserts
- Clean data lineage with proper source_document_id tracking
- Reduced storage and processing overhead

## Related Files Modified

- ✅ `api_server.py` - Removed `db_client.save_converted_data()` from `convert_from_storage_helper()`

## Related Files (No Changes Needed)

- ✅ `db_client.py` - No changes (single INSERT is fine when only one caller)
- ✅ `batch_processor.py` - No changes (doesn't save to DB)
- ✅ `generalLedgerConverter.py` - No changes (pure converter)

## Deployment Notes

1. Deploy updated `api_server.py` to Cloud Run
2. Run cleanup SQL to remove existing orphans
3. Monitor logs for any storage endpoint calls with `saved_to_db=true` (should now be false)
4. Verify transaction counts normalize in downstream systems

## Testing Checklist

- [ ] Deploy updated qbToJson API
- [ ] Upload new GL file → verify single record created
- [ ] Re-upload same GL → verify no duplicates
- [ ] Run cleanup SQL → remove orphans
- [ ] Verify `canonical_transactions` count drops to expected level
- [ ] Audit other projects for orphans
- [ ] Consider adding DB unique constraint for long-term prevention

