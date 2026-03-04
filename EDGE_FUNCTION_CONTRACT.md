# Edge Function → qbToJson API Contract

## Success Requirements for Edge Function

Based on the logs showing 401 errors from `Deno/2.1.4 (SupabaseEdgeRuntime)`, here's exactly what the edge function needs to do for successful uploads.

---

## ✅ Required for Success

### 1. **API Key Authentication** (CRITICAL)

**MUST include this header in EVERY request:**

```typescript
headers: {
  'x-api-key': Deno.env.get('QBTOJSON_API_KEY')!
}
```

**Without this header → 401 Unauthorized** ❌

---

### 2. **Request Format**

#### For Direct File Upload (General Ledger Example)

```typescript
const formData = new FormData();
formData.append('file', fileBlob, filename);
formData.append('save_to_db', 'true');
formData.append('project_id', projectId);
formData.append('source_document_id', documentId); // optional

const response = await fetch(
  'https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/general-ledger',
  {
    method: 'POST',
    headers: {
      'x-api-key': Deno.env.get('QBTOJSON_API_KEY')!  // REQUIRED!
    },
    body: formData
  }
);
```

---

## 📋 Complete Contract

### Request Headers

| Header | Required | Value | Notes |
|--------|----------|-------|-------|
| `x-api-key` | **YES** | `EWa7M/0xdOyribITySgZ1vZ6ilaqFpd3Mo29Rjyhtc4=` | Must be set in Edge Function secrets |

### Request Body (multipart/form-data)

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `file` | File | **YES** | The document file to convert |
| `save_to_db` | string | No | "true" to save to database |
| `project_id` | string | **YES if save_to_db=true** | Supabase project UUID |
| `source_document_id` | string | No | Optional document identifier |

---

## ✅ Success Response (200 OK)

### General Ledger Upload Success

```json
{
  "success": true,
  "data": {
    "header": { ... },
    "columns": { ... },
    "rows": {
      "row": [ ... ]
    }
  },
  "accounts": 45,
  "filename": "GeneralLedger.csv",
  "saved_to_db": true,
  "project_id": "proj_123",
  "coa_derived": true,        // NEW: Was COA auto-derived?
  "derived_count": 30          // NEW: Number of accounts derived
}
```

**Key Response Fields:**
- `success` (boolean): Always check this first
- `data` (object): The converted GL data
- `saved_to_db` (boolean): Confirms database save
- `coa_derived` (boolean): **NEW FEATURE** - true if COA was auto-derived
- `derived_count` (number): **NEW FEATURE** - how many accounts were derived

---

## ❌ Error Responses

### 401 Unauthorized (Current Issue)

```json
{
  "error": "Unauthorized",
  "details": "API key required. Include 'x-api-key' header."
}
```

**Cause:** Missing `x-api-key` header  
**Fix:** Add header with API key

### 400 Bad Request

```json
{
  "error": "No file provided"
}
```

**Cause:** Missing file in form data  
**Fix:** Ensure file is appended to FormData

### 500 Internal Server Error

```json
{
  "error": "Failed to convert file",
  "details": "Could not parse PDF structure"
}
```

**Cause:** File parsing failed  
**Fix:** Check file format, try different converter

---

## 📝 Complete Edge Function Example

### Supabase Edge Function Template

```typescript
// supabase/functions/process-quickbooks-file/index.ts

import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

serve(async (req) => {
  try {
    // Get the file from the request
    const formData = await req.formData();
    const file = formData.get('file') as File;
    const projectId = formData.get('project_id') as string;
    const documentId = formData.get('document_id') as string;
    
    if (!file || !projectId) {
      return new Response(
        JSON.stringify({ error: 'Missing file or project_id' }),
        { status: 400 }
      );
    }
    
    // Prepare request to qbToJson API
    const qbFormData = new FormData();
    qbFormData.append('file', file);
    qbFormData.append('save_to_db', 'true');
    qbFormData.append('project_id', projectId);
    if (documentId) {
      qbFormData.append('source_document_id', documentId);
    }
    
    // Call qbToJson API with API key
    const response = await fetch(
      'https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/general-ledger',
      {
        method: 'POST',
        headers: {
          'x-api-key': Deno.env.get('QBTOJSON_API_KEY')!  // CRITICAL!
        },
        body: qbFormData
      }
    );
    
    // Handle response
    if (!response.ok) {
      const error = await response.json();
      console.error('qbToJson API error:', error);
      return new Response(
        JSON.stringify({ 
          error: 'Failed to process document',
          details: error 
        }),
        { status: response.status }
      );
    }
    
    const result = await response.json();
    
    // Check if COA was auto-derived (NEW FEATURE)
    if (result.coa_derived) {
      console.log(`✅ Auto-derived ${result.derived_count} accounts from GL`);
      // You might want to notify the frontend about this
    }
    
    return new Response(
      JSON.stringify({
        success: true,
        data: result.data,
        coa_derived: result.coa_derived,
        derived_count: result.derived_count
      }),
      { 
        status: 200,
        headers: { 'Content-Type': 'application/json' }
      }
    );
    
  } catch (error) {
    console.error('Edge function error:', error);
    return new Response(
      JSON.stringify({ 
        error: 'Internal error',
        details: error.message 
      }),
      { status: 500 }
    );
  }
});
```

---

## 🔐 Setting Up Edge Function Secrets

### In Supabase Dashboard

1. Go to **Edge Functions** → **Secrets**
2. Add secret:
   ```
   Key: QBTOJSON_API_KEY
   Value: EWa7M/0xdOyribITySgZ1vZ6ilaqFpd3Mo29Rjyhtc4=
   ```
3. Save and redeploy the edge function

### Via CLI

```bash
supabase secrets set QBTOJSON_API_KEY=EWa7M/0xdOyribITySgZ1vZ6ilaqFpd3Mo29Rjyhtc4=
```

---

## 🧪 Testing

### Test with API Key

```bash
curl -X POST \
  -H "x-api-key: EWa7M/0xdOyribITySgZ1vZ6ilaqFpd3Mo29Rjyhtc4=" \
  -F "file=@GeneralLedger.csv" \
  -F "save_to_db=true" \
  -F "project_id=test-123" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/general-ledger
```

**Expected:** 200 OK with converted data

### Test without API Key

```bash
curl -X POST \
  -F "file=@GeneralLedger.csv" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/general-ledger
```

**Expected:** 401 Unauthorized

---

## 📊 Available Endpoints

All require `x-api-key` header:

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/api/convert/accounts` | POST | Chart of Accounts |
| `/api/convert/general-ledger` | POST | General Ledger (auto-derives COA) |
| `/api/convert/trial-balance` | POST | Trial Balance |
| `/api/convert/balance-sheet` | POST | Balance Sheet |
| `/api/convert/profit-loss` | POST | Profit & Loss |
| `/api/convert/cash-flow` | POST | Cash Flow |
| `/api/convert/journal-entries` | POST | Journal Entries |
| `/api/convert/accounts-payable` | POST | Accounts Payable |
| `/api/convert/accounts-receivable` | POST | Accounts Receivable |

**Public endpoints (no auth):**
- `GET /health`
- `GET /api/info`

---

## 🎯 Quick Checklist for Edge Function

- [ ] Edge function has `QBTOJSON_API_KEY` secret set
- [ ] Request includes `x-api-key` header
- [ ] FormData includes `file` field
- [ ] FormData includes `save_to_db=true` (if saving)
- [ ] FormData includes `project_id` (if saving)
- [ ] Response handling checks `success` field
- [ ] Response handling checks `coa_derived` field (for GL)
- [ ] Error handling for 401, 400, 500 status codes

---

## 📚 Related Documentation

- **Full API Docs:** `/Users/araboin/qofeai/qbToJson/API_DOCUMENTATION.md`
- **Frontend Integration:** `/Users/araboin/qofeai/qbToJson/FRONTEND_INTEGRATION.md`
- **COA Auto-Derive:** `/Users/araboin/qofeai/qbToJson/CHART_OF_ACCOUNTS_DERIVATION.md`

---

## 🚨 Current Issue Summary

**Problem:** Edge function is calling qbToJson API without `x-api-key` header  
**Evidence:** 20+ requests with 401 status in logs  
**Solution:** Add the header shown above  
**API Key:** `EWa7M/0xdOyribITySgZ1vZ6ilaqFpd3Mo29Rjyhtc4=`

Once the edge function includes the API key header, all 401 errors will stop and uploads will succeed!
