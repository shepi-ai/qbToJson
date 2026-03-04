# Document Converter API Documentation

## Overview

The Document Converter API is a REST API service that converts financial documents (QuickBooks reports) from various formats (CSV, XLSX, PDF) into standardized JSON structures. It supports 11 different document types and provides batch processing capabilities.

**Base URL (Production):** `https://qbtojson-7lqwugl3xa-uc.a.run.app`  
**Base URL (Local):** `http://localhost:5000`

## Features

- ✅ **11 Document Types**: Chart of Accounts, Balance Sheet, P&L, Trial Balance, Cash Flow, General Ledger, Journal Entries, AR/AP, Customer/Vendor Concentration
- ✅ **3 Input Formats**: CSV, XLSX, PDF
- ✅ **Direct Upload**: Upload files directly via multipart/form-data
- ✅ **Storage-based**: Convert files already in Supabase Storage
- ✅ **Batch Processing**: Process multiple files or ZIP archives
- ✅ **Database Integration**: Optionally save results to Supabase
- ✅ **Account Lookup**: Cache and lookup chart of accounts

## Authentication

**All API endpoints (except `/health` and `/api/info`) require API key authentication.**

### Required Header

Include the following header in all API requests:

```http
x-api-key: YOUR_API_KEY
```

### Example with cURL

```bash
curl -X POST \
  -H "x-api-key: YOUR_API_KEY" \
  -F "file=@TrialBalance.pdf" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/trial-balance
```

### Authentication Errors

**401 Unauthorized - Missing API Key:**
```json
{
  "error": "Unauthorized",
  "details": "API key required. Include 'x-api-key' header."
}
```

**401 Unauthorized - Invalid API Key:**
```json
{
  "error": "Unauthorized",
  "details": "Invalid API key"
}
```

### Public Endpoints (No Auth Required)

- `GET /health` - Health check
- `GET /api/info` - API information

## Common Headers

```http
Content-Type: multipart/form-data  # For file uploads
Content-Type: application/json      # For JSON requests
```

## Rate Limits

No rate limits are currently enforced, but reasonable usage is expected.

---

## Quick Start

### 1. Health Check

```bash
curl https://qbtojson-7lqwugl3xa-uc.a.run.app/health
```

**Response:**
```json
{
  "status": "healthy",
  "service": "Document Converter API",
  "db_configured": true,
  "endpoints": { ... }
}
```

### 2. Convert a Document

```bash
curl -X POST \
  -F "file=@TrialBalance.pdf" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/trial-balance
```

**Response:**
```json
{
  "success": true,
  "data": {
    "monthlyReports": [ ... ],
    "summary": { ... }
  },
  "months": 1,
  "filename": "TrialBalance.pdf"
}
```

---

## API Endpoints

### System Endpoints

#### GET /health
Health check endpoint

**Response:**
```json
{
  "status": "healthy",
  "service": "Document Converter API",
  "db_configured": boolean,
  "endpoints": { ... }
}
```

#### GET /api/info
Get detailed API information and examples

**Response:**
```json
{
  "name": "Document Converter API",
  "version": "1.0.0",
  "endpoints": [ ... ],
  "examples": { ... }
}
```

---

### Account Management

#### POST /api/accounts/load
Load Chart of Accounts into memory for ID lookups

**Request:**
```bash
curl -X POST \
  -F "file=@AccountList.csv" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/accounts/load
```

**Response:**
```json
{
  "success": true,
  "accounts_loaded": 156,
  "message": "Chart of Accounts loaded successfully"
}
```

#### POST /api/accounts/lookup
Look up account ID by name

**Request:**
```bash
curl -X POST \
  -H "Content-Type: application/json" \
  -d '{"name": "Cash"}' \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/accounts/lookup
```

**Response:**
```json
{
  "success": true,
  "account": {
    "id": "35",
    "name": "Cash",
    "type": "Bank",
    "balance": "1000.00"
  }
}
```

---

### Document Conversion Endpoints

All conversion endpoints follow the same pattern:
- **Method**: POST
- **Content-Type**: multipart/form-data
- **Accepts**: CSV, XLSX, PDF files (max 10MB)

#### Optional Form Parameters

- `save_to_db` (string): "true" to save results to database
- `project_id` (string): Required if save_to_db=true
- `source_document_id` (string): Optional document identifier

---

### 1. Chart of Accounts

#### POST /api/convert/accounts

**Request:**
```bash
curl -X POST \
  -F "file=@AccountList.csv" \
  -F "save_to_db=true" \
  -F "project_id=proj_123" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/accounts
```

**Response:**
```json
{
  "success": true,
  "data": [
    {
      "id": "1",
      "name": "Checking",
      "type": "Bank",
      "detail_type": "Checking",
      "balance": "1201.00",
      "currency": "USD"
    }
  ],
  "count": 45,
  "filename": "AccountList.csv",
  "saved_to_db": true,
  "project_id": "proj_123"
}
```

---

### 2. Balance Sheet

#### POST /api/convert/balance-sheet

**Request:**
```bash
curl -X POST \
  -F "file=@BalanceSheet.pdf" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/balance-sheet
```

**Response:**
```json
{
  "success": true,
  "data": [
    {
      "month": "DEC",
      "year": "2025",
      "startDate": "2025-12-01",
      "endDate": "2025-12-31",
      "assets": { ... },
      "liabilities": { ... },
      "equity": { ... }
    }
  ],
  "months": 1,
  "filename": "BalanceSheet.pdf"
}
```

---

### 3. Profit & Loss (Income Statement)

#### POST /api/convert/profit-loss

**Request:**
```bash
curl -X POST \
  -F "file=@ProfitLoss.xlsx" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/profit-loss
```

**Response:**
```json
{
  "success": true,
  "data": [
    {
      "month": "DEC",
      "year": "2025",
      "startDate": "2025-12-01",
      "endDate": "2025-12-31",
      "revenue": { ... },
      "expenses": { ... },
      "netIncome": "15000.00"
    }
  ],
  "months": 1,
  "filename": "ProfitLoss.xlsx"
}
```

---

### 4. Trial Balance

#### POST /api/convert/trial-balance

**Request:**
```bash
curl -X POST \
  -F "file=@TrialBalance.pdf" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/trial-balance
```

**Response:**
```json
{
  "success": true,
  "data": {
    "monthlyReports": [
      {
        "month": "DEC",
        "year": "2025",
        "startDate": "2025-12-01",
        "endDate": "2025-12-31",
        "report": {
          "header": { ... },
          "columns": { ... },
          "rows": {
            "row": [
              {
                "colData": [
                  {"value": "Checking", "id": "1"},
                  {"value": "1201.00"},
                  {"value": ""}
                ]
              }
            ]
          }
        }
      }
    ],
    "summary": {
      "requestStartDate": "2025-12-01",
      "requestEndDate": "2025-12-31",
      "totalMonths": 1,
      "accountingMethod": "Accrual"
    }
  },
  "months": 1,
  "filename": "TrialBalance.pdf"
}
```

---

### 5. Cash Flow Statement

#### POST /api/convert/cash-flow

**Request:**
```bash
curl -X POST \
  -F "file=@CashFlow.csv" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/cash-flow
```

**Response:**
```json
{
  "success": true,
  "data": [
    {
      "month": "DEC",
      "year": "2025",
      "operating": { ... },
      "investing": { ... },
      "financing": { ... },
      "netChange": "5000.00"
    }
  ],
  "months": 1,
  "filename": "CashFlow.csv"
}
```

---

### 6. General Ledger

#### POST /api/convert/general-ledger

**Auto-derives Chart of Accounts**: When you upload a General Ledger with `save_to_db=true`, the system automatically checks if a Chart of Accounts exists for your project. If not, it derives one from the GL transactions and saves it with `source_type='derived_from_gl'`.

**Request:**
```bash
curl -X POST \
  -F "file=@GeneralLedger.csv" \
  -F "save_to_db=true" \
  -F "project_id=proj_123" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/general-ledger
```

**Response:**
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
  "coa_derived": true,
  "derived_count": 30
}
```

**Response Fields:**
- `coa_derived` (boolean): `true` if a Chart of Accounts was automatically derived from the GL
- `derived_count` (number): Number of accounts that were derived (only present if `coa_derived` is true)

**How Auto-Derivation Works:**
1. You upload a General Ledger with `save_to_db=true` and `project_id`
2. System checks if Chart of Accounts exists for the project
3. If NO COA exists → System derives COA from GL transactions using pattern matching
4. Derived COA is saved with `data_type='chart_of_accounts'` and `source_type='derived_from_gl'`
5. Both GL and derived COA are now available in your project

---

### 7. Journal Entries

#### POST /api/convert/journal-entries

**Response Structure:**
```json
{
  "success": true,
  "data": {
    "entries": [ ... ]
  },
  "filename": "JournalEntries.pdf"
}
```

---

### 8. Accounts Payable

#### POST /api/convert/accounts-payable

**Response Structure:**
```json
{
  "success": true,
  "data": [
    {
      "vendor": "ABC Supplies",
      "amount": "1500.00",
      "dueDate": "2025-12-31",
      "aging": "Current"
    }
  ],
  "count": 15,
  "filename": "AP_Aging.csv"
}
```

---

### 9. Accounts Receivable

#### POST /api/convert/accounts-receivable

**Response Structure:**
```json
{
  "success": true,
  "data": [
    {
      "customer": "ABC Corp",
      "amount": "2500.00",
      "dueDate": "2025-12-31",
      "aging": "Current"
    }
  ],
  "count": 20,
  "filename": "AR_Aging.csv"
}
```

---

### 10. Customer Concentration

#### POST /api/convert/customer-concentration

**Response Structure:**
```json
{
  "success": true,
  "data": [
    {
      "customer": "ABC Corp",
      "revenue": "50000.00",
      "percentage": "25.5"
    }
  ],
  "count": 10,
  "filename": "CustomerConcentration.csv"
}
```

---

### 11. Vendor Concentration

#### POST /api/convert/vendor-concentration

**Response Structure:**
```json
{
  "success": true,
  "data": [
    {
      "vendor": "XYZ Supplies",
      "expenses": "30000.00",
      "percentage": "18.2"
    }
  ],
  "count": 12,
  "filename": "VendorConcentration.csv"
}
```

---

## Storage-Based Conversion

Convert documents already stored in Supabase Storage.

**Pattern:** POST `/api/convert-from-storage/{document-type}`

**Request Body:**
```json
{
  "file_path": "projects/proj_123/trial-balance.pdf",
  "project_id": "proj_123",
  "source_document_id": "doc_456"
}
```

**Available Endpoints:**
- `/api/convert-from-storage/accounts`
- `/api/convert-from-storage/balance-sheet`
- `/api/convert-from-storage/profit-loss`
- `/api/convert-from-storage/trial-balance`
- `/api/convert-from-storage/cash-flow`
- `/api/convert-from-storage/general-ledger`
- `/api/convert-from-storage/journal-entries`
- `/api/convert-from-storage/accounts-payable`
- `/api/convert-from-storage/accounts-receivable`
- `/api/convert-from-storage/customer-concentration`
- `/api/convert-from-storage/vendor-concentration`

**Example:**
```bash
curl -X POST \
  -H "Content-Type: application/json" \
  -d '{
    "file_path": "projects/proj_123/trial-balance.pdf",
    "project_id": "proj_123"
  }' \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert-from-storage/trial-balance
```

**Response:**
```json
{
  "success": true,
  "data": { ... },
  "months": 1,
  "file_path": "projects/proj_123/trial-balance.pdf",
  "saved_to_db": true,
  "project_id": "proj_123"
}
```

---

## Batch Processing

Process multiple files at once or ZIP archives containing multiple documents.

### Upload Multiple Files

```bash
curl -X POST \
  -F "files=@TrialBalance_Jan.pdf" \
  -F "files=@TrialBalance_Feb.pdf" \
  -F "files=@TrialBalance_Mar.pdf" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/batch/trial-balance
```

### Upload ZIP Archive

```bash
curl -X POST \
  -F "file=@financial_reports.zip" \
  https://qbtojson-7lqwugl3xa-uc.a.run.app/api/batch/trial-balance
```

### Available Batch Endpoints

#### POST /api/batch/balance-sheet
Process multiple Balance Sheet files

**Response:**
```json
{
  "success": true,
  "data": [ ... ],
  "files_processed": 3,
  "total_months": 3,
  "errors": []
}
```

#### POST /api/batch/profit-loss
Process multiple P&L files

#### POST /api/batch/trial-balance
Process multiple Trial Balance files

#### POST /api/batch/cash-flow
Process multiple Cash Flow files

#### POST /api/batch/general-ledger
Process multiple General Ledger files

#### POST /api/batch/mixed
Process mixed document types

**Response:**
```json
{
  "success": true,
  "results": {
    "trial_balance": { ... },
    "balance_sheet": { ... },
    "profit_loss": { ... }
  },
  "files_processed": 5,
  "errors": []
}
```

---

## Error Handling

### Error Response Format

```json
{
  "error": "Error message",
  "details": "Detailed error information"
}
```

### Common HTTP Status Codes

| Code | Description |
|------|-------------|
| 200  | Success |
| 400  | Bad Request (invalid file, missing parameters) |
| 404  | Endpoint not found |
| 500  | Internal Server Error |

### Common Errors

#### No File Provided
```json
{
  "error": "No file provided"
}
```

#### Invalid File Type
```json
{
  "error": "Invalid file type. Allowed types: csv, xlsx, pdf"
}
```

#### File Too Large
```json
{
  "error": "File too large. Maximum size: 10MB"
}
```

#### Conversion Failed
```json
{
  "error": "Failed to convert file",
  "details": "Could not parse PDF structure"
}
```

---

## File Format Support

### CSV
- **Single-month**: Single month per file
- **Multi-month**: Multiple months side-by-side
- **Encoding**: UTF-8
- **Delimiter**: Comma (,)

### XLSX
- **Single-month**: Single month with "As of [Date]" header
- **Multi-month**: Multiple months in columns
- **Formulas**: Automatically evaluated
- **Format**: Excel 2007+ (.xlsx)

### PDF
- **Single-month**: Single month with "As of [Date]" header
- **Multi-month**: Multiple months in columns
- **Text-based**: Requires extractable text (not scanned images)
- **Structure**: Columnar format with clear headers

---

## Data Types Reference

### Chart of Accounts
```typescript
{
  id: string
  name: string
  type: string
  detail_type: string
  balance: string
  currency: string
}
```

### Monthly Report (Balance Sheet, P&L, etc.)
```typescript
{
  month: string        // "JAN", "FEB", etc.
  year: string        // "2025"
  startDate: string   // "2025-01-01"
  endDate: string     // "2025-01-31"
  // ... report-specific fields
}
```

### Trial Balance
```typescript
{
  monthlyReports: Array<{
    month: string
    year: string
    startDate: string
    endDate: string
    report: {
      header: { ... }
      columns: { ... }
      rows: { row: Array<...> }
    }
  }>
  summary: {
    requestStartDate: string
    requestEndDate: string
    totalMonths: number
    accountingMethod: string
  }
}
```

---

## Best Practices

### 1. File Upload
- **Validate file size** before upload (max 10MB)
- **Check file extension** (.csv, .xlsx, .pdf)
- **Use secure_filename** to sanitize filenames
- **Handle errors gracefully** with try-catch

### 2. Batch Processing
- **Limit batch size** to reasonable number of files (< 50)
- **Use ZIP archives** for large batches
- **Monitor processing time**
- **Handle partial failures**

### 3. Database Integration
- **Set save_to_db=true** only when needed
- **Always provide project_id** when saving
- **Use source_document_id** for traceability
- **Check saved_to_db** in response

### 4. Error Handling
- **Check success field** in response
- **Log error details** for debugging
- **Implement retry logic** for transient failures
- **Validate input** before API call

### 5. Performance
- **Cache Chart of Accounts** using /api/accounts/load
- **Use storage-based conversion** for files already uploaded
- **Batch similar documents** together
- **Monitor API response times**

---

## Code Examples

See [FRONTEND_INTEGRATION.md](./FRONTEND_INTEGRATION.md) for detailed code examples in:
- JavaScript/TypeScript
- React
- Python
- cURL

---

## Support

For issues or questions:
- Check the `/health` and `/api/info` endpoints
- Review error messages in responses
- Consult this documentation
- Test with sample files from `sampleReports/` directory

---

## Changelog

### v1.0.0 (Current)
- ✅ 11 document types supported
- ✅ PDF, CSV, XLSX formats
- ✅ Batch processing
- ✅ Storage-based conversion
- ✅ Database integration
- ✅ Single-month trial balance parser (PDF & XLSX)

---

## API Version

**Current Version:** 1.0.0  
**Last Updated:** January 7, 2026  
**Service URL:** https://qbtojson-7lqwugl3xa-uc.a.run.app
