# Frontend Integration Guide

Complete guide for integrating the Document Converter API into your frontend application.

## Table of Contents

1. [JavaScript/TypeScript Examples](#javascripttypescript-examples)
2. [React Integration](#react-integration)
3. [TypeScript Types](#typescript-types)
4. [Error Handling](#error-handling)
5. [File Upload Best Practices](#file-upload-best-practices)
6. [Example Applications](#example-applications)

---

## JavaScript/TypeScript Examples

### Basic File Upload

```javascript
async function convertDocument(file, documentType, apiKey) {
  const formData = new FormData();
  formData.append('file', file);
  
  try {
    const response = await fetch(
      `https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/${documentType}`,
      {
        method: 'POST',
        headers: {
          'x-api-key': apiKey  // REQUIRED: API key authentication
        },
        body: formData
      }
    );
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    
    if (!result.success) {
      throw new Error(result.error || 'Conversion failed');
    }
    
    return result.data;
  } catch (error) {
    console.error('Error converting document:', error);
    throw error;
  }
}

// Usage
const fileInput = document.getElementById('fileInput');
const file = fileInput.files[0];
const apiKey = 'YOUR_API_KEY';  // Store securely, don't hardcode in production
const data = await convertDocument(file, 'trial-balance', apiKey);
console.log('Converted data:', data);
```

**⚠️ Security Note:** Never expose your API key in client-side code in production. Instead:
- Store API key in environment variables
- Make requests through your backend
- Or use a secure token exchange mechanism

### With Database Save

```javascript
async function convertAndSave(file, documentType, projectId, apiKey) {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('save_to_db', 'true');
  formData.append('project_id', projectId);
  
  const response = await fetch(
    `https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/${documentType}`,
    {
      method: 'POST',
      headers: {
        'x-api-key': apiKey
      },
      body: formData
    }
  );
  
  const result = await response.json();
  
  if (result.saved_to_db) {
    console.log('Successfully saved to database');
  }
  
  return result.data;
}
```

### Storage-Based Conversion

```javascript
async function convertFromStorage(filePath, projectId, documentType, apiKey) {
  const response = await fetch(
    `https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert-from-storage/${documentType}`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey
      },
      body: JSON.stringify({
        file_path: filePath,
        project_id: projectId
      })
    }
  );
  
  const result = await response.json();
  return result.data;
}

// Usage
const data = await convertFromStorage(
  'projects/proj_123/trial-balance.pdf',
  'proj_123',
  'trial-balance'
);
```

### Batch Upload

```javascript
async function batchConvert(files, documentType) {
  const formData = new FormData();
  
  // Append multiple files
  files.forEach(file => {
    formData.append('files', file);
  });
  
  const response = await fetch(
    `https://qbtojson-7lqwugl3xa-uc.a.run.app/api/batch/${documentType}`,
    {
      method: 'POST',
      body: formData
    }
  );
  
  const result = await response.json();
  return result;
}

// Usage with file input
const fileInput = document.getElementById('multiFileInput');
const files = Array.from(fileInput.files);
const apiKey = 'YOUR_API_KEY';
const result = await batchConvert(files, 'trial-balance', apiKey);
console.log(`Processed ${result.files_processed} files`);
```

### Batch Upload (with API Key)

```javascript
async function batchConvert(files, documentType, apiKey) {
  const formData = new FormData();
  
  // Append multiple files
  files.forEach(file => {
    formData.append('files', file);
  });
  
  const response = await fetch(
    `https://qbtojson-7lqwugl3xa-uc.a.run.app/api/batch/${documentType}`,
    {
      method: 'POST',
      headers: {
        'x-api-key': apiKey
      },
      body: formData
    }
  );
  
  const result = await response.json();
  return result;
}
```

### Account Lookup

```javascript
async function lookupAccount(accountName, apiKey) {
  const response = await fetch(
    'https://qbtojson-7lqwugl3xa-uc.a.run.app/api/accounts/lookup',
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey
      },
      body: JSON.stringify({ name: accountName })
    }
  );
  
  const result = await response.json();
  
  if (result.success) {
    return result.account;
  }
  
  return null;
}

// Usage
const account = await lookupAccount('Cash');
if (account) {
  console.log(`Account ID: ${account.id}`);
}
```

### General Ledger with Auto-Derive COA

```javascript
async function uploadGeneralLedgerWithAutoCOA(file, projectId, apiKey) {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('save_to_db', 'true');
  formData.append('project_id', projectId);
  
  const response = await fetch(
    'https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/general-ledger',
    {
      method: 'POST',
      headers: {
        'x-api-key': apiKey
      },
      body: formData
    }
  );
  
  const result = await response.json();
  
  if (result.success) {
    console.log('General Ledger processed:', result.data);
    
    // Check if COA was auto-derived
    if (result.coa_derived) {
      console.log(`✅ Auto-derived ${result.derived_count} accounts from General Ledger`);
      console.log('⚠️ Please review derived accounts against Balance Sheet and P&L');
    }
  }
  
  return result;
}

// Usage
const glFile = document.getElementById('fileInput').files[0];
const apiKey = 'YOUR_API_KEY';
const result = await uploadGeneralLedgerWithAutoCOA(glFile, 'proj_123', apiKey);
```

---

## React Integration

### Custom Hook for File Upload

```typescript
// useDocumentConverter.ts
import { useState } from 'react';

interface ConversionResult {
  success: boolean;
  data: any;
  filename?: string;
  months?: number;
  count?: number;
  saved_to_db?: boolean;
}

interface UseDocumentConverterOptions {
  apiBaseUrl?: string;
  onSuccess?: (result: ConversionResult) => void;
  onError?: (error: Error) => void;
}

export function useDocumentConverter(options: UseDocumentConverterOptions = {}) {
  const {
    apiBaseUrl = 'https://qbtojson-7lqwugl3xa-uc.a.run.app',
    onSuccess,
    onError
  } = options;
  
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<Error | null>(null);
  const [result, setResult] = useState<ConversionResult | null>(null);
  
  const convert = async (
    file: File,
    documentType: string,
    options?: {
      saveToDb?: boolean;
      projectId?: string;
      sourceDocumentId?: string;
    }
  ) => {
    setLoading(true);
    setError(null);
    
    try {
      const formData = new FormData();
      formData.append('file', file);
      
      if (options?.saveToDb) {
        formData.append('save_to_db', 'true');
        if (options.projectId) {
          formData.append('project_id', options.projectId);
        }
        if (options.sourceDocumentId) {
          formData.append('source_document_id', options.sourceDocumentId);
        }
      }
      
      const response = await fetch(
        `${apiBaseUrl}/api/convert/${documentType}`,
        {
          method: 'POST',
          body: formData
        }
      );
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      
      const data = await response.json();
      
      if (!data.success) {
        throw new Error(data.error || 'Conversion failed');
      }
      
      setResult(data);
      onSuccess?.(data);
      
      return data;
    } catch (err) {
      const error = err instanceof Error ? err : new Error('Unknown error');
      setError(error);
      onError?.(error);
      throw error;
    } finally {
      setLoading(false);
    }
  };
  
  const reset = () => {
    setLoading(false);
    setError(null);
    setResult(null);
  };
  
  return {
    convert,
    loading,
    error,
    result,
    reset
  };
}
```

### React Component Example

```typescript
// DocumentUploader.tsx
import React, { useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import { useDocumentConverter } from './useDocumentConverter';

interface DocumentUploaderProps {
  documentType: string;
  projectId?: string;
  onComplete?: (data: any) => void;
}

export function DocumentUploader({
  documentType,
  projectId,
  onComplete
}: DocumentUploaderProps) {
  const { convert, loading, error, result } = useDocumentConverter({
    onSuccess: (result) => {
      console.log('Conversion successful:', result);
      onComplete?.(result.data);
    },
    onError: (error) => {
      console.error('Conversion failed:', error);
    }
  });
  
  const onDrop = useCallback(
    async (acceptedFiles: File[]) => {
      if (acceptedFiles.length === 0) return;
      
      const file = acceptedFiles[0];
      await convert(file, documentType, {
        saveToDb: !!projectId,
        projectId
      });
    },
    [convert, documentType, projectId]
  );
  
  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'text/csv': ['.csv'],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/pdf': ['.pdf']
    },
    maxSize: 10 * 1024 * 1024, // 10MB
    multiple: false
  });
  
  return (
    <div className="document-uploader">
      <div
        {...getRootProps()}
        className={`dropzone ${isDragActive ? 'active' : ''}`}
      >
        <input {...getInputProps()} />
        {loading ? (
          <p>Converting...</p>
        ) : (
          <p>
            {isDragActive
              ? 'Drop the file here'
              : 'Drag & drop a file here, or click to select'}
          </p>
        )}
      </div>
      
      {error && (
        <div className="error">
          <p>Error: {error.message}</p>
        </div>
      )}
      
      {result && (
        <div className="success">
          <p>✅ Conversion successful!</p>
          <p>File: {result.filename}</p>
          {result.months && <p>Months: {result.months}</p>}
          {result.count && <p>Records: {result.count}</p>}
          {result.saved_to_db && <p>✅ Saved to database</p>}
        </div>
      )}
    </div>
  );
}
```

### Batch Upload Component

```typescript
// BatchUploader.tsx
import React, { useState } from 'react';

interface BatchUploaderProps {
  documentType: string;
  onComplete?: (results: any) => void;
}

export function BatchUploader({ documentType, onComplete }: BatchUploaderProps) {
  const [files, setFiles] = useState<File[]>([]);
  const [loading, setLoading] = useState(false);
  const [results, setResults] = useState<any>(null);
  const [error, setError] = useState<string | null>(null);
  
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setFiles(Array.from(e.target.files));
    }
  };
  
  const handleUpload = async () => {
    if (files.length === 0) return;
    
    setLoading(true);
    setError(null);
    
    try {
      const formData = new FormData();
      files.forEach(file => {
        formData.append('files', file);
      });
      
      const response = await fetch(
        `https://qbtojson-7lqwugl3xa-uc.a.run.app/api/batch/${documentType}`,
        {
          method: 'POST',
          body: formData
        }
      );
      
      const data = await response.json();
      
      if (!data.success) {
        throw new Error(data.error || 'Batch processing failed');
      }
      
      setResults(data);
      onComplete?.(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unknown error');
    } finally {
      setLoading(false);
    }
  };
  
  return (
    <div className="batch-uploader">
      <input
        type="file"
        multiple
        onChange={handleFileChange}
        accept=".csv,.xlsx,.pdf"
      />
      
      <div className="file-list">
        {files.map((file, index) => (
          <div key={index} className="file-item">
            {file.name} ({(file.size / 1024).toFixed(2)} KB)
          </div>
        ))}
      </div>
      
      <button onClick={handleUpload} disabled={loading || files.length === 0}>
        {loading ? 'Processing...' : `Upload ${files.length} file(s)`}
      </button>
      
      {error && <div className="error">{error}</div>}
      
      {results && (
        <div className="results">
          <p>✅ Processed {results.files_processed} files</p>
          {results.total_months && <p>Total months: {results.total_months}</p>}
          {results.errors?.length > 0 && (
            <div className="errors">
              <p>Errors:</p>
              <ul>
                {results.errors.map((err: string, i: number) => (
                  <li key={i}>{err}</li>
                ))}
              </ul>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
```

---

## TypeScript Types

```typescript
// types/api.ts

// Base response
export interface ApiResponse<T> {
  success: boolean;
  data: T;
  error?: string;
  details?: string;
}

// Account types
export interface Account {
  id: string;
  name: string;
  type: string;
  detail_type: string;
  balance: string;
  currency: string;
}

export interface AccountLookupResponse {
  success: boolean;
  account?: Account;
  fuzzy_match?: boolean;
  error?: string;
}

// Trial Balance types
export interface TrialBalanceRow {
  id: string | null;
  parentId: string | null;
  header: any | null;
  rows: any | null;
  summary: any | null;
  colData: Array<{
    attributes: any | null;
    value: string;
    id: string | null;
    href: string | null;
  }>;
  type: string | null;
  group: string | null;
}

export interface TrialBalanceReport {
  header: {
    time: string;
    reportName: string;
    reportBasis: string;
    startPeriod: string;
    endPeriod: string;
    summarizeColumnsBy: string;
    currency: string;
    option: Array<{ name: string; value: string }>;
  };
  columns: any;
  rows: {
    row: TrialBalanceRow[];
  };
}

export interface TrialBalanceMonthly {
  month: string;
  year: string;
  startDate: string;
  endDate: string;
  report: TrialBalanceReport;
}

export interface TrialBalanceData {
  monthlyReports: TrialBalanceMonthly[];
  summary: {
    requestStartDate: string;
    requestEndDate: string;
    totalMonths: number;
    accountingMethod: string;
  };
}

export interface TrialBalanceResponse extends ApiResponse<TrialBalanceData> {
  months: number;
  filename: string;
  saved_to_db?: boolean;
  project_id?: string;
}

// Balance Sheet types
export interface BalanceSheetMonthly {
  month: string;
  year: string;
  startDate: string;
  endDate: string;
  assets: any;
  liabilities: any;
  equity: any;
}

export interface BalanceSheetResponse extends ApiResponse<BalanceSheetMonthly[]> {
  months: number;
  filename: string;
  saved_to_db?: boolean;
  project_id?: string;
}

// P&L types
export interface ProfitLossMonthly {
  month: string;
  year: string;
  startDate: string;
  endDate: string;
  revenue: any;
  expenses: any;
  netIncome: string;
}

export interface ProfitLossResponse extends ApiResponse<ProfitLossMonthly[]> {
  months: number;
  filename: string;
  saved_to_db?: boolean;
  project_id?: string;
}

// Cash Flow types
export interface CashFlowMonthly {
  month: string;
  year: string;
  operating: any;
  investing: any;
  financing: any;
  netChange: string;
}

export interface CashFlowResponse extends ApiResponse<CashFlowMonthly[]> {
  months: number;
  filename: string;
  saved_to_db?: boolean;
  project_id?: string;
}

// AR/AP types
export interface AgingItem {
  customer?: string;
  vendor?: string;
  amount: string;
  dueDate: string;
  aging: string;
}

export interface AgingResponse extends ApiResponse<AgingItem[]> {
  count: number;
  filename: string;
  saved_to_db?: boolean;
  project_id?: string;
}

// Concentration types
export interface ConcentrationItem {
  customer?: string;
  vendor?: string;
  revenue?: string;
  expenses?: string;
  percentage: string;
}

export interface ConcentrationResponse extends ApiResponse<ConcentrationItem[]> {
  count: number;
  filename: string;
  saved_to_db?: boolean;
  project_id?: string;
}

// Batch processing types
export interface BatchResponse {
  success: boolean;
  data: any[];
  files_processed: number;
  total_months?: number;
  errors: string[];
}

// Storage conversion types
export interface StorageConversionRequest {
  file_path: string;
  project_id: string;
  source_document_id?: string;
}

// Document type union
export type DocumentType =
  | 'accounts'
  | 'balance-sheet'
  | 'profit-loss'
  | 'trial-balance'
  | 'cash-flow'
  | 'general-ledger'
  | 'journal-entries'
  | 'accounts-payable'
  | 'accounts-receivable'
  | 'customer-concentration'
  | 'vendor-concentration';
```

---

## Error Handling

### Comprehensive Error Handler

```typescript
class APIError extends Error {
  constructor(
    message: string,
    public statusCode: number,
    public details?: string
  ) {
    super(message);
    this.name = 'APIError';
  }
}

async function handleAPIResponse<T>(response: Response): Promise<T> {
  const contentType = response.headers.get('content-type');
  
  if (!contentType?.includes('application/json')) {
    throw new APIError(
      'Invalid response format',
      response.status,
      'Expected JSON response'
    );
  }
  
  const data = await response.json();
  
  if (!response.ok) {
    throw new APIError(
      data.error || 'API request failed',
      response.status,
      data.details
    );
  }
  
  if (!data.success) {
    throw new APIError(
      data.error || 'Operation failed',
      response.status,
      data.details
    );
  }
  
  return data;
}

// Usage
try {
  const response = await fetch(apiUrl, options);
  const result = await handleAPIResponse(response);
  console.log('Success:', result);
} catch (error) {
  if (error instanceof APIError) {
    console.error(`API Error (${error.statusCode}):`, error.message);
    if (error.details) {
      console.error('Details:', error.details);
    }
  } else {
    console.error('Unexpected error:', error);
  }
}
```

### React Error Boundary

```typescript
import React, { Component, ErrorInfo, ReactNode } from 'react';

interface Props {
  children: ReactNode;
  fallback?: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends Component<Props, State> {
  constructor(props: Props) {
    super(props);
    this.state = { hasError: false, error: null };
  }
  
  static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }
  
  componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error('Error caught by boundary:', error, errorInfo);
  }
  
  render() {
    if (this.state.hasError) {
      return (
        this.props.fallback || (
          <div className="error-boundary">
            <h2>Something went wrong</h2>
            <p>{this.state.error?.message}</p>
            <button onClick={() => this.setState({ hasError: false, error: null })}>
              Try again
            </button>
          </div>
        )
      );
    }
    
    return this.props.children;
  }
}
```

---

## File Upload Best Practices

### File Validation

```typescript
interface FileValidation {
  maxSize: number;
  allowedTypes: string[];
}

function validateFile(
  file: File,
  validation: FileValidation
): { valid: boolean; error?: string } {
  // Check file size
  if (file.size > validation.maxSize) {
    return {
      valid: false,
      error: `File too large. Maximum size: ${validation.maxSize / 1024 / 1024}MB`
    };
  }
  
  // Check file type
  const extension = file.name.split('.').pop()?.toLowerCase();
  if (!extension || !validation.allowedTypes.includes(extension)) {
    return {
      valid: false,
      error: `Invalid file type. Allowed: ${validation.allowedTypes.join(', ')}`
    };
  }
  
  return { valid: true };
}

// Usage
const validation: FileValidation = {
  maxSize: 10 * 1024 * 1024, // 10MB
  allowedTypes: ['csv', 'xlsx', 'pdf']
};

const result = validateFile(file, validation);
if (!result.valid) {
  alert(result.error);
  return;
}
```

### Progress Tracking

```typescript
function uploadWithProgress(
  file: File,
  url: string,
  onProgress: (percent: number) => void
): Promise<any> {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    
    // Track upload progress
    xhr.upload.addEventListener('progress', (e) => {
      if (e.lengthComputable) {
        const percent = (e.loaded / e.total) * 100;
        onProgress(percent);
      }
    });
    
    xhr.addEventListener('load', () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        resolve(JSON.parse(xhr.responseText));
      } else {
        reject(new Error(`HTTP ${xhr.status}: ${xhr.statusText}`));
      }
    });
    
    xhr.addEventListener('error', () => {
      reject(new Error('Network error'));
    });
    
    xhr.open('POST', url);
    
    const formData = new FormData();
    formData.append('file', file);
    
    xhr.send(formData);
  });
}

// Usage
await uploadWithProgress(
  file,
  'https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/trial-balance',
  (percent) => {
    console.log(`Upload progress: ${percent.toFixed(0)}%`);
  }
);
```

### Retry Logic

```typescript
async function fetchWithRetry<T>(
  url: string,
  options: RequestInit,
  maxRetries = 3,
  delay = 1000
): Promise<T> {
  let lastError: Error;
  
  for (let i = 0; i < maxRetries; i++) {
    try {
      const response = await fetch(url, options);
      return await handleAPIResponse<T>(response);
    } catch (error) {
      lastError = error instanceof Error ? error : new Error('Unknown error');
      
      // Don't retry on client errors (4xx)
      if (error instanceof APIError && error.statusCode >= 400 && error.statusCode < 500) {
        throw error;
      }
      
      // Wait before retrying
      if (i < maxRetries - 1) {
        await new Promise(resolve => setTimeout(resolve, delay * (i + 1)));
      }
    }
  }
  
  throw lastError!;
}
```

---

## Example Applications

### Complete React App

```typescript
// App.tsx
import React, { useState } from 'react';
import { DocumentUploader } from './components/DocumentUploader';
import { ErrorBoundary } from './components/ErrorBoundary';
import type { TrialBalanceData, DocumentType } from './types/api';

function App() {
  const [documentType, setDocumentType] = useState<DocumentType>('trial-balance');
  const [projectId] = useState('proj_123');
  const [convertedData, setConvertedData] = useState<any>(null);
  
  const handleComplete = (data: any) => {
    setConvertedData(data);
    console.log('Converted data:', data);
  };
  
  return (
    <ErrorBoundary>
      <div className="app">
        <h1>Document Converter</h1>
        
        <div className="controls">
          <label>
            Document Type:
            <select
              value={documentType}
              onChange={(e) => setDocumentType(e.target.value as DocumentType)}
            >
              <option value="trial-balance">Trial Balance</option>
              <option value="balance-sheet">Balance Sheet</option>
              <option value="profit-loss">Profit & Loss</option>
              <option value="cash-flow">Cash Flow</option>
              <option value="accounts">Chart of Accounts</option>
            </select>
          </label>
        </div>
        
        <DocumentUploader
          documentType={documentType}
          projectId={projectId}
          onComplete={handleComplete}
        />
        
        {convertedData && (
          <div className="results">
            <h2>Converted Data</h2>
            <pre>{JSON.stringify(convertedData, null, 2)}</pre>
          </div>
        )}
      </div>
    </ErrorBoundary>
  );
}

export default App;
```

### Vue.js Example

```vue
<!-- DocumentUploader.vue -->
<template>
  <div class="document-uploader">
    <input
      type="file"
      ref="fileInput"
      @change="handleFileChange"
      accept=".csv,.xlsx,.pdf"
    />
    
    <button @click="upload" :disabled="!file || loading">
      {{ loading ? 'Converting...' : 'Convert Document' }}
    </button>
    
    <div v-if="error" class="error">
      {{ error }}
    </div>
    
    <div v-if="result" class="success">
      <p>✅ Conversion successful!</p>
      <p>File: {{ result.filename }}</p>
      <p v-if="result.months">Months: {{ result.months }}</p>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue';

const props = defineProps<{
  documentType: string;
  projectId?: string;
}>();

const emit = defineEmits<{
  (e: 'complete', data: any): void;
}>();

const file = ref<File | null>(null);
const loading = ref(false);
const error = ref<string | null>(null);
const result = ref<any>(null);

const handleFileChange = (event: Event) => {
  const target = event.target as HTMLInputElement;
  file.value = target.files?.[0] || null;
  error.value = null;
  result.value = null;
};

const upload = async () => {
  if (!file.value) return;
  
  loading.value = true;
  error.value = null;
  
  try {
    const formData = new FormData();
    formData.append('file', file.value);
    
    if (props.projectId) {
      formData.append('save_to_db', 'true');
      formData.append('project_id', props.projectId);
    }
    
    const response = await fetch(
      `https://qbtojson-7lqwugl3xa-uc.a.run.app/api/convert/${props.documentType}`,
      {
        method: 'POST',
        body: formData
      }
    );
    
    const data = await response.json();
    
    if (!data.success) {
      throw new Error(data.error || 'Conversion failed');
    }
    
    result.value = data;
    emit('complete', data.data);
  } catch (err) {
    error.value = err instanceof Error ? err.message : 'Unknown error';
  } finally {
    loading.value = false;
  }
};
</script>
```

---

## Testing

### Jest Test Example

```typescript
// documentConverter.test.ts
import { convertDocument } from './documentConverter';

describe('Document Converter', () => {
  it('should convert a valid file', async () => {
    const mockFile = new File(['test'], 'test.csv', { type: 'text/csv' });
    
    global.fetch = jest.fn(() =>
      Promise.resolve({
        ok: true,
        json: () => Promise.resolve({
          success: true,
          data: { monthlyReports: [] },
          months: 1
        })
      })
    ) as jest.Mock;
    
    const result = await convertDocument(mockFile, 'trial-balance');
    
    expect(result).toBeDefined();
    expect(result.monthlyReports).toBeDefined();
  });
  
  it('should handle errors gracefully', async () => {
    const mockFile = new File(['test'], 'test.csv', { type: 'text/csv' });
    
    global.fetch = jest.fn(() =>
      Promise.resolve({
        ok: false,
        status: 400,
        json: () => Promise.resolve({
          error: 'Invalid file'
        })
      })
    ) as jest.Mock;
    
    await expect(convertDocument(mockFile, 'trial-balance')).rejects.toThrow();
  });
});
```

---

## Additional Resources

- [Backend API Documentation](./API_DOCUMENTATION.md)
- [Sample Reports](./sampleReports/)
- [Sample JSON Outputs](./sampleJson/)

For questions or issues, refer to the main API documentation or check the `/health` endpoint for API status.
