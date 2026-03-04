# qbToJson - CLAUDE.MD

## Overview

Transforms QuickBooks API data into JSON format for storage in Supabase. Handles data normalization and period detection.

## Tech Stack

- **Language**: Python 3.x
- **Deployment**: Cloud Run
- **Database**: Supabase

## Key Files

| File | Purpose |
|------|---------|
| `API_DOCUMENTATION.md` | API reference |
| `CHART_OF_ACCOUNTS_DERIVATION.md` | Account classification logic |
| `AUTO_DERIVE_DEPLOYMENT.md` | Auto-derive feature |
| `DATABASE_INTEGRATION_SUMMARY.md` | Supabase integration |
| `CONVERTER_SUMMARY.md` | Conversion logic overview |

## Commands

```bash
# Setup
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Run locally
python main.py

# Deploy
gcloud run deploy qb-to-json --source .
```

## Data Transformation Flow

```
QuickBooks API Response
    │
    ▼
Parse nested row structure
    │
    ▼
Flatten to account/period/amount
    │
    ▼
Derive FS Line Items (Column E)
    │
    ▼
Store in Supabase processed_data
```

## Output Format

```json
{
  "data_type": "trial_balance",
  "data": {
    "accounts": [
      {
        "Id": "123",
        "Name": "Cash",
        "AccountType": "Bank",
        "Classification": "Asset"
      }
    ],
    "monthlyBalances": {
      "2024-01-31": { "123": 50000, "124": 12000 }
    }
  }
}
```

## Related Repos

- `quickbooksapi` - Data source
- `shepi-ai-web` - Data consumer
- `cloudRunQBAuth` - Token management
