"""
Account Inference Converter
Derives Chart of Accounts from General Ledger data when CoA is not available.
Matches qbToJson output format exactly.
"""

import re
from typing import Dict, List, Tuple
from datetime import datetime


class AccountInferenceEngine:
    """Derives Chart of Accounts from General Ledger transactions"""
    
    # Classification patterns for account name matching
    CLASSIFICATION_PATTERNS = {
        'ASSET': [
            (r'\b(cash|checking|savings|bank)\b', 0.95),
            (r'\b(receivable|a/r)\b', 0.95),
            (r'\b(inventory|stock)\b', 0.90),
            (r'\b(equipment|vehicle|truck|machinery)\b', 0.90),
            (r'\b(prepaid|deferred)\b', 0.85),
            (r'\b(undeposited funds)\b', 0.95),
        ],
        'LIABILITY': [
            (r'\b(payable|a/p)\b', 0.95),
            (r'\b(credit card|mastercard|visa|amex)\b', 0.90),
            (r'\b(loan|note payable|debt)\b', 0.90),
            (r'\b(payroll|wages payable)\b', 0.85),
        ],
        'EQUITY': [
            (r'\b(equity|capital)\b', 0.95),
            (r'\b(retained earnings)\b', 0.95),
            (r'\b(opening balance equity)\b', 0.95),
            (r'\b(owner.?s draw|distribution)\b', 0.90),
        ],
        'REVENUE': [
            (r'\b(revenue|sales|income)\b', 0.90),
            (r'\b(service|consulting|design)\b', 0.85),
            (r'\b(landscaping|pest control)\b', 0.85),
        ],
        'EXPENSE': [
            (r'\b(expense|cost)\b', 0.90),
            (r'\b(wages|salary|payroll)\b', 0.90),
            (r'\b(rent|lease)\b', 0.90),
            (r'\b(utilities|telephone|gas|electric)\b', 0.90),
            (r'\b(advertising|marketing)\b', 0.90),
            (r'\b(insurance|legal|professional)\b', 0.90),
            (r'\b(maintenance|repair)\b', 0.90),
            (r'\b(supplies|materials|fuel)\b', 0.90),
            (r'\b(cogs|cost of goods sold)\b', 0.95),
        ],
    }
    
    # AccountType mapping based on classification and name patterns
    ACCOUNT_TYPE_PATTERNS = {
        'ASSET': {
            r'\b(cash|checking|savings|bank)\b': 'BANK',
            r'\b(receivable|a/r)\b': 'ACCOUNTS_RECEIVABLE',
            r'\b(inventory|stock)\b': 'OTHER_CURRENT_ASSET',
            r'\b(equipment|vehicle|truck|machinery)\b': 'FIXED_ASSET',
            r'\b(prepaid)\b': 'OTHER_CURRENT_ASSET',
            r'\b(undeposited)\b': 'OTHER_CURRENT_ASSET',
            'default': 'OTHER_CURRENT_ASSET'
        },
        'LIABILITY': {
            r'\b(payable|a/p)\b': 'ACCOUNTS_PAYABLE',
            r'\b(credit card|mastercard|visa)\b': 'CREDIT_CARD',
            r'\b(loan|note)\b': 'LONG_TERM_LIABILITY',
            'default': 'OTHER_CURRENT_LIABILITY'
        },
        'EQUITY': {
            r'\b(retained)\b': 'EQUITY',
            r'\b(opening)\b': 'EQUITY',
            'default': 'EQUITY'
        },
        'REVENUE': {
            'default': 'INCOME'
        },
        'EXPENSE': {
            r'\b(cogs|cost of goods)\b': 'COST_OF_GOODS_SOLD',
            'default': 'EXPENSE'
        }
    }
    
    def extract_accounts_from_gl(self, gl_data: Dict) -> Dict[str, Dict]:
        """
        Extract unique accounts from General Ledger data.
        Returns dict of {account_id: {name, stats}}
        """
        accounts = {}
        
        # Navigate the GL structure
        rows = gl_data.get('rows', {}).get('row', [])
        
        for row in rows:
            # Each account is in a header with nested transaction rows
            header = row.get('header', {})
            if header:
                col_data = header.get('colData', [])
                if col_data and len(col_data) > 0:
                    account_id = col_data[0].get('id')
                    account_name = col_data[0].get('value')
                    
                    if account_id and account_name:
                        # Initialize account if not seen before
                        if account_id not in accounts:
                            accounts[account_id] = {
                                'id': account_id,
                                'name': account_name,
                                'transaction_count': 0,
                                'total_debits': 0.0,
                                'total_credits': 0.0
                            }
                        
                        # Count transactions in nested rows
                        nested_rows = row.get('rows', {}).get('row', [])
                        accounts[account_id]['transaction_count'] += len(nested_rows)
                        
                        # Calculate totals from transactions
                        for txn in nested_rows:
                            txn_data = txn.get('colData', [])
                            if len(txn_data) >= 8:  # GL has 8 columns
                                amount_str = txn_data[6].get('value', '0')
                                try:
                                    amount = float(amount_str.replace(',', ''))
                                    if amount > 0:
                                        accounts[account_id]['total_debits'] += amount
                                    else:
                                        accounts[account_id]['total_credits'] += abs(amount)
                                except (ValueError, AttributeError):
                                    pass
        
        return accounts
    
    def infer_classification(self, account_name: str, stats: Dict) -> Tuple[str, float]:
        """
        Infer classification from account name using pattern matching.
        Returns (classification, confidence)
        """
        name_lower = account_name.lower()
        best_match = None
        best_confidence = 0.0
        
        # Try pattern matching first
        for classification, patterns in self.CLASSIFICATION_PATTERNS.items():
            for pattern, confidence in patterns:
                if re.search(pattern, name_lower, re.IGNORECASE):
                    if confidence > best_confidence:
                        best_match = classification
                        best_confidence = confidence
        
        # If found high-confidence match, return it
        if best_match and best_confidence >= 0.85:
            return best_match, best_confidence
        
        # Fallback: use natural balance
        debits = stats.get('total_debits', 0)
        credits = stats.get('total_credits', 0)
        
        if debits > credits * 1.5:
            # Debit balance suggests ASSET or EXPENSE
            return best_match or 'ASSET', 0.50
        elif credits > debits * 1.5:
            # Credit balance suggests LIABILITY, EQUITY, or REVENUE
            return best_match or 'LIABILITY', 0.50
        else:
            # Mixed or unclear
            return best_match or 'ASSET', 0.30
    
    def infer_account_type(self, classification: str, account_name: str) -> str:
        """Infer QuickBooks accountType from classification and name"""
        name_lower = account_name.lower()
        patterns = self.ACCOUNT_TYPE_PATTERNS.get(classification, {})
        
        for pattern, account_type in patterns.items():
            if pattern == 'default':
                continue
            if re.search(pattern, name_lower, re.IGNORECASE):
                return account_type
        
        return patterns.get('default', 'OTHER_CURRENT_ASSET')
    
    def calculate_current_balance(self, stats: Dict) -> float:
        """Calculate current balance from debits and credits"""
        debits = stats.get('total_debits', 0)
        credits = stats.get('total_credits', 0)
        return debits - credits
    
    def create_derived_coa(self, accounts: Dict[str, Dict]) -> List[Dict]:
        """
        Create Chart of Accounts in qbToJson format.
        Matches the exact structure from accounts.json.
        """
        derived_accounts = []
        current_time = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S.000+00:00")
        
        for account_id, stats in accounts.items():
            account_name = stats['name']
            
            # Infer classification
            classification, confidence = self.infer_classification(account_name, stats)
            
            # Infer accountType
            account_type = self.infer_account_type(classification, account_name)
            
            # Calculate balance
            current_balance = self.calculate_current_balance(stats)
            
            # Build account in qbToJson format (Option B: no extra metadata)
            account = {
                "id": account_id,
                "syncToken": "0",
                "metaData": {
                    "createdByRef": None,
                    "createTime": current_time,
                    "lastModifiedByRef": None,
                    "lastUpdatedTime": current_time,
                    "lastChangedInQB": None,
                    "synchronized": None
                },
                "customField": [],
                "attachableRef": [],
                "domain": "QBO",
                "status": None,
                "sparse": False,
                "name": account_name,
                "subAccount": False,
                "parentRef": None,
                "description": None,
                "fullyQualifiedName": account_name,
                "accountAlias": None,
                "txnLocationType": None,
                "active": True,
                "classification": classification,
                "accountType": account_type,
                "accountSubType": None,
                "accountPurposes": [],
                "acctNum": None,
                "acctNumExtn": None,
                "bankNum": None,
                "openingBalance": None,
                "openingBalanceDate": None,
                "currentBalance": round(current_balance, 2),
                "currentBalanceWithSubAccounts": round(current_balance, 2),
                "currencyRef": {
                    "value": "USD",
                    "name": "United States Dollar",
                    "type": None
                },
                "taxAccount": None,
                "taxCodeRef": None,
                "onlineBankingEnabled": None,
                "journalCodeRef": None,
                "accountEx": None,
                "finame": None
            }
            
            derived_accounts.append(account)
        
        # Sort by account ID for consistency
        derived_accounts.sort(key=lambda x: int(x['id']) if x['id'].isdigit() else x['id'])
        
        return derived_accounts
    
    def derive_chart_of_accounts(self, gl_data: Dict) -> List[Dict]:
        """
        Main method: derive complete Chart of Accounts from General Ledger.
        Returns list of accounts in qbToJson format.
        """
        # Step 1: Extract unique accounts
        accounts = self.extract_accounts_from_gl(gl_data)
        
        # Step 2: Infer classifications and create CoA
        derived_coa = self.create_derived_coa(accounts)
        
        return derived_coa


def convert_general_ledger_to_coa(gl_json_data: Dict) -> List[Dict]:
    """
    Convenience function to derive Chart of Accounts from General Ledger.
    
    Args:
        gl_json_data: General Ledger data from qbToJson
    
    Returns:
        List of accounts in qbToJson Chart of Accounts format
    """
    engine = AccountInferenceEngine()
    return engine.derive_chart_of_accounts(gl_json_data)
