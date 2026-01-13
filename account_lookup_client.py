#!/usr/bin/env python3
"""
Account Lookup Client
Provides functionality to look up account IDs from the API
"""

import requests
from typing import Optional, Dict, Any
import logging

logger = logging.getLogger(__name__)


class AccountLookupClient:
    """Client for looking up account IDs from the API"""
    
    def __init__(self, api_base_url: str = None):
        if api_base_url is None:
            api_base_url = "http://localhost:8080"
        self.api_base_url = api_base_url.rstrip('/')
        self.cache = {}  # Local cache to avoid repeated API calls
        
    def lookup_account_id(self, account_name: str) -> Optional[str]:
        """
        Look up account ID by name
        
        Args:
            account_name: The name of the account to look up
            
        Returns:
            The account ID if found, None otherwise
        """
        if not account_name:
            return None
            
        # Check cache first
        cache_key = account_name.strip().lower()
        if cache_key in self.cache:
            return self.cache[cache_key]
        
        try:
            # Call the API
            response = requests.post(
                f"{self.api_base_url}/api/accounts/lookup",
                json={"name": account_name},
                timeout=5
            )
            
            if response.status_code == 200:
                data = response.json()
                if data.get('success') and data.get('account'):
                    account_id = data['account']['id']
                    # Cache the result
                    self.cache[cache_key] = account_id
                    if data.get('fuzzy_match'):
                        logger.info(f"Fuzzy matched '{account_name}' to account ID {account_id}")
                    return account_id
            elif response.status_code == 404:
                logger.warning(f"Account not found: {account_name}")
                # Cache the negative result
                self.cache[cache_key] = None
                return None
            else:
                logger.error(f"API error looking up account '{account_name}': {response.status_code}")
                
        except requests.exceptions.RequestException as e:
            logger.error(f"Network error looking up account '{account_name}': {e}")
        except Exception as e:
            logger.error(f"Unexpected error looking up account '{account_name}': {e}")
            
        return None
    
    def load_accounts_file(self, file_path: str) -> bool:
        """
        Load a Chart of Accounts file into the API
        
        Args:
            file_path: Path to the accounts file
            
        Returns:
            True if successful, False otherwise
        """
        try:
            with open(file_path, 'rb') as f:
                files = {'file': f}
                response = requests.post(
                    f"{self.api_base_url}/api/accounts/load",
                    files=files,
                    timeout=30
                )
                
            if response.status_code == 200:
                data = response.json()
                if data.get('success'):
                    logger.info(f"Loaded {data.get('accounts_loaded', 0)} accounts")
                    # Clear cache since we have new data
                    self.cache.clear()
                    return True
            else:
                logger.error(f"Failed to load accounts file: {response.status_code}")
                
        except Exception as e:
            logger.error(f"Error loading accounts file: {e}")
            
        return False
    
    def is_api_available(self) -> bool:
        """Check if the API is available"""
        try:
            response = requests.get(f"{self.api_base_url}/health", timeout=2)
            return response.status_code == 200
        except:
            return False


# Singleton instance
_client_instance = None

def get_account_lookup_client(api_base_url: Optional[str] = None) -> AccountLookupClient:
    """Get or create the singleton account lookup client"""
    global _client_instance
    if _client_instance is None or (api_base_url and api_base_url != _client_instance.api_base_url):
        _client_instance = AccountLookupClient(api_base_url or "http://localhost:8080")
    return _client_instance
