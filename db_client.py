"""
Database Client for qbToJson API
Handles database operations and storage access via db-proxy
"""

import os
import requests
from typing import Dict, Optional
import logging

logger = logging.getLogger(__name__)


class DatabaseClient:
    """Client for Supabase database and storage operations"""
    
    def __init__(self):
        """Initialize database client with db-proxy credentials"""
        self.db_proxy_url = os.getenv(
            'DB_PROXY_URL',
            'https://sqwohcvobfnymsbzlfqr.supabase.co/functions/v1/db-proxy'
        )
        self.api_key = os.getenv('QBTOJSON_API_KEY')
        
        if not self.api_key:
            logger.warning("QBTOJSON_API_KEY not set - database saving will fail")
        
        self.headers = {
            'x-api-key': self.api_key,
            'x-service-name': 'qbtojson-api',
            'Content-Type': 'application/json'
        }
    
    def save_converted_data(self, 
                           project_id: str,
                           data_type: str,
                           data: Dict,
                           source_document_id: Optional[str] = None,
                           filename: Optional[str] = None,
                           source_type: str = 'qbtojson') -> bool:
        """
        Save converted data to database
        
        Args:
            project_id: Supabase project UUID
            data_type: Type of data (e.g., 'trial_balance', 'balance_sheet')
            data: Converted data dictionary
            source_document_id: Optional source document UUID
            filename: Optional original filename
            source_type: Source of the data (default: 'qbtojson', can be 'derived_from_gl')
            
        Returns:
            True if successful, False otherwise
        """
        if not self.api_key:
            logger.error("Cannot save to database: QBTOJSON_API_KEY not configured")
            return False
        
        try:
            # Extract record count if possible
            record_count = self._extract_record_count(data, data_type)
            
            # Prepare data for insertion
            insert_data = {
                'project_id': project_id,
                'source_type': source_type,
                'data_type': data_type,
                'data': data
            }
            
            if source_document_id:
                insert_data['source_document_id'] = source_document_id
            
            if record_count is not None:
                insert_data['record_count'] = record_count
            
            # Save to database via db-proxy
            payload = {
                'action': 'query',
                'table': 'processed_data',
                'operation': 'insert',
                'data': insert_data
            }
            
            response = requests.post(
                self.db_proxy_url,
                headers=self.headers,
                json=payload,
                timeout=30
            )
            
            response.raise_for_status()
            result = response.json()
            
            if result.get('success'):
                logger.info(f"✅ Saved {data_type} to database (project: {project_id})")
                return True
            else:
                logger.error(f"❌ Database save failed: {result}")
                return False
        
        except requests.exceptions.RequestException as e:
            logger.error(f"❌ Error saving to database: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                logger.error(f"Response: {e.response.text}")
            return False
        
        except Exception as e:
            logger.error(f"❌ Unexpected error saving to database: {str(e)}")
            return False
    
    def _extract_record_count(self, data: Dict, data_type: str) -> Optional[int]:
        """
        Extract record count from converted data
        
        Args:
            data: Converted data dictionary
            data_type: Type of data
            
        Returns:
            Number of records or None
        """
        try:
            if data_type == 'trial_balance':
                if 'monthlyReports' in data:
                    return len(data['monthlyReports'])
            
            elif data_type in ['balance_sheet', 'income_statement', 'cash_flow']:
                if isinstance(data, list):
                    return len(data)
            
            elif data_type in ['chart_of_accounts', 'accounts_payable', 'accounts_receivable']:
                if isinstance(data, list):
                    return len(data)
            
            elif data_type in ['general_ledger', 'journal_entries']:
                if 'rows' in data and 'row' in data['rows']:
                    return len(data['rows']['row'])
        
        except Exception as e:
            logger.warning(f"Could not extract record count: {str(e)}")
        
        return None
    
    def check_coa_exists(self, project_id: str) -> bool:
        """
        Check if Chart of Accounts exists for project
        
        Args:
            project_id: Supabase project UUID
            
        Returns:
            True if COA exists, False otherwise
        """
        if not self.api_key:
            logger.warning("Cannot check COA: QBTOJSON_API_KEY not configured")
            return False
        
        try:
            payload = {
                'action': 'query',
                'table': 'processed_data',
                'operation': 'select',
                'filters': {
                    'project_id': project_id,
                    'data_type': 'chart_of_accounts'
                },
                'columns': ['id'],
                'limit': 1
            }
            
            response = requests.post(
                self.db_proxy_url,
                headers=self.headers,
                json=payload,
                timeout=10
            )
            
            response.raise_for_status()
            result = response.json()
            
            if result.get('success') and result.get('data'):
                exists = len(result['data']) > 0
                logger.info(f"COA exists for project {project_id}: {exists}")
                return exists
            
            return False
        
        except Exception as e:
            logger.error(f"Error checking COA existence: {str(e)}")
            return False  # Assume doesn't exist on error
    
    def is_configured(self) -> bool:
        """Check if database saver is properly configured"""
        return self.api_key is not None
    
    def download_from_storage(self, file_path: str, bucket: str = 'documents') -> bytes:
        """
        Download file from Supabase Storage (in-memory, container-optimized)
        
        Args:
            file_path: Path to file in storage bucket (e.g., 'project-id/file.xlsx')
            bucket: Storage bucket name
            
        Returns:
            File content as bytes
        """
        if not self.api_key:
            raise ValueError("Cannot download from storage: QBTOJSON_API_KEY not configured")
        
        try:
            payload = {
                'action': 'storage',
                'operation': 'download',
                'bucket': bucket,
                'path': file_path
            }
            
            response = requests.post(
                self.db_proxy_url,
                headers=self.headers,
                json=payload,
                timeout=60
            )
            
            response.raise_for_status()
            result = response.json()
            
            if result.get('success') and result.get('data'):
                # Decode base64 data
                import base64
                file_data = base64.b64decode(result['data'])
                logger.info(f"Downloaded file from storage: {file_path} ({len(file_data)} bytes)")
                return file_data
            else:
                raise Exception(f"Storage download failed: {result}")
        
        except requests.exceptions.RequestException as e:
            logger.error(f"Error downloading from storage: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                logger.error(f"Response: {e.response.text}")
            raise
        
        except Exception as e:
            logger.error(f"Unexpected error downloading from storage: {str(e)}")
            raise
    
    def get_signed_url(self, file_path: str, bucket: str = 'documents', 
                      expires_in: int = 3600) -> str:
        """
        Get signed URL for file (alternative to direct download)
        
        Args:
            file_path: Path to file in storage bucket
            bucket: Storage bucket name  
            expires_in: URL expiration in seconds (default: 1 hour)
            
        Returns:
            Signed URL string
        """
        if not self.api_key:
            raise ValueError("Cannot get signed URL: QBTOJSON_API_KEY not configured")
        
        try:
            payload = {
                'action': 'storage',
                'operation': 'signed_url',
                'bucket': bucket,
                'path': file_path,
                'expires_in': expires_in
            }
            
            response = requests.post(
                self.db_proxy_url,
                headers=self.headers,
                json=payload,
                timeout=30
            )
            
            response.raise_for_status()
            result = response.json()
            
            if result.get('success') and result.get('signed_url'):
                logger.info(f"Generated signed URL for: {file_path}")
                return result['signed_url']
            else:
                raise Exception(f"Signed URL generation failed: {result}")
        
        except Exception as e:
            logger.error(f"Error getting signed URL: {str(e)}")
            raise


# Global instance for convenience
_db_client = None

def get_db_client() -> DatabaseClient:
    """Get global DatabaseClient instance"""
    global _db_client
    if _db_client is None:
        _db_client = DatabaseClient()
    return _db_client
