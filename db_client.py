"""
Database client for qbToJson API.

Uses the Supabase Python SDK directly (no more db-proxy).
"""

import os
import logging
from typing import Dict, Optional

from supabase import create_client, Client

logger = logging.getLogger(__name__)


class DatabaseClient:
    """Client for Supabase database and storage operations."""

    def __init__(self):
        self.supabase_url = os.getenv("SUPABASE_URL", "https://mdgmessqbfebrbvjtndz.supabase.co")
        self.service_role_key = os.getenv("SUPABASE_SERVICE_ROLE_KEY")

        self._client: Optional[Client] = None
        if not self.service_role_key:
            logger.warning("SUPABASE_SERVICE_ROLE_KEY not set - database saving will fail")

    @property
    def client(self) -> Client:
        if self._client is None:
            if not self.service_role_key:
                raise RuntimeError("SUPABASE_SERVICE_ROLE_KEY is not configured")
            self._client = create_client(self.supabase_url, self.service_role_key)
        return self._client

    def is_configured(self) -> bool:
        return self.service_role_key is not None

    def save_converted_data(self,
                            project_id: str,
                            data_type: str,
                            data: Dict,
                            source_document_id: Optional[str] = None,
                            filename: Optional[str] = None,
                            source_type: str = "qbtojson") -> bool:
        """Save converted data into processed_data."""
        if not self.is_configured():
            logger.error("Cannot save to database: SUPABASE_SERVICE_ROLE_KEY not configured")
            return False

        try:
            record_count = self._extract_record_count(data, data_type)

            insert_data = {
                "project_id": project_id,
                "source_type": source_type,
                "data_type": data_type,
                "data": data,
            }
            if source_document_id:
                insert_data["source_document_id"] = source_document_id
            if record_count is not None:
                insert_data["record_count"] = record_count

            self.client.table("processed_data").insert(insert_data).execute()
            logger.info(f"✅ Saved {data_type} to database (project: {project_id})")
            return True

        except Exception as e:
            logger.error(f"❌ Error saving to database: {e}")
            return False

    def _extract_record_count(self, data: Dict, data_type: str) -> Optional[int]:
        try:
            if data_type == "trial_balance":
                if "monthlyReports" in data:
                    return len(data["monthlyReports"])

            elif data_type in ("balance_sheet", "income_statement", "cash_flow"):
                if isinstance(data, list):
                    return len(data)

            elif data_type in ("chart_of_accounts", "accounts_payable", "accounts_receivable"):
                if isinstance(data, list):
                    return len(data)

            elif data_type in ("general_ledger", "journal_entries"):
                if "rows" in data and "row" in data["rows"]:
                    return len(data["rows"]["row"])

        except Exception as e:
            logger.warning(f"Could not extract record count: {e}")

        return None

    def check_coa_exists(self, project_id: str) -> bool:
        """Check if Chart of Accounts exists for project."""
        if not self.is_configured():
            logger.warning("Cannot check COA: SUPABASE_SERVICE_ROLE_KEY not configured")
            return False

        try:
            result = (
                self.client.table("processed_data")
                .select("id")
                .eq("project_id", project_id)
                .eq("data_type", "chart_of_accounts")
                .limit(1)
                .execute()
            )
            exists = bool(result.data)
            logger.info(f"COA exists for project {project_id}: {exists}")
            return exists

        except Exception as e:
            logger.error(f"Error checking COA existence: {e}")
            return False

    def download_from_storage(self, file_path: str, bucket: str = "documents") -> bytes:
        """Download a file from Supabase Storage."""
        if not self.is_configured():
            raise ValueError("Cannot download from storage: SUPABASE_SERVICE_ROLE_KEY not configured")

        try:
            file_data = self.client.storage.from_(bucket).download(file_path)
            logger.info(f"Downloaded file from storage: {file_path} ({len(file_data)} bytes)")
            return file_data
        except Exception as e:
            logger.error(f"Error downloading from storage: {e}")
            raise

    def get_signed_url(self, file_path: str, bucket: str = "documents",
                       expires_in: int = 3600) -> str:
        """Create a signed URL for a file in Supabase Storage."""
        if not self.is_configured():
            raise ValueError("Cannot get signed URL: SUPABASE_SERVICE_ROLE_KEY not configured")

        try:
            result = self.client.storage.from_(bucket).create_signed_url(file_path, expires_in)
            signed_url = result.get("signedURL") or result.get("signed_url")
            if not signed_url:
                raise Exception(f"Signed URL missing from response: {result}")
            logger.info(f"Generated signed URL for: {file_path}")
            return signed_url
        except Exception as e:
            logger.error(f"Error getting signed URL: {e}")
            raise


# Global instance for convenience
_db_client: Optional[DatabaseClient] = None


def get_db_client() -> DatabaseClient:
    """Get global DatabaseClient instance."""
    global _db_client
    if _db_client is None:
        _db_client = DatabaseClient()
    return _db_client
