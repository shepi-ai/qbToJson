#!/usr/bin/env python3
"""
LLM fallback extraction for General Ledger PDFs.

This module is intentionally conservative:
- PDF-only fallback path
- feature-flagged (ENABLE_LLM_GL_PDF_FALLBACK)
- deterministic validation + reconciliation after LLM extraction
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import requests

# Max chars of GL text per LLM call.  ~50 k chars ≈ 12 k input tokens, which
# leaves plenty of room for the model to produce clean JSON within max_tokens.
_CHUNK_MAX_CHARS = 50_000


def _chunk_text(text: str, max_chars: int = _CHUNK_MAX_CHARS) -> List[str]:
    """Split *text* into chunks of at most *max_chars* characters.

    Splits preferentially at blank lines (paragraph/page boundaries) so that
    GL account headers stay with their transaction rows.  Falls back to any
    newline when no blank line is available within the look-behind window.
    """
    if len(text) <= max_chars:
        return [text]

    chunks: List[str] = []
    start = 0
    while start < len(text):
        end = start + max_chars
        if end >= len(text):
            chunks.append(text[start:])
            break

        # Prefer a blank-line boundary within the last 20 % of the window.
        search_from = start + int(max_chars * 0.80)
        blank = text.rfind("\n\n", search_from, end)
        if blank >= search_from:
            split_at = blank + 2  # include the blank line in the preceding chunk
        else:
            # Fall back to the last newline in the window.
            nl = text.rfind("\n", search_from, end)
            split_at = (nl + 1) if nl >= search_from else end

        chunks.append(text[start:split_at])
        start = split_at

    return chunks


def _to_float(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "").replace("$", "")
    if not s:
        return 0.0
    if s.startswith("(") and s.endswith(")"):
        s = f"-{s[1:-1]}"
    try:
        return float(s)
    except ValueError:
        return 0.0


def _parse_date(value: str) -> Optional[str]:
    if not value:
        return None
    raw = str(value).strip()
    for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            dt = datetime.strptime(raw, fmt)
            return dt.strftime("%m/%d/%Y")
        except ValueError:
            continue
    return None


def _extract_json_block(text: str) -> Dict[str, Any]:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?", "", cleaned).strip()
        cleaned = re.sub(r"```$", "", cleaned).strip()

    candidates: List[str] = [cleaned]
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start >= 0 and end > start:
        candidates.append(cleaned[start : end + 1])

    last_err: Optional[Exception] = None
    for candidate in candidates:
        if not candidate:
            continue
        for attempt in (_json_identity, _json_fix_trailing_commas):
            payload = attempt(candidate)
            try:
                return json.loads(payload)
            except json.JSONDecodeError as e:
                last_err = e
                continue

    if last_err:
        raise last_err
    raise json.JSONDecodeError("Could not parse LLM JSON output", cleaned, 0)


def _json_identity(text: str) -> str:
    return text


def _json_fix_trailing_commas(text: str) -> str:
    # common LLM error: trailing commas before closing braces/brackets
    return re.sub(r",\s*([}\]])", r"\1", text)


def _extract_rows_loose(text: str) -> List[Dict[str, Any]]:
    """Best-effort recovery when full JSON payload is malformed.

    Attempts to parse object chunks within the rows array even if the outer JSON
    is truncated or contains minor formatting errors.
    """
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?", "", cleaned).strip()
        cleaned = re.sub(r"```$", "", cleaned).strip()

    key_idx = cleaned.find('"rows"')
    if key_idx < 0:
        key_idx = cleaned.find("rows")
    if key_idx < 0:
        return []

    arr_start = cleaned.find("[", key_idx)
    if arr_start < 0:
        return []

    # Scan the array region and pull balanced object chunks.
    chunks: List[str] = []
    i = arr_start + 1
    depth = 0
    obj_start = -1
    in_string = False
    escape = False

    while i < len(cleaned):
        ch = cleaned[i]

        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
        else:
            if ch == '"':
                in_string = True
            elif ch == "{":
                if depth == 0:
                    obj_start = i
                depth += 1
            elif ch == "}":
                if depth > 0:
                    depth -= 1
                    if depth == 0 and obj_start >= 0:
                        chunks.append(cleaned[obj_start : i + 1])
                        obj_start = -1
            elif ch == "]" and depth == 0:
                break
        i += 1

    rows: List[Dict[str, Any]] = []
    for chunk in chunks:
        candidate = _json_fix_trailing_commas(chunk)
        try:
            obj = json.loads(candidate)
            if isinstance(obj, dict):
                rows.append(obj)
        except Exception:
            continue

    return rows


@dataclass
class GLValidationResult:
    normalized_rows: List[Dict[str, Any]]
    issues: List[str]
    confidence: float
    reconciled: bool
    totals: Dict[str, float]


class GeneralLedgerLLMFallback:
    def __init__(self) -> None:
        self.enabled = os.getenv("ENABLE_LLM_GL_PDF_FALLBACK", "false").lower() == "true"
        self.api_key = os.getenv("ANTHROPIC_API_KEY")
        self.model = os.getenv("GL_LLM_MODEL", "claude-haiku-4-5-20251001")
        self.max_attempts = int(os.getenv("GL_LLM_MAX_ATTEMPTS", "2"))
        self.chunk_max_chars = int(os.getenv("GL_LLM_CHUNK_CHARS", str(_CHUNK_MAX_CHARS)))

    def is_available(self) -> bool:
        return self.enabled and bool(self.api_key)

    def extract_rows(self, pdf_text: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
        """Run extraction + retry with validation deltas.

        For long documents the text is automatically split into chunks so that
        each LLM call stays well within the output token budget.
        """
        chunks = _chunk_text(pdf_text, self.chunk_max_chars)
        if len(chunks) > 1:
            return self._extract_rows_chunked(pdf_text, chunks)
        return self._extract_rows_single(pdf_text)

    # ------------------------------------------------------------------
    # Internal extraction helpers
    # ------------------------------------------------------------------

    def _extract_rows_single(self, pdf_text: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
        """Single-call extraction path (original logic, used for short documents)."""
        last_rows: List[Dict[str, Any]] = []
        last_validation: Optional[GLValidationResult] = None
        last_parse_error: Optional[str] = None
        attempts_used = 0

        for attempt in range(1, self.max_attempts + 1):
            attempts_used = attempt
            prompt = self._build_prompt(pdf_text, last_validation, last_parse_error)
            response_text = self._call_anthropic(prompt)
            try:
                payload = _extract_json_block(response_text)
                last_parse_error = None
            except Exception as e:
                recovered_rows = _extract_rows_loose(response_text)
                if recovered_rows:
                    payload = {"rows": recovered_rows}
                    last_parse_error = f"Recovered rows from malformed JSON on attempt {attempt}: {e}"
                else:
                    last_parse_error = f"JSON parse error on attempt {attempt}: {e}"
                    continue

            rows = payload.get("rows", []) if isinstance(payload, dict) else []
            validation = self.validate_and_reconcile(rows)
            last_rows = validation.normalized_rows
            last_validation = validation

            if validation.reconciled and validation.confidence >= 0.75:
                break

        # STRICT RECONCILIATION ENFORCEMENT: Only accept perfectly balanced GLs
        if not last_validation or not last_validation.reconciled:
            delta = last_validation.totals['delta'] if last_validation else 0.0
            issues = last_validation.issues[:10] if last_validation else ["No validation performed"]
            raise ValueError(
                f"General Ledger could not be reconciled after {self.max_attempts} attempts. "
                f"Debits must equal Credits for accurate financial reporting. "
                f"Debit-Credit imbalance: ${abs(delta):.2f}. "
                f"Issues: {issues}"
            )

        return self._build_meta(last_rows, last_validation, last_parse_error, attempts_used, chunks_used=1)

    def _extract_rows_chunked(
        self,
        pdf_text: str,
        chunks: List[str],
    ) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
        """Multi-chunk extraction: one LLM call per chunk, merge, then validate."""
        all_rows: List[Dict[str, Any]] = []
        total_chunks = len(chunks)
        total_api_calls = 0
        chunk_issues: List[str] = []

        for chunk_idx, chunk_text in enumerate(chunks):
            chunk_rows, calls_used, issue = self._extract_chunk(
                chunk_text, chunk_idx, total_chunks
            )
            total_api_calls += calls_used
            all_rows.extend(chunk_rows)
            if issue:
                chunk_issues.append(f"chunk {chunk_idx+1}/{total_chunks}: {issue}")

        # Validate merged rows from all chunks.
        validation = self.validate_and_reconcile(all_rows)
        # Prepend any per-chunk extraction warnings.
        all_issues = chunk_issues + validation.issues
        is_reconciled = validation.reconciled and not chunk_issues

        # STRICT RECONCILIATION ENFORCEMENT: Only accept perfectly balanced GLs
        if not is_reconciled:
            delta = validation.totals.get('delta', 0.0)
            raise ValueError(
                f"General Ledger (chunked) could not be reconciled after {total_api_calls} API calls across {total_chunks} chunks. "
                f"Debits must equal Credits for accurate financial reporting. "
                f"Debit-Credit imbalance: ${abs(delta):.2f}. "
                f"Issues: {all_issues[:10]}"
            )

        meta = {
            "source": "llm_fallback_chunked",
            "chunks": total_chunks,
            "attempts": total_api_calls,
            "confidence": round(validation.confidence, 4),
            "reconciled": is_reconciled,
            "issues": all_issues[:25],
            "totals": validation.totals,
        }
        return validation.normalized_rows, meta

    def _extract_chunk(
        self,
        chunk_text: str,
        chunk_idx: int,
        total_chunks: int,
    ) -> Tuple[List[Dict[str, Any]], int, Optional[str]]:
        """Extract rows from a single chunk; return (rows, api_calls_used, issue_msg)."""
        last_parse_error: Optional[str] = None

        for attempt in range(1, self.max_attempts + 1):
            prompt = self._build_chunk_prompt(
                chunk_text, chunk_idx, total_chunks, last_parse_error
            )
            response_text = self._call_anthropic(prompt)

            try:
                payload = _extract_json_block(response_text)
                last_parse_error = None
            except Exception as e:
                recovered_rows = _extract_rows_loose(response_text)
                if recovered_rows:
                    return recovered_rows, attempt, f"Recovered from malformed JSON: {e}"
                last_parse_error = str(e)
                continue

            rows = payload.get("rows", []) if isinstance(payload, dict) else []
            if rows:
                return rows, attempt, None

        # All attempts exhausted with no rows.
        return [], self.max_attempts, last_parse_error or "no rows extracted"

    def _build_meta(
        self,
        last_rows: List[Dict[str, Any]],
        last_validation: Optional[GLValidationResult],
        last_parse_error: Optional[str],
        attempts_used: int,
        chunks_used: int,
    ) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
        meta: Dict[str, Any] = {
            "source": "llm_fallback",
            "attempts": attempts_used,
        }
        if last_validation:
            meta.update(
                {
                    "confidence": round(last_validation.confidence, 4),
                    "reconciled": last_validation.reconciled,
                    "issues": last_validation.issues[:25],
                    "totals": last_validation.totals,
                }
            )
        elif last_parse_error:
            meta.update(
                {
                    "confidence": 0.0,
                    "reconciled": False,
                    "issues": [last_parse_error],
                    "totals": {"total_debit": 0.0, "total_credit": 0.0, "delta": 0.0},
                }
            )
        return last_rows, meta

    def _call_anthropic(self, prompt: str) -> str:
        """Call Anthropic API with retry logic for timeout/rate limit errors."""
        import time
        
        headers = {
            "x-api-key": self.api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json",
        }
        payload = {
            "model": self.model,
            "max_tokens": 12000,
            "temperature": 0,
            "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}]}],
        }
        
        # Retry logic with exponential backoff
        max_retries = 3
        base_timeout = 120  # Increased from 60s to 120s
        
        for attempt in range(max_retries):
            try:
                timeout = base_timeout + (attempt * 30)  # 120s, 150s, 180s
                resp = requests.post(
                    "https://api.anthropic.com/v1/messages",
                    headers=headers,
                    json=payload,
                    timeout=timeout
                )
                resp.raise_for_status()
                data = resp.json()
                return data["content"][0]["text"]
                
            except requests.exceptions.ReadTimeout as e:
                if attempt < max_retries - 1:
                    wait_time = 2 ** attempt  # 1s, 2s, 4s
                    print(f"⚠️ Anthropic API timeout on attempt {attempt+1}/{max_retries}, retrying in {wait_time}s...")
                    time.sleep(wait_time)
                    continue
                else:
                    # Last attempt failed, re-raise
                    raise TimeoutError(
                        f"Anthropic API timed out after {max_retries} attempts. "
                        f"Last timeout was {timeout}s. Consider processing smaller chunks."
                    ) from e
                    
            except requests.exceptions.HTTPError as e:
                # Handle rate limiting (429) with retry
                if e.response.status_code == 429 and attempt < max_retries - 1:
                    wait_time = 5 * (2 ** attempt)  # 5s, 10s, 20s
                    print(f"⚠️ Rate limited by Anthropic API, retrying in {wait_time}s...")
                    time.sleep(wait_time)
                    continue
                else:
                    raise

    def _build_prompt(
        self,
        pdf_text: str,
        previous: Optional[GLValidationResult],
        parse_error: Optional[str] = None,
    ) -> str:
        schema, rules = self._schema_and_rules()
        hint = ""
        if parse_error:
            hint += f"\n\nPrevious response was not valid JSON. Fix this parse error: {parse_error}\n"
        if previous:
            hint += (
                "\n\nPrevious validation failed. Fix these deltas exactly:\n"
                f"- reconciled: {previous.reconciled}\n"
                f"- confidence: {previous.confidence:.4f}\n"
                f"- issues: {previous.issues[:8]}\n"
                f"- totals: {previous.totals}\n"
            )
        return f"{schema}\n{rules}{hint}\n\nGL TEXT:\n{pdf_text[:200000]}"

    def _build_chunk_prompt(
        self,
        chunk_text: str,
        chunk_idx: int,
        total_chunks: int,
        parse_error: Optional[str] = None,
    ) -> str:
        schema, rules = self._schema_and_rules()
        chunk_note = (
            f"\nThis is chunk {chunk_idx + 1} of {total_chunks} from a General Ledger PDF. "
            "Extract ALL transaction rows present in this portion only. "
            "Do not skip rows to save space.\n"
        )
        hint = ""
        if parse_error:
            hint = f"\n\nPrevious response was not valid JSON. Fix this parse error: {parse_error}\n"
        return f"{schema}\n{rules}{chunk_note}{hint}\n\nGL TEXT:\n{chunk_text}"

    @staticmethod
    def _schema_and_rules() -> Tuple[str, str]:
        schema = (
            "Return JSON only with shape: "
            "{\"rows\":[{\"date\":\"MM/DD/YYYY\",\"account\":\"...\",\"memo\":\"...\","
            "\"debit\":number,\"credit\":number,\"running_balance\":number|null,\"doc_num\":\"...\"}]}"
        )
        rules = (
            "Rules: date/account required. Exactly one of debit/credit should be > 0 per row. "
            "Do not invent rows. Keep debit/credit numeric. "
            "IMPORTANT: return strictly valid JSON (no trailing commas, no comments, no markdown)."
        )
        return schema, rules

    def validate_and_reconcile(self, rows: List[Dict[str, Any]]) -> GLValidationResult:
        issues: List[str] = []
        normalized: List[Dict[str, Any]] = []
        total_debit = 0.0
        total_credit = 0.0

        by_account: Dict[str, List[Dict[str, Any]]] = {}

        for idx, row in enumerate(rows):
            date_raw = str(row.get("date", "")).strip()
            account = str(row.get("account", "")).strip()
            memo = str(row.get("memo", "")).strip()
            parsed_date = _parse_date(date_raw)
            debit = max(_to_float(row.get("debit")), 0.0)
            credit = max(_to_float(row.get("credit")), 0.0)
            running_balance = row.get("running_balance")
            rb = _to_float(running_balance) if running_balance not in (None, "") else None
            doc_num = str(row.get("doc_num", "")).strip()

            if not parsed_date:
                issues.append(f"row {idx+1}: invalid/missing date")
            if not account:
                issues.append(f"row {idx+1}: missing account")
            if debit > 0 and credit > 0:
                issues.append(f"row {idx+1}: both debit and credit populated")
            if debit == 0 and credit == 0:
                issues.append(f"row {idx+1}: both debit and credit are zero")

            norm = {
                "date": parsed_date or date_raw,
                "account": account,
                "memo": memo,
                "debit": debit,
                "credit": credit,
                "running_balance": rb,
                "doc_num": doc_num,
            }
            normalized.append(norm)
            by_account.setdefault(account or "__missing__", []).append(norm)
            total_debit += debit
            total_credit += credit

        # per-period integrity
        period_delta = round(total_debit - total_credit, 2)
        if abs(period_delta) > 0.01:
            issues.append(f"period totals mismatch: debit-credit={period_delta:.2f}")

        # per-account running balance coherence when balances are available
        for account, a_rows in by_account.items():
            prev_balance: Optional[float] = None
            for r in a_rows:
                rb = r.get("running_balance")
                if rb is None:
                    continue
                net = r["debit"] - r["credit"]
                if prev_balance is not None:
                    expected = round(prev_balance + net, 2)
                    if abs(expected - rb) > 0.02:
                        issues.append(
                            f"account '{account}': running balance mismatch expected {expected:.2f}, got {rb:.2f}"
                        )
                        break
                prev_balance = rb

        # confidence scoring
        row_count = max(len(normalized), 1)
        penalty = min(len(issues) / (row_count * 0.6), 1.0)
        confidence = max(0.0, 1.0 - penalty)
        reconciled = len(issues) == 0

        return GLValidationResult(
            normalized_rows=normalized,
            issues=issues,
            confidence=confidence,
            reconciled=reconciled,
            totals={
                "total_debit": round(total_debit, 2),
                "total_credit": round(total_credit, 2),
                "delta": period_delta,
            },
        )
