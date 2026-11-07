"""Pure, offline-testable logic for the Sheets exporter.

Nothing in this module touches the network or the Apify platform, so it can
be unit-tested without credentials.
"""
from __future__ import annotations

import json
import re
import time
from typing import Any, Callable, Iterable, Iterator, Sequence

SHEETS_CELL_LIMIT = 10_000_000
DEFAULT_MAX_COLUMNS = 500
DEFAULT_CHUNK_ROWS = 5_000
CELLS_PER_REQUEST = 100_000
FREE_ROWS_PER_RUN = 2_000  # hardcoded on purpose: never read from user input

REQUIRED_SA_FIELDS = ("type", "client_email", "private_key", "token_uri")


class UserFacingError(Exception):
    """An error whose message is safe to show the user and explains the fix."""


# --------------------------------------------------------------------------- #
# Credentials
# --------------------------------------------------------------------------- #
def parse_service_account(raw: str | dict | None) -> dict:
    """Validate a pasted service-account key and return it as a dict.

    Error messages never include any part of the key.
    """
    if not raw:
        raise UserFacingError(
            "No service-account key provided. Create one in Google Cloud Console "
            "(IAM & Admin > Service Accounts > Keys > Add key > JSON) and paste the "
            "whole JSON file into the 'Service account key' field."
        )
    if isinstance(raw, dict):
        info = raw
    else:
        try:
            info = json.loads(raw)
        except json.JSONDecodeError:
            raise UserFacingError(
                "The service-account key is not valid JSON. Paste the entire "
                "downloaded .json file, including the curly braces."
            ) from None
    if not isinstance(info, dict):
        raise UserFacingError("The service-account key must be a JSON object.")
    if info.get("type") != "service_account":
        raise UserFacingError(
            "This key is not a service-account key (its 'type' is not "
            "'service_account'). OAuth client IDs and API keys will not work."
        )
    missing = [f for f in REQUIRED_SA_FIELDS if not info.get(f)]
    if missing:
        raise UserFacingError(
            "The service-account key is missing required fields: "
            + ", ".join(missing)
            + ". Download a fresh JSON key from Google Cloud Console."
        )
    return info


def redact(text: str, secrets: Iterable[str]) -> str:
    """Remove secret values (and anything that looks like a PEM key) from text."""
    out = text
    for s in secrets:
        if s:
            out = out.replace(s, "[REDACTED]")
    out = re.sub(
        r"-----BEGIN [A-Z ]*PRIVATE KEY-----.*?-----END [A-Z ]*PRIVATE KEY-----",
        "[REDACTED KEY]",
        out,
        flags=re.DOTALL,
    )
    return out


def secrets_from(info: dict | None) -> list[str]:
    if not info:
        return []
    # client_email is NOT secret: users must paste it into the Share dialog, and our
    # error messages need to show it.
    return [str(info.get(k, "")) for k in ("private_key", "private_key_id")]


# --------------------------------------------------------------------------- #
# Table shaping
# --------------------------------------------------------------------------- #
def flatten(record: Any, sep: str = ".", max_depth: int = 5) -> dict[str, Any]:
    """Flatten nested dicts into dot.path keys. Lists become JSON strings."""
    out: dict[str, Any] = {}

    def walk(value: Any, prefix: str, depth: int) -> None:
        if isinstance(value, dict) and depth < max_depth and value:
            for k, v in value.items():
                walk(v, f"{prefix}{sep}{k}" if prefix else str(k), depth + 1)
        else:
            out[prefix or "value"] = to_cell(value)

    walk(record, "", 0)
    return out


def to_cell(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float, str)):
        if isinstance(value, str) and len(value) > 50_000:
            return value[:49_990] + "…[cut]"  # Sheets cell limit is 50k chars
        return value
    return json.dumps(value, ensure_ascii=False, default=str)


def build_table(
    records: Iterable[Any], max_columns: int = DEFAULT_MAX_COLUMNS
) -> tuple[list[str], list[list[Any]], list[str]]:
    """Return (header, rows, dropped_columns). Column order = first seen."""
    flat = [flatten(r) for r in records]
    header: list[str] = []
    seen: set[str] = set()
    for row in flat:
        for k in row:
            if k not in seen:
                seen.add(k)
                header.append(k)
    dropped = header[max_columns:]
    header = header[:max_columns]
    rows = [[row.get(k, "") for k in header] for row in flat]
    return header, rows, dropped


# --------------------------------------------------------------------------- #
# Limits, chunking, retries, billing
# --------------------------------------------------------------------------- #
def preflight_cells(
    n_rows: int, n_cols: int, other_cells_in_spreadsheet: int = 0,
    limit: int = SHEETS_CELL_LIMIT,
) -> None:
    needed = (n_rows + 1) * max(n_cols, 1)  # +1 header row
    if needed + other_cells_in_spreadsheet > limit:
        raise UserFacingError(
            f"This export needs about {needed:,} cells, but Google Sheets allows "
            f"{limit:,} cells per spreadsheet (other tabs already use "
            f"{other_cells_in_spreadsheet:,}). Reduce rows or columns, choose an "
            "empty spreadsheet, or use the CSV/XLSX output instead."
        )


def rows_per_chunk(n_cols: int, max_rows: int = DEFAULT_CHUNK_ROWS,
                   cell_budget: int = CELLS_PER_REQUEST) -> int:
    """Rows per write request, bounded by cells so wide tables don't exceed request size."""
    return max(1, min(max_rows, cell_budget // max(n_cols, 1)))


_NUMERIC = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")


_PLAIN_NUMBER = re.compile(r"-?(0|[1-9]\d*)(\.\d+)?")
MAX_INT_DIGITS = 11  # Excel's General format shows longer integers as 1.23457E+12
MAX_SIG_DIGITS = 15  # a double holds ~15 significant digits exactly


def as_number(value: Any) -> Any:
    """Turn one plain numeric string ("22", "-7.25") into a number, else return it unchanged.

    Stays text: leading zeros ("02134"), "-0", "1e5", "+3", " 5", "1,000", integers over
    11 digits (barcodes, IDs) and decimals over 15 significant digits (precision).
    """
    if not isinstance(value, str) or not _PLAIN_NUMBER.fullmatch(value):
        return value
    whole, _, frac = value.lstrip("-").partition(".")
    if value.startswith("-") and not value.strip("-0."):
        return value  # "-0", "-0.00"
    if not frac:
        return int(value) if len(whole) <= MAX_INT_DIGITS else value
    sig = len((whole + frac).lstrip("0"))
    return float(value) if sig <= MAX_SIG_DIGITS else value


def _convertible(v: Any) -> bool:
    if isinstance(v, bool):
        return False
    return isinstance(v, (int, float)) or as_number(v) is not v


def numberize(rows: list[list[Any]]) -> list[list[Any]]:
    """Convert numeric strings column by column, only where EVERY non-empty value converts.

    A mixed column (e.g. SKUs "10023" and "A-77") is left entirely as text so it sorts
    consistently. Returns new rows; the input is not modified.
    """
    if not rows:
        return rows
    width = max(len(r) for r in rows)
    numeric = []
    for c in range(width):
        vals = [r[c] for r in rows if c < len(r) and r[c] != ""]
        if vals and any(isinstance(v, str) for v in vals) and all(_convertible(v) for v in vals):
            numeric.append(c)
    if not numeric:
        return rows
    out = [list(r) for r in rows]
    for r in out:
        for c in numeric:
            if c < len(r):
                r[c] = as_number(r[c])
    return out


def escape_formula(value: Any) -> Any:
    """Neutralize spreadsheet formula injection in CSV/XLSX output.

    Strings starting with = + - @ (or tab/CR) get a leading apostrophe, unless they
    are plain numbers such as "-5" or "+3.2".
    """
    if isinstance(value, str) and value and value[0] in "=+-@\t\r":
        if not _NUMERIC.match(value):
            return "'" + value
    return value


def chunked(rows: Sequence[Any], size: int = DEFAULT_CHUNK_ROWS) -> Iterator[Sequence[Any]]:
    if size <= 0:
        raise ValueError("size must be positive")
    for i in range(0, len(rows), size):
        yield rows[i : i + size]


def with_backoff(
    fn: Callable[[], Any],
    is_retryable: Callable[[Exception], bool],
    attempts: int = 6,
    base_delay: float = 2.0,
    sleep: Callable[[float], None] = time.sleep,
) -> Any:
    """Call fn, retrying retryable errors with exponential backoff (2,4,8,... s)."""
    for i in range(attempts):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            if i == attempts - 1 or not is_retryable(e):
                raise
            sleep(min(base_delay * (2**i), 64))
    raise RuntimeError("unreachable")


def billable_row_batches(rows_written: int, free_rows: int, batch: int = 1000) -> int:
    """Number of 1k-row events to charge: only rows actually written beyond the free tier.

    Partial batches round up (e.g. 10,001 rows with 10,000 free -> 1 event).
    """
    over = max(0, rows_written - max(0, free_rows))
    return -(-over // batch)


def friendly_google_error(status: int | None, message: str, sa_email: str) -> str:
    """Map common Google API failures to a plain-English fix."""
    m = message.lower()
    if any(k in m for k in ("pem", "private key", "invalid_grant", "deserialize key", "asn.1")):
        return (
            "Google rejected the service-account key (it is corrupted, revoked, or "
            "was edited). Download a fresh JSON key and paste the whole file unchanged."
        )
    if status == 404 or "not found" in m:
        return (
            "Spreadsheet not found. Check the spreadsheet ID (the long part of the "
            "URL between /d/ and /edit) and make sure it is shared with the service "
            f"account: {sa_email} (Share > add this email > Editor)."
        )
    if status == 403 and ("has not been used" in m or "disabled" in m):
        return (
            "The Google Sheets API is not enabled for your Google Cloud project. "
            "Enable it at console.cloud.google.com > APIs & Services > Library > "
            "Google Sheets API, wait a minute, and run again."
        )
    if status == 403:
        return (
            "Permission denied. Share the spreadsheet with the service account "
            f"{sa_email} as an Editor (Share button in Google Sheets)."
        )
    if status == 429:
        return "Google Sheets rate limit kept being hit after several retries. Try again in a few minutes or use a larger chunk size."
    return f"Google Sheets returned an error ({status}): {message}"


def append_idempotent(
    append_fn: Callable[[], Any],
    count_rows: Callable[[], int],
    rows_before: int,
    chunk_len: int,
    is_rate_limited: Callable[[Exception], bool],
    is_ambiguous: Callable[[Exception], bool],
    attempts: int = 6,
    base_delay: float = 2.0,
    sleep: Callable[[float], None] = time.sleep,
) -> None:
    """Append a chunk at most once, even if a request fails after Google applied it.

    - 429: rejected before being applied, so it is safe to resend.
    - Timeout / connection error / 5xx (ambiguous): re-count the table first. If the
      chunk landed, stop. If nothing landed, resend. If only part of it landed, stop
      with an error instead of guessing.
    """
    for i in range(attempts):
        try:
            append_fn()
            return
        except Exception as e:  # noqa: BLE001
            if i == attempts - 1 or not (is_rate_limited(e) or is_ambiguous(e)):
                raise
            sleep(min(base_delay * (2**i), 64))
            if is_ambiguous(e):
                now = count_rows()
                if now >= rows_before + chunk_len:
                    return  # it was applied; don't write it twice
                if now != rows_before:
                    raise UserFacingError(
                        f"A Google Sheets request failed midway: the tab now has {now - rows_before} "
                        f"of {chunk_len} rows from this batch. Stopped to avoid duplicates; "
                        "check the end of the tab, then re-run in Replace mode."
                    ) from None
    raise RuntimeError("unreachable")
