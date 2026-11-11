"""Apify Actor: export a dataset (or a public CSV/JSON URL) to Google Sheets, CSV and XLSX.

Auth is service-account only (v1): the user pastes their own key and shares the
target spreadsheet with the service account's email. We never store or log the key.
"""
from __future__ import annotations

import asyncio
import csv
import io
import json
import traceback

import httpx
from apify import Actor

from core import (
    DEFAULT_CHUNK_ROWS,
    DEFAULT_MAX_COLUMNS,
    FREE_ROWS_PER_RUN,
    append_idempotent,
    numberize,
    UserFacingError,
    billable_row_batches,
    build_table,
    chunked,
    escape_formula,
    friendly_google_error,
    google_status,
    parse_service_account,
    preflight_cells,
    redact,
    rows_per_chunk,
    secrets_from,
    with_backoff,
)

ROWS_EVENT = "rows-1k"  # $0.25 per 1k rows beyond the free tier (price set in Apify Console)
SHEETS_EVENT = "sheets-write"  # $0.01 per run that writes to Google Sheets


async def load_records(inp: dict) -> list:
    dataset_id = inp.get("datasetId")
    url = inp.get("sourceUrl")
    limit = int(inp.get("maxRows") or 1_000_000)
    if bool(dataset_id) == bool(url):
        raise UserFacingError("Provide exactly one source: 'Dataset ID' or 'Source URL'.")
    records: list = []
    if dataset_id:
        ds = await Actor.open_dataset(id=dataset_id, force_cloud=True)
        async for item in ds.iterate_items(limit=limit):
            records.append(item)
        return records
    async with httpx.AsyncClient(timeout=60, follow_redirects=True) as client:
        r = await client.get(url)
    if r.status_code != 200:
        raise UserFacingError(f"Could not download the source URL (HTTP {r.status_code}).")
    text = r.text
    stripped = text.lstrip()
    if stripped.startswith("[") or stripped.startswith("{"):
        try:
            data = json.loads(text)
            records = data if isinstance(data, list) else [data]
        except json.JSONDecodeError:  # maybe JSON Lines
            records = [json.loads(line) for line in text.splitlines() if line.strip()]
    else:
        records = list(csv.DictReader(io.StringIO(text)))
    return records[:limit]


def write_sheet(sa: dict, spreadsheet_id: str, sheet_name: str, mode: str,
                header: list, rows: list, chunk_rows: int) -> int:
    import gspread  # imported lazily so file-only runs don't need Google libs

    import requests

    def retryable(e: Exception) -> bool:
        if isinstance(e, (requests.exceptions.ConnectionError, requests.exceptions.Timeout)):
            return True
        code = getattr(getattr(e, "response", None), "status_code", None)
        return code in (429, 500, 502, 503, 504)

    gc = gspread.service_account_from_dict(sa)
    sh = with_backoff(lambda: gc.open_by_key(spreadsheet_id), retryable)
    worksheets = with_backoff(sh.worksheets, retryable)
    target = next((w for w in worksheets if w.title == sheet_name), None)
    other_cells = sum(w.row_count * w.col_count for w in worksheets if w is not target)
    per_request = rows_per_chunk(len(header), chunk_rows)

    if target is not None and mode == "append":
        existing_header = with_backoff(lambda: target.row_values(1), retryable)
        if existing_header and existing_header != [str(h) for h in header]:
            raise UserFacingError(
                f"Append aborted: the '{sheet_name}' tab has different columns than this export "
                f"(tab starts with {existing_header[:5]}, export starts with {header[:5]}). "
                "Use 'Replace' mode or a new tab name so columns don't get misaligned."
            )
        # Grid size is a safe upper bound for cells already used in the tab.
        preflight_cells(len(rows), len(header), other_cells + target.row_count * target.col_count)
        if not existing_header:
            with_backoff(lambda: target.update([header], "A1", value_input_option="RAW"), retryable)

        def rate_limited(e: Exception) -> bool:
            return getattr(getattr(e, "response", None), "status_code", None) == 429

        def ambiguous(e: Exception) -> bool:  # the request may have been applied
            return retryable(e) and not rate_limited(e)

        def count_rows() -> int:
            # Rows with any content (header included). Only called once up front and
            # after an ambiguous failure, never per chunk. Known edge case: a row whose
            # every cell is empty is not counted by get_all_values(), which could make an
            # applied chunk look partial. That gives a safe stop with an error, never a
            # duplicate. Exports of real records essentially never contain such rows.
            return len(with_backoff(target.get_all_values, retryable))

        table_rows = count_rows()
        written = 0
        for chunk in chunked(rows, per_request):
            # append_rows finds the true end of the table (blank first cells included)
            # and inserts rows, so existing data can never be overwritten. It is not
            # idempotent, so retries go through append_idempotent (no duplicates).
            append_idempotent(
                lambda c=chunk: target.append_rows(
                    list(c), value_input_option="RAW", insert_data_option="INSERT_ROWS", table_range="A1"),
                count_rows, table_rows, len(chunk), rate_limited, ambiguous,
            )
            table_rows += len(chunk)
            written += len(chunk)
        return written

    preflight_cells(len(rows), len(header), other_cells)
    total_rows = len(rows) + 1
    if target is None:
        target = with_backoff(lambda: sh.add_worksheet(
            title=sheet_name, rows=max(total_rows, 1), cols=max(len(header), 1)), retryable)
    else:
        with_backoff(target.clear, retryable)
        with_backoff(lambda: target.resize(rows=max(total_rows, 1), cols=max(len(header), 1)), retryable)

    with_backoff(lambda: target.update([header], "A1", value_input_option="RAW"), retryable)
    written = 0
    for chunk in chunked(rows, per_request):
        r0 = 2 + written
        with_backoff(lambda c=chunk, r=r0: target.update(list(c), f"A{r}", value_input_option="RAW"), retryable)
        written += len(chunk)
    return written


def to_xlsx(header: list, rows: list) -> bytes:
    from openpyxl import Workbook

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("data")
    ws.append([escape_formula(h) for h in header])
    for row in rows:
        ws.append([escape_formula(v) for v in row])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def to_csv(header: list, rows: list) -> bytes:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow([escape_formula(h) for h in header])
    w.writerows([escape_formula(v) for v in row] for row in rows)
    return buf.getvalue().encode("utf-8-sig")  # BOM so Excel opens UTF-8 correctly


async def main() -> None:
    async with Actor:
        inp = await Actor.get_input() or {}
        sa: dict | None = None
        try:
            spreadsheet_id = (inp.get("spreadsheetId") or "").strip()
            if spreadsheet_id:
                sa = parse_service_account(inp.get("serviceAccountKey"))
            records = await load_records(inp)
            if not records:
                raise UserFacingError("The source contains no rows, so there is nothing to export.")
            header, rows, dropped = build_table(records, int(inp.get("maxColumns") or DEFAULT_MAX_COLUMNS))
            if dropped:
                Actor.log.warning(f"{len(dropped)} columns over the column limit were dropped, e.g. {dropped[:5]}")
            # CSV keeps every value exactly as given; XLSX and Sheets get real numbers.
            typed_rows = numberize(rows) if inp.get("convertNumbers", True) else rows

            store = await Actor.open_key_value_store()
            formats = inp.get("outputFormats") or ["csv"]
            if "csv" in formats:
                await store.set_value("OUTPUT.csv", to_csv(header, rows), content_type="text/csv")
            if "xlsx" in formats:
                await store.set_value(
                    "OUTPUT.xlsx", to_xlsx(header, typed_rows),
                    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            written_to_sheet = 0
            if spreadsheet_id:
                try:
                    written_to_sheet = await asyncio.to_thread(
                        write_sheet, sa, spreadsheet_id, inp.get("sheetName") or "Export",
                        inp.get("mode") or "replace", header, typed_rows,
                        int(inp.get("chunkRows") or DEFAULT_CHUNK_ROWS),
                    )
                except UserFacingError:
                    raise
                except Exception as e:  # gspread.exceptions.APIError and friends
                    status = google_status(e)
                    raise UserFacingError(friendly_google_error(status, str(e), sa["client_email"])) from None

            exported = max(written_to_sheet, len(rows))
            events = billable_row_batches(exported, FREE_ROWS_PER_RUN)  # never from input
            charges = [(ROWS_EVENT, events)] + ([(SHEETS_EVENT, 1)] if written_to_sheet else [])
            for event_name, count in charges:
                if not count:
                    continue
                result = await Actor.charge(event_name=event_name, count=count)
                if getattr(result, "event_charge_limit_reached", False):
                    Actor.log.warning(
                        "Your maximum charge per run was reached, so part of this export was not billed. "
                        "Raise the run's max charge if you rely on larger exports.")

            summary = {"rows": len(rows), "columns": len(header), "rowsWrittenToSheet": written_to_sheet,
                       "droppedColumns": len(dropped), "formats": formats}
            await store.set_value("SUMMARY", summary)
            await Actor.set_status_message(f"Exported {len(rows):,} rows x {len(header)} columns.")
        except UserFacingError as e:
            await Actor.fail(status_message=redact(str(e), secrets_from(sa)))
        except Exception:  # noqa: BLE001 - last resort, never leak the key
            Actor.log.error(redact(traceback.format_exc(), secrets_from(sa)))
            await Actor.fail(status_message="Unexpected error. See the log; your key was not logged.")
