import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import core  # noqa: E402

SA = {
    "type": "service_account",
    "client_email": "bot@proj.iam.gserviceaccount.com",
    "private_key": "-----BEGIN PRIVATE KEY-----\nABC123\n-----END PRIVATE KEY-----\n",
    "private_key_id": "kid42",
    "token_uri": "https://oauth2.googleapis.com/token",
}


def test_parse_service_account_ok_from_string():
    assert core.parse_service_account(json.dumps(SA))["client_email"] == SA["client_email"]


@pytest.mark.parametrize("raw", [None, "", "not json", "[]", json.dumps({"type": "authorized_user"})])
def test_parse_service_account_rejects(raw):
    with pytest.raises(core.UserFacingError) as e:
        core.parse_service_account(raw)
    assert "ABC123" not in str(e.value)


def test_parse_service_account_missing_fields_never_echoes_key():
    bad = dict(SA, token_uri="")
    with pytest.raises(core.UserFacingError) as e:
        core.parse_service_account(json.dumps(bad))
    assert "token_uri" in str(e.value) and "ABC123" not in str(e.value)


def test_redact_removes_secrets_and_pem():
    text = f"failed for {SA['client_email']} key {SA['private_key']} id kid42"
    out = core.redact(text, core.secrets_from(SA))
    assert "ABC123" not in out and "kid42" not in out
    assert "bot@proj" in out  # the email is needed in user-facing fix instructions (M1)


def test_flatten_nested_and_lists():
    rec = {"a": 1, "b": {"c": 2, "d": {"e": None}}, "tags": ["x", "y"], "ok": True}
    assert core.flatten(rec) == {"a": 1, "b.c": 2, "b.d.e": "", "tags": '["x", "y"]', "ok": "TRUE"}


def test_flatten_respects_max_depth():
    rec = {"a": {"b": {"c": 1}}}
    # max_depth counts dict levels expanded, including the top-level record
    assert core.flatten(rec, max_depth=2) == {"a.b": '{"c": 1}'}
    assert core.flatten(rec, max_depth=1) == {"a": '{"b": {"c": 1}}'}


def test_long_strings_are_cut_to_sheets_limit():
    assert len(core.to_cell("x" * 60_000)) < 50_000


def test_build_table_union_of_keys_in_first_seen_order():
    header, rows, dropped = core.build_table([{"a": 1}, {"b": 2, "a": 3}])
    assert header == ["a", "b"] and rows == [[1, ""], [3, 2]] and dropped == []


def test_build_table_column_cap():
    header, rows, dropped = core.build_table([{f"k{i}": i for i in range(5)}], max_columns=3)
    assert header == ["k0", "k1", "k2"] and dropped == ["k3", "k4"] and rows == [[0, 1, 2]]


def test_preflight_cells():
    core.preflight_cells(999, 10)  # 10k cells fine
    with pytest.raises(core.UserFacingError):
        core.preflight_cells(1_000_000, 10)  # 10,000,010 > 10M
    with pytest.raises(core.UserFacingError):
        core.preflight_cells(10, 10, other_cells_in_spreadsheet=9_999_950)


def test_chunked():
    assert [list(c) for c in core.chunked(list(range(5)), 2)] == [[0, 1], [2, 3], [4]]
    with pytest.raises(ValueError):
        list(core.chunked([1], 0))


def test_with_backoff_retries_then_succeeds():
    calls, sleeps = [], []

    def fn():
        calls.append(1)
        if len(calls) < 3:
            raise RuntimeError("429")
        return "ok"

    assert core.with_backoff(fn, lambda e: "429" in str(e), sleep=sleeps.append) == "ok"
    assert sleeps == [2.0, 4.0]


def test_with_backoff_does_not_retry_fatal():
    sleeps = []
    with pytest.raises(ValueError):
        core.with_backoff(lambda: (_ for _ in ()).throw(ValueError("403")), lambda e: False, sleep=sleeps.append)
    assert sleeps == []


@pytest.mark.parametrize(
    "written,free,expected",
    [(0, 10_000, 0), (10_000, 10_000, 0), (10_001, 10_000, 1), (12_500, 10_000, 3), (5_000, 0, 5)],
)
def test_billable_row_batches(written, free, expected):
    assert core.billable_row_batches(written, free) == expected


def test_friendly_errors_name_the_service_account():
    msg = core.friendly_google_error(403, "The caller does not have permission", SA["client_email"])
    assert SA["client_email"] in msg and "Editor" in msg
    assert "enable" in core.friendly_google_error(403, "API has not been used in project", "x").lower()
    assert "not found" in core.friendly_google_error(404, "Requested entity was not found.", "x").lower()


def test_friendly_error_bad_key():
    assert "fresh json key" in core.friendly_google_error(None, "Unable to load PEM file", "x").lower()
    assert "fresh json key" in core.friendly_google_error(
        None, "Could not deserialize key data. Details: ASN.1 parsing error: invalid length", "x").lower()


def test_free_rows_is_a_constant_not_input():
    assert core.FREE_ROWS_PER_RUN == 2_000


def test_rows_per_chunk_bounded_by_cells():
    assert core.rows_per_chunk(500) == 200
    assert core.rows_per_chunk(4) == core.DEFAULT_CHUNK_ROWS
    assert core.rows_per_chunk(1_000_000) == 1


@pytest.mark.parametrize("raw,expected", [
    ("=HYPERLINK(\"x\")", "'=HYPERLINK(\"x\")"), ("+cmd", "'+cmd"), ("@SUM(A1)", "'@SUM(A1)"),
    ("-5", "-5"), ("+3.2", "+3.2"), ("-1e5", "-1e5"), ("- bullet", "'- bullet"), ("hello", "hello"), (7, 7), ("", ""),
])
def test_escape_formula(raw, expected):
    assert core.escape_formula(raw) == expected


class FakeSheet:
    """Applies appends, optionally failing after (or before) applying."""

    def __init__(self, rows=0, fail_plan=()):
        self.rows, self.plan, self.calls = rows, list(fail_plan), 0

    def append(self, n):
        self.calls += 1
        action = self.plan.pop(0) if self.plan else "ok"
        if action == "apply_then_timeout":
            self.rows += n
            raise TimeoutError("read timed out")
        if action == "timeout_before":
            raise TimeoutError("connect timed out")
        if action == "429":
            raise RuntimeError("429")
        if action == "partial_then_timeout":
            self.rows += n // 2
            raise TimeoutError("read timed out")
        self.rows += n


def _run(sheet, n=10):
    before = sheet.rows
    core.append_idempotent(
        lambda: sheet.append(n), lambda: sheet.rows, before, n,
        is_rate_limited=lambda e: "429" in str(e),
        is_ambiguous=lambda e: isinstance(e, TimeoutError),
        sleep=lambda s: None,
    )


def test_append_applied_then_timeout_is_not_duplicated():
    s = FakeSheet(rows=5, fail_plan=["apply_then_timeout"])
    _run(s)
    assert s.rows == 15 and s.calls == 1


def test_append_timeout_before_apply_is_retried():
    s = FakeSheet(rows=5, fail_plan=["timeout_before"])
    _run(s)
    assert s.rows == 15 and s.calls == 2


def test_append_429_is_retried():
    s = FakeSheet(fail_plan=["429", "429"])
    _run(s)
    assert s.rows == 10 and s.calls == 3


def test_append_partial_write_stops_with_clear_error():
    s = FakeSheet(fail_plan=["partial_then_timeout"])
    with pytest.raises(core.UserFacingError):
        _run(s)
    assert s.calls == 1


def test_as_number_converts_only_plain_numbers():
    assert core.as_number("22") == 22 and isinstance(core.as_number("22"), int)
    assert core.as_number("-7.25") == -7.25
    assert core.as_number("0") == 0 and core.as_number("0.5") == 0.5
    for keep in ["02134", "1e5", "+3", " 5", "1,000", "", "abc", "5.", ".5", "=1+1"]:
        assert core.as_number(keep) == keep
    assert core.as_number(7) == 7 and core.as_number(None) is None


def test_n1_decimals_over_15_significant_digits_stay_text():
    v = "123456789012345.123456789012345"
    assert core.as_number(v) == v
    assert core.as_number("12345678.1234567") == 12345678.1234567  # 15 sig digits: ok
    assert core.as_number("0.000000000000001") == 1e-15  # leading zeros aren't significant


def test_n2_integers_over_11_digits_stay_text():
    assert core.as_number("12345678901") == 12345678901
    assert core.as_number("4006381333931") == "4006381333931"  # 13-digit EAN barcode


def test_n3_mixed_column_stays_text_whole():
    rows = [["10023", "1.5", "", 7], ["A-77", "2", "", "8"], ["10024", "", "", ""]]
    out = core.numberize(rows)
    assert [r[0] for r in out] == ["10023", "A-77", "10024"]  # mixed SKU column untouched
    assert [r[1] for r in out] == [1.5, 2, ""]  # blanks don't block conversion
    assert [r[2] for r in out] == ["", "", ""]  # all-empty column untouched
    assert [r[3] for r in out] == [7, 8, ""]  # already-numeric values count as numeric
    assert rows[1][1] == "2"  # input not modified
    assert core.numberize([["TRUE"], [True]]) == [["TRUE"], [True]]


def test_n4_negative_zero_stays_text_and_option_exists():
    assert core.as_number("-0") == "-0" and core.as_number("-0.00") == "-0.00"
    import json, pathlib
    schema = json.loads((pathlib.Path(__file__).parents[1] / ".actor/input_schema.json").read_text())
    assert schema["properties"]["convertNumbers"]["default"] is True
