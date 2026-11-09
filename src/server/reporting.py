"""Everything about the report: which samples it contains, and remembering what was sent.

Each report only contains samples dated after the newest sample in the last report sent.
That date is saved in last_report_sent.json, once the report has actually been sent.
"""

import threading
from datetime import date, datetime, timezone
from pathlib import Path

from managers.mysql_manager import MySqlManager
from shared.json_serialization import from_json, to_json

LAST_REPORT_SENT_FILE = Path(__file__).resolve().parents[2] / "last_report_sent.json"

SAMPLE_COLUMNS = "id, name, sample_date, quantity, score, is_passed, checked_at"
SELECT_ALL_SAMPLES_QUERY = f"SELECT {SAMPLE_COLUMNS} FROM samples ORDER BY sample_date, id"
SELECT_NEW_SAMPLES_QUERY = (
    f"SELECT {SAMPLE_COLUMNS} FROM samples WHERE sample_date > %s ORDER BY sample_date, id"
)

# Each client has its own thread, so only one at a time may use the file.
_last_report_file_lock = threading.Lock()


class LastReportFileError(Exception):
    """last_report_sent.json can't be used. The message explains what to do."""


# region CREATING THE REPORT
def create_report() -> dict:
    """Returns a report of the samples that haven't been sent yet."""
    sent_up_to = _get_last_sample_date_sent()
    samples = _fetch_samples_after(sent_up_to)
    return {
        "samples_after": sent_up_to,
        "sample_count": len(samples),
        "samples": samples,
    }


def _fetch_samples_after(sent_up_to: date | None) -> list[dict]:
    with MySqlManager.from_environment() as database:
        if sent_up_to is None:
            return database.fetch_all(SELECT_ALL_SAMPLES_QUERY)
        return database.fetch_all(SELECT_NEW_SAMPLES_QUERY, (sent_up_to,))


# endregion


# region REMEMBERING WHAT WAS SENT
def mark_report_sent(report: dict):
    """Saves that this report was sent. A report with no samples changes nothing."""
    samples = report["samples"]
    if not samples:
        return

    newest_sample_date = max(sample["sample_date"] for sample in samples)
    with _last_report_file_lock:
        previous = _read_last_report_file()
        # Never move backwards, in case two clients' reports finish out of order.
        if previous is not None:
            newest_sample_date = max(newest_sample_date, previous["last_sample_date"])
        _write_last_report_file(
            {
                "last_sent_at": datetime.now(timezone.utc),
                "last_sample_date": newest_sample_date,
                "sample_count": len(samples),
            }
        )


def _get_last_sample_date_sent() -> date | None:
    with _last_report_file_lock:
        last_report = _read_last_report_file()
    return None if last_report is None else last_report["last_sample_date"]


def _read_last_report_file() -> dict | None:
    """Returns the file's contents, or None if no report has been sent yet."""
    if not LAST_REPORT_SENT_FILE.exists():
        return None
    try:
        last_report = from_json(LAST_REPORT_SENT_FILE.read_text(encoding="utf-8"))
        return {
            **last_report,
            "last_sample_date": date.fromisoformat(last_report["last_sample_date"]),
        }
    except (ValueError, KeyError, TypeError) as error:
        # Don't guess: treating this as "nothing sent yet" would resend every sample.
        raise LastReportFileError(
            f"{LAST_REPORT_SENT_FILE} is damaged. Fix it, or delete it to send all samples again."
        ) from error


def _write_last_report_file(last_report: dict):
    # Writing a temporary file first means a crash can't leave a half-written file.
    temporary_file = LAST_REPORT_SENT_FILE.with_suffix(".json.tmp")
    temporary_file.write_text(to_json(last_report, indent=2) + "\n", encoding="utf-8")
    temporary_file.replace(LAST_REPORT_SENT_FILE)


# endregion
