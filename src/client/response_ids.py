"""Creates the response IDs ("ID tanggapan") of the records, and remembers the last one.

An ID is the submission date and a counter: "261106-001", "261106-002", ... The counter
continues after the last ID in last-response.json when it has the same date, and starts
at 001 on a new date. The file is updated once the rows are in the master data, so the
master data never has to be read to number the next report.

last-response.json:
    {
        "last_response_id": "261106-003",
        "saved_at": "2026-11-06T15:16:12+07:00"
    }
"""

import json
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

from client.errors import ConfigurationError
from client.master_data_mapping import RESPONSE_ID_KEY

LAST_RESPONSE_FILE_NAME = "last-response.json"
LAST_RESPONSE_ID_KEY = "last_response_id"
SAVED_AT_KEY = "saved_at"
RESPONSE_ID_PATTERN = re.compile(r"(\d{6})-(\d{3,})")


def load_last_response_id(last_response_file: Path) -> str | None:
    """Returns the last response ID, or None if no report has been appended yet."""
    if not last_response_file.exists():
        return None
    try:
        values = json.loads(last_response_file.read_text(encoding="utf-8-sig"))
        response_id = values[LAST_RESPONSE_ID_KEY]
        if not isinstance(response_id, str) or RESPONSE_ID_PATTERN.fullmatch(response_id.strip()) is None:
            raise ValueError(f"{response_id!r} isn't an ID such as \"261106-001\"")
    except (OSError, ValueError, KeyError, TypeError) as error:
        # Don't guess: starting again at 001 could reuse IDs that are already in the master data.
        raise ConfigurationError(
            f"{last_response_file} is damaged ({error}). Set {LAST_RESPONSE_ID_KEY} to the "
            "last ID tanggapan in the master data, such as \"261106-003\"."
        ) from error
    return response_id.strip()


def assign_response_ids(
    records: list[dict[str, Any]],
    last_response_id: str | None,
    submission_date: date,
) -> list[dict[str, Any]]:
    """Returns the records, each with its response ID as the first field."""
    prefix = submission_date.strftime("%y%m%d")
    last_number = 0
    if last_response_id is not None:
        match = RESPONSE_ID_PATTERN.fullmatch(last_response_id)
        if match is not None and match[1] == prefix:
            last_number = int(match[2])

    return [
        {RESPONSE_ID_KEY: f"{prefix}-{last_number + offset:03d}", **record}
        for offset, record in enumerate(records, start=1)
    ]


def save_last_response_id(last_response_file: Path, response_id: str, saved_at: datetime) -> None:
    """Saves the last response ID. A temporary file first means a crash can't leave a
    half-written file."""
    content = {LAST_RESPONSE_ID_KEY: response_id, SAVED_AT_KEY: saved_at.isoformat()}
    temporary_file = last_response_file.with_name(f"{last_response_file.name}.tmp")
    temporary_file.write_text(json.dumps(content, indent=4) + "\n", encoding="utf-8")
    temporary_file.replace(last_response_file)
