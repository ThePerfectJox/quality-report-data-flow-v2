"""Appends the server's quality report to the Excel master data, then archives it. Runs on the client.

Data flow:
1. run_client.py receives the server's response and passes it to append_response_to_database().
   A saved response file can be appended with: python src/run_append_to_database.py [file]
2. client/report_conversion.py turns each sample of the summary_report view into a record,
   with the averages, the in-spec fractions (client/inspec.py, using the requirement of
   the sample's brand in config/client-config.json), its Group and Brand from the view, and the
   server's PLANT, LOKASI and BAGIAN from "server_config" in the response. Each record
   gets its response ID, continuing after the last one in last-response.json
   (client/response_ids.py).
3. client/excel_database.py appends the records, with their IDs, to the master data Excel
   Table set in MASTER_DATA of config/client-config.json, using the columns in
   client/master_data_mapping.py.
4. last-response.json is updated with the last ID appended.
5. client/archive.py writes an archive workbook with every view, calculated and metadata
   field: one per Set Sample in the report.

Errors are AppendDatabaseError subclasses, in client/errors.py. A PostAppendError means
the rows are already in the master data.
"""

import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from client.archive import build_archive_base_path, build_archive_root, create_archive_workbook, resolve_archive_path
from client.configuration import (
    CLIENT_CONFIG_FILE,
    ENV_FILE,
    PROJECT_ROOT,
    ClientConfig,
    load_client_configuration,
    read_report_json_path,
    resolve_project_path,
    validate_workbook_path,
    warn_about_moved_env_settings,
)
from client.errors import (
    AppendDatabaseError,
    ArchiveCreationError,
    ConfigurationError,
    DuplicateRecordError,
    LastResponseSaveError,
    PostAppendError,
)
from client.excel_database import (
    append_rows_to_excel_table,
    find_duplicate_records,
    find_used_response_ids,
    map_data_to_excel_columns,
    open_excel_table,
    read_table_rows,
    save_workbook_atomically,
    validate_excel_columns,
)
from client.master_data_mapping import INSPECTED_AT_KEY, RESPONSE_ID_KEY, SET_SAMPLE_KEY, WORKER_ID_KEY
from client.report_conversion import load_input_data, prepare_records, validate_input_structure
from client.response_ids import (
    LAST_RESPONSE_FILE_NAME,
    LAST_RESPONSE_ID_KEY,
    load_last_response_id,
    save_last_response_id,
)
from client.values import get_local_timezone
from managers.environment_manager import ENV_REPORT_JSON_PATH

LOGGER = logging.getLogger(__name__)

__all__ = [
    "AppendDatabaseError",
    "AppendResult",
    "ArchiveCreationError",
    "LastResponseSaveError",
    "PostAppendError",
    "append_data_to_excel_database",
    "append_response_to_database",
]


@dataclass(frozen=True)
class AppendResult:
    """What the append did."""

    database_path: Path
    worksheet_name: str | None
    table_name: str | None
    appended_row_count: int
    first_row: int | None
    last_row: int | None
    # One archive per Set Sample in the report, in the order of set_samples.
    archive_paths: list[Path]
    # The Set Samples of the appended rows, from the view, lowest first.
    set_samples: list[int]
    submission_datetime: datetime


def append_data_to_excel_database(
    input_json_path: Path | str | None = None,
    *,
    env_file: Path = ENV_FILE,
    client_config_file: Path = CLIENT_CONFIG_FILE,
    project_root: Path = PROJECT_ROOT,
    current_time: datetime | None = None,
) -> AppendResult:
    """Appends a saved report JSON file to the master data table, then writes the archive.

    Reads input_json_path, or REPORT_JSON_PATH from .env when it's not given. Raises an
    AppendDatabaseError subclass on failure. A PostAppendError means the master data
    append did succeed.
    """
    client_config = _load_settings(env_file, client_config_file, project_root)
    if input_json_path is not None:
        json_path = resolve_project_path(input_json_path, project_root)
    else:
        json_path = read_report_json_path(env_file, project_root)
        if json_path is None:
            raise ConfigurationError(
                f"No report JSON was given. Set {ENV_REPORT_JSON_PATH} in {env_file}, "
                "or pass the file path."
            )

    LOGGER.info("Reading %s", json_path)
    response = load_input_data(json_path)
    return _append_validated_response(
        response, client_config, current_time, project_root / LAST_RESPONSE_FILE_NAME
    )


def append_response_to_database(
    response: dict[str, Any],
    *,
    env_file: Path = ENV_FILE,
    client_config_file: Path = CLIENT_CONFIG_FILE,
    project_root: Path = PROJECT_ROOT,
    current_time: datetime | None = None,
) -> AppendResult:
    """Appends a server response, as returned by send_request(), to the master data table.

    Works like append_data_to_excel_database(), but takes the response directly instead
    of reading a JSON file, so REPORT_JSON_PATH isn't needed. The response must have the
    same structure as the saved JSON: "status", "responded_at" and "report.samples".
    """
    client_config = _load_settings(env_file, client_config_file, project_root)
    validate_input_structure(response, "The server response")
    return _append_validated_response(
        response, client_config, current_time, project_root / LAST_RESPONSE_FILE_NAME
    )


def _load_settings(env_file: Path, client_config_file: Path, project_root: Path) -> ClientConfig:
    """Loads config/client-config.json and checks the master data workbook exists."""
    warn_about_moved_env_settings(env_file)
    client_config = load_client_configuration(client_config_file, project_root)
    validate_workbook_path(client_config.excel.workbook_path)
    return client_config


def _append_validated_response(
    response: dict[str, Any],
    client_config: ClientConfig,
    current_time: datetime | None,
    last_response_file: Path,
) -> AppendResult:
    """Appends a validated response to the master data table, then writes the archive.

    Order: prepare and validate every record and give it its response ID, check the
    workbook and table, append and save the master data atomically, save the last
    response ID to last-response.json, then create the archives.
    """
    excel_config = client_config.excel
    local_timezone = get_local_timezone()
    last_response_id = load_last_response_id(last_response_file)
    prepared = prepare_records(response, client_config, local_timezone, current_time, last_response_id)
    # Checked now, so an unusable plant name stops the run before the master data changes.
    archive_root = build_archive_root(excel_config.workbook_path.parent, prepared.plant)
    LOGGER.info(
        "Submitted at %s, %d samples, brands %s",
        prepared.submission_datetime.isoformat(),
        len(prepared.records),
        ", ".join(prepared.brands) or "none",
    )

    if not prepared.records:
        LOGGER.info("The report has no samples. Nothing was appended or archived.")
        return AppendResult(
            excel_config.workbook_path, excel_config.worksheet_name, excel_config.table_name,
            0, None, None, [], [], prepared.submission_datetime,
        )

    target = open_excel_table(excel_config.workbook_path, excel_config.worksheet_name, excel_config.table_name)
    resolved_columns = validate_excel_columns(target.headers)
    existing_rows = read_table_rows(target)

    duplicates = find_duplicate_records(prepared.records, existing_rows, resolved_columns)
    if duplicates:
        examples = ", ".join(
            f"worker {record[WORKER_ID_KEY]} at {record[INSPECTED_AT_KEY]}" for record in duplicates[:5]
        )
        raise DuplicateRecordError(
            f"{len(duplicates)} of {len(prepared.records)} samples are already in "
            f"{target.table.displayName} (for example {examples}). This report may have "
            "been appended before. Nothing was written."
        )

    # The IDs come from last-response.json. This only catches a file that is behind the table.
    used_ids = find_used_response_ids(
        prepared.records, (row.get(resolved_columns[RESPONSE_ID_KEY]) for row in existing_rows)
    )
    if used_ids:
        raise DuplicateRecordError(
            f"The response IDs {', '.join(used_ids[:5])} are already in {target.table.displayName}, "
            f"so {last_response_file} is behind the master data. Set {LAST_RESPONSE_ID_KEY} in it "
            "to the last ID tanggapan in the master data. Nothing was written."
        )

    records = prepared.records
    excel_rows = [map_data_to_excel_columns(record, target.headers) for record in records]
    first_row, last_row = append_rows_to_excel_table(target, excel_rows)
    save_workbook_atomically(target)
    LOGGER.info(
        "Appended %d rows (worksheet rows %d-%d, IDs %s to %s) to %s!%s in %s",
        len(records), first_row, last_row, records[0][RESPONSE_ID_KEY], records[-1][RESPONSE_ID_KEY],
        target.worksheet.title, target.table.displayName, target.workbook_path,
    )
    appended = f"The master data append succeeded ({len(records)} rows added to {target.workbook_path})"

    # Saved before the archives, because the next report's IDs depend on it.
    last_id = records[-1][RESPONSE_ID_KEY]
    save_error: OSError | None = None
    try:
        save_last_response_id(last_response_file, last_id, datetime.now(local_timezone))
        LOGGER.info("Saved the last response ID %s to %s", last_id, last_response_file)
    except OSError as error:
        save_error = error

    # One archive per Set Sample, so each file name has the Set Sample of all its rows.
    records_by_set_sample: dict[int, list[dict[str, Any]]] = {}
    for record in records:
        records_by_set_sample.setdefault(record[SET_SAMPLE_KEY], []).append(record)

    archive_paths: list[Path] = []
    archive_error: Exception | None = None
    try:
        for set_sample, set_records in sorted(records_by_set_sample.items()):
            archive_path = resolve_archive_path(
                build_archive_base_path(archive_root, prepared.submission_datetime, set_sample)
            )
            create_archive_workbook(set_records, prepared.column_labels, archive_path)
            archive_paths.append(archive_path)
            LOGGER.info("Archive of Set Sample %d saved to %s", set_sample, archive_path)
    except (AppendDatabaseError, OSError, ValueError) as error:
        archive_error = error
    created = ", ".join(str(path) for path in archive_paths) or "none"

    if save_error is not None:
        archive_status = f"Archives created: {created}." if archive_error is None else (
            f"An archive also couldn't be created: {archive_error}. Archives created: {created}."
        )
        raise LastResponseSaveError(
            f"{appended}, but {last_response_file} couldn't be updated: {save_error}. Set "
            f"{LAST_RESPONSE_ID_KEY} in it to {last_id!r} before the next append, or the next "
            f"report will be rejected. {archive_status} Don't run the append again; the rows "
            "are already in the master data.",
            target.workbook_path,
            len(records),
        ) from save_error
    if archive_error is not None:
        raise ArchiveCreationError(
            f"{appended}, but an archive couldn't be created: {archive_error}. Archives created: "
            f"{created}. Don't run the append again; the rows are already in the master data.",
            target.workbook_path,
            len(records),
        ) from archive_error

    return AppendResult(
        target.workbook_path, target.worksheet.title, target.table.displayName,
        len(records), first_row, last_row, archive_paths,
        sorted(records_by_set_sample), prepared.submission_datetime,
    )
