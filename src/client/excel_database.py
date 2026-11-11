"""Appends rows to the Excel Table of the master data workbook, and saves it safely.

Columns are matched by name using client/master_data_mapping.py. The workbook is saved
through a temporary file, so a failed save can't damage it.
"""

import logging
import os
import tempfile
import zipfile
from copy import copy
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet

from client.errors import (
    ExcelDatabaseError,
    MissingExcelColumnError,
    TableHeaderError,
    TableNotFoundError,
    WorkbookLockedError,
    WorkbookSaveError,
    WorksheetNotFoundError,
)
from client.master_data_mapping import (
    DUPLICATE_CHECK_KEYS,
    DUPLICATE_FLOAT_DECIMALS,
    MASTER_DATA_COLUMN_MAPPING,
    REQUIRED_MASTER_DATA_KEYS,
    RESPONSE_ID_KEY,
)
from client.values import is_missing, normalize_header_name, to_nullable_float, write_cell_value
LOGGER = logging.getLogger(__name__)

# Where the worksheet and table names are set, for error messages.
SHEET_SETTING = "MASTER_DATA.SHEET_NAME in config/client-config.json"
TABLE_SETTING = "MASTER_DATA.TABLE_NAME in config/client-config.json"


@dataclass
class ExcelTableTarget:
    """A loaded master data workbook and the Excel Table to append to."""

    workbook_path: Path
    workbook: Workbook
    worksheet: Worksheet
    table: Table
    headers: list[str]
    file_signature: tuple[int, int]

    @property
    def bounds(self) -> tuple[int, int, int, int]:
        """Returns (first column, header row, last column, last row) of the table."""
        return range_boundaries(self.table.ref)

    @property
    def first_data_row(self) -> int:
        return self.bounds[1] + 1


# region COLUMN MAPPING
def resolve_excel_columns(
    excel_headers: list[str],
    mapping: dict[str, str] = MASTER_DATA_COLUMN_MAPPING,
) -> tuple[dict[str, str], list[str]]:
    """Matches the mapping to the table headers by name.

    Returns (source key -> actual table header, source keys whose column is missing).
    """
    headers_by_name = {normalize_header_name(header): header for header in excel_headers}
    resolved: dict[str, str] = {}
    missing: list[str] = []
    for source_key, column_name in mapping.items():
        header = headers_by_name.get(normalize_header_name(column_name))
        if header is None:
            missing.append(source_key)
        else:
            resolved[source_key] = header
    return resolved, missing


def validate_excel_columns(
    excel_headers: list[str],
    mapping: dict[str, str] = MASTER_DATA_COLUMN_MAPPING,
    required_keys: Iterable[str] = REQUIRED_MASTER_DATA_KEYS,
) -> dict[str, str]:
    """Raises MissingExcelColumnError for missing required columns; warns about other ones."""
    resolved, missing = resolve_excel_columns(excel_headers, mapping)
    required = set(required_keys)
    missing_required = [key for key in missing if key in required]
    missing_optional = [key for key in missing if key not in required]

    if missing_required:
        columns = ", ".join(repr(mapping[key]) for key in missing_required)
        raise MissingExcelColumnError(f"The Excel Table is missing required columns: {columns}")
    for key in missing_optional:
        LOGGER.warning(
            "The Excel Table has no column %r, so %r isn't written.", mapping[key], key
        )
    return resolved


def map_data_to_excel_columns(
    data: dict[str, Any],
    excel_headers: list[str],
    mapping: dict[str, str] = MASTER_DATA_COLUMN_MAPPING,
) -> dict[str, Any]:
    """Returns {table header: value} for the mapped fields whose column is in the table."""
    resolved, _ = resolve_excel_columns(excel_headers, mapping)
    return {header: data.get(source_key) for source_key, header in resolved.items()}


# endregion


# region DATABASE WORKBOOK
def check_workbook_not_locked(workbook_path: Path) -> None:
    """Raises WorkbookLockedError if the workbook is open in Excel or can't be written."""
    lock_file = workbook_path.with_name(f"~${workbook_path.name}")
    if lock_file.exists():
        raise WorkbookLockedError(
            f"{workbook_path} is open in Excel. Close it and try again. If Excel isn't "
            f"running, delete the leftover lock file {lock_file}."
        )
    try:
        with workbook_path.open("r+b"):
            pass
    except PermissionError as error:
        raise WorkbookLockedError(
            f"{workbook_path} is open in another program or is read-only. Close it and try again."
        ) from error
    except OSError as error:
        raise ExcelDatabaseError(f"{workbook_path} can't be opened: {error}") from error


def open_excel_table(
    workbook_path: Path,
    worksheet_name: str | None,
    table_name: str | None,
) -> ExcelTableTarget:
    """Loads the workbook and finds and checks the Excel Table to append to."""
    check_workbook_not_locked(workbook_path)
    file_signature = _get_file_signature(workbook_path)
    try:
        workbook = load_workbook(workbook_path)
    except PermissionError as error:
        raise WorkbookLockedError(f"{workbook_path} is locked by another program.") from error
    except (InvalidFileException, zipfile.BadZipFile, KeyError, ValueError, OSError) as error:
        raise ExcelDatabaseError(f"{workbook_path} isn't a readable .xlsx workbook: {error}") from error

    worksheet, table = locate_excel_table(workbook, worksheet_name, table_name)
    if table.totalsRowCount:
        raise ExcelDatabaseError(
            f"The Excel Table {table.displayName} has a totals row. Turn it off in Excel "
            "(Table Design > Total Row) so rows can be appended."
        )
    headers = read_table_headers(worksheet, table)
    return ExcelTableTarget(workbook_path, workbook, worksheet, table, headers, file_signature)


def locate_excel_table(
    workbook: Workbook,
    worksheet_name: str | None,
    table_name: str | None,
) -> tuple[Worksheet, Table]:
    """Returns the configured table, or the only table when none is configured."""
    if worksheet_name is not None:
        if worksheet_name not in workbook.sheetnames:
            raise WorksheetNotFoundError(
                f"The worksheet {worksheet_name!r} isn't in the workbook. "
                f"Worksheets: {workbook.sheetnames}. Check {SHEET_SETTING}."
            )
        worksheets = [workbook[worksheet_name]]
    else:
        worksheets = list(workbook.worksheets)

    tables = [(sheet, table) for sheet in worksheets for table in sheet.tables.values()]
    available = [f"{sheet.title}!{table.displayName}" for sheet, table in tables]

    if table_name is not None:
        matches = [
            (sheet, table)
            for sheet, table in tables
            if table.displayName.casefold() == table_name.casefold()
        ]
        if not matches:
            raise TableNotFoundError(
                f"The Excel Table {table_name!r} wasn't found. Tables: {available or 'none'}. "
                f"Check {TABLE_SETTING} and {SHEET_SETTING}."
            )
        return matches[0]

    if not tables:
        raise TableNotFoundError(
            "No Excel Table was found. The data must be inside an Excel Table "
            "(select the data, then Insert > Table)."
        )
    if len(tables) > 1:
        raise TableNotFoundError(
            f"The workbook has several Excel Tables: {available}. Set {SHEET_SETTING} and "
            f"{TABLE_SETTING}."
        )
    return tables[0]


def read_table_headers(worksheet: Worksheet, table: Table) -> list[str]:
    """Returns the table's header row. Raises TableHeaderError for blank or duplicate names."""
    if not table.headerRowCount:
        raise TableHeaderError(f"The Excel Table {table.displayName} has no header row.")

    min_col, header_row, max_col, _ = range_boundaries(table.ref)
    headers: list[str] = []
    blank_columns: list[str] = []
    for column in range(min_col, max_col + 1):
        value = worksheet.cell(header_row, column).value
        if value is None or not str(value).strip():
            blank_columns.append(get_column_letter(column))
        headers.append("" if value is None else str(value).strip())

    if blank_columns:
        raise TableHeaderError(
            f"The Excel Table {table.displayName} has blank headers in columns: {blank_columns}"
        )

    seen: dict[str, str] = {}
    duplicates: list[str] = []
    for header in headers:
        name = normalize_header_name(header)
        if name in seen:
            duplicates.append(header)
        seen[name] = header
    if duplicates:
        raise TableHeaderError(
            f"The Excel Table {table.displayName} has duplicate headers: {duplicates}"
        )
    return headers


def read_table_rows(target: ExcelTableTarget) -> list[dict[str, Any]]:
    """Returns the table's existing data rows as {header: value}."""
    min_col, _, max_col, last_row = target.bounds
    if last_row < target.first_data_row:
        return []
    rows = target.worksheet.iter_rows(
        min_row=target.first_data_row,
        max_row=last_row,
        min_col=min_col,
        max_col=max_col,
        values_only=True,
    )
    return [dict(zip(target.headers, row)) for row in rows]


def find_duplicate_records(
    records: list[dict[str, Any]],
    existing_rows: list[dict[str, Any]],
    resolved_columns: dict[str, str],
) -> list[dict[str, Any]]:
    """Returns the records whose DUPLICATE_CHECK_KEYS all match an existing table row."""
    headers = [resolved_columns[key] for key in DUPLICATE_CHECK_KEYS]
    existing: set[tuple[Any, ...]] = set()
    for row in existing_rows:
        fingerprint = _fingerprint(row.get(header) for header in headers)
        if any(part is not None for part in fingerprint):
            existing.add(fingerprint)
    return [
        record
        for record in records
        if _fingerprint(record.get(key) for key in DUPLICATE_CHECK_KEYS) in existing
    ]


def find_used_response_ids(records: list[dict[str, Any]], existing_ids: Iterable[Any]) -> list[str]:
    """Returns the records' response IDs that are already in the table.

    The IDs are created from last-response.json, so this only finds a file that is behind
    the master data, for example after it was deleted or edited.
    """
    existing = {str(value).strip() for value in existing_ids if value is not None}
    return [record[RESPONSE_ID_KEY] for record in records if record[RESPONSE_ID_KEY] in existing]


def append_rows_to_excel_table(target: ExcelTableTarget, rows: list[dict[str, Any]]) -> tuple[int, int]:
    """Writes the rows below the table in memory and expands the table to include them.

    Each row is {table header: value}. Unmapped columns stay empty, or get the column's
    table formula if it has one. New cells copy the formatting of the last table row.
    Returns the first and last worksheet row written.
    """
    if not rows:
        raise ValueError("There are no rows to append.")

    worksheet, table = target.worksheet, target.table
    min_col, header_row, max_col, last_row = target.bounds
    start_row = _find_first_free_row(target)
    end_row = start_row + len(rows) - 1
    _ensure_range_is_free(target, start_row, end_row)

    template_row = last_row if last_row > header_row else None
    formulas = _get_calculated_column_formulas(table, target.headers)
    for header in rows[0]:
        if header in formulas:
            LOGGER.warning("Column %r has a table formula, but the mapped value replaces it.", header)

    for row_number, row in zip(range(start_row, end_row + 1), rows):
        for offset, header in enumerate(target.headers):
            cell = worksheet.cell(row_number, min_col + offset)
            if template_row is not None and template_row != row_number:
                cell._style = copy(worksheet.cell(template_row, min_col + offset)._style)
            if header in row:
                write_cell_value(cell, row[header])
            elif header in formulas:
                cell.value = formulas[header]

    new_ref = f"{get_column_letter(min_col)}{header_row}:{get_column_letter(max_col)}{max(end_row, last_row)}"
    table.ref = new_ref
    if table.autoFilter is not None:
        table.autoFilter.ref = new_ref
    return start_row, end_row


def save_workbook_atomically(target: ExcelTableTarget) -> None:
    """Saves to a temporary file, checks it, then replaces the workbook in one step.

    If anything fails, the original workbook is left unchanged.
    """
    path = target.workbook_path
    expected_ref = target.table.ref
    file_descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.stem}.", suffix=".tmp.xlsx", dir=path.parent
    )
    os.close(file_descriptor)
    temporary_path = Path(temporary_name)

    try:
        try:
            target.workbook.save(temporary_path)
        except (OSError, ValueError, TypeError) as error:
            raise WorkbookSaveError(f"Saving {path} failed: {error}. The original is unchanged.") from error

        _verify_saved_workbook(temporary_path, target.worksheet.title, target.table.displayName, expected_ref)

        if _get_file_signature(path) != target.file_signature:
            raise WorkbookSaveError(
                f"{path} was changed by another program while this script ran. "
                "Nothing was written. Run the script again."
            )
        check_workbook_not_locked(path)

        try:
            os.replace(temporary_path, path)
        except PermissionError as error:
            raise WorkbookLockedError(
                f"{path} couldn't be replaced because it's open in another program. "
                "Close it and try again. The original is unchanged."
            ) from error
        except OSError as error:
            raise WorkbookSaveError(f"Replacing {path} failed: {error}. The original is unchanged.") from error
    finally:
        if temporary_path.exists():
            try:
                temporary_path.unlink()
            except OSError as error:
                LOGGER.warning("The temporary file %s couldn't be deleted: %s", temporary_path, error)


def _find_first_free_row(target: ExcelTableTarget) -> int:
    """Returns the row to start writing at. A table whose only data row is empty uses that row."""
    _, header_row, _, last_row = target.bounds
    if last_row == header_row + 1 and all(value is None for value in read_table_rows(target)[0].values()):
        return last_row
    return last_row + 1


def _ensure_range_is_free(target: ExcelTableTarget, start_row: int, end_row: int) -> None:
    """Raises ExcelDatabaseError if writing the rows would overwrite other content."""
    worksheet, table = target.worksheet, target.table
    min_col, _, max_col, last_row = target.bounds
    if start_row <= last_row:
        return  # The single empty data row of a new table.

    for row in worksheet.iter_rows(min_row=start_row, max_row=end_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if cell.value is not None:
                raise ExcelDatabaseError(
                    f"Cell {cell.coordinate} below the table isn't empty. Move that content so "
                    f"{end_row - start_row + 1} rows can be added below {table.displayName}."
                )

    for other in worksheet.tables.values():
        if other is table:
            continue
        other_min_col, other_min_row, other_max_col, other_max_row = range_boundaries(other.ref)
        if other_min_col <= max_col and other_max_col >= min_col and other_min_row <= end_row and other_max_row >= start_row:
            raise ExcelDatabaseError(
                f"The Excel Table {other.displayName} is below {table.displayName}, so rows can't be added."
            )

    for merged in worksheet.merged_cells.ranges:
        if merged.min_col <= max_col and merged.max_col >= min_col and merged.min_row <= end_row and merged.max_row >= start_row:
            raise ExcelDatabaseError(f"Merged cells {merged.coord} are where the new rows would go.")


def _get_calculated_column_formulas(table: Table, headers: list[str]) -> dict[str, str]:
    formulas: dict[str, str] = {}
    for header, column in zip(headers, table.tableColumns):
        formula = column.calculatedColumnFormula
        text = getattr(formula, "attr_text", formula)
        if text:
            formulas[header] = text if str(text).startswith("=") else f"={text}"
    return formulas


def _verify_saved_workbook(path: Path, worksheet_name: str, table_name: str, expected_ref: str) -> None:
    try:
        saved = load_workbook(path)
    except (InvalidFileException, zipfile.BadZipFile, KeyError, ValueError, OSError) as error:
        raise WorkbookSaveError(f"The saved workbook couldn't be read back: {error}") from error
    try:
        if worksheet_name not in saved.sheetnames or table_name not in saved[worksheet_name].tables:
            raise WorkbookSaveError("The saved workbook is missing the database table.")
        saved_ref = saved[worksheet_name].tables[table_name].ref
        if saved_ref != expected_ref:
            raise WorkbookSaveError(f"The saved table covers {saved_ref}, not {expected_ref}.")
    finally:
        saved.close()


def _get_file_signature(path: Path) -> tuple[int, int]:
    stat = path.stat()
    return stat.st_mtime_ns, stat.st_size


# endregion


# region HELPERS
def _fingerprint(values: Iterable[Any]) -> tuple[Any, ...]:
    return tuple(_fingerprint_part(value) for value in values)


def _fingerprint_part(value: Any) -> Any:
    if is_missing(value):
        return None
    if isinstance(value, datetime):
        # Excel stores times as fractions of a day, so compare to the nearest second.
        return (value + timedelta(microseconds=500_000)).replace(microsecond=0, tzinfo=None)
    number = to_nullable_float(value)
    if number is not None:
        return round(number, DUPLICATE_FLOAT_DECIMALS)
    return str(value).strip().casefold()


# endregion
