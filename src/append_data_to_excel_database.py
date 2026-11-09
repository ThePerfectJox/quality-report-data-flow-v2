"""Appends the quality report JSON to the Excel database table, then archives the batch.

Data flow:
1. run_client.py saves the server's JSON response. REPORT_JSON_PATH in .env points to it.
2. Each sample gets its calculated fields (average CW, BE and ME diameters, and the share
   of measurements in specification), plus the client metadata from client-config.json:
   Plant, Lokasi, Group, Brand, Set Sample and Bagian.
3. The rows are appended to the Excel Table in EXCEL_DATABASE_PATH, which is saved
   through a temporary file so a failed save can't damage it.
4. An archive workbook with the original, calculated and metadata fields is written to
   <folder of EXCEL_DATABASE_PATH>/archived-<Plant>/<ISO year>/<ISO week>/<yy_mm_dd_s_nnn>.xlsx,
   where nnn is a counter starting at 001 for each date and Set Sample.

Times are converted to Asia/Jakarta. Excel can't store a time zone, so Excel cells hold
Asia/Jakarta local times without one.

Set Sample boundary, with SET_SAMPLE_THRESHOLD "12:00":
- Submitted at or before 12:00:00.000000 local time: Set Sample 1.
- Submitted after it, even by a microsecond: Set Sample 2.

Usage: python src/append_data_to_excel_database.py
"""

import json
import logging
import math
import os
import re
import sys
import tempfile
import zipfile
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, tzinfo
from pathlib import Path
from typing import Any, Iterable
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, Cell
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from managers.environment_manager import (
    ENV_EXCEL_DATABASE_PATH,
    ENV_EXCEL_DATABASE_SHEET,
    ENV_EXCEL_DATABASE_TABLE,
    ENV_REPORT_JSON_PATH,
    EnvironmentManager,
    EnvironmentValueError,
)

LOGGER = logging.getLogger("append_data_to_excel_database")

# region CONFIGURATION
PROJECT_ROOT = Path(__file__).resolve().parents[1]
ENV_FILE = PROJECT_ROOT / ".env"
CLIENT_CONFIG_FILE = PROJECT_ROOT / "client-config.json"

LOCAL_TIMEZONE_NAME = "Asia/Jakarta"
SUPPORTED_WORKBOOK_SUFFIXES = {".xlsx"}

# The archive folder is "archived-" followed by the plant, for example "archived-RK1".
ARCHIVE_FOLDER_PREFIX = "archived-"
# Characters Windows doesn't allow in a folder name.
INVALID_FOLDER_NAME_CHARACTERS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
ARCHIVE_SHEET_NAME = "Data"
ARCHIVE_TABLE_NAME = "ArchivedData"
ARCHIVE_TABLE_STYLE = "TableStyleMedium2"
ARCHIVE_DATETIME_FORMAT = "yyyy-mm-dd hh:mm:ss"
# Each archive name ends with a 3-digit counter, _001 up to this number.
MAX_ARCHIVE_COUNTER = 999

# Keys in client-config.json.
CONFIG_PLANT = "PLANT"
CONFIG_LOKASI = "LOKASI"
CONFIG_GROUP = "GROUP"
CONFIG_BRAND = "BRAND"
CONFIG_BAGIAN = "BAGIAN"
CONFIG_SET_SAMPLE_THRESHOLD = "SET_SAMPLE_THRESHOLD"
CLIENT_CONFIG_KEYS = (
    CONFIG_PLANT,
    CONFIG_LOKASI,
    CONFIG_GROUP,
    CONFIG_BRAND,
    CONFIG_BAGIAN,
    CONFIG_SET_SAMPLE_THRESHOLD,
)

# Specification targets and tolerances. A value is in spec when
# target - tolerance <= value <= target + tolerance.
TARGET_CW = 2.02
TOLERANCE_CW = 0.14
TARGET_BE = 10.0
TOLERANCE_BE = 0.25
TARGET_ME = 8.0
TOLERANCE_ME = 0.25
# "Avg CW [1]" and "Avg CW [2]" are divided by this to give Avg Cw [First] and [Second].
CW_DIVISOR = 3

NULL_TEXT_VALUES = {"", "na", "n/a", "nan", "none", "null", "<na>", "nat", "-"}

# Source fields, as normalized by normalize_source_key(): "UB [1]" becomes "ub_1".
SOURCE_WORKER_ID_KEY = "id_ppskt"
SOURCE_START_KEY = "start"
CW_SOURCE_KEYS = ("avg_cw_1", "avg_cw_2")
BE_SOURCE_KEYS = tuple(f"ub_{number}" for number in range(1, 7))
ME_SOURCE_KEYS = tuple(f"uh_{number}" for number in range(1, 7))
# A sample without one of these keys means the JSON format changed, so it's rejected.
# Measurement values may still be empty; they're then left out of the averages.
REQUIRED_SOURCE_KEYS = (
    SOURCE_WORKER_ID_KEY,
    SOURCE_START_KEY,
    *CW_SOURCE_KEYS,
    *BE_SOURCE_KEYS,
    *ME_SOURCE_KEYS,
)

# Keys of the calculated and metadata fields. Original fields keep their normalized source key.
RESPONSE_ID_KEY = "tanggapan_id"
SUBMITTED_AT_KEY = "tanggal_pengiriman"
INSPECTED_AT_KEY = "tanggal_pemeriksaan"
WORKER_ID_KEY = "no_id_pekerja"
PLANT_KEY = "plant"
LOKASI_KEY = "lokasi"
GROUP_KEY = "group"
BRAND_KEY = "brand"
SET_SAMPLE_KEY = "set_sample"
BAGIAN_KEY = "bagian"

# Column names of the calculated and metadata fields in the archive workbook.
# Original fields keep their JSON name.
FIELD_LABELS = {
    RESPONSE_ID_KEY: "Tanggapan ID",
    SUBMITTED_AT_KEY: "Tanggal Pengiriman",
    INSPECTED_AT_KEY: "Tanggal Pemeriksaan",
    PLANT_KEY: "Plant",
    LOKASI_KEY: "Lokasi",
    GROUP_KEY: "Group",
    WORKER_ID_KEY: "No Id Pekerja",
    BRAND_KEY: "Brand",
    BAGIAN_KEY: "Bagian",
    SET_SAMPLE_KEY: "Set Sample",
    "avg_cw_first": "Avg Cw [First]",
    "avg_cw_second": "Avg Cw [Second]",
    "avg_cw": "Avg Cw",
    "avg_dia_be": "Avg Dia Be",
    "avg_dia_me": "Avg Dia Me",
    "in_spec_cw": "In Spec Cw",
    "in_spec_be": "In Spec Be",
    "in_spec_me": "In Spec Me",
}

# Source key -> column name in the database Excel Table. Only these fields are written.
# Columns are matched by name, ignoring case and repeated spaces, never by position.
EXCEL_COLUMN_MAPPING = {
    RESPONSE_ID_KEY: "ID tanggapan",
    SUBMITTED_AT_KEY: "Tanggal pengiriman",
    INSPECTED_AT_KEY: "Tanggal Pemeriksaan",
    PLANT_KEY: "Plant/Reg",
    LOKASI_KEY: "Lokasi",
    GROUP_KEY: "Group/Cell",
    WORKER_ID_KEY: "No ID Pekerja",
    BRAND_KEY: "Brand",
    SET_SAMPLE_KEY: "Set Sample",
    BAGIAN_KEY: "Bagian",
    "avg_cw": "Avg. CW",
    "avg_dia_be": "Avg. Dia. BE",
    "avg_dia_me": "Avg. Dia. ME",
    "in_spec_cw": "%In Spec CW",
    "in_spec_be": "%In Spec Dia. BE",
    "in_spec_me": "%In Spec Dia. ME",
}

# The append stops if the table lacks one of these columns. A mapped column that isn't
# required is skipped with a warning. Remove a key from this set to make its column optional.
REQUIRED_EXCEL_SOURCE_KEYS = frozenset(EXCEL_COLUMN_MAPPING)

# A sample counts as already in the database when all of these match an existing row.
# The JSON has no record ID, so these fields together identify one measurement.
DUPLICATE_CHECK_KEYS = (WORKER_ID_KEY, INSPECTED_AT_KEY, "avg_cw", "avg_dia_be", "avg_dia_me")
DUPLICATE_FLOAT_DECIMALS = 6
# endregion


# region ERRORS
class AppendDatabaseError(Exception):
    """The append can't go ahead. The message explains what to fix."""


class ConfigurationError(AppendDatabaseError):
    """.env or client-config.json is missing or contains an invalid value."""


class InvalidInputError(AppendDatabaseError):
    """The report JSON is missing, isn't valid JSON, or doesn't have the expected structure."""


class SubmissionTimestampError(InvalidInputError):
    """The report's submission timestamp is invalid."""


class ExcelDatabaseError(AppendDatabaseError):
    """The database workbook, worksheet or Excel Table can't be used."""


class WorkbookNotFoundError(ExcelDatabaseError):
    """The database workbook doesn't exist."""


class WorksheetNotFoundError(ExcelDatabaseError):
    """The configured worksheet isn't in the workbook."""


class TableNotFoundError(ExcelDatabaseError):
    """The Excel Table isn't in the workbook, or several tables match."""


class TableHeaderError(ExcelDatabaseError):
    """The table has blank or duplicate column headers."""


class MissingExcelColumnError(ExcelDatabaseError):
    """The table lacks a required mapped column."""


class WorkbookLockedError(ExcelDatabaseError):
    """The workbook is open in Excel or locked by another program."""


class WorkbookSaveError(ExcelDatabaseError):
    """The workbook couldn't be saved. The original file is unchanged."""


class DuplicateRecordError(AppendDatabaseError):
    """Samples in the report are already in the database."""


class ArchiveCollisionError(AppendDatabaseError):
    """No free archive filename is left for this date and Set Sample."""


class ArchiveCreationError(AppendDatabaseError):
    """The database append succeeded, but the archive workbook couldn't be created."""

    def __init__(self, message: str, database_path: Path, appended_row_count: int):
        super().__init__(message)
        self.database_path = database_path
        self.appended_row_count = appended_row_count


# endregion


# region DATA CLASSES
@dataclass(frozen=True)
class ExcelConfig:
    """Paths and table settings from .env."""

    workbook_path: Path
    worksheet_name: str | None
    table_name: str | None
    # Only needed when the report is read from a file, not passed in as a response.
    input_json_path: Path | None


@dataclass(frozen=True)
class ClientConfig:
    """The device metadata from client-config.json."""

    plant: str
    lokasi: str
    group: int | str
    brand: str
    bagian: str
    set_sample_threshold: time


@dataclass(frozen=True)
class PreparedRecords:
    """The validated report, ready to write. Records don't have a response ID yet."""

    records: list[dict[str, Any]]
    column_labels: dict[str, str]
    submission_datetime: datetime
    set_sample: int


@dataclass
class ExcelTableTarget:
    """A loaded database workbook and the Excel Table to append to."""

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


@dataclass(frozen=True)
class AppendResult:
    """What the append did."""

    database_path: Path
    worksheet_name: str | None
    table_name: str | None
    appended_row_count: int
    first_row: int | None
    last_row: int | None
    archive_path: Path | None
    set_sample: int
    submission_datetime: datetime


# endregion


# region CONFIGURATION LOADING
def get_local_timezone() -> ZoneInfo:
    """Returns the Asia/Jakarta time zone."""
    try:
        return ZoneInfo(LOCAL_TIMEZONE_NAME)
    except ZoneInfoNotFoundError as error:
        raise ConfigurationError(
            f"The time zone {LOCAL_TIMEZONE_NAME} isn't available. "
            "Install it with: python -m pip install -r requirements.txt"
        ) from error


def resolve_project_path(value: str | Path, project_root: Path = PROJECT_ROOT) -> Path:
    """Returns the path, with a relative path taken from the project folder."""
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = project_root / path
    return path.resolve()


def load_excel_configuration(
    env_file: Path = ENV_FILE,
    project_root: Path = PROJECT_ROOT,
) -> ExcelConfig:
    """Reads the Excel and input settings from .env and checks the workbook exists."""
    if not env_file.is_file():
        raise ConfigurationError(
            f"{env_file} was not found. Copy .env.example to .env and set "
            f"{ENV_EXCEL_DATABASE_PATH}."
        )

    environment = EnvironmentManager(env_file)
    try:
        workbook_value = environment.get_required(ENV_EXCEL_DATABASE_PATH)
    except EnvironmentValueError as error:
        raise ConfigurationError(str(error)) from error

    workbook_path = resolve_project_path(workbook_value, project_root)
    validate_workbook_path(workbook_path, ENV_EXCEL_DATABASE_PATH)

    input_json_value = _clean_optional_text(environment.get(ENV_REPORT_JSON_PATH))
    return ExcelConfig(
        workbook_path=workbook_path,
        worksheet_name=_clean_optional_text(environment.get(ENV_EXCEL_DATABASE_SHEET)),
        table_name=_clean_optional_text(environment.get(ENV_EXCEL_DATABASE_TABLE)),
        input_json_path=(
            resolve_project_path(input_json_value, project_root) if input_json_value else None
        ),
    )


def validate_workbook_path(workbook_path: Path, setting_name: str) -> None:
    """Raises WorkbookNotFoundError unless the path is an existing .xlsx file."""
    if not workbook_path.exists():
        raise WorkbookNotFoundError(
            f"The Excel database {workbook_path} doesn't exist. "
            f"Check {setting_name} in .env, or create the workbook first."
        )
    if not workbook_path.is_file():
        raise WorkbookNotFoundError(f"{setting_name} points to {workbook_path}, which isn't a file.")
    if workbook_path.suffix.lower() not in SUPPORTED_WORKBOOK_SUFFIXES:
        raise WorkbookNotFoundError(
            f"{workbook_path} isn't an .xlsx workbook. Save it as .xlsx in Excel first."
        )


def load_client_configuration(config_file: Path = CLIENT_CONFIG_FILE) -> ClientConfig:
    """Reads and checks the device metadata and the Set Sample threshold."""
    if not config_file.is_file():
        raise ConfigurationError(
            f"{config_file} was not found. Copy client-config.json.example to "
            f"{config_file.name} and fill in this device's values."
        )

    try:
        values = json.loads(config_file.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as error:
        raise ConfigurationError(f"{config_file} isn't valid JSON: {error}") from error
    except OSError as error:
        raise ConfigurationError(f"{config_file} can't be read: {error}") from error

    if not isinstance(values, dict):
        raise ConfigurationError(f"{config_file} must contain a JSON object.")

    missing_keys = [key for key in CLIENT_CONFIG_KEYS if key not in values]
    if missing_keys:
        raise ConfigurationError(f"{config_file} is missing: {', '.join(missing_keys)}")

    unknown_keys = sorted(set(values) - set(CLIENT_CONFIG_KEYS))
    if unknown_keys:
        LOGGER.warning("%s has unknown keys, which are ignored: %s", config_file, unknown_keys)

    def require_text(key: str) -> str:
        value = values[key]
        if not isinstance(value, str) or not value.strip():
            raise ConfigurationError(f"{key} in {config_file} must be non-empty text.")
        _ensure_excel_safe_text(value.strip(), f"{key} in {config_file}")
        return value.strip()

    group = values[CONFIG_GROUP]
    if isinstance(group, bool) or not isinstance(group, (int, str)) or not str(group).strip():
        raise ConfigurationError(f"{CONFIG_GROUP} in {config_file} must be a whole number or text.")

    try:
        threshold = parse_set_sample_threshold(values[CONFIG_SET_SAMPLE_THRESHOLD])
    except ValueError as error:
        raise ConfigurationError(f"{CONFIG_SET_SAMPLE_THRESHOLD} in {config_file}: {error}") from error

    return ClientConfig(
        plant=require_text(CONFIG_PLANT),
        lokasi=require_text(CONFIG_LOKASI),
        group=normalize_identifier(group),
        brand=require_text(CONFIG_BRAND),
        bagian=require_text(CONFIG_BAGIAN),
        set_sample_threshold=threshold,
    )


def parse_set_sample_threshold(value: Any) -> time:
    """Parses an HH:MM text, such as "12:00", into a time. Raises ValueError if invalid."""
    if not isinstance(value, str):
        raise ValueError(f"must be text in HH:MM format, such as \"12:00\", not {value!r}")
    match = re.fullmatch(r"(\d{2}):(\d{2})", value.strip())
    if match is None:
        raise ValueError(f"must use HH:MM format, such as \"12:00\", not {value!r}")
    hour, minute = int(match[1]), int(match[2])
    if hour > 23 or minute > 59:
        raise ValueError(f"{value!r} isn't a valid time of day (00:00 to 23:59)")
    return time(hour, minute)


def _clean_optional_text(value: str | None) -> str | None:
    if value is None:
        return None
    return value.strip() or None


# endregion


# region INPUT LOADING
def load_input_data(input_json_path: Path) -> dict[str, Any]:
    """Loads the report JSON saved by run_client.py and checks its structure."""
    if not input_json_path.is_file():
        raise InvalidInputError(
            f"The report JSON {input_json_path} was not found. Run src/run_client.py "
            f"first, or check {ENV_REPORT_JSON_PATH} in .env."
        )

    try:
        response = json.loads(input_json_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as error:
        raise InvalidInputError(f"{input_json_path} isn't valid JSON: {error}") from error
    except (OSError, UnicodeDecodeError) as error:
        raise InvalidInputError(f"{input_json_path} can't be read: {error}") from error

    validate_input_structure(response, input_json_path)
    return response


def validate_input_structure(response: Any, source: Path | str) -> None:
    """Raises InvalidInputError unless the response has the shape the server sends."""
    if not isinstance(response, dict):
        raise InvalidInputError(f"{source} must contain a JSON object.")

    status = response.get("status")
    if status != "ok":
        detail = response.get("message", "no message")
        raise InvalidInputError(f"{source} has status {status!r}, not 'ok' ({detail}).")

    report = response.get("report")
    if not isinstance(report, dict):
        raise InvalidInputError(f"{source} has no 'report' object.")

    samples = report.get("samples")
    if not isinstance(samples, list):
        raise InvalidInputError(f"{source} has no 'report.samples' list.")

    for index, sample in enumerate(samples, start=1):
        if not isinstance(sample, dict):
            raise InvalidInputError(f"Sample {index} in {source} isn't a JSON object.")

    sample_count = report.get("sample_count")
    if sample_count is not None and sample_count != len(samples):
        raise InvalidInputError(
            f"{source} says sample_count is {sample_count}, but it has {len(samples)} samples. "
            "The file may be incomplete."
        )


# endregion


# region SUBMISSION TIME AND SET SAMPLE
def parse_submission_datetime(value: Any, local_timezone: tzinfo) -> datetime:
    """Parses an ISO 8601 timestamp and returns it in the local time zone.

    A timestamp without a time zone is taken to be local (Asia/Jakarta) time.
    """
    if not isinstance(value, str) or not value.strip():
        raise SubmissionTimestampError(
            f"The submission timestamp 'responded_at' must be ISO 8601 text, not {value!r}."
        )
    try:
        parsed = datetime.fromisoformat(value.strip())
    except ValueError as error:
        raise SubmissionTimestampError(
            f"The submission timestamp 'responded_at' isn't a valid ISO 8601 time: {value!r}"
        ) from error

    if parsed.tzinfo is None:
        LOGGER.warning(
            "The submission timestamp %s has no time zone, so it's taken as %s time.",
            value,
            LOCAL_TIMEZONE_NAME,
        )
        return parsed.replace(tzinfo=local_timezone)
    return parsed.astimezone(local_timezone)


def resolve_submission_datetime(
    response: dict[str, Any],
    local_timezone: tzinfo,
    current_time: datetime | None = None,
) -> datetime:
    """Returns the report's submission time in the local time zone.

    Uses 'responded_at' from the response. Only when it's absent or null, uses the
    current time instead. A present but invalid value raises SubmissionTimestampError.
    """
    value = response.get("responded_at")
    if value is not None:
        return parse_submission_datetime(value, local_timezone)

    if current_time is None:
        current_time = datetime.now(local_timezone)
    elif current_time.tzinfo is None:
        raise SubmissionTimestampError("current_time must include a time zone.")

    LOGGER.warning(
        "The report has no 'responded_at' timestamp. Using the current time instead: %s",
        current_time.astimezone(local_timezone).isoformat(),
    )
    return current_time.astimezone(local_timezone)


def calculate_set_sample(
    submission_datetime: datetime,
    threshold: time,
    local_timezone: tzinfo | None = None,
) -> int:
    """Returns 2 if the submission is after the threshold in local time, otherwise 1.

    The comparison is submission_time > threshold, so a submission exactly at the
    threshold (for 12:00, at 12:00:00.000000) is Sample 1.
    """
    if submission_datetime.tzinfo is None:
        raise SubmissionTimestampError(
            "The submission time must include a time zone to calculate Set Sample."
        )
    local_timezone = local_timezone or get_local_timezone()
    local_time = submission_datetime.astimezone(local_timezone).time()
    return 2 if local_time > threshold else 1


# endregion


# region CALCULATION
def is_missing(value: Any) -> bool:
    """Returns True for None, a recognized empty text, or a number that isn't finite."""
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip().lower() in NULL_TEXT_VALUES
    if isinstance(value, float):
        return not math.isfinite(value)
    return False


def to_nullable_float(value: Any) -> float | None:
    """Returns the value as a finite float, or None if it's missing or not a number."""
    if is_missing(value) or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
    elif isinstance(value, str):
        try:
            number = float(value.strip().replace(",", "."))
        except ValueError:
            return None
    else:
        return None
    return number if math.isfinite(number) else None


def safe_divide(value: Any, divisor: Any) -> float | None:
    """Divides two values, returning None when either is missing or the divisor is 0."""
    numeric_value = to_nullable_float(value)
    numeric_divisor = to_nullable_float(divisor)
    if numeric_value is None or numeric_divisor is None or numeric_divisor == 0:
        return None
    result = numeric_value / numeric_divisor
    return result if math.isfinite(result) else None


def safe_mean(values: Iterable[Any]) -> float | None:
    """Returns the average of the values that aren't missing, or None if all are."""
    valid_values = [number for number in map(to_nullable_float, values) if number is not None]
    if not valid_values:
        return None
    average = sum(valid_values) / len(valid_values)
    return average if math.isfinite(average) else None


def check_in_spec(value: Any, target: float, tolerance: float) -> bool | None:
    """Returns whether target - tolerance <= value <= target + tolerance, or None if missing."""
    if tolerance < 0:
        raise ValueError("Tolerance cannot be negative.")
    numeric_value = to_nullable_float(value)
    if numeric_value is None:
        return None
    return target - tolerance <= numeric_value <= target + tolerance


def calculate_in_spec_percentage(
    values: Iterable[Any],
    target: float,
    tolerance: float,
) -> float | None:
    """Returns the fraction (0 to 1) of non-missing values in specification."""
    if tolerance < 0:
        raise ValueError("Tolerance cannot be negative.")
    valid_values = [number for number in map(to_nullable_float, values) if number is not None]
    if not valid_values:
        return None
    in_spec_count = sum(check_in_spec(number, target, tolerance) is True for number in valid_values)
    return in_spec_count / len(valid_values)


def calculate_derived_data(
    input_data: dict[str, Any],
    local_timezone: tzinfo | None = None,
) -> dict[str, Any]:
    """Calculates the report fields of one normalized sample.

    - Avg Cw [First] and [Second]: "Avg CW [1]" and "Avg CW [2]" divided by 3.
    - Avg Cw: the average of those two values.
    - Avg Dia Be and Avg Dia Me: the averages of UB [1-6] and UH [1-6].
    - In Spec Cw, Be and Me: the fraction (0 to 1) of those values within specification.
    Missing values are left out of averages and fractions. If all are missing, the result
    is None.
    """
    local_timezone = local_timezone or get_local_timezone()
    sample_name = f"sample of worker {input_data.get(SOURCE_WORKER_ID_KEY)!r}"

    def numeric(key: str) -> float | None:
        value = input_data.get(key)
        number = to_nullable_float(value)
        if number is None and not is_missing(value):
            LOGGER.warning("%s: %s is %r, which isn't a number. It's left out.", sample_name, key, value)
        return number

    avg_cw_first = safe_divide(numeric(CW_SOURCE_KEYS[0]), CW_DIVISOR)
    avg_cw_second = safe_divide(numeric(CW_SOURCE_KEYS[1]), CW_DIVISOR)
    be_values = [numeric(key) for key in BE_SOURCE_KEYS]
    me_values = [numeric(key) for key in ME_SOURCE_KEYS]

    return {
        INSPECTED_AT_KEY: parse_inspection_datetime(
            input_data.get(SOURCE_START_KEY), local_timezone, sample_name
        ),
        WORKER_ID_KEY: normalize_identifier(input_data.get(SOURCE_WORKER_ID_KEY), sample_name),
        "avg_cw_first": avg_cw_first,
        "avg_cw_second": avg_cw_second,
        "avg_cw": safe_mean([avg_cw_first, avg_cw_second]),
        "avg_dia_be": safe_mean(be_values),
        "avg_dia_me": safe_mean(me_values),
        "in_spec_cw": calculate_in_spec_percentage(
            [avg_cw_first, avg_cw_second], TARGET_CW, TOLERANCE_CW
        ),
        "in_spec_be": calculate_in_spec_percentage(be_values, TARGET_BE, TOLERANCE_BE),
        "in_spec_me": calculate_in_spec_percentage(me_values, TARGET_ME, TOLERANCE_ME),
    }


def build_metadata(
    client_config: ClientConfig,
    set_sample: int,
    submission_datetime: datetime,
) -> dict[str, Any]:
    """Returns the fields that are the same for every sample in the report."""
    return {
        SUBMITTED_AT_KEY: submission_datetime.replace(tzinfo=None),
        PLANT_KEY: client_config.plant,
        LOKASI_KEY: client_config.lokasi,
        GROUP_KEY: client_config.group,
        BRAND_KEY: client_config.brand,
        BAGIAN_KEY: client_config.bagian,
        SET_SAMPLE_KEY: set_sample,
    }


def combine_original_and_calculated_data(
    original_data: dict[str, Any],
    calculated_data: dict[str, Any],
    metadata: dict[str, Any],
) -> dict[str, Any]:
    """Returns metadata, then calculated fields, then the original fields.

    An original field with the same key as a calculated or metadata field is replaced by it. Use find_replaced_original_keys() to report those.
    """
    overlap = metadata.keys() & calculated_data.keys()
    if overlap:
        raise ValueError(f"Metadata and calculated fields share keys: {sorted(overlap)}")
    generated = {**metadata, **calculated_data}
    originals = {key: value for key, value in original_data.items() if key not in generated}
    return {**generated, **originals}


def find_replaced_original_keys(
    original_data: dict[str, Any],
    calculated_data: dict[str, Any],
    metadata: dict[str, Any],
) -> set[str]:
    """Returns the original keys that combine_original_and_calculated_data() replaces."""
    return original_data.keys() & (calculated_data.keys() | metadata.keys())


# endregion


# region NORMALIZATION
def normalize_source_key(name: str) -> str:
    """Returns the canonical key of a source field: "Avg CW [1]" and "avg_cw_1" both become
    "avg_cw_1". This is the only place source field names are converted."""
    return re.sub(r"[^0-9a-z]+", "_", name.strip().casefold()).strip("_")


def normalize_header_name(name: Any) -> str:
    """Returns an Excel header for comparing: case and repeated spaces are ignored."""
    return " ".join(str(name).split()).casefold()


def normalize_identifier(value: Any, context: str = "value") -> int | str:
    """Returns a whole number for digit-only IDs such as "317057", otherwise trimmed text.

    Existing database rows store worker IDs and groups as numbers. IDs with a leading
    zero stay text, so the zero isn't lost.
    """
    if isinstance(value, bool) or is_missing(value):
        raise InvalidInputError(f"{context}: the ID is missing.")
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    text = str(value).strip()
    return int(text) if re.fullmatch(r"[1-9]\d*", text) else text


def parse_inspection_datetime(value: Any, local_timezone: tzinfo, context: str) -> datetime:
    """Parses the sample's Start time into a local datetime without a time zone, for Excel."""
    if not isinstance(value, str) or is_missing(value):
        raise InvalidInputError(f"{context}: Start must be an ISO 8601 time, not {value!r}.")
    try:
        parsed = datetime.fromisoformat(value.strip())
    except ValueError as error:
        raise InvalidInputError(f"{context}: Start isn't a valid ISO 8601 time: {value!r}") from error
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(local_timezone).replace(tzinfo=None)
    return parsed


def normalize_source_record(
    raw_sample: dict[str, Any],
    sample_number: int,
) -> tuple[dict[str, Any], dict[str, str]]:
    """Returns the sample with canonical keys, and each canonical key's original name."""
    values: dict[str, Any] = {}
    labels: dict[str, str] = {}
    for raw_key, value in raw_sample.items():
        key = normalize_source_key(raw_key)
        if not key:
            raise InvalidInputError(f"Sample {sample_number} has a field with no usable name: {raw_key!r}")
        if key in values:
            raise InvalidInputError(
                f"Sample {sample_number} has fields {labels[key]!r} and {raw_key!r}, "
                "which mean the same field."
            )
        if isinstance(value, str):
            _ensure_excel_safe_text(value, f"Sample {sample_number}, field {raw_key!r}")
        _ensure_excel_safe_text(raw_key, f"Sample {sample_number}, field name")
        values[key] = value
        labels[key] = raw_key.strip()

    missing = [key for key in REQUIRED_SOURCE_KEYS if key not in values]
    if missing:
        raise InvalidInputError(
            f"Sample {sample_number} is missing these fields: {', '.join(missing)}. "
            "The report format may have changed."
        )
    return values, labels


def prepare_records(
    response: dict[str, Any],
    client_config: ClientConfig,
    local_timezone: tzinfo,
    current_time: datetime | None = None,
) -> PreparedRecords:
    """Validates every sample and builds the combined records, before any file is touched."""
    submission_datetime = resolve_submission_datetime(response, local_timezone, current_time)
    set_sample = calculate_set_sample(
        submission_datetime, client_config.set_sample_threshold, local_timezone
    )
    metadata = build_metadata(client_config, set_sample, submission_datetime)

    records: list[dict[str, Any]] = []
    column_labels: dict[str, str] = {RESPONSE_ID_KEY: FIELD_LABELS[RESPONSE_ID_KEY]}
    replaced_keys: set[str] = set()

    for sample_number, raw_sample in enumerate(response["report"]["samples"], start=1):
        original, labels = normalize_source_record(raw_sample, sample_number)
        calculated = calculate_derived_data(original, local_timezone)
        replaced_keys |= find_replaced_original_keys(original, calculated, metadata)
        record = combine_original_and_calculated_data(original, calculated, metadata)
        records.append(record)
        for key in record:
            column_labels.setdefault(key, FIELD_LABELS.get(key) or labels[key])

    if replaced_keys:
        LOGGER.warning(
            "These source fields have the same name as a calculated or metadata field. "
            "The calculated value is used: %s",
            sorted(replaced_keys),
        )

    return PreparedRecords(records, column_labels, submission_datetime, set_sample)


def _ensure_excel_safe_text(text: str, context: str) -> None:
    if ILLEGAL_CHARACTERS_RE.search(text):
        raise InvalidInputError(f"{context} contains control characters that Excel can't store.")


# endregion


# region COLUMN MAPPING
def resolve_excel_columns(
    excel_headers: list[str],
    mapping: dict[str, str] = EXCEL_COLUMN_MAPPING,
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
    mapping: dict[str, str] = EXCEL_COLUMN_MAPPING,
    required_keys: Iterable[str] = REQUIRED_EXCEL_SOURCE_KEYS,
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
    mapping: dict[str, str] = EXCEL_COLUMN_MAPPING,
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
                f"Worksheets: {workbook.sheetnames}. Check {ENV_EXCEL_DATABASE_SHEET} in .env."
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
                f"Check {ENV_EXCEL_DATABASE_TABLE} and {ENV_EXCEL_DATABASE_SHEET} in .env."
            )
        return matches[0]

    if not tables:
        raise TableNotFoundError(
            "No Excel Table was found. The data must be inside an Excel Table "
            "(select the data, then Insert > Table)."
        )
    if len(tables) > 1:
        raise TableNotFoundError(
            f"The workbook has several Excel Tables: {available}. Set "
            f"{ENV_EXCEL_DATABASE_SHEET} and {ENV_EXCEL_DATABASE_TABLE} in .env."
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


def assign_response_ids(
    records: list[dict[str, Any]],
    existing_ids: Iterable[Any],
    submission_date: date,
) -> list[dict[str, Any]]:
    """Gives each record an ID in the format "yymmdd-001", using the submission date.

    Numbering continues after the highest existing ID of the same date, so a second
    report on the same day doesn't reuse IDs.
    """
    prefix = submission_date.strftime("%y%m%d")
    pattern = re.compile(rf"{re.escape(prefix)}-(\d+)")
    existing = {str(value).strip() for value in existing_ids if value is not None}
    highest = max(
        (int(match[1]) for value in existing if (match := pattern.fullmatch(value))),
        default=0,
    )

    numbered: list[dict[str, Any]] = []
    for offset, record in enumerate(records, start=1):
        response_id = f"{prefix}-{highest + offset:03d}"
        if response_id in existing:
            raise DuplicateRecordError(f"The response ID {response_id} is already in the table.")
        numbered.append({RESPONSE_ID_KEY: response_id, **record})
    return numbered


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
                _write_cell_value(cell, row[header])
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


# region ARCHIVE
def build_archive_root(workbook_folder: Path, plant: str) -> Path:
    """Returns <workbook folder>/archived-<plant>, for example archived-RK1."""
    if not plant.strip() or INVALID_FOLDER_NAME_CHARACTERS.search(plant) or plant.strip(" .") != plant:
        raise ConfigurationError(
            f"PLANT {plant!r} can't be used in the archive folder name. It must not contain "
            '< > : " / \\ | ? * or start or end with a space or dot.'
        )
    return workbook_folder / f"{ARCHIVE_FOLDER_PREFIX}{plant}"


def build_archive_base_path(archive_root: Path, submission_datetime: datetime, set_sample: int) -> Path:
    """Returns <archive root>/<ISO year>/<ISO week, 2 digits>/<yy_mm_dd_s>.xlsx for the local date.

    This is the name without its counter. resolve_archive_path() adds the counter.
    """
    iso_year, iso_week, _ = submission_datetime.isocalendar()
    return (
        archive_root
        / f"{iso_year:04d}"
        / f"{iso_week:02d}"
        / f"{submission_datetime:%y_%m_%d}_{set_sample}.xlsx"
    )


def resolve_archive_path(base_path: Path) -> Path:
    """Adds the next counter to base_path: 26_11_06_1_001.xlsx, then _002, _003, ...

    The counter is one higher than the highest existing one, so a number is never reused,
    even if an earlier archive was deleted.
    """
    pattern = re.compile(rf"{re.escape(base_path.stem)}_(\d{{3}}){re.escape(base_path.suffix)}", re.IGNORECASE)
    existing = (
        [int(match[1]) for file in base_path.parent.iterdir() if (match := pattern.fullmatch(file.name))]
        if base_path.parent.is_dir()
        else []
    )
    counter = max(existing, default=0) + 1
    if counter > MAX_ARCHIVE_COUNTER:
        raise ArchiveCollisionError(
            f"{base_path.parent} already has archives up to {base_path.stem}_{MAX_ARCHIVE_COUNTER:03d}"
            f"{base_path.suffix}. No counter is left for this date and Set Sample."
        )
    return base_path.with_name(f"{base_path.stem}_{counter:03d}{base_path.suffix}")


def create_archive_workbook(
    records: list[dict[str, Any]],
    column_labels: dict[str, str],
    archive_path: Path,
) -> None:
    """Writes the records into a new workbook, inside the Excel Table ArchivedData.

    Never overwrites: raises ArchiveCollisionError if archive_path already exists.
    """
    if not records:
        raise ValueError("There are no records to archive.")

    keys = list(dict.fromkeys(key for record in records for key in record))
    headers = [column_labels.get(key, key) for key in keys]
    normalized = [normalize_header_name(header) for header in headers]
    if any(not header.strip() for header in headers) or len(set(normalized)) != len(normalized):
        raise TableHeaderError(f"The archive column names aren't unique and non-blank: {headers}")

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = ARCHIVE_SHEET_NAME
    for column, header in enumerate(headers, start=1):
        _write_cell_value(worksheet.cell(1, column), header)
    for row_number, record in enumerate(records, start=2):
        for column, key in enumerate(keys, start=1):
            cell = worksheet.cell(row_number, column)
            value = _to_archive_value(record.get(key))
            _write_cell_value(cell, value)
            if isinstance(value, datetime):
                cell.number_format = ARCHIVE_DATETIME_FORMAT

    table = Table(
        displayName=ARCHIVE_TABLE_NAME,
        ref=f"A1:{get_column_letter(len(keys))}{len(records) + 1}",
    )
    table.tableStyleInfo = TableStyleInfo(name=ARCHIVE_TABLE_STYLE, showRowStripes=True)
    worksheet.add_table(table)
    worksheet.freeze_panes = "A2"

    archive_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        # "x" creates the file only if it doesn't exist, so an archive is never overwritten.
        with archive_path.open("xb") as file:
            workbook.save(file)
    except FileExistsError as error:
        raise ArchiveCollisionError(f"{archive_path} was created by another process just now.") from error
    except BaseException:
        if archive_path.exists():
            archive_path.unlink()
        raise


def _to_archive_value(value: Any) -> Any:
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    return value


# endregion


# region SHARED HELPERS
def _write_cell_value(cell: Cell, value: Any) -> None:
    cell.value = value
    # Text starting with = must stay text, or Excel would run it as a formula.
    if isinstance(value, str) and cell.data_type == "f":
        cell.data_type = "s"


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


# region MAIN
def append_data_to_excel_database(
    input_json_path: Path | str | None = None,
    *,
    env_file: Path = ENV_FILE,
    client_config_file: Path = CLIENT_CONFIG_FILE,
    project_root: Path = PROJECT_ROOT,
    current_time: datetime | None = None,
) -> AppendResult:
    """Appends a report JSON file to the Excel database table, then writes the archive.

    Reads input_json_path, or REPORT_JSON_PATH from .env when it's not given. Raises an
    AppendDatabaseError subclass on failure. ArchiveCreationError means the database
    append did succeed.
    """
    excel_config = load_excel_configuration(env_file, project_root)
    client_config = load_client_configuration(client_config_file)
    if input_json_path is not None:
        json_path = resolve_project_path(input_json_path, project_root)
    elif excel_config.input_json_path is not None:
        json_path = excel_config.input_json_path
    else:
        raise ConfigurationError(
            f"No report JSON was given. Set {ENV_REPORT_JSON_PATH} in {env_file}, "
            "or pass the file path."
        )

    LOGGER.info("Reading %s", json_path)
    response = load_input_data(json_path)
    return _append_validated_response(response, excel_config, client_config, current_time)


def append_response_to_database(
    response: dict[str, Any],
    *,
    env_file: Path = ENV_FILE,
    client_config_file: Path = CLIENT_CONFIG_FILE,
    project_root: Path = PROJECT_ROOT,
    current_time: datetime | None = None,
) -> AppendResult:
    """Appends a server response, as returned by send_request(), to the Excel database table.

    Works like append_data_to_excel_database(), but takes the response directly instead
    of reading a JSON file, so REPORT_JSON_PATH isn't needed. The response must have the
    same structure as the saved JSON: "status", "responded_at" and "report.samples".
    """
    excel_config = load_excel_configuration(env_file, project_root)
    client_config = load_client_configuration(client_config_file)
    validate_input_structure(response, "The server response")
    return _append_validated_response(response, excel_config, client_config, current_time)


def _append_validated_response(
    response: dict[str, Any],
    excel_config: ExcelConfig,
    client_config: ClientConfig,
    current_time: datetime | None,
) -> AppendResult:
    """Appends a validated response to the database table, then writes the archive.

    Order: prepare and validate every record, check the workbook and table, append and
    save the database atomically, then create the archive.
    """
    local_timezone = get_local_timezone()
    prepared = prepare_records(response, client_config, local_timezone, current_time)
    # Checked now, so an unusable plant name stops the run before the database changes.
    archive_root = build_archive_root(excel_config.workbook_path.parent, client_config.plant)
    LOGGER.info(
        "Submitted at %s, Set Sample %s, %d samples",
        prepared.submission_datetime.isoformat(),
        prepared.set_sample,
        len(prepared.records),
    )

    if not prepared.records:
        LOGGER.info("The report has no samples. Nothing was appended or archived.")
        return AppendResult(
            excel_config.workbook_path, excel_config.worksheet_name, excel_config.table_name,
            0, None, None, None, prepared.set_sample, prepared.submission_datetime,
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

    records = assign_response_ids(
        prepared.records,
        (row.get(resolved_columns[RESPONSE_ID_KEY]) for row in existing_rows),
        prepared.submission_datetime.date(),
    )
    excel_rows = [map_data_to_excel_columns(record, target.headers) for record in records]
    first_row, last_row = append_rows_to_excel_table(target, excel_rows)
    save_workbook_atomically(target)
    LOGGER.info(
        "Appended %d rows (worksheet rows %d-%d) to %s!%s in %s",
        len(records), first_row, last_row, target.worksheet.title,
        target.table.displayName, target.workbook_path,
    )

    try:
        archive_path = resolve_archive_path(
            build_archive_base_path(
                archive_root,
                prepared.submission_datetime,
                prepared.set_sample,
            )
        )
        create_archive_workbook(records, prepared.column_labels, archive_path)
    except (AppendDatabaseError, OSError, ValueError) as error:
        raise ArchiveCreationError(
            f"The database append succeeded ({len(records)} rows added to {target.workbook_path}), "
            f"but the archive couldn't be created: {error}. Don't run the append again; "
            "the rows are already in the database.",
            target.workbook_path,
            len(records),
        ) from error

    LOGGER.info("Archive saved to %s", archive_path)
    return AppendResult(
        target.workbook_path, target.worksheet.title, target.table.displayName,
        len(records), first_row, last_row, archive_path,
        prepared.set_sample, prepared.submission_datetime,
    )


def _configure_logging() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


if __name__ == "__main__":
    _configure_logging()
    try:
        append_data_to_excel_database("c:/quality-report-data-flow-v2/test-response.json")
    except ArchiveCreationError as error:
        LOGGER.error("%s", error)
        sys.exit(2)
    except AppendDatabaseError as error:
        LOGGER.error("Nothing was written to the database: %s", error)
        sys.exit(1)
# endregion
