"""Turns the server's response (rows of the summary_report view) into master data records.

For each sample:
- The view's fields are given canonical keys: "UB [1]" becomes "ub_1".
- Group is the view's "ID Group / Cell". Brand is the part of "Nama Group / Cell" before the
  first "-": "SAI-2" gives "SAI".
- The averages are calculated here, and the in-spec fractions in client/inspec.py, using the
  sample's brand entry in INSPEC_REQUIREMENTS of client-config.json.
- Set Sample is the view's "Set Sample". A sample without one is rejected.
- Plant, Lokasi and Bagian come from "server_config" in the response, which the server reads
  from its server-config.json.

Times are converted to Asia/Jakarta. Excel can't store a time zone, so records hold
Asia/Jakarta local times without one.
"""

import json
import logging
from dataclasses import dataclass
from datetime import datetime, tzinfo
from pathlib import Path
from typing import Any

from client.configuration import CLIENT_CONFIG_LABEL, CONFIG_INSPEC_REQUIREMENTS, ClientConfig
from client.errors import ConfigurationError, InvalidInputError, SubmissionTimestampError
from client.inspec import (
    InspecRequirement,
    InspecRequirementError,
    calculate_inspec,
    find_requirement_for_brand,
)
from client.master_data_mapping import (
    ARCHIVE_FIELD_LABELS,
    BAGIAN_KEY,
    BE_SOURCE_KEYS,
    BRAND_KEY,
    BRAND_SEPARATOR,
    CW_SOURCE_KEYS,
    DEFECT_SOURCE_KEYS,
    GROUP_KEY,
    INSPECTED_AT_KEY,
    LOKASI_KEY,
    ME_SOURCE_KEYS,
    PLANT_KEY,
    REQUIRED_SOURCE_KEYS,
    RESPONSE_ID_KEY,
    SET_SAMPLE_KEY,
    SOURCE_GROUP_ID_KEY,
    SOURCE_GROUP_NAME_KEY,
    SOURCE_SET_SAMPLE_KEY,
    SOURCE_START_KEY,
    SOURCE_WORKER_ID_KEY,
    SUBMITTED_AT_KEY,
    WORKER_ID_KEY,
)
from client.response_ids import assign_response_ids
from client.values import (
    LOCAL_TIMEZONE_NAME,
    ensure_excel_safe_text,
    get_local_timezone,
    is_missing,
    normalize_identifier,
    normalize_source_key,
    safe_divide,
    safe_mean,
    to_nullable_float,
    to_number_or_text,
)
from managers.environment_manager import ENV_REPORT_JSON_PATH

LOGGER = logging.getLogger(__name__)

# "Avg CW [1]" and "Avg CW [2]" are divided by this to give Avg Cw [First] and [Second].
CW_DIVISOR = 3

# View fields that the calculation replaces on purpose with a cleaned value of the same
# field, so they aren't reported as overwritten.
CLEANED_SOURCE_KEYS = frozenset({SOURCE_SET_SAMPLE_KEY, *DEFECT_SOURCE_KEYS})


# "server_config" in the response: the server device's server-config.json.
RESPONSE_SERVER_CONFIG_KEY = "server_config"
SERVER_CONFIG_PLANT = "PLANT"
SERVER_CONFIG_LOKASI = "LOKASI"
SERVER_CONFIG_BAGIAN = "BAGIAN"
SERVER_CONFIG_KEYS = (SERVER_CONFIG_PLANT, SERVER_CONFIG_LOKASI, SERVER_CONFIG_BAGIAN)


@dataclass(frozen=True)
class ServerConfig:
    """Where the server device is, from "server_config" in the response."""

    plant: str
    lokasi: str
    bagian: str


@dataclass(frozen=True)
class PreparedRecords:
    """The validated report, ready to write. Each record starts with its response ID."""

    records: list[dict[str, Any]]
    column_labels: dict[str, str]
    submission_datetime: datetime
    # The server's PLANT. Names the archive folder.
    plant: str
    # The brands of the samples, in the order they first appear.
    brands: list[str]


# region INPUT
def load_input_data(input_json_path: Path) -> dict[str, Any]:
    """Loads a saved report JSON and checks its structure."""
    if not input_json_path.is_file():
        raise InvalidInputError(
            f"The report JSON {input_json_path} was not found. Check {ENV_REPORT_JSON_PATH} "
            "in .env, or pass the file path."
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

    read_server_config(response, source)

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


def read_server_config(response: dict[str, Any], source: Path | str = "The server response") -> ServerConfig:
    """Returns PLANT, LOKASI and BAGIAN from the response's "server_config"."""
    values = response.get(RESPONSE_SERVER_CONFIG_KEY)
    if not isinstance(values, dict):
        raise InvalidInputError(
            f"{source} has no '{RESPONSE_SERVER_CONFIG_KEY}' object. Update the server, so it "
            "sends PLANT, LOKASI and BAGIAN from its server-config.json."
        )

    def require_text(key: str) -> str:
        value = values.get(key)
        if not isinstance(value, str) or not value.strip():
            raise InvalidInputError(
                f"{key} in '{RESPONSE_SERVER_CONFIG_KEY}' of {source} must be non-empty text. "
                "Check server-config.json on the server."
            )
        ensure_excel_safe_text(value, f"{key} in '{RESPONSE_SERVER_CONFIG_KEY}' of {source}")
        return value.strip()

    return ServerConfig(
        plant=require_text(SERVER_CONFIG_PLANT),
        lokasi=require_text(SERVER_CONFIG_LOKASI),
        bagian=require_text(SERVER_CONFIG_BAGIAN),
    )


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
            ensure_excel_safe_text(value, f"Sample {sample_number}, field {raw_key!r}")
        ensure_excel_safe_text(raw_key, f"Sample {sample_number}, field name")
        values[key] = value
        labels[key] = raw_key.strip()

    missing = [key for key in REQUIRED_SOURCE_KEYS if key not in values]
    if missing:
        raise InvalidInputError(
            f"Sample {sample_number} is missing these fields: {', '.join(missing)}. "
            "The summary_report view may have changed; check client/master_data_mapping.py."
        )
    return values, labels


# endregion


# region SUBMISSION TIME AND SET SAMPLE
# The submission time fills Tanggal pengiriman, the response IDs and the archive date.
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


def parse_set_sample(value: Any, context: str) -> int:
    """Returns the view's Set Sample as a whole number. Raises InvalidInputError if it's
    empty or isn't a whole number from 1 up."""
    if is_missing(value):
        raise InvalidInputError(
            f"{context} has no Set Sample. Check ProcessH_Set of this sample in the database."
        )
    number = to_nullable_float(value)
    if number is None or not number.is_integer() or number < 1:
        raise InvalidInputError(f"{context}: Set Sample must be a whole number such as 1 or 2, not {value!r}.")
    return int(number)


# endregion


# region CALCULATION
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


def derive_brand(group_name: Any, context: str) -> str:
    """Returns the part of the group name before the first "-": "SAI-2" gives "SAI".

    A group name without "-" is the brand as a whole.
    """
    if not isinstance(group_name, str) or is_missing(group_name):
        raise InvalidInputError(
            f"{context}: Nama Group / Cell must be a group name such as \"SAI-2\", not {group_name!r}."
        )
    brand = group_name.split(BRAND_SEPARATOR, 1)[0].strip()
    if not brand:
        raise InvalidInputError(
            f"{context}: the group name {group_name!r} has no brand before {BRAND_SEPARATOR!r}."
        )
    return brand


def find_sample_requirement(
    inspec_requirements: dict[str, InspecRequirement],
    brand: str,
    group_name: str,
    context: str,
) -> InspecRequirement:
    """Returns the inspec requirement of the sample's brand."""
    try:
        return find_requirement_for_brand(inspec_requirements, brand)
    except InspecRequirementError as error:
        raise ConfigurationError(
            f"{context}: the brand {brand!r}, from the group {group_name!r}, has no entry in "
            f"{CONFIG_INSPEC_REQUIREMENTS} of {CLIENT_CONFIG_LABEL}. Brands: "
            f"{', '.join(inspec_requirements)}. Add it, then run the append again."
        ) from error


def calculate_derived_data(
    input_data: dict[str, Any],
    inspec_requirements: dict[str, InspecRequirement],
    local_timezone: tzinfo | None = None,
) -> dict[str, Any]:
    """Calculates the master data fields of one normalized sample.

    - Group: the view's "ID Group / Cell".
    - Brand: the part of the view's "Nama Group / Cell" before the first "-".
    - Avg Cw [First] and [Second]: "Avg CW [1]" and "Avg CW [2]" divided by 3.
    - Avg Cw: the average of those two values.
    - Avg Dia Be and Avg Dia Me: the averages of UB [1-6] and UH [1-6].
    - In Spec Cw, Be and Me: from client/inspec.py, with the brand's requirement.
    - Set Sample: the view's "Set Sample".
    - Defect counts: numbers stay numbers, so "2.00" from MySQL becomes 2.
    Missing values are left out of averages and fractions. If all are missing, the result
    is None.
    """
    local_timezone = local_timezone or get_local_timezone()
    sample_name = f"sample of worker {input_data.get(SOURCE_WORKER_ID_KEY)!r}"
    group_name = input_data.get(SOURCE_GROUP_NAME_KEY)
    brand = derive_brand(group_name, sample_name)
    inspec = find_sample_requirement(inspec_requirements, brand, group_name, sample_name)

    def numeric(key: str) -> float | None:
        value = input_data.get(key)
        number = to_nullable_float(value)
        if number is None and not is_missing(value):
            LOGGER.warning("%s: %s is %r, which isn't a number. It's left out.", sample_name, key, value)
        return number

    avg_cw_first = safe_divide(numeric(CW_SOURCE_KEYS[0]), CW_DIVISOR)
    avg_cw_second = safe_divide(numeric(CW_SOURCE_KEYS[1]), CW_DIVISOR)
    cw_values = [avg_cw_first, avg_cw_second]
    be_values = [numeric(key) for key in BE_SOURCE_KEYS]
    me_values = [numeric(key) for key in ME_SOURCE_KEYS]

    return {
        INSPECTED_AT_KEY: parse_inspection_datetime(
            input_data.get(SOURCE_START_KEY), local_timezone, sample_name
        ),
        WORKER_ID_KEY: normalize_identifier(input_data.get(SOURCE_WORKER_ID_KEY), sample_name),
        GROUP_KEY: normalize_identifier(input_data.get(SOURCE_GROUP_ID_KEY), f"{sample_name}, ID Group / Cell"),
        BRAND_KEY: brand,
        SET_SAMPLE_KEY: parse_set_sample(input_data.get(SOURCE_SET_SAMPLE_KEY), sample_name),
        "avg_cw_first": avg_cw_first,
        "avg_cw_second": avg_cw_second,
        "avg_cw": safe_mean(cw_values),
        "avg_dia_be": safe_mean(be_values),
        "avg_dia_me": safe_mean(me_values),
        **calculate_inspec(cw_values, me_values, be_values, inspec),
        **{key: to_number_or_text(input_data.get(key)) for key in DEFECT_SOURCE_KEYS},
    }


def build_metadata(server_config: ServerConfig, submission_datetime: datetime) -> dict[str, Any]:
    """Returns the fields that are the same for every sample in the report."""
    return {
        SUBMITTED_AT_KEY: submission_datetime.replace(tzinfo=None),
        PLANT_KEY: server_config.plant,
        LOKASI_KEY: server_config.lokasi,
        BAGIAN_KEY: server_config.bagian,
    }


def combine_original_and_calculated_data(
    original_data: dict[str, Any],
    calculated_data: dict[str, Any],
    metadata: dict[str, Any],
) -> dict[str, Any]:
    """Returns metadata, then calculated fields, then the original fields.

    An original field with the same key as a calculated or metadata field is replaced by it.
    Use find_replaced_original_keys() to report those.
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
    """Returns the original keys that combine_original_and_calculated_data() replaces by
    accident. The cleaned view fields in CLEANED_SOURCE_KEYS aren't included."""
    return (original_data.keys() & (calculated_data.keys() | metadata.keys())) - CLEANED_SOURCE_KEYS


# endregion


def prepare_records(
    response: dict[str, Any],
    client_config: ClientConfig,
    local_timezone: tzinfo,
    current_time: datetime | None = None,
    last_response_id: str | None = None,
) -> PreparedRecords:
    """Validates every sample and builds the combined records, before any file is touched.

    Each record gets its response ID here, continuing after last_response_id from
    last-response.json (see client/response_ids.py).
    """
    submission_datetime = resolve_submission_datetime(response, local_timezone, current_time)
    server_config = read_server_config(response)
    metadata = build_metadata(server_config, submission_datetime)

    records: list[dict[str, Any]] = []
    column_labels: dict[str, str] = {RESPONSE_ID_KEY: ARCHIVE_FIELD_LABELS[RESPONSE_ID_KEY]}
    replaced_keys: set[str] = set()
    brands: dict[str, None] = {}

    for sample_number, raw_sample in enumerate(response["report"]["samples"], start=1):
        original, labels = normalize_source_record(raw_sample, sample_number)
        calculated = calculate_derived_data(
            original, client_config.inspec_requirements, local_timezone
        )
        brands.setdefault(calculated[BRAND_KEY])
        replaced_keys |= find_replaced_original_keys(original, calculated, metadata)
        record = combine_original_and_calculated_data(original, calculated, metadata)
        records.append(record)
        for key in record:
            column_labels.setdefault(key, ARCHIVE_FIELD_LABELS.get(key) or labels[key])

    if replaced_keys:
        LOGGER.warning(
            "These view fields have the same name as a calculated or metadata field. "
            "The calculated value is used: %s",
            sorted(replaced_keys),
        )

    records = assign_response_ids(records, last_response_id, submission_datetime.date())
    return PreparedRecords(
        records, column_labels, submission_datetime, server_config.plant, list(brands)
    )
