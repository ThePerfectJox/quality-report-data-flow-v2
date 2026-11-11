"""Small helpers for reading report values and writing them to Excel."""

import math
import re
from typing import Any, Iterable
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, Cell

from client.errors import ConfigurationError, InvalidInputError

LOCAL_TIMEZONE_NAME = "Asia/Jakarta"
NULL_TEXT_VALUES = {"", "na", "n/a", "nan", "none", "null", "<na>", "nat", "-"}


def get_local_timezone() -> ZoneInfo:
    """Returns the Asia/Jakarta time zone."""
    try:
        return ZoneInfo(LOCAL_TIMEZONE_NAME)
    except ZoneInfoNotFoundError as error:
        raise ConfigurationError(
            f"The time zone {LOCAL_TIMEZONE_NAME} isn't available. "
            "Install it with: python -m pip install -r requirements.txt"
        ) from error


# region NUMBERS
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


def to_number_or_text(value: Any) -> int | float | str | None:
    """Returns a number for numeric values such as 3 or "3.00", trimmed text otherwise.

    MySQL DECIMAL values arrive as text, so they're turned back into numbers for Excel.
    Missing values become None.
    """
    if is_missing(value):
        return None
    number = to_nullable_float(value)
    if number is None:
        return str(value).strip()
    return int(number) if number.is_integer() else number


# endregion


# region NAMES AND IDS
def normalize_source_key(name: str) -> str:
    """Returns the canonical key of a source field: "Avg CW [1]" and "avg_cw_1" both become
    "avg_cw_1", and "Loose / Void" becomes "loose_void"."""
    return re.sub(r"[^0-9a-z]+", "_", name.strip().casefold()).strip("_")


def normalize_header_name(name: Any) -> str:
    """Returns an Excel header for comparing: case and repeated spaces are ignored."""
    return " ".join(str(name).split()).casefold()


def normalize_identifier(value: Any, context: str = "value") -> int | str:
    """Returns a whole number for digit-only IDs such as "317057", otherwise trimmed text.

    Existing master data rows store worker IDs and groups as numbers. IDs with a leading
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


# endregion


# region EXCEL
def ensure_excel_safe_text(text: str, context: str) -> None:
    """Raises InvalidInputError if the text has control characters Excel can't store."""
    if ILLEGAL_CHARACTERS_RE.search(text):
        raise InvalidInputError(f"{context} contains control characters that Excel can't store.")


def write_cell_value(cell: Cell, value: Any) -> None:
    """Writes a value. Text starting with = stays text, or Excel would run it as a formula."""
    cell.value = value
    if isinstance(value, str) and cell.data_type == "f":
        cell.data_type = "s"


# endregion
