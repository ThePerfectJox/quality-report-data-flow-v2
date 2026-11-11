"""In-spec ("inspec") calculation, using the targets and tolerances of each brand.

The requirements are in client-config.json under INSPEC_REQUIREMENTS, one entry per brand:

    "INSPEC_REQUIREMENTS": {
        "SAI": {
            "TARGET_CW": 2.02, "TOLERANCE_CW": 0.14,
            "TARGET_ME": 8.0,  "TOLERANCE_ME": 0.25,
            "TARGET_BE": 10.0, "TOLERANCE_BE": 0.25
        }
    }

configuration.py reads this section. Each sample uses the entry of its brand, which
report_conversion.py takes from the group name: "SAI-2" uses "SAI". A value is in spec when
target - tolerance <= value <= target + tolerance.
"""

import math
from dataclasses import dataclass
from typing import Any, Iterable

from client.master_data_mapping import BRAND_SEPARATOR
from client.values import to_nullable_float

# Keys of each brand's entry in INSPEC_REQUIREMENTS.
TARGET_CW = "TARGET_CW"
TOLERANCE_CW = "TOLERANCE_CW"
TARGET_ME = "TARGET_ME"
TOLERANCE_ME = "TOLERANCE_ME"
TARGET_BE = "TARGET_BE"
TOLERANCE_BE = "TOLERANCE_BE"
REQUIREMENT_KEYS = (TARGET_CW, TOLERANCE_CW, TARGET_ME, TOLERANCE_ME, TARGET_BE, TOLERANCE_BE)


class InspecRequirementError(ValueError):
    """An entry in INSPEC_REQUIREMENTS is missing or invalid."""


@dataclass(frozen=True)
class SpecLimit:
    """One measurement's target and tolerance."""

    target: float
    tolerance: float

    def __post_init__(self):
        if self.tolerance < 0:
            raise InspecRequirementError("Tolerance cannot be negative.")

    @property
    def low(self) -> float:
        return self.target - self.tolerance

    @property
    def high(self) -> float:
        return self.target + self.tolerance


@dataclass(frozen=True)
class InspecRequirement:
    """The CW, ME and BE specification of one brand."""

    brand: str
    cw: SpecLimit
    me: SpecLimit
    be: SpecLimit


# region REQUIREMENTS
def parse_inspec_requirements(values: Any) -> dict[str, InspecRequirement]:
    """Returns {brand: requirement} from the INSPEC_REQUIREMENTS object.

    Raises InspecRequirementError if an entry is missing a key or has an invalid number.
    """
    if not isinstance(values, dict) or not values:
        raise InspecRequirementError("must be an object with at least one brand.")

    requirements: dict[str, InspecRequirement] = {}
    for brand, entry in values.items():
        name = brand.strip()
        if not name:
            raise InspecRequirementError("has a brand with an empty name.")
        if BRAND_SEPARATOR in name:
            # The brand is the group name up to the first "-", so this name could never match.
            raise InspecRequirementError(
                f"has the brand {name!r}, which contains {BRAND_SEPARATOR!r}. Use the part of "
                f"the group name before {BRAND_SEPARATOR!r}, such as \"SAI\" for \"SAI-2\"."
            )
        if any(existing.casefold() == name.casefold() for existing in requirements):
            raise InspecRequirementError(f"has the brand {name!r} more than once.")
        requirements[name] = _parse_requirement(name, entry)
    return requirements


def find_requirement_for_brand(
    requirements: dict[str, InspecRequirement],
    brand: str,
) -> InspecRequirement:
    """Returns the requirement of the brand, ignoring case and surrounding spaces."""
    for name, requirement in requirements.items():
        if name.casefold() == brand.strip().casefold():
            return requirement
    raise InspecRequirementError(
        f"has no entry for the brand {brand!r}. Brands: {', '.join(requirements)}."
    )


def _parse_requirement(brand: str, entry: Any) -> InspecRequirement:
    if not isinstance(entry, dict):
        raise InspecRequirementError(f"{brand!r} must be an object with {', '.join(REQUIREMENT_KEYS)}.")

    missing = [key for key in REQUIREMENT_KEYS if key not in entry]
    if missing:
        raise InspecRequirementError(f"{brand!r} is missing: {', '.join(missing)}.")
    unknown = sorted(set(entry) - set(REQUIREMENT_KEYS))
    if unknown:
        raise InspecRequirementError(f"{brand!r} has unknown keys: {', '.join(unknown)}.")

    def number(key: str) -> float:
        value = entry[key]
        if value is None:
            raise InspecRequirementError(f"{key} of {brand!r} is empty. Enter a number.")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
            raise InspecRequirementError(f"{key} of {brand!r} must be a number, not {value!r}.")
        if key.startswith("TOLERANCE") and value < 0:
            raise InspecRequirementError(f"{key} of {brand!r} cannot be negative.")
        return float(value)

    return InspecRequirement(
        brand=brand,
        cw=SpecLimit(number(TARGET_CW), number(TOLERANCE_CW)),
        me=SpecLimit(number(TARGET_ME), number(TOLERANCE_ME)),
        be=SpecLimit(number(TARGET_BE), number(TOLERANCE_BE)),
    )


# endregion


# region CALCULATION
def check_in_spec(value: Any, limit: SpecLimit) -> bool | None:
    """Returns whether low <= value <= high, or None if the value is missing."""
    number = to_nullable_float(value)
    if number is None:
        return None
    return limit.low <= number <= limit.high


def calculate_in_spec_percentage(values: Iterable[Any], limit: SpecLimit) -> float | None:
    """Returns the fraction (0 to 1) of non-missing values in spec, or None if all are missing."""
    valid_values = [number for number in map(to_nullable_float, values) if number is not None]
    if not valid_values:
        return None
    in_spec_count = sum(check_in_spec(number, limit) is True for number in valid_values)
    return in_spec_count / len(valid_values)


def calculate_inspec(
    cw_values: Iterable[Any],
    me_values: Iterable[Any],
    be_values: Iterable[Any],
    requirement: InspecRequirement,
) -> dict[str, float | None]:
    """Returns the in-spec fractions of one sample: in_spec_cw, in_spec_me and in_spec_be."""
    return {
        "in_spec_cw": calculate_in_spec_percentage(cw_values, requirement.cw),
        "in_spec_me": calculate_in_spec_percentage(me_values, requirement.me),
        "in_spec_be": calculate_in_spec_percentage(be_values, requirement.be),
    }


# endregion
