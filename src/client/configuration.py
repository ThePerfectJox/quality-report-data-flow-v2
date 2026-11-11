"""Loads the client device settings from config/client-config.json.

    {
        "MASTER_DATA": {
            "EXCEL_PATH": "C:\\...\\Quality Inspection Report - Master Data.xlsx",
            "SHEET_NAME": "IPC_SKT",
            "TABLE_NAME": "TableIPC"
        },
        "INSPEC_REQUIREMENTS": { "SAI": { "TARGET_CW": 2.02, ... } }
    }

SHEET_NAME and TABLE_NAME may be empty when the workbook has only one Excel Table.
A relative EXCEL_PATH is taken from the project folder. Edit the file by hand, or with
the config editor: python src/run_config_editor.py

Only REPORT_JSON_PATH, used to append a saved response file, is still read from .env.
"""

import json
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from client.errors import ConfigurationError, WorkbookNotFoundError
from client.inspec import InspecRequirement, InspecRequirementError, parse_inspec_requirements
from managers.environment_manager import ENV_REPORT_JSON_PATH, EnvironmentManager

LOGGER = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
CONFIG_FOLDER = PROJECT_ROOT / "config"
ENV_FILE = PROJECT_ROOT / ".env"
CLIENT_CONFIG_FILE = CONFIG_FOLDER / "client-config.json"
CLIENT_CONFIG_LABEL = "config/client-config.json"
SUPPORTED_WORKBOOK_SUFFIXES = {".xlsx"}

# Keys in client-config.json. Plant, Lokasi and Bagian come from the server's
# server-config.json, and Group, Brand and Set Sample from the summary_report view.
CONFIG_MASTER_DATA = "MASTER_DATA"
CONFIG_INSPEC_REQUIREMENTS = "INSPEC_REQUIREMENTS"
CLIENT_CONFIG_KEYS = (CONFIG_MASTER_DATA, CONFIG_INSPEC_REQUIREMENTS)

# Keys of the MASTER_DATA section.
MASTER_DATA_EXCEL_PATH = "EXCEL_PATH"
MASTER_DATA_SHEET_NAME = "SHEET_NAME"
MASTER_DATA_TABLE_NAME = "TABLE_NAME"
MASTER_DATA_KEYS = (MASTER_DATA_EXCEL_PATH, MASTER_DATA_SHEET_NAME, MASTER_DATA_TABLE_NAME)

# .env settings that moved to MASTER_DATA. Still set ones are reported, because they're ignored.
MOVED_ENV_SETTINGS = {
    "EXCEL_DATABASE_PATH": MASTER_DATA_EXCEL_PATH,
    "EXCEL_DATABASE_SHEET": MASTER_DATA_SHEET_NAME,
    "EXCEL_DATABASE_TABLE": MASTER_DATA_TABLE_NAME,
}


@dataclass(frozen=True)
class ExcelConfig:
    """The master data workbook and its Excel Table, from MASTER_DATA."""

    workbook_path: Path
    worksheet_name: str | None
    table_name: str | None


@dataclass(frozen=True)
class ClientConfig:
    """Everything in config/client-config.json."""

    excel: ExcelConfig
    # {brand: requirement}. Each sample uses the entry of its brand.
    inspec_requirements: dict[str, InspecRequirement]


def resolve_project_path(value: str | Path, project_root: Path = PROJECT_ROOT) -> Path:
    """Returns the path, with a relative path taken from the project folder."""
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = project_root / path
    return path.resolve()


# region CLIENT CONFIG
def read_client_config_values(config_file: Path = CLIENT_CONFIG_FILE) -> dict[str, Any]:
    """Returns the JSON object in the file. Raises ConfigurationError if it's missing or invalid."""
    if not config_file.is_file():
        raise ConfigurationError(
            f"{config_file} was not found. Copy client-config.json.example to "
            f"{config_file.name} in the config folder, or run: python src/run_config_editor.py"
        )
    try:
        values = json.loads(config_file.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as error:
        raise ConfigurationError(f"{config_file} isn't valid JSON: {error}") from error
    except OSError as error:
        raise ConfigurationError(f"{config_file} can't be read: {error}") from error
    if not isinstance(values, dict):
        raise ConfigurationError(f"{config_file} must contain a JSON object.")
    return values


def parse_master_data(values: Any, project_root: Path = PROJECT_ROOT) -> ExcelConfig:
    """Returns the MASTER_DATA section. Raises ValueError if it's invalid.

    Doesn't check that the workbook exists: see validate_workbook_path().
    """
    if not isinstance(values, dict):
        raise ValueError(f"must be an object with {', '.join(MASTER_DATA_KEYS)}.")
    unknown = sorted(set(values) - set(MASTER_DATA_KEYS))
    if unknown:
        raise ValueError(f"has unknown keys: {', '.join(unknown)}.")

    excel_path = values.get(MASTER_DATA_EXCEL_PATH)
    if not isinstance(excel_path, str) or not excel_path.strip():
        raise ValueError(f"{MASTER_DATA_EXCEL_PATH} must be the path of the master data workbook.")
    workbook_path = resolve_project_path(excel_path.strip(), project_root)
    if workbook_path.suffix.lower() not in SUPPORTED_WORKBOOK_SUFFIXES:
        raise ValueError(
            f"{MASTER_DATA_EXCEL_PATH} must be an .xlsx workbook, not {workbook_path.name!r}."
        )

    def optional_text(key: str) -> str | None:
        value = values.get(key)
        if value is None:
            return None
        if not isinstance(value, str):
            raise ValueError(f"{key} must be text, or empty.")
        return value.strip() or None

    return ExcelConfig(
        workbook_path=workbook_path,
        worksheet_name=optional_text(MASTER_DATA_SHEET_NAME),
        table_name=optional_text(MASTER_DATA_TABLE_NAME),
    )


def load_client_configuration(
    config_file: Path = CLIENT_CONFIG_FILE,
    project_root: Path = PROJECT_ROOT,
) -> ClientConfig:
    """Reads and checks the master data settings and the inspec requirements."""
    values = read_client_config_values(config_file)

    missing_keys = [key for key in CLIENT_CONFIG_KEYS if key not in values]
    if missing_keys:
        raise ConfigurationError(
            f"{config_file} is missing: {', '.join(missing_keys)}. Add them by hand, or with "
            "the config editor: python src/run_config_editor.py"
        )

    unknown_keys = sorted(set(values) - set(CLIENT_CONFIG_KEYS))
    if unknown_keys:
        LOGGER.warning(
            "%s has keys that are no longer used, which are ignored: %s. Plant, Lokasi and "
            "Bagian come from the server's server-config.json, and Group, Brand and Set Sample "
            "from the report. You can remove these keys.",
            config_file,
            unknown_keys,
        )

    try:
        excel = parse_master_data(values[CONFIG_MASTER_DATA], project_root)
    except ValueError as error:
        raise ConfigurationError(f"{CONFIG_MASTER_DATA} in {config_file}: {error}") from error

    try:
        requirements = parse_inspec_requirements(values[CONFIG_INSPEC_REQUIREMENTS])
    except InspecRequirementError as error:
        raise ConfigurationError(f"{CONFIG_INSPEC_REQUIREMENTS} in {config_file} {error}") from error

    return ClientConfig(excel=excel, inspec_requirements=requirements)


def validate_workbook_path(workbook_path: Path) -> None:
    """Raises WorkbookNotFoundError unless the path is an existing .xlsx file."""
    setting = f"{CONFIG_MASTER_DATA}.{MASTER_DATA_EXCEL_PATH} in {CLIENT_CONFIG_LABEL}"
    if not workbook_path.exists():
        raise WorkbookNotFoundError(
            f"The Excel master data {workbook_path} doesn't exist. Check {setting}, "
            "or create the workbook first."
        )
    if not workbook_path.is_file():
        raise WorkbookNotFoundError(f"{setting} points to {workbook_path}, which isn't a file.")
    if workbook_path.suffix.lower() not in SUPPORTED_WORKBOOK_SUFFIXES:
        raise WorkbookNotFoundError(
            f"{workbook_path} isn't an .xlsx workbook. Save it as .xlsx in Excel first."
        )


# endregion


# region .env
def read_report_json_path(env_file: Path = ENV_FILE, project_root: Path = PROJECT_ROOT) -> Path | None:
    """Returns REPORT_JSON_PATH from .env, or None if it isn't set."""
    if not env_file.is_file():
        return None
    value = EnvironmentManager(env_file).get(ENV_REPORT_JSON_PATH)
    if value is None or not value.strip():
        return None
    return resolve_project_path(value.strip(), project_root)


def warn_about_moved_env_settings(env_file: Path = ENV_FILE) -> None:
    """Logs a warning for each Excel setting still in .env, because it's no longer read."""
    if not env_file.is_file():
        return
    environment = EnvironmentManager(env_file)
    for env_name, master_data_key in MOVED_ENV_SETTINGS.items():
        if environment.get(env_name):
            LOGGER.warning(
                "%s in %s is no longer used and is ignored. The setting is now %s.%s in %s.",
                env_name, env_file, CONFIG_MASTER_DATA, master_data_key, CLIENT_CONFIG_LABEL,
            )


# endregion
