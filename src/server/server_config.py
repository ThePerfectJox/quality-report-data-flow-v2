"""Reads config/server-config.json: where this server device is. Every report response includes it.

The file is read for each request, so a change takes effect without restarting the server.
"""

import json
from pathlib import Path

SERVER_CONFIG_FILE = Path(__file__).resolve().parents[2] / "config" / "server-config.json"

CONFIG_PLANT = "PLANT"
CONFIG_LOKASI = "LOKASI"
CONFIG_BAGIAN = "BAGIAN"
SERVER_CONFIG_KEYS = (CONFIG_PLANT, CONFIG_LOKASI, CONFIG_BAGIAN)


class ServerConfigError(Exception):
    """server-config.json is missing or invalid. The message explains what to fix."""


def load_server_config(config_file: Path = SERVER_CONFIG_FILE) -> dict[str, str]:
    """Returns {"PLANT": ..., "LOKASI": ..., "BAGIAN": ...}, each non-empty text."""
    if not config_file.is_file():
        raise ServerConfigError(
            f"{config_file} was not found. Create it with {', '.join(SERVER_CONFIG_KEYS)}."
        )
    try:
        values = json.loads(config_file.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as error:
        raise ServerConfigError(f"{config_file} isn't valid JSON: {error}") from error
    except OSError as error:
        raise ServerConfigError(f"{config_file} can't be read: {error}") from error

    if not isinstance(values, dict):
        raise ServerConfigError(f"{config_file} must contain a JSON object.")

    config: dict[str, str] = {}
    for key in SERVER_CONFIG_KEYS:
        value = values.get(key)
        if not isinstance(value, str) or not value.strip():
            raise ServerConfigError(f"{key} in {config_file} must be non-empty text.")
        config[key] = value.strip()
    return config
