"""Reads and updates the values in the .env file."""

import os
from pathlib import Path
from typing import overload

from dotenv import dotenv_values, set_key

ENV_FILE = Path(__file__).resolve().parents[2] / ".env"

ENV_DB_HOST = "DB_HOST"
ENV_DB_PORT = "DB_PORT"
ENV_DB_NAME = "DB_NAME"
ENV_DB_USER = "DB_USER"
ENV_DB_PASSWORD_ENCRYPTED = "DB_PASSWORD_ENCRYPTED"
ENV_DB_ENCRYPTION_KEY = "DB_ENCRYPTION_KEY"
ENV_SERVER_LISTEN_HOST = "SERVER_LISTEN_HOST"
ENV_CLIENT_BIND_IP = "CLIENT_BIND_IP"
ENV_CLIENT_DEFAULT_SERVER_HOST = "CLIENT_DEFAULT_SERVER_HOST"
ENV_REPORT_JSON_PATH = "REPORT_JSON_PATH"
ENV_EXCEL_DATABASE_PATH = "EXCEL_DATABASE_PATH"
ENV_EXCEL_DATABASE_SHEET = "EXCEL_DATABASE_SHEET"
ENV_EXCEL_DATABASE_TABLE = "EXCEL_DATABASE_TABLE"


class EnvironmentValueError(Exception):
    """A value in the environment is missing or invalid. The message explains which one."""


class EnvironmentManager:
    """Gives access to settings. A real environment variable overrides the .env file."""

    def __init__(self, env_file: Path = ENV_FILE):
        self._env_file = Path(env_file)
        # interpolate=False keeps values exactly as written, even if they contain a $.
        self._file_values = dotenv_values(self._env_file, interpolate=False)

    @property
    def env_file(self) -> Path:
        return self._env_file

    @overload
    def get(self, name: str, default: str) -> str: ...

    @overload
    def get(self, name: str, default: None = None) -> str | None: ...

    def get(self, name: str, default: str | None = None) -> str | None:
        """Returns the value, or the default if it's missing or empty."""
        return os.environ.get(name, self._file_values.get(name)) or default

    def get_required(self, name: str) -> str:
        """Returns the value. Raises EnvironmentValueError if it's missing or empty."""
        value = self.get(name)
        if value is None:
            raise EnvironmentValueError(
                f"{name} is not set. Add it to {self._env_file}, using .env.example as a guide."
            )
        return value

    def get_int(self, name: str, default: int | None = None) -> int | None:
        """Returns the value as a whole number, or the default if it's missing or empty."""
        value = self.get(name)
        if value is None:
            return default
        try:
            return int(value)
        except ValueError as error:
            raise EnvironmentValueError(f"{name} must be a whole number") from error

    def set_value(self, name: str, value: str):
        """Saves the value to the .env file, keeping everything else in the file."""
        self._env_file.touch(exist_ok=True)
        set_key(self._env_file, name, value)
        self._file_values[name] = value
