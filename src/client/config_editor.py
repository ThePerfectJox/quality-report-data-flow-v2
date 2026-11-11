"""A local web page for editing config/client-config.json: the master data workbook and the
inspec requirements. Start it with: python src/run_config_editor.py

The page is config-editor.html in the project folder. A browser can't write files itself,
so this small server shows the page and saves each section when its Save button is pressed.
It only listens on 127.0.0.1, so other devices can't reach it, and it only accepts requests
from its own page.

Each save checks the section with the same rules the append uses, then replaces only that
section in the file. The rest of the file is kept.
"""

import json
import os
import tempfile
from functools import partial
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any

from client.configuration import (
    CLIENT_CONFIG_FILE,
    CLIENT_CONFIG_KEYS,
    CLIENT_CONFIG_LABEL,
    CONFIG_INSPEC_REQUIREMENTS,
    CONFIG_MASTER_DATA,
    MASTER_DATA_EXCEL_PATH,
    MASTER_DATA_KEYS,
    MASTER_DATA_SHEET_NAME,
    MASTER_DATA_TABLE_NAME,
    PROJECT_ROOT,
    parse_master_data,
    read_client_config_values,
)
from client.errors import ConfigurationError
from client.inspec import REQUIREMENT_KEYS, InspecRequirementError, parse_inspec_requirements

EDITOR_HOST = "127.0.0.1"
DEFAULT_EDITOR_PORT = 8765
EDITOR_PAGE_FILE = PROJECT_ROOT / "config-editor.html"
MAX_REQUEST_BYTES = 64 * 1024

API_CONFIG = "/api/client-config"
API_SAVE_MASTER_DATA = "/api/client-config/master-data"
API_SAVE_INSPEC = "/api/client-config/inspec-requirements"
API_BROWSE_EXCEL = "/api/browse-excel"


class EditorRequestError(Exception):
    """The request can't be handled. The message is shown on the page."""


# region READING AND SAVING THE FILE
def read_editor_state(config_file: Path, project_root: Path) -> dict[str, Any]:
    """Returns the file's sections for the page, with a problem message for each invalid one.

    A missing file gives empty sections, so the page can create it.
    """
    if config_file.exists():
        values = read_client_config_values(config_file)
    else:
        values = {}

    master_data = values.get(CONFIG_MASTER_DATA)
    master_data = master_data if isinstance(master_data, dict) else {}
    master_data_problem = None
    resolved_path = None
    found = False
    try:
        excel = parse_master_data(values.get(CONFIG_MASTER_DATA), project_root)
        resolved_path = str(excel.workbook_path)
        found = excel.workbook_path.is_file()
    except ValueError as error:
        master_data_problem = str(error)

    inspec = values.get(CONFIG_INSPEC_REQUIREMENTS)
    inspec_problem = None
    try:
        parse_inspec_requirements(inspec)
    except InspecRequirementError as error:
        inspec_problem = f"{CONFIG_INSPEC_REQUIREMENTS} {error}"

    return {
        "ok": True,
        "config_file": CLIENT_CONFIG_LABEL,
        "config_path": str(config_file),
        "exists": config_file.exists(),
        "master_data": {key: _text(master_data.get(key)) for key in MASTER_DATA_KEYS},
        "master_data_status": {
            "problem": master_data_problem,
            "resolved_path": resolved_path,
            "found": found,
        },
        "inspec_requirements": inspec if isinstance(inspec, dict) else {},
        "inspec_problem": inspec_problem,
    }


def save_master_data(config_file: Path, project_root: Path, body: Any) -> dict[str, Any]:
    """Checks and saves the MASTER_DATA section. Returns the page's message."""
    if not isinstance(body, dict):
        raise EditorRequestError("The master data must be sent as an object.")
    section = {
        MASTER_DATA_EXCEL_PATH: _clean_path(body.get(MASTER_DATA_EXCEL_PATH)),
        MASTER_DATA_SHEET_NAME: _text(body.get(MASTER_DATA_SHEET_NAME)).strip(),
        MASTER_DATA_TABLE_NAME: _text(body.get(MASTER_DATA_TABLE_NAME)).strip(),
    }
    try:
        excel = parse_master_data(section, project_root)
    except ValueError as error:
        raise EditorRequestError(f"Not saved: {error}") from error

    _save_section(config_file, CONFIG_MASTER_DATA, section)
    warnings = []
    if not excel.workbook_path.is_file():
        warnings.append(
            f"Saved, but {excel.workbook_path} wasn't found. Check the path, or that "
            "OneDrive has synced the file, before the next append."
        )
    return {"ok": True, "message": "Master data saved.", "warnings": warnings}


def save_inspec_requirements(config_file: Path, body: Any) -> dict[str, Any]:
    """Checks and saves the INSPEC_REQUIREMENTS section. Returns the page's message."""
    values = body.get(CONFIG_INSPEC_REQUIREMENTS) if isinstance(body, dict) else None
    try:
        requirements = parse_inspec_requirements(values)
    except InspecRequirementError as error:
        raise EditorRequestError(f"Not saved: {CONFIG_INSPEC_REQUIREMENTS} {error}") from error

    # Saved in the usual key order, with the brand names trimmed as the append uses them.
    section = {
        brand.strip(): {key: values[brand][key] for key in REQUIREMENT_KEYS}
        for brand in values
    }
    _save_section(config_file, CONFIG_INSPEC_REQUIREMENTS, section)
    brands = ", ".join(requirements)
    return {"ok": True, "message": f"Inspec requirements saved for {brands}.", "warnings": []}


def _save_section(config_file: Path, key: str, section: Any) -> None:
    """Replaces one section and keeps the rest of the file. Written through a temporary
    file, so a failed save can't leave a half-written file."""
    values = read_client_config_values(config_file) if config_file.exists() else {}
    values[key] = section
    ordered = {name: values[name] for name in CLIENT_CONFIG_KEYS if name in values}
    ordered.update({name: value for name, value in values.items() if name not in ordered})

    config_file.parent.mkdir(parents=True, exist_ok=True)
    file_descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{config_file.name}.", suffix=".tmp", dir=config_file.parent
    )
    try:
        with os.fdopen(file_descriptor, "w", encoding="utf-8") as file:
            file.write(json.dumps(ordered, indent=4, ensure_ascii=False) + "\n")
        os.replace(temporary_name, config_file)
    except BaseException:
        if os.path.exists(temporary_name):
            os.unlink(temporary_name)
        raise


def _text(value: Any) -> str:
    return value if isinstance(value, str) else ""


def _clean_path(value: Any) -> str:
    """Trims the path, and the quotes Windows adds with "Copy as path"."""
    return _text(value).strip().strip('"').strip()


# endregion


# region FILE DIALOG
def browse_for_excel_file(current_path: str) -> str | None:
    """Opens a Windows file dialog for an .xlsx file. Returns its path, or None if cancelled."""
    try:
        import tkinter
        from tkinter import filedialog
    except ImportError as error:
        raise EditorRequestError(
            "The file dialog isn't available on this computer. Paste the path instead."
        ) from error

    current = Path(current_path) if current_path else None
    initial_folder = current.parent if current is not None and current.parent.is_dir() else Path.home()
    root = tkinter.Tk()
    try:
        root.withdraw()
        # Shows the dialog in front of the browser.
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            parent=root,
            title="Select the master data workbook",
            initialdir=str(initial_folder),
            filetypes=[("Excel workbook", "*.xlsx")],
        )
    finally:
        root.destroy()
    return str(Path(selected)) if selected else None


# endregion


# region WEB SERVER
class ConfigEditorRequestHandler(BaseHTTPRequestHandler):
    """Serves config-editor.html and the API it uses."""

    server_version = "ConfigEditor"

    def __init__(self, *args, config_file: Path, project_root: Path, page_file: Path, **kwargs):
        self.config_file = config_file
        self.project_root = project_root
        self.page_file = page_file
        super().__init__(*args, **kwargs)

    def do_GET(self):
        if not self._is_own_host():
            return
        path = self.path.split("?", 1)[0]
        if path in ("/", "/config-editor.html"):
            self._send_page()
        elif path == API_CONFIG:
            self._run(lambda: read_editor_state(self.config_file, self.project_root))
        else:
            self._send_json({"ok": False, "error": "Not found."}, HTTPStatus.NOT_FOUND)

    def do_POST(self):
        if not self._is_own_host() or not self._is_own_page():
            return
        path = self.path.split("?", 1)[0]
        routes = {
            API_SAVE_MASTER_DATA: lambda body: save_master_data(self.config_file, self.project_root, body),
            API_SAVE_INSPEC: lambda body: save_inspec_requirements(self.config_file, body),
            API_BROWSE_EXCEL: lambda body: {
                "ok": True,
                "path": browse_for_excel_file(_text(body.get("current_path")) if isinstance(body, dict) else ""),
            },
        }
        if path not in routes:
            self._send_json({"ok": False, "error": "Not found."}, HTTPStatus.NOT_FOUND)
            return
        body = self._read_json_body()
        if body is not None:
            self._run(lambda: routes[path](body))

    # Requests from other websites are refused, so they can't change the config.
    def _allowed_hosts(self) -> set[str]:
        port = self.server.server_address[1]
        return {f"{EDITOR_HOST}:{port}", f"localhost:{port}"}

    def _is_own_host(self) -> bool:
        if self.headers.get("Host", "") in self._allowed_hosts():
            return True
        self._send_json({"ok": False, "error": "Unknown host."}, HTTPStatus.FORBIDDEN)
        return False

    def _is_own_page(self) -> bool:
        origin = self.headers.get("Origin")
        if origin is not None and origin not in {f"http://{host}" for host in self._allowed_hosts()}:
            self._send_json({"ok": False, "error": "Requests from other pages are refused."}, HTTPStatus.FORBIDDEN)
            return False
        if not self.headers.get("Content-Type", "").startswith("application/json"):
            self._send_json({"ok": False, "error": "Expected JSON."}, HTTPStatus.UNSUPPORTED_MEDIA_TYPE)
            return False
        return True

    def _read_json_body(self) -> Any:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = -1
        if length < 0 or length > MAX_REQUEST_BYTES:
            self._send_json({"ok": False, "error": "The request is too large."}, HTTPStatus.REQUEST_ENTITY_TOO_LARGE)
            return None
        try:
            return json.loads(self.rfile.read(length).decode("utf-8") or "null")
        except (UnicodeDecodeError, json.JSONDecodeError):
            self._send_json({"ok": False, "error": "The request isn't valid JSON."}, HTTPStatus.BAD_REQUEST)
            return None

    def _run(self, action) -> None:
        """Runs an API action. Expected problems are shown on the page as a message."""
        try:
            result = action()
        except (EditorRequestError, ConfigurationError) as error:
            result = {"ok": False, "error": str(error)}
        except OSError as error:
            result = {"ok": False, "error": f"{CLIENT_CONFIG_LABEL} couldn't be saved: {error}"}
        self._send_json(result)

    def _send_page(self) -> None:
        try:
            content = self.page_file.read_bytes()
        except OSError:
            self._send_json({"ok": False, "error": f"{self.page_file} is missing."}, HTTPStatus.NOT_FOUND)
            return
        self._send(content, "text/html; charset=utf-8")

    def _send_json(self, data: dict[str, Any], status: HTTPStatus = HTTPStatus.OK) -> None:
        self._send(json.dumps(data, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8", status)

    def _send(self, content: bytes, content_type: str, status: HTTPStatus = HTTPStatus.OK) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.end_headers()
        self.wfile.write(content)

    def log_message(self, format: str, *args) -> None:
        # Only API requests are printed, so the terminal shows what was saved.
        if self.command == "POST":
            print(f"{self.command} {self.path} -> {args[1] if len(args) > 1 else ''}")


def create_config_editor_server(
    port: int = DEFAULT_EDITOR_PORT,
    config_file: Path = CLIENT_CONFIG_FILE,
    project_root: Path = PROJECT_ROOT,
    page_file: Path = EDITOR_PAGE_FILE,
) -> HTTPServer:
    """Returns the editor's server, listening on 127.0.0.1. Call serve_forever() to run it.

    It answers one request at a time, so the file dialog runs on the server's own thread.
    """
    handler = partial(
        ConfigEditorRequestHandler,
        config_file=config_file,
        project_root=project_root,
        page_file=page_file,
    )
    return HTTPServer((EDITOR_HOST, port), handler)


# endregion
