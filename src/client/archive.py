"""Writes the archive workbook of each append: every view field, plus the calculated and
metadata fields, at
<folder of MASTER_DATA.EXCEL_PATH>/archived-<Plant>/<ISO year>/<ISO week>/<yy_mm_dd_s_nnn>.xlsx,
where s is the Set Sample of the rows in the file (from the view) and nnn a counter starting
at 001 for each date and Set Sample. A report with both Set Samples gives two files.
"""

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from client.errors import ArchiveCollisionError, ConfigurationError, TableHeaderError
from client.values import normalize_header_name, write_cell_value

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


def build_archive_root(workbook_folder: Path, plant: str) -> Path:
    """Returns <workbook folder>/archived-<plant>, for example archived-RK1."""
    if not plant.strip() or INVALID_FOLDER_NAME_CHARACTERS.search(plant) or plant.strip(" .") != plant:
        raise ConfigurationError(
            f"PLANT {plant!r} from the server's server-config.json can't be used in the archive "
            "folder name. It must not contain "
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
        write_cell_value(worksheet.cell(1, column), header)
    for row_number, record in enumerate(records, start=2):
        for column, key in enumerate(keys, start=1):
            cell = worksheet.cell(row_number, column)
            value = _to_archive_value(record.get(key))
            write_cell_value(cell, value)
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
