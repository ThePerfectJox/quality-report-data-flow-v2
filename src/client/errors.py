"""Errors raised while appending a report to the Excel master data. Each message explains what to fix."""

from pathlib import Path


class AppendDatabaseError(Exception):
    """The append can't go ahead. The message explains what to fix."""


class ConfigurationError(AppendDatabaseError):
    """.env or client-config.json is missing or contains an invalid value."""


class InvalidInputError(AppendDatabaseError):
    """The report JSON is missing, isn't valid JSON, or doesn't have the expected structure."""


class SubmissionTimestampError(InvalidInputError):
    """The report's submission timestamp is invalid."""


class ExcelDatabaseError(AppendDatabaseError):
    """The master data workbook, worksheet or Excel Table can't be used."""


class WorkbookNotFoundError(ExcelDatabaseError):
    """The master data workbook doesn't exist."""


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
    """Samples in the report are already in the master data."""


class ArchiveCollisionError(AppendDatabaseError):
    """No free archive filename is left for this date and Set Sample."""


class PostAppendError(AppendDatabaseError):
    """The rows are in the master data, but a step after the append failed.

    Don't run the append again: the rows would be rejected as duplicates.
    """

    def __init__(self, message: str, database_path: Path, appended_row_count: int):
        super().__init__(message)
        self.database_path = database_path
        self.appended_row_count = appended_row_count


class ArchiveCreationError(PostAppendError):
    """The master data append succeeded, but the archive workbook couldn't be created."""


class LastResponseSaveError(PostAppendError):
    """The master data append succeeded, but last-response.json couldn't be updated."""
