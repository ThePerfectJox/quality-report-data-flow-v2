"""Tests for src/append_data_to_excel_database.py.

Run: python -m unittest discover -s tests -v
"""

import json
import sys
import tempfile
import unittest
from datetime import datetime, time, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableFormula

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import append_data_to_excel_database as module  # noqa: E402

JAKARTA = ZoneInfo("Asia/Jakarta")
NOON = time(12, 0)

TABLE_HEADERS = [
    "ID tanggapan",
    "Tanggal pengiriman",
    "Halaman terakhir",
    "Tanggal Pemeriksaan",
    "Plant/Reg",
    "Lokasi",
    "Group/Cell",
    "No ID Pekerja",
    "Brand",
    "Set Sample",
    "Bagian",
    "Avg. CW",
    "Avg. Dia. BE",
    "Avg. Dia. ME",
    "%In Spec CW",
    "%In Spec Dia. BE",
    "%In Spec Dia. ME",
    "Check",
]

SAMPLE = {
    "ID GL": "105569",
    "Nama GL": "Karsiati",
    "ID PPSKT": "317057",
    "Nama PPSKT": "Anjarwati",
    "UB [1]": 9.93, "UB [2]": 9.24, "UB [3]": 9.71, "UB [4]": 9.67, "UB [5]": 9.3, "UB [6]": 10.0,
    "UH [1]": 7.54, "UH [2]": 8.05, "UH [3]": 8.24, "UH [4]": 8.4, "UH [5]": 8.12, "UH [6]": 8.0,
    "Avg CW [1]": 6.06,
    "Avg CW [2]": 8.1,
    "Start": "2026-11-06T07:41:59",
    "End": "2026-11-06T07:46:03",
    "Tanggal": "2026-11-06",
}

CLIENT_CONFIG = {
    "PLANT": "RK1",
    "LOKASI": "U6",
    "GROUP": 99,
    "BRAND": "DKM12",
    "BAGIAN": "Rolling & Cutting",
    "SET_SAMPLE_THRESHOLD": "12:00",
}


class SetSampleTests(unittest.TestCase):
    def test_before_threshold_is_sample_1(self):
        submitted = datetime(2026, 11, 6, 11, 59, 59, tzinfo=JAKARTA)
        self.assertEqual(module.calculate_set_sample(submitted, NOON, JAKARTA), 1)

    def test_exactly_at_threshold_is_sample_1(self):
        submitted = datetime(2026, 11, 6, 12, 0, 0, tzinfo=JAKARTA)
        self.assertEqual(module.calculate_set_sample(submitted, NOON, JAKARTA), 1)

    def test_after_threshold_is_sample_2(self):
        just_after = datetime(2026, 11, 6, 12, 0, 0, 1, tzinfo=JAKARTA)
        self.assertEqual(module.calculate_set_sample(just_after, NOON, JAKARTA), 2)
        self.assertEqual(
            module.calculate_set_sample(datetime(2026, 11, 6, 15, 16, tzinfo=JAKARTA), NOON, JAKARTA), 2
        )

    def test_utc_time_is_compared_in_jakarta_time(self):
        # 05:30 UTC is 12:30 in Jakarta (UTC+7).
        submitted = datetime(2026, 11, 6, 5, 30, tzinfo=timezone.utc)
        self.assertEqual(module.calculate_set_sample(submitted, NOON, JAKARTA), 2)

    def test_time_without_time_zone_is_rejected(self):
        with self.assertRaises(module.SubmissionTimestampError):
            module.calculate_set_sample(datetime(2026, 11, 6, 12, 30), NOON, JAKARTA)

    def test_threshold_parsing(self):
        self.assertEqual(module.parse_set_sample_threshold("12:00"), NOON)
        for invalid in ("12", "24:00", "12:60", "noon", "12:00:00", 1200):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                module.parse_set_sample_threshold(invalid)


class ColumnMappingTests(unittest.TestCase):
    def test_missing_required_column_raises(self):
        headers = [header for header in TABLE_HEADERS if header != "Brand"]
        with self.assertRaises(module.MissingExcelColumnError) as context:
            module.validate_excel_columns(headers)
        self.assertIn("'Brand'", str(context.exception))

    def test_missing_optional_column_is_reported_and_skipped(self):
        headers = [header for header in TABLE_HEADERS if header != "Brand"]
        required = module.REQUIRED_EXCEL_SOURCE_KEYS - {"brand"}
        with self.assertLogs(module.LOGGER, "WARNING") as logs:
            resolved = module.validate_excel_columns(headers, required_keys=required)
        self.assertNotIn("brand", resolved)
        self.assertIn("Brand", logs.output[0])

    def test_headers_match_by_name_ignoring_case_and_spaces(self):
        headers = ["  set   SAMPLE ", "Plant/Reg"]
        row = module.map_data_to_excel_columns(
            {"set_sample": 2, "plant": "RK1", "unmapped": "x"}, headers
        )
        self.assertEqual(row, {"  set   SAMPLE ": 2, "Plant/Reg": "RK1"})


class ArchivePathTests(unittest.TestCase):
    def setUp(self):
        self._directory = tempfile.TemporaryDirectory()
        self.archive_root = module.build_archive_root(Path(self._directory.name), "RK1")
        self.submitted = datetime(2026, 11, 6, 9, 0, tzinfo=JAKARTA)

    def tearDown(self):
        self._directory.cleanup()

    def test_archive_folder_is_named_after_the_plant(self):
        self.assertEqual(self.archive_root, Path(self._directory.name) / "archived-RK1")

    def test_plant_unusable_as_folder_name_is_rejected(self):
        for plant in ("RK/1", "RK:1", "RK1.", " RK1", ""):
            with self.subTest(plant=plant), self.assertRaises(module.ConfigurationError):
                module.build_archive_root(Path(self._directory.name), plant)

    def test_base_path_uses_iso_year_and_week(self):
        path = module.build_archive_base_path(self.archive_root, self.submitted, 1)
        self.assertEqual(path, self.archive_root / "2026" / "45" / "26_11_06_1.xlsx")

    def test_first_archive_gets_counter_001(self):
        base = module.build_archive_base_path(self.archive_root, self.submitted, 1)
        self.assertEqual(module.resolve_archive_path(base).name, "26_11_06_1_001.xlsx")

    def test_counter_continues_after_the_highest_existing_archive(self):
        base = module.build_archive_base_path(self.archive_root, self.submitted, 1)
        base.parent.mkdir(parents=True)
        base.with_name("26_11_06_1_001.xlsx").touch()
        self.assertEqual(module.resolve_archive_path(base).name, "26_11_06_1_002.xlsx")
        # A deleted 002 isn't reused, and other Set Samples have their own counter.
        base.with_name("26_11_06_1_003.xlsx").touch()
        base.with_name("26_11_06_2_007.xlsx").touch()
        self.assertEqual(module.resolve_archive_path(base).name, "26_11_06_1_004.xlsx")

    def test_counter_limit_raises_collision(self):
        base = module.build_archive_base_path(self.archive_root, self.submitted, 1)
        base.parent.mkdir(parents=True)
        base.with_name("26_11_06_1_999.xlsx").touch()
        with self.assertRaises(module.ArchiveCollisionError):
            module.resolve_archive_path(base)

    def test_existing_archive_is_never_overwritten(self):
        path = module.resolve_archive_path(
            module.build_archive_base_path(self.archive_root, self.submitted, 1)
        )
        path.parent.mkdir(parents=True)
        path.write_bytes(b"existing")
        with self.assertRaises(module.ArchiveCollisionError):
            module.create_archive_workbook([{"plant": "RK1"}], {"plant": "Plant"}, path)
        self.assertEqual(path.read_bytes(), b"existing")


class AppendEndToEndTests(unittest.TestCase):
    """Runs the whole flow in a temporary project folder."""

    def setUp(self):
        self._directory = tempfile.TemporaryDirectory()
        self.root = Path(self._directory.name)
        self.workbook_path = self.root / "excel" / "database.xlsx"
        self.workbook_path.parent.mkdir()
        self._create_database_workbook()

        (self.root / ".env").write_text(
            "REPORT_JSON_PATH=report.json\n"
            "EXCEL_DATABASE_PATH=excel/database.xlsx\n"
            "EXCEL_DATABASE_SHEET=\n"
            "EXCEL_DATABASE_TABLE=\n",
            encoding="utf-8",
        )
        (self.root / "client-config.json").write_text(json.dumps(CLIENT_CONFIG), encoding="utf-8")
        second_sample = {**SAMPLE, "ID PPSKT": "319894", "Avg CW [1]": 8.4}
        self._write_report("2026-11-06T08:16:11.983512+00:00", [SAMPLE, second_sample])

    def tearDown(self):
        self._directory.cleanup()

    def _create_database_workbook(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "IPC_SKT"
        sheet.append(TABLE_HEADERS)
        existing = dict.fromkeys(TABLE_HEADERS)
        existing.update({"ID tanggapan": "261106-001", "No ID Pekerja": 111111, "Check": 2})
        sheet.append([existing[header] for header in TABLE_HEADERS])
        sheet["B2"].number_format = "yyyy-mm-dd"
        table = Table(displayName="TableIPC", ref=f"A1:R2")
        table._initialise_columns()
        for column, header in zip(table.tableColumns, TABLE_HEADERS):
            column.name = header
        table.tableColumns[-1].calculatedColumnFormula = TableFormula(attr_text="1+1")
        sheet.add_table(table)
        other = workbook.create_sheet("Spec")
        other["A1"] = 5
        other["B1"] = "=A1*2"
        workbook.save(self.workbook_path)

    def _write_report(self, responded_at, samples):
        report = {
            "request": "REQUEST QUALITY IPC REPORT DATA",
            "responded_at": responded_at,
            "status": "ok",
            "report": {"samples_after": None, "sample_count": len(samples), "samples": samples},
        }
        (self.root / "report.json").write_text(json.dumps(report), encoding="utf-8")

    def _run(self):
        with self.assertLogs(module.LOGGER, "INFO"):
            return module.append_data_to_excel_database(
                env_file=self.root / ".env",
                client_config_file=self.root / "client-config.json",
                project_root=self.root,
            )

    def test_successful_table_append(self):
        result = self._run()

        self.assertEqual((result.appended_row_count, result.first_row, result.last_row), (2, 3, 4))
        self.assertEqual(result.set_sample, 2)  # 08:16 UTC is 15:16 in Jakarta.

        workbook = load_workbook(self.workbook_path)
        sheet = workbook["IPC_SKT"]
        self.assertEqual(sheet.tables["TableIPC"].ref, "A1:R4")
        self.assertEqual(workbook["Spec"]["B1"].value, "=A1*2")

        headers = [cell.value for cell in sheet[1]]
        row = dict(zip(headers, (cell.value for cell in sheet[3])))
        self.assertEqual(row["ID tanggapan"], "261106-002")
        self.assertEqual(row["Tanggal pengiriman"], datetime(2026, 11, 6, 15, 16, 11, 984000))
        self.assertEqual(row["Tanggal Pemeriksaan"], datetime(2026, 11, 6, 7, 41, 59))
        self.assertEqual(row["Plant/Reg"], "RK1")
        self.assertEqual(row["Group/Cell"], 99)
        self.assertEqual(row["No ID Pekerja"], 317057)
        self.assertEqual(row["Set Sample"], 2)
        self.assertAlmostEqual(row["Avg. CW"], (6.06 / 3 + 8.1 / 3) / 2)
        self.assertAlmostEqual(row["%In Spec CW"], 0.5)  # 2.02 is in spec, 2.70 isn't.
        self.assertAlmostEqual(row["%In Spec Dia. BE"], 2 / 6)  # 9.93 and 10.0 are within 10 +/- 0.25.
        self.assertIsNone(row["Halaman terakhir"])
        self.assertEqual(row["Check"], "=1+1")
        self.assertEqual(sheet["B3"].number_format, "yyyy-mm-dd")
        self.assertEqual(sheet["A4"].value, "261106-003")

        self.assertEqual(
            result.archive_path,
            self.root / "excel" / "archived-RK1" / "2026" / "45" / "26_11_06_2_001.xlsx",
        )
        archive = load_workbook(result.archive_path)
        archive_sheet = archive["Data"]
        self.assertEqual(archive_sheet.tables["ArchivedData"].ref.split(":")[1][-1], "3")
        archive_headers = [cell.value for cell in archive_sheet[1]]
        for header in ("Tanggapan ID", "Plant", "Lokasi", "Group", "Brand", "Set Sample", "Bagian",
                       "Avg Cw [First]", "In Spec Me", "UB [1]", "Nama PPSKT", "Start"):
            self.assertIn(header, archive_headers)
        self.assertEqual(len(archive_headers), len(set(archive_headers)))

    def test_response_is_appended_without_a_json_file(self):
        response = json.loads((self.root / "report.json").read_text(encoding="utf-8"))
        (self.root / "report.json").unlink()
        # REPORT_JSON_PATH isn't needed when the response is passed in.
        (self.root / ".env").write_text("EXCEL_DATABASE_PATH=excel/database.xlsx\n", encoding="utf-8")

        with self.assertLogs(module.LOGGER, "INFO"):
            result = module.append_response_to_database(
                response,
                env_file=self.root / ".env",
                client_config_file=self.root / "client-config.json",
                project_root=self.root,
            )

        self.assertEqual((result.appended_row_count, result.first_row, result.last_row), (2, 3, 4))
        self.assertEqual(result.archive_path.name, "26_11_06_2_001.xlsx")
        sheet = load_workbook(self.workbook_path)["IPC_SKT"]
        self.assertEqual(sheet.tables["TableIPC"].ref, "A1:R4")

    def test_error_response_is_rejected(self):
        before = self.workbook_path.read_bytes()
        with self.assertRaises(module.InvalidInputError):
            module.append_response_to_database(
                {"request": "X", "status": "error", "message": "Unknown request"},
                env_file=self.root / ".env",
                client_config_file=self.root / "client-config.json",
                project_root=self.root,
            )
        self.assertEqual(self.workbook_path.read_bytes(), before)

    def test_file_append_without_report_json_path_is_reported(self):
        (self.root / ".env").write_text("EXCEL_DATABASE_PATH=excel/database.xlsx\n", encoding="utf-8")
        with self.assertRaises(module.ConfigurationError):
            module.append_data_to_excel_database(
                env_file=self.root / ".env",
                client_config_file=self.root / "client-config.json",
                project_root=self.root,
            )

    def test_same_report_twice_is_rejected_as_duplicate(self):
        self._run()
        before = self.workbook_path.read_bytes()
        with self.assertRaises(module.DuplicateRecordError):
            self._run()
        self.assertEqual(self.workbook_path.read_bytes(), before)

    def test_invalid_submission_timestamp_changes_nothing(self):
        self._write_report("not a time", [SAMPLE])
        before = self.workbook_path.read_bytes()
        with self.assertRaises(module.SubmissionTimestampError):
            module.append_data_to_excel_database(
                env_file=self.root / ".env",
                client_config_file=self.root / "client-config.json",
                project_root=self.root,
            )
        self.assertEqual(self.workbook_path.read_bytes(), before)

    def test_missing_workbook_is_reported(self):
        self.workbook_path.unlink()
        with self.assertRaises(module.WorkbookNotFoundError):
            module.load_excel_configuration(self.root / ".env", self.root)

    def test_missing_env_file_is_reported(self):
        with self.assertRaises(module.ConfigurationError):
            module.load_excel_configuration(self.root / "missing.env", self.root)


if __name__ == "__main__":
    unittest.main()
