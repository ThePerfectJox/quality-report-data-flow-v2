"""Tests for the client's append of the report to the Excel master data (src/client/).

Run: python -m unittest discover -s tests -v
"""

import json
import sys
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableFormula

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from client import append_to_database, archive, errors, excel_database, response_ids  # noqa: E402
from client import master_data_mapping as mapping  # noqa: E402
from client.report_conversion import parse_set_sample  # noqa: E402

JAKARTA = ZoneInfo("Asia/Jakarta")

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
    *mapping.REPORT_TO_MASTER_COLUMNS.values(),
    "Total VFI Cig.",
    "Avg. CW",
    "Avg. Dia. BE",
    "Avg. Dia. ME",
    "%In Spec CW",
    "%In Spec Dia. BE",
    "%In Spec Dia. ME",
    "Check",
]

# One row of the summary_report view in set-up/view.sql.
SAMPLE = {
    "ID Group / Cell": "105569",
    "Nama Group / Cell": "SAI-2",
    "ID PPSKT": "317057",
    "Nama PPSKT": "Anjarwati",
    "Set Sample": 1,
    **{column: 0 for column in mapping.REPORT_TO_MASTER_COLUMNS},
    "Loose / Void": "2.00",
    "Trial 1": 0,
    **{f"Reserved {number}": 0 for number in range(2, 11)},
    "UB [1]": 9.93, "UB [2]": 9.24, "UB [3]": 9.71, "UB [4]": 9.67, "UB [5]": 9.3, "UB [6]": 10.0,
    "UH [1]": 7.54, "UH [2]": 8.05, "UH [3]": 8.24, "UH [4]": 8.4, "UH [5]": 8.12, "UH [6]": 8.0,
    "Avg CW [1]": 6.06,
    "Avg CW [2]": 8.1,
    "Start": "2026-11-06T07:41:59",
    "End": "2026-11-06T07:46:03",
    "Tanggal": "2026-11-06",
}

SAI_REQUIREMENT = {
    "TARGET_CW": 2.02, "TOLERANCE_CW": 0.14,
    "TARGET_ME": 8.0, "TOLERANCE_ME": 0.25,
    "TARGET_BE": 10.0, "TOLERANCE_BE": 0.25,
}

CLIENT_CONFIG = {
    # Relative to the test's project folder.
    "MASTER_DATA": {"EXCEL_PATH": "excel/database.xlsx", "SHEET_NAME": "", "TABLE_NAME": ""},
    "INSPEC_REQUIREMENTS": {
        "SAI": SAI_REQUIREMENT,
        # A different CW spec, so the tests can tell which brand's entry was used.
        "DSS": {**SAI_REQUIREMENT, "TARGET_CW": 2.36, "TOLERANCE_CW": 0.35},
    },
}

# The server's server-config.json, as sent in the response.
SERVER_CONFIG = {"PLANT": "RK1", "LOKASI": "U6", "BAGIAN": "Rolling & Cutting"}


class SetSampleTests(unittest.TestCase):
    """Set Sample is the view's value, never calculated on the client."""

    def test_whole_numbers_are_accepted(self):
        for value, expected in ((1, 1), (2, 2), ("2", 2), (2.0, 2), ("1.00", 1)):
            with self.subTest(value=value):
                self.assertEqual(parse_set_sample(value, "sample"), expected)

    def test_empty_or_invalid_set_sample_is_rejected(self):
        for value in (None, "", "null", "x", 0, 1.5):
            with self.subTest(value=value), self.assertRaises(errors.InvalidInputError):
                parse_set_sample(value, "sample")


class ResponseIdTests(unittest.TestCase):
    def test_numbering_continues_on_the_same_date(self):
        records = response_ids.assign_response_ids([{}, {}], "261106009", date(2026, 11, 6))
        self.assertEqual([record["tanggapan_id"] for record in records], ["261106010", "261106011"])

    def test_numbering_starts_at_001_on_a_new_date(self):
        records = response_ids.assign_response_ids([{}], "261105007", date(2026, 11, 6))
        self.assertEqual(records[0]["tanggapan_id"], "261106001")

    def test_old_hyphenated_last_id_still_continues(self):
        records = response_ids.assign_response_ids([{}], "261106-009", date(2026, 11, 6))
        self.assertEqual(records[0]["tanggapan_id"], "261106010")

    def test_numbering_starts_at_001_without_a_last_id(self):
        records = response_ids.assign_response_ids([{}], None, date(2026, 11, 6))
        self.assertEqual(records[0]["tanggapan_id"], "261106001")


class ColumnMappingTests(unittest.TestCase):
    def test_missing_required_column_raises(self):
        headers = [header for header in TABLE_HEADERS if header != "Brand"]
        with self.assertRaises(errors.MissingExcelColumnError) as context:
            excel_database.validate_excel_columns(headers)
        self.assertIn("'Brand'", str(context.exception))

    def test_missing_optional_column_is_reported_and_skipped(self):
        headers = [header for header in TABLE_HEADERS if header != "Brand"]
        required = mapping.REQUIRED_MASTER_DATA_KEYS - {"brand"}
        with self.assertLogs(excel_database.LOGGER, "WARNING") as logs:
            resolved = excel_database.validate_excel_columns(headers, required_keys=required)
        self.assertNotIn("brand", resolved)
        self.assertIn("Brand", logs.output[0])

    def test_headers_match_by_name_ignoring_case_and_spaces(self):
        headers = ["  set   SAMPLE ", "Plant/Reg"]
        row = excel_database.map_data_to_excel_columns(
            {"set_sample": 2, "plant": "RK1", "unmapped": "x"}, headers
        )
        self.assertEqual(row, {"  set   SAMPLE ": 2, "Plant/Reg": "RK1"})


class ArchivePathTests(unittest.TestCase):
    def setUp(self):
        self._directory = tempfile.TemporaryDirectory()
        self.archive_root = archive.build_archive_root(Path(self._directory.name), "RK1")
        self.submitted = datetime(2026, 11, 6, 9, 0, tzinfo=JAKARTA)

    def tearDown(self):
        self._directory.cleanup()

    def test_archive_folder_is_named_after_the_plant(self):
        self.assertEqual(self.archive_root, Path(self._directory.name) / "archived-RK1")

    def test_plant_unusable_as_folder_name_is_rejected(self):
        for plant in ("RK/1", "RK:1", "RK1.", " RK1", ""):
            with self.subTest(plant=plant), self.assertRaises(errors.ConfigurationError):
                archive.build_archive_root(Path(self._directory.name), plant)

    def test_base_path_uses_iso_year_and_week(self):
        path = archive.build_archive_base_path(self.archive_root, self.submitted, 1)
        self.assertEqual(path, self.archive_root / "2026" / "45" / "26_11_06_1.xlsx")

    def test_first_archive_gets_counter_001(self):
        base = archive.build_archive_base_path(self.archive_root, self.submitted, 1)
        self.assertEqual(archive.resolve_archive_path(base).name, "26_11_06_1_001.xlsx")

    def test_counter_continues_after_the_highest_existing_archive(self):
        base = archive.build_archive_base_path(self.archive_root, self.submitted, 1)
        base.parent.mkdir(parents=True)
        base.with_name("26_11_06_1_001.xlsx").touch()
        self.assertEqual(archive.resolve_archive_path(base).name, "26_11_06_1_002.xlsx")
        # A deleted 002 isn't reused, and other Set Samples have their own counter.
        base.with_name("26_11_06_1_003.xlsx").touch()
        base.with_name("26_11_06_2_007.xlsx").touch()
        self.assertEqual(archive.resolve_archive_path(base).name, "26_11_06_1_004.xlsx")

    def test_counter_limit_raises_collision(self):
        base = archive.build_archive_base_path(self.archive_root, self.submitted, 1)
        base.parent.mkdir(parents=True)
        base.with_name("26_11_06_1_999.xlsx").touch()
        with self.assertRaises(errors.ArchiveCollisionError):
            archive.resolve_archive_path(base)

    def test_existing_archive_is_never_overwritten(self):
        path = archive.resolve_archive_path(
            archive.build_archive_base_path(self.archive_root, self.submitted, 1)
        )
        path.parent.mkdir(parents=True)
        path.write_bytes(b"existing")
        with self.assertRaises(errors.ArchiveCollisionError):
            archive.create_archive_workbook([{"plant": "RK1"}], {"plant": "Plant"}, path)
        self.assertEqual(path.read_bytes(), b"existing")


class AppendEndToEndTests(unittest.TestCase):
    """Runs the whole flow in a temporary project folder."""

    def setUp(self):
        self._directory = tempfile.TemporaryDirectory()
        self.root = Path(self._directory.name)
        self.workbook_path = self.root / "excel" / "database.xlsx"
        self.workbook_path.parent.mkdir()
        self._create_database_workbook()

        (self.root / ".env").write_text("REPORT_JSON_PATH=report.json\n", encoding="utf-8")
        self._write_client_config(CLIENT_CONFIG)
        # The master data's existing row has ID 261106001.
        self._write_last_response_id("261106001")
        second_sample = {**SAMPLE, "ID PPSKT": "319894", "Avg CW [1]": 8.4, "Set Sample": 2}
        self._write_report("2026-11-06T08:16:11.983512+00:00", [SAMPLE, second_sample])

    def tearDown(self):
        self._directory.cleanup()

    @property
    def last_response_file(self):
        return self.root / "last-response.json"

    def _write_last_response_id(self, response_id):
        self.last_response_file.write_text(
            json.dumps({"last_response_id": response_id, "saved_at": "2026-11-06T09:00:00+07:00"}),
            encoding="utf-8",
        )

    def _read_last_response_id(self):
        return json.loads(self.last_response_file.read_text(encoding="utf-8"))["last_response_id"]

    def _create_database_workbook(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "IPC_SKT"
        sheet.append(TABLE_HEADERS)
        existing = dict.fromkeys(TABLE_HEADERS)
        existing.update({"ID tanggapan": "261106001", "No ID Pekerja": 111111, "Check": 2})
        sheet.append([existing[header] for header in TABLE_HEADERS])
        sheet["B2"].number_format = "yyyy-mm-dd"
        last_column = sheet.cell(1, len(TABLE_HEADERS)).column_letter
        table = Table(displayName="TableIPC", ref=f"A1:{last_column}2")
        table._initialise_columns()
        for column, header in zip(table.tableColumns, TABLE_HEADERS):
            column.name = header
        table.tableColumns[-1].calculatedColumnFormula = TableFormula(attr_text="1+1")
        sheet.add_table(table)
        other = workbook.create_sheet("Spec")
        other["A1"] = 5
        other["B1"] = "=A1*2"
        workbook.save(self.workbook_path)
        self.last_column = last_column

    @property
    def client_config_file(self):
        return self.root / "config" / "client-config.json"

    def _write_client_config(self, config):
        self.client_config_file.parent.mkdir(exist_ok=True)
        self.client_config_file.write_text(json.dumps(config), encoding="utf-8")

    def _write_report(self, responded_at, samples, server_config=SERVER_CONFIG):
        report = {
            "request": "REQUEST QUALITY IPC REPORT DATA",
            "responded_at": responded_at,
            "status": "ok",
            "server_config": server_config,
            "report": {"samples_after": None, "sample_count": len(samples), "samples": samples},
        }
        (self.root / "report.json").write_text(json.dumps(report), encoding="utf-8")

    def _run(self):
        with self.assertLogs("client", "INFO"):
            return append_to_database.append_data_to_excel_database(
                env_file=self.root / ".env",
                client_config_file=self.client_config_file,
                project_root=self.root,
            )

    def test_successful_table_append(self):
        result = self._run()

        self.assertEqual((result.appended_row_count, result.first_row, result.last_row), (2, 3, 4))
        self.assertEqual(result.set_samples, [1, 2])  # From the view.

        workbook = load_workbook(self.workbook_path)
        sheet = workbook["IPC_SKT"]
        self.assertEqual(sheet.tables["TableIPC"].ref, f"A1:{self.last_column}4")
        self.assertEqual(workbook["Spec"]["B1"].value, "=A1*2")

        headers = [cell.value for cell in sheet[1]]
        row = dict(zip(headers, (cell.value for cell in sheet[3])))
        self.assertEqual(row["ID tanggapan"], "261106002")
        self.assertEqual(row["Tanggal pengiriman"], datetime(2026, 11, 6, 15, 16, 11, 984000))
        self.assertEqual(row["Tanggal Pemeriksaan"], datetime(2026, 11, 6, 7, 41, 59))
        self.assertEqual(row["Plant/Reg"], "RK1")  # From the server's server-config.json.
        self.assertEqual(row["Lokasi"], "U6")
        self.assertEqual(row["Bagian"], "Rolling & Cutting")
        self.assertEqual(row["Group/Cell"], 105569)  # From the view's ID Group / Cell.
        self.assertEqual(row["Brand"], "SAI")  # From the group name "SAI-2".
        self.assertEqual(row["No ID Pekerja"], 317057)
        # From the view. The submission time, 15:16 in Jakarta, plays no part.
        self.assertEqual(row["Set Sample"], 1)
        self.assertEqual(row["VFI NC LV"], 2)  # "2.00" from the view becomes a number.
        self.assertEqual(row["VFI NC Incompletely Filled"], 0)
        self.assertEqual(row["Total VFI Cig."], 2)  # Only Loose / Void is 2; the others are 0.
        self.assertAlmostEqual(row["Avg. CW"], (6.06 / 3 + 8.1 / 3) / 2)
        self.assertAlmostEqual(row["%In Spec CW"], 0.5)  # 2.02 is in spec, 2.70 isn't.
        self.assertAlmostEqual(row["%In Spec Dia. BE"], 2 / 6)  # 9.93 and 10.0 are within 10 +/- 0.25.
        self.assertIsNone(row["Halaman terakhir"])
        self.assertEqual(row["Check"], "=1+1")
        self.assertEqual(sheet["B3"].number_format, "yyyy-mm-dd")
        self.assertEqual(sheet["A4"].value, "261106003")
        self.assertEqual(self._read_last_response_id(), "261106003")
        second_row = dict(zip(headers, (cell.value for cell in sheet[4])))
        self.assertEqual(second_row["Set Sample"], 2)

        # One archive per Set Sample, each holding only its own row.
        week_folder = self.root / "excel" / "archived-RK1" / "2026" / "45"
        self.assertEqual(
            result.archive_paths,
            [week_folder / "26_11_06_1_001.xlsx", week_folder / "26_11_06_2_001.xlsx"],
        )
        for archive_path, worker_id in zip(result.archive_paths, ("317057", "319894")):
            archive_sheet = load_workbook(archive_path)["Data"]
            self.assertEqual(archive_sheet.tables["ArchivedData"].ref.split(":")[1][-1], "2")
            archive_row = dict(zip((cell.value for cell in archive_sheet[1]), (cell.value for cell in archive_sheet[2])))
            self.assertEqual(archive_row["ID PPSKT"], worker_id)
        archive_sheet = load_workbook(result.archive_paths[0])["Data"]
        archive_headers = [cell.value for cell in archive_sheet[1]]
        for header in ("Tanggapan ID", "Plant", "Lokasi", "Group", "Brand", "Set Sample", "Bagian",
                       "Avg Cw [First]", "In Spec Me", "UB [1]", "Nama PPSKT", "Start",
                       "Loose / Void", "Reserved 10"):
            self.assertIn(header, archive_headers)
        self.assertEqual(len(archive_headers), len(set(archive_headers)))

    def test_each_sample_uses_its_brand_requirement(self):
        dss_sample = {**SAMPLE, "ID PPSKT": "319894", "Nama Group / Cell": "dss-1"}
        self._write_report("2026-11-06T08:16:11+00:00", [SAMPLE, dss_sample])
        self._run()
        sheet = load_workbook(self.workbook_path)["IPC_SKT"]
        headers = [cell.value for cell in sheet[1]]
        sai_row = dict(zip(headers, (cell.value for cell in sheet[3])))
        dss_row = dict(zip(headers, (cell.value for cell in sheet[4])))
        self.assertAlmostEqual(sai_row["%In Spec CW"], 0.5)  # Only 2.02 is within 2.02 +/- 0.14.
        self.assertEqual(dss_row["Brand"], "dss")
        self.assertAlmostEqual(dss_row["%In Spec CW"], 1.0)  # 2.02 and 2.70 are within 2.36 +/- 0.35.

    def test_brand_without_requirement_is_rejected(self):
        self._write_report("2026-11-06T08:16:11+00:00", [{**SAMPLE, "Nama Group / Cell": "XYZ-1"}])
        before = self.workbook_path.read_bytes()
        with self.assertRaises(errors.ConfigurationError) as context:
            self._run_without_logs()
        self.assertIn("'XYZ'", str(context.exception))
        self.assertEqual(self.workbook_path.read_bytes(), before)

    def test_response_without_server_config_is_rejected(self):
        self._write_report("2026-11-06T08:16:11+00:00", [SAMPLE], server_config=None)
        before = self.workbook_path.read_bytes()
        with self.assertRaises(errors.InvalidInputError) as context:
            self._run_without_logs()
        self.assertIn("server_config", str(context.exception))
        self.assertEqual(self.workbook_path.read_bytes(), before)

    def _run_without_logs(self):
        return append_to_database.append_data_to_excel_database(
            env_file=self.root / ".env",
            client_config_file=self.client_config_file,
            project_root=self.root,
        )

    def test_response_is_appended_without_a_json_file(self):
        response = json.loads((self.root / "report.json").read_text(encoding="utf-8"))
        (self.root / "report.json").unlink()
        # .env and its REPORT_JSON_PATH aren't needed when the response is passed in.
        (self.root / ".env").unlink()

        with self.assertLogs("client", "INFO"):
            result = append_to_database.append_response_to_database(
                response,
                env_file=self.root / ".env",
                client_config_file=self.client_config_file,
                project_root=self.root,
            )

        self.assertEqual((result.appended_row_count, result.first_row, result.last_row), (2, 3, 4))
        self.assertEqual(
            [path.name for path in result.archive_paths],
            ["26_11_06_1_001.xlsx", "26_11_06_2_001.xlsx"],
        )
        sheet = load_workbook(self.workbook_path)["IPC_SKT"]
        self.assertEqual(sheet.tables["TableIPC"].ref, f"A1:{self.last_column}4")

    def test_error_response_is_rejected(self):
        before = self.workbook_path.read_bytes()
        with self.assertRaises(errors.InvalidInputError):
            append_to_database.append_response_to_database(
                {"request": "X", "status": "error", "message": "Unknown request"},
                env_file=self.root / ".env",
                client_config_file=self.client_config_file,
                project_root=self.root,
            )
        self.assertEqual(self.workbook_path.read_bytes(), before)

    def test_response_ids_continue_after_last_response_file(self):
        # The master data only has 261106001, so the IDs must come from the file.
        self._write_last_response_id("261106010")
        self._run()
        sheet = load_workbook(self.workbook_path)["IPC_SKT"]
        self.assertEqual([sheet["A3"].value, sheet["A4"].value], ["261106011", "261106012"])
        self.assertEqual(self._read_last_response_id(), "261106012")

    def test_second_report_continues_numbering(self):
        self._run()
        third_sample = {**SAMPLE, "ID PPSKT": "320001"}
        self._write_report("2026-11-06T09:00:00+00:00", [third_sample])
        self._run()
        sheet = load_workbook(self.workbook_path)["IPC_SKT"]
        self.assertEqual(sheet["A5"].value, "261106004")
        self.assertEqual(self._read_last_response_id(), "261106004")

    def test_last_response_file_behind_master_data_is_rejected(self):
        # Without the file, numbering starts at 261106001, which the master data already has.
        self.last_response_file.unlink()
        before = self.workbook_path.read_bytes()
        with self.assertRaises(errors.DuplicateRecordError) as context:
            self._run_without_logs()
        self.assertIn("261106001", str(context.exception))
        self.assertIn("last-response.json", str(context.exception))
        self.assertEqual(self.workbook_path.read_bytes(), before)
        self.assertFalse(self.last_response_file.exists())

    def test_damaged_last_response_file_is_rejected(self):
        self.last_response_file.write_text("{\"last_response_id\": \"not an id\"}", encoding="utf-8")
        before = self.workbook_path.read_bytes()
        with self.assertRaises(errors.ConfigurationError):
            self._run_without_logs()
        self.assertEqual(self.workbook_path.read_bytes(), before)

    def test_total_vfi_cig_is_the_sum_of_all_defect_counts(self):
        counts = {column: number for number, column in enumerate(mapping.REPORT_TO_MASTER_COLUMNS, start=1)}
        # Text from MySQL, an empty count and a non-number are handled like the view gives them.
        counts.update({"Loose / Void": "1.00", "Wrinkled": None, "TPIS": "x"})
        expected = sum(range(1, 23)) - 2 - 7  # Wrinkled (2) and TPIS (7) are left out.
        self._write_report("2026-11-06T08:16:11+00:00", [{**SAMPLE, **counts}])
        self._run()
        sheet = load_workbook(self.workbook_path)["IPC_SKT"]
        row = dict(zip((cell.value for cell in sheet[1]), (cell.value for cell in sheet[3])))
        self.assertEqual(row["Total VFI Cig."], expected)
        self.assertEqual(len(mapping.REPORT_TO_MASTER_COLUMNS), 22)

    def test_sample_with_empty_set_sample_is_rejected(self):
        self._write_report("2026-11-06T08:16:11+00:00", [{**SAMPLE, "Set Sample": None}])
        before = self.workbook_path.read_bytes()
        with self.assertRaises(errors.InvalidInputError) as context:
            self._run_without_logs()
        self.assertIn("Set Sample", str(context.exception))
        self.assertEqual(self.workbook_path.read_bytes(), before)

    def test_sample_without_a_view_column_is_rejected(self):
        incomplete = {key: value for key, value in SAMPLE.items() if key != "Clove Spot"}
        self._write_report("2026-11-06T08:16:11+00:00", [incomplete])
        before = self.workbook_path.read_bytes()
        with self.assertRaises(errors.InvalidInputError) as context:
            self._run_without_logs()
        self.assertIn("clove_spot", str(context.exception))
        self.assertEqual(self.workbook_path.read_bytes(), before)

    def test_file_append_without_report_json_path_is_reported(self):
        (self.root / ".env").unlink()
        with self.assertRaises(errors.ConfigurationError):
            self._run_without_logs()

    def test_same_report_twice_is_rejected_as_duplicate(self):
        self._run()
        before = self.workbook_path.read_bytes()
        with self.assertRaises(errors.DuplicateRecordError):
            self._run()
        self.assertEqual(self.workbook_path.read_bytes(), before)

    def test_invalid_submission_timestamp_changes_nothing(self):
        self._write_report("not a time", [SAMPLE])
        before = self.workbook_path.read_bytes()
        with self.assertRaises(errors.SubmissionTimestampError):
            self._run_without_logs()
        self.assertEqual(self.workbook_path.read_bytes(), before)

    def test_missing_workbook_is_reported(self):
        self.workbook_path.unlink()
        with self.assertRaises(errors.WorkbookNotFoundError) as context:
            self._run_without_logs()
        self.assertIn("MASTER_DATA.EXCEL_PATH", str(context.exception))

    def test_missing_client_config_is_reported(self):
        self.client_config_file.unlink()
        with self.assertRaises(errors.ConfigurationError):
            self._run_without_logs()

    def test_excel_settings_left_in_env_are_reported(self):
        (self.root / ".env").write_text(
            "REPORT_JSON_PATH=report.json\nEXCEL_DATABASE_PATH=old/database.xlsx\n", encoding="utf-8"
        )
        with self.assertLogs("client", "WARNING") as logs:
            self._run_without_logs()
        self.assertTrue(any("EXCEL_DATABASE_PATH" in line for line in logs.output))


if __name__ == "__main__":
    unittest.main()
