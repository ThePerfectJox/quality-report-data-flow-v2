"""Tests for src/server/reporting.py: remembering which samples were sent.

Run: python -m unittest discover -s tests -v
"""

import json
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from server import reporting  # noqa: E402
from shared.json_serialization import from_json, to_json  # noqa: E402

# The view gives End as a datetime with its time.
SAMPLES = [
    {"ID PPSKT": "317057", "End": datetime(2026, 11, 6, 7, 46, 3)},
    {"ID PPSKT": "319894", "End": datetime(2026, 11, 6, 9, 1, 0)},
]


class LastReportSentTests(unittest.TestCase):
    def setUp(self):
        self._directory = tempfile.TemporaryDirectory()
        patcher = mock.patch.object(
            reporting, "LAST_REPORT_SENT_FILE", Path(self._directory.name) / "last_report_sent.json"
        )
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(self._directory.cleanup)

    def _create_report(self, samples):
        with mock.patch.object(reporting, "_fetch_samples_after", return_value=samples) as fetch:
            report = reporting.create_report()
        return report, fetch.call_args.args[0]

    def test_report_includes_its_last_sample_end(self):
        report, _ = self._create_report(SAMPLES)
        self.assertEqual(report["last_sample_end"], datetime(2026, 11, 6, 9, 1, 0))
        self.assertIn('"last_sample_end": "2026-11-06T09:01:00"', to_json(report))

    def test_saved_file_can_be_read_back(self):
        report, _ = self._create_report(SAMPLES)
        reporting.mark_report_sent(report)
        saved = json.loads(reporting.LAST_REPORT_SENT_FILE.read_text(encoding="utf-8"))
        self.assertEqual(saved["last_sample_date"], "2026-11-06T09:01:00")

        # The next report asks for samples after that End, time included.
        next_report, sent_up_to = self._create_report([])
        self.assertEqual(sent_up_to, datetime(2026, 11, 6, 9, 1, 0))
        self.assertEqual(next_report["samples_after"], datetime(2026, 11, 6, 9, 1, 0))

    def test_sending_twice_works_and_never_moves_backwards(self):
        reporting.mark_report_sent(self._create_report(SAMPLES)[0])
        reporting.mark_report_sent(self._create_report(SAMPLES[:1])[0])
        self.assertEqual(reporting._get_last_sample_end_sent(), datetime(2026, 11, 6, 9, 1, 0))

    def test_report_resent_as_json_is_still_saved(self):
        report = from_json(to_json(self._create_report(SAMPLES)[0]))
        reporting.mark_report_sent(report)
        self.assertEqual(reporting._get_last_sample_end_sent(), datetime(2026, 11, 6, 9, 1, 0))

    def test_old_file_with_only_a_date_is_read(self):
        reporting.LAST_REPORT_SENT_FILE.write_text(
            json.dumps({"last_sample_date": "2026-11-06", "sample_count": 3}), encoding="utf-8"
        )
        self.assertEqual(reporting._get_last_sample_end_sent(), datetime(2026, 11, 6))

    def test_empty_report_changes_nothing(self):
        reporting.mark_report_sent(self._create_report([])[0])
        self.assertFalse(reporting.LAST_REPORT_SENT_FILE.exists())

    def test_damaged_file_is_reported(self):
        reporting.LAST_REPORT_SENT_FILE.write_text("{ broken", encoding="utf-8")
        with self.assertRaises(reporting.LastReportFileError):
            reporting._get_last_sample_end_sent()


if __name__ == "__main__":
    unittest.main()
