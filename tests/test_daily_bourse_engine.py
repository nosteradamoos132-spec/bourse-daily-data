import importlib.util
import json
import os
import sys
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ENGINE_PATH = ROOT / "scripts" / "daily_bourse_engine.py"
spec = importlib.util.spec_from_file_location("daily_bourse_engine", ENGINE_PATH)
engine = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = engine
spec.loader.exec_module(engine)


class TehranScheduleTests(unittest.TestCase):
    def dt(self, day, hour, minute=0):
        return datetime(
            day.year,
            day.month,
            day.day,
            hour,
            minute,
            tzinfo=engine.TEHRAN,
        )

    def test_jalali_date_conversion(self):
        self.assertEqual(
            engine.jalali_date_for(self.dt(date(2026, 11, 22), 18)),
            "1405-09-01",
        )

    def test_all_six_retry_windows(self):
        start = date(2026, 11, 22)
        tomorrow = date(2026, 11, 23)
        expected = [
            (self.dt(start, 18, 0), 0),
            (self.dt(start, 18, 30), 0),
            (self.dt(start, 19, 0), 0),
            (self.dt(start, 21, 0), 1),
            (self.dt(start, 21, 30), 1),
            (self.dt(start, 22, 0), 1),
            (self.dt(start, 23, 30), 2),
            (self.dt(tomorrow, 0, 0), 2),
            (self.dt(tomorrow, 0, 30), 2),
            (self.dt(tomorrow, 2, 0), 3),
            (self.dt(tomorrow, 2, 30), 3),
            (self.dt(tomorrow, 3, 0), 3),
            (self.dt(tomorrow, 4, 0), 4),
            (self.dt(tomorrow, 4, 30), 4),
            (self.dt(tomorrow, 5, 0), 4),
            (self.dt(tomorrow, 6, 0), 5),
            (self.dt(tomorrow, 6, 30), 5),
            (self.dt(tomorrow, 7, 0), 5),
        ]
        for moment, expected_stage in expected:
            with self.subTest(moment=moment.isoformat()):
                self.assertEqual(
                    engine.stage_for_now(moment, start),
                    expected_stage,
                )

    def test_cycle_closes_after_final_0600_stage(self):
        start = date(2026, 11, 22)
        tomorrow = date(2026, 11, 23)
        self.assertIsNone(engine.stage_for_now(self.dt(tomorrow, 7, 30), start))
        self.assertIsNone(engine.stage_for_now(self.dt(tomorrow, 8, 0), start))
        self.assertIsNone(
            engine.stage_for_now(self.dt(date(2026, 11, 24), 18), start)
        )

    def test_thursday_and_friday_cannot_start_new_cycle(self):
        thursday = self.dt(date(2026, 11, 26), 18)
        friday = self.dt(date(2026, 11, 27), 18)
        self.assertFalse(engine.is_start_day(thursday))
        self.assertFalse(engine.is_start_day(friday))

    def test_sunday_is_a_valid_nonholiday_start_day(self):
        sunday = self.dt(date(2026, 11, 22), 18)
        self.assertTrue(engine.is_start_day(sunday))

    def sample_report(self, report_date="1405-09-01"):
        return {
            "date": report_date,
            "time": "18:04:47",
            "state": "بسته",
            "index": 100000.0,
            "index_change": 100.0,
            "index_equalWeight": 50000.0,
            "index_equalWeight_change": 50.0,
            "mv": 1000.0,
            "tno": 2000.0,
            "tval": 3000.0,
            "tvol": 4000.0,
        }

    def test_previous_day_report_is_rejected_for_current_cycle(self):
        report = self.sample_report("1405-08-30")
        valid, reason = engine.valid_report(report, "1405-09-01")
        self.assertFalse(valid)
        self.assertIn("report date", reason)

    def test_identical_report_fingerprint_is_detected(self):
        first = self.sample_report()
        second = dict(reversed(list(first.items())))
        self.assertEqual(
            engine.report_fingerprint(first),
            engine.report_fingerprint(second),
        )
        second["tval"] += 1
        self.assertNotEqual(
            engine.report_fingerprint(first),
            engine.report_fingerprint(second),
        )

    def test_identical_duplicate_archive_files_are_ignored(self):
        report = self.sample_report()
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_dir = Path(temp_dir) / "archive"
            archive_dir.mkdir()
            (archive_dir / "index-1405-09-01.json").write_text(
                json.dumps(report, ensure_ascii=False), encoding="utf-8"
            )
            (archive_dir / "index-copy-1405-09-01.json").write_text(
                json.dumps(report, ensure_ascii=False), encoding="utf-8"
            )
            previous_dir = os.getcwd()
            try:
                os.chdir(temp_dir)
                records = engine.load_records()
            finally:
                os.chdir(previous_dir)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["date"], "1405-09-01")

    def test_conflicting_duplicate_archive_date_is_blocked(self):
        first = self.sample_report()
        second = self.sample_report()
        second["tval"] += 1
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_dir = Path(temp_dir) / "archive"
            archive_dir.mkdir()
            (archive_dir / "index-1405-09-01.json").write_text(
                json.dumps(first, ensure_ascii=False), encoding="utf-8"
            )
            (archive_dir / "index-copy-1405-09-01.json").write_text(
                json.dumps(second, ensure_ascii=False), encoding="utf-8"
            )
            previous_dir = os.getcwd()
            try:
                os.chdir(temp_dir)
                with self.assertRaisesRegex(RuntimeError, "CONFLICTING REPORTS"):
                    engine.load_records()
            finally:
                os.chdir(previous_dir)


if __name__ == "__main__":
    unittest.main()
