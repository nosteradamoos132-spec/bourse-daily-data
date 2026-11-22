import importlib.util
import sys
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


if __name__ == "__main__":
    unittest.main()
