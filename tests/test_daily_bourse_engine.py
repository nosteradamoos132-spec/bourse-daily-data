import importlib.util
import json
import os
import sys
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch

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

    def test_manual_resume_allowed_after_midnight_for_previous_cycle(self):
        start = date(2026, 11, 22)
        tomorrow = date(2026, 11, 23)
        cycle = {
            "started_at": self.dt(start, 18).isoformat(),
            "status": "fetching",
        }
        self.assertTrue(engine.manual_resume_allowed(self.dt(tomorrow, 0), cycle))
        self.assertTrue(engine.manual_resume_allowed(self.dt(tomorrow, 7, 29), cycle))
        self.assertFalse(engine.manual_resume_allowed(self.dt(tomorrow, 7, 30), cycle))
        self.assertFalse(engine.manual_resume_allowed(self.dt(start, 23, 59), cycle))
        self.assertFalse(engine.manual_resume_allowed(self.dt(date(2026, 11, 24), 1), cycle))

    def test_manual_resume_allows_pending_excel_repair_until_deadline(self):
        start = date(2026, 11, 22)
        tomorrow = date(2026, 11, 23)
        cycle = {
            "started_at": self.dt(start, 18).isoformat(),
            "status": "pending_excel",
        }
        self.assertTrue(engine.manual_resume_allowed(self.dt(tomorrow, 1), cycle))
        self.assertFalse(engine.manual_resume_allowed(self.dt(tomorrow, 7, 30), cycle))

    def test_manual_and_automatic_execution_use_separate_state_files_and_quotas(self):
        original_path = engine.STATE_PATH
        original_limit = engine.MAX_DAILY_REQUESTS
        original_mode = engine.MANUAL_MODE
        try:
            engine.configure_execution_scope(True)
            self.assertEqual(engine.STATE_PATH, engine.MANUAL_STATE_PATH)
            self.assertEqual(engine.MAX_DAILY_REQUESTS, 15)
            self.assertTrue(engine.MANUAL_MODE)

            engine.configure_execution_scope(False)
            self.assertEqual(engine.STATE_PATH, engine.AUTO_STATE_PATH)
            self.assertEqual(engine.MAX_DAILY_REQUESTS, 85)
            self.assertFalse(engine.MANUAL_MODE)
        finally:
            engine.STATE_PATH = original_path
            engine.MAX_DAILY_REQUESTS = original_limit
            engine.MANUAL_MODE = original_mode

    def test_commit_trigger_file_is_read_as_one_shot_command(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            trigger_path = Path(temp_dir) / "manual-resume-trigger.json"
            trigger_path.write_text(
                json.dumps(
                    {"action": "resume_active_cycle", "request_id": "test-unique-001"}
                ),
                encoding="utf-8",
            )
            with patch.object(engine, "MANUAL_TRIGGER_PATH", trigger_path):
                trigger = engine.read_manual_trigger()
            self.assertEqual(trigger["action"], "resume_active_cycle")
            self.assertEqual(trigger["request_id"], "test-unique-001")

    def test_invalid_commit_trigger_file_is_rejected(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            trigger_path = Path(temp_dir) / "manual-resume-trigger.json"
            trigger_path.write_text(
                json.dumps({"action": "unrelated", "request_id": "test-unique-002"}),
                encoding="utf-8",
            )
            with patch.object(engine, "MANUAL_TRIGGER_PATH", trigger_path):
                with self.assertRaisesRegex(RuntimeError, "invalid structure"):
                    engine.read_manual_trigger()

    def test_manual_command_can_recover_yesterday_if_schedule_never_started(self):
        today = date(2026, 11, 23)
        now = self.dt(today, 0, 15)
        state = {
            "active_cycle": None,
            "last_cycle": None,
            "last_completed_at": None,
        }
        with patch.object(engine, "is_start_day", return_value=True):
            with patch.object(engine, "read_archived_report", return_value=None):
                with patch.object(engine, "persist_state") as persist:
                    cycle = engine.start_manual_previous_day_cycle(state, now)

        self.assertIsNotNone(cycle)
        self.assertEqual(cycle["market_date"], "1405-09-01")
        self.assertEqual(cycle["started_at"], self.dt(date(2026, 11, 22), 18).isoformat())
        self.assertEqual(cycle["status"], "fetching")
        persist.assert_called_once()

    def test_automatic_recovery_restores_previous_trading_day_cycle(self):
        now = self.dt(date(2026, 11, 23), 0, 15)
        state = {
            "active_cycle": None,
            "last_cycle": None,
            "last_completed_at": None,
        }
        with patch.object(engine, "is_start_day", return_value=True):
            with patch.object(engine, "read_archived_report", return_value=None):
                with patch.object(engine, "persist_state") as persist:
                    cycle = engine.recover_automatic_previous_day_cycle(state, now)

        self.assertIsNotNone(cycle)
        self.assertEqual(cycle["market_date"], "1405-09-01")
        self.assertEqual(cycle["started_at"], self.dt(date(2026, 11, 22), 18).isoformat())
        self.assertEqual(cycle["status"], "fetching")
        self.assertEqual(engine.stage_for_now(now, engine.cycle_start_date(cycle)), 2)
        persist.assert_called_once()

    def test_scheduled_main_uses_automatic_recovery_after_midnight(self):
        now = self.dt(date(2026, 11, 23), 0, 15)
        cycle = {
            "market_date": "1405-09-01",
            "started_at": self.dt(date(2026, 11, 22), 18).isoformat(),
            "status": "fetching",
            "stage_attempts": [0, 0, 0, 0, 0, 0],
            "last_request_at": None,
        }
        state = {
            "active_cycle": None,
            "last_cycle": None,
            "last_completed_at": None,
            "daily_usage": {},
        }

        def recover(state_arg, now_arg):
            state_arg["active_cycle"] = cycle
            return cycle

        with patch.dict(os.environ, {"GITHUB_EVENT_NAME": "schedule"}):
            with patch.object(engine, "tehran_now", return_value=now):
                with patch.object(engine, "configure_execution_scope"):
                    with patch.object(engine, "load_state", return_value=state):
                        with patch.object(
                            engine,
                            "recover_automatic_previous_day_cycle",
                            side_effect=recover,
                        ) as recovery:
                            with patch.object(engine, "stage_for_now", return_value=2):
                                with patch.object(engine, "attempt_api_once", return_value=False) as attempt:
                                    result = engine.main()

        self.assertEqual(result, 1)
        recovery.assert_called_once_with(state, now)
        attempt.assert_called_once_with(state, now, 2)

    def test_automatic_recovery_records_missed_cycle_after_0730(self):
        now = self.dt(date(2026, 11, 23), 7, 30)
        state = {"active_cycle": None, "last_cycle": None}
        with patch.object(engine, "is_start_day", return_value=True):
            with patch.object(engine, "persist_state") as persist:
                result = engine.recover_automatic_previous_day_cycle(state, now)

        self.assertIs(result, False)
        self.assertIsNone(state["active_cycle"])
        self.assertEqual(state["last_cycle"]["market_date"], "1405-09-01")
        self.assertEqual(state["last_cycle"]["status"], "failed")
        persist.assert_called_once()

    def test_automatic_recovery_skips_previous_nontrading_day(self):
        now = self.dt(date(2026, 11, 28), 0, 15)
        state = {"active_cycle": None, "last_cycle": None}
        with patch.object(engine, "is_start_day", return_value=False):
            with patch.object(engine, "persist_state") as persist:
                result = engine.recover_automatic_previous_day_cycle(state, now)

        self.assertIsNone(result)
        self.assertIsNone(state["active_cycle"])
        persist.assert_not_called()

    def test_manual_recovery_is_blocked_at_0730(self):
        now = self.dt(date(2026, 11, 23), 7, 30)
        state = {"active_cycle": None, "last_cycle": None}
        with patch.object(engine, "is_start_day") as is_start_day:
            cycle = engine.start_manual_previous_day_cycle(state, now)
        self.assertIsNone(cycle)
        self.assertIsNone(state["active_cycle"])
        is_start_day.assert_not_called()

    def test_manual_recovery_does_not_reopen_a_closed_previous_day(self):
        now = self.dt(date(2026, 11, 23), 1, 0)
        state = {
            "active_cycle": None,
            "last_cycle": {"market_date": "1405-09-01", "status": "duplicate"},
        }
        with patch.object(engine, "is_start_day") as is_start_day:
            cycle = engine.start_manual_previous_day_cycle(state, now)
        self.assertIsNone(cycle)
        self.assertIsNone(state["active_cycle"])
        is_start_day.assert_not_called()

    def test_manual_resume_rejects_non_fetching_cycle(self):
        cycle = {
            "started_at": self.dt(date(2026, 11, 22), 18).isoformat(),
            "status": "duplicate",
        }
        self.assertFalse(
            engine.manual_resume_allowed(self.dt(date(2026, 11, 23), 1), cycle)
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
            duplicate_copy = dict(report)
            duplicate_copy["time"] = "18:19:02"
            (archive_dir / "index-1405-09-01.json").write_text(
                json.dumps(duplicate_copy, ensure_ascii=False), encoding="utf-8"
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


    def test_archived_date_duplicate_is_detected(self):
        report = self.sample_report("1405-08-30")
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_dir = Path(temp_dir)
            (archive_dir / "index-1405-08-30.json").write_text(
                json.dumps(report, ensure_ascii=False), encoding="utf-8"
            )
            with patch.object(engine, "ARCHIVE_DIR", archive_dir):
                reason = engine.find_duplicate_archived_report(report)

        self.assertIsNotNone(reason)
        self.assertIn("already exists", reason)

    def test_identical_market_values_on_another_date_are_not_duplicate(self):
        archived = self.sample_report("1405-08-30")
        incoming = self.sample_report("1405-09-01")
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_dir = Path(temp_dir)
            (archive_dir / "index-1405-08-30.json").write_text(
                json.dumps(archived, ensure_ascii=False), encoding="utf-8"
            )
            with patch.object(engine, "ARCHIVE_DIR", archive_dir):
                reason = engine.find_duplicate_archived_report(incoming)

        self.assertIsNone(reason)

    def test_new_market_values_are_not_duplicate(self):
        archived = self.sample_report("1405-08-30")
        incoming = self.sample_report("1405-09-01")
        incoming["index"] += 1
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_dir = Path(temp_dir)
            (archive_dir / "index-1405-08-30.json").write_text(
                json.dumps(archived, ensure_ascii=False), encoding="utf-8"
            )
            with patch.object(engine, "ARCHIVE_DIR", archive_dir):
                reason = engine.find_duplicate_archived_report(incoming)

        self.assertIsNone(reason)


    def test_duplicate_cycle_cannot_restart_later_same_day(self):
        now = self.dt(date(2026, 11, 22), 21, 0)
        state = {
            "active_cycle": None,
            "last_cycle": {
                "market_date": "1405-09-01",
                "status": "duplicate",
            },
            "last_completed_at": None,
        }
        with patch.object(engine, "persist_state") as persist:
            started = engine.start_new_cycle(state, now)

        self.assertFalse(started)
        self.assertIsNone(state["active_cycle"])
        persist.assert_not_called()

    def test_next_workday_cycle_is_not_blocked_by_rolling_24_hour_lock(self):
        next_workday = self.dt(date(2026, 11, 23), 18, 0)
        state = {
            "active_cycle": None,
            "last_cycle": {
                "market_date": "1405-09-01",
                "status": "completed",
            },
            "last_completed_at": self.dt(date(2026, 11, 22), 19, 0).isoformat(),
        }
        with patch.object(engine, "persist_state"):
            with patch.object(engine, "read_archived_report", return_value=None):
                with patch.object(engine, "stage_for_now", return_value=None):
                    started = engine.start_new_cycle(state, next_workday)

        self.assertTrue(started)
        self.assertEqual(state["active_cycle"]["market_date"], "1405-09-02")

    def test_failed_api_attempt_returns_failure_status(self):
        now = self.dt(date(2026, 11, 22), 18, 0)
        cycle = {
            "market_date": "1405-09-01",
            "started_at": now.isoformat(),
            "status": "fetching",
            "stage_attempts": [0, 0, 0, 0, 0, 0],
            "last_request_at": None,
        }
        state = {
            "schema_version": 1,
            "active_cycle": cycle,
            "last_completed_at": None,
            "last_cycle": None,
            "daily_usage": {},
        }
        with patch.dict(os.environ, {"BRSAPI_KEY": "test-key"}):
            with patch.object(engine, "persist_state"):
                with patch.object(engine, "tehran_now", return_value=now):
                    with patch.object(
                        engine,
                        "request_brsapi",
                        return_value=(None, None, 503, "temporary API failure"),
                    ):
                        result = engine.attempt_api_once(state, now, 0)

        self.assertIs(result, False)
        self.assertEqual(cycle["stage_attempts"][0], 1)
        self.assertEqual(cycle["last_attempt"]["result"], "failed_or_invalid")

    def test_duplicate_response_closes_cycle_without_registration(self):
        now = self.dt(date(2026, 11, 22), 18, 0)
        cycle = {
            "market_date": "1405-09-01",
            "started_at": now.isoformat(),
            "status": "fetching",
            "stage_attempts": [0, 0, 0, 0, 0, 0],
            "last_request_at": None,
        }
        state = {
            "schema_version": 1,
            "active_cycle": cycle,
            "last_completed_at": None,
            "last_cycle": None,
            "daily_usage": {},
        }
        with patch.dict(os.environ, {"BRSAPI_KEY": "test-key"}):
            with patch.object(engine, "persist_state") as persist:
                with patch.object(
                    engine,
                    "request_brsapi",
                    return_value=(None, None, 200, "DUPLICATE: old report already archived"),
                ):
                    engine.attempt_api_once(state, now, 0)

        self.assertIsNone(state["active_cycle"])
        self.assertEqual(state["last_cycle"]["status"], "duplicate")
        self.assertFalse(state["last_cycle"]["new_report_registered"])
        self.assertEqual(cycle["stage_attempts"][0], 1)
        self.assertEqual(persist.call_count, 2)

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



class ManualCurrentDayStartTests(unittest.TestCase):
    def dt(self, day, hour, minute=0):
        return datetime(
            day.year,
            day.month,
            day.day,
            hour,
            minute,
            tzinfo=engine.TEHRAN,
        )

    def test_manual_current_day_cycle_starts_after_market_close(self):
        today = date(2026, 11, 23)
        now = self.dt(today, 18, 19)
        state = {
            "active_cycle": None,
            "last_cycle": None,
            "last_completed_at": None,
        }
        with patch.object(engine, "is_start_day", return_value=True):
            with patch.object(engine, "read_archived_report", return_value=None):
                with patch.object(engine, "persist_state") as persist:
                    cycle = engine.start_manual_current_day_cycle(state, now)

        self.assertIsNotNone(cycle)
        self.assertEqual(cycle["market_date"], "1405-09-02")
        self.assertEqual(cycle["started_at"], self.dt(today, 18).isoformat())
        self.assertEqual(cycle["status"], "fetching")
        self.assertTrue(cycle["manual_current_day"])
        persist.assert_called_once()

    def test_manual_current_day_cycle_is_blocked_before_1800(self):
        today = date(2026, 11, 23)
        state = {"active_cycle": None, "last_cycle": None}
        with patch.object(engine, "is_start_day") as is_start_day:
            cycle = engine.start_manual_current_day_cycle(state, self.dt(today, 17, 59))
        self.assertIsNone(cycle)
        self.assertIsNone(state["active_cycle"])
        is_start_day.assert_not_called()

    def test_manual_current_day_cycle_can_reopen_closed_date(self):
        today = date(2026, 11, 23)
        state = {
            "active_cycle": None,
            "last_cycle": {"market_date": "1405-09-02", "status": "completed"},
        }
        with patch.object(engine, "is_start_day", return_value=True):
            with patch.object(engine, "read_archived_report", return_value=self.sample_report()):
                with patch.object(engine, "persist_state"):
                    cycle = engine.start_manual_current_day_cycle(state, self.dt(today, 18, 19))
        self.assertIsNotNone(cycle)
        self.assertEqual(cycle["status"], "fetching")
        self.assertTrue(cycle["manual_current_day"])

    
    def test_manual_current_day_existing_report_does_not_skip_api_cycle(self):
        today = date(2026, 11, 23)
        state = {"active_cycle": None, "last_cycle": None}
        existing = self.sample_report()
        with patch.object(engine, "is_start_day", return_value=True):
            with patch.object(engine, "read_archived_report", return_value=existing):
                with patch.object(engine, "persist_state"):
                    cycle = engine.start_manual_current_day_cycle(state, self.dt(today, 18, 19))
        self.assertEqual(cycle["status"], "fetching")
        self.assertIsNone(cycle["report_path"])

    def test_commit_trigger_accepts_immediate_current_day_action(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            trigger_path = Path(temp_dir) / "manual-resume-trigger.json"
            trigger_path.write_text(
                json.dumps(
                    {"action": "start_current_cycle", "request_id": "manual-now-001"}
                ),
                encoding="utf-8",
            )
            with patch.object(engine, "MANUAL_TRIGGER_PATH", trigger_path):
                trigger = engine.read_manual_trigger()
        self.assertEqual(trigger["action"], "start_current_cycle")
        self.assertEqual(trigger["request_id"], "manual-now-001")

class ManualOverwriteTests(unittest.TestCase):
    def dt(self):
        return datetime(2026, 11, 23, 18, 30, tzinfo=engine.TEHRAN)

    def report(self, value):
        return {
            "date": "1405-09-02",
            "time": "18:30:00",
            "state": "بسته",
            "index": 100000.0 + value,
            "index_change": 100.0,
            "index_equalWeight": 50000.0 + value,
            "index_equalWeight_change": 50.0,
            "mv": 1000.0,
            "tno": 2000.0,
            "tval": 3000.0 + value,
            "tvol": 4000.0 + value,
        }

    def make_state(self):
        now = self.dt()
        return {
            "schema_version": 1,
            "active_cycle": {
                "market_date": "1405-09-02",
                "started_at": now.isoformat(),
                "status": "fetching",
                "stage_attempts": [0, 0, 0, 0, 0, 0],
                "last_request_at": None,
            },
            "last_completed_at": None,
            "last_cycle": None,
            "daily_usage": {},
        }

    def test_automatic_execution_preserves_existing_json(self):
        old = self.report(1)
        new = self.report(2)
        with tempfile.TemporaryDirectory() as temp_dir:
            archive = Path(temp_dir)
            path = archive / "index-1405-09-02.json"
            path.write_text(json.dumps(old, ensure_ascii=False), encoding="utf-8")
            state = self.make_state()
            with patch.object(engine, "ARCHIVE_DIR", archive):
                with patch.object(engine, "MANUAL_MODE", False):
                    with patch.dict(os.environ, {"BRSAPI_KEY": "test-key"}):
                        with patch.object(engine, "persist_state"):
                            with patch.object(engine, "request_brsapi", return_value=(new, b"{}", 200, "valid")):
                                with patch.object(engine, "commit_and_push"):
                                    with patch.object(engine, "process_pending_excel", return_value=True):
                                        engine.attempt_api_once(state, self.dt(), 0)
            saved = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(saved["tval"], old["tval"])

    def test_manual_execution_replaces_existing_json_with_new_api_report(self):
        old = self.report(1)
        new = self.report(2)
        with tempfile.TemporaryDirectory() as temp_dir:
            archive = Path(temp_dir)
            path = archive / "index-1405-09-02.json"
            path.write_text(json.dumps(old, ensure_ascii=False), encoding="utf-8")
            state = self.make_state()
            with patch.object(engine, "ARCHIVE_DIR", archive):
                with patch.object(engine, "MANUAL_MODE", True):
                    with patch.dict(os.environ, {"BRSAPI_KEY": "test-key"}):
                        with patch.object(engine, "persist_state"):
                            with patch.object(engine, "request_brsapi", return_value=(new, b"{}", 200, "valid")):
                                with patch.object(engine, "commit_and_push") as commit:
                                    with patch.object(engine, "process_pending_excel", return_value=True) as excel:
                                        engine.attempt_api_once(state, self.dt(), 0)
            saved = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(saved["tval"], new["tval"])
            self.assertEqual(saved["tvol"], new["tvol"])
            commit.assert_called_once()
            excel.assert_called_once_with(state)

