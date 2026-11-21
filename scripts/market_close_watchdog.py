#!/usr/bin/env python3
import importlib.util
import json
import math
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

API = "https://api.brsapi.ir/Tsetmc/Index.php"
POLL_SECONDS = 300
CHAIN_AFTER_SECONDS = 5 * 60 * 60 + 35 * 60
WORKFLOW_FILE = "check-market-close.yml"
REPO = os.environ["GITHUB_REPOSITORY"]
TOKEN = os.environ["GH_TOKEN"]
API_KEY = os.environ.get("BRSAPI_KEY", "")

def gregorian_to_jalali(gy, gm, gd):
    month_days = [0,31,59,90,120,151,181,212,243,273,304,334]
    if gy > 1600:
        jy, gy = 979, gy - 1600
    else:
        jy, gy = 0, gy - 621
    gy2 = gy + 1 if gm > 2 else gy
    days = 365*gy + (gy2+3)//4 - (gy2+99)//100 + (gy2+399)//400 - 80 + gd + month_days[gm-1]
    jy += 33*(days//12053)
    days %= 12053
    jy += 4*(days//1461)
    days %= 1461
    if days > 365:
        jy += (days-1)//365
        days = (days-1)%365
    if days < 186:
        jm, jd = 1+days//31, 1+days%31
    else:
        jm, jd = 7+(days-186)//30, 1+(days-186)%30
    return jy, jm, jd

def today_jalali():
    now = datetime.now(ZoneInfo("Asia/Tehran"))
    y, m, d = gregorian_to_jalali(now.year, now.month, now.day)
    return now, f"{y:04d}-{m:02d}-{d:02d}", y

def friday_or_holiday(now, jalali_date, jalali_year):
    if now.weekday() == 4:
        print("SKIP: Friday in Iran; no report or workbook checks today.")
        return True
    calendar_path = Path(f"calendar/holidays-{jalali_year}.json")
    if not calendar_path.is_file():
        raise RuntimeError(f"Missing Iranian holiday calendar: {calendar_path}")
    calendar = json.loads(calendar_path.read_text(encoding="utf-8"))
    if calendar.get("year") != jalali_year:
        raise RuntimeError("Iranian holiday calendar year mismatch.")
    if jalali_date in set(calendar.get("holidays", [])):
        print(f"SKIP: {jalali_date} is an official Iranian public holiday.")
        return True
    return False

def valid_report(data, expected_date):
    if not isinstance(data, dict):
        return False
    if data.get("date") != expected_date or data.get("state") != "بسته":
        return False
    report_time = data.get("time")
    if not isinstance(report_time, str) or not re.fullmatch(r"\d{2}:\d{2}:\d{2}", report_time) or report_time < "17:00:00":
        return False
    required = (
        "index", "index_change", "index_equalWeight",
        "index_equalWeight_change", "mv", "tno", "tval", "tvol",
    )
    for key in required:
        value = data.get(key)
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
            return False
    return all(float(data[key]) > 0 for key in ("mv", "tno", "tval", "tvol"))

def read_archived_report(expected_date):
    path = Path(f"archive/index-{expected_date}.json")
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if valid_report(data, expected_date) else None
    except Exception as exc:
        print(f"Existing JSON archive is invalid: {exc}")
        return None

def git(*args, check=True):
    return subprocess.run(["git", *args], check=check, text=True)

def commit_and_push(paths, message):
    git("add", "--", *paths)
    changed = subprocess.run(
        ["git", "diff", "--cached", "--quiet"],
        check=False,
    ).returncode != 0
    if changed:
        git("config", "user.name", "github-actions[bot]")
        git("config", "user.email", "41898282+github-actions[bot]@users.noreply.github.com")
        git("commit", "-m", message)
    git("pull", "--rebase", "origin", "main")
    git("push", "origin", "HEAD:main")

def fetch_report(expected_date):
    if not API_KEY:
        raise RuntimeError("Missing GitHub secret BRSAPI_KEY.")
    query = urlencode({"key": API_KEY, "type": "1"})
    request = Request(
        f"{API}?{query}",
        headers={"Accept": "application/json", "User-Agent": "bourse-market-close-watchdog/1.0"},
    )
    with urlopen(request, timeout=45) as response:
        if response.status != 200:
            raise RuntimeError(f"BRSAPI HTTP status {response.status}")
        data = json.loads(response.read().decode("utf-8-sig"))
    print(
        f"BRSAPI response: date={data.get('date')!r}, "
        f"time={data.get('time')!r}, state={data.get('state')!r}"
    )
    return data if valid_report(data, expected_date) else None

def chain_next_run():
    url = f"https://api.github.com/repos/{REPO}/actions/workflows/{WORKFLOW_FILE}/dispatches"
    body = json.dumps({"ref": "main"}).encode("utf-8")
    request = Request(
        url,
        data=body,
        method="POST",
        headers={
            "Accept": "application/vnd.github+json",
            "Authorization": f"Bearer {TOKEN}",
            "X-GitHub-Api-Version": "2022-11-28",
            "Content-Type": "application/json",
            "User-Agent": "bourse-market-close-watchdog",
        },
    )
    with urlopen(request, timeout=30) as response:
        if response.status != 204:
            raise RuntimeError(f"Could not schedule continuation; HTTP {response.status}")
    print("Continuation queued: the watchdog will resume checking in a new run.")

def verify_workbook(expected_date, report):
    spec = importlib.util.spec_from_file_location(
        "update_index_excels", "scripts/update_index_excels.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    book = module.Book("archive/داده‌های روزانه بازار.xlsx", "داده‌های روزانه بازار")
    rows = book.rows()
    matching = [
        (rn, cells) for rn, cells in rows.items()
        if rn > 1 and module.canonical_date(book.value(cells.get(1))) == expected_date
    ]
    if len(matching) != 1:
        return False
    row_number, cells = matching[0]
    if row_number != 2:
        print(f"Workbook newest row is {row_number}; expected today's date in row 2.")
        return False
    expected = {
        2: float(report["index"]),
        3: float(report["index_equalWeight"]),
        4: float(report["tval"]) / 1_000_000_000,
        5: float(report["tvol"]) / 1_000_000_000,
    }
    for column, value in expected.items():
        actual = book.value(cells.get(column))
        if actual is None or not math.isclose(float(actual), value, rel_tol=1e-12, abs_tol=1e-9):
            print(f"Workbook verification failed in column {column}: {actual!r} != {value!r}")
            return False
    if book.number_format(cells.get(4)) != "#,##0.000":
        print("Workbook value display format verification failed.")
        return False
    if book.number_format(cells.get(5)) != "0.000":
        print("Workbook volume display format verification failed.")
        return False
    print(f"VERIFIED: today's report and all four market values are in workbook row 2 ({expected_date}).")
    return True

def main():
    now, expected_date, jy = today_jalali()
    print(f"Tehran date: {expected_date}; Tehran time: {now.isoformat()}")
    if friday_or_holiday(now, expected_date, jy):
        return 0

    started = time.monotonic()
    report = read_archived_report(expected_date)
    if report:
        print(f"REPORT OK: archive/index-{expected_date}.json already exists and is valid.")
    else:
        print("REPORT MISSING/INVALID: polling BRSAPI every five minutes until a valid final report is received.")

    while not report:
        if time.monotonic() - started >= CHAIN_AFTER_SECONDS:
            chain_next_run()
            return 0
        try:
            report = fetch_report(expected_date)
            if report:
                path = Path(f"archive/index-{expected_date}.json")
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
                commit_and_push([str(path)], f"Archive BRSAPI report {expected_date} (20:00 watchdog)")
                print(f"REPORT SAVED AND PUSHED: {path}")
            else:
                print("No valid final report yet; will retry.")
        except (HTTPError, URLError, TimeoutError, ValueError, json.JSONDecodeError, RuntimeError) as exc:
            print(f"Report attempt failed: {type(exc).__name__}: {exc}")
        except Exception as exc:
            print(f"Unexpected report attempt error: {type(exc).__name__}: {exc}")
        if not report:
            time.sleep(POLL_SECONDS)

    while True:
        if time.monotonic() - started >= CHAIN_AFTER_SECONDS:
            chain_next_run()
            return 0
        try:
            git("fetch", "origin", "main")
            git("reset", "--hard", "origin/main")
            subprocess.run([sys.executable, "scripts/update_index_excels.py"], check=True)
            subprocess.run([sys.executable, "scripts/merge_manual_history.py"], check=True)
            if verify_workbook(expected_date, report):
                commit_and_push(
                    ["archive/داده‌های روزانه بازار.xlsx"],
                    f"Verify daily market workbook {expected_date} (20:00 watchdog)",
                )
                git("fetch", "origin", "main")
                git("reset", "--hard", "origin/main")
                if verify_workbook(expected_date, report):
                    print("SUCCESS: report archive and unified Excel workbook are both verified on main.")
                    return 0
                print("Workbook verification after push failed; retrying.")
            else:
                print("Workbook is not yet correct; retrying after five minutes.")
        except Exception as exc:
            print(f"Workbook update/verification failed: {type(exc).__name__}: {exc}")
        time.sleep(POLL_SECONDS)

if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"FATAL: {type(exc).__name__}: {exc}", file=sys.stderr)
        raise
