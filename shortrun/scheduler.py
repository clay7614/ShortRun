from __future__ import annotations
import datetime as dt
import csv
import os
import re
import subprocess
from typing import List, Dict, Optional, Tuple
import concurrent.futures
import tempfile

_TASK_PREFIX = "ShortRun_"
_alias_re = re.compile(r"[^A-Za-z0-9_-]+")

def _sanitize(s: str) -> str:
    s = s.strip()
    s = _alias_re.sub("_", s)
    return s[:60]


def _task_name(alias: str, kind: str, suffix: Optional[str] = None) -> str:
    name = f"{_TASK_PREFIX}{_sanitize(alias)}_{kind}"
    if suffix:
        name += f"_{_sanitize(suffix)}"
    return name


def _quote(path: str) -> str:
    # schtasks の /TR に渡す文字列は二重引用符で囲む
    p = path.strip().strip('"')
    return f'"{p}"'


def _run(cmd: List[str]) -> subprocess.CompletedProcess:
    """サブプロセス実行（Windowsではコンソールを出さない）。"""
    kwargs = dict(capture_output=True, text=True, shell=False)
    if os.name == 'nt':
        # コンソールの点滅防止
        try:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= getattr(subprocess, 'STARTF_USESHOWWINDOW', 0)
            si.wShowWindow = 0  # SW_HIDE
        except Exception:
            si = None
        creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        kwargs.update(startupinfo=si, creationflags=creationflags)
    return subprocess.run(cmd, **kwargs)


def _ensure_author(task_name: str, author: str = "ShortRun") -> None:
    """Ensure the task has the given Author in its XML. Best-effort.
    Reads the XML, patches <RegistrationInfo><Author>, recreates the task.
    """
    try:
        cp = _run(["schtasks", "/Query", "/TN", task_name, "/XML"])
        if cp.returncode != 0 or not cp.stdout:
            return
        xml = cp.stdout
        # Insert or replace Author
        if "<Author>" in xml:
            new_xml = re.sub(r"<Author>.*?</Author>", f"<Author>{author}</Author>", xml, flags=re.S)
        else:
            # Try to inject under <RegistrationInfo>
            new_xml = re.sub(r"<RegistrationInfo>\s*", f"<RegistrationInfo><Author>{author}</Author>", xml, count=1)
        if not new_xml:
            return
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml", mode="w", encoding="utf-8") as f:
            f.write(new_xml)
            tmp = f.name
        try:
            _run(["schtasks", "/Create", "/TN", task_name, "/XML", tmp, "/F"])
        finally:
            try:
                os.remove(tmp)
            except Exception:
                pass
    except Exception:
        # Best-effort; ignore
        pass


def _get_author_and_enabled(task_name: str) -> Tuple[Optional[str], Optional[bool]]:
    """Return (Author, Enabled) by querying the task XML."""
    try:
        cp = _run(["schtasks", "/Query", "/TN", task_name, "/XML"])
        if cp.returncode != 0 or not cp.stdout:
            return (None, None)
        xml = cp.stdout
        m_a = re.search(r"<Author>(.*?)</Author>", xml, flags=re.S)
        author = m_a.group(1).strip() if m_a else None
        m_e = re.search(r"<Enabled>(true|false)</Enabled>", xml, flags=re.I)
        enabled = None
        if m_e:
            enabled = m_e.group(1).lower() == "true"
        return (author, enabled)
    except Exception:
        return (None, None)


def _append_schedule_window(cmd: List[str], *, sd: Optional[str] = None, ed: Optional[str] = None, et: Optional[str] = None, du: Optional[str] = None) -> None:
    """任意の開始日/終了日/終了時刻/期間をコマンドへ追加する。
    - sd: YYYY/MM/DD or YYYY-MM-DD
    - ed: YYYY/MM/DD or YYYY-MM-DD
    - et: HH:MM （24h）
    - du: HHH:MM or HH:MM （schtasks仕様に依存。バリデーションは緩め）
    """
    if sd:
        s = sd.replace('-', '/')
        # 最低限の形式チェック
        try:
            dt.datetime.strptime(s, "%Y/%m/%d")
            cmd.extend(["/SD", s])
        except ValueError:
            pass
    if ed:
        e = ed.replace('-', '/')
        try:
            dt.datetime.strptime(e, "%Y/%m/%d")
            cmd.extend(["/ED", e])
        except ValueError:
            pass
    if et:
        # HH:MM を想定、緩めに受け入れ
        try:
            dt.datetime.strptime(et, "%H:%M")
            cmd.extend(["/ET", et])
        except ValueError:
            pass
    if du:
        # 形式は環境依存のため簡易受け入れのみ
        if re.match(r"^\d{1,4}:\d{2}(:\d{2})?$", du):
            cmd.extend(["/DU", du])


def _append_repetition(cmd: List[str], *, rep_interval_min: Optional[int] = None, rep_duration: Optional[str] = None) -> None:
    """任意の繰り返し実行（Repetition）をコマンドへ追加する。
    - rep_interval_min: 繰り返し間隔(分)
    - rep_duration: 継続時間 HHH:MM 形式（schtasks の /DU に準拠）
    注意: /RI を指定する場合、/DU も必要（Task Scheduler の制約）。
    """
    if rep_interval_min is not None:
        try:
            ival = int(rep_interval_min)
            if ival >= 1:
                cmd.extend(["/RI", str(ival)])
        except Exception:
            pass
    if rep_duration:
        if re.match(r"^\d{1,4}:\d{2}(:\d{2})?$", rep_duration):
            cmd.extend(["/DU", rep_duration])


def _set_calendar_random_delay_minutes(task_name: str, minutes: int) -> None:
    """Patch task XML to set <RandomDelay> for Calendar(Time) triggers. Best-effort.
    minutes: total minutes for random delay. Writes PT{minutes}M.
    """
    try:
        if minutes is None or minutes <= 0:
            return
        cp = _run(["schtasks", "/Query", "/TN", task_name, "/XML"])
        if cp.returncode != 0 or not cp.stdout:
            return
        xml = cp.stdout
        # Find CalendarTrigger block
        pat = re.compile(r"(<CalendarTrigger[^>]*>)(.*?)(</CalendarTrigger>)", flags=re.S)
        m = pat.search(xml)
        if not m:
            # Some schedules might use <TimeTrigger>
            pat2 = re.compile(r"(<TimeTrigger[^>]*>)(.*?)(</TimeTrigger>)", flags=re.S)
            m = pat2.search(xml)
            if not m:
                return
        head, body, tail = m.group(1), m.group(2), m.group(3)
        # Replace or insert <RandomDelay>
        if re.search(r"<RandomDelay>.*?</RandomDelay>", body, flags=re.S):
            body2 = re.sub(r"<RandomDelay>.*?</RandomDelay>", f"<RandomDelay>PT{minutes}M</RandomDelay>", body)
        else:
            # Insert before closing tag, try after <StartBoundary> if exists
            if "</StartBoundary>" in body:
                body2 = body.replace("</StartBoundary>", "</StartBoundary><RandomDelay>PT" + str(minutes) + "M</RandomDelay>")
            else:
                body2 = body + f"<RandomDelay>PT{minutes}M</RandomDelay>"
        new_xml = xml[:m.start()] + head + body2 + tail + xml[m.end():]
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml", mode="w", encoding="utf-8") as f:
            f.write(new_xml)
            tmp = f.name
        try:
            _run(["schtasks", "/Create", "/TN", task_name, "/XML", tmp, "/F"])
        finally:
            try:
                os.remove(tmp)
            except Exception:
                pass
    except Exception:
        pass


def _set_repetition_stop_at_end(task_name: str, stop: bool) -> None:
    """Patch task XML to set <Repetition><StopAtDurationEnd> flag. Best-effort."""
    try:
        cp = _run(["schtasks", "/Query", "/TN", task_name, "/XML"])
        if cp.returncode != 0 or not cp.stdout:
            return
        xml = cp.stdout
        pat = re.compile(r"(<Repetition[^>]*>)(.*?)(</Repetition>)", flags=re.S)
        m = pat.search(xml)
        if not m:
            return
        head, body, tail = m.group(1), m.group(2), m.group(3)
        if re.search(r"<StopAtDurationEnd>.*?</StopAtDurationEnd>", body, flags=re.S):
            body2 = re.sub(r"<StopAtDurationEnd>.*?</StopAtDurationEnd>", f"<StopAtDurationEnd>{str(bool(stop)).lower()}</StopAtDurationEnd>", body)
        else:
            body2 = body + f"<StopAtDurationEnd>{str(bool(stop)).lower()}</StopAtDurationEnd>"
        new_xml = xml[:m.start()] + head + body2 + tail + xml[m.end():]
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml", mode="w", encoding="utf-8") as f:
            f.write(new_xml)
            tmp = f.name
        try:
            _run(["schtasks", "/Create", "/TN", task_name, "/XML", tmp, "/F"])
        finally:
            try:
                os.remove(tmp)
            except Exception:
                pass
    except Exception:
        pass


def list_tasks(alias: Optional[str] = None, *, author: Optional[str] = None) -> List[Dict[str, str]]:
    """タスク一覧を取得。
    - デフォルトは ShortRun_ プレフィックスのタスクのみを対象
    - author を指定した場合は、作成者（Author）が一致するタスクを対象にする
    いずれも簡易3列（TaskName, Next Run Time, Status）をCSVで取得し、
    author 指定時は XML から Author/Enabled を補完する。
    """
    cp = _run(["schtasks", "/Query", "/FO", "CSV", "/NH"])
    if cp.returncode != 0:
        return []
    results: List[Dict[str, str]] = []
    rows: List[Tuple[str, str, str]] = []
    for line in cp.stdout.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            parts = next(csv.reader([line]))
        except Exception:
            continue
        if not parts:
            continue
        name = (parts[0] or "").lstrip('\\')
        next_run = parts[1] if len(parts) > 1 else ''
        status = parts[2] if len(parts) > 2 else ''
        # 事前に配列へ溜める（author 指定時の並列処理のため）
        rows.append((name, next_run, status))

    if author is None:
        for name, next_run, status in rows:
            # 既定: 名前で ShortRun_ のみ
            if not name.startswith(_TASK_PREFIX):
                continue
            results.append({
                'TaskName': name,
                'SimpleName': name,
                'NextRunTime': next_run,
                'Status': status,
                'Author': None,
                'Enabled': None,
                'Schedule': '',
            })
    else:
        # Author 指定時: Microsoft 配下は除外し、XML 確認を並列化
        target_author = author.strip().lower()
        candidates = [(n, nr, st) for (n, nr, st) in rows if not n.lower().startswith("microsoft\\")]

        def _probe(tpl: Tuple[str, str, str]) -> Optional[Dict[str, str]]:
            n, nr, st = tpl
            a, en = _get_author_and_enabled(n)
            if (a or '').strip().lower() != target_author:
                return None
            return {
                'TaskName': n,
                'SimpleName': n,
                'NextRunTime': nr,
                'Status': st,
                'Author': a,
                'Enabled': en,
                'Schedule': '',
            }

        max_workers = max(2, min(8, (os.cpu_count() or 4)))
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
            for res in ex.map(_probe, candidates):
                if res:
                    results.append(res)
    # 追加の alias フィルタ（指定時）
    if alias is not None:
        prefix = _TASK_PREFIX + _sanitize(alias)
        results = [t for t in results if t['SimpleName'].startswith(prefix)]
    return results


def delete_task_by_simple_name(simple_name: str) -> None:
    _run(["schtasks", "/Delete", "/TN", simple_name, "/F"])


def delete_all_for_alias(alias: str) -> None:
    for t in list_tasks(alias):
        delete_task_by_simple_name(t['SimpleName'])


def ensure_logon_task(alias: str, exe_path: str, enabled: bool, *, elevated: bool = False, task_name: Optional[str] = None, delay_minutes: Optional[int] = None) -> None:
    name = _task_name(alias, "LOGON")
    if task_name:
        name = task_name
    if enabled:
        # 既存は上書き
        _run(["schtasks", "/Delete", "/TN", name, "/F"])
        cmd = [
            "schtasks", "/Create",
            "/TN", name,
            "/SC", "ONLOGON",
            "/TR", _quote(exe_path),
            "/RL", ("HIGHEST" if elevated else "LIMITED"),
            "/F",
        ]
        cp = _run(cmd)
        if cp.returncode != 0:
            raise RuntimeError(cp.stderr or cp.stdout)
        _ensure_author(name)
        # 任意: 起動遅延
        if isinstance(delay_minutes, int) and delay_minutes >= 0:
            try:
                _set_trigger_delay_minutes(name, delay_minutes, kind="LOGON")
            except Exception:
                pass
    else:
        _run(["schtasks", "/Delete", "/TN", name, "/F"])


def create_daily_task(alias: str, exe_path: str, hhmm: str, *, sd: Optional[str] = None, ed: Optional[str] = None, et: Optional[str] = None, du: Optional[str] = None, rep_interval_min: Optional[int] = None, rep_duration: Optional[str] = None, stop_at_rep_end: bool = False, random_delay_minutes: Optional[int] = None, utc: bool = False, elevated: bool = False, task_name: Optional[str] = None) -> None:
    # hh:mm を想定
    try:
        dt.datetime.strptime(hhmm, "%H:%M")
    except ValueError:
        raise ValueError("時刻は HH:MM 形式で指定してください")
    name = task_name or _task_name(alias, "DAILY", hhmm.replace(":", "-"))
    # 既存は上書き
    _run(["schtasks", "/Delete", "/TN", name, "/F"])
    cmd = [
        "schtasks", "/Create",
        "/TN", name,
        "/SC", "DAILY",
        "/ST", hhmm,
        "/TR", _quote(exe_path),
        "/RL", ("HIGHEST" if elevated else "LIMITED"),
        "/F",
    ]
    if utc:
        cmd.append("/Z")
    # 繰り返しを指定した場合は /DU は繰り返しの継続時間に使用するため、
    # ウィンドウ用の du は同時に指定しない（競合回避）。
    if rep_interval_min is None:
        _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=du)
    else:
        _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=None)
        _append_repetition(cmd, rep_interval_min=rep_interval_min, rep_duration=rep_duration)
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)
    # Post-create patches
    if random_delay_minutes:
        _set_calendar_random_delay_minutes(name, int(random_delay_minutes))
    if stop_at_rep_end and rep_interval_min is not None:
        _set_repetition_stop_at_end(name, True)


def create_once_task(alias: str, exe_path: str, date_str: str, hhmm: str, *, random_delay_minutes: Optional[int] = None, utc: bool = False, elevated: bool = False, task_name: Optional[str] = None) -> None:
    # date_str: YYYY/MM/DD or YYYY-MM-DD
    date_str = date_str.replace('-', '/')
    try:
        dt.datetime.strptime(date_str + ' ' + hhmm, "%Y/%m/%d %H:%M")
    except ValueError:
        raise ValueError("日付は YYYY/MM/DD、時刻は HH:MM で指定してください")
    name = task_name or _task_name(alias, "ONCE", date_str.replace('/', '-') + '_' + hhmm.replace(":", "-"))
    _run(["schtasks", "/Delete", "/TN", name, "/F"])
    cmd = [
        "schtasks", "/Create",
        "/TN", name,
        "/SC", "ONCE",
        "/SD", date_str,
        "/ST", hhmm,
        "/TR", _quote(exe_path),
        "/RL", ("HIGHEST" if elevated else "LIMITED"),
        "/F",
    ]
    if utc:
        cmd.append("/Z")
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)
    if random_delay_minutes:
        _set_calendar_random_delay_minutes(name, int(random_delay_minutes))


# 追加トリガー群 -----------------------------------------------------------

def _validate_hhmm(hhmm: str) -> None:
    try:
        dt.datetime.strptime(hhmm, "%H:%M")
    except ValueError:
        raise ValueError("時刻は HH:MM 形式で指定してください")


def ensure_onstart_task(alias: str, exe_path: str, enabled: bool, *, elevated: bool = False, task_name: Optional[str] = None, delay_minutes: Optional[int] = None) -> None:
    """Windows 起動時（ONSTART）のタスクを有効/無効にする。"""
    name = _task_name(alias, "ONSTART")
    if task_name:
        name = task_name
    if enabled:
        _run(["schtasks", "/Delete", "/TN", name, "/F"])
        cmd = [
            "schtasks", "/Create",
            "/TN", name,
            "/SC", "ONSTART",
            "/TR", _quote(exe_path),
            "/RL", ("HIGHEST" if elevated else "LIMITED"),
            "/F",
        ]
        cp = _run(cmd)
        if cp.returncode != 0:
            raise RuntimeError(cp.stderr or cp.stdout)
        _ensure_author(name)
        # 任意: 起動遅延
        if isinstance(delay_minutes, int) and delay_minutes >= 0:
            try:
                _set_trigger_delay_minutes(name, delay_minutes, kind="ONSTART")
            except Exception:
                pass
    else:
        _run(["schtasks", "/Delete", "/TN", name, "/F"])


def create_minutely_task(alias: str, exe_path: str, every_minutes: int, start_time: str, *, sd: Optional[str] = None, ed: Optional[str] = None, et: Optional[str] = None, du: Optional[str] = None, rep_interval_min: Optional[int] = None, rep_duration: Optional[str] = None, stop_at_rep_end: bool = False, random_delay_minutes: Optional[int] = None, utc: bool = False, elevated: bool = False, task_name: Optional[str] = None) -> None:
    """N分おき（MINUTE）。start_time は HH:MM。"""
    if every_minutes < 1 or every_minutes > 1439:
        raise ValueError("分間隔は 1〜1439 の範囲で指定してください")
    _validate_hhmm(start_time)
    name = task_name or _task_name(alias, "MINUTE", f"every{every_minutes}_at_{start_time.replace(':','-')}")
    _run(["schtasks", "/Delete", "/TN", name, "/F"])
    cmd = [
        "schtasks", "/Create",
        "/TN", name,
        "/SC", "MINUTE",
        "/MO", str(every_minutes),
        "/ST", start_time,
        "/TR", _quote(exe_path),
        "/RL", ("HIGHEST" if elevated else "LIMITED"),
        "/F",
    ]
    if utc:
        cmd.append("/Z")
    if rep_interval_min is None:
        _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=du)
    else:
        _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=None)
        _append_repetition(cmd, rep_interval_min=rep_interval_min, rep_duration=rep_duration)
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)
    if random_delay_minutes:
        _set_calendar_random_delay_minutes(name, int(random_delay_minutes))
    if stop_at_rep_end and rep_interval_min is not None:
        _set_repetition_stop_at_end(name, True)


def create_hourly_task(alias: str, exe_path: str, every_hours: int, start_time: str, *, sd: Optional[str] = None, ed: Optional[str] = None, et: Optional[str] = None, du: Optional[str] = None, rep_interval_min: Optional[int] = None, rep_duration: Optional[str] = None, stop_at_rep_end: bool = False, random_delay_minutes: Optional[int] = None, utc: bool = False, elevated: bool = False, task_name: Optional[str] = None) -> None:
    """N時間おき（HOURLY）。start_time は HH:MM。"""
    if every_hours < 1 or every_hours > 168:
        raise ValueError("時間間隔は 1〜168 の範囲で指定してください")
    _validate_hhmm(start_time)
    name = task_name or _task_name(alias, "HOURLY", f"every{every_hours}_at_{start_time.replace(':','-')}")
    _run(["schtasks", "/Delete", "/TN", name, "/F"])
    cmd = [
        "schtasks", "/Create",
        "/TN", name,
        "/SC", "HOURLY",
        "/MO", str(every_hours),
        "/ST", start_time,
        "/TR", _quote(exe_path),
        "/RL", ("HIGHEST" if elevated else "LIMITED"),
        "/F",
    ]
    if utc:
        cmd.append("/Z")
    if rep_interval_min is None:
        _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=du)
    else:
        _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=None)
        _append_repetition(cmd, rep_interval_min=rep_interval_min, rep_duration=rep_duration)
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)
    if random_delay_minutes:
        _set_calendar_random_delay_minutes(name, int(random_delay_minutes))
    if stop_at_rep_end and rep_interval_min is not None:
        _set_repetition_stop_at_end(name, True)


_WEEKDAYS = {"MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"}


def create_weekly_task(alias: str, exe_path: str, hhmm: str, days: List[str], weeks_interval: int = 1, *, sd: Optional[str] = None, ed: Optional[str] = None, et: Optional[str] = None, du: Optional[str] = None, rep_interval_min: Optional[int] = None, rep_duration: Optional[str] = None, stop_at_rep_end: bool = False, random_delay_minutes: Optional[int] = None, utc: bool = False, elevated: bool = False, task_name: Optional[str] = None) -> None:
    """毎週（WEEKLY）。days は ["MON", "TUE", ...]。"""
    _validate_hhmm(hhmm)
    if weeks_interval < 1 or weeks_interval > 52:
        raise ValueError("週間隔は 1〜52 の範囲で指定してください")
    days_norm = [d.strip().upper() for d in days if d.strip()]
    if not days_norm or any(d not in _WEEKDAYS for d in days_norm):
        raise ValueError("曜日は MON,TUE,WED,THU,FRI,SAT,SUN から指定してください")
    dstr = ",".join(days_norm)
    name = task_name or _task_name(alias, "WEEKLY", f"{dstr}_{hhmm.replace(':','-')}_every{weeks_interval}")
    _run(["schtasks", "/Delete", "/TN", name, "/F"])
    cmd = [
        "schtasks", "/Create",
        "/TN", name,
        "/SC", "WEEKLY",
        "/D", dstr,
        "/MO", str(weeks_interval),
        "/ST", hhmm,
        "/TR", _quote(exe_path),
        "/RL", ("HIGHEST" if elevated else "LIMITED"),
        "/F",
    ]
    if utc:
        cmd.append("/Z")
    if rep_interval_min is None:
        _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=du)
    else:
        _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=None)
        _append_repetition(cmd, rep_interval_min=rep_interval_min, rep_duration=rep_duration)
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)
    if random_delay_minutes:
        _set_calendar_random_delay_minutes(name, int(random_delay_minutes))
    if stop_at_rep_end and rep_interval_min is not None:
        _set_repetition_stop_at_end(name, True)


_MONTHS = {"JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"}
_MONTH_NUM_TO_ABBR = {
    1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN",
    7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC",
}


def create_monthly_task(alias: str, exe_path: str, hhmm: str, days: List[str], months: Optional[List[str]] = None, months_interval: int = 1, *, sd: Optional[str] = None, ed: Optional[str] = None, et: Optional[str] = None, du: Optional[str] = None, rep_interval_min: Optional[int] = None, rep_duration: Optional[str] = None, stop_at_rep_end: bool = False, random_delay_minutes: Optional[int] = None, utc: bool = False, elevated: bool = False, task_name: Optional[str] = None) -> None:
    """毎月（MONTHLY）。
    - days: ["1","15","LAST"] のような日付指定。
    - months: ["JAN","FEB",...] または ["1","3","12"]（数値） 省略可。
    - months_interval: 1〜12
    追加オプション: sd, ed, et, du
    """
    _validate_hhmm(hhmm)
    if months_interval < 1 or months_interval > 12:
        raise ValueError("月間隔は 1〜12 の範囲で指定してください")
    if not days:
        raise ValueError("日付を 1〜31 または LAST で指定してください（カンマ区切り可）")
    def _valid_day(s: str) -> bool:
        s = s.strip().upper()
        if s == "LAST":
            return True
        if not s.isdigit():
            return False
        v = int(s)
        return 1 <= v <= 31
    if any(not _valid_day(d) for d in days):
        raise ValueError("日付は 1〜31 または LAST を使用してください")
    dstr = ",".join([d.strip().upper() for d in days])
    mstr = None
    if months:
        norm: List[str] = []
        for m in months:
            ms = m.strip().upper()
            if not ms:
                continue
            if ms.isdigit():
                iv = int(ms)
                if 1 <= iv <= 12:
                    norm.append(_MONTH_NUM_TO_ABBR[iv])
                else:
                    raise ValueError("月は 1〜12 または JAN〜DEC で指定してください")
            else:
                if ms not in _MONTHS:
                    raise ValueError("月は 1〜12 または JAN〜DEC で指定してください")
                norm.append(ms)
        mstr = ",".join(norm)
    suffix = f"days_{dstr}_at_{hhmm.replace(':','-')}_every{months_interval}"
    if mstr:
        suffix += f"_{mstr}"
    name = task_name or _task_name(alias, "MONTHLY", suffix)
    _run(["schtasks", "/Delete", "/TN", name, "/F"])
    cmd = [
        "schtasks", "/Create",
        "/TN", name,
        "/SC", "MONTHLY",
        "/D", dstr,
        "/MO", str(months_interval),
        "/ST", hhmm,
        "/TR", _quote(exe_path),
        "/RL", ("HIGHEST" if elevated else "LIMITED"),
        "/F",
    ]
    if utc:
        cmd.append("/Z")
    if mstr:
        cmd.extend(["/M", mstr])
    if rep_interval_min is None:
        _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=du)
    else:
        _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=None)
        _append_repetition(cmd, rep_interval_min=rep_interval_min, rep_duration=rep_duration)
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)
    if random_delay_minutes:
        _set_calendar_random_delay_minutes(name, int(random_delay_minutes))
    if stop_at_rep_end and rep_interval_min is not None:
        _set_repetition_stop_at_end(name, True)


def create_onidle_task(alias: str, exe_path: str, idle_minutes: int = 10, *, elevated: bool = False, task_name: Optional[str] = None) -> None:
    """アイドル時（ONIDLE）。idle_minutes 分以上アイドルになったときに起動。"""
    if idle_minutes < 1 or idle_minutes > 999:
        raise ValueError("アイドル分は 1〜999 の範囲で指定してください")
    name = task_name or _task_name(alias, "ONIDLE", f"after{idle_minutes}m")
    _run(["schtasks", "/Delete", "/TN", name, "/F"])
    cmd = [
        "schtasks", "/Create",
        "/TN", name,
        "/SC", "ONIDLE",
        "/I", str(idle_minutes),
        "/TR", _quote(exe_path),
        "/RL", ("HIGHEST" if elevated else "LIMITED"),
        "/F",
    ]
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)


def rename_task(old_name: str, new_name: str) -> None:
    """Export a task to XML, create a new with the new name, then delete the old one."""
    # Export XML
    cp = _run(["schtasks", "/Query", "/TN", old_name, "/XML"])
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    xml = cp.stdout
    # Write to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xml", mode="w", encoding="utf-8") as f:
        f.write(xml)
        tmp = f.name
    try:
        # Create new
        cp2 = _run(["schtasks", "/Create", "/TN", new_name, "/XML", tmp, "/F"])
        if cp2.returncode != 0:
            raise RuntimeError(cp2.stderr or cp2.stdout)
        # Delete old
        _run(["schtasks", "/Delete", "/TN", old_name, "/F"])
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass


def change_task_enabled(name: str, enabled: bool) -> None:
    cmd = ["schtasks", "/Change", "/TN", name, "/ENABLE" if enabled else "/DISABLE"]
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)


def _set_trigger_delay_minutes(task_name: str, minutes: int, *, kind: str) -> None:
    """Set <Delay> for Logon/Boot triggers by patching task XML. best-effort."""
    try:
        cp = _run(["schtasks", "/Query", "/TN", task_name, "/XML"])
        if cp.returncode != 0 or not cp.stdout:
            return
        xml = cp.stdout
        trig = "LogonTrigger" if kind.upper() == "LOGON" else "BootTrigger"
        pat = re.compile(fr"(<{trig}[^>]*>)(.*?)(</{trig}>)", flags=re.S)
        m = pat.search(xml)
        if not m:
            return
        head, body, tail = m.group(1), m.group(2), m.group(3)
        # Replace or insert <Delay>PT{M}M
        if re.search(r"<Delay>.*?</Delay>", body, flags=re.S):
            body2 = re.sub(r"<Delay>.*?</Delay>", f"<Delay>PT{minutes}M</Delay>", body)
        else:
            # Insert before closing tag; try after Enabled if exists
            if "</Enabled>" in body:
                body2 = body.replace("</Enabled>", "</Enabled><Delay>PT" + str(minutes) + "M</Delay>")
            else:
                body2 = body + f"<Delay>PT{minutes}M</Delay>"
        new_xml = xml[:m.start()] + head + body2 + tail + xml[m.end():]
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml", mode="w", encoding="utf-8") as f:
            f.write(new_xml)
            tmp = f.name
        try:
            _run(["schtasks", "/Create", "/TN", task_name, "/XML", tmp, "/F"])
        finally:
            try:
                os.remove(tmp)
            except Exception:
                pass
    except Exception:
        # best-effort
        pass


def get_task_details(task_name: str) -> Dict[str, Optional[object]]:
    """Return best-effort parsed details of a task from its XML.
    Keys: Kind, StartTime, EndTime, EveryMinutes, EveryHours, Days, WeeksInterval,
          Months, MonthsInterval, OnceDate, StartDate, EndDate,
          Enabled, Command, Arguments, IdleMinutes, RepeatIntervalMinutes,
          RepeatDuration, RandomDelayMinutes, StopAtDurationEnd, Utc,
          WindowDuration.
    Values may be None if not applicable.
    """
    try:
        cp = _run(["schtasks", "/Query", "/TN", task_name, "/XML"])
        if cp.returncode != 0 or not cp.stdout:
            return {}
        xml = cp.stdout
        out: Dict[str, Optional[object]] = {
            'Kind': None,
            'StartTime': None,
            'EndTime': None,
            'EveryMinutes': None,
            'EveryHours': None,
            'Days': None,
            'WeeksInterval': None,
            'Months': None,
            'MonthsInterval': None,
            'OnceDate': None,
            'StartDate': None,
            'EndDate': None,
            'Enabled': None,
            'Command': None,
            'Arguments': None,
            'IdleMinutes': None,
            'RepeatIntervalMinutes': None,
            'RepeatDuration': None,
            'RandomDelayMinutes': None,
            'StopAtDurationEnd': None,
            'Utc': None,
            'WindowDuration': None,
        }

        def _get(tag: str) -> Optional[str]:
            m = re.search(fr"<{tag}>(.*?)</{tag}>", xml, flags=re.S)
            return m.group(1).strip() if m else None

        # Enabled
        en = _get('Enabled')
        if en is not None:
            out['Enabled'] = en.lower() == 'true'

        # Command/Arguments
        out['Command'] = _get('Command')
        out['Arguments'] = _get('Arguments')

        # Boundaries
        sb = _get('StartBoundary')
        eb = _get('EndBoundary')
        def _ymd(s: str) -> str:
            try:
                # 2025-08-27T14:30:00 -> 2025/08/27
                d = s.split('T', 1)[0]
                y, m, d2 = d.split('-')
                return f"{y}/{m}/{d2}"
            except Exception:
                return s
        def _hm(s: str) -> str:
            try:
                t = s.split('T', 1)[1]
                hh, mm, *_ = t.split(':')
                return f"{hh}:{mm}"
            except Exception:
                return ''
        def _parse_dt(s: str) -> Optional[dt.datetime]:
            try:
                # Remove trailing Z if present; ignore timezone for duration calc
                s2 = s.rstrip('Z')
                # Expect seconds present; if missing, pad
                if len(s2.split('T', 1)[1].split(':')) == 2:
                    s2 = s2 + ":00"
                return dt.datetime.strptime(s2, "%Y-%m-%dT%H:%M:%S")
            except Exception:
                return None
        if sb:
            out['StartDate'] = _ymd(sb)
            out['StartTime'] = _hm(sb)
            # crude UTC detection
            if sb.endswith('Z'):
                out['Utc'] = True
        if eb:
            out['EndDate'] = _ymd(eb)
            out['EndTime'] = _hm(eb)
        # Derive window duration if both boundaries exist
        if sb and eb:
            ds = _parse_dt(sb)
            de = _parse_dt(eb)
            if ds and de and de > ds:
                delta = de - ds
                hours = delta.days * 24 + delta.seconds // 3600
                mins = (delta.seconds % 3600) // 60
                out['WindowDuration'] = f"{hours}:{mins:02d}"

        # Quick kind checks
        if re.search(r"<LogonTrigger>", xml):
            out['Kind'] = 'ONLOGON'
            return out
        if re.search(r"<BootTrigger>", xml):
            out['Kind'] = 'ONSTART'
            return out
        if re.search(r"<IdleTrigger>", xml):
            out['Kind'] = 'ONIDLE'
            return out

    # Repetition
        m_rep = re.search(r"<Repetition>.*?<Interval>PT(\d+)([MH]).*?(?:<Duration>P(?:T\d+H)?(?:T\d+M)?)?", xml, flags=re.S)
        # Time/Calendar triggers
        if re.search(r"<TimeTrigger>", xml) and not m_rep:
            out['Kind'] = 'ONCE'
            # For ONCE, date/time derived from StartBoundary
            if sb:
                out['OnceDate'] = _ymd(sb)
            return out

        # DAILY
        if re.search(r"<ScheduleByDay>", xml):
            out['Kind'] = 'DAILY'
            # time from StartBoundary
            return out

        # WEEKLY
        if re.search(r"<ScheduleByWeek>", xml):
            out['Kind'] = 'WEEKLY'
            # WeeksInterval
            mw = re.search(r"<WeeksInterval>(\d+)</WeeksInterval>", xml)
            if mw:
                out['WeeksInterval'] = int(mw.group(1))
            # DaysOfWeek
            days_map = {
                'Monday': 'MON', 'Tuesday': 'TUE', 'Wednesday': 'WED',
                'Thursday': 'THU', 'Friday': 'FRI', 'Saturday': 'SAT', 'Sunday': 'SUN'
            }
            days: List[str] = []
            md = re.search(r"<DaysOfWeek>(.*?)</DaysOfWeek>", xml, flags=re.S)
            if md:
                block = md.group(1)
                for k, v in days_map.items():
                    if re.search(fr"<{k}\s*/>", block):
                        days.append(v)
            out['Days'] = days or None
            return out

        # MONTHLY
        if re.search(r"<ScheduleByMonth", xml):
            out['Kind'] = 'MONTHLY'
            # MonthsInterval (may appear as <MonthsInterval>)
            mi = re.search(r"<MonthsInterval>(\d+)</MonthsInterval>", xml)
            if mi:
                out['MonthsInterval'] = int(mi.group(1))
            # Months
            months_map = {
                'January': 'JAN', 'February': 'FEB', 'March': 'MAR', 'April': 'APR', 'May': 'MAY', 'June': 'JUN',
                'July': 'JUL', 'August': 'AUG', 'September': 'SEP', 'October': 'OCT', 'November': 'NOV', 'December': 'DEC'
            }
            mm = re.search(r"<Months>(.*?)</Months>", xml, flags=re.S)
            months: List[str] = []
            if mm:
                block = mm.group(1)
                for k, v in months_map.items():
                    if re.search(fr"<{k}\s*/>", block):
                        months.append(v)
            out['Months'] = months or None
            # Days
            days: List[str] = []
            dm = re.search(r"<DaysOfMonth>(.*?)</DaysOfMonth>", xml, flags=re.S)
            if dm:
                block = dm.group(1)
                for m in re.finditer(r"<Day>(\d+)</Day>", block):
                    days.append(m.group(1))
                if re.search(r"<LastDayOfMonth\s*/>", block):
                    days.append('LAST')
            out['Days'] = days or None
            return out

        # Repetition-based
        if m_rep:
            val = int(m_rep.group(1))
            unit = m_rep.group(2)
            # Interval
            interval_min = val if unit == 'M' else val * 60
            out['RepeatIntervalMinutes'] = interval_min
            # Duration（任意）: ISO8601 PT...H...M の簡易抽出
            dur_h = 0
            dur_m = 0
            m_dh = re.search(r"<Duration>P(?:T(\d+)H)?(?:T(\d+)M)?", xml)
            if m_dh:
                if m_dh.group(1):
                    dur_h = int(m_dh.group(1))
                if m_dh.group(2):
                    dur_m = int(m_dh.group(2))
                out['RepeatDuration'] = f"{dur_h}:{dur_m:02d}"
            # StopAtDurationEnd
            mse = re.search(r"<Repetition>.*?<StopAtDurationEnd>(true|false)</StopAtDurationEnd>.*?</Repetition>", xml, flags=re.S|re.I)
            if mse:
                out['StopAtDurationEnd'] = (mse.group(1).lower() == 'true')
            # MINUTE/HOURLYの同定（StartBoundaryの有無とは無関係に）
            if unit == 'M':
                out['Kind'] = 'MINUTE'
                out['EveryMinutes'] = val
            elif unit == 'H':
                out['Kind'] = 'HOURLY'
                out['EveryHours'] = val
            # 他の情報も残して返す
            return out

        # Idle settings
        mi = re.search(r"<IdleSettings>.*?<Duration>PT(\d+)M</Duration>.*?</IdleSettings>", xml, flags=re.S)
        if mi:
            out['IdleMinutes'] = int(mi.group(1))
            if out['Kind'] is None:
                out['Kind'] = 'ONIDLE'

        # RandomDelay for calendar trigger
        mr = re.search(r"<CalendarTrigger[^>]*>.*?<RandomDelay>PT(\d+)M</RandomDelay>.*?</CalendarTrigger>", xml, flags=re.S)
        if not mr:
            mr = re.search(r"<TimeTrigger[^>]*>.*?<RandomDelay>PT(\d+)M</RandomDelay>.*?</TimeTrigger>", xml, flags=re.S)
        if mr:
            try:
                out['RandomDelayMinutes'] = int(mr.group(1))
            except Exception:
                pass

        # Fallback: if we have a StartTime but no kind
        if out['StartTime']:
            out['Kind'] = 'DAILY'
        return out
    except Exception:
        return {}
