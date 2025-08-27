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


def ensure_logon_task(alias: str, exe_path: str, enabled: bool, *, elevated: bool = False, task_name: Optional[str] = None) -> None:
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
    else:
        _run(["schtasks", "/Delete", "/TN", name, "/F"])


def create_daily_task(alias: str, exe_path: str, hhmm: str, *, sd: Optional[str] = None, ed: Optional[str] = None, et: Optional[str] = None, du: Optional[str] = None, elevated: bool = False, task_name: Optional[str] = None) -> None:
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
    _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=du)
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)


def create_once_task(alias: str, exe_path: str, date_str: str, hhmm: str, *, elevated: bool = False, task_name: Optional[str] = None) -> None:
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
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)


# 追加トリガー群 -----------------------------------------------------------

def _validate_hhmm(hhmm: str) -> None:
    try:
        dt.datetime.strptime(hhmm, "%H:%M")
    except ValueError:
        raise ValueError("時刻は HH:MM 形式で指定してください")


def ensure_onstart_task(alias: str, exe_path: str, enabled: bool, *, elevated: bool = False, task_name: Optional[str] = None) -> None:
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
    else:
        _run(["schtasks", "/Delete", "/TN", name, "/F"])


def create_minutely_task(alias: str, exe_path: str, every_minutes: int, start_time: str, *, sd: Optional[str] = None, ed: Optional[str] = None, et: Optional[str] = None, du: Optional[str] = None, elevated: bool = False, task_name: Optional[str] = None) -> None:
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
    _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=du)
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)


def create_hourly_task(alias: str, exe_path: str, every_hours: int, start_time: str, *, sd: Optional[str] = None, ed: Optional[str] = None, et: Optional[str] = None, du: Optional[str] = None, elevated: bool = False, task_name: Optional[str] = None) -> None:
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
    _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=du)
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)


_WEEKDAYS = {"MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"}


def create_weekly_task(alias: str, exe_path: str, hhmm: str, days: List[str], weeks_interval: int = 1, *, sd: Optional[str] = None, ed: Optional[str] = None, et: Optional[str] = None, du: Optional[str] = None, elevated: bool = False, task_name: Optional[str] = None) -> None:
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
    _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=du)
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)


_MONTHS = {"JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"}
_MONTH_NUM_TO_ABBR = {
    1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN",
    7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC",
}


def create_monthly_task(alias: str, exe_path: str, hhmm: str, days: List[str], months: Optional[List[str]] = None, months_interval: int = 1, *, sd: Optional[str] = None, ed: Optional[str] = None, et: Optional[str] = None, du: Optional[str] = None, elevated: bool = False, task_name: Optional[str] = None) -> None:
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
    if mstr:
        cmd.extend(["/M", mstr])
    _append_schedule_window(cmd, sd=sd, ed=ed, et=et, du=du)
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
    _ensure_author(name)


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
