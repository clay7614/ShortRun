from __future__ import annotations
import datetime as dt
import os
import re
import subprocess
from typing import List, Dict, Optional

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
    return subprocess.run(cmd, capture_output=True, text=True, shell=False)


def list_tasks(alias: Optional[str] = None) -> List[Dict[str, str]]:
    # シンプルに /Query /FO CSV /NH で全件取得し、ShortRun_* だけ拾う
    cp = _run(["schtasks", "/Query", "/FO", "CSV", "/V", "/NH"])
    if cp.returncode != 0:
        return []
    lines = [l.strip() for l in cp.stdout.splitlines() if l.strip()]
    results: List[Dict[str, str]] = []
    for line in lines:
        # CSV 行: "TaskName","Next Run Time","Status", ...
        parts = []
        cur = ''
        in_q = False
        for ch in line:
            if ch == '"':
                in_q = not in_q
                continue
            if ch == ',' and not in_q:
                parts.append(cur)
                cur = ''
            else:
                cur += ch
        parts.append(cur)
        if not parts:
            continue
        name = parts[0]
        # タスク名の先頭に反映される場合、"\\" で始まる
        name = name.lstrip('\\')
        if not name.startswith(_TASK_PREFIX):
            continue
        simple_name = name
        if alias is not None and not simple_name.startswith(_TASK_PREFIX + _sanitize(alias)):
            continue
        next_run = parts[1] if len(parts) > 1 else ''
        status = parts[2] if len(parts) > 2 else ''
        schedule = parts[5] if len(parts) > 5 else ''  # Schedule Type
        results.append({
            'TaskName': name,
            'SimpleName': simple_name,
            'NextRunTime': next_run,
            'Status': status,
            'Schedule': schedule,
        })
    return results


def delete_task_by_simple_name(simple_name: str) -> None:
    _run(["schtasks", "/Delete", "/TN", simple_name, "/F"])


def delete_all_for_alias(alias: str) -> None:
    for t in list_tasks(alias):
        delete_task_by_simple_name(t['SimpleName'])


def ensure_logon_task(alias: str, exe_path: str, enabled: bool) -> None:
    name = _task_name(alias, "LOGON")
    if enabled:
        # 既存は上書き
        _run(["schtasks", "/Delete", "/TN", name, "/F"])
        cmd = [
            "schtasks", "/Create",
            "/TN", name,
            "/SC", "ONLOGON",
            "/TR", _quote(exe_path),
            "/RL", "LIMITED",
            "/F",
        ]
        cp = _run(cmd)
        if cp.returncode != 0:
            raise RuntimeError(cp.stderr or cp.stdout)
    else:
        _run(["schtasks", "/Delete", "/TN", name, "/F"])


def create_daily_task(alias: str, exe_path: str, hhmm: str) -> None:
    # hh:mm を想定
    try:
        dt.datetime.strptime(hhmm, "%H:%M")
    except ValueError:
        raise ValueError("時刻は HH:MM 形式で指定してください")
    name = _task_name(alias, "DAILY", hhmm.replace(":", "-"))
    # 既存は上書き
    _run(["schtasks", "/Delete", "/TN", name, "/F"])
    cmd = [
        "schtasks", "/Create",
        "/TN", name,
        "/SC", "DAILY",
        "/ST", hhmm,
        "/TR", _quote(exe_path),
        "/RL", "LIMITED",
        "/F",
    ]
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)


def create_once_task(alias: str, exe_path: str, date_str: str, hhmm: str) -> None:
    # date_str: YYYY/MM/DD or YYYY-MM-DD
    date_str = date_str.replace('-', '/')
    try:
        dt.datetime.strptime(date_str + ' ' + hhmm, "%Y/%m/%d %H:%M")
    except ValueError:
        raise ValueError("日付は YYYY/MM/DD、時刻は HH:MM で指定してください")
    name = _task_name(alias, "ONCE", date_str.replace('/', '-') + '_' + hhmm.replace(":", "-"))
    _run(["schtasks", "/Delete", "/TN", name, "/F"])
    cmd = [
        "schtasks", "/Create",
        "/TN", name,
        "/SC", "ONCE",
        "/SD", date_str,
        "/ST", hhmm,
        "/TR", _quote(exe_path),
        "/RL", "LIMITED",
        "/F",
    ]
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr or cp.stdout)
