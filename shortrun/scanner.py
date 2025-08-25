from __future__ import annotations
import os
import re
import sys
import json
import winreg
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple
import subprocess
import base64
import json

try:
    # pywin32
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover
    win32com = None  # type: ignore

START_MENU_DIRS = [
    os.path.join(os.environ.get("ProgramData", r"C:\\ProgramData"), r"Microsoft\Windows\Start Menu\Programs"),
    os.path.join(os.environ.get("APPDATA", r""), r"Microsoft\Windows\Start Menu\Programs"),
]

UNINSTALL_REG_PATHS = [
    (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Uninstall", winreg.KEY_WOW64_64KEY),
    (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Uninstall", winreg.KEY_WOW64_32KEY),
    (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Uninstall", 0),
]

_icon_path_re = re.compile(r"^\s*\"?(?P<path>[A-Za-z]:[^,\"]+?\.exe)\"?(?:,.*)?$")


@dataclass
class AppCandidate:
    name: str
    exe_path: str
    source: str  # uninstall64/uninstall32/startmenu_system/startmenu_user


def _run_no_window(args: List[str], **kwargs) -> subprocess.CompletedProcess:
    """Run subprocess without showing a console window on Windows.
    Returns CompletedProcess. Adds startupinfo/creationflags when os.name == 'nt'.
    """
    if os.name == 'nt':
        try:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= getattr(subprocess, 'STARTF_USESHOWWINDOW', 0)
            si.wShowWindow = 0  # SW_HIDE
        except Exception:
            si = None  # type: ignore
        creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        kwargs.setdefault('startupinfo', si)
        kwargs.setdefault('creationflags', creationflags)
        kwargs.setdefault('shell', False)
    return subprocess.run(args, **kwargs)


def _iter_registry_keys(root, subkey, access) -> Iterable[Tuple[str, str]]:
    try:
        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | access) as k:
            i = 0
            while True:
                try:
                    name = winreg.EnumKey(k, i)
                except OSError:
                    break
                i += 1
                yield (subkey, name)
    except FileNotFoundError:
        return


def _get_reg_values(root, path, name) -> Dict[str, str]:
    try:
        with winreg.OpenKey(root, os.path.join(path, name)) as sk:
            values: Dict[str, str] = {}
            j = 0
            while True:
                try:
                    vname, vdata, _ = winreg.EnumValue(sk, j)
                except OSError:
                    break
                j += 1
                if isinstance(vdata, str):
                    values[vname] = vdata
            return values
    except FileNotFoundError:
        return {}


def _extract_exe_from_display_icon(display_icon: str) -> Optional[str]:
    if not display_icon:
        return None
    m = _icon_path_re.match(display_icon)
    if m:
        exe = m.group("path")
        if os.path.isfile(exe):
            return exe
    # Fallback: 先頭のクォートを外して .exe を含む部分を探す
    s = display_icon.strip().strip('"')
    idx = s.lower().find('.exe')
    if idx != -1:
        exe = s[: idx + 4]
        if os.path.isfile(exe):
            return exe
    return None


def scan_uninstall() -> List[AppCandidate]:
    results: List[AppCandidate] = []
    for root, path, access in UNINSTALL_REG_PATHS:
        source = (
            "uninstall64" if access == winreg.KEY_WOW64_64KEY else (
                "uninstall32" if access == winreg.KEY_WOW64_32KEY else "uninstall_user"
            )
        )
        for _parent, name in _iter_registry_keys(root, path, access):
            vals = _get_reg_values(root, path, name)
            display_name = vals.get("DisplayName")
            if not display_name:
                continue
            display_icon = vals.get("DisplayIcon", "")
            install_loc = vals.get("InstallLocation", "")
            exe = _extract_exe_from_display_icon(display_icon)
            if not exe and install_loc and os.path.isdir(install_loc):
                # よくあるパターン: <InstallLocation>\<DisplayName>.exe
                guess = os.path.join(install_loc, f"{display_name}.exe")
                if os.path.isfile(guess):
                    exe = guess
            if exe:
                results.append(AppCandidate(name=display_name, exe_path=exe, source=source))
    return results


def _iter_shortcuts(root_dir: str) -> Iterable[str]:
    if not root_dir or not os.path.isdir(root_dir):
        return
    for base, _dirs, files in os.walk(root_dir):
        for fn in files:
            if fn.lower().endswith(".lnk"):
                yield os.path.join(base, fn)


def _resolve_lnk_target(path: str) -> Optional[str]:
    # First try via pywin32
    if win32com is not None:
        try:
            shell = win32com.client.Dispatch("WScript.Shell")  # type: ignore
            shortcut = shell.CreateShortCut(path)
            target = shortcut.TargetPath
            if target and target.lower().endswith(".exe") and os.path.isfile(target):
                return target
        except Exception:
            pass
    # Fallback: use PowerShell to read shortcut target to support packaged envs
    try:
        # Force Unicode (UTF-16LE) output to avoid codepage issues
        cmd = (
            "[Console]::OutputEncoding=[System.Text.Encoding]::Unicode; "
            "(New-Object -ComObject WScript.Shell).CreateShortcut('" + path.replace("'", "''") + "').TargetPath"
        )
        ps = [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            cmd,
        ]
        cp = _run_no_window(ps, capture_output=True, timeout=10)
        out_bytes = cp.stdout or b""
        # Decode as UTF-16LE (Unicode)
        target = (out_bytes.decode('utf-16le', errors='ignore') if out_bytes else '').strip().strip('"')
        if target and target.lower().endswith('.exe') and os.path.isfile(target):
            return target
    except Exception:
        pass
    return None


def _resolve_shortcuts_in_dir(root_dir: str) -> Dict[str, str]:
    """Resolve all .lnk targets under a directory using a single PowerShell process.
    Returns mapping: lnk_path -> target_exe (only valid .exe existing on disk).
    """
    if not root_dir or not os.path.isdir(root_dir):
        return {}
    # Build a PowerShell script and pass via -EncodedCommand (UTF-16LE base64)
    dir_escaped = root_dir.replace("'", "''")
    script = (
        "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8; "
        + "$dir = '{0}'; ".format(dir_escaped)
        + "Get-ChildItem -LiteralPath $dir -Recurse -Filter *.lnk -ErrorAction SilentlyContinue | "
        + "ForEach-Object { try { $s = (New-Object -ComObject WScript.Shell).CreateShortcut($_.FullName); "
        + "$t = $s.TargetPath; if ($t -and $t.ToLower().EndsWith('.exe') -and (Test-Path -LiteralPath $t)) { "
        + "[PSCustomObject]@{ Lnk=$_.FullName; Target=$t } } } catch { } } | ConvertTo-Json -Compress"
    )
    encoded = base64.b64encode(script.encode('utf-16le')).decode('ascii')
    args = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-EncodedCommand",
        encoded,
    ]
    try:
        res = _run_no_window(args, capture_output=True, timeout=15)
        data = res.stdout.decode('utf-8', errors='ignore').strip()
        if not data:
            return {}
        # ConvertTo-Json returns array or single object; normalize to list
        obj = json.loads(data)
        items = obj if isinstance(obj, list) else [obj]
        out: Dict[str, str] = {}
        for it in items:
            try:
                lnk = it.get('Lnk')
                tgt = it.get('Target')
                if lnk and tgt and tgt.lower().endswith('.exe') and os.path.isfile(tgt):
                    out[lnk] = tgt
            except Exception:
                continue
        return out
    except Exception:
        return {}


def scan_start_menu() -> List[AppCandidate]:
    results: List[AppCandidate] = []
    for i, d in enumerate(START_MENU_DIRS):
        src = "startmenu_system" if i == 0 else "startmenu_user"
        # Prefer batch resolution for performance in packaged envs
        mapping = _resolve_shortcuts_in_dir(d)
        if mapping:
            for lnk, exe in mapping.items():
                name = os.path.splitext(os.path.basename(lnk))[0]
                results.append(AppCandidate(name=name, exe_path=exe, source=src))
            continue
        # Fallback to per-link resolution
        for lnk in _iter_shortcuts(d) or []:
            exe = _resolve_lnk_target(lnk)
            if exe:
                name = os.path.splitext(os.path.basename(lnk))[0]
                results.append(AppCandidate(name=name, exe_path=exe, source=src))
    return results


def scan_all(dedup: bool = True) -> List[AppCandidate]:
    items = scan_uninstall() + scan_start_menu()
    if not dedup:
        return items
    seen: Dict[str, AppCandidate] = {}
    for it in items:
        key = os.path.normcase(os.path.abspath(it.exe_path))
        if key not in seen:
            seen[key] = it
    return list(seen.values())
