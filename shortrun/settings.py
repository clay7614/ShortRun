from __future__ import annotations
import json
import os
import sys
from typing import Any, Dict, Optional

try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover
    win32com = None  # type: ignore

APP_NAME = "ShortRun"


def _config_dir() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    d = os.path.join(base, APP_NAME)
    os.makedirs(d, exist_ok=True)
    return d


def _config_path() -> str:
    return os.path.join(_config_dir(), "config.json")


_DEFAULTS: Dict[str, Any] = {
    "theme": "system",  # system | light | dark
    "autostart": False,
    "last_tab": 0,
}


def load_config() -> Dict[str, Any]:
    path = _config_path()
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
    else:
        data = {}
    # apply defaults
    for k, v in _DEFAULTS.items():
        data.setdefault(k, v)
    return data


def save_config(cfg: Dict[str, Any]) -> None:
    path = _config_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# --- Autostart management ---

def _startup_dir() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, r"Microsoft\Windows\Start Menu\Programs\Startup")


def _startup_link_path() -> str:
    return os.path.join(_startup_dir(), f"{APP_NAME}.lnk")


def is_autostart_enabled() -> bool:
    return os.path.isfile(_startup_link_path())


def set_autostart(enabled: bool) -> None:
    link = _startup_link_path()
    if enabled:
        if win32com is None:
            raise RuntimeError("pywin32 が必要です (win32com.client)")
        try:
            shell = win32com.client.Dispatch("WScript.Shell")  # type: ignore
            shortcut = shell.CreateShortcut(link)
            shortcut.TargetPath = sys.executable
            shortcut.Arguments = "-m shortrun"
            shortcut.WorkingDirectory = os.path.dirname(os.path.dirname(__file__))
            shortcut.WindowStyle = 7  # Minimized
            shortcut.Description = APP_NAME
            shortcut.Save()
        except Exception as e:
            raise RuntimeError(f"スタートアップ登録に失敗: {e}")
    else:
        try:
            if os.path.isfile(link):
                os.remove(link)
        except Exception as e:
            raise RuntimeError(f"スタートアップ解除に失敗: {e}")


# --- Convenience setters ---

def set_theme(cfg: Dict[str, Any], theme: str) -> Dict[str, Any]:
    if theme not in ("system", "light", "dark"):
        theme = "system"
    cfg = dict(cfg)
    cfg["theme"] = theme
    save_config(cfg)
    return cfg


def set_last_tab(cfg: Dict[str, Any], index: int) -> Dict[str, Any]:
    cfg = dict(cfg)
    cfg["last_tab"] = int(index)
    save_config(cfg)
    return cfg
