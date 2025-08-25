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


# 自動起動関連はユーザー要望により削除


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
