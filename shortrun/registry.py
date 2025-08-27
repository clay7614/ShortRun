from __future__ import annotations
import os
import re
import winreg
from dataclasses import dataclass
from typing import List, Optional

APP_PATHS_KEY = r"Software\\Microsoft\\Windows\\CurrentVersion\\App Paths"
MARKER_NAME = "ShortRun"
MARKER_VALUE = "1"

_alias_re = re.compile(r"^[A-Za-z0-9_-]{1,64}$")


@dataclass
class AliasEntry:
    alias: str
    exe_path: str
    # ShortRun 専用の付加情報
    run_as_admin: bool = False
    comment: Optional[str] = None


def _app_paths_subkey_name(alias: str) -> str:
    # Win+R は拡張子無しでも解決するため、App Paths 上は <alias>.exe で登録する
    return f"{alias}.exe"


def validate_alias(alias: str) -> None:
    if not _alias_re.match(alias):
        raise ValueError("エイリアスは英数字、ハイフン、アンダースコアで 1〜64 文字にしてください。")


def _open_app_paths_key(access=winreg.KEY_READ):
    return winreg.OpenKey(winreg.HKEY_CURRENT_USER, APP_PATHS_KEY, 0, access)


def _ensure_app_paths_key():
    return winreg.CreateKey(winreg.HKEY_CURRENT_USER, APP_PATHS_KEY)


def list_aliases() -> List[AliasEntry]:
    try:
        with _open_app_paths_key() as k:
            entries: List[AliasEntry] = []
            i = 0
            while True:
                try:
                    name = winreg.EnumKey(k, i)
                except OSError:
                    break
                i += 1
                try:
                    with winreg.OpenKey(k, name) as sk:
                        # ShortRun フラグを持つもののみ表示
                        try:
                            marker, _ = winreg.QueryValueEx(sk, MARKER_NAME)
                            if marker != MARKER_VALUE:
                                continue
                        except FileNotFoundError:
                            continue
                        try:
                            exe_path, _ = winreg.QueryValueEx(sk, None)
                        except FileNotFoundError:
                            continue
                        # 追加フラグ（存在しない場合は False）
                        run_admin = False
                        try:
                            ra, _ = winreg.QueryValueEx(sk, "RunAsAdmin")
                            # REG_SZ/REG_DWORD いずれも受け入れる
                            if isinstance(ra, int):
                                run_admin = bool(ra)
                            else:
                                run_admin = str(ra).strip().lower() in ("1", "true", "yes")
                        except FileNotFoundError:
                            run_admin = False
                        alias = name[:-4] if name.lower().endswith(".exe") else name
                        entries.append(AliasEntry(alias=alias, exe_path=exe_path, run_as_admin=run_admin))
                except OSError:
                    continue
            return entries
    except FileNotFoundError:
        return []


def get_alias(alias: str) -> Optional[AliasEntry]:
    name = _app_paths_subkey_name(alias)
    try:
        with _open_app_paths_key() as k:
            with winreg.OpenKey(k, name) as sk:
                try:
                    marker, _ = winreg.QueryValueEx(sk, MARKER_NAME)
                    if marker != MARKER_VALUE:
                        return None
                except FileNotFoundError:
                    return None
                exe_path, _ = winreg.QueryValueEx(sk, None)
                # 追加フラグ（存在しない場合は False）
                run_admin = False
                try:
                    ra, _ = winreg.QueryValueEx(sk, "RunAsAdmin")
                    if isinstance(ra, int):
                        run_admin = bool(ra)
                    else:
                        run_admin = str(ra).strip().lower() in ("1", "true", "yes")
                except FileNotFoundError:
                    run_admin = False
                return AliasEntry(alias=alias, exe_path=exe_path, run_as_admin=run_admin)
    except FileNotFoundError:
        return None


def add_alias(alias: str, exe_path: str, overwrite: bool = False) -> AliasEntry:
    validate_alias(alias)
    exe_path = os.path.abspath(exe_path)
    if not os.path.isfile(exe_path):
        raise FileNotFoundError(f"EXE が見つかりません: {exe_path}")
    name = _app_paths_subkey_name(alias)
    _ensure_app_paths_key()

    # 既存チェック
    existing: Optional[AliasEntry] = None
    try:
        with _open_app_paths_key() as k:
            with winreg.OpenKey(k, name) as sk:
                current_path, _ = winreg.QueryValueEx(sk, None)
                try:
                    marker, _ = winreg.QueryValueEx(sk, MARKER_NAME)
                except FileNotFoundError:
                    marker = None
                existing = AliasEntry(alias=alias, exe_path=current_path, comment=marker)
    except FileNotFoundError:
        pass

    if existing is not None and not overwrite:
        raise FileExistsError("同名のエイリアスが既に存在します。上書きするには overwrite=True を指定してください。")

    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, os.path.join(APP_PATHS_KEY, name)) as sk:
        # 既定値にフルパス、Path にディレクトリ、ShortRun フラグ
        winreg.SetValueEx(sk, None, 0, winreg.REG_SZ, exe_path)
        winreg.SetValueEx(sk, "Path", 0, winreg.REG_SZ, os.path.dirname(exe_path))
        winreg.SetValueEx(sk, MARKER_NAME, 0, winreg.REG_SZ, MARKER_VALUE)
        # 既定は管理者権限で実行しない
        try:
            winreg.SetValueEx(sk, "RunAsAdmin", 0, winreg.REG_DWORD, 0)
        except Exception:
            try:
                winreg.SetValueEx(sk, "RunAsAdmin", 0, winreg.REG_SZ, "0")
            except Exception:
                pass

    return AliasEntry(alias=alias, exe_path=exe_path, run_as_admin=False)


def remove_alias(alias: str) -> None:
    name = _app_paths_subkey_name(alias)
    try:
        with _open_app_paths_key(winreg.KEY_READ | winreg.KEY_WRITE) as k:
            # ShortRun フラグがあるもののみ削除
            with winreg.OpenKey(k, name, 0, winreg.KEY_READ | winreg.KEY_WRITE) as sk:
                try:
                    marker, _ = winreg.QueryValueEx(sk, MARKER_NAME)
                    if marker != MARKER_VALUE:
                        raise PermissionError("ShortRun が作成していないエイリアスは削除しません。")
                except FileNotFoundError:
                    raise PermissionError("ShortRun が作成していないエイリアスは削除しません。")
            winreg.DeleteKey(k, name)
    except FileNotFoundError:
        return


def update_alias(old_alias: str, new_alias: str, new_exe_path: str, overwrite: bool = False) -> AliasEntry:
    """エイリアス名やパスを更新する。
    - alias 変更なし: 既存キーの既定値(Path含む)を更新
    - alias 変更あり: new_alias を作成し old_alias を削除（上書き許可が無ければ既存チェック）
    戻り値は最終的な AliasEntry。
    """
    validate_alias(new_alias)
    new_exe_path = os.path.abspath(new_exe_path)
    if not os.path.isfile(new_exe_path):
        raise FileNotFoundError(f"EXE が見つかりません: {new_exe_path}")

    if old_alias == new_alias:
        # そのまま更新
        name = _app_paths_subkey_name(old_alias)
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, os.path.join(APP_PATHS_KEY, name)) as sk:
            winreg.SetValueEx(sk, None, 0, winreg.REG_SZ, new_exe_path)
            winreg.SetValueEx(sk, "Path", 0, winreg.REG_SZ, os.path.dirname(new_exe_path))
            winreg.SetValueEx(sk, MARKER_NAME, 0, winreg.REG_SZ, MARKER_VALUE)
        # RunAsAdmin は既存値を温存
        run_admin = False
        try:
            ra, _ = winreg.QueryValueEx(sk, "RunAsAdmin")
            if isinstance(ra, int):
                run_admin = bool(ra)
            else:
                run_admin = str(ra).strip().lower() in ("1", "true", "yes")
        except Exception:
            run_admin = False
        return AliasEntry(alias=new_alias, exe_path=new_exe_path, run_as_admin=run_admin)

    # alias が変わる場合
    # 既存 new_alias の有無確認
    if get_alias(new_alias) is not None and not overwrite:
        raise FileExistsError("同名のエイリアスが既に存在します。上書きするには overwrite=True を指定してください。")

    # 新規作成
    new_name = _app_paths_subkey_name(new_alias)
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, os.path.join(APP_PATHS_KEY, new_name)) as sk:
        winreg.SetValueEx(sk, None, 0, winreg.REG_SZ, new_exe_path)
        winreg.SetValueEx(sk, "Path", 0, winreg.REG_SZ, os.path.dirname(new_exe_path))
        winreg.SetValueEx(sk, MARKER_NAME, 0, winreg.REG_SZ, MARKER_VALUE)
        try:
            winreg.SetValueEx(sk, "RunAsAdmin", 0, winreg.REG_DWORD, 0)
        except Exception:
            try:
                winreg.SetValueEx(sk, "RunAsAdmin", 0, winreg.REG_SZ, "0")
            except Exception:
                pass

    # 旧キー削除（ShortRun フラグ確認の上）
    remove_alias(old_alias)
    return AliasEntry(alias=new_alias, exe_path=new_exe_path, run_as_admin=False)


def set_run_as_admin(alias: str, run_as_admin: bool) -> None:
    """エイリアスに対して「管理者として実行」フラグを設定する。"""
    name = _app_paths_subkey_name(alias)
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, os.path.join(APP_PATHS_KEY, name)) as sk:
        try:
            winreg.SetValueEx(sk, "RunAsAdmin", 0, winreg.REG_DWORD, 1 if run_as_admin else 0)
        except Exception:
            try:
                winreg.SetValueEx(sk, "RunAsAdmin", 0, winreg.REG_SZ, "1" if run_as_admin else "0")
            except Exception:
                pass
