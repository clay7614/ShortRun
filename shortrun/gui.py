from __future__ import annotations
import os
import sys
import re
import subprocess
import ctypes
import time
import threading
from typing import List, Optional, Callable, Tuple
import datetime as dt

import flet as ft

# --- Flet compatibility shims for PyInstaller/runtime ---
# 一部の環境で ft.icons や ft.colors がエクスポートされない場合があるため補完する
def _ensure_ft_attr(attr: str, candidates: list[tuple[str, str]]):
    if hasattr(ft, attr):
        return
    mod = None
    for pkg, name in candidates:
        try:
            mod = __import__(pkg, fromlist=[name])
            obj = getattr(mod, name)
            setattr(ft, attr, obj)
            return
        except Exception:
            continue
    # 最後の手段: ダミーを挿す（文字列名をそのまま返す）
    class _Dummy:
        def __getattr__(self, n):
            return n.lower()
    setattr(ft, attr, _Dummy())

_ensure_ft_attr('icons', [('flet', 'icons'), ('flet_core', 'icons')])
_ensure_ft_attr('colors', [('flet', 'colors'), ('flet_core', 'colors')])

from . import registry
from . import scanner
from . import settings
from . import scheduler

_slug_re = re.compile(r"[^A-Za-z0-9_-]+")


def _slugify(name: str) -> str:
    s = name.strip().lower()
    s = _slug_re.sub("_", s)
    s = s.strip("_")
    return s[:64] if s else "app"


_banner_seq = 0

# --- Assets path helper ------------------------------------------------------
def _asset_path(name: str) -> str:
    """Return absolute path to asset, compatible with PyInstaller."""
    try:
        base = getattr(sys, "_MEIPASS", None) or os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        return os.path.join(base, "assets", name)
    except Exception:
        return os.path.join("assets", name)

def _assets_dir() -> str:
    """Return absolute path to assets directory for ft.app(assets_dir=...)."""
    try:
        base = getattr(sys, "_MEIPASS", None) or os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        d = os.path.join(base, "assets")
        return d
    except Exception:
        return os.path.abspath("assets")


def _post_ui(page: ft.Page, fn: Callable[[], None]):
    """UIスレッドに処理をディスパッチするヘルパ。
    PyInstaller 環境ではバックグラウンドスレッドからの page.update() が無視されることがあるため、
    必ずこの関数経由で UI 変更 + update を行う。
    """
    try:
        poster = getattr(page, "_sr_post_ui", None)
        if callable(poster):
            poster(fn)
        else:
            fn()
    except Exception:
        try:
            fn()
        except Exception:
            pass


def _show_banner(page: ft.Page, message: str, *, error: bool = False, duration: float = 2.5):
    """ページ上部から短時間スライド表示する通知（オーバーレイ）。
    レイアウトを押し下げないためダイアログ下の余白が発生しません。
    """
    global _banner_seq
    _banner_seq += 1
    seq = _banner_seq

    # flet の icons/colors が PyInstaller で見つからない環境でも確実に表示する
    bgcolor = "#EF9A9A" if error else "#90CAF9"  # red200 / blue200
    icon = "error_outline" if error else "info_outline"

    # オーバーレイ用バナーの生成/取得
    banner_row = getattr(page, "_sr_banner_wrapper", None)
    banner_box: Optional[ft.Container] = getattr(page, "_sr_banner", None)
    created = False
    if banner_box is None:
        banner_box = ft.Container(
            bgcolor=bgcolor,
            padding=ft.padding.symmetric(horizontal=16, vertical=10),
            border_radius=8,
            content=ft.Row([ft.Icon(icon, size=18, color="#000000"), ft.Text(message, color="#000000")], spacing=8),
            opacity=0,
            offset=ft.Offset(0, -1),
            animate_opacity=300,
        )
        banner_row = ft.Row([banner_box], alignment=ft.MainAxisAlignment.CENTER)
        page._sr_banner = banner_box
        page._sr_banner_wrapper = banner_row
        page.overlay.append(banner_row)
        created = True
    else:
        banner_box.bgcolor = bgcolor
        banner_box.content = ft.Row([ft.Icon(icon, size=18, color="#000000"), ft.Text(message, color="#000000")], spacing=8)

    # 初回表示は一度マウントを確定してからアニメーション開始
    if created:
        _post_ui(page, lambda: page.update())

    # 表示（スライドイン）
    def _apply_show():
        banner_box.offset = ft.Offset(0, 0)
        banner_box.opacity = 1
        page.update()
    _post_ui(page, _apply_show)

    # 自動クローズ（スライドアウト）
    def _auto_close_sync(s: int, d: float, box: ft.Container):
        try:
            time.sleep(max(0.5, d))
            def _apply_close():
                if s == _banner_seq and getattr(page, "_sr_banner", None) is box:
                    box.opacity = 0
                    box.offset = ft.Offset(0, -1)
                    page.update()
            _post_ui(page, _apply_close)
        except Exception:
            pass

    try:
        threading.Thread(target=_auto_close_sync, args=(seq, duration, banner_box), daemon=True).start()
    except Exception:
        pass


def _show_error(page: ft.Page, message: str):
    _show_banner(page, message, error=True)


def _show_info(page: ft.Page, message: str):
    _show_banner(page, message, error=False)


class AliasTabUI:
    def __init__(self, page: ft.Page):
        self.page = page
        # 別タブへ変更通知（例えばスキャン結果の再描画）
        self.on_alias_changed: Optional[Callable[[], None]] = None
        self.alias_list = ft.ListView(expand=True, spacing=4, padding=8)
        # ソート状態（プログラム一覧）
        self._sort_key: str = "alias"  # alias|path
        self._sort_asc: bool = True
        self.add_alias_field = ft.TextField(label="名称", width=220, tooltip="Win+R で起動する短い名前。英数・-・_ を推奨")
        self.add_path_field = ft.TextField(label="ファイルパス", expand=True, tooltip="起動したいファイルのパス")
        self.file_picker = ft.FilePicker(on_result=self._on_file_picked)
        # FilePicker は overlay に追加するのが推奨
        if self.file_picker not in self.page.overlay:
            self.page.overlay.append(self.file_picker)
        self.open_file_btn = ft.IconButton(ft.icons.FOLDER_OPEN, tooltip="ファイルを選択")
        self.open_file_btn.on_click = self._pick_file
        self.add_btn = ft.ElevatedButton(text="追加", icon=ft.icons.ADD, on_click=self._add_alias, tooltip="入力中の名称とファイルパスでプログラムを登録")
        self.refresh_btn = ft.ElevatedButton(text="再読み込み", icon=ft.icons.REFRESH, tooltip="保存したプログラムを再読み込み", on_click=lambda e: self.refresh())

        # ヘッダー（列名 + ソート）
        self.ha_name_btn = ft.TextButton()
        self.ha_path_btn = ft.TextButton()
        def _set_sort_alias(key: str):
            if self._sort_key == key:
                self._sort_asc = not self._sort_asc
            else:
                self._sort_key = key
                self._sort_asc = True
            self._render_alias_header()
            self.refresh()
        self.ha_name_btn.on_click = lambda e: _set_sort_alias("alias")
        self.ha_path_btn.on_click = lambda e: _set_sort_alias("path")
        self._render_alias_header()
        self.alias_header_row = ft.Row([
            ft.Container(content=self.ha_name_btn, width=180),
            ft.Container(content=self.ha_path_btn, expand=True),
            ft.Container(width=160),  # 操作列のスペース
        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)

        # 旧スケジュール設定 UIは統合のため削除し、行ごとの「スケジュール」ボタンからダイアログを開く
        self.current_alias: Optional[str] = None
        self.current_path: Optional[str] = None

        # 作成者検索ボタン（OSのタスクを作成者で検索）
        author_search_btn = ft.IconButton(
            ft.icons.SEARCH,
            tooltip="OSのスケジュールを作成者で検索",
            on_click=lambda e: self._open_author_search_dialog(),
        )

        self._view = ft.Container(
            content=ft.Column([
                ft.Row([
                    self.add_alias_field,
                    self.add_path_field,
                    self.open_file_btn,
                    self.add_btn,
                    self.refresh_btn,
                    author_search_btn,
                ]),
                ft.Divider(),
                self.alias_header_row,
                self.alias_list,
            ], expand=True, spacing=10),
            padding=ft.padding.only(top=12, left=12, right=12, bottom=8),
        )

    def view(self) -> ft.Control:
        return self._view

    def prefill(self, exe_path: str, alias_name: Optional[str] = None):
        """アプリ一覧からの呼び出しで入力欄を自動セット"""
        self.add_path_field.value = exe_path
        if alias_name:
            self.add_alias_field.value = alias_name
        else:
            base = os.path.splitext(os.path.basename(exe_path))[0]
            self.add_alias_field.value = _slugify(base)
        # 入力フォーカス
        self.add_alias_field.focus()
        self.page.update()

        self.current_alias = self.add_alias_field.value
        self.current_path = self.add_path_field.value
        self._refresh_schedule_toggle()

    def refresh(self):
        # 一旦「読み込み中…」を表示
        self.alias_list.controls.clear()
        self.alias_list.controls.append(ft.Row([ft.ProgressRing(), ft.Text("読み込み中...")]))
        try:
            self.page.update()
        except Exception:
            pass

        entries = registry.list_aliases()
        # 再描画
        self.alias_list.controls.clear()
        if not entries:
            self.alias_list.controls.append(ft.Text("登録された名称はありません。", color=ft.colors.GREY))
        else:
            # ソート
            if self._sort_key == "alias":
                entries.sort(key=lambda x: (x.alias or "").lower(), reverse=not self._sort_asc)
            else:
                entries.sort(key=lambda x: (x.exe_path or "").lower(), reverse=not self._sort_asc)
            for ent in entries:
                self.alias_list.controls.append(self._alias_row(ent))
        self.page.update()

    def _refresh_schedule_toggle(self):
        # 統合後も現在選択中の alias/path は維持
        self.page.update()

    def _alias_row(self, ent: registry.AliasEntry) -> ft.Control:
        def open_schedule_dialog(_: ft.ControlEvent):
            # スケジュール設定をまとめたダイアログ
            alias = ent.alias
            path = ent.exe_path
            # open schedule dialog

            # コントロール（共通）
            logon_sw = ft.Switch(label="ログオン時に起動", value=False, disabled=True)
            onstart_sw = ft.Switch(label="Windows起動時に起動", value=False, disabled=True)

            # 毎日
            daily_tf = ft.TextField(label="時刻 (HH:MM)", width=140)
            # 毎分
            min_every = ft.TextField(label="間隔(分)", width=140)
            min_start = ft.TextField(label="開始時刻 HH:MM", width=160)
            # 毎時
            hr_every = ft.TextField(label="間隔(時間)", width=160)
            hr_start = ft.TextField(label="開始時刻 HH:MM", width=160)
            # 毎週
            weekly_time = ft.TextField(label="時刻 (HH:MM)", width=160)
            weekly_interval = ft.TextField(label="間隔(週)", width=120)
            wd_labels = ["月","火","水","木","金","土","日"]
            wd_check: list[ft.Checkbox] = [ft.Checkbox(label=l, value=False) for l in wd_labels]
            # 毎月
            monthly_time = ft.TextField(label="時刻 (HH:MM)", width=160)
            monthly_days = ft.TextField(label="日付 (例: 1,15,LAST)", width=240)
            monthly_months = ft.TextField(label="対象月 (例: 1,2,3) ※任意", width=240, tooltip="JAN〜DEC も可。数値は1〜12")
            monthly_interval = ft.TextField(label="間隔(月)", width=120)
            # 1回のみ
            once_date = ft.TextField(label="日付 (YYYY/MM/DD)", width=180)
            once_time = ft.TextField(label="時刻 (HH:MM)", width=140)
            # 共通スケジュール名（任意）
            schedule_name_tf = ft.TextField(label="スケジュール名 (任意)", width=420, tooltip="未入力なら自動命名。種別や時刻を変えると自動更新します")
            name_user_override = {"v": False}

            # 日付ピッカー（必要時に生成して開く: 表示崩れ防止）
            def _open_dp(target_tf: ft.TextField):
                def _on_change(ev: ft.ControlEvent):
                    try:
                        if ev.control.value:
                            target_tf.value = ev.control.value.strftime('%Y/%m/%d')
                        if dp in self.page.overlay:
                            self.page.overlay.remove(dp)
                        self.page.update()
                    except Exception:
                        pass
                dp = ft.DatePicker(on_change=_on_change)
                self.page.overlay.append(dp)
                def _on_dismiss(_: ft.ControlEvent):
                    try:
                        if dp in self.page.overlay:
                            self.page.overlay.remove(dp)
                        self.page.update()
                    except Exception:
                        pass
                dp.on_dismiss = _on_dismiss
                dp.open = True
                self.page.update()

            once_pick_btn = ft.IconButton(ft.icons.CALENDAR_MONTH, tooltip="カレンダーから選択", on_click=lambda e: _open_dp(once_date))

            # 期間オプション（共通）
            sd_tf = ft.TextField(label="開始日 YYYY/MM/DD", width=190, tooltip="空欄可")
            ed_tf = ft.TextField(label="終了日 YYYY/MM/DD", width=190, tooltip="空欄可")
            et_tf = ft.TextField(label="終了時刻 HH:MM", width=160, tooltip="空欄可")
            du_tf = ft.TextField(label="期間 HHH:MM", width=160, tooltip="空欄可。例 12:00")
            sd_pick_btn = ft.IconButton(ft.icons.CALENDAR_TODAY, tooltip="開始日を選択", on_click=lambda e: _open_dp(sd_tf))
            ed_pick_btn = ft.IconButton(ft.icons.EVENT, tooltip="終了日を選択", on_click=lambda e: _open_dp(ed_tf))

            # デフォルト値（現在日時）
            now = dt.datetime.now()
            hhmm_now = now.strftime("%H:%M")
            ymd_now = now.strftime("%Y/%m/%d")
            daily_tf.value = hhmm_now
            min_start.value = hhmm_now
            hr_start.value = hhmm_now
            weekly_time.value = hhmm_now
            monthly_time.value = hhmm_now
            once_date.value = ymd_now
            once_time.value = hhmm_now
            sd_tf.value = ymd_now

            # トグル状態のみ更新
            def refresh_toggles():
                tasks = scheduler.list_tasks(alias)
                logon_sw.value = any(t['SimpleName'].endswith('_LOGON') for t in tasks)
                onstart_sw.value = any(t['SimpleName'].endswith('_ONSTART') for t in tasks)
                self.page.update()

            def delete_all(e: ft.ControlEvent):
                try:
                    scheduler.delete_all_for_alias(alias)
                    self.page.close(dlg)
                    _show_info(self.page, "スケジュールを削除しました")
                except Exception as ex:
                    _show_error(self.page, f"削除に失敗: {ex}")

            # 遅延時間（任意）を分入力（ドロップダウン外に配置）
            idle_minutes = ft.TextField(label="遅延時間(分)", width=140, tooltip="空欄なら変更しません")

            # 種別選択（ドロップダウン・単一選択）
            schedule_type_dd = ft.Dropdown(
                label="スケジュール種別",
                width=300,
                options=[
                    ft.dropdown.Option("DAILY", "毎日"),
                    ft.dropdown.Option("MIN", "毎分"),
                    ft.dropdown.Option("HOUR", "毎時"),
                    ft.dropdown.Option("WEEK", "毎週"),
                    ft.dropdown.Option("MONTH", "毎月"),
                    ft.dropdown.Option("ONCE", "1回のみ"),
                ],
                value=None,
                padding=ft.padding.only(bottom=12)
            )
            schedule_type_dd.disabled = True

            # 各セクション（表示は選択に応じて切替）
            sec_daily = ft.Container(content=ft.Column([ft.Row([daily_tf], spacing=8)], spacing=8), visible=False)
            sec_min = ft.Container(content=ft.Column([ft.Row([min_every, min_start], spacing=8)], spacing=8), visible=False)
            sec_hour = ft.Container(content=ft.Column([ft.Row([hr_every, hr_start], spacing=8)], spacing=8), visible=False)
            sec_week = ft.Container(content=ft.Column([
                ft.Row([weekly_time, weekly_interval], spacing=8),
                ft.Row(wd_check, spacing=8),
            ], spacing=8), visible=False)
            sec_month = ft.Container(content=ft.Column([
                ft.Row([monthly_time, monthly_interval], spacing=8),
                ft.Row([monthly_days, monthly_months], spacing=8),
            ], spacing=8), visible=False)
            sec_once = ft.Container(content=ft.Column([
                ft.Row([once_date, once_pick_btn, once_time], spacing=8),
            ], spacing=8), visible=False)

            # 既定名の算出（scheduler と同等の命名規則に合わせる）
            _re_sanitize = re.compile(r"[^A-Za-z0-9_-]+")
            def _sanitize(s: str) -> str:
                s = (s or "").strip()
                s = _re_sanitize.sub("_", s)
                return s[:60]
            def _month_abbr_list(vals: list[str]) -> list[str]:
                """Convert numeric months to JAN..DEC; pass through JAN.. etc."""
                m_abbr = {1:"JAN",2:"FEB",3:"MAR",4:"APR",5:"MAY",6:"JUN",7:"JUL",8:"AUG",9:"SEP",10:"OCT",11:"NOV",12:"DEC"}
                out: list[str] = []
                for v in vals:
                    vs = (v or "").strip().upper()
                    if not vs:
                        continue
                    if vs.isdigit():
                        iv = int(vs)
                        if 1 <= iv <= 12:
                            out.append(m_abbr[iv])
                    else:
                        out.append(vs)
                return out
            def _wd_tokens() -> list[str]:
                labels = ["月","火","水","木","金","土","日"]
                token = {"月":"MON","火":"TUE","水":"WED","木":"THU","金":"FRI","土":"SAT","日":"SUN"}
                res: list[str] = []
                for i, cb in enumerate(wd_check):
                    if cb.value:
                        res.append(token[labels[i]])
                return res
            def _compute_default_name() -> str:
                typ = schedule_type_dd.value
                if not typ:
                    return ""
                alias_safe = _sanitize(alias)
                def _tn(kind: str, suffix: str | None = None) -> str:
                    name = f"ShortRun_{alias_safe}_{kind}"
                    if suffix:
                        name += f"_{_sanitize(suffix)}"
                    return name
                if typ == "DAILY":
                    hhmm = (daily_tf.value or "").strip()
                    return _tn("DAILY", hhmm.replace(":","-"))
                if typ == "MIN":
                    every = (min_every.value or "").strip() or "1"
                    st = (min_start.value or "").strip()
                    return _tn("MINUTE", f"every{every}_at_{st.replace(':','-')}")
                if typ == "HOUR":
                    every = (hr_every.value or "").strip() or "1"
                    st = (hr_start.value or "").strip()
                    return _tn("HOURLY", f"every{every}_at_{st.replace(':','-')}")
                if typ == "WEEK":
                    d = _wd_tokens()
                    hhmm = (weekly_time.value or "").strip()
                    interval = (weekly_interval.value or "").strip() or "1"
                    dstr = ",".join(d) if d else "MON"
                    return _tn("WEEKLY", f"{dstr}_{hhmm.replace(':','-')}_every{interval}")
                if typ == "MONTH":
                    hhmm = (monthly_time.value or "").strip()
                    d = ",".join([(s or "").strip().upper() for s in (monthly_days.value or "").split(',') if s.strip()]) or "1"
                    m = [s for s in (monthly_months.value or "").split(',') if s.strip()]
                    m_abbr = _month_abbr_list(m)
                    interval = (monthly_interval.value or "").strip() or "1"
                    suffix = f"days_{d}_at_{hhmm.replace(':','-')}_every{interval}"
                    if m_abbr:
                        suffix += f"_{','.join(m_abbr)}"
                    return _tn("MONTHLY", suffix)
                if typ == "ONCE":
                    d = (once_date.value or "").strip().replace('/','-')
                    t_ = (once_time.value or "").strip().replace(':','-')
                    return _tn("ONCE", f"{d}_{t_}")
                return ""

            def _maybe_set_default_name(_: Optional[ft.ControlEvent] = None):
                if not name_user_override["v"]:
                    schedule_name_tf.value = _compute_default_name()
                    try:
                        self.page.update()
                    except Exception:
                        pass

            # ユーザーが名前を変更したら以降自動更新しない
            def _on_name_change(_: ft.ControlEvent):
                name_user_override["v"] = True if (schedule_name_tf.value or "").strip() else False
            schedule_name_tf.on_change = _on_name_change

            def _update_sections(_: Optional[ft.ControlEvent] = None):
                typ = schedule_type_dd.value
                sec_daily.visible = typ == "DAILY"
                sec_min.visible = typ == "MIN"
                sec_hour.visible = typ == "HOUR"
                sec_week.visible = typ == "WEEK"
                sec_month.visible = typ == "MONTH"
                sec_once.visible = typ == "ONCE"
                self.page.update()

            schedule_type_dd.on_change = lambda e: (_update_sections(e), _maybe_set_default_name(e))

            # ローディング表示（初期は見せる）
            loading_row = ft.Row([ft.ProgressRing(), ft.Text("読み込み中...")], spacing=8)
            loading_box = ft.Container(content=loading_row, visible=True, padding=ft.padding.only(top=8,left=8))

            # 期間指定（左）
            window_opts = ft.Container(
                expand=True,
                content=ft.Column([
                    ft.Container(
                        content=ft.Text("期間指定", color=ft.colors.GREY),
                        padding=ft.padding.only(bottom=8),
                    ),
                    ft.Row([sd_tf, sd_pick_btn, ed_tf, ed_pick_btn], spacing=8),
                    ft.Row([et_tf, du_tf], spacing=8),
                ], spacing=8),
            )

            # 遅延時間（右）
            idle_opts = ft.Container(
                expand=True,
                content=ft.Column([
                    ft.Container(
                        content=ft.Text("遅延時間", color=ft.colors.GREY),
                        padding=ft.padding.only(bottom=8),
                    ),
                    ft.Row([idle_minutes], spacing=8),
                ], spacing=8),
            )

            # オプション（任意）全体をまとめる
            options_group = ft.Container(
                content=ft.Column([
                    ft.Container(
                        content=ft.Text("オプション（任意）", color=ft.colors.GREY),
                        padding=ft.padding.only(bottom=8),
                    ),
                    ft.Column([window_opts, idle_opts], spacing=16),
                ], spacing=8),
            )

            content = ft.Container(
                width=920,
                height=620,
                content=ft.Column([
                    loading_box,
                    schedule_name_tf,
                    ft.Text(f"対象: {alias}", color=ft.colors.GREY),
                    logon_sw,
                    onstart_sw,
                    ft.Divider(),
                    ft.Text("適用する種類を選択", color=ft.colors.GREY),
                    schedule_type_dd,
                    sec_daily,
                    sec_min,
                    sec_hour,
                    sec_week,
                    sec_month,
                    sec_once,
                    ft.Divider(),
                    options_group,
                ], tight=True, spacing=8, scroll=ft.ScrollMode.ALWAYS),
            )

            # 初期表示反映
            _update_sections()
            _maybe_set_default_name()

            def save_all(e: ft.ControlEvent):
                errors: List[str] = []
                elevated = bool(settings.load_config().get("run_as_admin", False))
                # ログオン
                try:
                    scheduler.ensure_logon_task(alias, path, bool(logon_sw.value), elevated=elevated)
                except Exception as ex:
                    errors.append(f"ログオン時: {ex}")
                # 起動時
                try:
                    scheduler.ensure_onstart_task(alias, path, bool(onstart_sw.value), elevated=elevated)
                except Exception as ex:
                    errors.append(f"起動時: {ex}")
                typ = schedule_type_dd.value
                # 毎日
                if typ == "DAILY":
                    hhmm = (daily_tf.value or '').strip()
                    if not hhmm:
                        errors.append("毎日: 時刻を入力してください")
                    else:
                        try:
                            scheduler.create_daily_task(
                                alias, path, hhmm,
                                sd=(sd_tf.value or '').strip() or None,
                                ed=(ed_tf.value or '').strip() or None,
                                et=(et_tf.value or '').strip() or None,
                                du=(du_tf.value or '').strip() or None,
                                elevated=elevated,
                                task_name=((schedule_name_tf.value or '').strip() or None),
                            )
                        except Exception as ex:
                            errors.append(f"毎日: {ex}")
                # 毎分
                if typ == "MIN":
                    try:
                        every = int((min_every.value or '0').strip())
                        st = (min_start.value or '').strip()
                        scheduler.create_minutely_task(
                            alias, path, every, st,
                            sd=(sd_tf.value or '').strip() or None,
                            ed=(ed_tf.value or '').strip() or None,
                            et=(et_tf.value or '').strip() or None,
                            du=(du_tf.value or '').strip() or None,
                            elevated=elevated,
                            task_name=((schedule_name_tf.value or '').strip() or None),
                        )
                    except Exception as ex:
                        errors.append(f"毎分: {ex}")
                # 毎時
                if typ == "HOUR":
                    try:
                        every = int((hr_every.value or '0').strip())
                        st = (hr_start.value or '').strip()
                        scheduler.create_hourly_task(
                            alias, path, every, st,
                            sd=(sd_tf.value or '').strip() or None,
                            ed=(ed_tf.value or '').strip() or None,
                            et=(et_tf.value or '').strip() or None,
                            du=(du_tf.value or '').strip() or None,
                            elevated=elevated,
                            task_name=((schedule_name_tf.value or '').strip() or None),
                        )
                    except Exception as ex:
                        errors.append(f"毎時: {ex}")
                # 毎週
                if typ == "WEEK":
                    try:
                        sel_days = [cb.label for cb in wd_check if cb.value]
                        interval = int((weekly_interval.value or '1').strip() or '1')
                        jp_to_en = {"月":"MON","火":"TUE","水":"WED","木":"THU","金":"FRI","土":"SAT","日":"SUN"}
                        days = [jp_to_en.get(d, d) for d in sel_days]
                        if not days:
                            raise ValueError("曜日を1つ以上選択してください")
                        if not (weekly_time.value or '').strip():
                            raise ValueError("時刻を入力してください")
                        scheduler.create_weekly_task(
                            alias, path, weekly_time.value.strip(), days, interval,
                            sd=(sd_tf.value or '').strip() or None,
                            ed=(ed_tf.value or '').strip() or None,
                            et=(et_tf.value or '').strip() or None,
                            du=(du_tf.value or '').strip() or None,
                            elevated=elevated,
                            task_name=((schedule_name_tf.value or '').strip() or None),
                        )
                    except Exception as ex:
                        errors.append(f"毎週: {ex}")
                # 毎月
                if typ == "MONTH":
                    try:
                        d = [x.strip() for x in (monthly_days.value or '').split(',') if x.strip()]
                        m = [x.strip() for x in (monthly_months.value or '').split(',') if x.strip()] if (monthly_months.value or '').strip() else None
                        interval = int((monthly_interval.value or '1').strip() or '1')
                        scheduler.create_monthly_task(
                            alias, path, monthly_time.value.strip(), d, m, interval,
                            sd=(sd_tf.value or '').strip() or None,
                            ed=(ed_tf.value or '').strip() or None,
                            et=(et_tf.value or '').strip() or None,
                            du=(du_tf.value or '').strip() or None,
                            elevated=elevated,
                            task_name=((schedule_name_tf.value or '').strip() or None),
                        )
                    except Exception as ex:
                        errors.append(f"毎月: {ex}")
                # 1回
                if typ == "ONCE":
                    od = (once_date.value or '').strip()
                    ot = (once_time.value or '').strip()
                    if od and ot:
                        try:
                            scheduler.create_once_task(alias, path, od, ot, elevated=elevated, task_name=((schedule_name_tf.value or '').strip() or None))
                        except Exception as ex:
                            errors.append(f"1回のみ: {ex}")
                    else:
                        errors.append("「1回のみ」は日付と時刻を両方入力してください")
                # 遅延時間（任意）
                try:
                    im_str = (idle_minutes.value or "").strip()
                    if im_str:
                        scheduler.create_onidle_task(
                            alias, path, int(im_str), elevated=elevated,
                            task_name=((schedule_name_tf.value or '').strip() or None)
                        )
                except Exception as ex:
                    errors.append(f"遅延時間: {ex}")

                # トグルのみ更新
                refresh_toggles()

                if errors:
                    _show_error(self.page, "\n".join(errors))
                else:
                    _show_info(self.page, "保存しました")
                    try:
                        self.page.close(dlg)
                    except Exception:
                        pass

            # Enterで保存
            for tf in [daily_tf, min_every, min_start, hr_every, hr_start, weekly_time, weekly_interval, monthly_time, monthly_days, monthly_months, monthly_interval, once_date, once_time, sd_tf, ed_tf, et_tf, du_tf, idle_minutes]:
                try:
                    tf.on_submit = save_all
                    tf.on_change = _maybe_set_default_name
                except Exception:
                    pass
            try:
                for cb in wd_check:
                    cb.on_change = _maybe_set_default_name
            except Exception:
                pass

            # 保存ボタンは読み込み完了まで無効化
            save_btn = ft.TextButton("保存", on_click=save_all, disabled=True)

            dlg = ft.AlertDialog(
                modal=True,
                title=ft.Text(f"スケジュール設定: {alias}"),
                content=content,
                actions=[
                    save_btn,
                    ft.TextButton("閉じる", on_click=lambda e: self.page.close(dlg)),
                ],
            )
            self.page.open(dlg)

            # 非同期でトグル状態を取得し、完了後にUIを有効化
            def _fetch_toggles():
                try:
                    tasks = scheduler.list_tasks(alias)
                except Exception:
                    tasks = []
                def _apply_after():
                    try:
                        logon_sw.value = any(t['SimpleName'].endswith('_LOGON') for t in tasks)
                        onstart_sw.value = any(t['SimpleName'].endswith('_ONSTART') for t in tasks)
                        logon_sw.disabled = False
                        onstart_sw.disabled = False
                        schedule_type_dd.disabled = False
                        loading_box.visible = False
                        save_btn.disabled = False
                    finally:
                        self.page.update()
                _post_ui(self.page, _apply_after)
            threading.Thread(target=_fetch_toggles, daemon=True).start()

        return ft.Container(
            content=ft.Row([
                ft.Text(ent.alias, width=180, weight=ft.FontWeight.BOLD),
                ft.Text(ent.exe_path, expand=True, selectable=True),
                ft.IconButton(ft.icons.SCHEDULE, tooltip="スケジュール設定を開く", on_click=open_schedule_dialog),
                ft.IconButton(ft.icons.PLAY_ARROW, tooltip="プログラムの起動", on_click=lambda e, p=ent.exe_path, ra=getattr(ent, 'run_as_admin', False): self._launch(p, ra)),
                ft.IconButton(ft.icons.EDIT, tooltip="名称とパスを編集", on_click=lambda e, entry=ent: self._edit_alias(entry)),
                ft.IconButton(ft.icons.DELETE, tooltip="プログラムを削除", on_click=lambda e, a=ent.alias: self._confirm_remove(a)),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            padding=6,
        )

    def _open_author_search_dialog(self):
        page = self.page
        author_tf = ft.TextField(label="作成者", value="ShortRun", width=260)
        progress = ft.Row([ft.ProgressRing(), ft.Text("検索中...")], visible=False)
        list_view = ft.ListView(expand=True, spacing=4, padding=4)

        def load():
            auth = (author_tf.value or "").strip() or "ShortRun"
            try:
                tasks = scheduler.list_tasks(author=auth)
            except Exception:
                tasks = []
            def apply():
                list_view.controls.clear()
                if not tasks:
                    list_view.controls.append(ft.Text("該当タスクなし", color=ft.colors.GREY))
                else:
                    for t in tasks:
                        nm = t.get('SimpleName') or t.get('TaskName')
                        en = t.get('Enabled')
                        nx = t.get('NextRunTime') or ''
                        stxt = "有効" if en else ("無効" if en is not None else (t.get('Status') or ''))
                        list_view.controls.append(
                            ft.Row([
                                ft.Text(nm, expand=True),
                                ft.Text(stxt, width=80, color=ft.colors.GREY),
                                ft.Text(nx, width=180),
                            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
                        )
                progress.visible = False
                try:
                    page.update()
                except Exception:
                    pass
            _post_ui(page, apply)

        def do_search(_: ft.ControlEvent = None):
            progress.visible = True
            list_view.controls.clear()
            try:
                page.update()
            except Exception:
                pass
            threading.Thread(target=load, daemon=True).start()

        author_tf.on_submit = do_search

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text("作成者でOSスケジュール検索"),
            content=ft.Container(
                width=760,
                height=520,
                content=ft.Column([
                    ft.Row([author_tf, ft.TextButton("検索", on_click=do_search)]),
                    progress,
                    ft.Divider(),
                    list_view,
                ], expand=True),
            ),
            actions=[
                ft.TextButton("閉じる", on_click=lambda e: page.close(dlg)),
            ],
        )
        page.open(dlg)

    def _render_alias_header(self):
        def label_for(key: str, jp: str) -> str:
            if self._sort_key == key:
                return f"{jp} {'▲' if self._sort_asc else '▼'}"
            return jp
        self.ha_name_btn.text = label_for("alias", "名称")
        self.ha_path_btn.text = label_for("path", "パス")
        try:
            self.page.update()
        except Exception:
            pass

    def _launch(self, path: str, run_as_admin: bool = False):
        try:
            lp = (path or "").lower()
            # 管理者要求時は拡張子に関係なく ShellExecuteW("runas") を優先
            if run_as_admin and os.name == 'nt':
                try:
                    ctypes.windll.shell32.ShellExecuteW(None, "runas", path, None, None, 1)
                    return
                except Exception:
                    pass
            if lp.endswith(".lnk") or lp.endswith(".url"):
                # ショートカット/URL は Shell 経由で開く
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen([path], close_fds=True)
        except Exception as ex:
            _show_error(self.page, f"起動に失敗しました: {ex}")

    def _confirm_remove(self, alias: str):
        def do_remove(_: ft.ControlEvent):
            try:
                registry.remove_alias(alias)
                self.page.close(dlg)
                _show_info(self.page, f"削除しました: {alias}")
                self.refresh()
                if self.on_alias_changed:
                    try:
                        self.on_alias_changed()
                    except Exception:
                        pass
            except Exception as ex:
                _show_error(self.page, f"削除に失敗: {ex}")
        dlg = ft.AlertDialog(
            title=ft.Text("確認"),
            content=ft.Text(f"プログラム '{alias}' を削除しますか？"),
            actions=[
                ft.TextButton("キャンセル", on_click=lambda e: self.page.close(dlg)),
                ft.TextButton("削除", on_click=do_remove),
            ],
        )
        self.page.open(dlg)

    def _edit_alias(self, entry: registry.AliasEntry):
        page = self.page
        alias_tf = ft.TextField(label="名称", value=entry.alias, width=220)
        path_tf = ft.TextField(label="ファイルパス", value=entry.exe_path, expand=True)
        admin_sw = ft.Switch(label="管理者として実行", value=bool(getattr(entry, "run_as_admin", False)))

        # ローカル FilePicker（編集用）
        def _on_pick(res: ft.FilePickerResultEvent):
            if res.files:
                path_tf.value = res.files[0].path
                page.update()

        fp = ft.FilePicker(on_result=_on_pick)
        if fp not in page.overlay:
            page.overlay.append(fp)
        pick_btn = ft.IconButton(
            ft.icons.FOLDER_OPEN,
            tooltip="ファイル / ショートカットを選択",
            on_click=lambda e: fp.pick_files(
                allow_multiple=False,
                file_type=ft.FilePickerFileType.ANY,
                dialog_title="ファイルまたはショートカットを選択",
            ),
        )

        content = ft.Container(
            width=560,
            content=ft.Column([
                ft.Row([alias_tf], alignment=ft.MainAxisAlignment.START),
                ft.Row([path_tf, pick_btn], alignment=ft.MainAxisAlignment.START),
                ft.Row([admin_sw], alignment=ft.MainAxisAlignment.START),
            ], spacing=10, tight=True),
        )

        def do_save(e: ft.ControlEvent, *, overwrite: bool = False):
            new_alias = (alias_tf.value or "").strip()
            new_path = (path_tf.value or "").strip().strip('"')
            if not new_alias or not new_path:
                _show_error(page, "名称とファイルのパスを入力してください。")
                return
            try:
                registry.update_alias(entry.alias, new_alias, new_path, overwrite=overwrite)
                try:
                    registry.set_run_as_admin(new_alias, bool(admin_sw.value))
                except Exception:
                    pass
                page.close(dlg)
                _show_info(page, f"更新しました: {entry.alias} → {new_alias}")
                self.refresh()
                if self.on_alias_changed:
                    try:
                        self.on_alias_changed()
                    except Exception:
                        pass
            except FileExistsError:
                # 上書き確認
                def confirm_over(_: ft.ControlEvent):
                    try:
                        registry.update_alias(entry.alias, new_alias, new_path, overwrite=True)
                        try:
                            registry.set_run_as_admin(new_alias, bool(admin_sw.value))
                        except Exception:
                            pass
                        page.close(confirm)
                        page.close(dlg)
                        _show_info(page, f"上書きしました: {new_alias}")
                        self.refresh()
                        if self.on_alias_changed:
                            try:
                                self.on_alias_changed()
                            except Exception:
                                pass
                    except Exception as ex:
                        _show_error(page, f"上書きに失敗: {ex}")
                confirm = ft.AlertDialog(
                    title=ft.Text("既存の名称"),
                    content=ft.Text(f"{new_alias} は既に存在します。上書きしますか？"),
                    actions=[
                        ft.TextButton("キャンセル", on_click=lambda _: page.close(confirm)),
                        ft.TextButton("上書き", on_click=confirm_over),
                    ],
                )
                page.open(confirm)
            except Exception as ex:
                _show_error(page, f"更新に失敗: {ex}")

        # Enterで保存
        alias_tf.on_submit = do_save
        path_tf.on_submit = do_save

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text("プログラムの編集"),
            content=content,
            actions=[
                ft.TextButton("保存", on_click=do_save),
                ft.TextButton("閉じる", on_click=lambda e: page.close(dlg)),
            ],
        )
        page.open(dlg)

    def _pick_file(self, e: ft.ControlEvent):
        self.file_picker.pick_files(
            allow_multiple=False,
            file_type=ft.FilePickerFileType.ANY,
            dialog_title="ファイルまたはショートカットを選択",
        )

    def _on_file_picked(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.add_path_field.value = e.files[0].path
            if not self.add_alias_field.value:
                base = os.path.splitext(os.path.basename(e.files[0].path))[0]
                self.add_alias_field.value = _slugify(base)
            self.page.update()

    def _add_alias(self, e: ft.ControlEvent):
        alias = self.add_alias_field.value.strip()
        path = self.add_path_field.value.strip().strip('"')
        if not alias or not path:
            _show_error(self.page, "名称とファイルのパスを入力してください。")
            return
        try:
            registry.add_alias(alias, path, overwrite=False)
            _show_info(self.page, f"登録しました: {alias}")
            self.add_alias_field.value = ""
            self.add_path_field.value = ""
            self.refresh()
            if self.on_alias_changed:
                try:
                    self.on_alias_changed()
                except Exception:
                    pass
        except FileExistsError:
            def do_overwrite(_: ft.ControlEvent):
                try:
                    registry.add_alias(alias, path, overwrite=True)
                    _show_info(self.page, f"上書きしました: {alias}")
                    self.page.close(dialog)
                    self.refresh()
                    if self.on_alias_changed:
                        try:
                            self.on_alias_changed()
                        except Exception:
                            pass
                except Exception as ex:
                    _show_error(self.page, f"上書きに失敗: {ex}")
            dialog = ft.AlertDialog(
                title=ft.Text("既存の名称"),
                content=ft.Text("同名の名称が存在します。上書きしますか？"),
                actions=[
                    ft.TextButton("キャンセル", on_click=lambda _: self.page.close(dialog)),
                    ft.TextButton("上書き", on_click=do_overwrite),
                ],
            )
            # use page.open for reliability
            self.page.open(dialog)
        except Exception as ex:
            _show_error(self.page, f"登録に失敗: {ex}")

    def _on_toggle_logon(self, e: ft.ControlEvent):
        if not self.current_alias or not self.current_path:
            _show_error(self.page, "スケジュール対象のプログラムを一覧から選択してください")
            self.logon_switch.value = False
            self.page.update()
            return
        try:
            scheduler.ensure_logon_task(self.current_alias, self.current_path, bool(self.logon_switch.value))
            _show_info(self.page, "ログオン時起動を更新しました")
        except Exception as ex:
            _show_error(self.page, f"更新に失敗: {ex}")
            self._refresh_schedule_toggle()

    def _on_add_daily(self, e: ft.ControlEvent):
        if not self.current_alias or not self.current_path:
            _show_error(self.page, "スケジュール対象のプログラムを一覧から選択してください")
            return
        try:
            scheduler.create_daily_task(self.current_alias, self.current_path, (self.daily_time.value or '').strip())
            _show_info(self.page, "毎日スケジュールを追加しました")
        except Exception as ex:
            _show_error(self.page, f"追加に失敗: {ex}")

    def _on_add_once(self, e: ft.ControlEvent):
        if not self.current_alias or not self.current_path:
            _show_error(self.page, "スケジュール対象のプログラムを一覧から選択してください")
            return
        try:
            scheduler.create_once_task(
                self.current_alias,
                self.current_path,
                (self.once_date.value or '').strip(),
                (self.once_time.value or '').strip(),
            )
            _show_info(self.page, "「1回のみ」スケジュールを追加しました")
        except Exception as ex:
            _show_error(self.page, f"追加に失敗: {ex}")

class ScanTabUI:
    def __init__(self, page: ft.Page, on_alias_added: Optional[Callable[[], None]] = None, on_request_prefill: Optional[Callable[[str, str], None]] = None, *, cfg: Optional[dict] = None):
        self.page = page
        self.on_alias_added = on_alias_added
        self.on_request_prefill = on_request_prefill
        self.cfg = cfg or {}
        self.items: List[scanner.AppCandidate] = []
        self._hidden_paths: set[str] = set()  # 既に登録済みの exe パス（正規化）
        self._visible_keys: set[str] = set()  # 現在表示中の候補キー（正規化パス）
        # ソート状態
        self._sort_key: str = "name"  # name|exe|source
        self._sort_asc: bool = True
        self.filter_field = ft.TextField(
            hint_text="アプリ名でフィルタ（-で除外）",
            expand=True,
            on_change=lambda e: self._render_list(),
            tooltip="スペース区切りで複数語を指定できます。-foo のように先頭に - を付けると除外します。"
        )
        self.scan_btn = ft.ElevatedButton("再読み込み", icon=ft.icons.REFRESH, on_click=lambda e: self.scan(), tooltip="アプリの候補一覧を再取得")
        self.bulk_add_btn = ft.ElevatedButton("選択したプログラムを一括追加", icon=ft.icons.ADD_TASK, on_click=lambda e: self._bulk_add(), tooltip="チェック済みの候補をまとめて登録")
        self.list_view = ft.ListView(expand=True, spacing=4, padding=8)

        # ヘッダー（列名 + ソート）
        def _btn(label: str, on_click: Callable[[ft.ControlEvent], None], *, width: Optional[int] = None, expand: bool = False) -> ft.Control:
            b = ft.TextButton(text=label, on_click=on_click)
            if width is not None:
                return ft.Container(content=b, width=width)
            if expand:
                return ft.Container(content=b, expand=True)
            return b

        self.h_name_btn = ft.TextButton()
        self.h_path_btn = ft.TextButton()
        self.h_source_btn = ft.TextButton()
        # ダミー幅（チェックボックス列/操作列）
        left_pad = ft.Container(width=28)
        right_pad = ft.Container(width=44)
        self.header_row = ft.Row([
            left_pad,
            ft.Container(content=self.h_name_btn, width=240),
            ft.Container(content=self.h_path_btn, expand=True),
            ft.Container(content=self.h_source_btn, width=140),
            right_pad,
        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)

        def _set_sort(key: str):
            if self._sort_key == key:
                self._sort_asc = not self._sort_asc
            else:
                self._sort_key = key
                self._sort_asc = True
            self._render_header()
            self._render_list()

        self.h_name_btn.on_click = lambda e: _set_sort("name")
        self.h_path_btn.on_click = lambda e: _set_sort("exe")
        self.h_source_btn.on_click = lambda e: _set_sort("source")

        # 件数ステータスと選択解除
        self.status_text = ft.Text("該当: 0 / 選択: 0", color=ft.colors.GREY)
        self.clear_sel_btn = ft.ElevatedButton(
            "選択を解除",
            icon=ft.icons.CANCEL_PRESENTATION,
            on_click=lambda e: self._clear_selection(),
            tooltip="チェック済みの選択をすべて解除します",
        )

        self._render_header()
        self._view = ft.Container(
            content=ft.Column([
                ft.Row([self.filter_field, self.clear_sel_btn, self.scan_btn, self.bulk_add_btn]),
                ft.Divider(),
                ft.Row([self.status_text], alignment=ft.MainAxisAlignment.END),
                self.header_row,
                self.list_view,
            ], expand=True),
            padding=ft.padding.only(top=12, left=12, right=12, bottom=8),
        )
        # 選択状態
        self._selected: set[str] = set()
        self._scanning: bool = False

    def view(self) -> ft.Control:
        return self._view

    def scan(self):
        if self._scanning:
            return
        self._scanning = True
        self.list_view.controls.clear()
        self.list_view.controls.append(ft.Row([ft.ProgressRing(), ft.Text("読み込み中...")]))
        try:
            self.page.update()
        except Exception:
            pass

        def _work():
            try:
                show_uninst = bool((self.cfg or {}).get("show_uninstallers", False))
                items = scanner.scan_all(show_uninstallers=show_uninst)

                def _apply():
                    self.items = items
                    # 現在の登録済みパスを収集
                    try:
                        regs = registry.list_aliases()
                        self._hidden_paths = {os.path.normcase(os.path.abspath(e.exe_path)) for e in regs}
                    except Exception:
                        self._hidden_paths = set()
                    self._render_list()
                    self._scanning = False

                _post_ui(self.page, _apply)
            except Exception:
                def _apply_err():
                    self.list_view.controls.clear()
                    self.list_view.controls.append(ft.Text("読み込みに失敗しました", color=ft.colors.RED))
                    try:
                        self.page.update()
                    except Exception:
                        pass
                    self._scanning = False

                _post_ui(self.page, _apply_err)

        threading.Thread(target=_work, daemon=True).start()

    def _render_list(self):
        q = (self.filter_field.value or "").strip().lower()
        self.list_view.controls.clear()

        def norm(p: str) -> str:
            try:
                return os.path.normcase(os.path.abspath(p))
            except Exception:
                return p

        # 除外語（-prefix）対応のフィルタ
        def parse_terms(text: str) -> tuple[list[str], list[str]]:
            if not text:
                return [], []
            parts = [t for t in re.split(r"\s+", text) if t]
            include = [t for t in parts if not t.startswith("-")]
            exclude = [t[1:] for t in parts if t.startswith("-") and len(t) > 1]
            return include, exclude

        include_terms, exclude_terms = parse_terms(q)

        def passes(i: scanner.AppCandidate) -> bool:
            name = (i.name or "").lower()
            base = os.path.basename(i.exe_path).lower()
            # 検索対象は名称とパス（ソースは対象外）
            hay = [name, base]
            # include は AND（すべて含む）
            if include_terms and not all(any(t in h for h in hay) for t in include_terms):
                return False
            # exclude は OR（いずれか含めば除外）
            if any(t and (t in name or t in base) for t in exclude_terms):
                return False
            # include 指定が無い場合は全件対象
            return True

        matched_all = [i for i in self.items if passes(i)]
        # 既に登録された exe パスは除外
        matched = [i for i in matched_all if norm(i.exe_path) not in self._hidden_paths]
        # ソート
        def key_name(i: scanner.AppCandidate) -> str:
            return (i.name or "").lower()

        def key_exe(i: scanner.AppCandidate) -> str:
            try:
                return (i.exe_path or "").lower()
            except Exception:
                return i.exe_path or ""

        def key_source(i: scanner.AppCandidate) -> str:
            return (i.source or "").lower()

        key_funcs = {"name": key_name, "exe": key_exe, "source": key_source}
        kf = key_funcs.get(self._sort_key, key_name)
        matched.sort(key=kf, reverse=not self._sort_asc)
        # 表示中キーを保存し件数更新
        try:
            self._visible_keys = {norm(i.exe_path) for i in matched}
        except Exception:
            self._visible_keys = set()
        self._update_status(len(matched))
        if not matched:
            self.list_view.controls.append(ft.Text("該当なし", color=ft.colors.GREY))
        else:
            for it in matched:
                self.list_view.controls.append(self._row(it))
        self.page.update()

    def _render_header(self):
        def label_for(key: str, jp: str) -> str:
            if self._sort_key == key:
                return f"{jp} {'▲' if self._sort_asc else '▼'}"
            return jp
        self.h_name_btn.text = label_for("name", "アプリ名")
        self.h_path_btn.text = label_for("exe", "パス")
        self.h_source_btn.text = label_for("source", "ソース")
        try:
            self.page.update()
        except Exception:
            pass

    def _row(self, it: scanner.AppCandidate) -> ft.Control:
        def toggle_selected(e: ft.ControlEvent):
            key = os.path.normcase(os.path.abspath(it.exe_path))
            if key in self._selected:
                self._selected.remove(key)
                cb.value = False
            else:
                self._selected.add(key)
                cb.value = True
            # 件数のみ更新
            try:
                self._update_status(len(self._visible_keys))
            except Exception:
                pass
            self.page.update()

        def create_alias(_: ft.ControlEvent):
            # ダイアログで名称を編集してから登録
            suggested = _slugify(it.name) if it.name else _slugify(os.path.splitext(os.path.basename(it.exe_path))[0])
            alias_tf = ft.TextField(label="名称", value=suggested, width=320)

            def do_register(e: ft.ControlEvent, *, overwrite: bool = False):
                name = (alias_tf.value or "").strip()
                if not name:
                    _show_error(self.page, "名称を入力してください")
                    return
                try:
                    registry.add_alias(name, it.exe_path, overwrite=overwrite)
                    self.page.close(dlg)
                    _show_info(self.page, ("上書きしました: " if overwrite else "登録しました: ") + name)
                    if self.on_alias_added:
                        try:
                            self.on_alias_added()
                        except Exception:
                            pass
                    # 追加した exe は非表示にして再描画
                    try:
                        self._hidden_paths.add(os.path.normcase(os.path.abspath(it.exe_path)))
                    except Exception:
                        pass
                    self._render_list()
                except FileExistsError:
                    # 上書き確認
                    def confirm_over(_: ft.ControlEvent):
                        try:
                            registry.add_alias(name, it.exe_path, overwrite=True)
                            self.page.close(confirm)
                            self.page.close(dlg)
                            _show_info(self.page, f"上書きしました: {name}")
                            if self.on_alias_added:
                                try:
                                    self.on_alias_added()
                                except Exception:
                                    pass
                            try:
                                self._hidden_paths.add(os.path.normcase(os.path.abspath(it.exe_path)))
                            except Exception:
                                pass
                            self._render_list()
                        except Exception as ex:
                            _show_error(self.page, f"上書きに失敗: {ex}")
                    confirm = ft.AlertDialog(
                        title=ft.Text("既存の名称"),
                        content=ft.Text(f"{name} は既に存在します。上書きしますか？"),
                        actions=[
                            ft.TextButton("キャンセル", on_click=lambda _: self.page.close(confirm)),
                            ft.TextButton("上書き", on_click=confirm_over),
                        ],
                    )
                    self.page.open(confirm)
                except Exception as ex:
                    _show_error(self.page, f"登録に失敗: {ex}")

            # Enterで保存
            alias_tf.on_submit = do_register

            dlg = ft.AlertDialog(
                title=ft.Text("アプリを追加"),
                content=ft.Container(
                    width=420,
                    content=ft.Column([
                        ft.Text(os.path.basename(it.exe_path), color=ft.colors.GREY),
                        alias_tf,
                    ], spacing=10, tight=True),
                ),
                actions=[
                    ft.TextButton("キャンセル", on_click=lambda e: self.page.close(dlg)),
                    ft.TextButton("登録", on_click=do_register),
                ],
            )
            self.page.open(dlg)

        cb = ft.Checkbox(value=False, on_change=toggle_selected, tooltip="一括追加の対象として選択/解除")
        return ft.Container(
            content=ft.Row([
                cb,
                ft.Text(it.name, width=240, weight=ft.FontWeight.BOLD),
                ft.Text(it.exe_path, expand=True, selectable=True),
                ft.Text(it.source, width=140, color=ft.colors.GREY),
                ft.IconButton(ft.icons.ADD, tooltip="このアプリを追加", on_click=create_alias),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            padding=6,
        )

    def _bulk_add(self):
        # 選択されたエントリをまとめて追加（ダイアログで名前編集）
        targets = [i for i in self.items if os.path.normcase(os.path.abspath(i.exe_path)) in self._selected]
        if not targets:
            _show_info(self.page, "一括追加する項目をチェックしてください")
            return

        # 行ごとの入力欄を生成
        rows: list[tuple[scanner.AppCandidate, ft.TextField]] = []
        for it in targets:
            suggested = _slugify(it.name) if it.name else _slugify(os.path.splitext(os.path.basename(it.exe_path))[0])
            tf = ft.TextField(value=suggested, width=300)
            rows.append((it, tf))

        overwrite_cb = ft.Checkbox(label="既存の名称は上書きする", value=False)

        items_column = ft.Column([
            ft.Row([ft.Text(os.path.basename(it.exe_path), width=280), tf], spacing=12)
            for (it, tf) in rows
        ], spacing=8)

        content = ft.Container(
            width=700,
            height=480,
            content=ft.Column([
                ft.Text("一括追加: 名称を確認・編集してから登録します", color=ft.colors.GREY),
                ft.Container(content=items_column, expand=True, bgcolor=ft.colors.with_opacity(0.02, ft.colors.BLACK), padding=8),
                overwrite_cb,
            ], spacing=10, tight=True, scroll=ft.ScrollMode.ALWAYS),
        )

        def do_register(_: ft.ControlEvent):
            added = 0
            errors: List[Tuple[str, str]] = []
            for it, tf in rows:
                name = (tf.value or "").strip()
                if not name:
                    errors.append((os.path.basename(it.exe_path), "名称が未入力"))
                    continue
                try:
                    registry.add_alias(name, it.exe_path, overwrite=bool(overwrite_cb.value))
                    added += 1
                    try:
                        self._hidden_paths.add(os.path.normcase(os.path.abspath(it.exe_path)))
                    except Exception:
                        pass
                except Exception as ex:
                    errors.append((name, str(ex)))

            self.page.close(dlg)
            msg = f"追加: {added}"
            if errors:
                msg += f" / 失敗: {len(errors)}"
            _show_info(self.page, msg)

            # 再描画＆選択解除
            self._selected.clear()
            self._render_list()
            if self.on_alias_added:
                try:
                    self.on_alias_added()
                except Exception:
                    pass

        # Enterで保存（最後のテキストフィールドに送る）
        try:
            rows[-1][1].on_submit = do_register
        except Exception:
            pass

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text("アプリを一括追加"),
            content=content,
            actions=[
                ft.TextButton("キャンセル", on_click=lambda e: self.page.close(dlg)),
                ft.TextButton("追加", on_click=do_register),
            ],
        )
        self.page.open(dlg)

    def _update_status(self, matched_count: int):
        try:
            selected_in_view = sum(1 for k in self._selected if k in self._visible_keys)
            self.status_text.value = f"該当: {matched_count} / 選択: {selected_in_view}"
        except Exception:
            self.status_text.value = f"該当: {matched_count} / 選択: 0"
        try:
            self.page.update()
        except Exception:
            pass

    def _clear_selection(self):
        # 全選択解除（再描画せずに既存行のチェックだけ外してステータス更新）
        try:
            self._selected.clear()
        except Exception:
            self._selected = set()
        # 既存行のチェックボックスをオフ
        try:
            for item in list(self.list_view.controls):
                row = getattr(item, "content", None)
                ctrls = getattr(row, "controls", None)
                if isinstance(ctrls, list) and ctrls:
                    cb = ctrls[0]
                    if isinstance(cb, ft.Checkbox):
                        cb.value = False
        except Exception:
            pass
        # 件数ステータスのみ更新
        try:
            self._update_status(len(getattr(self, "_visible_keys", set())))
        except Exception:
            pass
        try:
            self.page.update()
        except Exception:
            pass


class SettingsTabUI:
    def __init__(self, page: ft.Page, cfg: dict, on_config_changed: Optional[Callable[[], None]] = None):
        self.page = page
        self.cfg = cfg
        self._on_config_changed = on_config_changed
        # テーマ
        self.theme_dropdown = ft.Dropdown(
            label="テーマ",
            value=cfg.get("theme", "system"),
            options=[
                ft.dropdown.Option("system", "システムに合わせる"),
                ft.dropdown.Option("light", "ライト"),
                ft.dropdown.Option("dark", "ダーク"),
            ],
            on_change=self._on_theme_changed,
            width=260,
        )
        self._view = ft.Container(
            content=ft.Column([
                ft.Text("設定", weight=ft.FontWeight.BOLD, size=18),
                self.theme_dropdown,
                ft.Row([
                    ft.Switch(
                        label="管理者として実行",
                        value=bool(cfg.get("run_as_admin", False)),
                        on_change=self._on_toggle_run_as_admin,
                        tooltip="スケジュール設定を行う場合、管理者権限が必要になります",
                    )
                ]),
                ft.Row([
                    ft.Switch(
                        label="アプリ一覧にアンインストーラを表示する(要再起動)",
                        value=bool(cfg.get("show_uninstallers", False)),
                        on_change=self._on_toggle_uninstaller,
                    )
                ]),
                ft.Divider(),
                ft.Text("その他"),
                ft.Text("Tips: エラーや不具合が発生した場合はgithubのREADMEを参考にするか、\nissuesでの報告をお願いします。", color=ft.colors.GREY),
                # GitHub ヘルプアイコン（ICO画像）。テーマに応じて白/黒へ反転させる
                self._build_help_icon(),
            ], expand=False, spacing=12),
            padding=ft.padding.only(top=12, left=12, right=12, bottom=8),
        )

    def view(self) -> ft.Control:
        return self._view

    def _build_help_icon(self) -> ft.Control:
        # 画像本体
        try:
            icon_path = _asset_path("github.ico")
        except Exception:
            icon_path = os.path.join("assets", "github.ico")
        self.help_img = ft.Image(src=icon_path, width=26, height=26)
        # 可能ならブレンドモードで単色化を有効にする
        try:
            bm = getattr(ft, "BlendMode", None)
            if bm is not None:
                self.help_img.color_blend_mode = bm.SRC_IN
        except Exception:
            pass
        # コンテナでクリック/ツールチップを付与
        help_btn = ft.Container(
            content=self.help_img,
            tooltip="ヘルプ (GitHub) を開く",
            on_click=lambda e: self.page.launch_url(
                "https://github.com/clay7614/ShortRun?tab=readme-ov-file#%E4%B8%80%E8%88%AC%E3%83%A6%E3%83%BC%E3%82%B6%E3%83%BC%E5%90%91%E3%81%91"
            ),
            padding=4,
        )
        # 初期色の適用
        self._apply_help_icon_theme()
        return help_btn

    def _apply_help_icon_theme(self):
        # 現在のテーマから明暗を推定し、アイコン色を白/黒で反転
        try:
            # page.theme_mode が LIGHT/DARK/SYSTEM のいずれか
            mode = getattr(self.page, "theme_mode", None)
            is_dark = False
            if mode == ft.ThemeMode.DARK:
                is_dark = True
            elif mode == ft.ThemeMode.LIGHT:
                is_dark = False
            else:
                # SYSTEM の場合は platform_brightness があれば参照（無ければライト扱い）
                pb = getattr(self.page, "platform_brightness", None)
                # flet では "dark" / "light" の文字列または Brightness 列挙が入る可能性に配慮
                if str(pb).lower().endswith("dark"):
                    is_dark = True
            self.help_img.color = ft.colors.WHITE if is_dark else ft.colors.BLACK
        except Exception:
            # フォールバック: 黒
            try:
                self.help_img.color = ft.colors.BLACK
            except Exception:
                pass

    def _on_theme_changed(self, e: ft.ControlEvent):
        theme = self.theme_dropdown.value or "system"
        self.cfg = settings.set_theme(self.cfg, theme)
        # 適用
        if theme == "light":
            self.page.theme_mode = ft.ThemeMode.LIGHT
        elif theme == "dark":
            self.page.theme_mode = ft.ThemeMode.DARK
        else:
            self.page.theme_mode = ft.ThemeMode.SYSTEM
        # ヘルプアイコンの色も更新
        try:
            self._apply_help_icon_theme()
        except Exception:
            pass
        self.page.update()

    # 自動起動設定は削除済み

    def _on_toggle_uninstaller(self, e: ft.ControlEvent):
        show = bool(e.control.value)
        self.cfg = settings.set_show_uninstallers(self.cfg, show)
        if callable(self._on_config_changed):
            try:
                self._on_config_changed()
            except Exception:
                pass
        self.page.update()

    def _on_toggle_run_as_admin(self, e: ft.ControlEvent):
        val = bool(e.control.value)
        self.cfg = settings.set_run_as_admin(self.cfg, val)
        # ヒントを表示
        if val:
            _show_info(self.page, "管理者として実行を有効にしました")
        else:
            _show_info(self.page, "管理者として実行を無効にしました")
        self.page.update()


class ScheduleTabUI:
    def __init__(self, page: ft.Page):
        self.page = page
        self.alias_filter = ft.TextField(hint_text="スケジュール名でフィルタ", width=240, on_change=lambda e: self.refresh())
        self.refresh_btn = ft.IconButton(ft.icons.REFRESH, tooltip="再読み込み", on_click=lambda e: self.refresh())
        self.list_view = ft.ListView(expand=True, spacing=4, padding=8)
        self._refreshing = False
        self._last_refresh_ts = 0.0
        self._view = ft.Container(
            content=ft.Column([
                ft.Row([self.alias_filter, self.refresh_btn]),
                ft.Divider(),
                self.list_view,
            ], expand=True),
            padding=ft.padding.only(top=12, left=12, right=12, bottom=8),
        )

    def view(self) -> ft.Control:
        return self._view

    def refresh(self):
        # 短時間の連続呼び出しを抑止（簡易デバウンス）
        now = time.time()
        if self._refreshing or (now - self._last_refresh_ts) < 0.3:
            return
        self._refreshing = True
        self._last_refresh_ts = now
        q = (self.alias_filter.value or "").strip().lower()
        # 一旦「読み込み中…」を表示
        self.list_view.controls.clear()
        self.list_view.controls.append(ft.Row([ft.ProgressRing(), ft.Text("読み込み中...")]))
        try:
            self.page.update()
        except Exception:
            pass

        def _work():
            try:
                # Author=ShortRun のタスクを一覧表示
                tasks = scheduler.list_tasks(author="ShortRun")
            except Exception as ex:
                tasks = []

            def _apply():
                self.list_view.controls.clear()
                _tasks = tasks
                if q:
                    _tasks = [t for t in _tasks if q in t['SimpleName'].lower()]
                if not _tasks:
                    self.list_view.controls.append(ft.Text("タスクはありません", color=ft.colors.GREY))
                else:
                    for t in _tasks:
                        name = t['SimpleName']
                        when = t.get('NextRunTime', '')
                        sched = t.get('Schedule', '')
                        enabled_flag = t.get('Enabled')
                        status_str = ("有効" if enabled_flag else ("無効" if enabled_flag is not None else t.get('Status','')))
                        def make_row(simple_name=name, when_text=when, sched_text=sched, status_text=status_str, enabled_val=enabled_flag):
                            def do_delete(_: ft.ControlEvent):
                                def _confirm(_: ft.ControlEvent):
                                    try:
                                        scheduler.delete_task_by_simple_name(simple_name)
                                        self.page.close(dlg)
                                        _show_info(self.page, f"削除しました: {simple_name}")
                                        self.refresh()
                                    except Exception as ex:
                                        _show_error(self.page, f"削除に失敗: {ex}")
                                dlg = ft.AlertDialog(
                                    title=ft.Text("確認"),
                                    content=ft.Text(f"タスク '{simple_name}' を削除しますか？"),
                                    actions=[
                                        ft.TextButton("キャンセル", on_click=lambda e: self.page.close(dlg)),
                                        ft.TextButton("削除", on_click=_confirm),
                                    ],
                                )
                                self.page.open(dlg)
                            def do_edit(_: ft.ControlEvent):
                                try:
                                    enabled_now = bool(enabled_val) if enabled_val is not None else ((status_text or '').strip().lower() != 'disabled')
                                    name_tf = ft.TextField(label="タスク名", value=simple_name, width=420)
                                    # スイッチのラベルは状態に応じて切り替え
                                    en_sw = ft.Switch(label=("有効" if enabled_now else "無効"), value=enabled_now)
                                    def _on_sw_change(ev: ft.ControlEvent):
                                        try:
                                            ev.control.label = "有効" if bool(ev.control.value) else "無効"
                                            self.page.update()
                                        except Exception:
                                            pass
                                    en_sw.on_change = _on_sw_change
                                    # 保存アクション
                                    def _apply_edit(new_name_val: str, enabled_val: bool):
                                        new_name = (new_name_val or '').strip()
                                        if not new_name:
                                            _show_error(self.page, "タスク名を入力してください")
                                            return
                                        old = simple_name
                                        try:
                                            # 先にリネーム、次に有効/無効変更（新名で適用）
                                            target_name = old
                                            if new_name != old:
                                                scheduler.rename_task(old, new_name)
                                                target_name = new_name
                                            # 有効状態の変更
                                            if enabled_val != enabled_now:
                                                scheduler.change_task_enabled(target_name, enabled_val)
                                            _show_info(self.page, "保存しました")
                                            try:
                                                self.page.close(dlg)
                                            except Exception:
                                                pass
                                            # キーボードハンドラ復元
                                            try:
                                                self.page.on_keyboard_event = prev_kb
                                            except Exception:
                                                pass
                                            self.refresh()
                                        except Exception as ex:
                                            _show_error(self.page, f"保存に失敗: {ex}")

                                    dlg = ft.AlertDialog(
                                        modal=True,
                                        title=ft.Text("タスクの編集"),
                                        content=ft.Column([
                                            name_tf,
                                            en_sw,
                                        ], tight=True, spacing=8),
                                        actions=[
                                            ft.TextButton("キャンセル", on_click=lambda e: (self.page.close(dlg), setattr(self.page, 'on_keyboard_event', prev_kb))),
                                            ft.TextButton("保存", on_click=lambda e: _apply_edit(name_tf.value or '', en_sw.value)),
                                        ],
                                    )
                                    # Enterキーで保存（ダイアログ表示中のみ有効）
                                    prev_kb = getattr(self.page, 'on_keyboard_event', None)
                                    def _kb(ev: ft.KeyboardEvent):
                                        try:
                                            k = str(getattr(ev, 'key', '')).lower()
                                            if 'enter' in k:
                                                _apply_edit(name_tf.value or '', en_sw.value)
                                        except Exception:
                                            pass
                                    try:
                                        self.page.on_keyboard_event = _kb
                                    except Exception:
                                        pass
                                    # TextField でも Enter で保存可能に
                                    try:
                                        name_tf.on_submit = lambda e: _apply_edit(name_tf.value or '', en_sw.value)
                                    except Exception:
                                        pass
                                    self.page.open(dlg)
                                except Exception as ex:
                                    _show_error(self.page, f"編集ダイアログの表示に失敗: {ex}")
                            return ft.Container(
                                content=ft.Row([
                                    ft.Text(simple_name, expand=True),
                                    ft.Text(when_text, width=200),
                                    ft.Text(sched_text, width=160, color=ft.colors.GREY),
                                    ft.Text(status_text or "", width=80, color=ft.colors.GREY),
                                    ft.IconButton(ft.icons.EDIT, tooltip="名前変更・有効/無効", on_click=do_edit),
                                    ft.IconButton(ft.icons.DELETE, tooltip="このタスクを削除", on_click=do_delete),
                                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                                padding=6,
                            )
                        self.list_view.controls.append(make_row())
                try:
                    self.page.update()
                finally:
                    self._refreshing = False

            _post_ui(self.page, _apply)

        threading.Thread(target=_work, daemon=True).start()

def main(page: ft.Page):
    page.title = "ShortRUN"
    # Window icon (absolute path is more reliable for desktop windows)
    try:
        ico = _asset_path("shortrun.ico")
        png = _asset_path("shortrun.png")
        if os.path.isfile(ico):
            page.window.icon = ico
        elif os.path.isfile(png):
            page.window.icon = png
    except Exception:
        pass
    page.window.width = 980
    page.window.height = 680
    page.horizontal_alignment = ft.CrossAxisAlignment.STRETCH

    # UIスレッド実行ヘルパ（pubsub）を登録しておくと、バックグラウンドからの UI 更新確認が容易
    try:
        def _sr_ui_consumer(msg):
            try:
                if isinstance(msg, tuple) and len(msg) == 2 and msg[0] == "__sr_ui__" and callable(msg[1]):
                    msg[1]()
            except Exception:
                pass
        page.pubsub.subscribe(_sr_ui_consumer)
        page._sr_post_ui = lambda fn: page.pubsub.send_all(("__sr_ui__", fn))
    except Exception:
        page._sr_post_ui = lambda fn: fn()

    cfg = settings.load_config()
    # 初期テーマ適用
    if cfg.get("theme") == "light":
        page.theme_mode = ft.ThemeMode.LIGHT
    elif cfg.get("theme") == "dark":
        page.theme_mode = ft.ThemeMode.DARK
    else:
        page.theme_mode = ft.ThemeMode.SYSTEM

    alias_ui = AliasTabUI(page)
    scan_ui = ScanTabUI(page, on_alias_added=alias_ui.refresh, cfg=cfg)
    # 双方向の通知: エイリアス変更時に探索を再描画
    alias_ui.on_alias_changed = scan_ui.scan
    settings_ui = SettingsTabUI(page, cfg, on_config_changed=lambda: scan_ui.scan())
    schedule_tab = ScheduleTabUI(page)

    def on_tab_changed(e: ft.ControlEvent):
        # タブ順: 0:アプリ一覧, 1:スケジュール一覧, 2:プログラム, 3:設定
        if e.control.selected_index == 1:
            # スケジュール一覧に切替時のみリフレッシュ
            try:
                schedule_tab.refresh()
            except Exception:
                pass
        if e.control.selected_index == 2:
            # プログラムタブに切替時のみ、エイリアスと探索を更新
            try:
                alias_ui.refresh()
            except Exception:
                pass
            try:
                scan_ui.scan()
            except Exception:
                pass
        # タブ位置を保存
        settings.set_last_tab(cfg, e.control.selected_index)

    tabs = ft.Tabs(
        expand=True,
        selected_index=0,
        animation_duration=450,
        tabs=[
            ft.Tab(text="アプリ一覧", content=scan_ui.view()),
            ft.Tab(text="スケジュール一覧", content=schedule_tab.view()),
            ft.Tab(text="プログラム", content=alias_ui.view()),
            ft.Tab(text="設定", content=settings_ui.view()),
        ],
        on_change=on_tab_changed,
    )

    # + 押下でプレフィル＆タブ遷移するコールバックを接続
    def go_to_alias(exe_path: str, alias_name: str):
        alias_ui.prefill(exe_path, alias_name)
        # プログラムタブへ遷移（タブ順: 0:アプリ一覧, 1:スケジュール一覧, 2:プログラム, 3:設定）
        tabs.selected_index = 2
        page.update()

    scan_ui.on_request_prefill = go_to_alias

    # Slight top/left padding to avoid clipping; ensure container expands so inner ListView scrolls
    page.add(ft.Container(content=tabs, padding=ft.padding.only(left=8, top=8), expand=True))
    # 初期化
    scan_ui.scan()
    alias_ui.refresh()


def run_app():
    # Serve assets for dev and packaged runs
    ft.app(target=main, assets_dir=_assets_dir())
