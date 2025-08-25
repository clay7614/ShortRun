from __future__ import annotations
import os
import re
import subprocess
from typing import List, Optional, Callable, Tuple

import flet as ft

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


def _show_error(page: ft.Page, message: str):
    page.snack_bar = ft.SnackBar(ft.Text(message), bgcolor=ft.colors.RED_300)
    page.snack_bar.open = True
    page.update()


def _show_info(page: ft.Page, message: str):
    page.snack_bar = ft.SnackBar(ft.Text(message))
    page.snack_bar.open = True
    page.update()


class AliasTabUI:
    def __init__(self, page: ft.Page):
        self.page = page
        self.alias_list = ft.ListView(expand=True, spacing=4, padding=8)
        self.add_alias_field = ft.TextField(label="プログラム名", width=220)
        self.add_path_field = ft.TextField(label="EXE パス", expand=True)
        self.file_picker = ft.FilePicker(on_result=self._on_file_picked)
        # FilePicker は overlay に追加するのが推奨
        if self.file_picker not in self.page.overlay:
            self.page.overlay.append(self.file_picker)
        self.open_file_btn = ft.IconButton(ft.icons.FOLDER_OPEN, tooltip="EXE を選択", on_click=self._pick_file)
        self.add_btn = ft.ElevatedButton(text="追加", icon=ft.icons.ADD, on_click=self._add_alias)
        self.refresh_btn = ft.IconButton(ft.icons.REFRESH, tooltip="更新", on_click=lambda e: self.refresh())

        # スケジュール設定 UI
        self.logon_switch = ft.Switch(label="ログオン時に自動起動（選択中のエイリアス）", value=False, on_change=self._on_toggle_logon)
        self.daily_time = ft.TextField(label="毎日 起動時刻 (HH:MM)", width=160)
        self.daily_btn = ft.OutlinedButton("毎日スケジュール追加", on_click=self._on_add_daily)
        self.once_date = ft.TextField(label="1回のみ 日付 (YYYY/MM/DD)", width=160)
        self.once_time = ft.TextField(label="時刻 (HH:MM)", width=120)
        self.once_btn = ft.OutlinedButton("1回のみスケジュール追加", on_click=self._on_add_once)
        self.current_alias: Optional[str] = None
        self.current_path: Optional[str] = None

        self._view = ft.Container(
            content=ft.Column([
                ft.Row([
                    self.add_alias_field,
                    self.add_path_field,
                    self.open_file_btn,
                    self.add_btn,
                    self.refresh_btn,
                ]),
                ft.Divider(),
                ft.Text("スケジュール設定", weight=ft.FontWeight.BOLD),
                ft.Row([self.logon_switch]),
                ft.Row([self.daily_time, self.daily_btn]),
                ft.Row([self.once_date, self.once_time, self.once_btn]),
                ft.Divider(),
                self.alias_list,
            ], expand=True, spacing=10),
            padding=ft.padding.only(top=12, left=12, right=12, bottom=8),
        )

    def view(self) -> ft.Control:
        return self._view

    def prefill(self, exe_path: str, alias_name: Optional[str] = None):
        """アプリ探索からの呼び出しで入力欄を自動セット"""
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
        self.alias_list.controls.clear()
        entries = registry.list_aliases()
        if not entries:
            self.alias_list.controls.append(ft.Text("登録されたプログラム名はありません。", color=ft.colors.GREY))
        else:
            for ent in sorted(entries, key=lambda x: x.alias.lower()):
                self.alias_list.controls.append(self._alias_row(ent))
        self.page.update()

    def _refresh_schedule_toggle(self):
        if self.current_alias:
            tasks = scheduler.list_tasks(self.current_alias)
            # LOGON タスクが存在するかでスイッチ状態を決定
            self.logon_switch.value = any(t['SimpleName'].endswith('_LOGON') for t in tasks)
        else:
            self.logon_switch.value = False
        self.page.update()

    def _alias_row(self, ent: registry.AliasEntry) -> ft.Control:
        def select_for_schedule(_: ft.ControlEvent):
            # スケジュール対象として選択
            self.current_alias = ent.alias
            self.current_path = ent.exe_path
            self._refresh_schedule_toggle()
            _show_info(self.page, f"スケジュール対象: {ent.alias}")

        def show_tasks(_: ft.ControlEvent):
            tasks = scheduler.list_tasks(ent.alias)
            lines = [f"- {t['SimpleName']} | 次回: {t['NextRunTime']} | {t['Schedule']}" for t in tasks] or ["(なし)"]
            dlg = ft.AlertDialog(title=ft.Text(f"スケジュール: {ent.alias}"), content=ft.Column([ft.Text(l) for l in lines]))
            self.page.dialog = dlg
            dlg.open = True
            self.page.update()

        def delete_all_tasks(_: ft.ControlEvent):
            try:
                scheduler.delete_all_for_alias(ent.alias)
                _show_info(self.page, f"スケジュールを削除しました: {ent.alias}")
                self._refresh_schedule_toggle()
            except Exception as ex:
                _show_error(self.page, f"削除に失敗: {ex}")

        return ft.Container(
            content=ft.Row([
                ft.Text(ent.alias, width=180, weight=ft.FontWeight.BOLD),
                ft.Text(ent.exe_path, expand=True, selectable=True),
                ft.IconButton(ft.icons.SCHEDULE, tooltip="スケジュールを表示", on_click=show_tasks),
                ft.IconButton(ft.icons.CHECK_CIRCLE, tooltip="スケジュール対象に選択", on_click=select_for_schedule),
                ft.IconButton(ft.icons.DELETE_FOREVER, tooltip="スケジュールを全削除", on_click=delete_all_tasks),
                ft.IconButton(ft.icons.PLAY_ARROW, tooltip="起動テスト", on_click=lambda e, p=ent.exe_path: self._launch(p)),
                ft.IconButton(ft.icons.DELETE, tooltip="削除", on_click=lambda e, a=ent.alias: self._remove(a)),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            padding=6,
        )

    def _launch(self, path: str):
        try:
            subprocess.Popen([path], close_fds=True)
        except Exception as ex:
            _show_error(self.page, f"起動に失敗しました: {ex}")

    def _remove(self, alias: str):
        try:
            registry.remove_alias(alias)
            _show_info(self.page, f"削除しました: {alias}")
            self.refresh()
        except Exception as ex:
            _show_error(self.page, f"削除に失敗: {ex}")

    def _pick_file(self, e: ft.ControlEvent):
        self.file_picker.pick_files(allow_multiple=False, allowed_extensions=["exe"], dialog_title="実行ファイルを選択")

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
            _show_error(self.page, "プログラム名と EXE パスを入力してください。")
            return
        try:
            registry.add_alias(alias, path, overwrite=False)
            _show_info(self.page, f"登録しました: {alias}")
            self.add_alias_field.value = ""
            self.add_path_field.value = ""
            self.refresh()
        except FileExistsError:
            def do_overwrite(_: ft.ControlEvent):
                try:
                    registry.add_alias(alias, path, overwrite=True)
                    _show_info(self.page, f"上書きしました: {alias}")
                    self.page.close(dialog)
                    self.refresh()
                except Exception as ex:
                    _show_error(self.page, f"上書きに失敗: {ex}")
            dialog = ft.AlertDialog(
                title=ft.Text("既存のプログラム名"),
                content=ft.Text("同名のプログラム名が存在します。上書きしますか？"),
                actions=[
                    ft.TextButton("キャンセル", on_click=lambda _: self.page.close(dialog)),
                    ft.TextButton("上書き", on_click=do_overwrite),
                ],
            )
            self.page.dialog = dialog
            dialog.open = True
            self.page.update()
        except Exception as ex:
            _show_error(self.page, f"登録に失敗: {ex}")

    def _on_toggle_logon(self, e: ft.ControlEvent):
        if not self.current_alias or not self.current_path:
            _show_error(self.page, "スケジュール対象のエイリアスを一覧から選択してください")
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
            _show_error(self.page, "スケジュール対象のエイリアスを一覧から選択してください")
            return
        try:
            scheduler.create_daily_task(self.current_alias, self.current_path, (self.daily_time.value or '').strip())
            _show_info(self.page, "毎日スケジュールを追加しました")
        except Exception as ex:
            _show_error(self.page, f"追加に失敗: {ex}")

    def _on_add_once(self, e: ft.ControlEvent):
        if not self.current_alias or not self.current_path:
            _show_error(self.page, "スケジュール対象のエイリアスを一覧から選択してください")
            return
        try:
            scheduler.create_once_task(
                self.current_alias,
                self.current_path,
                (self.once_date.value or '').strip(),
                (self.once_time.value or '').strip(),
            )
            _show_info(self.page, "1回のみスケジュールを追加しました")
        except Exception as ex:
            _show_error(self.page, f"追加に失敗: {ex}")


class ScanTabUI:
    def __init__(self, page: ft.Page, on_alias_added: Optional[Callable[[], None]] = None, on_request_prefill: Optional[Callable[[str, str], None]] = None):
        self.page = page
        self.on_alias_added = on_alias_added
        self.on_request_prefill = on_request_prefill
        self.items: List[scanner.AppCandidate] = []
        self.filter_field = ft.TextField(hint_text="アプリ名でフィルタ", expand=True, on_change=lambda e: self._render_list())
        self.scan_btn = ft.ElevatedButton("再スキャン", icon=ft.icons.SEARCH, on_click=lambda e: self.scan())
        self.bulk_add_btn = ft.ElevatedButton("選択を一括追加", icon=ft.icons.ADD_TASK, on_click=lambda e: self._bulk_add())
        self.list_view = ft.ListView(expand=True, spacing=4, padding=8)
        self._view = ft.Container(
            content=ft.Column([
                ft.Row([self.filter_field, self.scan_btn, self.bulk_add_btn]),
                ft.Divider(),
                self.list_view,
            ], expand=True),
            padding=ft.padding.only(top=12, left=12, right=12, bottom=8),
        )
        # 選択状態
        self._selected: set[str] = set()

    def view(self) -> ft.Control:
        return self._view

    def scan(self):
        self.list_view.controls.clear()
        self.list_view.controls.append(ft.Row([ft.ProgressRing(), ft.Text("スキャン中...")]))
        self.page.update()
        # スキャン（同期）
        items = scanner.scan_all()
        self.items = items
        self._render_list()

    def _render_list(self):
        q = (self.filter_field.value or "").strip().lower()
        self.list_view.controls.clear()
        matched = [i for i in self.items if (q in i.name.lower() or q in os.path.basename(i.exe_path).lower())]
        if not matched:
            self.list_view.controls.append(ft.Text("該当なし", color=ft.colors.GREY))
        else:
            for it in sorted(matched, key=lambda x: x.name.lower()):
                self.list_view.controls.append(self._row(it))
        self.page.update()

    def _row(self, it: scanner.AppCandidate) -> ft.Control:
        def toggle_selected(e: ft.ControlEvent):
            key = os.path.normcase(os.path.abspath(it.exe_path))
            if key in self._selected:
                self._selected.remove(key)
                cb.value = False
            else:
                self._selected.add(key)
                cb.value = True
            self.page.update()

        def create_alias(_: ft.ControlEvent):
            # + クリックで即時にエイリアスを登録（既存時は上書き確認）
            suggested = _slugify(it.name) if it.name else _slugify(os.path.splitext(os.path.basename(it.exe_path))[0])
            try:
                registry.add_alias(suggested, it.exe_path, overwrite=False)
                _show_info(self.page, f"登録しました: {suggested}")
                if self.on_alias_added:
                    try:
                        self.on_alias_added()
                    except Exception:
                        pass
            except FileExistsError:
                # 上書き確認
                def do_over(_: ft.ControlEvent):
                    try:
                        registry.add_alias(suggested, it.exe_path, overwrite=True)
                        _show_info(self.page, f"上書きしました: {suggested}")
                        if self.on_alias_added:
                            try:
                                self.on_alias_added()
                            except Exception:
                                pass
                        self.page.close(confirm)
                    except Exception as ex:
                        _show_error(self.page, f"上書きに失敗: {ex}")
                confirm = ft.AlertDialog(
                    title=ft.Text("既存のプログラム名"),
                    content=ft.Text(f"{suggested} は既に存在します。上書きしますか？"),
                    actions=[
                        ft.TextButton("キャンセル", on_click=lambda _: self.page.close(confirm)),
                        ft.TextButton("上書き", on_click=do_over),
                    ],
                )
                self.page.dialog = confirm
                confirm.open = True
                self.page.update()
            except Exception as ex:
                _show_error(self.page, f"登録に失敗: {ex}")

        cb = ft.Checkbox(value=False, on_change=toggle_selected)
        return ft.Container(
            content=ft.Row([
                cb,
                ft.Text(it.name, width=240, weight=ft.FontWeight.BOLD),
                ft.Text(it.exe_path, expand=True, selectable=True),
                ft.Text(it.source, width=140, color=ft.colors.GREY),
                ft.IconButton(ft.icons.ADD, tooltip="直接追加", on_click=create_alias),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            padding=6,
        )

    def _bulk_add(self):
        # 選択されたエントリをまとめて追加
        targets = [i for i in self.items if os.path.normcase(os.path.abspath(i.exe_path)) in self._selected]
        if not targets:
            _show_info(self.page, "一括追加する項目をチェックしてください")
            return
        errors: List[Tuple[str, str]] = []
        added = 0
        for it in targets:
            alias = _slugify(it.name) if it.name else _slugify(os.path.splitext(os.path.basename(it.exe_path))[0])
            try:
                registry.add_alias(alias, it.exe_path, overwrite=False)
                added += 1
            except FileExistsError:
                # 既存はスキップ（必要なら上書きオプションを追加可能）
                continue
            except Exception as ex:
                errors.append((alias, str(ex)))
        msg = f"追加: {added}"
        if errors:
            msg += f" / 失敗: {len(errors)}"
        _show_info(self.page, msg)
        if self.on_alias_added:
            try:
                self.on_alias_added()
            except Exception:
                pass


class SettingsTabUI:
    def __init__(self, page: ft.Page, cfg: dict):
        self.page = page
        self.cfg = cfg
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
        # 自動起動
        self.autostart_switch = ft.Switch(
            label="ログイン時に自動起動",
            value=settings.is_autostart_enabled(),
            on_change=self._on_autostart_changed,
        )
        self._view = ft.Container(
            content=ft.Column([
                ft.Text("設定", weight=ft.FontWeight.BOLD, size=18),
                self.theme_dropdown,
                self.autostart_switch,
                ft.Divider(),
                ft.Text("その他"),
                ft.Text("・エイリアスのエクスポート/インポート（今後対応予定）", color=ft.colors.GREY),
            ], expand=False, spacing=12),
            padding=ft.padding.only(top=12, left=12, right=12, bottom=8),
        )

    def view(self) -> ft.Control:
        return self._view

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
        self.page.update()

    def _on_autostart_changed(self, e: ft.ControlEvent):
        try:
            settings.set_autostart(bool(self.autostart_switch.value))
            _show_info(self.page, "自動起動を更新しました")
        except Exception as ex:
            _show_error(self.page, f"自動起動の更新に失敗: {ex}")
            # 失敗したらUIを元に戻す
            self.autostart_switch.value = not bool(self.autostart_switch.value)
            self.page.update()


def main(page: ft.Page):
    page.title = "ShortRun"
    page.window_width = 980
    page.window_height = 680
    page.horizontal_alignment = ft.CrossAxisAlignment.STRETCH

    cfg = settings.load_config()
    # 初期テーマ適用
    if cfg.get("theme") == "light":
        page.theme_mode = ft.ThemeMode.LIGHT
    elif cfg.get("theme") == "dark":
        page.theme_mode = ft.ThemeMode.DARK
    else:
        page.theme_mode = ft.ThemeMode.SYSTEM

    alias_ui = AliasTabUI(page)
    scan_ui = ScanTabUI(page, on_alias_added=alias_ui.refresh)
    settings_ui = SettingsTabUI(page, cfg)

    def on_tab_changed(e: ft.ControlEvent):
        # プログラム名タブに切り替えたら最新化
        if e.control.selected_index == 1:
            try:
                alias_ui.refresh()
            except Exception:
                pass
        # タブ位置を保存
        settings.set_last_tab(cfg, e.control.selected_index)

    tabs = ft.Tabs(
        expand=True,
        selected_index=int(cfg.get("last_tab", 0)),
        tabs=[
            ft.Tab(text="アプリ探索", content=scan_ui.view()),
            ft.Tab(text="プログラム名", content=alias_ui.view()),
            ft.Tab(text="設定", content=settings_ui.view()),
        ],
        on_change=on_tab_changed,
    )

    # + 押下でプレフィル＆タブ遷移するコールバックを接続
    def go_to_alias(exe_path: str, alias_name: str):
        alias_ui.prefill(exe_path, alias_name)
        tabs.selected_index = 1
        page.update()

    scan_ui.on_request_prefill = go_to_alias

    page.add(tabs)
    # 初期化
    scan_ui.scan()
    alias_ui.refresh()


def run_app():
    ft.app(target=main)
