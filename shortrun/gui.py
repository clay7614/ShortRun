from __future__ import annotations
import os
import re
import subprocess
import time
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


_banner_seq = 0


def _show_banner(page: ft.Page, message: str, *, error: bool = False, duration: float = 2.5):
    """ページ上部から短時間スライド表示する通知。
    - error=True の場合は赤系、通常は青系。
    - duration 秒後に自動で閉じる（最新表示を優先）。
    """
    global _banner_seq
    _banner_seq += 1
    seq = _banner_seq

    bgcolor = ft.colors.RED_200 if error else ft.colors.BLUE_200
    icon = ft.icons.ERROR_OUTLINE if error else ft.icons.INFO_OUTLINE

    # 既存のバナーを再利用 or 作成
    if getattr(page, "banner", None) is None:
        page.banner = ft.Banner(
            bgcolor=bgcolor,
            leading=ft.Icon(icon),
            content=ft.Text(message),
            actions=[
                ft.TextButton("閉じる", on_click=lambda e, p=page: (setattr(p.banner, "open", False), p.update())),
            ],
        )
    else:
        page.banner.bgcolor = bgcolor
        page.banner.leading = ft.Icon(icon)
        page.banner.content = ft.Text(message)
        page.banner.actions = [
            ft.TextButton("閉じる", on_click=lambda e, p=page: (setattr(p.banner, "open", False), p.update())),
        ]

    page.banner.open = True
    page.update()

    # 非同期に自動クローズ（別スレッドで待機してから閉じる）
    def _auto_close_sync(s: int, d: float):
        try:
            time.sleep(max(0.5, d))
            if s == _banner_seq and getattr(page, "banner", None) is not None:
                page.banner.open = False
                page.update()
        except Exception:
            pass

    try:
        page.run_task(lambda: _auto_close_sync(seq, duration))
    except Exception:
        # run_task が使えない環境でも例外は無視（自動閉じは諦める）
        pass


def _show_error(page: ft.Page, message: str):
    _show_banner(page, message, error=True)


def _show_info(page: ft.Page, message: str):
    _show_banner(page, message, error=False)


class AliasTabUI:
    def __init__(self, page: ft.Page):
        self.page = page
        self.alias_list = ft.ListView(expand=True, spacing=4, padding=8)
        self.add_alias_field = ft.TextField(label="プログラム名", width=220, tooltip="Win+R で起動する短い名前。英数・-・_ を推奨")
        self.add_path_field = ft.TextField(label="実行ファイルパス", expand=True, tooltip="起動したい実行ファイルのパス")
        self.file_picker = ft.FilePicker(on_result=self._on_file_picked)
        # FilePicker は overlay に追加するのが推奨
        if self.file_picker not in self.page.overlay:
            self.page.overlay.append(self.file_picker)
        self.open_file_btn = ft.IconButton(ft.icons.FOLDER_OPEN, tooltip="実行ファイルを選択")
        self.open_file_btn.on_click = self._pick_file
        self.add_btn = ft.ElevatedButton(text="追加", icon=ft.icons.ADD, on_click=self._add_alias, tooltip="入力中のプログラム名と実行ファイルパスでプログラムを登録")
        self.refresh_btn = ft.IconButton(ft.icons.REFRESH, tooltip="保存したプログラムを再読み込み", on_click=lambda e: self.refresh())

        # 旧スケジュール設定 UIは統合のため削除し、行ごとの「スケジュール」ボタンからダイアログを開く
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
        # 統合後も現在選択中の alias/path は維持
        self.page.update()

    def _alias_row(self, ent: registry.AliasEntry) -> ft.Control:
        def open_schedule_dialog(_: ft.ControlEvent):
            # スケジュール設定をまとめたダイアログ
            alias = ent.alias
            path = ent.exe_path
            # コントロール
            logon_sw = ft.Switch(label="ログオン時に起動", value=False)
            daily_tf = ft.TextField(label="毎日 時刻 (HH:MM)", width=140)
            daily_btn = ft.OutlinedButton("毎日を追加")
            once_date = ft.TextField(label="1回 日付 (YYYY/MM/DD)", width=180)
            once_time = ft.TextField(label="1回 時刻 (HH:MM)", width=140)
            add_once_btn = ft.OutlinedButton("1回を追加")

            def add_daily(e: ft.ControlEvent):
                try:
                    scheduler.create_daily_task(alias, path, (daily_tf.value or '').strip())
                    _show_info(self.page, "「毎日」スケジュールを追加しました")
                except Exception as ex:
                    _show_error(self.page, f"追加に失敗: {ex}")

            def add_once(e: ft.ControlEvent):
                try:
                    scheduler.create_once_task(alias, path, (once_date.value or '').strip(), (once_time.value or '').strip())
                    _show_info(self.page, "「1回のみ」スケジュールを追加しました")
                except Exception as ex:
                    _show_error(self.page, f"追加に失敗: {ex}")

        def delete_all(e: ft.ControlEvent):
                try:
            scheduler.delete_all_for_alias(alias)
            self.page.close(dlg)
            _show_info(self.page, "スケジュールを削除しました")
                except Exception as ex:
                    _show_error(self.page, f"削除に失敗: {ex}")

            daily_btn.on_click = add_daily
            add_once_btn.on_click = add_once

            # 既存一覧（のちほど計算）
            lines: List[str] = ["読み込み中..."]

            content = ft.Container(
                width=860,
                content=ft.Column([
                    logon_sw,
                    ft.Row([daily_tf, daily_btn]),
                    ft.Row([once_date, once_time, add_once_btn]),
                    ft.Divider(),
                    ft.Text("現在のスケジュール"),
                    ft.Column([ft.Text(l) for l in lines], scroll=ft.ScrollMode.AUTO, height=220),
                ], tight=True, spacing=8),
            )

            def save_all(e: ft.ControlEvent):
                errors: List[str] = []
                # ログオン
                try:
                    scheduler.ensure_logon_task(alias, path, bool(logon_sw.value))
                except Exception as ex:
                    errors.append(f"ログオン時: {ex}")
                # 毎日
                hhmm = (daily_tf.value or '').strip()
                if hhmm:
                    try:
                        scheduler.create_daily_task(alias, path, hhmm)
                    except Exception as ex:
                        errors.append(f"毎日: {ex}")
                # 1回
                od = (once_date.value or '').strip()
                ot = (once_time.value or '').strip()
                if od or ot:
                    if od and ot:
                        try:
                            scheduler.create_once_task(alias, path, od, ot)
                        except Exception as ex:
                            errors.append(f"1回のみ: {ex}")
                    else:
                        errors.append("「1回のみ」は日付と時刻を両方入力してください")

                # 一覧再描画
                tasks2 = scheduler.list_tasks(alias)
                logon_sw.value = any(t['SimpleName'].endswith('_LOGON') for t in tasks2)
                new_lines = [f"- {t['SimpleName']} | 次回: {t['NextRunTime']} | {t['Schedule']}" for t in tasks2] or ["(なし)"]
                if isinstance(content.content, ft.Column):
                    content.content.controls[5] = ft.Column([ft.Text(l) for l in new_lines], scroll=ft.ScrollMode.AUTO, height=220)
                self.page.update()

                if errors:
                    _show_error(self.page, "\n".join(errors))
                else:
                    _show_info(self.page, "保存しました")

            dlg = ft.AlertDialog(
                modal=True,
                title=ft.Text(f"スケジュール設定: {alias}"),
                content=content,
                actions=[
                    ft.TextButton("保存", on_click=save_all),
                    ft.TextButton("全削除", on_click=delete_all),
                    ft.TextButton("閉じる", on_click=lambda e: self.page.close(dlg)),
                ],
            )
            self.page.open(dlg)

            # 内容を計算して更新（同期）
            tasks = scheduler.list_tasks(alias)
            logon_sw.value = any(t['SimpleName'].endswith('_LOGON') for t in tasks)
            lines = [f"- {t['SimpleName']} | 次回: {t['NextRunTime']} | {t['Schedule']}" for t in tasks] or ["(なし)"]
            # 最後の Column を置き換え
            if isinstance(content.content, ft.Column):
                col_controls = content.content.controls
                # 0: switch, 1: daily row, 2: once row, 3: divider, 4: title, 5: list
                col_controls[5] = ft.Column([ft.Text(l) for l in lines], scroll=ft.ScrollMode.AUTO, height=220)
            self.page.update()

        return ft.Container(
            content=ft.Row([
                ft.Text(ent.alias, width=180, weight=ft.FontWeight.BOLD),
                ft.Text(ent.exe_path, expand=True, selectable=True),
                ft.IconButton(ft.icons.SCHEDULE, tooltip="スケジュール設定を開く", on_click=open_schedule_dialog),
                ft.IconButton(ft.icons.PLAY_ARROW, tooltip="プログラムの起動", on_click=lambda e, p=ent.exe_path: self._launch(p)),
                ft.IconButton(ft.icons.EDIT, tooltip="プログラム名とパスを編集", on_click=lambda e, entry=ent: self._edit_alias(entry)),
                ft.IconButton(ft.icons.DELETE, tooltip="プログラムを削除", on_click=lambda e, a=ent.alias: self._remove(a)),
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

    def _edit_alias(self, entry: registry.AliasEntry):
        page = self.page
        alias_tf = ft.TextField(label="プログラム名", value=entry.alias, width=220)
        path_tf = ft.TextField(label="実行ファイルパス", value=entry.exe_path, expand=True)

        # ローカル FilePicker（編集用）
        def _on_pick(res: ft.FilePickerResultEvent):
            if res.files:
                path_tf.value = res.files[0].path
                page.update()

        fp = ft.FilePicker(on_result=_on_pick)
        if fp not in page.overlay:
            page.overlay.append(fp)
        pick_btn = ft.IconButton(ft.icons.FOLDER_OPEN, tooltip="実行ファイルを選択", on_click=lambda e: fp.pick_files(allow_multiple=False, allowed_extensions=["exe"], dialog_title="実行ファイルを選択"))

        content = ft.Column([
            ft.Row([alias_tf]),
            ft.Row([path_tf, pick_btn]),
        ], spacing=10)

        def do_save(e: ft.ControlEvent, *, overwrite: bool = False):
            new_alias = (alias_tf.value or "").strip()
            new_path = (path_tf.value or "").strip().strip('"')
            if not new_alias or not new_path:
                _show_error(page, "プログラム名と実行ファイルのパスを入力してください。")
                return
            try:
                registry.update_alias(entry.alias, new_alias, new_path, overwrite=overwrite)
                page.close(dlg)
                _show_info(page, f"更新しました: {entry.alias} → {new_alias}")
                self.refresh()
            except FileExistsError:
                # 上書き確認
                def confirm_over(_: ft.ControlEvent):
                    try:
                        registry.update_alias(entry.alias, new_alias, new_path, overwrite=True)
                        page.close(confirm)
                        page.close(dlg)
                        _show_info(page, f"上書きしました: {new_alias}")
                        self.refresh()
                    except Exception as ex:
                        _show_error(page, f"上書きに失敗: {ex}")
                confirm = ft.AlertDialog(
                    title=ft.Text("既存のプログラム名"),
                    content=ft.Text(f"{new_alias} は既に存在します。上書きしますか？"),
                    actions=[
                        ft.TextButton("キャンセル", on_click=lambda _: page.close(confirm)),
                        ft.TextButton("上書き", on_click=confirm_over),
                    ],
                )
                page.open(confirm)
            except Exception as ex:
                _show_error(page, f"更新に失敗: {ex}")

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
            _show_error(self.page, "プログラム名と実行ファイルのパスを入力してください。")
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
    def __init__(self, page: ft.Page, on_alias_added: Optional[Callable[[], None]] = None, on_request_prefill: Optional[Callable[[str, str], None]] = None):
        self.page = page
        self.on_alias_added = on_alias_added
        self.on_request_prefill = on_request_prefill
        self.items: List[scanner.AppCandidate] = []
        self.filter_field = ft.TextField(hint_text="アプリ名でフィルタ", expand=True, on_change=lambda e: self._render_list(), tooltip="表示中の候補を部分一致で絞り込み")
        self.scan_btn = ft.ElevatedButton("再スキャン", icon=ft.icons.SEARCH, on_click=lambda e: self.scan(), tooltip="アプリの候補一覧を再取得")
        self.bulk_add_btn = ft.ElevatedButton("選択を一括追加", icon=ft.icons.ADD_TASK, on_click=lambda e: self._bulk_add(), tooltip="チェック済みの候補をまとめて登録")
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
                self.page.open(confirm)
            except Exception as ex:
                _show_error(self.page, f"登録に失敗: {ex}")

        cb = ft.Checkbox(value=False, on_change=toggle_selected, tooltip="一括追加の対象として選択/解除")
        return ft.Container(
            content=ft.Row([
                cb,
                ft.Text(it.name, width=240, weight=ft.FontWeight.BOLD),
                ft.Text(it.exe_path, expand=True, selectable=True),
                ft.Text(it.source, width=140, color=ft.colors.GREY),
                ft.IconButton(ft.icons.ADD, tooltip="このアプリを推奨名で直接追加（既存は確認）", on_click=create_alias),
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
        self._view = ft.Container(
            content=ft.Column([
                ft.Text("設定", weight=ft.FontWeight.BOLD, size=18),
                self.theme_dropdown,
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

    # 自動起動設定は削除済み


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
