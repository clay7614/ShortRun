# ShortRun

ShortRun は、Windows の「ファイル名を指定して実行」（Win + R）で入力する短い名前から、登録したアプリを起動できるようにするツールです。  
また、ツール内で様々な自動起動の設定を行うことが出来ます。（要管理者権限）  
スタートメニューなどからインストール済みアプリを見つけて、クリック操作だけで登録できます。  

## 一般ユーザー向け

1. 以下リンクまたはReleasesから「ShortRun.exe」をダウンロードして起動します（インストール不要）。
https://github.com/clay7614/ShortRun/releases/latest
2. 画面上部の「再読み込み」を押すと、インストール済みアプリの候補が一覧表示されます。
3. 追加したいアプリを選んで登録します。「プログラム」タブでプログラム名と実行ファイルパスを手入力で追加することもできます。
4. 登録後は Win + R を押して、登録したプログラム名（例: `note`, `edge`）を入力するとアプリが起動します。

主な機能
- プログラム名の追加・編集・削除（ユーザー権限のみで登録）
- スタートメニューやアンインストール情報から EXE 候補を自動収集して一覧化
- スケジュール起動（任意の登録アプリに対し、ログオン時/起動時/毎日/毎分/毎時/毎週/毎月/1回のみ/アイドル時 を設定可能）

ヒント
- **「エラー：アクセスが拒否されました」**と表示される場合、管理者権限で起動してください。
- プログラム名は英数字・ハイフン・アンダースコア等を用いた分かり易い名称がおすすめです。
- UWP/ストアアプリ/ブラウザ内インストールアプリ，等は対象外です。
- ポータブルアプリです。不要になったら，登録したプログラム名はアプリ内の削除機能で消し、exe を削除してください。

## ライセンス
本プロジェクトはリポジトリの `LICENSE` に従います。

---

## 開発者向け

要件
- Windows、Python 3.11 以降推奨
- PowerShell（同梱の Windows 標準で可）

開発実行

```powershell
python -m venv .venv ; .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m shortrun
```

ビルド（PyInstaller）

```powershell
pyinstaller --clean --noconfirm .\ShortRun.spec
```

メモ
- exe アイコンとウィンドウアイコンは `assets/shortrun.ico` を使用しています。
- プログラム名の登録は `HKCU\Software\Microsoft\Windows\CurrentVersion\App Paths\<alias>.exe` に書き込みます。
- アプリ探索はアンインストールレジストリとスタートメニューの .lnk を解析します。
- スケジュール機能は `schtasks` を利用します。サブプロセスは Windows でコンソール非表示（CREATE_NO_WINDOW）で実行しています。
