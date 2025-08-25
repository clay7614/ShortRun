try:
    # PyInstallerなどで __package__ が未設定の場合は絶対インポートで対応
    from shortrun.gui import run_app as _run
except Exception:
    # パッケージ実行 (python -m shortrun) では相対でも可
    from .gui import run_app as _run

if __name__ == "__main__":
    _run()
