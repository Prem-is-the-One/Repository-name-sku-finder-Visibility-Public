from __future__ import annotations

import json
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

ROOT = Path(__file__).resolve().parents[1]
SETTINGS_FILE = ROOT / "settings.json"
EXPORT_SCRIPT = ROOT / "tools" / "export_excel.py"


def load_settings():
    if not SETTINGS_FILE.exists():
        raise FileNotFoundError(
            "settings.json is missing. Copy settings.example.json to settings.json first."
        )
    with SETTINGS_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


def run_export():
    print("\nExcel change detected. Rebuilding website data...")
    result = subprocess.run([sys.executable, str(EXPORT_SCRIPT)], cwd=ROOT)
    if result.returncode != 0:
        print("Export failed. The website was not pushed.")
        return False
    return True


def git_push():
    try:
        subprocess.run(["git", "add", "data/data.json", "data/images"], cwd=ROOT, check=True)

        diff = subprocess.run(
            ["git", "diff", "--cached", "--quiet"],
            cwd=ROOT,
        )
        if diff.returncode == 0:
            print("No website data changes to publish.")
            return

        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        subprocess.run(
            ["git", "commit", "-m", f"Auto update Excel data {stamp}"],
            cwd=ROOT,
            check=True,
        )
        subprocess.run(["git", "push"], cwd=ROOT, check=True)
        print("Published to GitHub.")
    except FileNotFoundError:
        print("Git is not installed or is not available in PATH.")
    except subprocess.CalledProcessError as exc:
        print(f"Git publish failed (exit code {exc.returncode}).")


class ExcelHandler(FileSystemEventHandler):
    def __init__(self, excel_file: Path, auto_push: bool):
        self.excel_file = excel_file.resolve()
        self.auto_push = auto_push
        self.last_run = 0.0

    def _is_target(self, event) -> bool:
        try:
            return Path(event.src_path).resolve() == self.excel_file
        except Exception:
            return False

    def on_modified(self, event):
        if event.is_directory or not self._is_target(event):
            return

        now = time.time()
        if now - self.last_run < 2:
            return
        self.last_run = now

        # Excel can fire the file event before the save has fully completed.
        time.sleep(1.5)

        if run_export() and self.auto_push:
            git_push()

    def on_moved(self, event):
        if event.is_directory:
            return
        try:
            dest = Path(event.dest_path).resolve()
        except Exception:
            return
        if dest == self.excel_file:
            time.sleep(1.5)
            if run_export() and self.auto_push:
                git_push()


def main():
    settings = load_settings()
    excel_file = Path(settings["excel_file"]).expanduser()
    auto_push = bool(settings.get("auto_git_push", False))

    if not excel_file.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_file}")

    print(f"Watching: {excel_file}")
    print(f"Auto GitHub push: {'ON' if auto_push else 'OFF'}")
    print("Leave this terminal open while you are editing Excel.")
    print("Press Ctrl+C to stop.")

    # Initial build
    if run_export() and auto_push:
        git_push()

    handler = ExcelHandler(excel_file, auto_push)
    observer = Observer()
    observer.schedule(handler, str(excel_file.parent), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


if __name__ == "__main__":
    main()
