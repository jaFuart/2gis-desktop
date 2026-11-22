"""Local job history and the browser worker for the native application."""

import csv
import os
import shutil
import sqlite3
import sys
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime
from pathlib import Path

import psutil

import parse_2gis


def data_directory():
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    return base / "2GIS Export"


class History:
    def __init__(self, path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(path)
        self.db.row_factory = sqlite3.Row
        self.db.execute("PRAGMA journal_mode=WAL")
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                id INTEGER PRIMARY KEY, created TEXT NOT NULL, query TEXT NOT NULL,
                city TEXT NOT NULL, max_cards INTEGER, target TEXT NOT NULL,
                status TEXT NOT NULL, rows INTEGER NOT NULL DEFAULT 0,
                file TEXT, message TEXT NOT NULL DEFAULT '', owner INTEGER NOT NULL,
                owner_started REAL NOT NULL
            )
        """)
        running = self.db.execute(
            "SELECT id, owner, owner_started FROM jobs WHERE status = 'running'"
        ).fetchall()
        for job in running:
            try:
                alive = abs(psutil.Process(job["owner"]).create_time() - job["owner_started"]) < 0.01
            except psutil.NoSuchProcess:
                alive = False
            except psutil.AccessDenied:
                alive = True
            if not alive:
                self.finish(job["id"], "interrupted", 0, None, "Приложение завершилось во время сбора.")

    def start(self, query, city, limit, target):
        with self.db:
            return self.db.execute("""
                INSERT INTO jobs (created, query, city, max_cards, target, status, owner, owner_started)
                VALUES (?, ?, ?, ?, ?, 'running', ?, ?)
            """, (datetime.now().astimezone().isoformat(timespec="seconds"), query, city, limit,
                  str(target), os.getpid(), psutil.Process().create_time())).lastrowid

    def finish(self, job, status, rows, file, message):
        with self.db:
            self.db.execute("UPDATE jobs SET status=?, rows=?, file=?, message=? WHERE id=?",
                            (status, rows, str(file) if file else None, message, job))

    def list(self):
        return self.db.execute("SELECT * FROM jobs ORDER BY id DESC").fetchall()


def run_worker(command, send, stopped):
    rows = 0
    warning = ""

    def progress(event, data):
        nonlocal rows, warning
        if event == "row":
            rows = data["count"]
        if event == "warning":
            warning = data["message"]
        send.send((event, data))

    try:
        with open(os.devnull, "w", encoding="utf-8") as quiet, redirect_stdout(quiet), redirect_stderr(quiet):
            if stopped.is_set():
                send.send(("finish", {"status": "cancelled", "count": 0, "message": ""}))
                return
            from seleniumbase.core.browser_launcher import override_driver_dir

            cache = data_directory() / "drivers"
            cache.mkdir(parents=True, exist_ok=True)
            override_driver_dir(str(cache))
            os.chdir(Path(command[command.index("--output") + 1]).parent)
            rc = parse_2gis.main(command, on_progress=progress, is_cancelled=stopped.is_set)
        if rc == 130:
            status = "cancelled"
        elif rc == 2:
            status = "empty"
        else:
            status = "completed" if rc == 0 else ("partial" if rows else "error")
        send.send(("finish", {"status": status, "count": rows, "message": warning}))
    except Exception as error:
        send.send(("finish", {"status": "error", "count": rows, "message": str(error)}))
    finally:
        send.close()


def stop_worker(process):
    """Stop descendants of this worker only if graceful cancellation timed out."""
    try:
        owner = psutil.Process(process.pid)
        children = owner.children(recursive=True)
        owner.kill()
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        children = []
        if process.is_alive():
            process.kill()
    for child in reversed(children):
        try:
            child.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass


def publish_result(staged, target):
    with staged.open("rb") as source, target.open("xb") as destination:
        try:
            shutil.copyfileobj(source, destination)
        except OSError:
            destination.close()
            target.unlink(missing_ok=True)
            raise


def save_checkpoint(path, rows):
    values = [[parse_2gis.clean(row.get(column)) for column in parse_2gis.COLUMNS] for row in rows]
    if path.suffix == ".xlsx":
        parse_2gis.write_xlsx(path, parse_2gis.COLUMNS, values)
    else:
        with path.open("x", encoding="utf-8-sig", newline="") as file:
            writer = csv.writer(file, delimiter=";")
            writer.writerow(parse_2gis.COLUMNS)
            writer.writerows(values)
