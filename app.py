"""Native search and export window for 2GIS Export."""

import multiprocessing as mp
import os
import re
import subprocess
import sys
import tempfile
import time
import tkinter as tk
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from urllib.parse import urlsplit

from desktop_backend import History, data_directory, publish_result, run_worker, save_checkpoint, stop_worker


CITIES = {
    "Москва": "moscow", "Санкт-Петербург": "spb", "Казань": "kazan",
    "Екатеринбург": "ekaterinburg", "Новосибирск": "novosibirsk", "Томск": "tomsk",
    "Краснодар": "krasnodar", "Горно-Алтайск": "gornoaltaysk",
}
STATUS = {
    "ready": ("Готово к поиску", "#EDF1F3", "#36454F"),
    "running": ("В работе", "#E5F0FF", "#174675"),
    "connecting": ("Подключение", "#E5F0FF", "#174675"),
    "searching": ("Поиск", "#E5F0FF", "#174675"),
    "collecting": ("Сбор карточек", "#E2F2EF", "#176358"),
    "saving": ("Сохранение", "#E2F2EF", "#176358"),
    "stopping": ("Отменяется", "#FCEED7", "#784800"),
    "completed": ("Готово", "#DDF3E6", "#18583C"),
    "partial": ("Частичная выгрузка", "#FCEED7", "#784800"),
    "cancelled": ("Отменено", "#EDF1F3", "#36454F"),
    "empty": ("Нет результатов", "#EDF1F3", "#36454F"),
    "interrupted": ("Прервано", "#FCEED7", "#784800"),
    "error": ("Ошибка", "#FCE5E5", "#8B2D2D"),
}
BG = "#F2F5F4"
PANEL = "#FFFFFF"
INK = "#172B29"
MUTED = "#506460"
ACCENT = "#186C59"


def city_code(value):
    text = value.strip()
    label = next((name for name in CITIES if name.casefold() == text.casefold()), "")
    if label:
        return CITIES[label]
    parsed = urlsplit(text)
    if parsed.hostname in {"2gis.ru", "www.2gis.ru"}:
        text = parsed.path.strip("/").split("/")[0]
    if not re.fullmatch(r"[a-z0-9_-]+", text.casefold(), flags=re.ASCII):
        raise ValueError("Выберите город или введите его короткий код из ссылки 2ГИС.")
    return text.casefold()


def file_name(query):
    label = re.sub(r'[^\w -]', ' ', query).strip()[:55].rstrip() or "карточки"
    return f"2gis-{label}-{datetime.now():%Y%m%d-%H%M%S}.xlsx"


def open_file(path):
    if sys.platform == "win32":
        os.startfile(str(path))
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])


@dataclass
class Job:
    key: int
    target: Path
    staged: Path
    temp: object
    process: object
    receive: object
    stopped: object
    limit: int | None
    rows: list = field(default_factory=list)
    stop_at: float | None = None
    finish: dict | None = None
    warning: str = ""
    found: int = 0
    page: int = 0
    pages: int | None = None
    total: int | None = None


class ExportApp:
    def __init__(self, root, history=None):
        self.root = root
        self.history = history or History(data_directory() / "history.sqlite3")
        self.context = mp.get_context("spawn")
        self.job = None
        self.close_after_job = False
        self.latest_file = None
        self.last_folder = Path.home() / "Documents"
        self.query = tk.StringVar()
        self.city = tk.StringVar(value="Москва")
        self.limit = tk.StringVar()
        self.detail = tk.StringVar()
        self.metrics = tk.StringVar()
        self.history_detail = tk.StringVar()
        self.activity = False
        root.title("2ГИС Экспорт")
        root.geometry("850x685")
        root.minsize(730, 610)
        root.configure(background=BG)
        root.protocol("WM_DELETE_WINDOW", self.close)
        self.configure_styles()
        self.create_widgets()
        self.refresh_history()
        self.status("ready", "Введите запрос, как в 2ГИС.")
        self.query_input.focus_set()

    def configure_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", font=("TkDefaultFont", 10), foreground=INK)
        style.configure("TButton", padding=(13, 8), background=PANEL, bordercolor="#C9D7D2",
                        lightcolor=PANEL, darkcolor=PANEL)
        style.map("TButton", background=[("active", "#E7F0ED"), ("disabled", "#F0F3F2")])
        style.configure("Primary.TButton", background=ACCENT, foreground="white",
                        bordercolor=ACCENT, lightcolor=ACCENT, darkcolor=ACCENT, padding=(18, 10))
        style.map("Primary.TButton", background=[("disabled", "#DDE9E5"), ("active", "#115745")],
                  foreground=[("disabled", MUTED)])
        style.configure("TEntry", padding=8, fieldbackground=PANEL, bordercolor="#BDD1CA")
        style.configure("TCombobox", padding=8, fieldbackground=PANEL, bordercolor="#BDD1CA")
        style.configure("TNotebook", background=BG, borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure("TNotebook.Tab", padding=(17, 9), background="#E3ECE9")
        style.map("TNotebook.Tab", background=[("selected", PANEL)])
        style.configure("Treeview", rowheight=33, background=PANEL, fieldbackground=PANEL, borderwidth=0)
        style.configure("Treeview.Heading", background="#E7EFEC", foreground=INK, padding=(6, 8))
        style.map("Treeview", background=[("selected", "#DCEDE8")], foreground=[("selected", INK)])
        style.configure("Horizontal.TProgressbar", background=ACCENT, troughcolor="#DFE9E5",
                        thickness=10, bordercolor="#DFE9E5", lightcolor=ACCENT, darkcolor=ACCENT)

    def create_widgets(self):
        header = tk.Frame(self.root, background="#133D34", padx=30, pady=22)
        header.pack(fill="x")
        tk.Label(header, text="2ГИС Экспорт", font=("TkDefaultFont", 23, "bold"),
                 foreground="white", background="#133D34").pack(anchor="w")
        tk.Label(header, text="Поиск организаций и выгрузка в Excel", font=("TkDefaultFont", 10),
                 foreground="#C9E4DD", background="#133D34").pack(anchor="w", pady=(5, 0))

        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill="both", expand=True, padx=26, pady=(18, 24))
        search = tk.Frame(self.tabs, background=PANEL, padx=24, pady=20)
        archive = tk.Frame(self.tabs, background=PANEL, padx=24, pady=20)
        self.tabs.add(search, text="Новая выгрузка")
        self.tabs.add(archive, text="История")
        search.columnconfigure(0, weight=1)
        search.columnconfigure(1, weight=1)

        def label(parent, text, **settings):
            return tk.Label(parent, text=text, background=PANEL, foreground=INK,
                            anchor="w", **settings)

        label(search, "Запрос", font=("TkDefaultFont", 11, "bold")).grid(row=0, column=0, sticky="w")
        self.query_input = ttk.Entry(search, textvariable=self.query)
        self.query_input.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5, 4))
        self.query_input.bind("<Return>", lambda _: self.start())
        label(search, "Например, «Кафе», «Студии йоги» или «Пекарни рядом с вокзалом»",
              font=("TkDefaultFont", 9)).grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 14))
        label(search, "Город").grid(row=3, column=0, sticky="w")
        label(search, "Лимит карточек · необязательно").grid(row=3, column=1, sticky="w", padx=(14, 0))
        self.city_input = ttk.Combobox(search, textvariable=self.city, values=tuple(CITIES))
        self.city_input.grid(row=4, column=0, sticky="ew", pady=(5, 0))
        self.limit_input = ttk.Entry(search, textvariable=self.limit)
        self.limit_input.grid(row=4, column=1, sticky="ew", padx=(14, 0), pady=(5, 0))
        controls = tk.Frame(search, background=PANEL)
        controls.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(16, 19))
        self.start_button = ttk.Button(controls, text="Найти и выгрузить", command=self.start, style="Primary.TButton")
        self.start_button.pack(side="left")
        self.cancel_button = ttk.Button(controls, text="Отменить", command=self.confirm_cancel, state="disabled")
        self.cancel_button.pack(side="left", padx=(8, 0))
        ttk.Separator(search).grid(row=6, column=0, columnspan=2, sticky="ew", pady=(0, 17))
        label(search, "Ход работы", font=("TkDefaultFont", 12, "bold")).grid(row=7, column=0, sticky="w")
        self.badge = tk.Label(search, padx=10, pady=4, font=("TkDefaultFont", 9, "bold"))
        self.badge.grid(row=7, column=1, sticky="e")
        self.bar = ttk.Progressbar(search, mode="determinate", maximum=1)
        self.bar.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(13, 7))
        tk.Label(search, textvariable=self.metrics, background=PANEL, foreground=INK, anchor="w").grid(
            row=9, column=0, columnspan=2, sticky="ew")
        tk.Label(search, textvariable=self.detail, background=PANEL, foreground=MUTED, anchor="w",
                 justify="left", wraplength=650).grid(row=10, column=0, columnspan=2, sticky="ew", pady=(4, 8))
        self.open_button = ttk.Button(search, text="Открыть выгрузку", command=self.open_latest, state="disabled")
        self.open_button.grid(row=11, column=0, sticky="w")

        label(archive, "История запросов", font=("TkDefaultFont", 14, "bold")).pack(anchor="w", pady=(0, 12))
        table = tk.Frame(archive, background=PANEL)
        table.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table, columns=("query", "city", "time", "status", "count"),
                                 show="headings", selectmode="browse")
        headings = (("query", "Запрос", 190), ("city", "Город", 110), ("time", "Дата и время", 126),
                    ("status", "Статус", 132), ("count", "Карточки", 70))
        for column, name, width in headings:
            self.tree.heading(column, text=name)
            self.tree.column(column, width=width, minwidth=62, stretch=column == "query",
                             anchor="center" if column == "count" else "w")
        for status, (_, _, color) in STATUS.items():
            self.tree.tag_configure(status, foreground=color)
        scrollbar = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.history_selected)
        self.tree.bind("<Double-1>", lambda _: self.open_history())
        tk.Label(archive, textvariable=self.history_detail, background=PANEL, foreground=MUTED, anchor="w",
                 wraplength=650, justify="left").pack(fill="x", pady=(10, 6))
        actions = tk.Frame(archive, background=PANEL)
        actions.pack(fill="x")
        self.repeat_button = ttk.Button(actions, text="Повторить запрос", command=self.repeat, state="disabled")
        self.repeat_button.pack(side="left")
        self.history_open_button = ttk.Button(actions, text="Открыть файл", command=self.open_history, state="disabled")
        self.history_open_button.pack(side="left", padx=(8, 0))

    def status(self, key, text):
        label, color, foreground = STATUS[key]
        self.badge.configure(text=label, background=color, foreground=foreground)
        self.detail.set(text)

    def refresh_history(self):
        selected = self.tree.selection()
        self.jobs = {str(row["id"]): row for row in self.history.list()}
        if self.tree.get_children():
            self.tree.delete(*self.tree.get_children())
        names = {code: label for label, code in CITIES.items()}
        for key, row in self.jobs.items():
            date = datetime.fromisoformat(row["created"]).astimezone().strftime("%d.%m.%Y %H:%M")
            self.tree.insert("", "end", iid=key, tags=(row["status"],),
                             values=(row["query"], names.get(row["city"], row["city"]), date,
                                     STATUS[row["status"]][0], row["rows"]))
        if selected and selected[0] in self.jobs:
            self.tree.selection_set(selected[0])
        self.history_selected()

    def selected(self):
        ids = self.tree.selection()
        return self.jobs.get(ids[0]) if ids else None

    def history_selected(self, _event=None):
        row = self.selected()
        self.repeat_button.configure(state="normal" if row and not self.job else "disabled")
        self.history_open_button.configure(state="normal" if row and row["file"] else "disabled")
        self.history_detail.set((row["message"] or row["file"] or "") if row else
                                ("Запросы ещё не запускались." if not self.jobs else "Выберите запрос."))

    def repeat(self):
        row = self.selected()
        if row and not self.job:
            self.query.set(row["query"])
            self.city.set(next((name for name, code in CITIES.items() if code == row["city"]), row["city"]))
            self.limit.set(str(row["max_cards"] or ""))
            self.tabs.select(0)
            self.query_input.focus_set()

    def start(self):
        if self.job:
            return
        try:
            query = self.query.get().strip()
            if not query:
                raise ValueError("Введите поисковый запрос.")
            city = city_code(self.city.get())
            raw = self.limit.get().strip()
            if raw and (not raw.isdecimal() or not 1 <= int(raw) <= 1_048_575):
                raise ValueError("Лимит должен быть целым числом от 1 до 1 048 575.")
            limit = int(raw) if raw else None
        except ValueError as error:
            messagebox.showwarning("Проверьте запрос", str(error), parent=self.root)
            return
        destination = filedialog.asksaveasfilename(
            parent=self.root, title="Сохранить выгрузку", initialfile=file_name(query),
            initialdir=self.last_folder if self.last_folder.is_dir() else Path.home(),
            defaultextension=".xlsx", filetypes=(("Книга Excel", "*.xlsx"), ("CSV", "*.csv")),
        )
        if not destination:
            return
        self.start_export(query, city, limit, Path(destination))

    def start_export(self, query, city, limit, target):
        if target.exists() or target.suffix.casefold() not in {".xlsx", ".csv"}:
            messagebox.showwarning("Файл выгрузки", "Выберите новый файл с расширением .xlsx или .csv.",
                                   parent=self.root)
            return
        self.last_folder = target.parent
        folder = tempfile.TemporaryDirectory(prefix="2gis-export-")
        staged = Path(folder.name) / ("export" + target.suffix.casefold())
        command = ["--uc", "--city", city, "--query", query, "--output", str(staged)]
        if limit:
            command.extend(("--limit", str(limit)))
        receive, send = self.context.Pipe(duplex=False)
        stopped = self.context.Event()
        process = self.context.Process(target=run_worker, args=(command, send, stopped))
        key = self.history.start(query, city, limit, target)
        try:
            process.start()
        except (OSError, RuntimeError) as error:
            receive.close()
            folder.cleanup()
            self.history.finish(key, "error", 0, None, str(error))
            self.status("error", "Не удалось запустить сбор.")
            self.refresh_history()
            return
        finally:
            send.close()
        self.job = Job(key, target, staged, folder, process, receive, stopped, limit)
        self.latest_file = None
        self.start_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.open_button.configure(state="disabled")
        self.metrics.set("")
        self.status("connecting", "Запускаю фоновый Chrome. При первом запуске загрузка драйвера может занять время.")
        self.bar.configure(mode="indeterminate")
        self.bar.start(12)
        self.activity = True
        self.refresh_history()
        self.root.after(100, self.poll)

    def poll(self):
        job = self.job
        if not job:
            return
        self.read_events(job)
        if job.stop_at is not None and time.monotonic() - job.stop_at > 30 and job.process.is_alive():
            stop_worker(job.process)
        if job.process.is_alive():
            self.root.after(100, self.poll)
        else:
            job.process.join()
            self.read_events(job)
            self.finish(job)

    def read_events(self, job):
        try:
            while job.receive.poll():
                event, data = job.receive.recv()
                self.handle_event(job, event, data)
        except (EOFError, OSError):
            pass

    def handle_event(self, job, event, data):
        if event == "finish":
            job.finish = data
            return
        if event == "warning":
            job.warning = data["message"]
        elif event == "page":
            job.page, job.pages = data["page"], data.get("pages")
            job.total = data.get("total")
            job.found += data["found"]
        elif event == "row":
            job.rows.append(data["row"])
        if job.stop_at is not None:
            return
        if event == "searching":
            self.status("searching", f"Открываю страницу {data['page']} выдачи.")
        elif event == "card":
            self.status("collecting", data["name"] or "Читаю карточку организации.")
        elif event == "saving":
            self.status("saving", "Записываю строки в выбранный формат.")
        elif event == "warning":
            self.detail.set(job.warning)
        page = f"Страница {job.page}" if job.page else ""
        if page and job.pages:
            page += f" из {job.pages}"
        self.metrics.set("   ·   ".join(filter(None, (page, f"Найдено {job.found}", f"Собрано {len(job.rows)}"))))
        total = min(job.total, job.limit) if job.total and job.limit else job.total or job.limit
        if isinstance(total, int) and total > 0:
            self.stop_animation()
            self.bar.configure(mode="determinate", maximum=total, value=min(len(job.rows), total))

    def stop_animation(self):
        if self.activity:
            self.bar.stop()
            self.activity = False

    def confirm_cancel(self, close_after=False):
        if not self.job or self.job.stop_at is not None:
            return
        if messagebox.askyesno("Отменить выгрузку?", "Сбор остановится. Уже собранные карточки будут сохранены.",
                               icon="warning", default="no", parent=self.root):
            self.close_after_job = close_after
            self.job.stop_at = time.monotonic()
            self.job.stopped.set()
            self.cancel_button.configure(state="disabled")
            self.status("stopping", "Завершаю текущий запрос и сохраняю собранные карточки.")

    def finish(self, job):
        result = job.finish or {"status": "cancelled" if job.stop_at is not None else "error",
                                "message": job.warning or "Рабочий процесс завершился до окончания сбора."}
        state = result["status"]
        message = result.get("message") or ""
        file = None
        try:
            if (not job.finish or state == "error") and job.rows:
                checkpoint = job.staged.with_stem("checkpoint")
                save_checkpoint(checkpoint, job.rows)
                publish_result(checkpoint, job.target)
                file = job.target
                if state == "error":
                    state = "partial"
            elif job.staged.exists() and state in {"completed", "partial", "cancelled"}:
                publish_result(job.staged, job.target)
                file = job.target
        except (OSError, ValueError) as error:
            state, message = "error", f"Не удалось сохранить файл: {error}"
        finally:
            job.receive.close()
            job.process.close()
            job.temp.cleanup()
        labels = {"completed": "Выгрузка сохранена.", "empty": "Проверьте запрос или выберите другой город.",
                  "cancelled": "Собранная часть выгрузки сохранена." if file else "Сбор остановлен.",
                  "partial": "Сохранены доступные карточки.", "error": "Выгрузка завершилась с ошибкой."}
        detail = message or labels[state]
        self.history.finish(job.key, state, result.get("count", len(job.rows)), file, detail)
        self.job = None
        self.latest_file = file
        self.start_button.configure(state="normal")
        self.cancel_button.configure(state="disabled")
        self.open_button.configure(state="normal" if file else "disabled")
        self.stop_animation()
        self.bar.configure(mode="determinate", maximum=1, value=1 if state == "completed" else 0)
        self.status(state, detail)
        self.refresh_history()
        if self.close_after_job:
            self.root.destroy()

    def open_latest(self):
        if self.latest_file:
            self.open(self.latest_file)

    def open_history(self):
        row = self.selected()
        if row and row["file"]:
            self.open(Path(row["file"]))

    def open(self, file):
        try:
            if not file.is_file():
                raise FileNotFoundError("Файл был перемещён или удалён.")
            open_file(file)
        except OSError as error:
            messagebox.showwarning("Не удалось открыть файл", str(error), parent=self.root)

    def close(self):
        if self.job:
            self.confirm_cancel(close_after=True)
        else:
            self.root.destroy()


def main():
    root = tk.Tk()
    ExportApp(root)
    root.mainloop()


if __name__ == "__main__":
    mp.freeze_support()
    main()
