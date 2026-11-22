"""Native interface for searching and exporting 2GIS organizations."""

import multiprocessing as mp
import re
import sys
import tempfile
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QTimer, Qt, QUrl
from PySide6.QtGui import QColor, QDesktopServices, QFont, QFontDatabase, QIcon, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCompleter, QFileDialog, QFrame, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox, QProgressBar,
    QPushButton, QScrollArea, QStackedWidget, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

from cities import ASSETS, CITIES, LABELS, city_code, search_key
from desktop_backend import History, data_directory, publish_result, run_worker, save_checkpoint, stop_worker


STATUS = {
    "ready": ("Готово к поиску", "#2a374b", "#cad7ec"),
    "running": ("В работе", "#233e5d", "#abd2ff"),
    "connecting": ("Подключение", "#233e5d", "#abd2ff"),
    "searching": ("Поиск", "#233e5d", "#abd2ff"),
    "collecting": ("Сбор карточек", "#204741", "#9df2d2"),
    "saving": ("Сохранение", "#204741", "#9df2d2"),
    "stopping": ("Отменяется", "#523c27", "#ffda9d"),
    "completed": ("Готово", "#204741", "#9df2d2"),
    "partial": ("Частичная выгрузка", "#523c27", "#ffda9d"),
    "cancelled": ("Отменено", "#2a374b", "#cad7ec"),
    "empty": ("Нет результатов", "#2a374b", "#cad7ec"),
    "interrupted": ("Прервано", "#523c27", "#ffda9d"),
    "error": ("Ошибка", "#512c3e", "#ffc9d4"),
}


def file_name(query):
    label = re.sub(r'[^\w -]', ' ', query).strip()[:55].rstrip() or "карточки"
    return f"2gis-{label}-{datetime.now():%Y%m%d-%H%M%S}.xlsx"


def label(text="", style="", *, wrap=False):
    widget = QLabel(text)
    widget.setTextFormat(Qt.TextFormat.PlainText)
    widget.setObjectName(style)
    widget.setWordWrap(wrap)
    return widget


def card(style="card"):
    widget = QFrame()
    widget.setObjectName(style)
    layout = QVBoxLayout(widget)
    layout.setContentsMargins(20, 17, 20, 17)
    layout.setSpacing(12)
    return widget, layout


def button(text, action, style=""):
    widget = QPushButton(text)
    widget.setObjectName(style)
    widget.setCursor(Qt.CursorShape.PointingHandCursor)
    widget.clicked.connect(action)
    return widget


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


class ExportApp(QMainWindow):
    def __init__(self, history=None):
        super().__init__()
        self.history = history or History(data_directory() / "history.sqlite3")
        self.context = mp.get_context("spawn")
        self.job = None
        self.close_after_job = False
        self.latest_file = None
        self.last_folder = Path.home() / "Documents"
        self.setWindowTitle("2ГИС Экспорт")
        self.setWindowIcon(QIcon(str(ASSETS / "app.png")))
        self.resize(1040, 690)
        self.setMinimumSize(760, 600)
        self.setStyleSheet((ASSETS / "ui.qss").read_text(encoding="utf-8"))
        self.poller = QTimer(self, interval=100)
        self.poller.timeout.connect(self.poll)
        self.create_widgets()
        self.refresh_history()
        self.status("ready", "Введите запрос и укажите место сохранения.")
        self.query_input.setFocus()

    def create_widgets(self):
        root = QWidget(objectName="workspace")
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        sidebar = QWidget(objectName="sidebar")
        sidebar.setFixedWidth(210)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(16, 29, 16, 20)
        side.setSpacing(8)
        mark = label()
        mark.setPixmap(QPixmap(str(ASSETS / "app.png")).scaled(55, 55,
                       Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        side.addWidget(mark)
        side.addWidget(label("2ГИС Экспорт", "brand"))
        side.addSpacing(17)
        self.new_tab = button("Новая выгрузка", lambda: self.navigate(0), "nav")
        self.history_tab = button("История запросов", lambda: self.navigate(1), "nav")
        for tab in (self.new_tab, self.history_tab):
            tab.setCheckable(True)
            side.addWidget(tab)
        side.addStretch()
        layout.addWidget(sidebar)
        self.pages = QStackedWidget()
        for make_page in (self.create_search, self.create_history):
            page = QWidget(objectName="page")
            panel = QVBoxLayout(page)
            panel.setContentsMargins(32, 26, 32, 27)
            panel.setSpacing(16)
            make_page(panel)
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setWidget(page)
            self.pages.addWidget(scroll)
        layout.addWidget(self.pages, 1)
        self.setCentralWidget(root)
        self.navigate(0)

    def navigate(self, index):
        self.pages.setCurrentIndex(index)
        self.new_tab.setChecked(index == 0)
        self.history_tab.setChecked(index == 1)
        if index == 1:
            self.refresh_history()

    def create_search(self, layout):
        layout.addWidget(label("Новая выгрузка", "heading"))
        layout.addWidget(label("Найдите организации в 2ГИС и сохраните результаты.", "caption"))

        form, fields = card()
        fields.addWidget(label("Поиск организаций", "sectionTitle"))
        fields.addWidget(label("Запрос", "fieldLabel"))
        self.query_input = QLineEdit(placeholderText="Например, кафе или студии йоги")
        self.query_input.setAccessibleName("Поисковый запрос")
        self.query_input.returnPressed.connect(self.start)
        fields.addWidget(self.query_input)
        options = QHBoxLayout()
        options.setSpacing(12)
        cities = QVBoxLayout()
        cities.setSpacing(5)
        cities.addWidget(label("Город", "fieldLabel"))
        self.city_input = QLineEdit("Москва", placeholderText="Начните вводить название города")
        self.city_input.setAccessibleName("Город 2ГИС")
        self.city_input.textEdited.connect(self.suggest_cities)
        self.completer = QCompleter(list(CITIES), self.city_input)
        self.completer.setCompletionMode(QCompleter.CompletionMode.UnfilteredPopupCompletion)
        self.completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.completer.setMaxVisibleItems(9)
        self.city_input.setCompleter(self.completer)
        city_row = QHBoxLayout()
        city_row.setSpacing(6)
        city_row.addWidget(self.city_input, 1)
        city_dropdown = button("▾", self.show_all_cities)
        city_dropdown.setAccessibleName("Показать список городов")
        city_dropdown.setFixedWidth(40)
        city_row.addWidget(city_dropdown)
        cities.addLayout(city_row)
        cities.addWidget(label(f"{len(CITIES)} городов и регионов · можно вставить ссылку 2ГИС", "fieldHint", wrap=True))
        options.addLayout(cities, 3)
        limits = QVBoxLayout()
        limits.setSpacing(5)
        limits.addWidget(label("Лимит карточек", "fieldLabel"))
        self.limit_input = QLineEdit(placeholderText="Без лимита")
        self.limit_input.setAccessibleName("Лимит карточек")
        limits.addWidget(self.limit_input)
        limits.addStretch()
        options.addLayout(limits, 2)
        fields.addLayout(options)
        controls = QHBoxLayout()
        self.start_button = button("Найти и выгрузить", self.start, "primary")
        self.cancel_button = button("Отменить", self.confirm_cancel)
        self.cancel_button.setEnabled(False)
        controls.addWidget(self.start_button)
        controls.addWidget(self.cancel_button)
        controls.addStretch()
        fields.addLayout(controls)
        layout.addWidget(form)

        progress, result = card()
        top = QHBoxLayout()
        top.addWidget(label("Ход работы", "sectionTitle"))
        top.addStretch()
        self.badge = label("", "badge")
        top.addWidget(self.badge)
        result.addLayout(top)
        self.detail = label("", "muted", wrap=True)
        result.addWidget(self.detail)
        self.bar = QProgressBar(textVisible=False, minimum=0, maximum=1, value=0)
        self.bar.setFixedHeight(10)
        result.addWidget(self.bar)
        metrics = QHBoxLayout()
        metrics.setSpacing(8)
        self.metric_values = []
        for name in ("Страница", "Найдено", "Собрано"):
            tile, tile_layout = card("metricCard")
            tile_layout.setContentsMargins(11, 6, 11, 8)
            tile_layout.setSpacing(1)
            value = label("—", "metric")
            self.metric_values.append(value)
            tile_layout.addWidget(value)
            tile_layout.addWidget(label(name, "metricName"))
            metrics.addWidget(tile, 1)
        result.addLayout(metrics)
        self.open_button = button("Открыть выгрузку", self.open_latest)
        self.open_button.hide()
        result.addWidget(self.open_button, 0, Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(progress)
        layout.addStretch()

    def suggest_cities(self, text):
        search = search_key(text)
        options = [city for city in CITIES if search in search_key(city)]
        if search:
            code = next((slug for word, slug in (("питер", "spb"), ("спб", "spb"), ("екб", "ekaterinburg"))
                         if word.startswith(search)), None)
            if code and LABELS.get(code) and LABELS[code] not in options:
                options.insert(0, LABELS[code])
        model = self.completer.model()
        model.setStringList(options)
        self.completer.setCompletionPrefix("")
        if options:
            self.completer.complete()
        else:
            self.completer.popup().hide()

    def show_all_cities(self):
        self.city_input.setFocus()
        self.suggest_cities("")

    def create_history(self, layout):
        layout.addWidget(label("История запросов", "heading"))
        layout.addWidget(label("Откройте готовый файл или повторите запрос с прежними параметрами.", "caption", wrap=True))
        container, box = card()
        self.tree = QTableWidget(0, 5)
        self.tree.setHorizontalHeaderLabels(("Запрос", "Город", "Дата и время", "Статус", "Карточки"))
        self.tree.verticalHeader().hide()
        self.tree.setAlternatingRowColors(True)
        self.tree.setShowGrid(False)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.tree.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tree.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tree.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for index, width in enumerate((118, 120, 123, 72), 1):
            self.tree.setColumnWidth(index, width)
        self.tree.verticalHeader().setDefaultSectionSize(39)
        self.tree.setMinimumHeight(220)
        self.tree.itemSelectionChanged.connect(self.history_selected)
        self.tree.itemDoubleClicked.connect(lambda _: self.open_history())
        box.addWidget(self.tree, 1)
        self.history_detail = label("", "muted", wrap=True)
        box.addWidget(self.history_detail)
        controls = QHBoxLayout()
        self.repeat_button = button("Повторить запрос", self.repeat)
        self.history_open_button = button("Открыть файл", self.open_history)
        controls.addWidget(self.repeat_button)
        controls.addWidget(self.history_open_button)
        controls.addStretch()
        box.addLayout(controls)
        layout.addWidget(container, 1)

    def status(self, key, message):
        title, background, foreground = STATUS[key]
        self.badge.setText(title)
        self.badge.setStyleSheet(f"background: {background}; color: {foreground};")
        self.detail.setText(message)

    def refresh_history(self):
        selected = self.selected()
        self.jobs = self.history.list()
        self.tree.setRowCount(len(self.jobs))
        for index, row in enumerate(self.jobs):
            date = datetime.fromisoformat(row["created"]).astimezone().strftime("%d.%m.%Y %H:%M")
            title, background, foreground = STATUS[row["status"]]
            values = (row["query"], LABELS.get(row["city"], row["city"]), date, title, row["rows"])
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 3:
                    item.setBackground(QColor(background))
                    item.setForeground(QColor(foreground))
                if column == 4:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tree.setItem(index, column, item)
            if selected and selected["id"] == row["id"]:
                self.tree.selectRow(index)
        self.history_selected()

    def selected(self):
        index = self.tree.currentRow()
        jobs = getattr(self, "jobs", ())
        return jobs[index] if 0 <= index < len(jobs) else None

    def history_selected(self):
        row = self.selected()
        self.repeat_button.setEnabled(bool(row and not self.job))
        self.history_open_button.setEnabled(bool(row and row["file"]))
        self.history_detail.setText((row["message"] or row["file"] or "") if row else
                                    ("Выберите запрос." if self.jobs else "Запросы ещё не запускались."))

    def repeat(self):
        row = self.selected()
        if row and not self.job:
            self.query_input.setText(row["query"])
            self.city_input.setText(LABELS.get(row["city"], row["city"]))
            self.limit_input.setText(str(row["max_cards"] or ""))
            self.navigate(0)
            self.query_input.setFocus()

    def start(self):
        if self.job:
            return
        try:
            query = self.query_input.text().strip()
            if not query:
                raise ValueError("Введите поисковый запрос.")
            city = city_code(self.city_input.text())
            raw = self.limit_input.text().strip()
            if raw and (not raw.isdecimal() or not 1 <= int(raw) <= 1_048_575):
                raise ValueError("Введите целое число карточек от 1 до 1 048 575.")
            limit = int(raw) if raw else None
        except ValueError as error:
            QMessageBox.warning(self, "Проверьте запрос", str(error))
            return
        folder = self.last_folder if self.last_folder.is_dir() else Path.home()
        destination, kind = QFileDialog.getSaveFileName(
            self, "Сохранить выгрузку", str(folder / file_name(query)), "Книга Excel (*.xlsx);;CSV (*.csv)"
        )
        if destination:
            target = Path(destination)
            if not target.suffix:
                target = target.with_suffix(".csv" if kind.startswith("CSV") else ".xlsx")
            self.start_export(query, city, limit, target)

    def start_export(self, query, city, limit, target):
        if target.exists() or target.suffix.casefold() not in {".xlsx", ".csv"}:
            QMessageBox.warning(self, "Файл выгрузки", "Выберите новый файл с расширением .xlsx или .csv.")
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
        self.start_button.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.open_button.hide()
        self.set_metrics("—", "0", "0")
        self.status("connecting", "Запускаю фоновый Chrome. При первом запуске может понадобиться загрузка драйвера.")
        self.bar.setRange(0, 0)
        self.refresh_history()
        self.poller.start()

    def poll(self):
        job = self.job
        if not job:
            return
        self.read_events(job)
        if job.stop_at is not None and time.monotonic() - job.stop_at > 30 and job.process.is_alive():
            stop_worker(job.process)
        if not job.process.is_alive():
            job.process.join()
            self.read_events(job)
            self.poller.stop()
            self.finish(job)

    def read_events(self, job):
        try:
            while job.receive.poll():
                event, data = job.receive.recv()
                self.handle_event(job, event, data)
        except (EOFError, OSError):
            pass

    def set_metrics(self, page, found, collected):
        for widget, value in zip(self.metric_values, (page, found, collected)):
            widget.setText(str(value))

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
            self.status("saving", "Сохраняю собранные карточки.")
        elif event == "warning":
            self.detail.setText(job.warning)
        page = f"{job.page} / {job.pages}" if job.page and job.pages else job.page or "—"
        self.set_metrics(page, job.found, len(job.rows))
        total = min(job.total, job.limit) if job.total and job.limit else job.total or job.limit
        if isinstance(total, int) and total > 0:
            self.bar.setRange(0, total)
            self.bar.setValue(min(len(job.rows), total))

    def confirm_cancel(self, close_after=False):
        if not self.job or self.job.stop_at is not None:
            return
        response = QMessageBox.question(
            self, "Отменить выгрузку?", "Сбор остановится. Уже собранные карточки будут сохранены.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No,
        )
        if response == QMessageBox.StandardButton.Yes:
            self.close_after_job = close_after
            self.job.stop_at = time.monotonic()
            self.job.stopped.set()
            self.cancel_button.setEnabled(False)
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
        self.start_button.setEnabled(True)
        self.cancel_button.setEnabled(False)
        self.open_button.setVisible(bool(file))
        self.bar.setRange(0, 1)
        self.bar.setValue(1 if state == "completed" else 0)
        self.status(state, detail)
        self.refresh_history()
        if self.close_after_job:
            self.close()

    def open_latest(self):
        if self.latest_file:
            self.open(self.latest_file)

    def open_history(self):
        row = self.selected()
        if row and row["file"]:
            self.open(Path(row["file"]))

    def open(self, file):
        if not file.is_file():
            QMessageBox.warning(self, "Не удалось открыть файл", "Файл был перемещён или удалён.")
        elif not QDesktopServices.openUrl(QUrl.fromLocalFile(str(file))):
            QMessageBox.warning(self, "Не удалось открыть файл", str(file))

    def closeEvent(self, event):
        if self.job:
            event.ignore()
            self.confirm_cancel(close_after=True)
        else:
            event.accept()


def interface_font():
    registered = QFontDatabase.addApplicationFont(str(ASSETS / "Inter.ttf"))
    families = QFontDatabase.applicationFontFamilies(registered) if registered >= 0 else []
    fallback = "Segoe UI" if sys.platform == "win32" else "Helvetica Neue"
    return QFont(families[0] if families else fallback, 10)


def main():
    if sys.platform == "win32":
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("jaFuart.2GISExport")
    application = QApplication(sys.argv)
    application.setStyle("Fusion")
    application.setApplicationName("2ГИС Экспорт")
    application.setWindowIcon(QIcon(str(ASSETS / "app.png")))
    application.setFont(interface_font())
    window = ExportApp()
    window.show()
    return application.exec()


if __name__ == "__main__":
    mp.freeze_support()
    sys.exit(main())
