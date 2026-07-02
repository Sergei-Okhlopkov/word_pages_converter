import traceback
from threading import Thread
from tkinter import StringVar, filedialog
from tkinter import ttk

from shared.themes import STATUS_ICONS
from shared.widgets import WrapLabel
from tools.project_hyperlinks.processor import process_file

TOOL_ID = "project_hyperlinks"
TOOL_TITLE = "Гиперссылки для проектов"
TOOL_DESCRIPTION = (
    "Загрузите CSV или Excel: в первом столбце (#) со 2-й строки номера "
    "будут преобразованы в гиперссылки на tracker.rddm.team. "
    "Результат всегда сохраняется в Excel (.xlsx)."
)


class ProjectHyperlinksPanel(ttk.Frame):
    def __init__(self, parent, style: ttk.Style):
        super().__init__(parent)
        self.style = style

        self.input_var = StringVar()
        self.status_var = StringVar(value="Ожидание запуска")
        self.status_icon_var = StringVar(value=STATUS_ICONS["idle"])
        self.is_busy = False

        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)

        header = WrapLabel(self, text=TOOL_TITLE, font=("Segoe UI", 14, "bold"))
        header.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        WrapLabel(self, text=TOOL_DESCRIPTION).grid(row=1, column=0, sticky="ew", pady=(0, 16))

        card = ttk.Frame(self, padding=16)
        card.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="Входной файл (CSV или Excel):").grid(row=0, column=0, columnspan=3, sticky="ew")
        self.input_entry = ttk.Entry(card, textvariable=self.input_var)
        self.input_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.pick_btn = ttk.Button(card, text="Выбрать файл", command=self.pick_input)
        self.pick_btn.grid(row=1, column=2, sticky="e", padx=(8, 0))

        self.run_btn = ttk.Button(card, text="Создать гиперссылки", command=self.on_process)
        self.run_btn.grid(row=2, column=0, sticky="w", pady=(12, 0))

        status_row = ttk.Frame(self)
        status_row.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        status_row.columnconfigure(2, weight=1)

        ttk.Label(status_row, text="Статус").grid(row=0, column=0, sticky="nw", padx=(0, 6))
        self.status_icon_lbl = ttk.Label(status_row, textvariable=self.status_icon_var, font=("Segoe UI Emoji", 12))
        self.status_icon_lbl.grid(row=0, column=1, sticky="nw", padx=(0, 8))
        self.status_lbl = WrapLabel(status_row, textvariable=self.status_var)
        self.status_lbl.grid(row=0, column=2, sticky="ew")

        legend_card = ttk.Frame(self, padding=12)
        legend_card.grid(row=4, column=0, sticky="nsew")
        legend_card.columnconfigure(0, weight=1)
        ttk.Label(legend_card, text="Легенда иконок").grid(row=0, column=0, sticky="ew", pady=(0, 8))
        WrapLabel(
            legend_card,
            text="◌ ожидание    ⏳ идёт работа    ✅ успех    ❌ ошибка",
            font=("Segoe UI", 10),
        ).grid(row=1, column=0, sticky="ew")

    def set_status(self, kind: str, text: str):
        self.status_icon_var.set(STATUS_ICONS.get(kind, STATUS_ICONS["idle"]))
        self.status_var.set(text)

    def set_busy(self, busy: bool):
        self.is_busy = busy
        state = "disabled" if busy else "normal"
        self.run_btn.configure(state=state)
        self.pick_btn.configure(state=state)
        self.input_entry.configure(state=state)

    def pick_input(self):
        filetypes = [
            ("CSV и Excel", "*.csv *.xlsx *.xlsm"),
            ("CSV", "*.csv"),
            ("Excel", "*.xlsx *.xlsm"),
        ]
        selected = filedialog.askopenfilename(title="Выберите входной файл", filetypes=filetypes)
        if selected:
            self.input_var.set(selected)
            self.set_status("idle", "Файл выбран. Нажмите «Создать гиперссылки».")

    def on_process(self):
        if self.is_busy:
            return
        input_path = self.input_var.get().strip()
        if not input_path:
            self.set_status("error", "Не выбран входной файл.")
            return

        self.set_busy(True)
        self.set_status("working", "Обработка файла...")

        worker = Thread(target=self._process_worker, args=(input_path,), daemon=True)
        worker.start()

    def _process_worker(self, input_path: str):
        try:
            output = process_file(input_path)
            self.after(0, self._on_success, str(output))
        except Exception as exc:
            details = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            self.after(0, self._on_error, details)

    def _on_success(self, output_path: str):
        self.set_status("success", f"Готово: {output_path}")
        self.set_busy(False)

    def _on_error(self, details: str):
        self.set_status("error", f"Ошибка: {details}")
        self.set_busy(False)
