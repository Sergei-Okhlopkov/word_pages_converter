import traceback
from threading import Thread
from tkinter import DoubleVar, StringVar, filedialog
from tkinter import ttk

import pythoncom

from shared.themes import STATUS_ICONS
from shared.widgets import OutputPathOptions, WrapLabel
from tools.word_to_pdf_pages.converter import convert_document

TOOL_ID = "word_to_pdf_pages"
TOOL_TITLE = 'Конвертация Word в «PDF в Word\'е»'
TOOL_DESCRIPTION = "Каждая страница исходного документа вставляется в новый DOCX как изображение."


class WordToPdfPagesPanel(ttk.Frame):
    def __init__(self, parent, style: ttk.Style, settings: dict):
        super().__init__(parent)
        self.style = style
        self.settings = settings

        self.input_var = StringVar()
        self.status_var = StringVar(value="Ожидание запуска")
        self.status_icon_var = StringVar(value=STATUS_ICONS["idle"])
        self.progress_var = DoubleVar(value=0.0)
        self.is_busy = False

        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        header = WrapLabel(self, text=TOOL_TITLE, font=("Segoe UI", 14, "bold"))
        header.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        WrapLabel(self, text=TOOL_DESCRIPTION).grid(row=1, column=0, sticky="ew", pady=(0, 16))

        card = ttk.Frame(self, padding=16)
        card.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="Входной Word-файл:").grid(row=0, column=0, columnspan=3, sticky="ew")
        self.input_entry = ttk.Entry(card, textvariable=self.input_var)
        self.input_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.pick_btn = ttk.Button(card, text="Выбрать файл", command=self.pick_input)
        self.pick_btn.grid(row=1, column=2, sticky="e", padx=(8, 0))

        self.run_btn = ttk.Button(card, text="Конвертировать", command=self.on_convert)
        self.run_btn.grid(row=2, column=0, sticky="w", pady=(12, 0))

        self.output_options = OutputPathOptions(card, TOOL_ID, self.settings)

        status_row = ttk.Frame(self)
        status_row.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        status_row.columnconfigure(2, weight=1)

        ttk.Label(status_row, text="Статус").grid(row=0, column=0, sticky="nw", padx=(0, 6))
        self.status_icon_lbl = ttk.Label(status_row, textvariable=self.status_icon_var, font=("Segoe UI Emoji", 12))
        self.status_icon_lbl.grid(row=0, column=1, sticky="nw", padx=(0, 8))
        self.status_lbl = WrapLabel(status_row, textvariable=self.status_var)
        self.status_lbl.grid(row=0, column=2, sticky="ew")

        self.progress = ttk.Progressbar(self, variable=self.progress_var, maximum=100, mode="determinate")
        self.progress.grid(row=4, column=0, sticky="ew", pady=(0, 12))

        legend_card = ttk.Frame(self, padding=12)
        legend_card.grid(row=5, column=0, sticky="nsew")
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
        self.output_options.set_busy(busy)

    def pick_input(self):
        filetypes = [
            (
                "Word and text documents",
                "*.doc *.docx *.docm *.docb *.dot *.dotx *.dotm *.rtf *.odt *.txt *.htm *.html *.mht *.mhtml *.xml",
            )
        ]
        selected = filedialog.askopenfilename(title="Выберите входной файл", filetypes=filetypes)
        if selected:
            self.input_var.set(selected)
            self.set_status("idle", "Файл выбран. Нажмите «Конвертировать».")

    def on_convert(self):
        if self.is_busy:
            return
        input_path = self.input_var.get().strip()
        if not input_path:
            self.set_status("error", "Не выбран входной файл.")
            return

        path_error = self.output_options.validate()
        if path_error:
            self.set_status("error", path_error)
            return

        output_dir = self.output_options.get_output_dir()

        self.progress_var.set(0)
        self.set_busy(True)
        self.set_status("working", "Подготовка конвертации...")

        worker = Thread(target=self._convert_worker, args=(input_path, output_dir), daemon=True)
        worker.start()

    def _ui_progress(self, stage: str, current=None, total=None):
        if stage == "status":
            self.set_status("working", str(current))
            return

        if not total:
            return
        total = max(1, int(total))
        current = int(current)

        if stage == "render":
            percent = 10 + (current / total) * 45
            self.set_status("working", f"Рендер страниц: {current}/{total}")
        elif stage == "insert":
            percent = 55 + (current / total) * 44
            self.set_status("working", f"Сборка DOCX: {current}/{total}")
        else:
            return

        self.progress_var.set(min(99, percent))

    def _convert_worker(self, input_path: str, output_dir):
        def callback(*args):
            self.after(0, self._ui_progress, *args)

        try:
            pythoncom.CoInitialize()
            output = convert_document(input_path, on_progress=callback, output_dir=output_dir)
            self.after(0, self._on_success, str(output))
        except Exception as exc:
            details = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            self.after(0, self._on_error, details)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _on_success(self, output_path: str):
        self.progress_var.set(100)
        self.set_status("success", f"Готово: {output_path}")
        self.set_busy(False)

    def _on_error(self, details: str):
        self.progress_var.set(0)
        self.set_status("error", f"Ошибка: {details}")
        self.set_busy(False)
