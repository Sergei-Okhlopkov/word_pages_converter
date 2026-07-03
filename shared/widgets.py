import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import BooleanVar, StringVar, filedialog, ttk

from shared.settings import get_tool_settings, save_tool_output_dir

OUTPUT_PATH_CHECKBOX_TEXT = "Сохранить по выбранному пути"
OUTPUT_PATH_HINT = (
    "Без галочки результат сохраняется рядом с исходным файлом. "
    "С галочкой — в папку, выбранную кнопкой «Выбрать путь»."
)
OUTPUT_PATH_PLACEHOLDER = "Папка не выбрана"


class OutputPathOptions:
    """Checkbox, optional folder picker, and hint for custom output directory."""

    def __init__(self, card: ttk.Frame, tool_id: str, settings: dict, options_row_index: int = 3):
        self.tool_id = tool_id
        self.settings = settings

        tool_settings = get_tool_settings(settings, tool_id)
        self.use_custom_var = BooleanVar(value=False)
        self.output_dir_var = StringVar(value=tool_settings.get("output_dir", ""))

        options_row = ttk.Frame(card)
        options_row.grid(row=options_row_index, column=0, columnspan=3, sticky="ew", pady=(8, 0))

        self.checkbox = ttk.Checkbutton(
            options_row,
            text=OUTPUT_PATH_CHECKBOX_TEXT,
            variable=self.use_custom_var,
            command=self._on_toggle,
        )
        self.checkbox.grid(row=0, column=0, sticky="w")

        self.pick_dir_btn = ttk.Button(options_row, text="Выбрать путь", command=self.pick_output_dir)
        self.pick_dir_btn.grid(row=0, column=1, sticky="w", padx=(12, 0))

        self.path_entry = ttk.Entry(card, state="readonly")
        self.path_entry.grid(row=options_row_index + 1, column=0, columnspan=3, sticky="ew", pady=(8, 0))

        self.hint = WrapLabel(card, text=OUTPUT_PATH_HINT, style="Muted.TLabel", font=("Segoe UI", 9))
        self.hint.grid(row=options_row_index + 2, column=0, columnspan=3, sticky="ew", pady=(6, 0))

        self._sync_path_display()
        self._sync_controls()

    def _on_toggle(self):
        self._sync_controls()

    def _sync_controls(self):
        enabled = self.use_custom_var.get()
        self.pick_dir_btn.configure(state="normal" if enabled else "disabled")

    def _sync_path_display(self):
        if self.output_dir_var.get().strip():
            self.path_entry.configure(state="normal")
            self.path_entry.delete(0, "end")
            self.path_entry.insert(0, self.output_dir_var.get().strip())
            self.path_entry.configure(state="readonly")
        else:
            self.path_entry.configure(state="normal")
            self.path_entry.delete(0, "end")
            self.path_entry.insert(0, OUTPUT_PATH_PLACEHOLDER)
            self.path_entry.configure(state="readonly")

    def pick_output_dir(self):
        initial = self.output_dir_var.get().strip() or None
        selected = filedialog.askdirectory(title="Выберите папку для сохранения", initialdir=initial)
        if selected:
            self.output_dir_var.set(selected)
            self._sync_path_display()
            self._persist()

    def get_output_dir(self) -> Path | None:
        if not self.use_custom_var.get():
            return None
        path = self.output_dir_var.get().strip()
        if not path or path == OUTPUT_PATH_PLACEHOLDER:
            return None
        return Path(path)

    def validate(self) -> str | None:
        if self.use_custom_var.get() and self.get_output_dir() is None:
            return "Не выбрана папка для сохранения."
        return None

    def set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        self.checkbox.configure(state=state)
        if busy or not self.use_custom_var.get():
            self.pick_dir_btn.configure(state="disabled")
        else:
            self.pick_dir_btn.configure(state="normal")

    def _persist(self):
        output_dir = self.output_dir_var.get().strip()
        if output_dir == OUTPUT_PATH_PLACEHOLDER:
            output_dir = ""
        save_tool_output_dir(self.settings, self.tool_id, output_dir)


class WrapLabel(ttk.Label):
    """Label that wraps text to the current widget width."""

    def __init__(self, master=None, wrap_padding: int = 2, **kwargs):
        self._wrap_padding = wrap_padding
        kwargs.setdefault("justify", "left")
        super().__init__(master, **kwargs)
        self.bind("<Configure>", self._update_wraplength, add=True)

    def _update_wraplength(self, event=None):
        width = self.winfo_width()
        if width > 1:
            wrap = max(1, width - self._wrap_padding)
            if self.cget("wraplength") != wrap:
                self.configure(wraplength=wrap)


class ToolMenuButton(tk.Frame):
    """Sidebar menu item with wrapped label and selectable border."""

    def __init__(self, master, text: str, command: Callable[[], None], wraplength: int = 212, **kwargs):
        super().__init__(master, cursor="hand2", **kwargs)
        self._command = command
        self._active = False
        self._wraplength = wraplength

        self.label = tk.Label(
            self,
            text=text,
            anchor="w",
            justify="left",
            cursor="hand2",
            bd=0,
            padx=10,
            pady=8,
            wraplength=wraplength,
        )
        self.label.pack(fill="x")

        for widget in (self, self.label):
            widget.bind("<Button-1>", self._on_click)
        self.bind("<Configure>", self._update_wraplength, add=True)

    def _on_click(self, _event=None):
        self._command()

    def _update_wraplength(self, event=None):
        width = self.winfo_width()
        if width > 1:
            wrap = max(1, min(self._wraplength, width - 20))
            self.label.configure(wraplength=wrap)

    def apply_theme(self, colors: dict, active: bool) -> None:
        self._active = active

        if active:
            bg = colors["sidebar_active"]
            fg = colors["text"]
            border = colors["accent"]
            border_width = 2
            font = ("Segoe UI", 10, "bold")
        else:
            bg = colors["sidebar"]
            fg = colors["muted"]
            border = colors["sidebar"]
            border_width = 1
            font = ("Segoe UI", 10)

        self.configure(
            bg=bg,
            highlightbackground=border,
            highlightcolor=border,
            highlightthickness=border_width,
        )
        self.label.configure(bg=bg, fg=fg, font=font)
