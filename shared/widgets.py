import tkinter as tk
from collections.abc import Callable
from tkinter import ttk


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
            wrap = max(1, min(self._wraplength, width - 24))
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
