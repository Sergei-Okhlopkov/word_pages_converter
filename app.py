from tkinter import Tk, ttk

from shared.settings import load_settings, save_settings
from shared.themes import THEMES
from shared.widgets import ToolMenuButton
from tools.project_hyperlinks.panel import (
    TOOL_ID as HYPERLINKS_ID,
    TOOL_TITLE as HYPERLINKS_TITLE,
    ProjectHyperlinksPanel,
)
from tools.word_to_pdf_pages.panel import (
    TOOL_ID as CONVERTER_ID,
    TOOL_TITLE as CONVERTER_TITLE,
    WordToPdfPagesPanel,
)

TOOLS = [
    {"id": CONVERTER_ID, "title": CONVERTER_TITLE, "panel_cls": WordToPdfPagesPanel},
    {"id": HYPERLINKS_ID, "title": HYPERLINKS_TITLE, "panel_cls": ProjectHyperlinksPanel},
]

SIDEBAR_WIDTH = 260
SIDEBAR_PAD = (10, 12, 12, 12)


class App:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("Pages Converter")
        self.root.geometry("960x640")
        self.root.minsize(960, 640)

        self.settings = load_settings()
        self.theme_name = self.settings.get("theme", "light")
        if self.theme_name not in THEMES:
            self.theme_name = "light"

        self.style = ttk.Style()
        self.tool_buttons: dict[str, ToolMenuButton] = {}
        self.panels: dict[str, ttk.Frame] = {}
        self.active_tool_id: str | None = None

        self._build_ui()
        self.apply_theme()
        self.show_tool(TOOLS[0]["id"])

    def _build_ui(self):
        self.root.columnconfigure(0, weight=0, minsize=SIDEBAR_WIDTH)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.sidebar = ttk.Frame(self.root, padding=SIDEBAR_PAD, style="Sidebar.TFrame")
        self.sidebar.grid(row=0, column=0, sticky="nsew")

        menu_wrap = SIDEBAR_WIDTH - SIDEBAR_PAD[0] - SIDEBAR_PAD[2] - 12

        ttk.Label(self.sidebar, text="Инструменты", font=("Segoe UI", 12, "bold"), style="Sidebar.TLabel").pack(
            anchor="w", pady=(0, 12)
        )

        for tool in TOOLS:
            btn = ToolMenuButton(
                self.sidebar,
                text=tool["title"],
                command=lambda tid=tool["id"]: self.show_tool(tid),
                wraplength=menu_wrap,
            )
            btn.pack(fill="x", pady=4)
            self.tool_buttons[tool["id"]] = btn

        self.theme_btn = ttk.Button(self.sidebar, text="Тема: ...", command=self.toggle_theme)
        self.theme_btn.pack(side="bottom", fill="x", pady=(16, 0))

        self.content = ttk.Frame(self.root, padding=16)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(0, weight=1)

        for tool in TOOLS:
            panel = tool["panel_cls"](self.content, self.style, self.settings)
            panel.grid(row=0, column=0, sticky="nsew")
            self.panels[tool["id"]] = panel

    def show_tool(self, tool_id: str):
        if tool_id not in self.panels:
            return

        for tid, panel in self.panels.items():
            if tid == tool_id:
                panel.tkraise()
            if tid in self.tool_buttons:
                self._set_tool_button_active(tid, tid == tool_id)

        self.active_tool_id = tool_id

    def _set_tool_button_active(self, tool_id: str, active: bool):
        colors = THEMES[self.theme_name]
        self.tool_buttons[tool_id].apply_theme(colors, active=active)

    def apply_theme(self):
        colors = THEMES[self.theme_name]
        self.root.configure(bg=colors["bg"])
        self.style.theme_use("clam")

        self.style.configure(".", background=colors["bg"], foreground=colors["text"])
        self.style.configure("TFrame", background=colors["bg"])
        self.style.configure("Sidebar.TFrame", background=colors["sidebar"], borderwidth=0, relief="flat")
        self.style.configure("TLabel", background=colors["bg"], foreground=colors["text"], font=("Segoe UI", 10))
        self.style.configure("Muted.TLabel", background=colors["bg"], foreground=colors["muted"], font=("Segoe UI", 9))
        self.style.configure("Sidebar.TLabel", background=colors["sidebar"], foreground=colors["text"], font=("Segoe UI", 10))
        self.style.configure(
            "TCheckbutton",
            background=colors["bg"],
            foreground=colors["text"],
            focuscolor=colors["bg"],
            borderwidth=0,
            relief="flat",
            font=("Segoe UI", 10),
        )
        self.style.map(
            "TCheckbutton",
            background=[
                ("active", colors["panel"]),
                ("pressed", colors["panel"]),
                ("disabled", colors["bg"]),
            ],
            foreground=[
                ("active", colors["text"]),
                ("disabled", colors["muted"]),
            ],
            indicatorcolor=[
                ("selected", "#ffffff"),
                ("!selected", colors["muted"]),
            ],
            indicatorbackground=[
                ("selected", colors["accent"]),
                ("active", colors["panel"]),
                ("pressed", colors["panel"]),
                ("!selected", colors["entry_bg"]),
            ],
        )
        self.style.configure(
            "TButton",
            background=colors["accent"],
            foreground="#ffffff",
            padding=(10, 8),
            borderwidth=1,
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )
        self.style.map(
            "TButton",
            background=[("active", colors["accent"])],
            foreground=[("disabled", "#9ca3af"), ("active", "#ffffff")],
        )
        self.style.configure(
            "TEntry",
            fieldbackground=colors["entry_bg"],
            foreground=colors["entry_fg"],
            bordercolor=colors["muted"],
            lightcolor=colors["muted"],
            darkcolor=colors["muted"],
            padding=6,
        )
        self.style.configure(
            "Horizontal.TProgressbar",
            background=colors["accent"],
            troughcolor=colors["panel"],
            bordercolor=colors["panel"],
            lightcolor=colors["accent"],
            darkcolor=colors["accent"],
        )

        self.sidebar.configure(style="Sidebar.TFrame")
        self.content.configure(style="TFrame")
        self.theme_btn.configure(text=f"Тема: {'Тёмная' if self.theme_name == 'dark' else 'Светлая'}")

        if self.active_tool_id:
            colors = THEMES[self.theme_name]
            for tid, btn in self.tool_buttons.items():
                btn.apply_theme(colors, active=(tid == self.active_tool_id))

    def toggle_theme(self):
        self.theme_name = "dark" if self.theme_name == "light" else "light"
        self.settings["theme"] = self.theme_name
        save_settings(self.settings)
        self.apply_theme()


def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
