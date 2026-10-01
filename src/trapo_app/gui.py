"""Tk-Oberflächen: Dateiauswahl und Sortierdialog für Treffpunkte."""
import tkinter as tk
from contextlib import contextmanager
from tkinter import filedialog

# ──────────────────────────────────────────────
#  Farb-Palette
# ──────────────────────────────────────────────
BG = "#F7F7F8"
PANEL_BG = "#FFFFFF"
ACCENT = "#5B6AF0"  # Blau-Violett
SUCCESS = "#22C55E"
SUCCESS_DK = "#16A34A"
TEXT = "#1A1A2E"
BORDER = "#E2E4EE"
SELECTED_BG = "#EEF0FD"
BTN_BG = "#EDEDF5"
BTN_HOVER = "#DDDDF0"

WINDOW_WIDTH, WINDOW_HEIGHT = 440, 560


@contextmanager
def _hidden_root():
    root = tk.Tk()
    root.withdraw()
    try:
        yield root
    finally:
        root.destroy()


def get_file_ui():
    """Lässt den Nutzer eine Word-, CSV- oder Excel-Datei auswählen."""
    with _hidden_root():
        return filedialog.askopenfilename(
            title="Wähle eine Datei aus",
            filetypes=[("Word-,CSV- oder Excel-Datei", "*.docx *.xlsx *.xls *.csv")]
        )


def get_several_files_ui(ending=""):
    """Lässt den Nutzer mehrere Dateien auswählen, ohne `ending` Word- und Exceltabellen."""
    with _hidden_root():
        if ending:
            file_paths = filedialog.askopenfilenames(
                title="Wähle mehrere Dateien aus",
                filetypes=[("Dateien", ending)]
            )
        else:
            file_paths = filedialog.askopenfilenames(
                title="Wähle mehrere Word- oder Exceltabellen aus",
                filetypes=[("Word- und Excel-Dateien", "*.docx *.xlsx *.xls")]
            )
        return list(file_paths)


def get_sorted_tps(tps):
    """Zeigt die Treffpunkte in einem Sortierdialog und gibt die gewählte Reihenfolge zurück."""
    root = tk.Tk()
    app = ReorderableListApp(root, tps)
    root.mainloop()
    root.destroy()
    return app.get_result()


class ReorderableListApp:
    """
    Tkinter-Fenster mit einer umsortierbaren Liste (Drag & Drop, Pfeiltasten, Buttons).

    Verwendung:
        root = tk.Tk()
        app  = ReorderableListApp(root, ["Apfel", "Banane", "Kirsche"])
        root.mainloop()
        print(app.get_result())   # finale Reihenfolge als Liste
    """

    def __init__(self, root: tk.Tk, items: list[str]):
        self.root = root
        self.items = list(items)
        self.result = None
        self._drag_start_idx: int | None = None

        self._configure_window()
        self._build_ui()

    def _configure_window(self):
        self.root.title("Liste sortieren")
        self.root.minsize(360, 400)
        self.root.configure(bg=BG)
        self.root.resizable(True, True)

        # Fenster zentrieren
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - WINDOW_WIDTH) // 2
        y = (self.root.winfo_screenheight() - WINDOW_HEIGHT) // 2
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x}+{y}")

    def _build_ui(self):
        self._build_header()

        content = tk.Frame(self.root, bg=BG, padx=20, pady=16)
        content.pack(fill=tk.BOTH, expand=True)

        self._build_listbox(content)

        btn_row = tk.Frame(content, bg=BG, pady=12)
        btn_row.pack(fill=tk.X)
        self._make_button(btn_row, "↑  Nach oben", self.move_up).pack(side=tk.LEFT, padx=(0, 6))
        self._make_button(btn_row, "↓  Nach unten", self.move_down).pack(side=tk.LEFT)

        done_btn = tk.Button(
            content,
            text="✓   Fertig",
            font=("Helvetica", 12, "bold"),
            bg=SUCCESS, fg="black",
            activebackground=SUCCESS_DK, activeforeground="white",
            relief=tk.FLAT, cursor="hand2",
            padx=24, pady=10,
            command=self._on_done,
        )
        done_btn.pack(fill=tk.X)
        self._add_hover(done_btn, SUCCESS, SUCCESS_DK)

        self._refresh(keep_selection=None)
        self.listbox.bind("<ButtonPress-1>", self._drag_start)
        self.listbox.bind("<B1-Motion>", self._drag_motion)
        self.listbox.bind("<ButtonRelease-1>", self._drag_release)
        self.listbox.bind("<Up>", lambda _: self.move_up())
        self.listbox.bind("<Down>", lambda _: self.move_down())
        self.root.bind("<Return>", lambda _: self._on_done())

    def _build_header(self):
        header = tk.Frame(self.root, bg=ACCENT, pady=18)
        header.pack(fill=tk.X)
        tk.Label(
            header, text="Reihenfolge anpassen",
            font=("Helvetica", 15, "bold"),
            bg=ACCENT, fg="white"
        ).pack()
        tk.Label(
            header,
            text="Drag & Drop · Pfeiltasten · ↑↓ Buttons",
            font=("Helvetica", 9),
            bg=ACCENT, fg="#C7CCFA"
        ).pack(pady=(2, 0))

    def _build_listbox(self, parent):
        border = tk.Frame(parent, bg=BORDER, bd=0)
        border.pack(fill=tk.BOTH, expand=True)
        inner = tk.Frame(border, bg=PANEL_BG, bd=0)
        inner.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        self.listbox = tk.Listbox(
            inner,
            font=("Helvetica", 11),
            bg=PANEL_BG,
            fg=TEXT,
            selectbackground=SELECTED_BG,
            selectforeground=TEXT,
            activestyle="none",
            relief=tk.FLAT,
            highlightthickness=0,
            borderwidth=0,
            selectborderwidth=0,
            cursor="hand2",
        )
        scrollbar = tk.Scrollbar(
            inner, orient=tk.VERTICAL,
            command=self.listbox.yview,
            troughcolor=BG, bg=BORDER,
        )
        self.listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def _make_button(self, parent, label: str, command) -> tk.Button:
        btn = tk.Button(
            parent, text=label,
            font=("Helvetica", 10),
            bg=BTN_BG, fg=TEXT,
            activebackground=BTN_HOVER, activeforeground=TEXT,
            relief=tk.FLAT, cursor="hand2",
            padx=14, pady=7,
            command=command,
        )
        self._add_hover(btn, BTN_BG, BTN_HOVER)
        return btn

    @staticmethod
    def _add_hover(widget: tk.Widget, normal: str, hover: str):
        widget.bind("<Enter>", lambda _: widget.configure(bg=hover))
        widget.bind("<Leave>", lambda _: widget.configure(bg=normal))

    def _refresh(self, keep_selection: int | None):
        """Listbox neu zeichnen, Auswahl ggf. wiederherstellen."""
        self.listbox.delete(0, tk.END)
        for i, item in enumerate(self.items):
            self.listbox.insert(tk.END, f"  {i + 1:>2}.  {item}")

        if keep_selection is not None:
            idx = max(0, min(keep_selection, len(self.items) - 1))
            self.listbox.select_set(idx)
            self.listbox.see(idx)
            self.listbox.activate(idx)

    def _current(self) -> int | None:
        selection = self.listbox.curselection()
        return selection[0] if selection else None

    def _swap(self, i: int, j: int):
        self.items[i], self.items[j] = self.items[j], self.items[i]
        self._refresh(j)

    def move_up(self):
        idx = self._current()
        if idx is not None and idx > 0:
            self._swap(idx, idx - 1)

    def move_down(self):
        idx = self._current()
        if idx is not None and idx < len(self.items) - 1:
            self._swap(idx, idx + 1)

    def _drag_start(self, event: tk.Event):
        self._drag_start_idx = self.listbox.nearest(event.y)
        self.listbox.select_clear(0, tk.END)
        self.listbox.select_set(self._drag_start_idx)

    def _drag_motion(self, event: tk.Event):
        if self._drag_start_idx is None:
            return
        target = self.listbox.nearest(event.y)
        if target != self._drag_start_idx:
            self._swap(self._drag_start_idx, target)
            self._drag_start_idx = target

    def _drag_release(self, _event: tk.Event):
        self._drag_start_idx = None

    def _on_done(self):
        self.result = list(self.items)
        self.root.quit()

    def get_result(self) -> list[str]:
        """Gibt die finale Liste zurück (nach mainloop)."""
        return self.result if self.result is not None else self.items
