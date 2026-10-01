"""Tk-Oberflächen: Dateiauswahl und Sortierdialog für Treffpunkte."""
import tkinter as tk
from contextlib import contextmanager
from tkinter import filedialog, messagebox, ttk

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


def get_part_names():
    """Fragt die Namen der Teillisten ab; gibt None zurück, wenn der Dialog abgebrochen wurde."""
    root = tk.Tk()
    app = PartNamesApp(root)
    root.mainloop()
    root.destroy()
    return app.result


def get_part_assignments(parts, meeting_points):
    """
    Lässt den Nutzer jedem Treffpunkt eine Teilliste zuordnen.
    Gibt {Teilliste: [Treffpunkte]} zurück oder None, wenn der Dialog abgebrochen wurde.
    """
    root = tk.Tk()
    app = AssignmentApp(root, parts, meeting_points)
    root.mainloop()
    root.destroy()
    return app.result


def _style_window(root, title, width=WINDOW_WIDTH, height=WINDOW_HEIGHT):
    root.title(title)
    root.minsize(360, 400)
    root.configure(bg=BG)
    x = (root.winfo_screenwidth() - width) // 2
    y = (root.winfo_screenheight() - height) // 2
    root.geometry(f"{width}x{height}+{x}+{y}")


def _header(root, title, subtitle):
    header = tk.Frame(root, bg=ACCENT, pady=18)
    header.pack(fill=tk.X)
    tk.Label(header, text=title, font=("Helvetica", 15, "bold"), bg=ACCENT, fg="white").pack()
    tk.Label(header, text=subtitle, font=("Helvetica", 9), bg=ACCENT, fg="#C7CCFA").pack(pady=(2, 0))


def _done_button(parent, command):
    btn = tk.Button(
        parent, text="✓   Fertig", font=("Helvetica", 12, "bold"),
        bg=SUCCESS, fg="black", activebackground=SUCCESS_DK, activeforeground="white",
        relief=tk.FLAT, cursor="hand2", padx=24, pady=10, command=command,
    )
    btn.bind("<Enter>", lambda _: btn.configure(bg=SUCCESS_DK))
    btn.bind("<Leave>", lambda _: btn.configure(bg=SUCCESS))
    return btn


class PartNamesApp:
    """Tkinter-Fenster zum Eingeben der Namen der Teillisten (z.B. Nord, Südwest)."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.names: list[str] = []
        self.result = None

        _style_window(root, "Teillisten anlegen")
        _header(root, "Namen der Teillisten", "Name eingeben · Enter oder Hinzufügen")

        content = tk.Frame(root, bg=BG, padx=20, pady=16)
        content.pack(fill=tk.BOTH, expand=True)

        entry_row = tk.Frame(content, bg=BG)
        entry_row.pack(fill=tk.X)
        self.entry = tk.Entry(entry_row, font=("Helvetica", 12), relief=tk.FLAT,
                              highlightthickness=1, highlightbackground=BORDER)
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)
        tk.Button(entry_row, text="Hinzufügen", font=("Helvetica", 10), bg=BTN_BG, fg=TEXT,
                  relief=tk.FLAT, cursor="hand2", padx=14, pady=7,
                  command=self._add).pack(side=tk.LEFT, padx=(6, 0))

        self.listbox = tk.Listbox(content, font=("Helvetica", 11), bg=PANEL_BG, fg=TEXT,
                                  selectbackground=SELECTED_BG, selectforeground=TEXT,
                                  activestyle="none", relief=tk.FLAT, highlightthickness=1,
                                  highlightbackground=BORDER)
        self.listbox.pack(fill=tk.BOTH, expand=True, pady=12)

        buttons = tk.Frame(content, bg=BG)
        buttons.pack(fill=tk.X, pady=(0, 12))
        tk.Button(buttons, text="Ausgewählte entfernen", font=("Helvetica", 10), bg=BTN_BG, fg=TEXT,
                  relief=tk.FLAT, cursor="hand2", padx=14, pady=7,
                  command=self._remove).pack(side=tk.LEFT)

        _done_button(content, self._on_done).pack(fill=tk.X)

        self.entry.bind("<Return>", lambda _: self._add())
        self.listbox.bind("<BackSpace>", lambda _: self._remove())
        self.entry.focus_set()

    def _add(self):
        name = self.entry.get().strip()
        if not name:
            return
        if name.lower() in (n.lower() for n in self.names):
            messagebox.showwarning("Doppelter Name", f"'{name}' gibt es schon.")
            return
        self.names.append(name)
        self.listbox.insert(tk.END, f"  {name}")
        self.entry.delete(0, tk.END)

    def _remove(self):
        for idx in reversed(self.listbox.curselection()):
            self.listbox.delete(idx)
            del self.names[idx]

    def _on_done(self):
        if not self.names:
            messagebox.showwarning("Keine Teillisten", "Bitte lege mindestens eine Teilliste an.")
            return
        self.result = list(self.names)
        self.root.quit()


class AssignmentApp:
    """Tkinter-Fenster, in dem jedem Treffpunkt per Auswahlliste eine Teilliste zugeordnet wird."""

    def __init__(self, root: tk.Tk, parts: list[str], meeting_points: list[str]):
        self.root = root
        self.meeting_points = list(meeting_points)
        self.result = None

        _style_window(root, "Treffpunkte zuordnen", height=640)
        _header(root, "Treffpunkte zuordnen", "Wähle für jeden Treffpunkt die passende Teilliste")

        content = tk.Frame(root, bg=BG, padx=20, pady=16)
        content.pack(fill=tk.BOTH, expand=True)
        _done_button(content, self._on_done).pack(side=tk.BOTTOM, fill=tk.X, pady=(12, 0))

        canvas = tk.Canvas(content, bg=BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(content, orient=tk.VERTICAL, command=canvas.yview)
        rows = tk.Frame(canvas, bg=BG)
        rows.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=rows, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.choices = []
        for row, tp in enumerate(self.meeting_points):
            tk.Label(rows, text=tp or "(leer)", font=("Helvetica", 11), bg=BG, fg=TEXT,
                     anchor="w", width=24).grid(row=row, column=0, sticky="w", pady=3)
            choice = tk.StringVar()
            ttk.Combobox(rows, textvariable=choice, values=parts, state="readonly",
                         width=18).grid(row=row, column=1, padx=(8, 0), pady=3)
            self.choices.append(choice)

    def _on_done(self):
        missing = [tp or "(leer)" for tp, choice in zip(self.meeting_points, self.choices) if not choice.get()]
        if missing:
            messagebox.showwarning("Zuordnung unvollständig",
                                   "Diese Treffpunkte haben noch keine Teilliste:\n\n" + "\n".join(missing))
            return
        result = {}
        for tp, choice in zip(self.meeting_points, self.choices):
            result.setdefault(choice.get(), []).append(tp)
        self.result = result
        self.root.quit()


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
