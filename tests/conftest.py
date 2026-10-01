import sys
import types

try:
    import tkinter  # noqa: F401
except ImportError:  # headless Python builds without Tk: stub it so the modules stay importable
    tk = types.ModuleType("tkinter")
    tk.Tk = tk.Frame = tk.Widget = tk.Event = tk.Button = object
    filedialog = types.ModuleType("tkinter.filedialog")
    tk.filedialog = filedialog
    sys.modules["tkinter"] = tk
    sys.modules["tkinter.filedialog"] = filedialog
