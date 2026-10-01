import sys
import types

try:
    import tkinter  # noqa: F401
except ImportError:  # headless Python builds without Tk: stub it so the modules stay importable
    tk = types.ModuleType("tkinter")
    tk.Tk = tk.Frame = tk.Widget = tk.Event = tk.Button = object
    for submodule in ("filedialog", "messagebox", "ttk"):
        stub = types.ModuleType(f"tkinter.{submodule}")
        setattr(tk, submodule, stub)
        sys.modules[f"tkinter.{submodule}"] = stub
    sys.modules["tkinter"] = tk
