import os
import sys
import json
import zipfile
import shutil
import threading
import logging
import time
import queue
import tkinter.font as tkfont
from concurrent.futures import ThreadPoolExecutor, TimeoutError
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# --- PLATFORM-SPECIFIC CHECK (Windows only) ---
if os.name != 'nt':
    tk.Tk().withdraw()
    messagebox.showerror(
        "Unsupported OS",
        "This app embeds Excel via Win32 and runs only on Windows."
    )
    sys.exit(1)

# --- STABLE DEPENDENCIES (preinstalled per your original) ---
import xlwings as xw
import networkx as nx
import matplotlib
matplotlib.use('Agg')  # render offscreen; we embed canvas separately
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from SPARQLWrapper import SPARQLWrapper, JSON
import win32gui
import win32con
import win32api

APP_NAME = "LEAN Digital Twin (Pro Edition)"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(APP_DIR, "excel_files")
SETTINGS_PATH = os.path.join(APP_DIR, "settings.json")
TAGS_PATH = os.path.join(APP_DIR, "tags.json")
LOG_DIR = os.path.join(APP_DIR, "logs")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)

# --- Logging ---
logging.basicConfig(
    filename=os.path.join(LOG_DIR, "app.log"),
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger("ldt")

# --- Small utilities ---
def safe_json_load(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default

def safe_json_dump(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    os.replace(tmp, path)

def now_str():
    return time.strftime("%Y-%m-%d %H:%M:%S")

# -------------------------
# UI Components
# -------------------------
class Toast(tk.Toplevel):
    """Simple transient toast notification."""
    def __init__(self, parent, text, duration=2000):
        super().__init__(parent)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=text).pack()
        self.update_idletasks()
        x = parent.winfo_rootx() + parent.winfo_width() - self.winfo_width() - 20
        y = parent.winfo_rooty() + parent.winfo_height() - self.winfo_height() - 40
        self.geometry(f"+{x}+{y}")
        self.after(duration, self.destroy)

def show_toast(parent, text, ms=2000):
    try:
        Toast(parent, text, duration=ms)
    except Exception:
        pass

class SelectionDialog(tk.Toplevel):
    """Dialog for selecting multiple items from a list."""
    def __init__(self, parent, title="Select Items", item_list=None, node_type=None):
        super().__init__(parent)
        self.parent = parent
        self.result = None
        self.node_type = node_type
        self.title(title)
        self.geometry("520x440")
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        # center
        self.update_idletasks()
        x = parent.winfo_rootx() + 100
        y = parent.winfo_rooty() + 100
        self.geometry(f"+{x}+{y}")

        self._create_widgets(item_list or [])

    def _create_widgets(self, item_list):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)
        hdr = f"Select {self.node_type.replace('_', ' ').title()}s from Repository:" if self.node_type else "Select items:"
        ttk.Label(main_frame, text=hdr, style='Header.TLabel').pack(anchor="w", pady=(0, 6))

        # filter bar
        filt = ttk.Frame(main_frame)
        filt.pack(fill="x", pady=(0, 8))
        ttk.Label(filt, text="Filter:").pack(side="left", padx=(0, 6))
        self.search_var = tk.StringVar()
        ent = ttk.Entry(filt, textvariable=self.search_var)
        ent.pack(side="left", fill="x", expand=True)
        ent.focus()
        self.search_var.trace_add('write', self._filter_items)

        # list
        self.original_items = sorted(item_list)
        boxfrm = ttk.Frame(main_frame)
        boxfrm.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(boxfrm, selectmode="extended")
        self.listbox.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(boxfrm, command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)
        self._populate_listbox(self.original_items)

        # footer
        self.selection_info = ttk.Label(main_frame, text="0 items selected", style='Info.TLabel')
        self.selection_info.pack(anchor="w", pady=(6, 6))
        self.listbox.bind("<<ListboxSelect>>", self._update_selection_info)

        btns = ttk.Frame(main_frame)
        btns.pack(fill="x")
        ttk.Button(btns, text="Select All", command=self._select_all).pack(side="left")
        ttk.Button(btns, text="Clear", command=self._clear_all).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=self._cancel_clicked).pack(side="right")
        ttk.Button(btns, text="OK", command=self._ok_clicked).pack(side="right", padx=6)

    def _populate_listbox(self, items):
        self.listbox.delete(0, tk.END)
        for item in items:
            self.listbox.insert(tk.END, item)
        self._update_selection_info()

    def _filter_items(self, *_):
        term = self.search_var.get().lower().strip()
        if not term:
            filtered = self.original_items
        else:
            filtered = [i for i in self.original_items if term in i.lower()]
        self._populate_listbox(filtered)

    def _update_selection_info(self, _=None):
        self.selection_info.config(text=f"{len(self.listbox.curselection())} of {self.listbox.size()} selected")

    def _select_all(self):
        self.listbox.select_set(0, tk.END)
        self._update_selection_info()

    def _clear_all(self):
        self.listbox.selection_clear(0, tk.END)
        self._update_selection_info()

    def _ok_clicked(self):
        self.result = [self.listbox.get(i) for i in self.listbox.curselection()]
        self.destroy()

    def _cancel_clicked(self):
        self.result = None
        self.destroy()

# -------------------------
# Main App
# -------------------------
class LeanDigitalTwin(ttk.Frame):
    """
    Lean Digital Twin – Pro Edition
    - SPA-like notebook UI, background SPARQL, Excel embedding.
    """
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.pack(fill="both", expand=True)

        self.initialization_ok = True
        self.settings = safe_json_load(SETTINGS_PATH, {
            "theme": "light",
            "recent_repos": [],
            "repo_url": "",
            "sparql_prefix": "PREFIX ex: <http://example.org/pumps#>"
        })
        self.tag_associations = safe_json_load(TAGS_PATH, {})
        self.executor = ThreadPoolExecutor(max_workers=4)
        self.future_tasks = set()
        self.stop_flag = False

        self.properties = []
        self.current_excel_path = None
        self.graph = nx.DiGraph()
        self.xl_app = None
        self.current_xl_db = None
        self.active_node = None
        self.canvas = None
        self.is_dragging = False
        self.drag_data = None
        self.drag_source_type = None
        self.tree = None
        self.current_item_id = None
        self.fetched_nodes = {'all': [], 'equipment': [], 'sub_equipment': [], 'asset': [], 'plant': [], 'unit': [], 'area': []}
        self.last_preview_ts = None

        # --- shared variables for repo/prefix (fix for geometry mixing) ---
        self.repo_var = tk.StringVar(value=self.settings.get("repo_url", ""))
        self.prefix_var = tk.StringVar(value=self.settings.get("sparql_prefix", "PREFIX ex: <http://example.org/pumps#>"))

        self._configure_styles()
        if not self._probe_excel():
            self.initialization_ok = False
            return
        self._build_ui()
        self._wire_shortcuts()
        self._refresh_file_and_tag_lists()
        # ask after UI is ready
        self.master.after(50, self._show_open_or_new_dialog)

    def _show_open_or_new_dialog(self):
        if getattr(self, "_open_new_shown", False):
            return
        self._open_new_shown = True

        dlg = tk.Toplevel(self.master)
        dlg.title("Welcome")
        dlg.transient(self.master)
        dlg.grab_set()
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Start a project", style="Header.TLabel").pack(anchor="w")
        ttk.Label(frm, text="Open an existing project or create a new one.", style="Info.TLabel").pack(anchor="w",
                                                                                                       pady=(2, 12))

        btns = ttk.Frame(frm)
        btns.pack(fill="x")

        def do_open():
            dlg.destroy()
            self._open_project()

        def do_new():
            dlg.destroy()
            self._create_new_project()

        ttk.Button(btns, text="Open Project…", command=do_open).pack(side="left")
        ttk.Button(btns, text="Create New Project", command=do_new).pack(side="right")

        dlg.update_idletasks()
        x = self.master.winfo_rootx() + (self.master.winfo_width() - dlg.winfo_width()) // 2
        y = self.master.winfo_rooty() + (self.master.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{x}+{y}")

    def _create_new_project(self):
        try:
            self.tag_associations = {}
            self.settings["recent_repos"] = []
            self._persist_state()
            self._refresh_file_and_tag_lists()
            show_toast(self.master, "New project created")
        except Exception as e:
            messagebox.showerror("New Project", f"{e}")