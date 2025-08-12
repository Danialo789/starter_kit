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
        # Project-scoped paths (we keep app settings global; project content saved via actions)
        self.project_dir = self.settings.get("last_project_dir", APP_DIR)
        self.data_dir = os.path.join(self.project_dir, "excel_files")
        os.makedirs(self.data_dir, exist_ok=True)
        self.tags_path = os.path.join(self.project_dir, "tags.json")
        self.tag_associations = safe_json_load(self.tags_path if os.path.exists(self.tags_path) else TAGS_PATH, {})
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
        self.fetched_nodes = {'all': [], 'equipment': [], 'sub_equipment': [], 'asset': []}
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

    # ---------- lifecycle ----------
    def _probe_excel(self):
        try:
            app_check = xw.App(visible=True, add_book=False)
            app_check.quit()
            return True
        except Exception as e:
            self.master.withdraw()
            messagebox.showerror("Excel Not Found",
                f"Could not connect to Microsoft Excel. Ensure it’s installed.\n\n{e}")
            return False

    def on_close(self):
        try:
            if self.xl_app and self.current_xl_db:
                if not self.current_xl_db.api.Saved:
                    res = messagebox.askyesnocancel("Save Changes?",
                        f"Save changes to '{self.current_xl_db.name}'?", icon='warning')
                    if res is True:
                        self.current_xl_db.save()
                    elif res is None:
                        return
            self._cleanup_xlwings()
            self._persist_state()
        finally:
            self.stop_flag = True
            self.executor.shutdown(wait=False, cancel_futures=True)
            self.master.destroy()

    def _persist_state(self):
        # settings
        self.settings["repo_url"] = self.repo_var.get().strip()
        self.settings["sparql_prefix"] = self.prefix_var.get().strip()
        self.settings["last_project_dir"] = self.project_dir
        safe_json_dump(SETTINGS_PATH, self.settings)
        # project tags
        safe_json_dump(self.tags_path, self.tag_associations)

    # ---------- styles / theme ----------
    def _configure_styles(self):
        self.style = ttk.Style(self.master)
        # Keep builtin themes for portability
        base_theme = 'clam'
        try:
            self.style.theme_use(base_theme)
        except Exception:
            pass

        # palette
        if self.settings.get("theme") == "dark":
            bg = "#22252a"; fg = "#e8eaed"; acc = "#3d7eff"
            panel = "#2b2f36"; border = "#3a3f47"
        else:
            bg = "#f7f7fb"; fg = "#1f2328"; acc = "#2357ff"
            panel = "#ffffff"; border = "#dcdfe4"

        self.master.configure(bg=bg)
        self.style.configure("TFrame", background=panel)
        self.style.configure("TLabel", background=panel, foreground=fg, font=('Segoe UI', 10))
        self.style.configure("TButton", font=('Segoe UI', 10, 'bold'), padding=6)
        self.style.configure("TEntry", fieldbackground="#ffffff")
        self.style.configure("TNotebook", background=panel, tabposition='n')
        self.style.configure("TNotebook.Tab", font=('Segoe UI', 10, 'bold'), padding=[10, 6])
        self.style.configure("Header.TLabel", font=('Segoe UI', 11, 'bold'))
        self.style.configure("Info.TLabel", font=('Segoe UI', 9, 'italic'))
        self.style.configure("Badge.TLabel", font=('Segoe UI', 9, 'bold'), foreground=acc)
        self.style.map("TButton", highlightcolor=[("active", acc)], foreground=[("active", fg)])

    def _toggle_theme(self):
        self.settings["theme"] = "dark" if self.settings.get("theme") != "dark" else "light"
        self._configure_styles()
        show_toast(self.master, f"Theme: {self.settings['theme'].title()}")

    # ---------- UI Build ----------
    def _build_ui(self):
        self.master.title(APP_NAME)
        self.master.geometry("1460x900")
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)

        # Menubar: Project actions
        menubar = tk.Menu(self.master)
        project_menu = tk.Menu(menubar, tearoff=0)
        project_menu.add_command(label="New Project…", command=self._new_project_folder)
        project_menu.add_command(label="Open Project…", command=self._open_project_folder)
        project_menu.add_separator()
        project_menu.add_command(label="Save Project", command=self._save_project_folder)
        project_menu.add_command(label="Save Project As…", command=lambda: self._save_project_folder(force_ask=True))
        project_menu.add_separator()
        project_menu.add_command(label="Export Project Zip…", command=self._export_project)
        project_menu.add_command(label="Import Project Zip…", command=self._import_project)
        project_menu.add_separator()
        project_menu.add_command(label="Exit", command=self.on_close)
        menubar.add_cascade(label="Project", menu=project_menu)
        self.master.config(menu=menubar)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(10, 10))

        # Tabs
        self.tab_graphical = ttk.Frame(self.nb, padding=10)
        self.tab_graphdb = ttk.Frame(self.nb, padding=10)
        self.tab_excel = ttk.Frame(self.nb, padding=10)
        self.tab_functionalities = ttk.Frame(self.nb, padding=10)
        self.tab_asset_hierarchy = ttk.Frame(self.nb, padding=10)
        frame_repo_tools = ttk.Frame(self.nb, padding=10)

        # Make Repository Tools FIRST so it’s easy to find
        self.nb.add(frame_repo_tools, text="Repository")
        self.nb.add(self.tab_graphical, text="Graphical Model")
        self.nb.add(self.tab_graphdb, text="GraphDB → Excel")
        self.nb.add(self.tab_excel, text="Datasheet Editor")
        self.nb.add(self.tab_functionalities, text="Functionalities")
        self.nb.add(self.tab_asset_hierarchy, text="Asset Hierarchy")

        # Build tab contents
        self._build_repository_tools_tab(frame_repo_tools)
        self._build_graphical_model_tab(self.tab_graphical)
        self._build_graphdb_tab(self.tab_graphdb)  # (now without its duplicate Connection group)
        self._build_datasheet_editor_tab(self.tab_excel)
        self._build_functionalities_tab(self.tab_functionalities)
        self._build_plant_hierarchy_tab(self.tab_asset_hierarchy)

        # bottom/status bar
        status_bar = ttk.Frame(self)
        status_bar.pack(fill="x", padx=10, pady=(0, 6))
        self.status_lbl = ttk.Label(status_bar, text="Ready")
        self.status_lbl.pack(side="left")
        self.badge_nodes = ttk.Label(status_bar, text="Nodes: 0 | Eq:0 Sub:0 Asset:0", style="Badge.TLabel")
        self.badge_nodes.pack(side="right", padx=(0, 10))
        self.progress = ttk.Progressbar(status_bar, mode="indeterminate", length=160)

        # bottom buttons bar
        self.bottom = ttk.Frame(self)
        self.bottom.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(self.bottom, text="Save Project…", command=lambda: self._save_project_folder(force_ask=True)).pack(side="left")
        ttk.Button(self.bottom, text="Open Project…", command=self._open_project_folder).pack(side="left", padx=6)
        ttk.Button(self.bottom, text="Export Project…", command=self._export_project).pack(side="left", padx=(12,0))
        ttk.Button(self.bottom, text="Import Project…", command=self._import_project).pack(side="left", padx=6)
        ttk.Button(self.bottom, text="Theme", command=self._toggle_theme).pack(side="right")  # moved here from top
        ttk.Button(self.bottom, text="Save", command=self._safe_save_all).pack(side="right", padx=(0, 6))

    def _wire_shortcuts(self):
        self.master.bind("<Control-s>", lambda e: self._safe_save_all())
        self.master.bind("<Control-S>", lambda e: self._save_project_folder())
        self.master.bind("<Control-o>", lambda e: self._open_project_folder())
        self.master.bind("<F5>", lambda e: self._update_data_model())
        self.master.bind("<Control-f>", lambda e: self._focus_first_filter())

    def _focus_first_filter(self):
        try:
            self.node_display_filter.focus_set()
        except Exception:
            pass

    def _with_progress(self, running=True):
        if running:
            self.progress.pack(side="right")
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.pack_forget()

    def _set_status(self, text, ms=None):
        self.status_lbl.config(text=text)
        if ms:
            self.master.after(ms, lambda: self.status_lbl.config(text="Ready"))

    # ---------- Tabs ----------
    def _build_graphical_model_tab(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)

        ctrl = ttk.Frame(parent)
        ctrl.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(ctrl, text="Select Node(s) to Visualize:", style='Header.TLabel').grid(row=0, column=0, sticky="w")
        filtfrm = ttk.Frame(ctrl)
        filtfrm.grid(row=1, column=0, sticky="ew", pady=(6, 4))
        ttk.Label(filtfrm, text="Filter:").pack(side="left", padx=(0, 6))
        self.node_display_filter_var = tk.StringVar()
        self.node_display_filter = ttk.Entry(filtfrm, textvariable=self.node_display_filter_var)
        self.node_display_filter.pack(side="left", fill="x", expand=True)
        self.node_display_filter_var.trace_add('write', lambda *_: self._filter_listbox(self.node_listbox_display, self.node_display_filter_var.get()))

        self.node_listbox_display = tk.Listbox(ctrl, height=7, selectmode="extended", exportselection=False)
        self.node_listbox_display.grid(row=2, column=0, sticky="ew")
        self.node_listbox_display.bind("<<ListboxSelect>>", lambda e: self._update_data_model())

        btnbar = ttk.Frame(ctrl)
        btnbar.grid(row=3, column=0, sticky="ew", pady=6)
        ttk.Button(btnbar, text="Refresh Model (F5)", command=self._update_data_model).pack(side="left")
        ttk.Button(btnbar, text="Clear Graph", command=self._clear_graph).pack(side="left", padx=6)

        self.graph_frame = ttk.Frame(parent, relief=tk.SUNKEN, borderwidth=1)
        self.graph_frame.grid(row=1, column=0, sticky="nsew")

    def _build_graphdb_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        # Node selection (now the only section here)
        node_select = ttk.LabelFrame(parent, text="Master Node List (Double-click to set Active)", padding=10)
        node_select.grid(row=0, column=0, sticky="nsew")
        node_select.columnconfigure(0, weight=1)
        node_select.rowconfigure(2, weight=1)

        topbar = ttk.Frame(node_select)
        topbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(topbar, text="Fetch All Nodes", command=self._fetch_nodes).pack(side="left")
        ttk.Button(topbar, text="Categorize", command=self._categorize_nodes).pack(side="left", padx=6)

        fl = ttk.Frame(node_select)
        fl.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        ttk.Label(fl, text="Filter:").pack(side="left", padx=(0, 6))
        self.node_filter_var = tk.StringVar()
        ent = ttk.Entry(fl, textvariable=self.node_filter_var)
        ent.pack(side="left", fill="x", expand=True)
        self.node_filter_var.trace_add('write',
                                       lambda *_: self._filter_listbox(self.node_listbox, self.node_filter_var.get()))

        self.node_listbox = tk.Listbox(node_select, height=10, exportselection=False)
        self.node_listbox.grid(row=2, column=0, sticky="nsew")
        self.node_listbox.bind("<Double-1>", lambda e: self._on_node_manual_select())

    def _build_datasheet_editor_tab(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(1, weight=1)

        left = ttk.Frame(parent, padding=5)
        left.grid(row=0, column=0, rowspan=2, sticky="ns")
        left.rowconfigure(3, weight=1)

        ttk.Label(left, text="Select a Tag to View:", style='Header.TLabel').grid(row=0, column=0, sticky="w")
        self.tag_selector_combobox = ttk.Combobox(left, state="readonly")
        self.tag_selector_combobox.grid(row=1, column=0, sticky="ew", pady=(4, 6))
        self.tag_selector_combobox.bind("<<ComboboxSelected>>", lambda e: self._on_tag_selected_in_editor_tab())

        ttk.Label(left, text="Datasheets for Tag:", style='Header.TLabel').grid(row=2, column=0, sticky="w", pady=(6, 2))
        self.datasheet_listbox_for_tag = tk.Listbox(left, height=16, exportselection=False)
        self.datasheet_listbox_for_tag.grid(row=3, column=0, sticky="nsew")
        self.datasheet_listbox_for_tag.bind("<Double-1>", self._load_file_from_list)

        # right pane
        main_pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        main_pane.grid(row=1, column=1, sticky="nsew", padx=10, pady=5)

        table_container = ttk.Frame(main_pane)
        main_pane.add(table_container, weight=3)
        table_container.rowconfigure(1, weight=1)
        table_container.columnconfigure(0, weight=1)

        sheet_controls = ttk.Frame(table_container)
        sheet_controls.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(sheet_controls, text="Activate Sheet:").pack(side=tk.LEFT)
        self.sheet_selector_combobox = ttk.Combobox(sheet_controls, state="readonly", width=28)
        self.sheet_selector_combobox.pack(side=tk.LEFT, padx=6)
        self.sheet_selector_combobox.bind("<<ComboboxSelected>>", lambda e: self._on_sheet_selected())

        ttk.Button(sheet_controls, text="Open Excel Window", command=self._popout_excel).pack(side="left", padx=6)
        ttk.Button(sheet_controls, text="Save As…", command=self._save_excel_as).pack(side="left")

        self.excel_frame = ttk.Frame(table_container, relief=tk.SUNKEN, borderwidth=1)
        self.excel_frame.grid(row=1, column=0, sticky="nsew")
        self.excel_frame.bind("<Configure>", self._resize_excel_window)

        info = ttk.Frame(main_pane, padding=10)
        info.columnconfigure(0, weight=1)
        main_pane.add(info, weight=1)

        ttk.Label(info, text="Tag Information", style='Header.TLabel').grid(row=0, column=0, sticky="w")
        self.tag_text_embed = tk.Text(info, height=4, width=30, font=('Segoe UI', 10), relief=tk.SOLID, borderwidth=1, state='disabled')
        self.tag_text_embed.grid(row=1, column=0, sticky="ew", pady=(4, 8))

        active = ttk.LabelFrame(info, text="Active Node for Mapping", padding=10)
        active.grid(row=2, column=0, sticky="ew")
        self.active_node_display_label = ttk.Label(active, text="None Selected", style='Header.TLabel')
        self.active_node_display_label.pack()

        ttk.Label(info, text="Associated Node Properties (Double-click to preview):", style='Header.TLabel').grid(row=3, column=0, sticky="w", pady=(10, 4))
        self.properties_text_embed = tk.Text(info, height=10, width=30, font=('Segoe UI', 10), relief=tk.SOLID, borderwidth=1, state='disabled')
        self.properties_text_embed.grid(row=4, column=0, sticky="nsew")
        self.properties_text_embed.bind("<Double-1>", self._on_property_double_click)

        paste = ttk.LabelFrame(info, text="Drag-and-Drop Live Data", padding=10)
        paste.grid(row=5, column=0, sticky="nsew", pady=(10, 0))
        paste.columnconfigure(0, weight=1)
        prop_display = ttk.LabelFrame(paste, text="Selected Property")
        prop_display.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.selected_property_display = ttk.Label(prop_display, text="(None)", font=('Segoe UI', 10, 'bold'))
        self.selected_property_display.pack(padx=5, pady=5)
        self.live_value_button = ttk.Button(paste, text="Copy Value: (none)")
        self.live_value_button.value = None
        self.live_value_button.configure(command=lambda: self._initiate_drag_copy(self.live_value_button.value, 'value'))
        self.live_value_button.grid(row=1, column=0, sticky="ew", pady=5)
        self.live_unit_button = ttk.Button(paste, text="Copy Unit: (none)")
        self.live_unit_button.value = None
        self.live_unit_button.configure(command=lambda: self._initiate_drag_copy(self.live_unit_button.value, 'unit'))
        self.live_unit_button.grid(row=2, column=0, sticky="ew", pady=5)
        self.preview_ts_lbl = ttk.Label(paste, text="Last preview: —", style="Info.TLabel")
        self.preview_ts_lbl.grid(row=3, column=0, sticky="w")

    def _init_tree_icons_and_fonts(self):
        """Create tiny 16x16 icons procedurally (no external files) + bold font for buckets."""
        # Fonts
        base = tkfont.nametofont('TkDefaultFont')
        self.font_bold = tkfont.Font(self.master, **base.actual())
        self.font_bold.configure(weight='bold')

        # Simple pixel painter for 16x16 icons
        def make_icon(bg, fg=None, kind="square"):
            img = tk.PhotoImage(width=16, height=16)
            img.put(bg, to=(0, 0, 16, 16))
            if kind == "folder":
                # folder tab
                tab = "#ffffff" if fg is None else fg
                img.put(tab, to=(1, 1, 9, 5))
                # folder body
                body = "#f6d365" if fg is None else fg
                img.put(body, to=(1, 5, 15, 15))
            elif kind == "leaf":  # Plant
                leaf = "#ffffff" if fg is None else fg
                for x in range(3, 13):
                    y = int(8 + 3 * (x - 8) / 8)
                    img.put(leaf, to=(x, y, x + 1, y + 2))
                img.put(leaf, to=(7, 4, 9, 6))
            elif kind == "gear":  # Equipment
                gear = "#ffffff" if fg is None else fg
                img.put(gear, to=(3, 7, 13, 9))
                img.put(gear, to=(7, 3, 9, 13))
            elif kind == "box":  # Asset
                box = "#ffffff" if fg is None else fg
                img.put(box, to=(3, 3, 13, 13))
                img.put(bg, to=(5, 5, 11, 11))
            elif kind == "chip":  # Sub-Equipment
                chip = "#ffffff" if fg is None else fg
                img.put(chip, to=(3, 5, 13, 11))
                # pins
                for x in (2, 13):
                    img.put(chip, to=(x, 6, x + 1, 10))
            elif kind == "layers":  # Unit
                lay = "#ffffff" if fg is None else fg
                img.put(lay, to=(3, 5, 13, 8))
                img.put(lay, to=(4, 9, 12, 12))
            elif kind == "grid":  # Area
                line = "#ffffff" if fg is None else fg
                for x in (5, 10):
                    img.put(line, to=(x, 3, x + 1, 13))
                for y in (6, 10):
                    img.put(line, to=(3, y, 13, y + 1))
            elif kind == "project":
                bar = "#ffffff" if fg is None else fg
                img.put(bar, to=(2, 3, 14, 6))
                img.put(bar, to=(2, 8, 10, 11))
                img.put(bar, to=(2, 12, 8, 14))
            return img

        # Keep references on self so Tk doesn't GC them
        self.icons = {
            "project": make_icon("#4a6cf7", "#dbe3ff", "project"),
            "folder": make_icon("#f0b84d", "#ffe6b3", "folder"),
            "plant": make_icon("#3cb371", "#eaffea", "leaf"),
            "unit": make_icon("#6a9ff0", "#e6f0ff", "layers"),
            "area": make_icon("#f5974e", "#ffe9d8", "grid"),
            "equipment": make_icon("#6b7280", "#e5e7eb", "gear"),
            "sub_equipment": make_icon("#8b5cf6", "#ede9fe", "chip"),
            "asset": make_icon("#10b981", "#d1fae5", "box"),
        }

        # Map entity/bucket tags -> icon key
        self.icon_for_type = {
            'project_node': "project",
            'bucket_plants': "folder",
            'bucket_units': "folder",
            'bucket_areas': "folder",
            'bucket_equipment': "folder",
            'bucket_sub_equipment': "folder",
            'bucket_assets': "folder",
            'plant_node': "plant",
            'unit_node': "unit",
            'area_node': "area",
            'equipment_node': "equipment",
            'sub_equipment_node': "sub_equipment",
            'asset_node': "asset",
        }

    def _insert_bucket(self, parent_iid, label, tag):
        """Create a visible bucket under parent (bold + folder icon), or return existing one."""
        # Reuse existing
        for child in self.tree.get_children(parent_iid):
            if tag in (self.tree.item(child, 'tags') or ()): 
                return child
        icon_key = self.icon_for_type.get(tag, 'folder')
        iid = self.tree.insert(
            parent_iid, "end", text=label, open=True,
            tags=(tag, 'bucket_bold'), image=self.icons[icon_key]
        )
        return iid

    def _ensure_buckets(self, owner_iid, owner_type):
        """Create the right buckets under each entity node so structure is always visible."""
        if owner_type == 'plant_node':
            self._insert_bucket(owner_iid, "Units", "bucket_units")
            self._insert_bucket(owner_iid, "Areas", "bucket_areas")
            self._insert_bucket(owner_iid, "Equipment", "bucket_equipment")
            self._insert_bucket(owner_iid, "Assets", "bucket_assets")
        elif owner_type == 'unit_node':
            self._insert_bucket(owner_iid, "Areas", "bucket_areas")
            self._insert_bucket(owner_iid, "Equipment", "bucket_equipment")
            self._insert_bucket(owner_iid, "Assets", "bucket_assets")
        elif owner_type == 'area_node':
            self._insert_bucket(owner_iid, "Equipment", "bucket_equipment")
            self._insert_bucket(owner_iid, "Assets", "bucket_assets")
        elif owner_type == 'equipment_node':
            self._insert_bucket(owner_iid, "Sub-Equipment", "bucket_sub_equipment")
            self._insert_bucket(owner_iid, "Assets", "bucket_assets")
        elif owner_type == 'sub_equipment_node':
            self._insert_bucket(owner_iid, "Assets", "bucket_assets")

    def _get_node_type(self, iid):
        tags = self.tree.item(iid, "tags")
        return tags[0] if tags else 'project_node'  # treat top as project

    def _get_insert_parent(self):
        sel = self.tree.selection()
        return sel[0] if sel else getattr(self, "project_root", "")

    def _add_child_clicked(self):
        parent_iid = self._get_insert_parent()
        parent_type = self._get_node_type(parent_iid)
        choices = self.allowed_children.get(parent_type, [])
        if not choices:
            return
        m = tk.Menu(self.master, tearoff=0)
        for child_tag in choices:
            m.add_command(label=f"➕ {self.pretty_type.get(child_tag, child_tag)}",
                          command=lambda ct=child_tag: self._create_child(parent_iid, ct))
        x, y = self.master.winfo_pointerx(), self.master.winfo_pointery()
        m.post(x, y)

    def _create_child(self, parent_iid, child_tag):
        # If adding a bucket directly
        if child_tag.startswith("bucket_"):
            self._insert_bucket(parent_iid, self.pretty_type.get(child_tag, child_tag), child_tag)
            return

        # Determine bucket to drop this entity into
        parent_type = self._get_node_type(parent_iid)
        owner_iid = parent_iid
        if parent_type.startswith("bucket_"):
            bucket_iid = parent_iid
            owner_iid = self.tree.parent(parent_iid)
            owner_type = self._get_node_type(owner_iid)
        else:
            owner_type = parent_type
            bucket_for_child = {
                'plant_node': 'bucket_plants',  # only under project
                'unit_node': 'bucket_units',
                'area_node': 'bucket_areas',
                'equipment_node': 'bucket_equipment',
                'sub_equipment_node': 'bucket_sub_equipment',
                'asset_node': 'bucket_assets',
            }
            needed_bucket_tag = bucket_for_child.get(child_tag)
            if not needed_bucket_tag:
                messagebox.showwarning("Not allowed", f"Cannot create {self.pretty_type.get(child_tag)} here.")
                return
            # Ensure buckets exist
            if child_tag != 'plant_node':
                self._ensure_buckets(owner_iid, owner_type)
            # Find/create the bucket
            bucket_iid = None
            for c in self.tree.get_children(owner_iid):
                if needed_bucket_tag in (self.tree.item(c, 'tags') or ()): 
                    bucket_iid = c; 
                    break
            if not bucket_iid:
                bucket_iid = self._insert_bucket(owner_iid, self.pretty_type.get(needed_bucket_tag, needed_bucket_tag),
                                                 needed_bucket_tag)

        pretty = self.pretty_type.get(child_tag, child_tag)
        icon_key = self.icon_for_type.get(child_tag, 'folder')

        # Free-text entities
        if child_tag in ('plant_node', 'unit_node', 'area_node'):
            name = simpledialog.askstring(f"Create {pretty}", f"Enter {pretty} name:")
            if not name: return
            new_iid = self.tree.insert(bucket_iid, "end", text=name, open=True, tags=(child_tag,),
                                       image=self.icons[icon_key])
            self._ensure_buckets(new_iid, child_tag)
            self._set_status(f"Created {pretty}: {name}", 3000)
            return

        # Repo-backed entities
        if not self.fetched_nodes['all']:
            messagebox.showwarning("No Nodes", "Fetch repository nodes first (GraphDB → Excel tab).")
            return

        node_type_map = {
            'equipment_node': 'equipment',
            'sub_equipment_node': 'sub_equipment',
            'asset_node': 'asset',
        }
        repo_key = node_type_map.get(child_tag, 'all')
        repo_nodes = self.fetched_nodes.get(repo_key) or self.fetched_nodes['all']
        if not repo_nodes:
            messagebox.showwarning("No Nodes", f"No {repo_key.replace('_', ' ')} nodes found.")
            return

        titles = {
            'equipment_node': "Select Equipment",
            'sub_equipment_node': "Select Sub-Equipment",
            'asset_node': "Select Assets",
        }
        dlg = SelectionDialog(self.master, title=titles.get(child_tag, "Select Items"),
                              item_list=repo_nodes, node_type=repo_key)
        self.master.wait_window(dlg)
        if not dlg.result: return

        for name in dlg.result:
            nid = self.tree.insert(bucket_iid, "end", text=name, open=True, tags=(child_tag,),
                                   image=self.icons[icon_key])
            self._ensure_buckets(nid, child_tag)
        self._set_status(f"Added {len(dlg.result)} {repo_key.replace('_', ' ')} node(s).", 3000)

    def _on_tree_select(self, _=None):
        """Render a small inline details row directly under the selected node."""
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        self._render_inline_details(iid)

    def _render_inline_details(self, iid):
        """Create/update a single 'inline info' child under iid with some details."""
        # Remove any previous inline info child for this node
        for child in self.tree.get_children(iid):
            if 'inline_info' in (self.tree.item(child, 'tags') or ()): 
                self.tree.delete(child)

        node_type = self._get_node_type(iid)
        node_name = self.tree.item(iid, "text")

        # Placeholder while we fetch
        info_iid = self.tree.insert(iid, "end", text="(loading…)", tags=('inline_info',))
        self.tree.item(iid, open=True)

        # For repo-backed nodes, fetch a quick preview via SPARQL.
        # For free-text nodes, just show the type + name.
        if node_type in ('equipment_node', 'sub_equipment_node', 'asset_node'):
            prefix = self.prefix_var.get().strip()
            # Try: count of literal properties + one example value/unit
            q = f"""{prefix}
            SELECT (COUNT(?p) AS ?propCount) (SAMPLE(?v) AS ?sampleVal) (SAMPLE(?u) AS ?sampleUnit) WHERE {{
                {{ ex:{node_name} ?p ?o . FILTER (isLiteral(?o)) BIND(?o AS ?v) }}
                UNION
                {{ ex:{node_name} ?prop ?b .
                   ?b ex:hasValue ?v .
                   OPTIONAL {{ ?b ex:hasUnit ?u . }} }}
            }}"""

            fut, _ = self._run_sparql_query_bg(q, timeout=20)

            def after(res):
                if isinstance(res, Exception) or not res:
                    text = f"{self.pretty_type.get(node_type)} • {node_name}"
                else:
                    row = res[0]
                    cnt = row.get('propCount', {}).get('value')
                    val = row.get('sampleVal', {}).get('value')
                    unit = row.get('sampleUnit', {}).get('value')
                    if val and unit:
                        sample = f"{val} {unit}"
                    elif val:
                        sample = val
                    else:
                        sample = "—"
                    text = f"{self.pretty_type.get(node_type)} • {node_name}  |  props: {cnt or '0'}  |  sample: {sample}"
                # Update inline node text
                if self.tree.exists(info_iid):
                    self.tree.item(info_iid, text=text)

            self._track_future(fut, after)
        else:
            # Free-text types
            text = f"{self.pretty_type.get(node_type)} • {node_name}"
            self.tree.item(info_iid, text=text)

    def _build_plant_hierarchy_tab(self, parent):
        # Layout: LEFT = Tree (expands), RIGHT = reserved frame (Visio placeholder)
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=2)  # tree expands more
        parent.columnconfigure(1, weight=1)  # placeholder for future Visio

        # Init icons/fonts once we have a Tk root
        self._init_tree_icons_and_fonts()

        # ---- LEFT: hierarchy tree ----
        left = ttk.Frame(parent, padding=(6, 6, 6, 6))
        left.grid(row=0, column=0, sticky="nsew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(left)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        vs.grid(row=0, column=1, sticky="ns")
        hs = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        hs.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)

        # Bold style for bucket tags
        try:
            self.tree.tag_configure('bucket_bold', font=self.font_bold)
        except Exception:
            pass

        # Root: Project
        self.project_root = self.tree.insert(
            "", "end", text="Project", open=True,
            tags=('project_node',), image=self.icons['project']
        )
        # Under Project: Plants bucket (always visible)
        self._insert_bucket(self.project_root, "Plants", "bucket_plants")

        # Context menu (right-click) and '+' key to add
        def _context_menu(event):
            iid = self.tree.identify_row(event.y)
            if iid:
                self.tree.selection_set(iid)
            parent_iid = self._get_insert_parent()
            parent_type = self._get_node_type(parent_iid)

            m = tk.Menu(self.master, tearoff=0)
            add_menu = tk.Menu(m, tearoff=0)
            for child_tag in self.allowed_children.get(parent_type, []):
                add_menu.add_command(
                    label=f"➕ {self.pretty_type.get(child_tag, child_tag)}",
                    command=lambda ct=child_tag: self._create_child(parent_iid, ct)
                )
            if add_menu.index("end") is not None:
                m.add_cascade(label="Add", menu=add_menu)
            m.add_command(label="Rename", command=self._rename_item)
            m.add_command(label="Delete", command=self._delete_item)
            m.post(event.x_root, event.y_root)

        self.tree.bind("<Button-3>", _context_menu)
        self.tree.bind("+", lambda e: self._add_child_clicked())
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # ---- RIGHT: reserved frame for Visio (still blank) ----
        self.visio_frame = ttk.Frame(parent, relief=tk.SUNKEN, padding=6)
        self.visio_frame.grid(row=0, column=1, sticky="nsew")

        # Allowed children map (true hierarchy)
        self.allowed_children = {
            'project_node': ['bucket_plants'],
            'bucket_plants': ['plant_node'],

            'plant_node': ['bucket_units'],
            'bucket_units': ['unit_node'],

            'unit_node': ['bucket_areas', 'bucket_equipment', 'bucket_assets'],
            'bucket_areas': ['area_node'],
            'bucket_equipment': ['equipment_node'],
            'bucket_assets': ['asset_node'],

            'area_node': ['bucket_equipment', 'bucket_assets'],

            'equipment_node': ['bucket_sub_equipment', 'bucket_assets'],
            'bucket_sub_equipment': ['sub_equipment_node'],

            'sub_equipment_node': ['bucket_assets'],

            'asset_node': [],
        }

        self.pretty_type = {
            'project_node': 'Project',
            'plant_node': 'Plant',
            'unit_node': 'Unit',
            'area_node': 'Area',
            'equipment_node': 'Equipment',
            'sub_equipment_node': 'Sub-Equipment',
            'asset_node': 'Asset',
            'bucket_plants': 'Plants',
            'bucket_units': 'Units',
            'bucket_areas': 'Areas',
            'bucket_equipment': 'Equipment',
            'bucket_sub_equipment': 'Sub-Equipment',
            'bucket_assets': 'Assets',
        }

    def _build_repository_tools_tab(self, parent):
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(4, weight=1)

        # Use the existing StringVars; DO NOT re-create them here
        ttk.Label(parent, text="Repository URL:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(parent, textvariable=self.repo_var).grid(row=0, column=1, sticky="ew", padx=6, pady=6)

        ttk.Label(parent, text="SPARQL Prefix:").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(parent, textvariable=self.prefix_var).grid(row=1, column=1, sticky="ew", padx=6, pady=6)

        # Test connection button (reuse the existing _test_connection method)
        ttk.Button(parent, text="Test Connection", command=self._test_connection) \
            .grid(row=2, column=0, columnspan=2, sticky="w", padx=6, pady=6)

        # Simple inline query tester (kept per your original tab)
        ttk.Label(parent, text="SPARQL Query Tester:").grid(row=3, column=0, sticky="nw", padx=6, pady=(6, 0))
        self.query_text = tk.Text(parent, height=10)
        self.query_text.grid(row=3, column=1, sticky="nsew", padx=6, pady=(6, 0))

        ttk.Button(parent, text="Run Query", command=self._run_sparql_query) \
            .grid(row=4, column=1, sticky="e", padx=6, pady=(6, 6))

        ttk.Label(parent, text="Results:").grid(row=5, column=0, sticky="nw", padx=6, pady=(6, 0))
        self.query_results = tk.Text(parent, height=10, state="disabled")
        self.query_results.grid(row=5, column=1, sticky="nsew", padx=6, pady=(6, 6))

    def _test_repository_connection(self):
        repo = self.repo_var.get().strip()
        prefix = self.prefix_var.get().strip()
        if not repo:
            messagebox.showerror("Error", "Please enter a repository URL.")
            return
        # Simulate test (replace with your real connection logic)
        try:
            # TODO: actually try to connect
            # e.g., requests.get(repo) or SPARQLWrapper test
            self._set_status(f"Successfully connected to repository: {repo}", 3000)
            messagebox.showinfo("Success", f"Connected to {repo} with prefix '{prefix}'")
        except Exception as e:
            messagebox.showerror("Connection Failed", str(e))

    def _run_sparql_query(self):
        repo = self.repo_var.get().strip()
        query = self.query_text.get("1.0", "end").strip()
        if not repo or not query:
            messagebox.showerror("Error", "Please enter both repository and query.")
            return
        try:
            # TODO: replace with actual SPARQL query execution
            fake_result = f"Running query against {repo}:\n\n{query}\n\n(Result rows here)"
            self.query_results.config(state="normal")
            self.query_results.delete("1.0", "end")
            self.query_results.insert("1.0", fake_result)
            self.query_results.config(state="disabled")
        except Exception as e:
            messagebox.showerror("Query Error", str(e))

    def _build_functionalities_tab(self, parent):
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        sub = ttk.Notebook(parent)
        sub.grid(row=0, column=0, sticky="nsew")
        frame_tags = ttk.Frame(sub, padding=10)
        frame_files = ttk.Frame(sub, padding=10)
        sub.add(frame_tags, text="Tag Management")
        sub.add(frame_files, text="File Management")

        # Tag subtab
        frame_tags.columnconfigure(1, weight=1)
        frame_tags.rowconfigure(1, weight=1)
        create = ttk.LabelFrame(frame_tags, text="Create or Update Tag", padding=10)
        create.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        create.columnconfigure(1, weight=1)
        ttk.Label(create, text="Tag Name:").grid(row=0, column=0, sticky="w")
        self.entry_tag = ttk.Entry(create)
        self.entry_tag.grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(create, text="Associate Node(s):").grid(row=1, column=0, sticky="w", pady=(6,0))
        self.node_combobox = ttk.Combobox(create, state="readonly")
        self.node_combobox.grid(row=1, column=1, sticky="ew", padx=6, pady=(6,0))

        ttk.Label(create, text="Associate Datasheet(s) [Creates a Copy]:").grid(row=2, column=0, sticky="w", pady=(6,0))
        self.datasheet_combobox = ttk.Combobox(create, state="readonly")
        self.datasheet_combobox.grid(row=2, column=1, sticky="ew", padx=6, pady=(6,0))

        ttk.Button(create, text="Create/Update Tag", command=self._add_tag).grid(row=3, column=1, sticky="e", pady=(10,0))

        view = ttk.LabelFrame(frame_tags, text="View Tag Associations", padding=10)
        view.grid(row=1, column=0, columnspan=2, sticky="nsew")
        view.columnconfigure(1, weight=1)
        view.rowconfigure(3, weight=1)

        ttk.Label(view, text="Existing Tags (Double-click to view)").grid(row=0, column=0, columnspan=2, sticky="w")
        self.tag_listbox = tk.Listbox(view, height=6, exportselection=False)
        self.tag_listbox.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4,6))
        self.tag_listbox.bind("<Double-1>", lambda e: self._show_tag_connections())

        ttk.Label(view, text="Associated Nodes").grid(row=2, column=0, sticky="w")
        self.nodes_display = tk.Listbox(view, height=7, exportselection=False)
        self.nodes_display.grid(row=3, column=0, sticky="nsew", padx=(0,6))

        ttk.Label(view, text="Associated Datasheets").grid(row=2, column=1, sticky="w")
        self.datasheets_display = tk.Listbox(view, height=7, exportselection=False)
        self.datasheets_display.grid(row=3, column=1, sticky="nsew")
        self.datasheets_display.bind("<Double-1>", self._load_datasheet_from_functionalities_tab)

        # File subtab
        frame_files.columnconfigure((0,1), weight=1)
        frame_files.rowconfigure(2, weight=1)
        ttk.Label(frame_files, text="Manage all imported Excel files in the application library.", style='Info.TLabel').grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(frame_files, text="Assigned Datasheets", style='Header.TLabel').grid(row=1, column=0, sticky="w")
        self.assigned_files_listbox = tk.Listbox(frame_files, height=12)
        self.assigned_files_listbox.grid(row=2, column=0, sticky="nsew", padx=(0,6), pady=6)

        ttk.Label(frame_files, text="Unassigned Datasheets (Templates)", style='Header.TLabel').grid(row=1, column=1, sticky="w")
        self.unassigned_files_listbox = tk.Listbox(frame_files, height=12)
        self.unassigned_files_listbox.grid(row=2, column=1, sticky="nsew", padx=(6,0), pady=6)

        bar = ttk.Frame(frame_files)
        bar.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6,0))
        ttk.Button(bar, text="Import Template…", command=self._import_datasheet_template).pack(side="left")
        ttk.Button(bar, text="Remove Selected", command=self._remove_file).pack(side="left", padx=6)

    # ---------- helpers ----------
    def _filter_listbox(self, listbox, term):
        term = term.lower().strip()
        items = list(listbox.get(0, "end"))
        # store original on widget
        if not hasattr(listbox, "_all_items"):
            listbox._all_items = items
        if not term:
            filtered = listbox._all_items
        else:
            filtered = [i for i in listbox._all_items if term in i.lower()]
        listbox.delete(0, tk.END)
        for i in filtered:
            listbox.insert(tk.END, i)

    def _filter_tree(self):
        term = self.tree_filter_var.get().lower().strip()
        # naive approach: collapse/expand by text match
        for iid in self.tree.get_children(""):
            self._filter_tree_recursive(iid, term)

    def _filter_tree_recursive(self, iid, term):
        text = self.tree.item(iid, "text").lower()
        match = (term in text) if term else True
        children = self.tree.get_children(iid)
        child_match = False
        for c in children:
            child_match |= self._filter_tree_recursive(c, term)
        show = match or child_match
        self.tree.detach(iid) if not show else self.tree.reattach(iid, self.tree.parent(iid), "end")
        return show

    # ---------- connection & SPARQL ----------
    def _test_connection(self):
        url = self.repo_var.get().strip()
        if not url:
            messagebox.showwarning("Missing URL", "Enter a repository URL.")
            return
        self._with_progress(True)
        self._set_status("Testing connection…")
        def job():
            try:
                s = SPARQLWrapper(url)
                s.setQuery("ASK { ?s ?p ?o }")
                s.setReturnFormat(JSON)
                r = s.query().convert()
                return bool(r.get("boolean", False))
            except Exception as e:
                return e
        fut = self.executor.submit(job)
        self._track_future(fut, lambda res: self._after_test_connection(res))

    def _after_test_connection(self, res):
        self._with_progress(False)
        if isinstance(res, Exception):
            messagebox.showerror("Connection Failed", f"{res}")
            self._set_status("Connection failed.", 4000)
        else:
            show_toast(self.master, "Connection OK")
            self._set_status("Connection OK", 3000)

    def _run_sparql_query_bg(self, query, timeout=20):
        url = self.repo_var.get().strip()
        if not url:
            raise RuntimeError("Repository URL is not set.")
        def run():
            s = SPARQLWrapper(url)
            s.setQuery(query)
            s.setReturnFormat(JSON)
            return s.query().convert()["results"]["bindings"]
        fut = self.executor.submit(run)
        return fut, timeout

    def _track_future(self, future, callback):
        self.future_tasks.add(future)
        def done(_):
            self.future_tasks.discard(future)
            try:
                res = future.result()
            except TimeoutError as e:
                res = e
            except Exception as e:
                res = e
            self.master.after(0, lambda: callback(res))
        future.add_done_callback(done)

    # ---------- node fetching & model ----------
    def _fetch_nodes(self):
        prefix_str = self.prefix_var.get().strip()
        if not prefix_str or '<' not in prefix_str or '>' not in prefix_str:
            messagebox.showerror("Invalid Prefix", "Provide a valid SPARQL Prefix (e.g. PREFIX ex: <http://example.org#>)")
            return
        self._with_progress(True)
        self._set_status("Fetching all nodes from repository…")
        try:
            uri_base = prefix_str.split('<')[1].split('>')[0]
        except Exception:
            messagebox.showerror("Invalid Prefix", "Could not parse base URI from prefix.")
            self._with_progress(False)
            return
        query = f'''
            SELECT DISTINCT ?resource WHERE {{
                {{ ?resource ?p ?o . }} UNION {{ ?s ?p ?resource . }}
                FILTER(ISIRI(?resource) && STRSTARTS(STR(?resource), "{uri_base}"))
            }} ORDER BY ?resource
        '''
        fut, to = self._run_sparql_query_bg(query, timeout=45)
        self._track_future(fut, lambda res: self._after_fetch_nodes(res))

    def _after_fetch_nodes(self, res):
        self._with_progress(False)
        if isinstance(res, Exception):
            messagebox.showerror("Fetch Failed", f"{res}")
            self._set_status("Failed to fetch nodes.", 4000)
            return
        nodes = sorted(list(set(r["resource"]["value"].split('#')[-1].split('/')[-1] for r in res)))
        self.fetched_nodes['all'] = nodes
        self._update_node_lists(nodes)
        show_toast(self.master, f"Fetched {len(nodes)} nodes")
        self._set_status("Nodes fetched.", 3000)

    def _update_node_lists(self, nodes):
        self.node_listbox.delete(0, tk.END)
        self.node_listbox_display.delete(0, tk.END)
        for n in nodes:
            self.node_listbox.insert(tk.END, n)
            self.node_listbox_display.insert(tk.END, n)
        self.node_combobox['values'] = nodes
        self._update_badges()

    def _categorize_nodes(self):
        if not self.fetched_nodes['all']:
            messagebox.showwarning("No Nodes", "Fetch nodes first.")
            return
        self._with_progress(True)
        self._set_status("Categorizing nodes by type…")

        prefix = self.prefix_var.get().strip()
        queries = {
            'equipment': f"""{prefix} SELECT DISTINCT ?x WHERE {{ ?x a ex:Equipment . }} ORDER BY ?x""",
            'sub_equipment': f"""{prefix} SELECT DISTINCT ?x WHERE {{ ?x a ex:SubEquipment . }} ORDER BY ?x""",
            'asset': f"""{prefix} SELECT DISTINCT ?x WHERE {{ ?x a ex:Asset . }} ORDER BY ?x"""
        }

        results = {}

        def make_cb(kind):
            def cb(res):
                results[kind] = [] if isinstance(res, Exception) else [
                    r['x']['value'].split('#')[-1].split('/')[-1] for r in res
                ]
                if len(results) == 3:
                    self._with_progress(False)
                    self.fetched_nodes.update(results)
                    self._update_badges()
                    show_toast(self.master, f"Categorized: Eq {len(results['equipment'])} | Sub {len(results['sub_equipment'])} | Asset {len(results['asset'])}")
                    self._set_status("Categorization complete.", 4000)
            return cb

        for k, q in queries.items():
            fut, to = self._run_sparql_query_bg(q, timeout=30)
            self._track_future(fut, make_cb(k))

    def _update_badges(self):
        a = len(self.fetched_nodes.get('all', []))
        e = len(self.fetched_nodes.get('equipment', []))
        s = len(self.fetched_nodes.get('sub_equipment', []))
        t = len(self.fetched_nodes.get('asset', []))
        self.badge_nodes.config(text=f"Nodes: {a} | Eq:{e} Sub:{s} Asset:{t}")

    def _update_data_model(self, event=None):
        sel = self.node_listbox_display.curselection()
        if not sel:
            return
        nodes = [self.node_listbox_display.get(i) for i in sel]
        repo_url, prefix = self.repo_var.get().strip(), self.prefix_var.get().strip()
        if not all([repo_url, prefix, nodes]):
            messagebox.showwarning("Missing Data", "Repository URL, Prefix, and selected nodes are required.")
            return
        self._clear_graph()
        self._with_progress(True)
        self._set_status("Fetching data model…")

        node_conditions = " || ".join([f"sameTerm(?subject, ex:{n}) || sameTerm(?object, ex:{n})" for n in nodes])
        query = f"""{prefix}
            SELECT ?subject ?predicate ?object WHERE {{
                ?subject ?predicate ?object .
                FILTER({node_conditions})
            }}
        """
        fut, to = self._run_sparql_query_bg(query, timeout=45)
        def after(res):
            self._with_progress(False)
            if isinstance(res, Exception):
                messagebox.showerror("SPARQL Error", f"{res}")
                self._set_status("Error loading model.", 4000)
                return
            self.graph.clear()
            def local(x):
                if not isinstance(x, str): return x
                return x.split('#')[-1].split('/')[-1]
            for r in res:
                s = local(r.get("subject", {}).get("value", ""))
                p = local(r.get("predicate", {}).get("value", ""))
                o = local(r.get("object", {}).get("value", ""))
                if s and p and o:
                    self.graph.add_edge(s, o, label=p)
            self._draw_graph()
            self._set_status("Data model loaded.", 4000)
        self._track_future(fut, after)

    def _draw_graph(self):
        for w in self.graph_frame.winfo_children():
            w.destroy()
        if not self.graph.nodes():
            self._set_status("No data for selection.", 4000)
            return
        try:
            fig, ax = plt.subplots(figsize=(10, 8))
            pos = nx.spring_layout(self.graph, k=0.7, iterations=50)
            nx.draw(self.graph, pos, ax=ax, with_labels=True, node_color='#a0cbe2',
                    node_size=2500, font_size=10, font_weight='bold', width=1.5,
                    edge_color='gray', arrows=True)
            edge_labels = nx.get_edge_attributes(self.graph, 'label')
            nx.draw_networkx_edge_labels(self.graph, pos, edge_labels=edge_labels, font_color='firebrick', font_size=9)
            fig.tight_layout()
            self.canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
            self.canvas.draw()
            self.canvas.get_tk_widget().pack(expand=True, fill="both")
            plt.close(fig)
        except Exception as e:
            logger.exception("Graph draw error")
            messagebox.showerror("Graphing Error", f"{e}")

    def _clear_graph(self):
        for w in self.graph_frame.winfo_children():
            w.destroy()
        self.graph.clear()
        if self.canvas:
            self.canvas = None
        self._set_status("Graph cleared.", 2000)

    # ---------- Excel embedding ----------
    def _resize_excel_window(self, _=None):
        if self.xl_app and self.xl_app.hwnd:
            try:
                win32gui.MoveWindow(self.xl_app.hwnd, 0, 0, self.excel_frame.winfo_width(),
                                    self.excel_frame.winfo_height(), True)
            except win32gui.error:
                self.xl_app = None

    def _embed_excel_window(self):
        if not self.xl_app or not self.current_xl_db: return
        try:
            frame_hwnd = self.excel_frame.winfo_id()
            excel_hwnd = self.xl_app.hwnd
            win32gui.SetParent(excel_hwnd, frame_hwnd)
            style = win32gui.GetWindowLong(excel_hwnd, win32con.GWL_STYLE)
            style &= ~(win32con.WS_CAPTION | win32con.WS_THICKFRAME)
            win32gui.SetWindowLong(excel_hwnd, win32con.GWL_STYLE, style)
            self._resize_excel_window()
        except Exception as e:
            logger.exception("Embedding error")
            messagebox.showerror("Embedding Error", f"Could not embed Excel.\n\n{e}")
            self._cleanup_xlwings()

    def _popout_excel(self):
        try:
            if self.xl_app and self.xl_app.hwnd:
                win32gui.SetParent(self.xl_app.hwnd, 0)
                self._set_status("Excel popped out.", 3000)
        except Exception:
            pass

    def _load_excel_file(self, file_path):
        try:
            abs_path = os.path.abspath(file_path)
            if not os.path.exists(abs_path):
                messagebox.showerror("File Not Found", f"The file could not be found:\n{abs_path}")
                return
            self._cancel_drag_mode()
            if self.xl_app is None:
                self.xl_app = xw.App(visible=True, add_book=False)
            self.current_xl_db = self.xl_app.books.open(abs_path)
            self.current_xl_db.activate()
            self.master.after(200, self._embed_excel_window)
            sheet_names = [s.name for s in self.current_xl_db.sheets]
            self.sheet_selector_combobox['values'] = sheet_names
            if sheet_names:
                self.sheet_selector_combobox.set(sheet_names[0])
            self.current_excel_path = abs_path
            self._set_status(f"Embedding '{os.path.basename(file_path)}'…", 4000)
        except Exception as e:
            logger.exception("xlwings load error")
            messagebox.showerror("Excel Load Error", f"{e}")
            self._cleanup_xlwings()
            self.current_xl_db = None
            self.sheet_selector_combobox['values'] = []
            self.sheet_selector_combobox.set('')

    def _on_sheet_selected(self):
        name = self.sheet_selector_combobox.get()
        if not name or not self.current_xl_db:
            return
        try:
            self.current_xl_db.sheets[name].activate()
            self._set_status(f"Activated sheet '{name}'.", 2500)
        except Exception as e:
            messagebox.showerror("Sheet Activation Error", f"{e}")

    def _cleanup_xlwings(self):
        if self.xl_app:
            try:
                if self.xl_app.hwnd:
                    win32gui.SetParent(self.xl_app.hwnd, 0)
                self.xl_app.quit()
            except Exception:
                pass
            self.xl_app = None
            self._set_status("Excel instance closed.", 2000)

    # ---------- Drag copy ----------
    def _initiate_drag_copy(self, value_to_drag, source_type):
        if value_to_drag in (None, ""):
            messagebox.showinfo("No Value", f"There is no {source_type} to drag.")
            return
        if self.is_dragging and self.drag_source_type == source_type:
            self._cancel_drag_mode()
        else:
            self.is_dragging = True
            self.drag_data = value_to_drag
            self.drag_source_type = source_type
            self.master.config(cursor="hand2")
            self._set_status(f"DRAGGING {source_type.upper()}: click in Excel to drop.", 0)
            self._poll_for_drop()

    def _cancel_drag_mode(self):
        self.is_dragging = False
        self.drag_data = None
        self.drag_source_type = None
        self.master.config(cursor="")
        self.live_value_button.config(text=f"Copy Value: {self.live_value_button.value or '(none)'}")
        self.live_unit_button.config(text=f"Copy Unit: {self.live_unit_button.value or '(none)'}")
        self._set_status("Ready", 0)

    def _poll_for_drop(self):
        if not self.is_dragging: return
        if win32api.GetKeyState(0x01) < 0:
            sx, sy = win32gui.GetCursorPos()
            fx = self.excel_frame.winfo_rootx()
            fy = self.excel_frame.winfo_rooty()
            fw = self.excel_frame.winfo_width()
            fh = self.excel_frame.winfo_height()
            if fx <= sx <= fx + fw and fy <= sy <= fy + fh:
                try:
                    xl_range = self.xl_app.api.ActiveWindow.RangeFromPoint(sx, sy)
                    if xl_range:
                        sheet = self.current_xl_db.sheets.active
                        com = xl_range.MergeArea.Cells(1, 1)
                        write_cell = sheet.range(com.Address)
                        write_cell.value = self.drag_data
                        self.current_xl_db.save()
                        self._set_status(f"Pasted '{self.drag_data}' to {write_cell.address.replace('$','')}.", 4000)
                        self._cancel_drag_mode()
                        return
                except Exception as e:
                    self._set_status(f"Drop Error: {e}", 3000)
                    self._cancel_drag_mode()
                    return
        self.master.after(100, self._poll_for_drop)

    # ---------- Files / templates ----------
    def _import_datasheet_template(self):
        path = filedialog.askopenfilename(
            title="Select a Datasheet Template",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
        )
        if not path: return
        try:
            fn = os.path.basename(path)
            dest = os.path.join(self.data_dir, fn)
            if os.path.exists(dest):
                if not messagebox.askyesno("Overwrite?", f"'{fn}' exists. Overwrite?"):
                    return
            shutil.copy2(path, dest)
            self._refresh_file_and_tag_lists()
            show_toast(self.master, f"Imported {fn}")
            self._set_status("Library updated.", 3000)
        except Exception as e:
            messagebox.showerror("Import Error", f"{e}")

    def _save_excel_as(self):
        if not self.current_excel_path or not self.current_xl_db:
            messagebox.showwarning("No File", "Open a datasheet first.")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx *.xlsm")],
                                                 title="Save Datasheet As")
        if not save_path: return
        try:
            self.current_xl_db.save(os.path.abspath(save_path))
            fn = os.path.basename(save_path)
            lib_path = os.path.join(self.data_dir, fn)
            if os.path.normpath(save_path).lower() != os.path.normpath(lib_path).lower():
                shutil.copyfile(save_path, lib_path)
            self._refresh_file_and_tag_lists()
            messagebox.showinfo("Success", f"Exported as '{fn}' and added to library.")
        except Exception as e:
            messagebox.showerror("Save As Error", f"{e}")

    def _remove_file(self):
        file_name = None
        if self.assigned_files_listbox.curselection():
            file_name = self.assigned_files_listbox.get(self.assigned_files_listbox.curselection())
        elif self.unassigned_files_listbox.curselection():
            file_name = self.unassigned_files_listbox.get(self.unassigned_files_listbox.curselection())
        if not file_name:
            messagebox.showwarning("No Selection", "Select a file to remove.")
            return
        if self.current_xl_db and self.current_xl_db.name == file_name:
            self.current_xl_db.close()
            self.current_xl_db = None
            self.sheet_selector_combobox['values'] = []
            self.sheet_selector_combobox.set('')
        if messagebox.askyesno("Confirm Removal", f"Delete '{file_name}' from library and untag it?"):
            try:
                os.remove(os.path.join(self.data_dir, file_name))
                for tag in self.tag_associations:
                    if file_name in self.tag_associations[tag].get('datasheets', []):
                        self.tag_associations[tag]['datasheets'].remove(file_name)
                self._refresh_file_and_tag_lists()
                messagebox.showinfo("Removed", f"'{file_name}' deleted.")
            except Exception as e:
                messagebox.showerror("Error", f"{e}")

    def _load_file_from_list(self, event=None):
        lb = event.widget if event else self.datasheet_listbox_for_tag
        if not lb.curselection(): return
        file_name = lb.get(lb.curselection())
        self.nb.select(self.tab_excel)
        self._load_excel_file(os.path.join(self.data_dir, file_name))

    def _load_datasheet_from_functionalities_tab(self, event=None):
        """Open the datasheet double-clicked in the 'Functionalities > Tag Management' panel."""
        if not self.datasheets_display.curselection():
            return
        file_name = self.datasheets_display.get(self.datasheets_display.curselection())
        self.nb.select(self.tab_excel)
        self._load_excel_file(os.path.join(self.data_dir, file_name))

    # ---------- Tags ----------
    def _refresh_file_and_tag_lists(self):
        try:
            os.makedirs(self.data_dir, exist_ok=True)
            all_files = sorted([f for f in os.listdir(self.data_dir) if f.lower().endswith(('.xlsx', '.xls', '.xlsm'))])
            assigned = set()
            for tdata in self.tag_associations.values():
                assigned.update(tdata.get('datasheets', []))
            self.assigned_files_listbox.delete(0, tk.END)
            self.unassigned_files_listbox.delete(0, tk.END)
            for f in all_files:
                (self.assigned_files_listbox if f in assigned else self.unassigned_files_listbox).insert(tk.END, f)
            self.datasheet_combobox['values'] = all_files

            tags = sorted(self.tag_associations.keys())
            self.tag_listbox.delete(0, tk.END)
            for t in tags:
                self.tag_listbox.insert(tk.END, t)
            self.tag_selector_combobox['values'] = tags
        except Exception as e:
            logger.exception("refresh lists")

    def _add_tag(self):
        tag = self.entry_tag.get().strip()
        node = self.node_combobox.get().strip()
        template = self.datasheet_combobox.get().strip()

        if not tag:
            messagebox.showwarning("Input Required", "Tag name cannot be empty.")
            return

        datasheet_to_assign = None
        if template:
            if not node:
                messagebox.showwarning("Node Required", "Select a node to associate with the new datasheet copy.")
                return
            name, ext = os.path.splitext(template)
            default_new_name = f"{node}_datasheet{ext}"
            new_filename = simpledialog.askstring("Name New Datasheet", "Enter filename for the copy:", initialvalue=default_new_name)
            if not new_filename:
                return
            src = os.path.join(self.data_dir, template)
            dst = os.path.join(self.data_dir, new_filename)
            if os.path.exists(dst):
                if not messagebox.askyesno("Overwrite?", f"'{new_filename}' exists. Overwrite?"):
                    return
            try:
                shutil.copy2(src, dst)
                datasheet_to_assign = new_filename
                show_toast(self.master, f"Created '{new_filename}'")
            except Exception as e:
                messagebox.showerror("File Copy Error", f"{e}")
                return

        if tag not in self.tag_associations:
            self.tag_associations[tag] = {'nodes': [], 'datasheets': []}

        if node and node not in self.tag_associations[tag]['nodes']:
            self.tag_associations[tag]['nodes'].append(node)

        if datasheet_to_assign and datasheet_to_assign not in self.tag_associations[tag]['datasheets']:
            self.tag_associations[tag]['datasheets'].append(datasheet_to_assign)

        self._refresh_file_and_tag_lists()
        self._show_tag_connections(tag_name=tag)
        self._persist_state()

    def _show_tag_connections(self, event=None, tag_name=None):
        tag = tag_name or (self.tag_listbox.get(self.tag_listbox.curselection()) if self.tag_listbox.curselection() else None)
        if not tag: return
        self.nodes_display.delete(0, tk.END)
        self.datasheets_display.delete(0, tk.END)
        if tag in self.tag_associations:
            for node in self.tag_associations[tag].get('nodes', []):
                self.nodes_display.insert(tk.END, node)
            for ds in self.tag_associations[tag].get('datasheets', []):
                self.datasheets_display.insert(tk.END, ds)

    def _on_tag_selected_in_editor_tab(self):
        tag_name = self.tag_selector_combobox.get()
        if not tag_name: return
        self._cancel_drag_mode()
        self.selected_property_display.config(text="(None)")

        nodes = self.tag_associations.get(tag_name, {}).get('nodes', [])
        if len(nodes) == 1:
            self._set_active_node(nodes[0])
        else:
            self._set_active_node(None)

        self.datasheet_listbox_for_tag.delete(0, tk.END)
        if tag_name in self.tag_associations:
            for ds in self.tag_associations[tag_name].get('datasheets', []):
                self.datasheet_listbox_for_tag.insert(tk.END, ds)
        self._display_tag_info_in_editor_view(tag_name)

    def _display_tag_info_in_editor_view(self, tag_name):
        self.tag_text_embed.config(state='normal'); self.properties_text_embed.config(state='normal')
        self.tag_text_embed.delete(1.0, tk.END); self.properties_text_embed.delete(1.0, tk.END)

        if tag_name not in self.tag_associations:
            self.tag_text_embed.insert(tk.END, "Tag not found.")
            self.tag_text_embed.config(state='disabled'); self.properties_text_embed.config(state='disabled')
            return

        self.tag_text_embed.insert(tk.END, f"Tag: {tag_name}")
        nodes = self.tag_associations[tag_name].get('nodes', [])
        if not nodes:
            self.properties_text_embed.insert(tk.END, "Tag has no associated nodes.")
        else:
            node_list = '\n- '.join(nodes)
            self.properties_text_embed.insert(tk.END, f"Associated Nodes:\n- {node_list}\n\nProperties (from all nodes):\n")

        self.tag_text_embed.config(state='disabled'); self.properties_text_embed.config(state='disabled')
        if not nodes: return

        self._set_status(f"Fetching properties for '{tag_name}'…")
        def job():
            all_props = set()
            prefix = self.prefix_var.get().strip()
            for n in nodes:
                q = f"""{prefix}
                    SELECT DISTINCT ?p WHERE {{
                        ex:{n} ?p ?o .
                        FILTER (isLiteral(?o) || isBlank(?o))
                    }}
                """
                try:
                    bindings = self._run_sparql_query_bg(q)[0].result(timeout=25)
                except Exception:
                    continue
                for r in bindings or []:
                    p = r['p']['value']
                    name = p.split('#')[-1]
                    if name not in ["hasValue", "hasUnit", "a", "type"]:
                        all_props.add(name)
            return sorted(all_props)
        fut = self.executor.submit(job)
        def after(res):
            self.properties_text_embed.config(state='normal')
            if isinstance(res, Exception) or not res:
                self.properties_text_embed.insert(tk.END, "(No direct properties found)")
            else:
                self.properties_text_embed.insert(tk.END, "\n".join(f"- {p}" for p in res))
            self.properties_text_embed.config(state='disabled')
            self._set_status("Info loaded.", 3000)
        self._track_future(fut, after)

    def _on_property_double_click(self, event=None):
        try:
            index = self.properties_text_embed.index(f"@{event.x},{event.y} linestart")
            line_end = self.properties_text_embed.index(f"{index} lineend")
            line_text = self.properties_text_embed.get(index, line_end)
            import re
            m = re.search(r"^\s*[-*]?\s*([\w_]+)", line_text)
            if not m: return
            prop = m.group(1)
            self.selected_property_display.config(text=prop)
            if not self.active_node:
                self.live_value_button.value = None; self.live_unit_button.value = None
                self.live_value_button.config(text="Copy Value: (No Active Node)")
                self.live_unit_button.config(text="Copy Unit: (No Active Node)")
                messagebox.showinfo("Info", "Set an active node to preview.")
                return
            node = self.active_node
            self._set_status(f"Fetching preview for {prop}…")
            prefix = self.prefix_var.get().strip()
            q = f"""{prefix} SELECT ?value ?unit WHERE {{
                ex:{node} ex:{prop} ?b .
                ?b ex:hasValue ?value .
                OPTIONAL {{ ?b ex:hasUnit ?unit . }}
            }} LIMIT 1"""
            fut, to = self._run_sparql_query_bg(q, timeout=20)
            def after(res):
                if isinstance(res, Exception) or not res:
                    val, uni = (None, None)
                else:
                    val = res[0]["value"]["value"]
                    uni = res[0].get("unit", {}).get("value")
                self.live_value_button.value = val
                self.live_unit_button.value = uni
                self.live_value_button.config(text=f"Copy Value: {val or '(none)'}")
                self.live_unit_button.config(text=f"Copy Unit: {uni or '(none)'}")
                self.last_preview_ts = now_str()
                self.preview_ts_lbl.config(text=f"Last preview: {self.last_preview_ts}")
                self._set_status("Preview loaded.", 3000)
            self._track_future(fut, after)
        except (tk.TclError, IndexError):
            pass

    # ---------- Active node ----------
    def _on_node_manual_select(self):
        if not self.node_listbox.curselection(): return
        node_name = self.node_listbox.get(self.node_listbox.curselection())
        self._set_active_node(node_name)

    def _set_active_node(self, node_name):
        self.active_node = node_name
        if node_name:
            self.active_node_display_label.config(text=node_name)
            try:
                idx = self.node_listbox.get(0, "end").index(node_name)
                self.node_listbox.selection_clear(0, tk.END)
                self.node_listbox.selection_set(idx)
                self.node_listbox.see(idx)
            except ValueError:
                pass
        else:
            self.active_node_display_label.config(text="None (Select a tag with one node)")
            self.node_listbox.selection_clear(0, tk.END)

    # ---------- Tree context menu ----------
    def _show_context_menu(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        self.tree.selection_set(item_id)
        self.current_item_id = item_id
        tags = self.tree.item(item_id, "tags")
        item_type = tags[0] if tags else None
        m = tk.Menu(self.master, tearoff=0)
        if item_type == 'root':
            m.add_command(label="Create Plant", command=self._create_plant)
        elif item_type == 'plant_node':
            m.add_command(label="Create Unit", command=self._create_unit)
            m.add_command(label="Create Area", command=self._create_area)
        elif item_type == 'unit_node':
            m.add_command(label="Create Area", command=self._create_area)
            m.add_command(label="Create Equipment", command=lambda: self._create_from_repository('equipment_node'))
            m.add_command(label="Create Assets", command=lambda: self._create_from_repository('asset_node'))
        elif item_type == 'area_node':
            m.add_command(label="Create Equipment", command=lambda: self._create_from_repository('equipment_node'))
            m.add_command(label="Create Assets", command=lambda: self._create_from_repository('asset_node'))
        elif item_type == 'equipment_node':
            m.add_command(label="Create Sub-Equipment", command=lambda: self._create_from_repository('sub_equipment_node'))
            m.add_command(label="Create Assets", command=lambda: self._create_from_repository('asset_node'))
        m.add_separator()
        m.add_command(label="Rename", command=self._rename_item)
        m.add_command(label="Delete", command=self._delete_item)
        if m.index("end") is not None:
            m.post(event.x_root, event.y_root)

    def _rename_item(self):
        iid = self.current_item_id or (self.tree.selection()[0] if self.tree.selection() else None)
        if not iid: return
        txt = self.tree.item(iid, "text")
        new = simpledialog.askstring("Rename", "New name:", initialvalue=txt)
        if new:
            self.tree.item(iid, text=new)

    def _delete_item(self):
        iid = self.current_item_id or (self.tree.selection()[0] if self.tree.selection() else None)
        if not iid: return
        if messagebox.askyesno("Delete", f"Delete '{self.tree.item(iid, 'text')}'?"):
            self.tree.delete(iid)

    def _create_plant(self):
        name = simpledialog.askstring("Create Plant", "Enter Plant Name:")
        if name:
            self.tree.insert(self.current_item_id, "end", text=name, open=True, tags=('plant_node',))

    def _create_unit(self):
        name = simpledialog.askstring("Create Unit", "Enter Unit Name:")
        if name:
            self.tree.insert(self.current_item_id, "end", text=name, open=True, tags=('unit_node',))

    def _create_area(self):
        name = simpledialog.askstring("Create Area", "Enter Area Name:")
        if name:
            self.tree.insert(self.current_item_id, "end", text=name, open=True, tags=('area_node',))

    def _create_from_repository(self, node_tag):
        if not self.fetched_nodes['all']:
            messagebox.showwarning("No Nodes Available", "Fetch repository nodes first (GraphDB → Excel tab).")
            return
        node_type_map = {'equipment_node': 'equipment', 'sub_equipment_node': 'sub_equipment', 'asset_node': 'asset'}
        node_type = node_type_map.get(node_tag, 'all')
        repo_nodes = self.fetched_nodes.get(node_type) or self.fetched_nodes['all']
        if not repo_nodes:
            messagebox.showwarning("No Nodes Available", f"No {node_type.replace('_',' ')} nodes found.")
            return
        titles = {
            'equipment_node': "Select Equipment from Repository",
            'sub_equipment_node': "Select Sub-Equipment from Repository",
            'asset_node': "Select Assets from Repository"
        }
        dlg = SelectionDialog(self.master, title=titles.get(node_tag, "Select Items"), item_list=repo_nodes, node_type=node_type)
        self.master.wait_window(dlg)
        if dlg.result:
            for name in dlg.result:
                self.tree.insert(self.current_item_id, "end", text=name, open=True, tags=(node_tag,))
            self._set_status(f"Added {len(dlg.result)} {node_type.replace('_',' ')} node(s).", 3000)

    # ---------- Project serialization ----------
    def _serialize_tree(self):
        if not self.tree or not self.project_root:
            return {}

        def build(iid):
            tags = tuple(self.tree.item(iid, 'tags') or ())
            if 'inline_info' in tags:
                return None
            node = {
                'text': self.tree.item(iid, 'text'),
                'tags': list(tags),
                'open': bool(self.tree.item(iid, 'open')),
                'children': []
            }
            for c in self.tree.get_children(iid):
                child_node = build(c)
                if child_node is not None:
                    node['children'].append(child_node)
            return node

        return build(self.project_root)

    def _rebuild_tree_from_serialized(self, data):
        if not data:
            return
        # Clear existing tree
        for iid in self.tree.get_children(""):
            self.tree.delete(iid)

        def icon_for(tags_list):
            tag0 = tags_list[0] if tags_list else 'project_node'
            return self.icons.get(self.icon_for_type.get(tag0, 'folder'))

        def insert(node, parent):
            tags = tuple(node.get('tags', []))
            text = node.get('text', '')
            open_state = bool(node.get('open', True))
            img = icon_for(list(tags))
            # add bucket_bold on bucket_* nodes
            tags_tuple = tags + ( 'bucket_bold', ) if any(str(t).startswith('bucket_') for t in tags) else tags
            iid = self.tree.insert(parent, 'end', text=text, open=open_state, tags=tags_tuple, image=img)
            for child in node.get('children', []):
                insert(child, iid)
            return iid

        # Build root and keep reference
        self.project_root = insert(data, '')

    # ---------- Save/Load project ----------
    def _export_project(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".zip",
            filetypes=[("Project Zip", "*.zip")],
            title="Export Project (project.json + tags.json + library)"
        )
        if not path: return
        try:
            project_data = self._serialize_tree()
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
                # metadata
                z.writestr("settings.json", json.dumps(self.settings, indent=2))
                z.writestr("tags.json", json.dumps(self.tag_associations, indent=2))
                z.writestr("project.json", json.dumps(project_data, indent=2))
                # files
                for f in os.listdir(self.data_dir):
                    fp = os.path.join(self.data_dir, f)
                    if os.path.isfile(fp):
                        z.write(fp, arcname=os.path.join("excel_files", f))
            show_toast(self.master, "Project exported")
        except Exception as e:
            messagebox.showerror("Export Failed", f"{e}")

    def _import_project(self):
        path = filedialog.askopenfilename(
            title="Import Project", filetypes=[("Project Zip", "*.zip")]
        )
        if not path: return
        try:
            imported_settings = None
            imported_tags = None
            imported_project = None
            with zipfile.ZipFile(path, "r") as z:
                for name in z.namelist():
                    if name.endswith("settings.json"):
                        imported_settings = json.loads(z.read(name).decode("utf-8"))
                    elif name.endswith("tags.json"):
                        imported_tags = json.loads(z.read(name).decode("utf-8"))
                    elif name.endswith("project.json"):
                        imported_project = json.loads(z.read(name).decode("utf-8"))
                    elif name.startswith("excel_files/") and not name.endswith("/"):
                        os.makedirs(self.data_dir, exist_ok=True)
                        target = os.path.join(self.data_dir, os.path.basename(name))
                        with open(target, "wb") as f:
                            f.write(z.read(name))
            # Apply imported content
            if imported_settings:
                self.settings.update(imported_settings)
                self.repo_var.set(self.settings.get("repo_url", ""))
                self.prefix_var.set(self.settings.get("sparql_prefix", ""))
                self._configure_styles()
            if imported_tags is not None:
                self.tag_associations = imported_tags
                safe_json_dump(self.tags_path, self.tag_associations)
            if imported_project is not None:
                self._rebuild_tree_from_serialized(imported_project)
            self._refresh_file_and_tag_lists()
            show_toast(self.master, "Project imported")
        except Exception as e:
            messagebox.showerror("Import Failed", f"{e}")

    def _save_project_folder(self, force_ask=False):
        # Determine target directory
        if force_ask or not self.project_dir or not os.path.isdir(self.project_dir):
            chosen = filedialog.askdirectory(title="Choose Project Folder to Save")
            if not chosen:
                return
            self.project_dir = chosen
            self.data_dir = os.path.join(self.project_dir, "excel_files")
            self.tags_path = os.path.join(self.project_dir, "tags.json")
            os.makedirs(self.data_dir, exist_ok=True)
        try:
            # Save tags
            safe_json_dump(self.tags_path, self.tag_associations)
            # Save tree as project.json
            proj_path = os.path.join(self.project_dir, "project.json")
            safe_json_dump(proj_path, self._serialize_tree())
            # Copy library
            os.makedirs(self.data_dir, exist_ok=True)
            src_files = [f for f in os.listdir(self.data_dir) if os.path.isfile(os.path.join(self.data_dir, f))]
            # Already in place as we pointed data_dir to project_dir/excel_files; still ensure exists
            # Persist app settings reference
            self._persist_state()
            show_toast(self.master, f"Project saved")
            self._set_status(f"Saved to {self.project_dir}", 4000)
        except Exception as e:
            messagebox.showerror("Save Project Failed", f"{e}")

    def _new_project_folder(self):
        chosen = filedialog.askdirectory(title="Create/Select Empty Project Folder")
        if not chosen:
            return
        # Confirm emptiness or consent to use
        if os.listdir(chosen):
            if not messagebox.askyesno("Use Non-Empty Folder?", "Folder is not empty. Use it for a new project anyway?"):
                return
        # Reset current project data in UI
        self.project_dir = chosen
        self.data_dir = os.path.join(self.project_dir, "excel_files")
        self.tags_path = os.path.join(self.project_dir, "tags.json")
        os.makedirs(self.data_dir, exist_ok=True)
        self.tag_associations = {}
        self._refresh_file_and_tag_lists()
        # Reset tree to default
        self._rebuild_tree_from_serialized({
            'text': 'Project', 'tags': ['project_node'], 'open': True,
            'children': [ { 'text': 'Plants', 'tags': ['bucket_plants'], 'open': True, 'children': [] } ]
        })
        self._persist_state()
        show_toast(self.master, "New project initialized")

    def _open_project_folder(self):
        chosen = filedialog.askdirectory(title="Open Project Folder")
        if not chosen:
            return
        try:
            # Load tags
            tags_path = os.path.join(chosen, "tags.json")
            project_path = os.path.join(chosen, "project.json")
            excel_dir = os.path.join(chosen, "excel_files")

            if not os.path.exists(tags_path) and not os.path.exists(project_path) and not os.path.isdir(excel_dir):
                messagebox.showerror("Not a Project", "Selected folder does not appear to be a project (missing project.json/tags.json/excel_files).")
                return

            # Switch project pointers
            self.project_dir = chosen
            self.data_dir = excel_dir
            self.tags_path = tags_path
            os.makedirs(self.data_dir, exist_ok=True)

            # Load tags
            if os.path.exists(tags_path):
                self.tag_associations = safe_json_load(tags_path, {})
            else:
                self.tag_associations = {}

            # Load and rebuild tree
            if os.path.exists(project_path):
                data = safe_json_load(project_path, {})
                self._rebuild_tree_from_serialized(data)
            else:
                # Minimal default
                self._rebuild_tree_from_serialized({
                    'text': 'Project', 'tags': ['project_node'], 'open': True,
                    'children': [ { 'text': 'Plants', 'tags': ['bucket_plants'], 'open': True, 'children': [] } ]
                })

            # Update UI lists, save settings pointer
            self._refresh_file_and_tag_lists()
            self._persist_state()
            show_toast(self.master, f"Opened project")
            self._set_status(f"Opened: {self.project_dir}", 4000)
        except Exception as e:
            messagebox.showerror("Open Project Failed", f"{e}")

    def _safe_save_all(self):
        try:
            if self.current_xl_db:
                self.current_xl_db.save()
            self._persist_state()
            show_toast(self.master, "Saved")
        except Exception as e:
            messagebox.showerror("Save Failed", f"{e}")

    # ---------- Query tester ----------
    def _open_query_tester(self):
        win = tk.Toplevel(self.master)
        win.title("SPARQL Query Tester")
        win.geometry("820x520")
        frm = ttk.Frame(win, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Write a SPARQL query (results limited):", style="Header.TLabel").pack(anchor="w")
        txt = tk.Text(frm, height=12, font=('Consolas', 10))
        txt.pack(fill="both", expand=False, pady=(6,6))
        txt.insert("1.0", "SELECT * WHERE { ?s ?p ?o } LIMIT 25")
        bar = ttk.Frame(frm); bar.pack(fill="x")
        ttk.Button(bar, text="Run", command=lambda: run_query()).pack(side="left")
        ttk.Button(bar, text="Close", command=win.destroy).pack(side="right")
        cols = ("var","value")
        tree = ttk.Treeview(frm, columns=cols, show="headings", height=10)
        for c in cols: tree.heading(c, text=c.title())
        tree.pack(fill="both", expand=True, pady=(6,0))

        def run_query():
            q = txt.get("1.0", "end").strip()
            if not q:
                return
            self._with_progress(True)
            def job():
                try:
                    s = SPARQLWrapper(self.repo_var.get().strip())
                    s.setQuery(q)
                    s.setReturnFormat(JSON)
                    return s.query().convert()
                except Exception as e:
                    return e
            fut = self.executor.submit(job)
            def after(res):
                self._with_progress(False)
                tree.delete(*tree.get_children())
                if isinstance(res, Exception):
                    messagebox.showerror("Query Error", f"{res}")
                    return
                bindings = res.get("results", {}).get("bindings", [])
                for b in bindings:
                    for k,v in b.items():
                        tree.insert("", "end", values=(k, v.get("value","")))
                if not bindings:
                    show_toast(self.master, "No rows")
            self._track_future(fut, after)

# -------------------------
# Main
# -------------------------
def main():
    root = tk.Tk()
    app = LeanDigitalTwin(root)
    if app.initialization_ok:
        root.mainloop()

if __name__ == "__main__":
    main()