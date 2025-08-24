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

    def _build_graphical_model_tab(self, parent):
        parent.rowconfigure(1, weight=1)  # graph area takes most space
        parent.columnconfigure(1, weight=1)  # middle column (node selection) expands

        # TOP ROW: controls
        top_frame = ttk.Frame(parent)
        top_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        top_frame.columnconfigure(1, weight=1)

        # LEFT COLUMN: controls
        left = ttk.Frame(top_frame, padding=8)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        left.columnconfigure(0, weight=1)

        # 1. Semantic Network URL
        ttk.Label(left, text="Semantic Network URL:").grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.url_var = tk.StringVar(value=self.settings.get("repo_url", ""))
        url_entry = ttk.Entry(left, textvariable=self.url_var, width=40)
        url_entry.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        # 2. SPARQL Prefix
        ttk.Label(left, text="SPARQL Prefix:").grid(row=2, column=0, sticky="w", pady=(0, 2))
        prefix_entry = ttk.Entry(left, textvariable=self.prefix_var, width=40)
        prefix_entry.grid(row=3, column=0, sticky="ew", pady=(0, 8))

        # 3. Test Connection and Fetch All Nodes
        btn_frame = ttk.Frame(left)
        btn_frame.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(btn_frame, text="Test Connection", command=self._test_connection).pack(side="left", padx=(0, 4))
        ttk.Button(btn_frame, text="Fetch All Nodes", command=self._fetch_nodes).pack(side="left")

        # 4. SPARQL Query Tester
        ttk.Label(left, text="SPARQL Query:").grid(row=5, column=0, sticky="w", pady=(0, 2))
        self.query_text = tk.Text(left, height=6, width=40)
        self.query_text.grid(row=6, column=0, sticky="ew", pady=(0, 4))
        ttk.Button(left, text="Execute Query", command=self._execute_query).grid(row=7, column=0, sticky="ew", pady=(0, 8))

        # 5. Query Results
        ttk.Label(left, text="Query Results:").grid(row=8, column=0, sticky="w", pady=(0, 2))
        self.results_text = tk.Text(left, height=8, width=40)
        self.results_text.grid(row=9, column=0, sticky="ew", pady=(0, 8))

        # MIDDLE COLUMN: Select Nodes to Visualize
        middle = ttk.Frame(top_frame, padding=8)
        middle.grid(row=0, column=1, sticky="nsew", padx=10)
        middle.columnconfigure(0, weight=1)
        middle.rowconfigure(1, weight=1)

        ttk.Label(middle, text="Select Nodes to Visualize:", style='Header.TLabel').grid(row=0, column=0, sticky="w", pady=(0, 4))
        
        # Filter for nodes
        filter_frame = ttk.Frame(middle)
        filter_frame.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        ttk.Label(filter_frame, text="Filter:").pack(side="left", padx=(0, 6))
        self.node_display_filter = tk.StringVar()
        filter_entry = ttk.Entry(filter_frame, textvariable=self.node_display_filter)
        filter_entry.pack(side="left", fill="x", expand=True)
        self.node_display_filter.trace_add('write', self._filter_node_display)

        # Node listbox
        self.node_listbox_display = tk.Listbox(middle, height=20, exportselection=False)
        self.node_listbox_display.grid(row=2, column=0, sticky="nsew", pady=(0, 4))
        self.node_listbox_display.bind("<Double-1>", self._on_node_double_click)

        # Scrollbar for node listbox
        node_scrollbar = ttk.Scrollbar(middle, orient="vertical", command=self.node_listbox_display.yview)
        node_scrollbar.grid(row=2, column=1, sticky="ns")
        self.node_listbox_display.configure(yscrollcommand=node_scrollbar.set)

        # Buttons for node selection
        node_btn_frame = ttk.Frame(middle)
        node_btn_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        ttk.Button(node_btn_frame, text="Select All", command=self._select_all_nodes).pack(side="left")
        ttk.Button(node_btn_frame, text="Clear Selection", command=self._clear_node_selection).pack(side="left", padx=6)
        ttk.Button(node_btn_frame, text="Visualize Selected", command=self._visualize_selected_nodes).pack(side="right")

        # RIGHT COLUMN: Node Properties
        right = ttk.Frame(top_frame, padding=8)
        right.grid(row=0, column=2, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        ttk.Label(right, text="Node Properties (double-click node to view):").grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.properties_text = tk.Text(right, height=30, width=50)
        self.properties_text.grid(row=1, column=0, sticky="nsew", pady=(0, 8))

        # Scrollbar for properties
        prop_scrollbar = ttk.Scrollbar(right, orient="vertical", command=self.properties_text.yview)
        prop_scrollbar.grid(row=1, column=1, sticky="ns")
        self.properties_text.configure(yscrollcommand=prop_scrollbar.set)

        # MAIN GRAPH AREA (below controls)
        graph_frame = ttk.LabelFrame(parent, text="Graph Visualization", padding=10)
        graph_frame.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=6, pady=6)
        graph_frame.columnconfigure(0, weight=1)
        graph_frame.rowconfigure(0, weight=1)

        # Create matplotlib figure and canvas
        self.fig, self.ax = plt.subplots(figsize=(10, 8))
        self.canvas = FigureCanvasTkAgg(self.fig, graph_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")

        # Bind canvas for node selection
        self.canvas.mpl_connect('button_press_event', self._on_graph_click)

        # Initialize empty graph
        self._update_graph()

    def _filter_node_display(self, *args):
        """Filter the node display listbox based on search term"""
        search_term = self.node_display_filter.get().lower().strip()
        self.node_listbox_display.delete(0, tk.END)
        
        if not hasattr(self, 'fetched_nodes') or not self.fetched_nodes.get('all'):
            return
            
        for node in self.fetched_nodes['all']:
            if not search_term or search_term in node.lower():
                self.node_listbox_display.insert(tk.END, node)

    def _select_all_nodes(self):
        """Select all nodes in the display listbox"""
        self.node_listbox_display.select_set(0, tk.END)

    def _clear_node_selection(self):
        """Clear selection in the node display listbox"""
        self.node_listbox_display.selection_clear(0, tk.END)

    def _visualize_selected_nodes(self):
        """Visualize the currently selected nodes"""
        selected_indices = self.node_listbox_display.curselection()
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select nodes to visualize.")
            return
            
        selected_nodes = [self.node_listbox_display.get(i) for i in selected_indices]
        self.selected_nodes_for_graph = selected_nodes
        self._update_graph()
        self._set_status(f"Visualizing {len(selected_nodes)} selected nodes", 3000)

    def _on_node_double_click(self, event=None):
        """Handle double-click on node in the display listbox"""
        if not self.node_listbox_display.curselection():
            return
        node_name = self.node_listbox_display.get(self.node_listbox_display.curselection())
        self._display_node_properties(node_name)

    def _display_node_properties(self, node_name):
        """Display properties of selected node in properties text area"""
        self.properties_text.config(state='normal')
        self.properties_text.delete(1.0, tk.END)

        self.properties_text.insert(tk.END, f"Node: {node_name}\n\n")

        # Fetch properties from SPARQL
        prefix = self.prefix_var.get().strip()
        if not prefix:
            self.properties_text.insert(tk.END, "No SPARQL prefix configured.")
            self.properties_text.config(state='disabled')
            return

        query = f"""{prefix}
        SELECT ?p ?o WHERE {{
            ex:{node_name} ?p ?o .
            FILTER (isLiteral(?o))
        }} ORDER BY ?p"""

        fut, _ = self._run_sparql_query_bg(query, timeout=20)

        def after(res):
            self.properties_text.config(state='normal')
            if isinstance(res, Exception) or not res:
                self.properties_text.insert(tk.END, "No properties found or error occurred.")
            else:
                for r in res:
                    prop = r['p']['value'].split('#')[-1]
                    value = r['o']['value']
                    self.properties_text.insert(tk.END, f"{prop}: {value}\n")
            self.properties_text.config(state='disabled')

        self._track_future(fut, after)

    def _update_node_lists(self, nodes):
        # Update the node display listbox in Graphical Model tab
        if hasattr(self, "node_listbox_display") and self.node_listbox_display.winfo_exists():
            self.node_listbox_display.delete(0, tk.END)
            for n in nodes:
                self.node_listbox_display.insert(tk.END, n)

        # Functionalities combobox (optional)
        if hasattr(self, "node_combobox"):
            self.node_combobox['values'] = nodes

        self._update_badges()

    def _update_badges(self):
        a = len(self.fetched_nodes.get('all', []))
        e = len(self.fetched_nodes.get('equipment', []))
        s = len(self.fetched_nodes.get('sub_equipment', []))
        t = len(self.fetched_nodes.get('asset', []))
        if hasattr(self, 'badge_nodes'):
            self.badge_nodes.config(text=f"Nodes: {a} | Eq:{e} Sub:{s} Asset:{t}")

    def _update_graph(self):
        """Update the graph visualization"""
        self.ax.clear()

        if not hasattr(self, 'selected_nodes_for_graph') or not self.selected_nodes_for_graph:
            self.ax.text(0.5, 0.5, 'No nodes selected for visualization\nSelect nodes and click "Visualize Selected"',
                         ha='center', va='center', transform=self.ax.transAxes)
            self.canvas.draw()
            return

        # Create graph from selected nodes
        G = nx.Graph()
        G.add_nodes_from(self.selected_nodes_for_graph)

        # Add edges if we have relationship data
        # You can extend this to add edges based on your data

        # Draw the graph
        pos = nx.spring_layout(G, k=1, iterations=50)
        self.node_positions = pos

        nx.draw(G, pos, ax=self.ax, with_labels=True,
                node_color='lightblue', node_size=1000,
                font_size=8, font_weight='bold')

        self.canvas.draw()

    def _on_graph_click(self, event):
        """Handle clicks on graph nodes to show properties"""
        if event.inaxes != self.ax:
            return

        # Find clicked node
        for node, pos in self.node_positions.items():
            if abs(event.xdata - pos[0]) < 0.1 and abs(event.ydata - pos[1]) < 0.1:
                self._display_node_properties(node)
                break

    def _test_connection(self):
        url = self.url_var.get().strip()
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

    def _fetch_nodes(self):
        url = self.url_var.get().strip()
        prefix_str = self.prefix_var.get().strip()
        if not url:
            messagebox.showwarning("Missing URL", "Enter a repository URL.")
            return
            
        self._with_progress(True)
        self._set_status("Fetching all nodes from semantic network…")

        import re
        base = None
        m = re.search(r"<([^>]+)>", prefix_str)
        if m:
            base = m.group(1).strip()

        # First attempt: with namespace filter if we have one
        if base:
            query = f'''
                SELECT DISTINCT ?resource WHERE {{
                  {{ ?resource ?p ?o . }} UNION {{ ?s ?p ?resource . }}
                  FILTER(ISIRI(?resource) && STRSTARTS(STR(?resource), "{base}"))
                }} ORDER BY ?resource
            '''
        else:
            query = '''
                SELECT DISTINCT ?resource WHERE {
                  { ?resource ?p ?o . } UNION { ?s ?p ?resource . }
                  FILTER(ISIRI(?resource))
                } ORDER BY ?resource LIMIT 5000
            '''

        fut, to = self._run_sparql_query_bg(query, timeout=45)

        def after_first(res):
            if isinstance(res, Exception) or not res:
                # Fallback: fetch all IRIs without prefix filtering (bounded)
                fallback_q = '''
                    SELECT DISTINCT ?resource WHERE {
                      { ?resource ?p ?o . } UNION { ?s ?p ?resource . }
                      FILTER(ISIRI(?resource))
                    } ORDER BY ?resource LIMIT 5000
                '''
                fut2, _ = self._run_sparql_query_bg(fallback_q, timeout=45)
                self._track_future(fut2, lambda r2: self._after_fetch_nodes_resilient(r2, base))
            else:
                self._after_fetch_nodes_resilient(res, base)

        self._track_future(fut, after_first)

    def _after_fetch_nodes_resilient(self, res, base_hint=None):
        self._with_progress(False)
        if isinstance(res, Exception):
            messagebox.showerror(
                "Fetch Failed",
                f"Could not fetch nodes.\n\nHint:\n- Ensure the Semantic Network URL points to the SPARQL endpoint (often ends with /sparql)\n- Verify the SPARQL Prefix namespace matches your data\n\nDetails:\n{res}"
            )
            self._set_status("Failed to fetch nodes.", 4000)
            return

        # Normalize to local names
        def localize(u):
            u = u.get("resource", {}).get("value", "")
            if not u: return None
            part = u.split('#')[-1]
            part = part.split('/')[-1] if '/' in part else part
            return part or None

        nodes = [localize(r) for r in (res or [])]
        nodes = sorted({n for n in nodes if n})

        # If filtered by base produced nothing, try deriving dominant namespace and warn
        if not nodes and res:
            try:
                uris = [r["resource"]["value"] for r in res if "resource" in r]

                def ns(u):
                    return u.rsplit('#', 1)[0] if '#' in u else u.rsplit('/', 1)[0]

                from collections import Counter
                common_ns = Counter(ns(u) for u in uris).most_common(1)[0][0]
                messagebox.showwarning(
                    "Prefix Mismatch?",
                    f"No results matched the configured prefix.\nDetected common namespace:\n{common_ns}\n\nUpdate SPARQL Prefix to this base."
                )
            except Exception:
                pass

        self.fetched_nodes['all'] = nodes
        self._update_node_lists(nodes)
        show_toast(self.master, f"Fetched {len(nodes)} nodes")
        self._set_status("Nodes fetched.", 3000)

    def _run_sparql_query_bg(self, query, timeout=20):
        url = self.url_var.get().strip()
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

    def _execute_query(self):
        """Execute SPARQL query from the query text widget and display results"""
        query = self.query_text.get("1.0", "end").strip()
        if not query:
            messagebox.showwarning("No Query", "Please enter a SPARQL query.")
            return

        # Clear previous results
        self.results_text.delete("1.0", "end")
        self.results_text.insert("1.0", "Executing query...\n")
        self._with_progress(True)

        try:
            fut, _ = self._run_sparql_query_bg(query, timeout=30)

            def after_query(res):
                self._with_progress(False)
                self.results_text.delete("1.0", "end")

                if isinstance(res, Exception):
                    self.results_text.insert("1.0", f"Query Error: {str(res)}")
                elif not res:
                    self.results_text.insert("1.0", "No results returned.")
                else:
                    # Format results nicely
                    result_text = f"Query returned {len(res)} results:\n\n"
                    for i, result in enumerate(res[:100]):  # Limit to first 100 results
                        result_text += f"Result {i + 1}:\n"
                        for key, value in result.items():
                            result_text += f"  {key}: {value.get('value', 'N/A')}\n"
                        result_text += "\n"

                    if len(res) > 100:
                        result_text += f"... and {len(res) - 100} more results (showing first 100)"

                    self.results_text.insert("1.0", result_text)

                self._set_status("Query completed", 3000)

            self._track_future(fut, after_query)

        except Exception as e:
            self._with_progress(False)
            self.results_text.delete("1.0", "end")
            self.results_text.insert("1.0", f"Error executing query: {str(e)}")
            self._set_status("Query failed", 3000)

    def _with_progress(self, running=True):
        if running:
            if hasattr(self, 'progress'):
                self.progress.pack(side="right")
                self.progress.start(12)
        else:
            if hasattr(self, 'progress'):
                self.progress.stop()
                self.progress.pack_forget()

    def _set_status(self, text, ms=None):
        if hasattr(self, 'status_lbl'):
            self.status_lbl.config(text=text)
            if ms:
                self.master.after(ms, lambda: self.status_lbl.config(text="Ready"))

    def _build_ui(self):
        self.master.title(APP_NAME)
        self.master.geometry("1460x900")
        menubar = tk.Menu(self.master)

        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Open Project…", command=self._open_project, accelerator="Ctrl+O")
        file_menu.add_command(label="Save Project…", command=self._save_project, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        view_menu = tk.Menu(menubar, tearoff=False)
        view_menu.add_command(label="Toggle Theme", command=self._toggle_theme)
        menubar.add_cascade(label="View", menu=view_menu)

        self.master.config(menu=menubar)
        self.master.bind("<Control-o>", lambda e: self._open_project())
        self.master.bind("<Control-s>", lambda e: self._safe_save_all())

        # Notebook for tabs
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(10, 10))

        # Tabs
        self.tab_graphical = ttk.Frame(self.nb, padding=10)
        self.tab_asset_hierarchy = ttk.Frame(self.nb, padding=10)
        self.tab_functionalities = ttk.Frame(self.nb, padding=10)
        self.tab_excel = ttk.Frame(self.nb, padding=10)

        # Add tabs
        self.nb.add(self.tab_graphical, text="Graphical Model")
        self.nb.add(self.tab_asset_hierarchy, text="Asset Hierarchy")
        self.nb.add(self.tab_functionalities, text="Functionalities")
        self.nb.add(self.tab_excel, text="Datasheet Editor")

        # Build tab contents
        self._build_graphical_model_tab(self.tab_graphical)
        # Note: Other tab methods would be implemented here
        # self._build_plant_hierarchy_tab(self.tab_asset_hierarchy)
        # self._build_functionalities_tab(self.tab_functionalities)
        # self._build_datasheet_editor_tab(self.tab_excel)

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
        ttk.Button(self.bottom, text="Open Project…", command=self._open_project).pack(side="left")
        ttk.Button(self.bottom, text="Save Project…", command=self._save_project).pack(side="left", padx=6)
        ttk.Button(self.bottom, text="Theme", command=self._toggle_theme).pack(side="right")
        ttk.Button(self.bottom, text="Save", command=self._safe_save_all).pack(side="right", padx=(0, 6))

    def _configure_styles(self):
        self.style = ttk.Style(self.master)
        try:
            self.style.theme_use('clam')
        except Exception:
            pass

        is_dark = self.settings.get("theme") == "dark"
        if is_dark:
            self.theme = {
                "bg": "#0f1115",
                "panel": "#171a21",
                "fg": "#e5e7eb",
                "muted": "#9aa4b2",
                "accent": "#60a5fa",
                "accent_on": "#ffffff",
                "border": "#2a2f3a",
                "input_bg": "#0f1218",
                "input_fg": "#e5e7eb",
                "select_bg": "#2563eb",
                "select_fg": "#ffffff",
                "edge": "#6b7280",
            }
        else:
            self.theme = {
                "bg": "#f6f7fb",
                "panel": "#ffffff",
                "fg": "#1f2328",
                "muted": "#6b7280",
                "accent": "#2357ff",
                "accent_on": "#ffffff",
                "border": "#d0d7de",
                "input_bg": "#ffffff",
                "input_fg": "#1f2328",
                "select_bg": "#2357ff",
                "select_fg": "#ffffff",
                "edge": "#8a9199",
            }

        t = self.theme
        self.master.configure(bg=t["bg"])

        # TTK core
        self.style.configure("TFrame", background=t["panel"])
        self.style.configure("TLabel", background=t["panel"], foreground=t["fg"], font=('Segoe UI', 10))
        self.style.configure("Header.TLabel", background=t["panel"], foreground=t["fg"], font=('Segoe UI', 11, 'bold'))
        self.style.configure("Info.TLabel", background=t["panel"], foreground=t["muted"],
                             font=('Segoe UI', 9, 'italic'))
        self.style.configure("Badge.TLabel", background=t["panel"], foreground=t["accent"],
                             font=('Segoe UI', 9, 'bold'))

        self.style.configure("TButton", padding=6, font=('Segoe UI', 10, 'bold'), background=t["panel"],
                             foreground=t["fg"])
        self.style.map("TButton",
                       foreground=[("active", t["accent_on"]), ("pressed", t["accent_on"])],
                       background=[("active", t["accent"]), ("pressed", t["accent"])],
                       relief=[("pressed", "sunken"), ("!pressed", "raised")]
                       )

        # Inputs
        self.style.configure("TEntry", fieldbackground=t["input_bg"], foreground=t["input_fg"])
        self.style.configure("TCombobox", fieldbackground=t["input_bg"], foreground=t["input_fg"])

        # Notebook / tabs
        self.style.configure("TNotebook", background=t["bg"])
        self.style.configure("TNotebook.Tab",
                             background=t["panel"], foreground=t["fg"], padding=[10, 6], focuscolor=t["accent"])
        self.style.map("TNotebook.Tab",
                       background=[("selected", t["accent"])],
                       foreground=[("selected", t["accent_on"])],
                       expand=[("selected", [2, 2, 2, 0])]
                       )

        # Progress bar
        self.style.configure("Horizontal.TProgressbar", background=t["accent"], troughcolor=t["border"])

    def _probe_excel(self):
        try:
            app_check = xw.App(visible=True, add_book=False)
            app_check.quit()
            return True
        except Exception as e:
            self.master.withdraw()
            messagebox.showerror("Excel Not Found",
                f"Could not connect to Microsoft Excel. Ensure it's installed.\n\n{e}")
            return False

    def _toggle_theme(self):
        self.settings["theme"] = "dark" if self.settings.get("theme") != "dark" else "light"
        self._configure_styles()
        show_toast(self.master, f"Theme: {self.settings['theme'].title()}")

    def _wire_shortcuts(self):
        self.master.bind("<Control-s>", lambda e: self._safe_save_all())
        self.master.bind("<F5>", lambda e: self._update_data_model())

    def _open_project(self):
        path = filedialog.askopenfilename(title="Open Project", filetypes=[("Project Zip", "*.zip")])
        if not path: return
        try:
            with zipfile.ZipFile(path, "r") as z:
                for name in z.namelist():
                    if name.endswith("settings.json"):
                        self.settings = json.loads(z.read(name).decode("utf-8"))
                    elif name.endswith("tags.json"):
                        self.tag_associations = json.loads(z.read(name).decode("utf-8"))
                    elif name.startswith("excel_files/") and not name.endswith("/"):
                        target = os.path.join(DATA_DIR, os.path.basename(name))
                        with open(target, "wb") as f:
                            f.write(z.read(name))
            self.url_var.set(self.settings.get("repo_url", ""))
            self.prefix_var.set(self.settings.get("sparql_prefix", ""))
            self._configure_styles()
            show_toast(self.master, "Project opened")
        except Exception as e:
            messagebox.showerror("Open Failed", f"{e}")

    def _save_project(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".zip",
            filetypes=[("Project Zip", "*.zip")],
            title="Save Project (tags.json + library)"
        )
        if not path: return
        try:
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("settings.json", json.dumps(self.settings, indent=2))
                z.writestr("tags.json", json.dumps(self.tag_associations, indent=2))
                for f in os.listdir(DATA_DIR):
                    fp = os.path.join(DATA_DIR, f)
                    if os.path.isfile(fp):
                        z.write(fp, arcname=os.path.join("excel_files", f))
            show_toast(self.master, "Project saved")
        except Exception as e:
            messagebox.showerror("Save Failed", f"{e}")

    def _safe_save_all(self):
        try:
            self._persist_state()
            show_toast(self.master, "Saved")
        except Exception as e:
            messagebox.showerror("Save Failed", f"{e}")

    def _persist_state(self):
        # settings
        self.settings["repo_url"] = self.url_var.get().strip()
        self.settings["sparql_prefix"] = self.prefix_var.get().strip()
        safe_json_dump(SETTINGS_PATH, self.settings)
        safe_json_dump(TAGS_PATH, self.tag_associations)

    def _refresh_file_and_tag_lists(self):
        # This would be implemented to refresh file and tag lists
        pass

    def on_close(self):
        try:
            self._persist_state()
        finally:
            self.stop_flag = True
            self.executor.shutdown(wait=False, cancel_futures=True)
            self.master.destroy()

    def _update_data_model(self, event=None):
        # This would be implemented to update the data model
        pass

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