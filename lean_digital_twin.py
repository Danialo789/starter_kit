import os
import shutil
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import re

# --- PLATFORM-SPECIFIC CHECK ---
if os.name != 'nt':
    messagebox.showerror("Unsupported OS",
                         "This version of the application uses Windows-specific features (pywin32) to embed Excel and can only run on Windows.")
    exit()

# --- STABLE DEPENDENCIES ---
import xlwings as xw
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from SPARQLWrapper import SPARQLWrapper, JSON
import win32gui
import win32con
import win32api


class SelectionDialog(tk.Toplevel):
    """Dialog for selecting multiple items from a list."""
    
    def __init__(self, parent, title="Select Items", item_list=None, node_type=None):
        super().__init__(parent)
        self.parent = parent
        self.result = None
        self.node_type = node_type
        self.title(title)
        self.geometry("500x400")
        self.transient(parent)
        self.grab_set()
        
        # Center the dialog
        self.geometry("+%d+%d" % (parent.winfo_rootx() + 50, parent.winfo_rooty() + 50))
        
        self._create_widgets(item_list or [])
        
    def _create_widgets(self, item_list):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # Header with node type information
        if self.node_type:
            ttk.Label(main_frame, text=f"Select {self.node_type.replace('_', ' ').title()}s from GraphDB Repository:", 
                     font=('Segoe UI', 10, 'bold')).pack(anchor="w", pady=(0, 5))
        else:
            ttk.Label(main_frame, text="Select items to add:").pack(anchor="w", pady=(0, 5))
        
        # Search frame
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(search_frame, text="Filter:").pack(side="left", padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side="left", fill="x", expand=True)
        self.search_var.trace('w', self._filter_items)
        
        # Store original items for filtering
        self.original_items = item_list[:]
        
        # Listbox with scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.listbox = tk.Listbox(list_frame, selectmode="multiple", yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)
        
        # Populate listbox
        self._populate_listbox(item_list)
        
        # Selection info
        self.selection_info = ttk.Label(main_frame, text="0 items selected", font=('Segoe UI', 9, 'italic'))
        self.selection_info.pack(anchor="w", pady=(0, 10))
        
        # Bind selection event
        self.listbox.bind("<<ListboxSelect>>", self._update_selection_info)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x")
        
        ttk.Button(button_frame, text="Select All", command=self._select_all).pack(side="left", padx=(0, 5))
        ttk.Button(button_frame, text="Clear All", command=self._clear_all).pack(side="left", padx=(0, 5))
        ttk.Button(button_frame, text="Cancel", command=self._cancel_clicked).pack(side="right")
        ttk.Button(button_frame, text="OK", command=self._ok_clicked).pack(side="right", padx=(5, 0))
        
    def _populate_listbox(self, items):
        self.listbox.delete(0, tk.END)
        for item in items:
            self.listbox.insert(tk.END, item)
            
    def _filter_items(self, *args):
        search_term = self.search_var.get().lower()
        if not search_term:
            filtered_items = self.original_items
        else:
            filtered_items = [item for item in self.original_items if search_term in item.lower()]
        
        self._populate_listbox(filtered_items)
        self._update_selection_info()
        
    def _update_selection_info(self, event=None):
        selected_count = len(self.listbox.curselection())
        total_count = self.listbox.size()
        self.selection_info.config(text=f"{selected_count} of {total_count} items selected")
        
    def _select_all(self):
        self.listbox.select_set(0, tk.END)
        self._update_selection_info()
        
    def _clear_all(self):
        self.listbox.selection_clear(0, tk.END)
        self._update_selection_info()
        
    def _ok_clicked(self):
        selected_indices = self.listbox.curselection()
        self.result = [self.listbox.get(i) for i in selected_indices]
        self.destroy()
        
    def _cancel_clicked(self):
        self.result = None
        self.destroy()


class LeanDigitalTwin(tk.Tk):
    """
    An application for interacting with a GraphDB repository, visualizing data models,
    and linking data to an embedded Excel datasheet with drag-and-drop.
    [VERSION: XLWINGS - Final with Template Cloning]
    """

    def __init__(self):
        super().__init__()
        self.initialization_ok = True

        try:
            app_check = xw.App(visible=False, add_book=False)
            app_check.quit()
        except Exception as e:
            self.withdraw()
            messagebox.showerror("Excel Not Found",
                                 "Could not connect to Microsoft Excel. Please ensure it is installed and accessible.\n\n"
                                 f"Error: {e}")
            self.initialization_ok = False
            return

        self.title("LEAN Digital Twin (Final Edition)")
        self.geometry("1400x850")

        self.properties = []
        self.current_excel_path = None
        self.tag_associations = {}
        self.graph = nx.DiGraph()
        self.xl_app = None
        self.current_xl_db = None
        self.active_node = None

        # Drag and Drop related variables
        self.is_dragging = False
        self.drag_data = None
        self.drag_source_type = None

        # Tree view related variables
        self.tree = None
        self.current_item_id = None
        
        # Store fetched nodes by type for filtered selection
        self.fetched_nodes = {
            'all': [],
            'equipment': [],
            'sub_equipment': [],
            'asset': []
        }

        self.storage_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excel_files")
        os.makedirs(self.storage_dir, exist_ok=True)

        self._configure_styles()
        self._create_main_ui()

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.after(100, self._update_file_lists)
        self.after(100, self._update_tag_lists)

    def on_closing(self):
        """Handle the window close event, prompting to save changes if necessary."""
        if self.xl_app and self.current_xl_db:
            if not self.current_xl_db.api.Saved:
                response = messagebox.askyesnocancel(
                    "Save Changes?",
                    f"Do you want to save the changes you made to '{self.current_xl_db.name}'?",
                    icon='warning'
                )
                if response is True:
                    self.current_xl_db.save()
                elif response is None:
                    return
            else:
                if not messagebox.askyesno("Exit Application",
                                           "An Excel file is currently embedded. Are you sure you want to exit?"):
                    return

        self._cleanup_xlwings()
        self.destroy()

    def _cleanup_xlwings(self):
        if self.xl_app:
            try:
                if self.xl_app.hwnd:
                    win32gui.SetParent(self.xl_app.hwnd, 0)
                self.xl_app.quit()
                self.xl_app = None
                self._update_status("Excel instance closed.", 2000)
            except Exception as e:
                print(f"Could not quit Excel gracefully: {e}")

    def _configure_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure('TLabel', font=('Segoe UI', 10))
        self.style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=5)
        self.style.configure('TEntry', font=('Segoe UI', 10))
        self.style.configure('TNotebook.Tab', font=('Segoe UI', 10, 'bold'), padding=[10, 5])
        self.style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))
        self.style.configure('Info.TLabel', font=('Segoe UI', 9, 'italic'))
        self.style.configure('ActiveNode.TLabel', font=('Segoe UI', 11, 'bold'), foreground='blue')

    def _create_main_ui(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(expand=True, fill="both")

        self.main_notebook = ttk.Notebook(main_frame)
        self.main_notebook.pack(expand=True, fill="both")

        self.tab_graphical = ttk.Frame(self.main_notebook, padding=10)
        self.tab_graphdb = ttk.Frame(self.main_notebook, padding=10)
        self.tab_excel = ttk.Frame(self.main_notebook, padding=10)
        self.tab_functionalities = ttk.Frame(self.main_notebook, padding=10)
        self.tab_asset_hierarchy = ttk.Frame(self.main_notebook, padding=10)

        self.main_notebook.add(self.tab_graphical, text="Graphical Model")
        self.main_notebook.add(self.tab_graphdb, text="GraphDB → Excel")
        self.main_notebook.add(self.tab_excel, text="Datasheet Editor")
        self.main_notebook.add(self.tab_functionalities, text="Functionalities")
        self.main_notebook.add(self.tab_asset_hierarchy, text="Asset Hierarchy")

        self._build_graphical_model_tab(self.tab_graphical)
        self._build_graphdb_tab(self.tab_graphdb)
        self._build_datasheet_editor_tab(self.tab_excel)
        self._build_functionalities_tab(self.tab_functionalities)
        self._build_plant_hierarchy_tab(self.tab_asset_hierarchy)

        self.status_bar = ttk.Label(self, text="Ready", relief=tk.SUNKEN, anchor='w', padding=5)
        self.status_bar.pack(side="bottom", fill="x")

    def _update_status(self, message, clear_after_ms=None):
        self.status_bar.config(text=message)
        if clear_after_ms:
            self.after(clear_after_ms, lambda: self.status_bar.config(text="Ready"))

    def _build_graphical_model_tab(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)

        controls_frame = ttk.Frame(parent)
        controls_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        ttk.Label(controls_frame, text="Select Node(s) to Visualize:", style='Header.TLabel').pack(anchor="w")

        self.node_listbox_display = tk.Listbox(controls_frame, height=6, selectmode="extended", exportselection=False)
        self.node_listbox_display.pack(fill="x", expand=True, pady=5)
        self.node_listbox_display.bind("<<ListboxSelect>>", lambda event: self._update_data_model(event))

        button_bar = ttk.Frame(controls_frame)
        button_bar.pack(fill="x", pady=5)

        ttk.Button(button_bar, text="Refresh Model", command=self._update_data_model).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_bar, text="Clear Graph", command=self._clear_graph).pack(side=tk.LEFT)

        self.graph_frame = ttk.Frame(parent, relief=tk.SUNKEN, borderwidth=1)
        self.graph_frame.grid(row=1, column=0, sticky="nsew")

    def _build_graphdb_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        conn_frame = ttk.LabelFrame(parent, text="1. Connection", padding=10)
        conn_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        conn_frame.columnconfigure(1, weight=1)

        ttk.Label(conn_frame, text="Repo URL:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.entry_repo = ttk.Entry(conn_frame)
        self.entry_repo.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(conn_frame, text="SPARQL Prefix:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.entry_prefix = ttk.Entry(conn_frame)
        self.entry_prefix.insert(0, "PREFIX ex: <http://example.org/pumps#>")
        self.entry_prefix.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        node_select_frame = ttk.LabelFrame(parent,
                                           text="2. Master Node List (Double-click to set Active Node manually)",
                                           padding=10)
        node_select_frame.grid(row=1, column=0, sticky="nsew", pady=10)
        node_select_frame.rowconfigure(1, weight=1)
        node_select_frame.columnconfigure(0, weight=1)

        # Button frame for fetch operations
        fetch_button_frame = ttk.Frame(node_select_frame)
        fetch_button_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        
        ttk.Button(fetch_button_frame, text="Fetch All Nodes", command=self._fetch_nodes).pack(side="left", padx=(0, 5))
        ttk.Button(fetch_button_frame, text="Categorize Nodes", command=self._categorize_nodes).pack(side="left")
        
        self.node_listbox = tk.Listbox(node_select_frame, height=6, exportselection=False)
        self.node_listbox.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.node_listbox.bind("<Double-1>", lambda event: self._on_node_manual_select(event))

    def _build_datasheet_editor_tab(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(1, weight=1)

        left_pane = ttk.Frame(parent, padding=5)
        left_pane.grid(row=0, column=0, rowspan=2, sticky="ns", pady=5)
        left_pane.rowconfigure(3, weight=1)

        ttk.Label(left_pane, text="Select a Tag to View:", style='Header.TLabel').grid(row=0, column=0, sticky="w",
                                                                                       pady=(0, 5))

        self.tag_selector_combobox = ttk.Combobox(left_pane, state="readonly")
        self.tag_selector_combobox.grid(row=1, column=0, sticky="ew")
        self.tag_selector_combobox.bind("<<ComboboxSelected>>",
                                        lambda event: self._on_tag_selected_in_editor_tab(event))

        ttk.Label(left_pane, text="Datasheets for Selected Tag:", style='Header.TLabel').grid(row=2, column=0,
                                                                                              sticky="w", pady=(10, 5))

        self.datasheet_listbox_for_tag = tk.Listbox(left_pane, height=15, exportselection=False)
        self.datasheet_listbox_for_tag.grid(row=3, column=0, sticky="nsew")
        self.datasheet_listbox_for_tag.bind("<Double-1>", lambda event: self._load_file_from_list(event))

        main_pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        main_pane.grid(row=1, column=1, sticky="nsew", padx=10, pady=5)

        table_container = ttk.Frame(main_pane)
        main_pane.add(table_container, weight=3)
        table_container.rowconfigure(1, weight=1)
        table_container.columnconfigure(0, weight=1)

        sheet_controls_frame = ttk.Frame(table_container)
        sheet_controls_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        ttk.Label(sheet_controls_frame, text="Activate Sheet:").pack(side=tk.LEFT, padx=(0, 5))

        self.sheet_selector_combobox = ttk.Combobox(sheet_controls_frame, state="readonly")
        self.sheet_selector_combobox.pack(side=tk.LEFT, fill="x", expand=True)
        self.sheet_selector_combobox.bind("<<ComboboxSelected>>", lambda event: self._on_sheet_selected(event))

        self.excel_frame = ttk.Frame(table_container, relief=tk.SUNKEN, borderwidth=1)
        self.excel_frame.grid(row=1, column=0, sticky="nsew")
        self.excel_frame.bind("<Configure>", self._resize_excel_window)

        info_frame = ttk.Frame(main_pane, padding=10)
        info_frame.columnconfigure(0, weight=1)
        main_pane.add(info_frame, weight=1)

        ttk.Label(info_frame, text="Tag Information", style='Header.TLabel').grid(row=0, column=0, sticky="w",
                                                                                  pady=(0, 5))
        self.tag_text_embed = tk.Text(info_frame, height=4, width=30, font=('Segoe UI', 10), relief=tk.SOLID,
                                      borderwidth=1, state='disabled')
        self.tag_text_embed.grid(row=1, column=0, sticky="ew")

        active_node_frame = ttk.LabelFrame(info_frame, text="Active Node for Mapping", padding=10)
        active_node_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.active_node_display_label = ttk.Label(active_node_frame, text="None Selected", style='ActiveNode.TLabel')
        self.active_node_display_label.pack(pady=2)

        ttk.Label(info_frame, text="Associated Node Properties (Double-click to preview):", style='Header.TLabel').grid(
            row=3, column=0, sticky="w", pady=(10, 5))
        self.properties_text_embed = tk.Text(info_frame, height=8, width=30, font=('Segoe UI', 10), relief=tk.SOLID,
                                             borderwidth=1, state='disabled')
        self.properties_text_embed.grid(row=4, column=0, sticky="nsew")
        self.properties_text_embed.bind("<Double-1>", self._on_property_double_click)

        paste_frame = ttk.LabelFrame(info_frame, text="Drag-and-Drop Live Data", padding=10)
        paste_frame.grid(row=5, column=0, sticky="nsew", pady=(10, 0))
        paste_frame.columnconfigure(0, weight=1)

        prop_display_frame = ttk.LabelFrame(paste_frame, text="Selected Property")
        prop_display_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.selected_property_display = ttk.Label(prop_display_frame, text="(None)", font=('Segoe UI', 10, 'bold'),
                                                   foreground='navy')
        self.selected_property_display.pack(padx=5, pady=5)

        self.live_value_button = ttk.Button(paste_frame, text="Copy Value: (none)")
        self.live_value_button.value = None
        self.live_value_button.configure(
            command=lambda: self._initiate_drag_copy(self.live_value_button.value, 'value'))
        self.live_value_button.grid(row=1, column=0, sticky="ew", pady=(5, 5))

        self.live_unit_button = ttk.Button(paste_frame, text="Copy Unit: (none)")
        self.live_unit_button.value = None
        self.live_unit_button.configure(command=lambda: self._initiate_drag_copy(self.live_unit_button.value, 'unit'))
        self.live_unit_button.grid(row=2, column=0, sticky="ew", pady=5)

    def _build_plant_hierarchy_tab(self, parent):
        """Build the Plant Hierarchy tab with a tree view for asset management."""
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        
        # Create tree view
        tree_frame = ttk.Frame(parent)
        tree_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        
        # Tree with scrollbars
        self.tree = ttk.Treeview(tree_frame)
        self.tree.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Initialize tree with root
        root_item = self.tree.insert("", "end", text="Plant Hierarchy", open=True, tags=('root',))
        
        # Bind right-click event
        self.tree.bind("<Button-3>", self._show_context_menu)

    def _show_context_menu(self, event):
        """Determines which context menu to show based on the clicked item's tag."""
        # Identify the item that was clicked
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        # Select the clicked item in the tree
        self.tree.selection_set(item_id)
        self.current_item_id = item_id

        # Get the tag of the item to determine its type
        item_tags = self.tree.item(item_id, "tags")
        item_type = item_tags[0] if item_tags else None

        # Create the appropriate context menu
        context_menu = tk.Menu(self, tearoff=0)

        if item_type == 'root':
            context_menu.add_command(label="Create Plant", command=self._create_plant)
        elif item_type == 'plant_node':
            context_menu.add_command(label="Create Unit", command=self._create_unit)
            context_menu.add_command(label="Create Area", command=self._create_area)
        elif item_type == 'unit_node':
            context_menu.add_command(label="Create Area", command=self._create_area)
            context_menu.add_command(label="Create Equipment",
                                     command=lambda: self._create_from_repository('equipment_node'))
            context_menu.add_command(label="Create Assets", 
                                     command=lambda: self._create_from_repository('asset_node'))
        elif item_type == 'area_node':
            context_menu.add_command(label="Create Equipment",
                                     command=lambda: self._create_from_repository('equipment_node'))
            context_menu.add_command(label="Create Assets", 
                                     command=lambda: self._create_from_repository('asset_node'))
        elif item_type == 'equipment_node':
            context_menu.add_command(label="Create Sub-Equipment",
                                     command=lambda: self._create_from_repository('sub_equipment_node'))
            context_menu.add_command(label="Create Assets", 
                                     command=lambda: self._create_from_repository('asset_node'))

        # Display the menu at the cursor's location
        if context_menu.index("end") is not None:
            context_menu.post(event.x_root, event.y_root)

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

    def _create_asset(self):
        name = simpledialog.askstring("Create Asset", "Enter Asset Name:")
        if name:
            self.tree.insert(self.current_item_id, "end", text=name, open=False, tags=('asset_node',))

    def _categorize_nodes(self):
        """Categorize fetched nodes into Equipment, Sub-Equipment, and Asset based on GraphDB queries."""
        if not self.fetched_nodes['all']:
            messagebox.showwarning("No Nodes", "Please fetch nodes from the repository first.")
            return
            
        self._update_status("Categorizing nodes by type...")
        
        def task():
            try:
                prefix = self.entry_prefix.get().strip()
                
                # Reset categorized nodes
                self.fetched_nodes['equipment'] = []
                self.fetched_nodes['sub_equipment'] = []
                self.fetched_nodes['asset'] = []
                
                # Query for Equipment nodes
                equipment_query = f"""{prefix}
                SELECT DISTINCT ?equipment WHERE {{
                    ?equipment a ex:Equipment .
                }} ORDER BY ?equipment"""
                
                # Query for Sub-Equipment nodes  
                sub_equipment_query = f"""{prefix}
                SELECT DISTINCT ?subequipment WHERE {{
                    ?subequipment a ex:SubEquipment .
                }} ORDER BY ?subequipment"""
                
                # Query for Asset nodes
                asset_query = f"""{prefix}
                SELECT DISTINCT ?asset WHERE {{
                    ?asset a ex:Asset .
                }} ORDER BY ?asset"""
                
                # Execute queries and categorize
                equipment_results = self._run_sparql_query(equipment_query)
                if equipment_results:
                    self.fetched_nodes['equipment'] = [
                        res['equipment']['value'].split('#')[-1].split('/')[-1] 
                        for res in equipment_results
                    ]
                
                sub_equipment_results = self._run_sparql_query(sub_equipment_query)
                if sub_equipment_results:
                    self.fetched_nodes['sub_equipment'] = [
                        res['subequipment']['value'].split('#')[-1].split('/')[-1] 
                        for res in sub_equipment_results
                    ]
                
                asset_results = self._run_sparql_query(asset_query)
                if asset_results:
                    self.fetched_nodes['asset'] = [
                        res['asset']['value'].split('#')[-1].split('/')[-1] 
                        for res in asset_results
                    ]
                
                def update_ui():
                    total_categorized = (len(self.fetched_nodes['equipment']) + 
                                       len(self.fetched_nodes['sub_equipment']) + 
                                       len(self.fetched_nodes['asset']))
                    
                    self._update_status(
                        f"Categorized {total_categorized} nodes: "
                        f"{len(self.fetched_nodes['equipment'])} Equipment, "
                        f"{len(self.fetched_nodes['sub_equipment'])} Sub-Equipment, "
                        f"{len(self.fetched_nodes['asset'])} Assets", 
                        5000
                    )
                
                self.after(0, update_ui)
                
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Categorization Error", 
                                                          f"Failed to categorize nodes: {e}"))
                self.after(0, lambda: self._update_status("Error categorizing nodes.", 4000))
        
        threading.Thread(target=task, daemon=True).start()

    def _create_from_repository(self, node_tag):
        """Creates nodes from repository with type-specific filtering."""
        if not self.fetched_nodes['all']:
            messagebox.showwarning("No Nodes Available", 
                                 "Please fetch nodes from the GraphDB repository first using the 'GraphDB → Excel' tab.")
            return
        
        # Determine node type and get appropriate list
        node_type_map = {
            'equipment_node': 'equipment',
            'sub_equipment_node': 'sub_equipment', 
            'asset_node': 'asset'
        }
        
        node_type = node_type_map.get(node_tag, 'all')
        
        # Get the appropriate node list
        if node_type in self.fetched_nodes and self.fetched_nodes[node_type]:
            repo_nodes = self.fetched_nodes[node_type]
        else:
            # Fallback to all nodes if specific type not categorized
            repo_nodes = self.fetched_nodes['all']
            
        if not repo_nodes:
            messagebox.showwarning("No Nodes Available", 
                                 f"No {node_type.replace('_', ' ')} nodes found in the repository.\n"
                                 "Try running 'Categorize Nodes' first, or check your GraphDB schema.")
            return

        # Create dialog title based on node type
        dialog_titles = {
            'equipment_node': "Select Equipment from GraphDB Repository",
            'sub_equipment_node': "Select Sub-Equipment from GraphDB Repository", 
            'asset_node': "Select Assets from GraphDB Repository"
        }
        
        dialog_title = dialog_titles.get(node_tag, "Select Items from GraphDB Repository")
        
        # Show selection dialog with filtering
        dialog = SelectionDialog(self, title=dialog_title, item_list=repo_nodes, node_type=node_type)

        # Wait for user selection
        self.wait_window(dialog)

        if dialog.result:  # Check if the user made a selection
            for item_name in dialog.result:
                self.tree.insert(self.current_item_id, "end", text=item_name, open=True, tags=(node_tag,))
            
            # Update status with selection info
            self._update_status(f"Added {len(dialog.result)} {node_type.replace('_', ' ')} node(s) to hierarchy.", 3000)

    def _build_functionalities_tab(self, parent):
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

        sub_notebook = ttk.Notebook(parent)
        sub_notebook.grid(row=0, column=0, sticky="nsew")

        frame_tags = ttk.Frame(sub_notebook, padding=10)
        frame_excel_files = ttk.Frame(sub_notebook, padding=10)

        sub_notebook.add(frame_tags, text="Tag Management")
        sub_notebook.add(frame_excel_files, text="File Management")

        self._build_tags_subtab(frame_tags)
        self._build_excel_files_subtab(frame_excel_files)

    def _build_excel_files_subtab(self, parent):
        parent.columnconfigure((0, 1), weight=1)
        parent.rowconfigure(2, weight=1)

        ttk.Label(parent, text="Manage all imported Excel files in the application library.",
                  style='Info.TLabel').grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        ttk.Label(parent, text="Assigned Datasheets", style='Header.TLabel').grid(row=1, column=0, sticky="w", padx=5)
        self.assigned_files_listbox = tk.Listbox(parent, height=10)
        self.assigned_files_listbox.grid(row=2, column=0, sticky="nsew", padx=(0, 5), pady=5)

        ttk.Label(parent, text="Unassigned Datasheets (Templates)", style='Header.TLabel').grid(row=1, column=1,
                                                                                                sticky="w", padx=5)
        self.unassigned_files_listbox = tk.Listbox(parent, height=10)
        self.unassigned_files_listbox.grid(row=2, column=1, sticky="nsew", padx=(5, 0), pady=5)

        button_bar = ttk.Frame(parent)
        button_bar.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        ttk.Button(button_bar, text="Import Datasheet Template...", command=self._import_datasheet_template).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_bar, text="Remove Selected File from Library", command=self._remove_file).pack(side=tk.LEFT,
                                                                                                         padx=5)

    def _build_tags_subtab(self, parent):
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(1, weight=1)

        create_frame = ttk.LabelFrame(parent, text="Create or Update Tag", padding=10)
        create_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        create_frame.columnconfigure(1, weight=1)

        ttk.Label(create_frame, text="Tag Name:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.entry_tag = ttk.Entry(create_frame)
        self.entry_tag.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(create_frame, text="Associate Node(s):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.node_combobox = ttk.Combobox(create_frame, state="readonly")
        self.node_combobox.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(create_frame, text="Associate Datasheet(s) [Creates a Copy]:").grid(row=2, column=0, sticky="w",
                                                                                      padx=5, pady=2)
        self.datasheet_combobox = ttk.Combobox(create_frame, state="readonly")
        self.datasheet_combobox.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        ttk.Button(create_frame, text="Create/Update Tag", command=self._add_tag).grid(row=3, column=1, sticky="e",
                                                                                       padx=5, pady=10)

        view_frame = ttk.LabelFrame(parent, text="View Tag Associations", padding=10)
        view_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        view_frame.columnconfigure(1, weight=1)
        view_frame.rowconfigure(1, weight=1)

        ttk.Label(view_frame, text="Existing Tags (Double-click to view)").grid(row=0, column=0, columnspan=2,
                                                                                sticky="w", padx=5)
        self.tag_listbox = tk.Listbox(view_frame, height=5, exportselection=False)
        self.tag_listbox.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        self.tag_listbox.bind("<Double-1>", lambda event: self._show_tag_connections(event))

        ttk.Label(view_frame, text="Associated Nodes").grid(row=2, column=0, sticky="w", padx=5, pady=(10, 0))
        self.nodes_display = tk.Listbox(view_frame, height=5, exportselection=False)
        self.nodes_display.grid(row=3, column=0, sticky="nsew", padx=5, pady=5)
        view_frame.rowconfigure(3, weight=1)

        ttk.Label(view_frame, text="Associated Datasheets").grid(row=2, column=1, sticky="w", padx=5, pady=(10, 0))
        self.datasheets_display = tk.Listbox(view_frame, height=5, exportselection=False)
        self.datasheets_display.grid(row=3, column=1, sticky="nsew", padx=(5, 5), pady=5)
        self.datasheets_display.bind("<Double-1>", lambda event: self._load_datasheet_from_functionalities_tab(event))

    # --- CORE LOGIC METHODS ---

    def _resize_excel_window(self, event=None):
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
            messagebox.showerror("Embedding Error", f"Could not embed the Excel window.\n\n{e}")
            self._cleanup_xlwings()

    def _load_excel_file(self, file_path):
        try:
            abs_path = os.path.abspath(file_path)
            if not os.path.exists(abs_path):
                messagebox.showerror("File Not Found", f"The file could not be found at:\n{abs_path}")
                return

            self._cancel_drag_mode()

            if self.xl_app is None:
                self.xl_app = xw.App(visible=True, add_book=False)

            self.current_xl_db = self.xl_app.books.open(abs_path)
            self.current_xl_db.activate()

            self.after(200, self._embed_excel_window)

            sheet_names = [sheet.name for sheet in self.current_xl_db.sheets]

            self.sheet_selector_combobox['values'] = sheet_names
            if sheet_names:
                self.sheet_selector_combobox.set(sheet_names[0])

            self.current_excel_path = abs_path
            self._update_status(f"Embedding '{os.path.basename(file_path)}'...", 4000)

        except Exception as e:
            messagebox.showerror("xlwings Load Error",
                                 f"Failed to open or embed the Excel file.\n\nError: {e}")
            self._cleanup_xlwings()
            self.current_xl_db = None
            self.sheet_selector_combobox['values'] = []
            self.sheet_selector_combobox.set('')

    def _on_sheet_selected(self, event=None):
        selected_sheet = self.sheet_selector_combobox.get()
        if selected_sheet:
            self._display_sheet(selected_sheet)

    def _display_sheet(self, sheet_name):
        if not self.current_xl_db: return
        try:
            self.current_xl_db.sheets[sheet_name].activate()
            self._update_status(f"Activated sheet '{sheet_name}'.", 3000)
        except Exception as e:
            messagebox.showerror("Sheet Activation Error", f"Could not activate sheet '{sheet_name}'.\n\nError: {e}")

    def _initiate_drag_copy(self, value_to_drag, source_type):
        if value_to_drag is None:
            messagebox.showinfo("No Value", f"There is no {source_type} to drag.")
            return

        if self.is_dragging and self.drag_source_type == source_type:
            self._cancel_drag_mode()
        else:
            self.is_dragging = True
            self.drag_data = value_to_drag
            self.drag_source_type = source_type
            self.config(cursor="hand2")
            self._update_status(f"DRAGGING {source_type.upper()}: Click cell in Excel to drop.", 0)
            self._poll_for_drop()

    def _cancel_drag_mode(self):
        self.is_dragging = False
        self.drag_data = None
        self.drag_source_type = None
        self.config(cursor="")
        self.live_value_button.config(text=f"Copy Value: {self.live_value_button.value or '(none)'}")
        self.live_unit_button.config(text=f"Copy Unit: {self.live_unit_button.value or '(none)'}")
        self._update_status("Ready", 0)

    def _poll_for_drop(self):
        if not self.is_dragging: return

        if win32api.GetKeyState(0x01) < 0:
            screen_x, screen_y = win32gui.GetCursorPos()
            frame_x = self.excel_frame.winfo_rootx()
            frame_y = self.excel_frame.winfo_rooty()
            frame_w = self.excel_frame.winfo_width()
            frame_h = self.excel_frame.winfo_height()

            if frame_x <= screen_x <= frame_x + frame_w and frame_y <= screen_y <= frame_y + frame_h:
                try:
                    xl_range = self.xl_app.api.ActiveWindow.RangeFromPoint(screen_x, screen_y)
                    if xl_range:
                        sheet = self.current_xl_db.sheets.active
                        write_cell_com = xl_range.MergeArea.Cells(1, 1)
                        write_cell = sheet.range(write_cell_com.Address)

                        write_cell.value = self.drag_data
                        self.current_xl_db.save()

                        self._update_status(f"Pasted '{self.drag_data}' to {write_cell.address.replace('$', '')}.",
                                            4000)
                        self._cancel_drag_mode()
                        return
                except Exception as e:
                    self._update_status(f"Drop Error: {e}", 3000)
                    self._cancel_drag_mode()
                    return

        self.after(100, self._poll_for_drop)

    def _import_datasheet_template(self):
        path = filedialog.askopenfilename(
            title="Select a Datasheet Template to Import",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
        )
        if not path: return

        try:
            filename = os.path.basename(path)
            internal_path = os.path.join(self.storage_dir, filename)

            if os.path.exists(internal_path):
                if not messagebox.askyesno("Overwrite?",
                                           f"The file '{filename}' already exists in the library. Overwrite it?"):
                    return

            shutil.copy(path, internal_path)
            self._update_file_lists()
            messagebox.showinfo("Import Successful", f"Template '{filename}' imported successfully to the library.")
            self._update_status("Datasheet library updated.", 4000)
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import template:\n{e}")

    def _update_data_model(self, event=None):
        selected_indices = self.node_listbox_display.curselection()
        if not selected_indices: return

        nodes = [self.node_listbox_display.get(i) for i in selected_indices]
        repo_url, prefix = self.entry_repo.get().strip(), self.entry_prefix.get().strip()

        if not all([repo_url, prefix, nodes]):
            messagebox.showwarning("Missing Data", "Repository URL, Prefix, and a selected node are required.")
            return

        self._clear_graph()
        self._update_status("Fetching data model...")

        def get_local_name(uri):
            if not isinstance(uri, str): return uri
            return uri.split('#')[-1].split('/')[-1]

        def task():
            node_conditions = " || ".join(
                [f"sameTerm(?subject, ex:{node}) || sameTerm(?object, ex:{node})" for node in nodes])
            query = f"{prefix}\nSELECT ?subject ?predicate ?object WHERE {{ ?subject ?predicate ?object . FILTER({node_conditions}) }}"
            results = self._run_sparql_query(query)
            if results is None: return

            self.graph.clear()
            for res in results:
                subject = get_local_name(res.get("subject", {}).get("value", ""))
                predicate = get_local_name(res.get("predicate", {}).get("value", ""))
                obj = get_local_name(res.get("object", {}).get("value", ""))
                if all([subject, predicate, obj]):
                    self.graph.add_edge(subject, obj, label=predicate)

            self.after(0, self._draw_graph)

        threading.Thread(target=task, daemon=True).start()

    def _draw_graph(self):
        for widget in self.graph_frame.winfo_children():
            widget.destroy()

        if not self.graph.nodes():
            self._update_status("No data found for selected nodes.", 4000)
            return

        try:
            fig, ax = plt.subplots(figsize=(10, 8))
            pos = nx.spring_layout(self.graph, k=0.7, iterations=50)
            nx.draw(self.graph, pos, ax=ax, with_labels=True, node_color='#a0cbe2', node_size=2500, font_size=10,
                    font_weight='bold', width=1.5, edge_color='gray', arrows=True)
            edge_labels = nx.get_edge_attributes(self.graph, 'label')
            nx.draw_networkx_edge_labels(self.graph, pos, edge_labels=edge_labels, font_color='firebrick', font_size=9)
            fig.tight_layout()

            self.canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
            self.canvas.draw()
            self.canvas.get_tk_widget().pack(expand=True, fill="both")
            self._update_status("Data model loaded.", 4000)
            plt.close(fig)
        except Exception as e:
            messagebox.showerror("Graphing Error", f"An error occurred while drawing the graph: {e}")

    def _clear_graph(self):
        for widget in self.graph_frame.winfo_children():
            widget.destroy()
        self.graph.clear()
        if hasattr(self, 'canvas'):
            del self.canvas
        self._update_status("Graph cleared.")

    def _run_sparql_query(self, query):
        repo_url = self.entry_repo.get().strip()
        if not repo_url:
            self.after(0, lambda: messagebox.showerror("Error", "Repository URL is not set."))
            return None
        try:
            sparql = SPARQLWrapper(repo_url)
            sparql.setQuery(query)
            sparql.setReturnFormat(JSON)
            return sparql.query().convert()["results"]["bindings"]
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("SPARQL Error", f"Failed to execute query:\n{e}"))
            self.after(0, lambda: self._update_status(f"SPARQL Error: {e}", 5000))
            return None

    def _fetch_nodes(self):
        prefix_str = self.entry_prefix.get().strip()
        if not prefix_str or '<' not in prefix_str or '>' not in prefix_str:
            messagebox.showerror("Invalid Prefix",
                                 "Please provide a valid SPARQL Prefix (e.g., PREFIX ex: <http://example.org#>)")
            return
        self._update_status("Fetching all nodes from repository...")

        def task():
            try:
                uri_base = prefix_str.split('<')[1].split('>')[0]
                query = f'SELECT DISTINCT ?resource WHERE {{ {{ ?resource ?p ?o . }} UNION {{ ?s ?p ?resource . }} FILTER(ISIRI(?resource) && STRSTARTS(STR(?resource), "{uri_base}")) }} ORDER BY ?resource'
                results = self._run_sparql_query(query)
                if results is None:
                    self.after(0, lambda: self._update_status("Failed to fetch nodes.", 4000))
                    return
                nodes = sorted(list(set(res["resource"]["value"].split('#')[-1].split('/')[-1] for res in results)))
                self.after(0, lambda: self._update_node_lists(nodes))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", f"Failed to parse prefix or fetch nodes: {e}"))
                self.after(0, lambda: self._update_status("Error fetching nodes.", 4000))

        threading.Thread(target=task, daemon=True).start()

    def _update_node_lists(self, nodes):
        self.node_listbox.delete(0, tk.END)
        self.node_listbox_display.delete(0, tk.END)
        
        # Store all fetched nodes
        self.fetched_nodes['all'] = nodes
        
        for node in nodes:
            self.node_listbox.insert(tk.END, node)
            self.node_listbox_display.insert(tk.END, node)
        self.node_combobox['values'] = nodes
        self._update_status(f"{len(nodes)} nodes fetched. Use 'Categorize Nodes' to organize by type.", 4000)

    def _update_file_lists(self):
        try:
            all_files_in_dir = sorted(
                [f for f in os.listdir(self.storage_dir) if f.endswith(('.xlsx', '.xls', '.xlsm'))])

            assigned_files = set()
            for tag_data in self.tag_associations.values():
                assigned_files.update(tag_data.get('datasheets', []))

            self.assigned_files_listbox.delete(0, tk.END)
            self.unassigned_files_listbox.delete(0, tk.END)

            for f in all_files_in_dir:
                if f in assigned_files:
                    self.assigned_files_listbox.insert(tk.END, f)
                else:
                    self.unassigned_files_listbox.insert(tk.END, f)

            self.datasheet_combobox['values'] = all_files_in_dir
        except Exception as e:
            print(f"Error updating file list: {e}")

    def _update_tag_lists(self):
        try:
            tags = sorted(self.tag_associations.keys())
            self.tag_listbox.delete(0, tk.END)
            for tag in tags:
                self.tag_listbox.insert(tk.END, tag)
            self.tag_selector_combobox['values'] = tags
        except Exception as e:
            print(f"Error updating tag lists: {e}")

    def _save_excel_as(self):
        if not self.current_excel_path or not self.current_xl_db:
            messagebox.showwarning("No File", "A datasheet must be open in Excel to use 'Save As'.")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx *.xlsm")],
                                                 title="Save Datasheet As")
        if not save_path: return
        try:
            abs_save_path = os.path.abspath(save_path)
            self.current_xl_db.save(abs_save_path)

            filename = os.path.basename(abs_save_path)
            internal_path = os.path.join(self.storage_dir, filename)

            if os.path.normpath(abs_save_path).lower() != os.path.normpath(internal_path).lower():
                shutil.copyfile(abs_save_path, internal_path)

            self._update_file_lists()
            messagebox.showinfo("Success",
                                f"File exported as '{filename}' and a copy was added/updated in the library.")
        except Exception as e:
            messagebox.showerror("Save As Error", f"Failed to save the file: {e}")

    def _load_file_from_list(self, event=None):
        if not event.widget.curselection(): return
        file_name = event.widget.get(event.widget.curselection())

        if self.current_xl_db and self.current_xl_db.name == file_name:
            self.current_xl_db.activate()
        else:
            self._load_excel_file(os.path.join(self.storage_dir, file_name))

    def _on_tag_selected_in_editor_tab(self, event=None):
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
            for datasheet in self.tag_associations[tag_name].get('datasheets', []):
                self.datasheet_listbox_for_tag.insert(tk.END, datasheet)
        self._display_tag_info_in_editor_view(tag_name)

    def _on_property_double_click(self, event=None):
        try:
            index = self.properties_text_embed.index(f"@{event.x},{event.y} linestart")
            line_end = self.properties_text_embed.index(f"{index} lineend")
            line_text = self.properties_text_embed.get(index, line_end)

            match = re.search(r"^\s*[-*]?\s*([\w_]+)", line_text)
            if not match: return
            prop = match.group(1)

            if not prop: return

            self.selected_property_display.config(text=prop)

            if not self.active_node:
                self.live_value_button.value = None
                self.live_unit_button.value = None
                self.live_value_button.config(text="Copy Value: (No Active Node)")
                self.live_unit_button.config(text="Copy Unit: (No Active Node)")
                messagebox.showinfo("Info", "An active node must be set to see a live value preview.")
                return

            node = self.active_node
            self._update_status(f"Fetching preview for {prop}...")

            def task():
                prefix = self.entry_prefix.get().strip()
                query_bnode = f"{prefix} SELECT ?value ?unit WHERE {{ ex:{node} ex:{prop} ?bnode . ?bnode ex:hasValue ?value . OPTIONAL {{ ?bnode ex:hasUnit ?unit . }} }}"
                res_bnode = self._run_sparql_query(query_bnode)
                val, uni = (
                res_bnode[0]["value"]["value"], res_bnode[0].get("unit", {}).get("value")) if res_bnode else (
                None, None)

                def update_ui():
                    self.live_value_button.value = val
                    self.live_unit_button.value = uni
                    self.live_value_button.config(text=f"Copy Value: {val or '(none)'}")
                    self.live_unit_button.config(text=f"Copy Unit: {uni or '(none)'}")
                    self._update_status("Preview loaded.", 4000)

                self.after(0, update_ui)

            threading.Thread(target=task, daemon=True).start()

        except (tk.TclError, IndexError):
            pass

    def _display_tag_info_in_editor_view(self, tag_name):
        self.tag_text_embed.config(state='normal')
        self.properties_text_embed.config(state='normal')
        self.tag_text_embed.delete(1.0, tk.END)
        self.properties_text_embed.delete(1.0, tk.END)

        if tag_name not in self.tag_associations:
            self.tag_text_embed.insert(tk.END, "Tag not found.")
            self.tag_text_embed.config(state='disabled')
            self.properties_text_embed.config(state='disabled')
            return

        self.tag_text_embed.insert(tk.END, f"Tag: {tag_name}")
        nodes = self.tag_associations[tag_name].get('nodes', [])

        if not nodes:
            self.properties_text_embed.insert(tk.END, "Tag has no associated nodes.")
        else:
            node_list_str = '\n- '.join(nodes)
            self.properties_text_embed.insert(tk.END,
                                              f"Associated Nodes:\n- {node_list_str}\n\nProperties (from all nodes):\n")

        self.tag_text_embed.config(state='disabled')
        self.properties_text_embed.config(state='disabled')

        if not nodes: return

        self._update_status(f"Fetching properties for tag '{tag_name}'...")

        def task():
            all_properties = set()
            prefix = self.entry_prefix.get().strip()
            for node in nodes:
                query = f"""{prefix}
                SELECT DISTINCT ?p WHERE {{ 
                    ex:{node} ?p ?o .
                    FILTER (isLiteral(?o) || isBlank(?o))
                }}"""
                results = self._run_sparql_query(query)
                if results:
                    all_properties.update(p.split('#')[-1] for p in (res['p']['value'] for res in results) if
                                          p.split('#')[-1] not in ["hasValue", "hasUnit", "a", "type"])
            prop_text = "\n".join(f"- {p}" for p in sorted(list(all_properties))) or "(No direct properties found)"

            def update_text():
                self.properties_text_embed.config(state='normal')
                self.properties_text_embed.insert(tk.END, prop_text)
                self.properties_text_embed.config(state='disabled')
                self._update_status(f"Info loaded for tag '{tag_name}'.", 4000)

            self.after(0, update_text)

        threading.Thread(target=task, daemon=True).start()

    def _remove_file(self):
        file_name = None
        if self.assigned_files_listbox.curselection():
            file_name = self.assigned_files_listbox.get(self.assigned_files_listbox.curselection())
        elif self.unassigned_files_listbox.curselection():
            file_name = self.unassigned_files_listbox.get(self.unassigned_files_listbox.curselection())

        if not file_name:
            messagebox.showwarning("No Selection", "Please select a file to remove from the library.")
            return

        if self.current_xl_db and self.current_xl_db.name == file_name:
            self.current_xl_db.close()
            self.current_xl_db = None
            self.sheet_selector_combobox['values'] = []
            self.sheet_selector_combobox.set('')

        if messagebox.askyesno("Confirm Removal",
                               f"Are you sure you want to permanently delete '{file_name}' from the library? This will also un-tag it from any associations."):
            try:
                os.remove(os.path.join(self.storage_dir, file_name))
                for tag in self.tag_associations:
                    if file_name in self.tag_associations[tag]['datasheets']:
                        self.tag_associations[tag]['datasheets'].remove(file_name)
                self._update_file_lists()
                self._on_tag_selected_in_editor_tab()
                messagebox.showinfo("Success", f"'{file_name}' has been removed from the library.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to remove file: {e}")

    def _add_tag(self):
        tag = self.entry_tag.get().strip()
        node = self.node_combobox.get().strip()
        template_datasheet = self.datasheet_combobox.get().strip()

        if not tag:
            messagebox.showwarning("Input Required", "Tag name cannot be empty.")
            return

        datasheet_to_assign = None

        # --- NEW TEMPLATE CLONING LOGIC ---
        if template_datasheet:
            if not node:
                messagebox.showwarning("Node Required",
                                       "Please select a node to associate with the new datasheet copy.")
                return

            # Propose a new name based on the node
            name, ext = os.path.splitext(template_datasheet)
            default_new_name = f"{node}_datasheet{ext}"

            new_filename = simpledialog.askstring(
                "Name New Datasheet",
                "Enter a filename for the new datasheet copy:",
                initialvalue=default_new_name
            )

            if not new_filename:  # User cancelled
                return

            # Create the new file
            source_path = os.path.join(self.storage_dir, template_datasheet)
            dest_path = os.path.join(self.storage_dir, new_filename)

            if os.path.exists(dest_path):
                if not messagebox.askyesno("Overwrite?", f"The file '{new_filename}' already exists. Overwrite it?"):
                    return

            try:
                shutil.copy2(source_path, dest_path)
                datasheet_to_assign = new_filename
                messagebox.showinfo("Success", f"Created and assigned '{new_filename}'.")
            except Exception as e:
                messagebox.showerror("File Copy Error", f"Could not create datasheet copy:\n{e}")
                return
        # --- END NEW LOGIC ---

        # Now, perform the association
        if tag not in self.tag_associations:
            self.tag_associations[tag] = {'nodes': [], 'datasheets': []}

        if node and node not in self.tag_associations[tag]['nodes']:
            self.tag_associations[tag]['nodes'].append(node)

        if datasheet_to_assign and datasheet_to_assign not in self.tag_associations[tag]['datasheets']:
            self.tag_associations[tag]['datasheets'].append(datasheet_to_assign)

        self._update_tag_lists()
        self._update_file_lists()
        self._show_tag_connections(tag_name=tag)

    def _show_tag_connections(self, event=None, tag_name=None):
        tag = tag_name or (
            self.tag_listbox.get(self.tag_listbox.curselection()) if self.tag_listbox.curselection() else None)
        if not tag: return
        self.nodes_display.delete(0, tk.END)
        self.datasheets_display.delete(0, tk.END)
        if tag in self.tag_associations:
            for node in self.tag_associations[tag].get('nodes', []):
                self.nodes_display.insert(tk.END, node)
            for datasheet in self.tag_associations[tag].get('datasheets', []):
                self.datasheets_display.insert(tk.END, datasheet)

    def _load_datasheet_from_functionalities_tab(self, event=None):
        if not event.widget.curselection(): return
        file_name = event.widget.get(event.widget.curselection())
        self.main_notebook.select(2)
        self._load_file_from_list(event)

    def _on_node_manual_select(self, event=None):
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


if __name__ == "__main__":
    app = LeanDigitalTwin()

    if app.initialization_ok:
        app.mainloop()