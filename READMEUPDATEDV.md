# EXCEL-CLASSIFIER-APP
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import re
import numpy as np
import pickle
import threading
import queue
import csv  # For CSV sniffing/robust reading
import io  # For reading strings like a file


class ExcelClassifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Classifier Viewer")

        try:
            # self.root.iconphoto(True, tk.PhotoImage(file='your_icon.png'))
            pass
        except tk.TclError:
            print("Icon not found. Using default icon.")

        self.root.geometry("1200x700")

        self.style = ttk.Style()

        # --- THEME LOGIC ---
        self.current_theme = "light"
        self.setup_themes()
        self.style.theme_use(self.current_theme)
        # --- END THEME ---

        # --- Menu Bar Setup ---
        self.current_job_file_path = None
        self.setup_menu_bar()
        # --- END Menu ---

        # --- Sidebar (Right Frame) Setup ---
        self.right_frame = ttk.Frame(self.root, width=200)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10)
        self.right_frame.pack_propagate(False)

        # --- Sidebar Toggle Frame & Button ---
        self.sidebar_toggle_frame = ttk.Frame(self.root, width=20)
        self.sidebar_toggle_frame.pack(side=tk.RIGHT, fill=tk.Y, pady=10)

        self.sidebar_toggle_button = ttk.Button(
            self.sidebar_toggle_frame,
            text=">",
            command=self.hide_sidebar,
            style="Toggle.TButton",
            width=2
        )
        self.sidebar_toggle_button.pack(side=tk.LEFT, fill=tk.Y)
        # --- END Toggle ---

        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))  # Reduced bottom padding

        self.board_control_frame = ttk.Frame(self.main_frame)
        self.board_control_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        self.board_label = ttk.Label(self.board_control_frame, text="Select Board:")
        self.board_label.pack(side=tk.LEFT, padx=(0, 5))

        self.board_selector = ttk.Combobox(self.board_control_frame, state="disabled", width=20)
        self.board_selector.pack(side=tk.LEFT, padx=(0, 10))
        self.board_selector.bind("<<ComboboxSelected>>", self.on_board_select)

        self.mpin_group_frame = ttk.Frame(self.board_control_frame)
        self.mpin_group_frame.pack(side=tk.LEFT, padx=10)

        self.mpin_group_label = ttk.Label(self.mpin_group_frame, text="Multi-Pin Group:")
        self.mpin_group_label.pack(side=tk.LEFT, padx=(0, 5))

        self.mpin_group_selector = ttk.Combobox(self.board_control_frame, state="disabled", width=20)
        self.mpin_group_selector.pack(side=tk.LEFT, padx=(0, 10))
        self.mpin_group_selector.bind("<<ComboboxSelected>>", self.on_mpin_group_select)

        self.pin_filter_frame = ttk.Frame(self.board_control_frame)
        self.pin_filter_frame.pack(side=tk.LEFT, padx=10)

        self.pin_filter_label = ttk.Label(self.pin_filter_frame, text="Pin Filter:")
        self.pin_filter_label.pack(side=tk.LEFT, padx=(0, 5))

        self.pin_filter_all_btn = ttk.Button(self.pin_filter_frame, text="All Parts", style="Filter.TButton",
                                             command=lambda: self.set_pin_filter("all"), state=tk.NORMAL)
        self.pin_filter_all_btn.pack(side=tk.LEFT, padx=2)

        self.pin_filter_2pin_btn = ttk.Button(self.pin_filter_frame, text="2-Pin", style="Filter.TButton",
                                              command=lambda: self.set_pin_filter("2pin"), state=tk.NORMAL)
        self.pin_filter_2pin_btn.pack(side=tk.LEFT, padx=2)

        self.pin_filter_mpin_btn = ttk.Button(self.pin_filter_frame, text="Multi-Pin (Rest)", style="Filter.TButton",
                                              command=lambda: self.set_pin_filter("mpin"), state=tk.NORMAL)
        self.pin_filter_mpin_btn.pack(side=tk.LEFT, padx=2)

        self.group_filter_frame = ttk.Frame(self.board_control_frame)
        self.group_filter_frame.pack(side=tk.LEFT, padx=10)

        self.group_filter_label = ttk.Label(self.group_filter_frame, text="Component Type:")
        self.group_filter_label.pack(side=tk.LEFT, padx=(0, 5))

        self.group_filter_all_btn = ttk.Button(self.group_filter_frame, text="All", style="Filter.TButton",
                                               command=lambda: self.set_group_filter("all"))
        self.group_filter_all_btn.pack(side=tk.LEFT, padx=2)

        self.group_filter_single_btn = ttk.Button(self.group_filter_frame, text="Single", style="Filter.TButton",
                                                  command=lambda: self.set_group_filter("single"))
        self.group_filter_single_btn.pack(side=tk.LEFT, padx=2)

        self.group_filter_grouped_btn = ttk.Button(self.group_filter_frame, text="Grouped (/)", style="Filter.TButton",
                                                   command=lambda: self.set_group_filter("grouped"))
        self.group_filter_grouped_btn.pack(side=tk.LEFT, padx=2)

        self.style_filter = tk.StringVar()
        self.part_filter = tk.StringVar()
        self.remark_filter = tk.StringVar()

        self.style_trace_id = self.style_filter.trace_add("write", self.on_text_filter_change)
        self.part_trace_id = self.part_filter.trace_add("write", self.on_text_filter_change)
        self.remark_trace_id = self.remark_filter.trace_add("write", self.on_text_filter_change)

        self.setup_menubuttons()

        self.workspace_frame = ttk.Frame(self.main_frame)
        self.workspace_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(self.workspace_frame, show="tree headings", selectmode="extended")

        vsb = ttk.Scrollbar(self.workspace_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.workspace_frame, orient="horizontal", command=self.tree.xview)

        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree.pack(fill=tk.BOTH, expand=True)

        # --- Status Bar Frame ---
        self.status_bar_frame = ttk.Frame(self.root, height=20, relief=tk.SUNKEN)
        self.status_bar_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label = ttk.Label(self.status_bar_frame, text="Ready.", anchor=tk.W)
        self.status_label.pack(fill=tk.X, padx=10, pady=2)
        # --- END Status Bar ---

        self.full_df = None
        self.original_df = None
        self.current_filtered_df = None
        self.current_highlight = None
        self.current_view = "edited"
        self.current_board_view = "All Boards"
        self.current_pin_filter = "all"
        self.current_group_filter = "all"
        self.current_mpin_group_filter = "All"
        self.board_list = []

        # New: Search state for Ctrl+F
        self.search_term_var = tk.StringVar()
        self.current_search_query = ""
        self.search_window = None  # To track the popup

        # The core columns expected from a TRD-style file
        self.trd_columns = [
            "Board", "Step", "ID", "Style", "Part", "Remark", "ACT", "STD", "Tst",
            "HL", "LL", "A", "B", "EA", "EB", "LC", "Result"
        ]

        # All columns including internal and computed ones (es_remarks added here)
        self.all_columns = self.trd_columns + [
            "Classification", "Override", "Is_Grouped", "Parallel", "Coverage",
            "Pin_Type", "MPin_Group", "es_remarks"
        ]

        # --- Config data instances ---
        # RefDes prefixes are used to determine the category in _get_component_category
        self.summary_config_data = [
            ("Resistor", "R"),
            ("Capacitor", "C"),
            # Transistors (Q), Diodes (D), Zener (Z), LED, TR
            ("Diode/Transistor", "D, Q, Z, LED, TR"),
            # Inductor (L), Ferrite Bead (FB)
            ("Inductor/FB", "L, FB"),
            # Relays (K, RL)
            ("Relay", "K, RL"),
            # Integrated Circuit (U), IC
            ("IC", "U, IC"),
            # Connector (J), Pin Header (P), Connector (CN)
            ("Connector", "J, P, CN"),
            ("Others", "")
        ]

        # Config for the new es_remarks field: (Special Name, Full Form)
        # REMOVED the third 'Remarks' column as requested.
        self.remarks_config_data = [
            ("nm", "not mounted")
        ]
        # --- END Config data ---

        self.reset_tag_colors()
        self.update_pin_filter_buttons()
        self.update_group_filter_buttons()

        # --- Bindings ---
        self.root.bind("<Control-f>", self.open_search_popup)
        self.root.bind("<Control-F>", self.open_search_popup)
        # --- End Bindings ---

        # --- Threading and Queue setup ---
        self.thread_queue = queue.Queue()
        self.loading_window = None
        # --- END Threading ---

    def open_search_popup(self, event=None):
        """Opens the Ctrl+F search pop-up window."""
        if self.full_df is None:
            messagebox.showwarning("No Data", "Please load a data file first to enable searching.")
            return

        if self.search_window and self.search_window.winfo_exists():
            self.search_window.lift()
            return

        self.search_window = tk.Toplevel(self.root)
        self.search_window.title("Find Component")
        self.search_window.transient(self.root)
        self.search_window.grab_set()

        # Center the window near the top of the main window
        self.root.update_idletasks()
        try:
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 200
            y = self.root.winfo_y() + 50
            self.search_window.geometry(f"400x100+{x}+{y}")
        except tk.TclError:
            # Fallback if window dimensions are not yet ready
            self.search_window.geometry("400x100")

        self.search_window.resizable(False, False)

        frame = ttk.Frame(self.search_window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Search Term:").pack(side=tk.LEFT, padx=5)

        # Initialize the entry with the current query if any
        self.search_term_var.set(self.current_search_query)
        search_entry = ttk.Entry(frame, textvariable=self.search_term_var, width=30)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        def apply_search():
            new_query = self.search_term_var.get().strip()
            self.current_search_query = new_query.upper() if new_query else ""
            self.refresh_treeview()
            if self.search_window:
                self.search_window.lift()

        def clear_search():
            self.search_term_var.set("")
            self.current_search_query = ""
            self.refresh_treeview()
            if self.search_window:
                self.search_window.lift()

        search_btn = ttk.Button(frame, text="Search", command=apply_search)
        search_btn.pack(side=tk.LEFT, padx=(0, 5))

        clear_btn = ttk.Button(frame, text="Clear", command=clear_search)
        clear_btn.pack(side=tk.LEFT)

        search_entry.bind("<Return>", lambda e: apply_search())
        search_entry.focus_set()

        self.search_window.protocol("WM_DELETE_WINDOW", self.search_window.destroy)

    # --- Menu Bar Setup ---
    def setup_menu_bar(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # --- File Menu ---
        self.file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=self.file_menu)

        self.file_menu.add_command(label="New", command=self.new_file)
        self.file_menu.add_command(label="Open Job File...", command=self.import_job_file)

        import_menu = tk.Menu(self.file_menu, tearoff=0)
        import_menu.add_command(label="Data File (.xlsx, .csv, .trd)...", command=self.open_excel)
        import_menu.add_command(label="Configuration (.csv)...", command=self.import_config_csv)
        import_menu.add_command(label="Job File (.job)...", command=self.import_job_file)
        self.file_menu.add_cascade(label="Import", menu=import_menu)

        self.file_menu.add_separator()
        self.file_menu.add_command(label="Save", command=self.save_job_file, state=tk.DISABLED)
        self.file_menu.add_command(label="Save As...", command=self.save_job_file_as, state=tk.DISABLED)

        export_menu = tk.Menu(self.file_menu, tearoff=0)
        export_menu.add_command(label="Export Current View (.xlsx)...", command=self.export_current_view)
        export_menu.add_command(label="Export Filtered Board (Multi-Sheet)...",
                                command=self.export_filtered_board_multisheet)
        export_menu.add_command(label="Export All Boards (Separate Files)...", command=self.export_all_boards_separate)

        self.file_menu.add_cascade(label="Export", menu=export_menu, state=tk.DISABLED)

        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.root.quit)

        # --- Edit Menu ---
        self.edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=self.edit_menu)

        classification_menu = tk.Menu(self.edit_menu, tearoff=0)
        classification_menu.add_command(label="Change Selected to P",
                                        command=lambda: self.batch_update_classification("P"))
        classification_menu.add_command(label="Change Selected to F",
                                        command=lambda: self.batch_update_classification("F"))
        classification_menu.add_command(label="Change Selected to NC",
                                        command=lambda: self.batch_update_classification("NC"))
        classification_menu.add_separator()
        classification_menu.add_command(label="Clear Override on Selected",
                                        command=lambda: self.batch_update_classification(pd.NA))

        self.edit_menu.add_cascade(label="Change Classification", menu=classification_menu, state=tk.DISABLED)

        filter_menu = tk.Menu(self.edit_menu, tearoff=0)
        filter_menu.add_command(label="Filter P", command=lambda: self.highlight_rows("P"))
        filter_menu.add_command(label="Filter F", command=lambda: self.highlight_rows("F"))
        filter_menu.add_command(label="Filter NC", command=lambda: self.highlight_rows("NC"))
        filter_menu.add_separator()
        filter_menu.add_command(label="Reset Filters", command=self.reset_highlight)

        self.edit_menu.add_cascade(label="Filter By Classification", menu=filter_menu, state=tk.DISABLED)

        self.edit_menu.add_separator()
        self.edit_menu.add_command(label="Summary Config...", command=self.open_summary_config,
                                   state=tk.NORMAL)  # Always enabled
        self.edit_menu.add_command(label="Remarks Config...", command=self.open_remarks_config,
                                   state=tk.NORMAL)  # Always enabled

        # --- View Menu ---
        self.view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=self.view_menu)

        self.view_menu.add_command(label="Preview Export...", command=self.preview_export, state=tk.DISABLED)
        self.view_menu.add_command(label="Summary View...", command=self.open_summary_view, state=tk.DISABLED)
        self.view_menu.add_separator()

        self.view_mode_var = tk.StringVar(value="edited")
        self.view_menu.add_radiobutton(label="Show Edited File", variable=self.view_mode_var, value="edited",
                                       command=self.show_edited_file, state=tk.DISABLED)
        self.view_menu.add_radiobutton(label="Show Past File (Original)", variable=self.view_mode_var, value="past",
                                       command=self.show_past_file, state=tk.DISABLED)

        self.view_menu.add_separator()
        self.view_menu.add_command(label="Toggle Theme", command=self.toggle_theme)

    # --- END Menu Bar Setup ---

    # --- Theme setup ---
    def setup_themes(self):
        # Define colors
        self.colors = {
            "light": {
                "bg": "#F0F0F0", "fg": "#000000", "bg_widget": "#E0E0E0", "fg_widget": "#000000",
                "bg_entry": "#FFFFFF", "fg_entry": "#000000", "bg_tree": "#FFFFFF", "fg_tree": "#000000",
                "bg_select": "#0078D7", "fg_select": "#FFFFFF", "bg_disabled": "#F0F0F0", "fg_disabled": "#A0A0A0",
                "bg_status": "#F0F0F0", "fg_status": "#000000"
            },
            "dark": {
                "bg": "#2E2E2E", "fg": "#E0E0E0", "bg_widget": "#4A4A4A", "fg_widget": "#E0E0E0",
                "bg_entry": "#3E3E3E", "fg_entry": "#E0E0E0", "bg_tree": "#252525", "fg_tree": "#E0E0E0",
                "bg_select": "#0078D7", "fg_select": "#FFFFFF", "bg_disabled": "#2E2E2E", "fg_disabled": "#5A5A5A",
                "bg_status": "#202020", "fg_status": "#E0E0E0"
            }
        }

        # --- Create Light Theme ---
        c = self.colors["light"]
        self.style.theme_create("light", parent="clam", settings={
            ".": {"configure": {"background": c["bg"], "foreground": c["fg"], "bordercolor": c["bg_widget"]}},
            "TFrame": {"configure": {"background": c["bg"]}},
            "TLabel": {"configure": {"background": c["bg"], "foreground": c["fg"]}},
            "TButton": {"configure": {"background": c["bg_widget"], "foreground": c["fg_widget"], "padding": 6,
                                      "relief": "flat"},
                        "map": {"background": [("active", "#D0D0D0"), ("disabled", c["bg_disabled"])]}},
            "Treeview": {
                "configure": {"background": c["bg_tree"], "foreground": c["fg_tree"], "fieldbackground": c["bg_tree"]},
                "map": {"background": [("selected", c["bg_select"])], "foreground": [("selected", c["fg_select"])]}},
            "Treeview.Heading": {
                "configure": {"background": c["bg_widget"], "foreground": c["fg_widget"], "padding": (5, 5),
                              "relief": "flat"}},
            "TCombobox": {"configure": {"fieldbackground": c["bg_entry"], "foreground": c["fg_entry"],
                                        "selectbackground": c["bg_entry"], "selectforeground": c["fg_entry"],
                                        "arrowcolor": c["fg_widget"]},
                          "map": {"background": [("readonly", c["bg_entry"])],
                                  "fieldbackground": [("readonly", c["bg_entry"])]}},
            "TEntry": {"configure": {"fieldbackground": c["bg_entry"], "foreground": c["fg_entry"],
                                     "insertcolor": c["fg_entry"]}},
            "TMenubutton": {"configure": {"background": c["bg_widget"], "foreground": c["fg_widget"]}},
            "TScrollbar": {
                "configure": {"background": c["bg"], "troughcolor": c["bg_widget"], "arrowcolor": c["fg_widget"]}},
            "Filter.TButton": {"configure": {"padding": 6, "relief": "flat"},
                               "map": {"background": [('selected', c["bg_select"]), ('!selected', c["bg_widget"])],
                                       "foreground": [('selected', c["fg_select"]), ('!selected', c["fg_widget"])]}},
            "Action.TCombobox": {"configure": {"padding": 6, "relief": "flat", "arrowsize": 0},
                                 "map": {"background": [("readonly", c["bg_widget"]), ("active", "#D0D0D0")],
                                         "fieldbackground": [("readonly", c["bg_widget"])],
                                         "foreground": [("readonly", c["fg_widget"])],
                                         "selectbackground": [("readonly", c["bg_widget"])],
                                         "selectforeground": [("readonly", c["fg_widget"])]}},
            "Toggle.TButton": {"configure": {"padding": (2, 5), "relief": "flat"},
                               "map": {"background": [("active", "#D0D0D0"), ("!active", c["bg_widget"])]}},
            # Status Bar Style
            "Status.TLabel": {"configure": {"background": c["bg_status"], "foreground": c["fg_status"]}},
            "Status.TFrame": {"configure": {"background": c["bg_status"]}}

        })

        # --- Create Dark Theme ---
        c = self.colors["dark"]
        self.style.theme_create("dark", parent="clam", settings={
            ".": {"configure": {"background": c["bg"], "foreground": c["fg"], "bordercolor": c["bg_widget"]}},
            "TFrame": {"configure": {"background": c["bg"]}},
            "TLabel": {"configure": {"background": c["bg"], "foreground": c["fg"]}},
            "TButton": {"configure": {"background": c["bg_widget"], "foreground": c["fg_widget"], "padding": 6,
                                      "relief": "flat"},
                        "map": {"background": [("active", "#5A5A5A"), ("disabled", c["bg_disabled"])]}},
            "Treeview": {
                "configure": {"background": c["bg_tree"], "foreground": c["fg_tree"], "fieldbackground": c["bg_tree"]},
                "map": {"background": [("selected", c["bg_select"])], "foreground": [("selected", c["fg_select"])]}},
            "Treeview.Heading": {
                "configure": {"background": c["bg_widget"], "foreground": c["fg_widget"], "padding": (5, 5),
                              "relief": "flat"}},
            "TCombobox": {"configure": {"fieldbackground": c["bg_entry"], "foreground": c["fg_entry"],
                                        "selectbackground": c["bg_entry"], "selectforeground": c["fg_entry"],
                                        "arrowcolor": c["fg_widget"]},
                          "map": {"background": [("readonly", c["bg_entry"])],
                                  "fieldbackground": [("readonly", c["bg_entry"])]}},
            "TEntry": {"configure": {"fieldbackground": c["bg_entry"], "foreground": c["fg_entry"],
                                     "insertcolor": c["fg_entry"]}},
            "TMenubutton": {"configure": {"background": c["bg_widget"], "foreground": c["fg_widget"]}},
            "TScrollbar": {
                "configure": {"background": c["bg"], "troughcolor": c["bg_widget"], "arrowcolor": c["fg_widget"]}},
            "Filter.TButton": {"configure": {"padding": 6, "relief": "flat"},
                               "map": {"background": [('selected', c["bg_select"]), ('!selected', c["bg_widget"])],
                                       "foreground": [('selected', c["fg_select"]), ('!selected', c["fg_widget"])]}},
            "Action.TCombobox": {"configure": {"padding": 6, "relief": "flat", "arrowsize": 0},
                                 "map": {"background": [("readonly", c["bg_widget"]), ("active", "#5A5A5A")],
                                         "fieldbackground": [("readonly", c["bg_widget"])],
                                         "foreground": [("readonly", c["fg_widget"])],
                                         "selectbackground": [("readonly", c["bg_widget"])],
                                         "selectforeground": [("readonly", c["fg_widget"])]}},
            "Toggle.TButton": {"configure": {"padding": (2, 5), "relief": "flat"},
                               "map": {"background": [("active", "#5A5A5A"), ("!active", c["bg_widget"])]}},
            # Status Bar Style
            "Status.TLabel": {"configure": {"background": c["bg_status"], "foreground": c["fg_status"]}},
            "Status.TFrame": {"configure": {"background": c["bg_status"]}}
        })

    # --- Theme toggle function ---
    def toggle_theme(self):
        if self.current_theme == "light":
            self.current_theme = "dark"
        else:
            self.current_theme = "light"

        self.style.theme_use(self.current_theme)
        self.reset_tag_colors()

        # Update status bar background explicitly if required by the theme
        c = self.colors[self.current_theme]
        self.status_bar_frame.config(style="Status.TFrame")
        self.status_label.config(style="Status.TLabel")

    # --- Methods to hide/show sidebar ---
    def hide_sidebar(self):
        self.right_frame.pack_forget()
        self.sidebar_toggle_frame.pack_forget()
        self.main_frame.pack_forget()

        self.sidebar_toggle_frame.pack(side=tk.RIGHT, fill=tk.Y, pady=10)
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))

        self.sidebar_toggle_button.config(text="<", command=self.show_sidebar)

    def show_sidebar(self):
        self.right_frame.pack_forget()
        self.sidebar_toggle_frame.pack_forget()
        self.main_frame.pack_forget()

        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10)
        self.sidebar_toggle_frame.pack(side=tk.RIGHT, fill=tk.Y, pady=10)
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))

        self.sidebar_toggle_button.config(text=">", command=self.hide_sidebar)

    # --- Status Bar Update ---
    def update_status_bar(self, current_count, total_count_in_board):
        if self.full_df is None:
            self.status_label.config(text="Ready. No data loaded.")
            return

        board_name = self.current_board_view
        view_mode = "Edited" if self.current_view == "edited" else "Original"

        if total_count_in_board == 0:
            status_text = f"Board: {board_name} | View: {view_mode} | Status: No data for this board."
        elif current_count == total_count_in_board:
            status_text = f"Board: {board_name} | View: {view_mode} | Showing all {current_count} items."
        else:
            status_text = (f"Board: {board_name} | View: {view_mode} | "
                           f"Showing {current_count} of {total_count_in_board} items (Filtered).")

        self.status_label.config(text=status_text)

    def _set_buttons_state(self, state):
        """Helper to enable/disable all main control buttons and menus."""

        # File Menu
        self.file_menu.entryconfig("Save", state=state)
        self.file_menu.entryconfig("Save As...", state=state)
        self.file_menu.entryconfig("Export", state=state)

        # Edit Menu
        self.edit_menu.entryconfig("Change Classification", state=state)
        self.edit_menu.entryconfig("Filter By Classification", state=state)

        # View Menu
        self.view_menu.entryconfig("Preview Export...", state=state)
        self.view_menu.entryconfig("Summary View...", state=state)
        # Radio buttons are handled by post_load_setup/show_edited_file/show_past_file

        # Sidebar Buttons
        try:
            self.reset_button.config(state=state)
            self.reload_logic_button.config(state=state)
            self.stats_button.config(state=state)
        except tk.TclError:
            pass

    def show_loading_screen(self, message="Loading..."):
        """Displays a modal loading window with a progress bar."""
        if self.loading_window:
            self.loading_window.destroy()

        self.loading_window = tk.Toplevel(self.root)
        self.loading_window.transient(self.root)
        self.loading_window.grab_set()
        self.loading_window.title("Processing")

        # Center the window
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 150
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 50
        self.loading_window.geometry(f"300x100+{x}+{y}")
        self.loading_window.resizable(False, False)

        frame = ttk.Frame(self.loading_window, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text=message).pack(pady=(0, 10))

        progress = ttk.Progressbar(frame, mode='indeterminate', length=250)
        progress.pack()
        progress.start(10)

    def hide_loading_screen(self):
        """Hides the loading window."""
        if self.loading_window:
            self.loading_window.destroy()
            self.loading_window = None

    def check_thread_queue(self):
        """Checks the queue for results from the worker thread."""
        try:
            result = self.thread_queue.get_nowait()

            self.hide_loading_screen()
            self._set_buttons_state(tk.NORMAL)
            self.handle_thread_result(result)

        except queue.Empty:
            self.root.after(100, self.check_thread_queue)

    def handle_thread_result(self, result):
        """Processes the result from the worker thread in the main GUI thread."""

        task_type = result.get('type')
        status = result.get('status')

        if status == 'error':
            messagebox.showerror(f"Error in {task_type}", f"An error occurred:\n\n{result.get('error')}")
            return

        if task_type == 'data_load' and status == 'success':
            (self.full_df, self.original_df, board_names, file_path) = result['data']
            file_name = os.path.basename(file_path)
            self.current_job_file_path = None
            self.root.title(f"Excel Classifier Viewer - {file_name} (Unsaved)")
            self.post_load_setup(board_names)

        elif task_type == 'job_load' and status == 'success':
            (self.full_df, self.original_df, self.summary_config_data, self.remarks_config_data, file_path) = result[
                'data']
            self.current_job_file_path = file_path
            self.root.title(f"Excel Classifier Viewer - {os.path.basename(file_path)}")
            # When loading a job file, derive board names from the loaded data
            board_names = self.full_df["Board"].unique().tolist()
            self.post_load_setup(board_names_from_trd=board_names)

        elif task_type == 'config_load' and status == 'success':
            config_name = result.get('config_name')
            new_data = result.get('data')

            if config_name == "summary":
                self.summary_config_data.clear()
                self.summary_config_data.extend([tuple(row) for row in new_data])
            elif config_name == "remarks":
                self.remarks_config_data.clear()
                self.remarks_config_data.extend([tuple(row) for row in new_data])

            messagebox.showinfo("Config Import Successful",
                                f"Configuration data updated from:\n{result.get('file_path')}")

        elif task_type == 'preview' and status == 'success':
            self._create_preview_window(result['data'])

        elif task_type == 'export_current' and status == 'success':
            messagebox.showinfo("Export Successful", f"Current view saved successfully to:\n{result.get('file_path')}")

        elif task_type == 'export_multisheet' and status == 'success':
            messagebox.showinfo("Export Successful",
                                f"Filtered multi-sheet file saved successfully to:\n{result.get('file_path')}")

        elif task_type == 'export_all_boards' and status == 'success':
            messagebox.showinfo("Export Successful",
                                f"Successfully exported {result.get('count')} board files to:\n{result.get('folder_path')}")

        elif task_type == 'save_job' and status == 'success':
            file_path = result.get('file_path')
            self.current_job_file_path = file_path
            self.root.title(f"Excel Classifier Viewer - {os.path.basename(file_path)}")
            messagebox.showinfo("Save Successful", f".job file saved successfully to:\n{file_path}")

        elif task_type == 'remarks_update' and status == 'success':
            self.refresh_treeview()
            messagebox.showinfo("Success",
                                "The 'es_remarks' column has been updated based on the current configuration.")

    # --- Sidebar Setup ---
    def setup_menubuttons(self):
        self.filter_label = ttk.Label(self.right_frame, text="Classifier Filter", font=("Arial", 14, "bold"))
        self.filter_label.pack(side=tk.TOP, pady=10)

        self.reset_button = ttk.Button(self.right_frame, text="Reset View", command=self.reset_highlight,
                                       state=tk.DISABLED)
        self.reset_button.pack(side=tk.TOP, pady=(10, 5), padx=10, fill=tk.X)

        self.reload_logic_button = ttk.Button(self.right_frame, text="Reload P/F/NC Logic",
                                              command=self.reload_original_logic, state=tk.DISABLED)
        self.reload_logic_button.pack(side=tk.TOP, pady=(5, 20), padx=10, fill=tk.X)

        self.stats_button = ttk.Button(self.right_frame, text="Calculate Stats", command=self.calculate_stats,
                                       state=tk.DISABLED)
        self.stats_button.pack(side=tk.TOP, pady=(20, 5), padx=10, fill=tk.X)

        self.text_filter_separator = ttk.Separator(self.right_frame, orient='horizontal')
        self.text_filter_separator.pack(side=tk.TOP, pady=10, fill=tk.X)

        self.text_filter_label = ttk.Label(self.right_frame, text="Text Filters", font=("Arial", 12, "bold"))
        self.text_filter_label.pack(side=tk.TOP, pady=(0, 5))

        self.style_filter_label = ttk.Label(self.right_frame, text="Style (starts with):")
        self.style_filter_label.pack(side=tk.TOP, fill=tk.X, padx=10)
        self.style_filter_entry = ttk.Entry(self.right_frame, textvariable=self.style_filter)
        self.style_filter_entry.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(0, 10))

        self.part_filter_label = ttk.Label(self.right_frame, text="Part (starts with):")
        self.part_filter_label.pack(side=tk.TOP, fill=tk.X, padx=10)
        self.part_filter_entry = ttk.Entry(self.right_frame, textvariable=self.part_filter)
        self.part_filter_entry.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(0, 10))

        self.remark_filter_label = ttk.Label(self.right_frame, text="Remark (starts with):")
        self.remark_filter_label.pack(side=tk.TOP, fill=tk.X, padx=10)
        self.remark_filter_entry = ttk.Entry(self.right_frame, textvariable=self.remark_filter)
        self.remark_filter_entry.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(0, 10))

    # --- End Sidebar Setup ---

    def open_summary_config(self):
        """Opens the Summary Configuration window."""
        title = "Summary Config"
        columns = ["Component Name", "Short Name"]
        self._create_config_view(title, columns, self.summary_config_data, self.summary_config_data)

    def open_remarks_config(self):
        """Opens the Remarks Configuration window, which includes the logic run button."""
        title = "Remarks Config"
        # UPDATED: Only two columns are needed now
        columns = ["Special Names in Remark", "Full Form"]
        self._create_config_view(title, columns, self.remarks_config_data, self.remarks_config_data)

    def open_summary_view(self):
        """
        Opens a view summarizing the classification status (C/NC/PC) of each component part.
        """
        if self.full_df is None:
            messagebox.showwarning("No Data", "Please import a data file first to generate the Summary View.")
            return

        title = "Summary View"
        columns = ["Part", "C/NC", "Short Name"]

        # This logic generates a list of unique parts and determines their C/NC status
        unique_parts = self.full_df[self.full_df['Part'].notna() & (self.full_df['Part'] != "")]['Part'].unique()
        unique_parts.sort()

        data = []
        for part in unique_parts:
            category_name = self._get_component_category(part)
            part_classifications = self.full_df.loc[self.full_df['Part'] == part, 'Classification'].dropna().unique()

            is_covered = any(c in ['P', 'F'] for c in part_classifications)
            is_not_covered = 'NC' in part_classifications

            c_nc = ""
            if is_covered and is_not_covered:
                # Although this simple view doesn't explicitly track PC, we denote complexity
                c_nc = "C/NC"
            elif is_covered:
                c_nc = "C"
            elif is_not_covered:
                c_nc = "NC"
            else:
                c_nc = "Unclassified"

            data.append((part, c_nc, category_name))

        # Pass None as the data_list_to_edit, as this view is read-only
        self._create_config_view(title, columns, data, None)

    def _get_part_short_name(self, part_str):
        """
        Kept for compatibility, but mainly superseded by _get_component_category.
        Removes all numeric digits from a string (e.g., 'R10' -> 'R').
        """
        if pd.isna(part_str):
            return ""
        return re.sub(r'\d+', '', str(part_str))

    def _get_component_category(self, part_name):
        """
        Classifies the component based on its Reference Designator prefix
        and the user-configurable rules in self.summary_config_data.
        """
        if pd.isna(part_name) or not str(part_name).strip():
            return "Others"

        part_name = str(part_name).upper().strip()

        # Iterate through the configurable summary data
        for component_name, short_names_str in self.summary_config_data:
            if not short_names_str.strip():
                continue

            # Split the short names (prefixes) by comma and strip whitespace
            prefixes = [p.strip() for p in short_names_str.split(',') if p.strip()]

            for prefix in prefixes:
                if part_name.startswith(prefix):
                    return component_name

        return "Others"

    def _create_config_view(self, title, columns, data, data_list_to_edit=None):
        """Helper function to create a simple Toplevel window with a Treeview."""
        config_win = tk.Toplevel(self.root)
        config_win.title(title)
        config_win.transient(self.root)
        config_win.grab_set()
        config_win.geometry("600x300")

        frame = ttk.Frame(config_win, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        if not data and data_list_to_edit is None:
            ttk.Label(frame, text="No data to display.").pack()
            close_btn = ttk.Button(frame, text="Close", command=config_win.destroy)
            close_btn.pack(side=tk.RIGHT, pady=10)
            return

        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')

        for col in columns:
            header_width = len(col) * 9 + 20
            tree.heading(col, text=col)
            tree.column(col, width=max(150, header_width), anchor="w")

        for i, row_data in enumerate(data):
            # Ensure row_data is a tuple of strings for consistency
            tree.insert("", "end", iid=i, values=tuple(str(x) for x in row_data))

        tree.pack(fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

        if data_list_to_edit is not None:
            # Existing edit buttons
            new_row_btn = ttk.Button(btn_frame, text="New Row",
                                     command=lambda: self._open_new_row_popup(tree, config_win, data_list_to_edit,
                                                                              columns))
            new_row_btn.pack(side=tk.LEFT, padx=(0, 5))

            delete_btn = ttk.Button(btn_frame, text="Delete Row",
                                    command=lambda: self._delete_selected_rows(tree, data_list_to_edit, config_win))
            delete_btn.pack(side=tk.LEFT, padx=5)

            edit_btn = ttk.Button(btn_frame, text="Edit",
                                  command=lambda: self._open_edit_popup(tree, config_win, data_list_to_edit))
            edit_btn.pack(side=tk.LEFT, padx=5)

            # --- NEW: Run Logic Button for Remarks Config ---
            if title == "Remarks Config":
                run_logic_btn = ttk.Button(btn_frame, text="Run Remarks Logic",
                                           command=lambda: self._run_remarks_logic_update(config_win))
                run_logic_btn.pack(side=tk.LEFT, padx=(15, 5))
            # --- END NEW ---

        close_btn = ttk.Button(btn_frame, text="Close", command=config_win.destroy)
        close_btn.pack(side=tk.RIGHT)

        self.root.wait_window(config_win)

    def _open_edit_popup(self, tree, parent_window, data_list):
        """Opens a new popup window to edit a selected row from a config tree."""
        selected_item = tree.selection()

        if not selected_item or len(selected_item) > 1:
            messagebox.showwarning("Selection Error", "Please select exactly one row to edit.", parent=parent_window)
            return

        iid = selected_item[0]
        item_values = tree.item(iid, 'values')
        columns = list(tree['columns'])
        data_index = int(iid)  # The iid is currently the index in the list

        edit_win = tk.Toplevel(parent_window)
        edit_win.transient(parent_window)
        edit_win.grab_set()
        edit_win.title(f"Edit Entry (Row {data_index + 1})")

        main_frame = ttk.Frame(edit_win, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.X, expand=True)

        entry_vars = []
        for i, col in enumerate(columns):
            lbl = ttk.Label(form_frame, text=f"{col}:")
            lbl.grid(row=i, column=0, sticky="w", padx=5, pady=5)

            var = tk.StringVar(value=item_values[i])
            entry_vars.append(var)

            ent = ttk.Entry(form_frame, textvariable=var, width=50)
            ent.grid(row=i, column=1, sticky="ew", padx=5, pady=5)

        form_frame.columnconfigure(1, weight=1)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        # --- Clear Fields Button ---
        def clear_fields():
            for var in entry_vars:
                var.set("")

        clear_btn = ttk.Button(btn_frame, text="Clear Fields", command=clear_fields)
        clear_btn.pack(side=tk.LEFT, padx=5)
        # --- END ---

        save_btn = ttk.Button(btn_frame, text="Save",
                              command=lambda: self._save_edit_data(data_list, data_index, entry_vars, tree, iid,
                                                                   edit_win))
        save_btn.pack(side=tk.RIGHT, padx=5)

        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=edit_win.destroy)
        cancel_btn.pack(side=tk.RIGHT)

    def _save_edit_data(self, data_list, data_index, entry_vars, tree, iid, edit_win):
        """Saves the data from the edit popup back to the source list and treeview."""
        try:
            new_values = tuple(var.get() for var in entry_vars)

            # Check if all fields are empty (prevents saving an empty row on edit)
            if not any(val.strip() for val in new_values):
                if messagebox.askyesno("Confirm Delete",
                                       "All fields are empty. Do you want to delete this row instead?",
                                       parent=edit_win):
                    self._delete_selected_rows(tree, data_list, edit_win, iid_to_delete=iid)
                    edit_win.destroy()
                    return
                return  # Keep edit window open

            # Update the underlying data list (data_list is passed by reference)
            data_list[data_index] = new_values

            # Update the treeview item visually
            tree.item(iid, values=new_values)
            edit_win.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Could not save changes: {e}", parent=edit_win)

    def _open_new_row_popup(self, tree, parent_window, data_list, columns):
        """Opens a new popup window to create a new row for a config tree."""
        new_row_win = tk.Toplevel(parent_window)
        new_row_win.transient(parent_window)
        new_row_win.grab_set()
        new_row_win.title("Add New Entry")

        main_frame = ttk.Frame(new_row_win, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.X, expand=True)

        entry_vars = []
        for i, col in enumerate(columns):
            lbl = ttk.Label(form_frame, text=f"{col}:")
            lbl.grid(row=i, column=0, sticky="w", padx=5, pady=5)

            var = tk.StringVar(value="")
            entry_vars.append(var)

            ent = ttk.Entry(form_frame, textvariable=var, width=50)
            ent.grid(row=i, column=1, sticky="ew", padx=5, pady=5)

        form_frame.columnconfigure(1, weight=1)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        # --- Clear Fields Button ---
        def clear_fields():
            for var in entry_vars:
                var.set("")

        clear_btn = ttk.Button(btn_frame, text="Clear Fields", command=clear_fields)
        clear_btn.pack(side=tk.LEFT, padx=5)
        # --- END ---

        save_btn = ttk.Button(btn_frame, text="Save",
                              command=lambda: self._save_new_row(data_list, entry_vars, tree, new_row_win))
        save_btn.pack(side=tk.RIGHT, padx=5)

        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=new_row_win.destroy)
        cancel_btn.pack(side=tk.RIGHT)

    def _save_new_row(self, data_list, entry_vars, tree, new_row_win):
        """Saves the data from the new row popup to the source list and treeview."""
        try:
            new_values = tuple(var.get() for var in entry_vars)

            if not any(val.strip() for val in new_values):
                messagebox.showwarning("Empty Row", "Cannot add an empty row.", parent=new_row_win)
                return

            data_list.append(new_values)

            # Clear and re-populate the tree with the new data to correctly establish IIDs
            tree.delete(*tree.get_children())
            for i, row_data in enumerate(data_list):
                tree.insert("", "end", iid=i, values=tuple(str(x) for x in row_data))

            # Select and show the new row
            new_iid = len(data_list) - 1
            tree.selection_set(new_iid)
            tree.see(new_iid)

            new_row_win.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Could not save new row: {e}", parent=new_row_win)

    def _delete_selected_rows(self, tree, data_list, parent_window, iid_to_delete=None):
        """Deletes one or more selected rows from a config tree and its data list."""
        if iid_to_delete is not None:
            selected_items = [iid_to_delete]
        else:
            selected_items = tree.selection()

        if not selected_items:
            messagebox.showwarning("No Selection", "Please select one or more rows to delete.", parent=parent_window)
            return

        if not messagebox.askyesno("Confirm Deletion",
                                   f"Are you sure you want to delete {len(selected_items)} selected row(s)?\n\nThis action cannot be undone.",
                                   parent=parent_window):
            return

        try:
            # Get a set of indices (as integers) to delete
            indices_to_delete = {int(iid) for iid in selected_items}

            # Create a new data list, excluding items at the selected indices
            new_data_list = [item for i, item in enumerate(data_list) if i not in indices_to_delete]

            # Update the original list in-place to maintain the reference
            data_list.clear()
            data_list.extend(new_data_list)

            # Clear and re-populate the tree with the new data, re-establishing IIDs
            tree.delete(*tree.get_children())
            for i, row_data in enumerate(data_list):
                tree.insert("", "end", iid=i, values=tuple(str(x) for x in row_data))

        except Exception as e:
            messagebox.showerror("Error", f"Could not delete rows: {e}", parent=parent_window)

    def _get_mpin_group(self, part_str):
        """
        Extracts the base component name from a multi-pin part string.
        e.g., "Q1_1_1" -> "Q1", "U10(pin1)" -> "U10"
        """
        if pd.isna(part_str):
            return pd.NA

        part_str = str(part_str).strip()
        if "_" not in part_str and "(" not in part_str:
            return pd.NA

        # Handle parentheses first (e.g., U10(pin1))
        paren_index = part_str.find("(")
        if paren_index != -1:
            return part_str[:paren_index]

        # Handle typical multi-pin format (e.g., Q1_1_1)
        first_underscore_index = part_str.find("_")
        if first_underscore_index != -1:
            return part_str[:first_underscore_index]

        return part_str

    def parse_trd_file(self, file_path):
        """Parses a TRD file structure into a single DataFrame."""
        all_board_dfs = []
        board_names = []

        current_board_name = "Board 1"
        current_data_lines = []
        headers = []
        in_component_block = False
        board_name_pattern = re.compile(r"====================(Board .*)====================")

        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue

                board_match = board_name_pattern.match(line)
                if board_match:
                    # Save previous board data
                    if in_component_block and current_data_lines and headers:
                        df = pd.DataFrame(current_data_lines, columns=headers)
                        df["Board"] = current_board_name
                        all_board_dfs.append(df)
                        current_data_lines = []

                    current_board_name = board_match.group(1).strip()
                    if current_board_name not in board_names:
                        board_names.append(current_board_name)
                    in_component_block = False
                    continue

                if line.startswith("Data Mapping="):
                    # Headers are defined here
                    headers = [h.strip() for h in line.split("=")[1].split(',') if h.strip()]
                    continue

                if line.startswith("***************Component***************"):
                    in_component_block = True
                    continue

                if line.startswith("++++++++++++++++++++++++++++++++++++++++++++++++++"):
                    # End of component block/file
                    if in_component_block and current_data_lines and headers:
                        df = pd.DataFrame(current_data_lines, columns=headers)
                        df["Board"] = current_board_name
                        all_board_dfs.append(df)
                    current_data_lines = []
                    in_component_block = False
                    continue

                if in_component_block and headers:
                    # Data lines are semicolon-separated in typical TRD
                    current_data_lines.append(line.split(';'))

        # Check for trailing data if file ended abruptly
        if in_component_block and current_data_lines and headers:
            df = pd.DataFrame(current_data_lines, columns=headers)
            df["Board"] = current_board_name
            all_board_dfs.append(df)

        if not all_board_dfs:
            raise ValueError("Could not find component data in TRD file.")

        combined_df = pd.concat(all_board_dfs, ignore_index=True)
        return combined_df, board_names

    def expand_rows_by_remark(self, df):
        """
        Expands rows where 'Remark' contains '/' or ',' to create separate rows
        for each component mentioned in the remark, inheriting most data from the parent row.
        """
        new_rows = []
        split_pattern = re.compile(r'[/,]')

        for _, row in df.iterrows():
            remark_str = str(row['Remark']).strip()
            part_str = str(row['Part']).strip()

            if '/' in remark_str or ',' in remark_str:
                parts_in_remark = split_pattern.split(remark_str)

                # Use the original part name as the "Parallel with" identifier
                main_part = part_str if part_str else row['ID']  # Fallback to ID if part is missing
                parallel_with = f"Parallel with {main_part}"

                for part in parts_in_remark:
                    part_cleaned = part.strip()

                    if part_cleaned:
                        new_row = row.to_dict()
                        new_row['Remark'] = part_cleaned
                        new_row['Part'] = part_cleaned  # The new Part is the item in the Remark

                        new_row['Is_Grouped'] = True
                        new_row['Parallel'] = parallel_with
                        new_rows.append(new_row)
            else:
                # Normal row: check if the Remark is actually a Part name (e.g., for multi-pins)
                new_row = row.to_dict()
                remark_val = remark_str
                part_val = part_str

                # Check for specific multi-pin format pattern in remark but not in part
                if "_" in remark_val and "_" not in part_val and not part_val:
                    new_row['Part'] = remark_val
                else:
                    new_row['Part'] = part_val

                new_row['Remark'] = remark_val
                new_row['Is_Grouped'] = False
                new_row['Parallel'] = ""
                new_rows.append(new_row)

        return pd.DataFrame(new_rows)

    def open_excel(self):
        """Opens file dialog for data import and starts the thread to load the data."""
        file_path = filedialog.askopenfilename(
            title="Open Data File",
            filetypes=[
                ("Data Files", "*.xlsx *.xls *.csv *.trd"),
                ("Excel/CSV Files", "*.xlsx *.xls *.csv"),
                ("Excel Files", "*.xlsx *.xls"),
                ("CSV Files", "*.csv"),
                ("TRD Files", "*.trd"),
                ("All Files", "*.*")
            ]
        )
        if not file_path:
            return

        self._set_buttons_state(tk.DISABLED)
        self.show_loading_screen("Loading data file...")

        threading.Thread(
            target=self._thread_load_data,
            args=(file_path,),
            daemon=True
        ).start()

        self.check_thread_queue()

    def _thread_load_data(self, file_path):
        """Worker thread for loading Excel/CSV/TRD files."""
        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            df = None
            board_names_from_trd = []

            if file_ext in ['.xlsx', '.xls']:
                df = pd.read_excel(file_path, engine='openpyxl')
                if "Board" not in df.columns:
                    df["Board"] = "Default"

            # --- Robust CSV Import Logic mimicking TRD parser ---
            elif file_ext == '.csv':

                all_board_dfs = []
                current_board_name = "Default"
                current_data_lines = []
                headers = []
                in_component_block = False

                BOARD_START_PATTERN = re.compile(r"====================(Board .*)====================")
                DATA_MAP_PREFIX = "Data Mapping="
                COMPONENT_START_MARKER = "***************Component***************"
                DATA_END_MARKER = "++++++++++++++++++++++++++++++++++++++++++++++++++"

                # 1. Read the raw file content
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content_lines = [line.strip() for line in f.readlines()]

                # 2. Pre-parse lines to find markers and extract data
                for line in content_lines:
                    if not line: continue

                    # A. Detect Data Mapping/Headers
                    if line.startswith(DATA_MAP_PREFIX):
                        headers = [h.strip() for h in line.split("=")[1].split(',') if h.strip()]
                        continue

                    # B. Detect Board Name change
                    board_match = BOARD_START_PATTERN.match(line)
                    if board_match:
                        # Save previous board data if in component block
                        if in_component_block and current_data_lines and headers:
                            df_block = pd.DataFrame(current_data_lines, columns=headers)
                            df_block["Board"] = current_board_name
                            all_board_dfs.append(df_block)
                            current_data_lines = []

                        current_board_name = board_match.group(1).strip()
                        board_names_from_trd.append(current_board_name)
                        in_component_block = False
                        continue

                    # C. Detect Component Data Start
                    if COMPONENT_START_MARKER in line:
                        in_component_block = True
                        continue

                    # D. Detect Data End Marker
                    if DATA_END_MARKER in line:
                        if in_component_block and current_data_lines and headers:
                            df_block = pd.DataFrame(current_data_lines, columns=headers)
                            df_block["Board"] = current_board_name
                            all_board_dfs.append(df_block)
                        in_component_block = False
                        current_data_lines = []
                        continue

                    # E. Collect data lines if inside the component block
                    if in_component_block and headers:
                        try:
                            # Use StringIO and csv.reader for robust parsing of comma-separated data
                            reader = csv.reader(io.StringIO(line))
                            row = next(reader)

                            # Handle row length mismatch
                            if len(row) < len(headers):
                                row.extend([''] * (len(headers) - len(row)))
                            if len(row) > len(headers):
                                row = row[:len(headers)]

                            current_data_lines.append(row)
                        except Exception:
                            # Skip lines that cannot be parsed as a data row
                            pass

                # Final check for any remaining unsaved data after EOF
                if in_component_block and current_data_lines and headers:
                    df_block = pd.DataFrame(current_data_lines, columns=headers)
                    df_block["Board"] = current_board_name
                    all_board_dfs.append(df_block)

                if not all_board_dfs:
                    raise ValueError(
                        "Could not find component data or consistent headers in the CSV file using TRD markers.")

                df = pd.concat(all_board_dfs, ignore_index=True)

            # --- END Robust CSV Import ---

            elif file_ext == '.trd':
                # Existing TRD parser
                df, board_names_from_trd = self.parse_trd_file(file_path)

            else:
                raise ValueError(f"File type '{file_ext}' is not supported for data import.")

            # Continue processing the DataFrame (df) normally
            df = df.dropna(axis=1, how='all')

            # --- Robust Column Initialization ---
            required_internal_cols = [
                "Classification", "Override", "Is_Grouped", "Parallel",
                "Coverage", "Pin_Type", "MPin_Group", "es_remarks"
            ]

            for col in required_internal_cols:
                if col not in df.columns:
                    # Initialize with empty string for es_remarks, pd.NA otherwise
                    df[col] = '' if col == 'es_remarks' else pd.NA
            # --- End Robust Column Initialization ---

            # Check if necessary TRD columns exist before proceeding
            required_cols_present = all(col in df.columns for col in self.trd_columns)

            if not required_cols_present:
                missing_cols = [col for col in self.trd_columns if col not in df.columns]
                raise ValueError(
                    f"Imported file is missing crucial columns defined in Data Mapping: {', '.join(missing_cols)}")

            # Create a clean subset with only TRD columns for expansion
            clean_df = df[[col for col in self.trd_columns if col in df.columns]].copy()

            # Expand rows based on Remark for parallel components
            expanded_df = self.expand_rows_by_remark(clean_df)

            # Re-calculate derived columns (these are not carried over from the original df after expansion)
            expanded_df["Override"] = pd.NA

            # Temporary Pin_Type calculation (uses Part name)
            expanded_df["Pin_Type"] = expanded_df["Part"].astype(str).apply(
                lambda x: "mpin" if "_" in x or "(" in x else "2pin")

            # Default MPin_Group only for original multi-pin style parts
            expanded_df["MPin_Group"] = expanded_df.apply(
                lambda row: self._get_mpin_group(row["Part"]) if row["Pin_Type"] == "mpin" else pd.NA, axis=1)

            classification_map = {
                'P': 'P', 'F': 'F', 'N': 'NC'
            }
            expanded_df["Classification"] = expanded_df["Result"].map(classification_map).fillna(expanded_df["Result"])
            expanded_df["Coverage"] = expanded_df["Classification"].apply(self.get_coverage_text)

            expanded_df['es_remarks'] = ''

            full_df = expanded_df.reset_index(drop=True)

            # --- NEW AGGREGATION LOGIC: Group repeated 2-Pin parts ---

            # 1. Identify 2-Pin parts that are repeated (more than one row)
            # We only care about parts not already grouped (Is_Grouped == False)
            mask_simple_parts = (full_df["Pin_Type"] == "2pin") & (full_df["Is_Grouped"] == False)

            repeated_parts = full_df[mask_simple_parts]['Part'].value_counts()
            repeated_parts = repeated_parts[repeated_parts > 1].index.tolist()

            # 2. For these repeated simple parts, set the MPin_Group to the Part name itself.
            if repeated_parts:
                # Convert Part to string for reliable masking
                full_df['Part_Str'] = full_df['Part'].astype(str)

                for part_name in repeated_parts:
                    # Find all indices in the full_df where Part equals the repeated name AND it's a simple 2-pin part
                    mask = (full_df['Part_Str'] == part_name) & mask_simple_parts
                    full_df.loc[mask, 'MPin_Group'] = part_name

                full_df.drop(columns=['Part_Str'], inplace=True)

            # --- End NEW AGGREGATION LOGIC ---

            # Ensure original_df is a clean copy with all required columns
            original_df = full_df.copy()

            result_data = (full_df, original_df, board_names_from_trd, file_path)
            self.thread_queue.put({
                'type': 'data_load',
                'status': 'success',
                'data': result_data
            })

        except Exception as e:
            self.thread_queue.put({
                'type': 'data_load',
                'status': 'error',
                'error': str(e)
            })

    def import_config_csv(self):
        """
        Opens a file dialog to import configuration data (Summary or Remarks config) from a CSV file.
        """
        file_path = filedialog.askopenfilename(
            title="Import Configuration CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        self._set_buttons_state(tk.DISABLED)  # Disable controls while configuration loads
        self.show_loading_screen("Loading configuration data...")

        threading.Thread(
            target=self._thread_load_config_csv,
            args=(file_path,),
            daemon=True
        ).start()

        self.check_thread_queue()

    def _thread_load_config_csv(self, file_path):
        """
        Worker thread to read configuration data from a CSV file and update config lists.
        """
        try:
            # Attempt to read CSV.
            config_df = pd.read_csv(file_path).fillna('').astype(str)

            if config_df.empty:
                raise ValueError("CSV file is empty or could not be read.")

            # --- Determine which configuration list to update based on column headers ---

            # Summary Config columns: ["Component Name", "Short Name"] (2 columns)
            summary_cols = ["Component Name", "Short Name"]
            # Remarks Config columns: ["Special Names in Remark", "Full Form"] (2 columns)
            remarks_cols = ["Special Names in Remark", "Full Form"]

            header_list = [col.strip() for col in config_df.columns.tolist()]

            updated_config_name = None
            new_data = []

            # Check for Summary Config match
            if all(col in header_list for col in summary_cols) and len(header_list) == len(summary_cols):
                new_data = config_df[summary_cols].values.tolist()
                updated_config_name = "summary"

            # Check for Remarks Config match (now checking for 2 columns)
            elif all(col in header_list for col in remarks_cols) and len(header_list) == len(remarks_cols):
                new_data = config_df[remarks_cols].values.tolist()
                updated_config_name = "remarks"

            if not updated_config_name:
                raise ValueError(
                    "CSV columns do not match expected formats for Summary Config (2 columns) or Remarks Config (2 columns).")

            # --- Update the live list in the main thread handler ---
            self.thread_queue.put({
                'type': 'config_load',
                'status': 'success',
                'file_path': file_path,
                'config_name': updated_config_name,
                'data': new_data
            })

        except Exception as e:
            self.thread_queue.put({
                'type': 'config_load',
                'status': 'error',
                'error': str(e)
            })

    def import_job_file(self):
        """Opens file dialog for job file import and starts the thread to load the data."""
        file_path = filedialog.askopenfilename(
            title="Open Job File",
            filetypes=[("Job Files", "*.job"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        self._set_buttons_state(tk.DISABLED)
        self.show_loading_screen("Loading job file...")

        threading.Thread(
            target=self._thread_load_job,
            args=(file_path,),
            daemon=True
        ).start()

        self.check_thread_queue()

    def _thread_load_job(self, file_path):
        """Worker thread for loading .job files. Now loads config data."""
        try:
            with open(file_path, 'rb') as f:
                save_data = pickle.load(f)

            if 'full_df' not in save_data or 'original_df' not in save_data:
                self.thread_queue.put({
                    'type': 'job_load',
                    'status': 'error',
                    'error': "Invalid or corrupted .job file. Missing core dataframes."
                })
                return

            full_df = save_data['full_df']
            original_df = save_data['original_df']

            # Ensure 'es_remarks' column exists in loaded dataframes for older job files
            if 'es_remarks' not in full_df.columns:
                full_df['es_remarks'] = ''
            else:
                full_df['es_remarks'] = full_df['es_remarks'].fillna('')

            if 'es_remarks' not in original_df.columns:
                original_df['es_remarks'] = ''
            else:
                original_df['es_remarks'] = original_df['es_remarks'].fillna('')

            # --- Load Configs with fallback for old files ---
            summary_config = save_data.get('summary_config', self.summary_config_data)
            # Handle potential loading of old 3-column remarks config from older files
            remarks_config_raw = save_data.get('remarks_config', self.remarks_config_data)

            # If loaded config has 3 elements, truncate to 2 to match new structure
            remarks_config = [r[:2] for r in remarks_config_raw]

            # --- END ---

            result_data = (full_df, original_df, summary_config, remarks_config, file_path)
            self.thread_queue.put({
                'type': 'job_load',
                'status': 'success',
                'data': result_data
            })

        except Exception as e:
            self.thread_queue.put({
                'type': 'job_load',
                'status': 'error',
                'error': str(e)
            })

    def save_job_file_as(self):
        """Prompts for a new file path and saves the job file."""
        if self.full_df is None:
            messagebox.showwarning("No Data", "There is no data to save.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Job File As...",
            defaultextension=".job",
            filetypes=[("Job Files", "*.job"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        self._set_buttons_state(tk.DISABLED)
        self.show_loading_screen("Saving .job file...")

        threading.Thread(
            target=self._thread_save_job,
            args=(file_path,),
            daemon=True
        ).start()

        self.check_thread_queue()

    def save_job_file(self):
        """Saves the job file to the current path, or calls save_job_file_as if no path exists."""
        if self.full_df is None:
            messagebox.showwarning("No Data", "There is no data to save.")
            return

        if self.current_job_file_path is None:
            self.save_job_file_as()
        else:
            self._set_buttons_state(tk.DISABLED)
            self.show_loading_screen("Saving .job file...")

            threading.Thread(
                target=self._thread_save_job,
                args=(self.current_job_file_path,),
                daemon=True
            ).start()

            self.check_thread_queue()

    def _thread_save_job(self, file_path):
        """Worker thread for saving .job files. Saves dataframes and config data."""
        try:
            save_data = {
                'full_df': self.full_df,
                'original_df': self.original_df,
                # --- Save Configs (remarks_config is now 2-column) ---
                'summary_config': self.summary_config_data,
                'remarks_config': self.remarks_config_data
                # --- END ---
            }

            with open(file_path, 'wb') as f:
                pickle.dump(save_data, f)

            self.thread_queue.put({
                'type': 'save_job',
                'status': 'success',
                'file_path': file_path
            })

        except Exception as e:
            self.thread_queue.put({
                'type': 'save_job',
                'status': 'error',
                'error': str(e)
            })

    def new_file(self):
        """Resets the application state for a new file."""
        if self.full_df is not None:
            if not messagebox.askyesno("Confirm New File",
                                       "Are you sure you want to create a new file?\n\nAll unsaved changes will be lost."):
                return

        self.full_df = None
        self.original_df = None
        self.current_filtered_df = None
        self.current_job_file_path = None

        self.root.title("Excel Classifier Viewer")
        self.tree.delete(*self.tree.get_children())

        self.board_selector.config(values=[], state=tk.DISABLED)
        self.board_selector.set("")

        self.mpin_group_selector.config(values=[], state=tk.DISABLED)
        self.mpin_group_selector.set("All")

        self.reset_highlight()
        self._set_buttons_state(tk.DISABLED)

        self.view_mode_var.set("edited")
        self.view_menu.entryconfig("Show Edited File", state=tk.DISABLED)
        self.view_menu.entryconfig("Show Past File (Original)", state=tk.DISABLED)

        self.update_status_bar(0, 0)

    def post_load_setup(self, board_names_from_trd):
        """Final setup steps after data has been successfully loaded into dataframes."""
        if self.full_df is None:
            return

        self.show_sidebar()

        # Set up board selector
        unique_boards = self.full_df["Board"].unique()
        if board_names_from_trd:
            # Use names extracted directly from TRD, plus "All Boards"
            self.board_list = ["All Boards"] + sorted(list(set(board_names_from_trd)))
        elif len(unique_boards) > 1 or (len(unique_boards) == 1 and unique_boards[0] != "Default"):
            # Use names from DataFrame, plus "All Boards"
            self.board_list = ["All Boards"] + sorted(list(unique_boards))
        else:
            self.board_list = ["All Boards"]

        self.board_selector.config(values=self.board_list, state="readonly" if len(self.board_list) > 1 else "disabled")
        self.board_selector.set("All Boards")
        self.current_board_view = "All Boards"

        # --- UPDATED: Set up mpin group selector using ALL non-NA MPin_Group values (including new repeated parts) ---
        mpin_groups = self.full_df["MPin_Group"].unique()
        mpin_groups = sorted([g for g in mpin_groups if pd.notna(g)])
        self.mpin_group_list = ["All"] + mpin_groups
        self.mpin_group_selector.config(values=self.mpin_group_list,
                                        state="readonly" if len(self.mpin_group_list) > 1 else "disabled")
        self.mpin_group_selector.set("All")
        self.current_mpin_group_filter = "All"
        # --- END UPDATED ---

        self.reset_highlight()
        self.refresh_treeview()

        self.view_mode_var.set("edited")
        self.view_menu.entryconfig("Show Edited File", state=tk.DISABLED)
        self.view_menu.entryconfig("Show Past File (Original)", state=tk.NORMAL)
        self.reload_logic_button.config(state=tk.NORMAL)  # Enable reload logic button

        self.current_view = "edited"
        self._set_buttons_state(tk.NORMAL)

    def on_board_select(self, event=None):
        self.current_board_view = self.board_selector.get()
        self.refresh_treeview()

    def on_mpin_group_select(self, event=None):
        self.current_mpin_group_filter = self.mpin_group_selector.get()
        self.refresh_treeview()

    def export_current_view(self):
        """Exports the currently filtered/displayed data to a single Excel sheet."""
        if self.current_filtered_df is None or self.current_filtered_df.empty:
            messagebox.showwarning("No Data", "There is no data in the current view to export.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Current View",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        self._set_buttons_state(tk.DISABLED)
        self.show_loading_screen("Exporting current view...")

        # Export columns including the new es_remarks
        df_to_export = self.current_filtered_df.copy()

        threading.Thread(
            target=self._thread_export_current,
            args=(file_path, df_to_export),
            daemon=True
        ).start()

        self.check_thread_queue()

    def _thread_export_current(self, file_path, df_to_export):
        """Worker thread for exporting the current view."""
        try:
            # Include 'es_remarks' in the exported columns
            export_cols = self.trd_columns + ["Classification", "MPin_Group", "Parallel", "Coverage", "es_remarks"]
            export_cols_exist = [col for col in export_cols if col in df_to_export.columns]
            export_df = df_to_export[export_cols_exist].copy()

            export_df.to_excel(file_path, sheet_name="Current View", index=False)

            self.thread_queue.put({
                'type': 'export_current',
                'status': 'success',
                'file_path': file_path
            })
        except Exception as e:
            self.thread_queue.put({
                'type': 'export_current',
                'status': 'error',
                'error': str(e)
            })

    def export_filtered_board_multisheet(self):
        """Exports the current single board view into a multi-sheet Excel file with summary and breakdowns."""
        if self.full_df is None:
            messagebox.showwarning("No Data", "There is no data to process. Please open a file first.")
            return

        if self.current_board_view == "All Boards":
            messagebox.showwarning("Board Not Selected",
                                   "Please select a single board from the dropdown to use this function.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Filtered Excel File (Multi-Sheet)",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        self._set_buttons_state(tk.DISABLED)
        self.show_loading_screen("Generating multi-sheet export...")

        threading.Thread(
            target=self._thread_export_multisheet,
            args=(file_path,),
            daemon=True
        ).start()

        self.check_thread_queue()

    def _thread_export_multisheet(self, file_path):
        """Worker thread for exporting multi-sheet file."""
        try:
            export_data_dict = self._get_filtered_export_data()
            if export_data_dict is None:
                raise ValueError("No data returned from filter. (Board not selected?)")

            # --- Calculate Component Status at the Group Level (for Group_Status column) ---

            df_all = export_data_dict['all_filtered']
            mpin_groups = df_all[df_all["Pin_Type"] == "mpin"].groupby("MPin_Group")

            group_status_map = {}
            for mpin_group, group in mpin_groups:
                pin_classifications = group["Classification"].unique()
                has_covered_pins = any(x in ['P', 'F'] for x in pin_classifications)
                has_nc_pins = 'NC' in pin_classifications

                if has_covered_pins and has_nc_pins:
                    status = 'PC'  # Partially Covered
                elif has_covered_pins:
                    status = 'C'  # Fully Covered
                else:
                    status = 'NC'  # Not Covered
                group_status_map[mpin_group] = status

            # Helper function to apply the group status to pins
            def apply_group_status(df):
                if df.empty:
                    df['Group_Status'] = pd.NA
                    return df

                # Apply group status to MPin rows, keep Classification for others
                df['Group_Status'] = df.apply(
                    lambda row: group_status_map.get(row['MPin_Group'], row['Classification'])
                    if row['Pin_Type'] == 'mpin' else row['Classification'],
                    axis=1
                )
                return df

            # 2. Apply status to relevant dataframes for export
            df_2pin_combined = pd.concat([export_data_dict['df_2pin_c'], export_data_dict['df_2pin_nc']],
                                         ignore_index=True)
            df_mpin_combined = pd.concat([export_data_dict['df_mpin_c'], export_data_dict['df_mpin_nc']],
                                         ignore_index=True)
            df_parallel_combined = pd.concat([export_data_dict['df_parallel_c'], export_data_dict['df_parallel_nc']],
                                             ignore_index=True)

            # 3. Process Config Data and Summary Data (order is important)
            summary_df = self._calculate_summary_data(export_data_dict)
            config_df = self._calculate_config_data(df_all)  # Calculates Config sheet data including Remarks

            # 4. Prepare export columns including the new Group_Status and es_remarks columns
            export_cols = self.trd_columns + ["Classification", "MPin_Group", "Parallel", "Coverage", "Group_Status",
                                              "es_remarks"]

            # Note: 2-Pin doesn't need Group_Status explicitly if we just use Classification,
            # but for consistency in multi-pin logic:
            df_mpin_combined = apply_group_status(df_mpin_combined)
            df_parallel_combined = apply_group_status(df_parallel_combined)

            # Ensure only existing columns are exported for the final sheets
            def filter_export_cols(df_to_filter):
                cols_to_use = [col for col in export_cols if col in df_to_filter.columns]
                return df_to_filter[cols_to_use]

            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                filter_export_cols(df_2pin_combined).to_excel(writer, sheet_name='2-Pin', index=False)
                filter_export_cols(df_mpin_combined).to_excel(writer, sheet_name='Multi-Pin', index=False)
                filter_export_cols(df_parallel_combined).to_excel(writer, sheet_name='Parallel', index=False)
                config_df.to_excel(writer, sheet_name='Config Data',
                                   index=False)  # Config Data sheet uses _calculate_config_data

            self.thread_queue.put({
                'type': 'export_multisheet',
                'status': 'success',
                'file_path': file_path
            })
        except Exception as e:
            self.thread_queue.put({
                'type': 'export_multisheet',
                'status': 'error',
                'error': str(e)
            })

    def export_all_boards_separate(self):
        """Exports all boards to individual Excel files."""
        if self.full_df is None:
            messagebox.showwarning("No Data", "There is no data to export. Please open a file first.")
            return

        folder_path = filedialog.askdirectory(title="Select Folder to Save Board Files")
        if not folder_path:
            return

        self._set_buttons_state(tk.DISABLED)
        self.show_loading_screen("Exporting all boards...")

        threading.Thread(
            target=self._thread_export_all_boards,
            args=(folder_path,),
            daemon=True
        ).start()

        self.check_thread_queue()

    def _thread_export_all_boards(self, folder_path):
        """Worker thread for exporting all boards to separate files."""
        try:
            # Include 'es_remarks'
            export_cols = self.trd_columns + ["Classification", "MPin_Group", "Parallel", "Coverage", "es_remarks"]
            export_cols_exist = [col for col in export_cols if col in self.full_df.columns]

            # Get the list of unique board names
            board_names = self.full_df["Board"].unique().tolist()

            base_filename = "exported_board_data"

            for board in board_names:
                board_df = self.full_df[self.full_df["Board"] == board][export_cols_exist].copy()

                # Sanitize board name for filename
                safe_board_name = re.sub(r'[\\/*?:"<>|]', "", board).replace(" ", "_")
                file_name = f"{base_filename}_{safe_board_name}.xlsx"
                full_path = os.path.join(folder_path, file_name)

                board_df.to_excel(full_path, index=False)

            self.thread_queue.put({
                'type': 'export_all_boards',
                'status': 'success',
                'folder_path': folder_path,
                'count': len(board_names)
            })

        except Exception as e:
            self.thread_queue.put({
                'type': 'export_all_boards',
                'status': 'error',
                'error': str(e)
            })

    def _get_filtered_export_data(self):
        """Applies all current filters to the dataframe and breaks it down into categories."""
        if self.full_df is None:
            return None

        # Determine which source DF to use
        source_df = self.full_df if self.current_view == "edited" else self.original_df
        if source_df is None:
            return None

        # Columns needed for filtering, grouping, and final export/config sheets
        required_cols = list(set(self.trd_columns + self.all_columns))

        # Ensure all columns exist in the source_df before copying
        for col in required_cols:
            if col not in source_df.columns:
                source_df[col] = pd.NA

        filtered_df = source_df.copy()

        # Apply Board Filter
        if self.current_board_view != "All Boards":
            filtered_df = filtered_df[filtered_df["Board"] == self.current_board_view]

        if filtered_df.empty:
            return None  # Return None if board filter yields empty result

        # Apply Pin Type Filter
        if self.current_pin_filter == "2pin":
            filtered_df = filtered_df[filtered_df["Pin_Type"] == "2pin"]
        elif self.current_pin_filter == "mpin":
            filtered_df = filtered_df[filtered_df["Pin_Type"] == "mpin"]

        # Apply Grouped/Single Filter
        if self.current_group_filter == "single":
            filtered_df = filtered_df[filtered_df["Is_Grouped"] == False]
        elif self.current_group_filter == "grouped":
            filtered_df = filtered_df[filtered_df["Is_Grouped"] == True]

        # Apply MPin Group Filter
        if self.current_mpin_group_filter != "All":
            filtered_df = filtered_df[filtered_df["MPin_Group"] == self.current_mpin_group_filter]

        # Apply Text Filters
        style_query = self.style_filter.get().upper()
        if style_query:
            df_to_show = filtered_df[filtered_df["Style"].astype(str).str.upper().str.startswith(style_query)]

        part_query = self.part_filter.get().upper()
        if part_query:
            filtered_df = filtered_df[filtered_df["Part"].astype(str).str.upper().str.startswith(part_query)]

        remark_query = self.remark_filter.get().upper()
        if remark_query:
            filtered_df = filtered_df[filtered_df["Remark"].astype(str).str.upper().str.startswith(remark_query)]

        # Apply Classification Highlight Filter
        if self.current_highlight:
            filtered_df = filtered_df[filtered_df["Classification"] == self.current_highlight]

        # --- Apply Ctrl+F Global Search Filter ---
        if self.current_search_query:
            query = self.current_search_query

            # Columns to search against (must handle NaNs by casting to string)
            search_cols = ['Part', 'Remark', 'Style', 'ID', 'MPin_Group']

            # Combine all search columns into one mask using OR (|) logic
            mask = False
            for col in search_cols:
                if col in filtered_df.columns:
                    # Use str.contains() for general search
                    mask |= filtered_df[col].astype(str).str.upper().str.contains(query, na=False)

            filtered_df = filtered_df[mask]
        # --- End Ctrl+F Filter ---

        # If no rows remain after filtering
        if filtered_df.empty:
            return None

        # --- Breakdown for multi-sheet export ---
        mask_2pin = filtered_df["Pin_Type"] == "2pin"
        mask_is_grouped = filtered_df["Is_Grouped"] == True
        mask_c = lambda df: df['Classification'].isin(['P', 'F'])
        mask_nc = lambda df: df['Classification'] == 'NC'

        # Columns that should appear on the main export sheets (includes es_remarks now)
        export_cols = self.trd_columns + ["Classification", "MPin_Group", "Parallel", "Coverage", "es_remarks"]
        export_cols_exist = [col for col in export_cols if col in filtered_df.columns]

        # Columns needed for internal calculation/config sheet (includes all internal cols)
        required_cols = list(set(self.trd_columns + self.all_columns))
        all_filtered_cols_exist = [col for col in required_cols if col in filtered_df.columns]

        df_2pin = filtered_df[mask_2pin].copy()
        df_mpin = filtered_df[~mask_2pin].copy()
        df_parallel = filtered_df[mask_is_grouped].copy()

        return {
            'df_2pin_c': df_2pin[mask_c(df_2pin)][export_cols_exist].copy(),
            'df_2pin_nc': df_2pin[mask_nc(df_2pin)][export_cols_exist].copy(),
            'df_mpin_c': df_mpin[mask_c(df_mpin)][export_cols_exist].copy(),
            'df_mpin_nc': df_mpin[mask_nc(df_mpin)][export_cols_exist].copy(),
            'df_parallel_c': df_parallel[mask_c(df_parallel)][export_cols_exist].copy(),
            'df_parallel_nc': df_parallel[mask_nc(df_parallel)][export_cols_exist].copy(),
            'all_filtered': filtered_df[all_filtered_cols_exist].copy()
        }

    def preview_export(self):
        """Prepares and displays the export preview window."""
        if self.full_df is None:
            messagebox.showwarning("No Data", "There is no data to process. Please open a file first.")
            return
        if self.current_board_view == "All Boards":
            messagebox.showwarning("Board Not Selected",
                                   "Please select a single board from the dropdown to use this function.")
            return

        self._set_buttons_state(tk.DISABLED)
        self.show_loading_screen("Generating preview data...")

        threading.Thread(
            target=self._thread_generate_preview,
            daemon=True
        ).start()

        self.check_thread_queue()

    def _thread_generate_preview(self):
        """Worker thread for generating preview data."""
        try:
            export_data_dict = self._get_filtered_export_data()
            if export_data_dict is None:
                raise ValueError("No data returned from filter. (Board not selected or filters too strict)")

            self.thread_queue.put({
                'type': 'preview',
                'status': 'success',
                'data': export_data_dict
            })
        except Exception as e:
            self.thread_queue.put({
                'type': 'preview',
                'status': 'error',
                'error': str(e)
            })

    def _create_preview_window(self, data_dict):
        """Creates the Toplevel window to display the export preview tabs."""
        preview_win = tk.Toplevel(self.root)
        preview_win.title(f"Export Preview for: {self.current_board_view} (Filters Applied)")
        preview_win.geometry("900x600")
        preview_win.transient(self.root)
        preview_win.grab_set()

        themed_frame = ttk.Frame(preview_win)
        themed_frame.pack(fill=tk.BOTH, expand=True)

        notebook = ttk.Notebook(themed_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tab_summary = ttk.Frame(notebook)
        tab_2pin = ttk.Frame(notebook)
        tab_mpin = ttk.Frame(notebook)
        tab_parallel = ttk.Frame(notebook)
        tab_config = ttk.Frame(notebook)

        notebook.add(tab_summary, text="Summary")
        notebook.add(tab_2pin, text="2-Pin")
        notebook.add(tab_mpin, text="Multi-Pin")
        notebook.add(tab_parallel, text="Parallel")
        notebook.add(tab_config, text="Config Data")

        self._populate_summary_tab(tab_summary, data_dict)

        # 2-Pin Tab Content
        frame_2pin_c = ttk.Frame(tab_2pin)
        frame_2pin_c.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame_2pin_c, text="2-Pin Covered (C)", font=("Arial", 12, "bold")).pack(pady=(10, 2))
        self._populate_treeview_tab(frame_2pin_c, data_dict['df_2pin_c'])
        ttk.Separator(frame_2pin_c, orient='horizontal').pack(fill='x', pady=5)

        frame_2pin_nc = ttk.Frame(tab_2pin)
        frame_2pin_nc.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame_2pin_nc, text="2-Pin Not Covered (NC)", font=("Arial", 12, "bold")).pack(pady=(10, 2))
        self._populate_treeview_tab(frame_2pin_nc, data_dict['df_2pin_nc'])

        # Multi-Pin Tab Content
        frame_mpin_c = ttk.Frame(tab_mpin)
        frame_mpin_c.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame_mpin_c, text="Multi-Pin Covered (C)", font=("Arial", 12, "bold")).pack(pady=(10, 2))
        self._populate_treeview_tab(frame_mpin_c, data_dict['df_mpin_c'])
        ttk.Separator(frame_mpin_c, orient='horizontal').pack(fill='x', pady=5)

        frame_mpin_nc = ttk.Frame(tab_mpin)
        frame_mpin_nc.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame_mpin_nc, text="Multi-Pin Not Covered (NC)", font=("Arial", 12, "bold")).pack(pady=(10, 2))
        self._populate_treeview_tab(frame_mpin_nc, data_dict['df_mpin_nc'])

        # Parallel Tab Content
        frame_parallel_c = ttk.Frame(tab_parallel)
        frame_parallel_c.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame_parallel_c, text="Parallel Covered (C)", font=("Arial", 12, "bold")).pack(pady=(10, 2))
        self._populate_treeview_tab(frame_parallel_c, data_dict['df_parallel_c'])
        ttk.Separator(frame_parallel_c, orient='horizontal').pack(fill='x', pady=5)

        frame_parallel_nc = ttk.Frame(tab_parallel)
        frame_parallel_nc.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame_parallel_nc, text="Parallel Not Covered (NC)", font=("Arial", 12, "bold")).pack(pady=(10, 2))
        self._populate_treeview_tab(frame_parallel_nc, data_dict['df_parallel_nc'])

        self._populate_config_tab(tab_config, data_dict['all_filtered'])

        ttk.Button(themed_frame, text="Close Preview", command=preview_win.destroy).pack(pady=10)

    def _get_col_width(self, col_name, data):
        """Estimates column width based on header and max data length."""
        header_len = len(str(col_name))

        if data.empty:
            max_data_len = 0
        else:
            try:
                # Calculate max length of string representation of data
                max_data_len = data.astype(str).str.len().max()
            except Exception:
                max_data_len = 0

        # Max width calculation: (max length + padding) * average char width
        # Use 7 for font size ~10-12
        return max(75, (max(header_len, int(max_data_len or 0)) + 2) * 7)

    def _populate_treeview_tab(self, parent_frame, dataframe):
        """Populates a generic treeview within a preview tab, including a search bar."""
        # Container for this specific table section including search
        container = ttk.Frame(parent_frame)
        container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        if dataframe is None or dataframe.empty:
            ttk.Label(container, text="No data for this category.").pack(padx=10, pady=10)
            return

        # --- Search Bar Setup ---
        search_frame = ttk.Frame(container)
        search_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=(0, 5))
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        # Preserve original data in a list of value tuples for searching
        data_rows = [tuple(str(v) for v in row) for _, row in dataframe.iterrows()]

        tree = ttk.Treeview(container, columns=list(dataframe.columns), show="headings")

        def perform_search():
            query = search_var.get().lower()
            # Clear current items
            tree.delete(*tree.get_children())

            # Filter and insert
            for row_values in data_rows:
                if not query or any(query in val.lower() for val in row_values):
                    tree.insert("", "end", values=row_values)

        search_btn = ttk.Button(search_frame, text="Search", command=perform_search)
        search_btn.pack(side=tk.LEFT, padx=(0, 2))

        # Bind Enter key to search
        search_entry.bind("<Return>", lambda e: perform_search())

        def reset_search():
            search_var.set("")
            perform_search()

        reset_btn = ttk.Button(search_frame, text="Reset", command=reset_search)
        reset_btn.pack(side=tk.LEFT)
        # --- End Search Bar ---

        cols = list(dataframe.columns)

        # Configure scrollbars and treeview
        vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')

        # Configure columns
        for col in cols:
            col_width = self._get_col_width(col, dataframe[col])
            tree.heading(col, text=col)
            tree.column(col, width=col_width, anchor="w", stretch=True)

        # Initial population (loads all data since query starts empty)
        perform_search()

        tree.pack(fill=tk.BOTH, expand=True)

    def _calculate_summary_data(self, data_dict):
        """Calculates component-level coverage statistics for the Summary tab."""
        df = data_dict['all_filtered']
        if df is None or df.empty:
            # Provide empty DataFrame with correct column names
            cols = ["S.No", "Item", "Total Qty", "Covered", "Not Covered", "PC", "Coverage %"]
            return pd.DataFrame(columns=cols)

        # Dictionary to store stats: {Category: {Total: 0, Covered: 0, NC: 0, Partial: 0}}
        stats = {}
        category_order = []

        # Initialize stats based on config order
        for name, _ in self.summary_config_data:
            stats[name] = {"Total": 0, "Covered": 0, "NC": 0, "Partial": 0}
            category_order.append(name)

        # Ensure "Others" is a valid key if missing from config
        if "Others" not in stats:
            stats["Others"] = {"Total": 0, "Covered": 0, "NC": 0, "Partial": 0}

        if category_order and category_order[-1] == "Others":
            category_order.pop()

        get_category = self._get_component_category

        # 1. Group items by component name/group ID to determine COMPONENT-LEVEL status

        # A. Multi-Pin Components (Group by MPin_Group)
        df_mpin = df[df["Pin_Type"] == "mpin"]
        for mpin_group, group in df_mpin.groupby("MPin_Group"):
            if pd.isna(mpin_group) or not mpin_group.strip(): continue

            cat = get_category(mpin_group)
            if cat not in stats: cat = "Others"

            pin_classifications = group["Classification"].unique()
            has_covered_pins = any(x in ['P', 'F'] for x in pin_classifications)
            has_nc_pins = 'NC' in pin_classifications

            stats[cat]["Total"] += 1

            if has_covered_pins and has_nc_pins:
                stats[cat]["Partial"] += 1
            elif has_covered_pins:
                stats[cat]["Covered"] += 1
            else:
                stats[cat]["NC"] += 1

        # B. Single/2-Pin Components (Group by Part, excluding multi-pin and parallel splits)
        # Filter for rows that are NOT multi-pin AND NOT part of a grouped/parallel split
        # Only count the unique part names that fall into this category
        df_single = df[(df["Pin_Type"] == "2pin") & (df["Is_Grouped"] == False)]

        for part, group in df_single.groupby("Part"):
            if pd.isna(part) or not part.strip(): continue

            cat = get_category(part)
            if cat not in stats: cat = "Others"

            results = group["Classification"].unique()
            is_covered = any(x in ['P', 'F'] for x in results)

            # For 2-pin/single components, we treat it as binary (Covered or Not)
            if is_covered:
                stats[cat]["Covered"] += 1
            else:
                stats[cat]["NC"] += 1
            stats[cat]["Total"] += 1

        # C. Parallel Components (Group by Parallel field, excluding single/multi-pin parts)
        # Note: Parallel parts are already captured by df_single logic if they are 2-pin,
        # so this logic is often complex to avoid double counting.
        # For simplicity, based on the assumption that 'Is_Grouped=True' means the row is a parallel split item:
        df_parallel_split = df[df["Is_Grouped"] == True]

        for parallel_group, group in df_parallel_split.groupby("Parallel"):
            if pd.isna(parallel_group) or not parallel_group.strip(): continue

            # Use the "main_part" which is included in the Parallel string (e.g., "R10") for categorization
            match = re.search(r'Parallel with (.*)', parallel_group)
            main_part = match.group(1).strip() if match else parallel_group

            cat = get_category(main_part)
            if cat not in stats: cat = "Others"

            results = group["Classification"].unique()
            has_covered_pins = any(x in ['P', 'F'] for x in results)
            has_nc_pins = 'NC' in results

            # We count the main component once
            stats[cat]["Total"] += 1

            if has_covered_pins and has_nc_pins:
                stats[cat]["Partial"] += 1
            elif has_covered_pins:
                stats[cat]["Covered"] += 1
            else:
                stats[cat]["NC"] += 1

        # Combine results into DataFrame
        rows = []
        s_no = 1
        total_summary = {"Total": 0, "Covered": 0, "NC": 0, "Partial": 0}

        # Include "Others" in the final category list for display if it has data
        final_categories = category_order
        if "Others" not in final_categories and stats.get("Others", {"Total": 0})["Total"] > 0:
            final_categories.append("Others")

        for cat in final_categories:
            data = stats.get(cat, {"Total": 0, "Covered": 0, "NC": 0, "Partial": 0})
            total = data["Total"]

            if total == 0: continue

            cov = data["Covered"]
            nc = data["NC"]
            par = data["Partial"]

            # Aggregate totals for the grand summary
            total_summary["Total"] += total
            total_summary["Covered"] += cov
            total_summary["NC"] += nc
            total_summary["Partial"] += par

            # Percentage is based on (Covered + Partially Covered) / Total
            perc = ((cov + par) / total) * 100 if total > 0 else 0.0

            rows.append((s_no, cat, total, cov, nc, par, f"{perc:.1f}%"))
            s_no += 1

        # Add Grand Total Row
        gt_total = total_summary["Total"]
        gt_cov = total_summary["Covered"]
        gt_nc = total_summary["NC"]
        gt_par = total_summary["Partial"]

        gt_perc = ((gt_cov + gt_par) / gt_total) * 100 if gt_total > 0 else 0.0

        rows.append(("", "Total", gt_total, gt_cov, gt_nc, gt_par, f"{gt_perc:.1f}%"))

        # Renaming the column headers to use the short form PC
        cols = ["S.No", "Item", "Total Qty", "Covered", "Not Covered", "PC", "Coverage %"]

        # We must re-index the DataFrame columns to match the new headers
        df_summary = pd.DataFrame(rows, columns=cols)

        return df_summary

    def _populate_summary_tab(self, parent_frame, data_dict):
        f = ttk.Frame(parent_frame, padding=20)
        f.pack(fill=tk.BOTH, expand=True)

        # Use short forms in the column setup
        cols = ["S.No", "Item", "Total Qty", "Covered", "Not Covered", "PC", "Coverage %"]
        tree = ttk.Treeview(f, columns=cols, show="headings", height=10)

        # Configure columns with short form headers
        tree.heading("S.No", text="S.No")
        tree.column("S.No", width=50, anchor="center")
        tree.heading("Item", text="Item")
        tree.column("Item", width=120, anchor="w")
        tree.heading("Total Qty", text="Total Qty")
        tree.column("Total Qty", width=80, anchor="center")
        tree.heading("Covered", text="Covered")
        tree.column("Covered", width=80, anchor="center")
        tree.heading("Not Covered", text="Not Covered")
        tree.column("Not Covered", width=100, anchor="center")
        tree.heading("PC", text="PC")  # Using PC here (Partially Covered)
        tree.column("PC", width=100, anchor="center")
        tree.heading("Coverage %", text="Coverage %")
        tree.column("Coverage %", width=100, anchor="center")

        summary_df = self._calculate_summary_data(data_dict)

        for index, row in summary_df.iterrows():
            tree.insert("", "end", values=list(row))

        tree.pack(fill=tk.X, expand=False, pady=10)

    def _calculate_config_data(self, all_filtered_df):
        """Generates the data for the Config Data export sheet."""
        config_df = pd.DataFrame()
        if all_filtered_df is None or all_filtered_df.empty:
            return config_df

        df = all_filtered_df.copy().reset_index(drop=True)

        config_df["S.No"] = np.arange(1, len(df) + 1)
        config_df["Part"] = df["Part"]
        config_df["STD"] = df["STD"]
        config_df["HL"] = df["HL"]
        config_df["LL"] = df["LL"]
        config_df["C/NC"] = df["Classification"].apply(
            lambda x: 'C' if x in ['P', 'F'] else ('NC' if x == 'NC' else ''))

        # Calculate Group Status for remarks
        mpin_groups = df[df["Pin_Type"] == "mpin"].groupby("MPin_Group")
        group_status_map = {}
        for mpin_group, group in mpin_groups:
            pin_classifications = group["Classification"].unique()
            has_covered_pins = any(x in ['P', 'F'] for x in pin_classifications)
            has_nc_pins = 'NC' in pin_classifications
            if has_covered_pins and has_nc_pins:
                status = ' (PC)'
            elif has_covered_pins:
                status = ' (C)'
            else:
                status = ' (NC)'
            group_status_map[mpin_group] = status

        def combine_remarks(row):
            remarks = []
            group_status_suffix = ""

            # Add Group Status suffix if it's a multi-pin component
            if row["Pin_Type"] == 'mpin' and pd.notna(row['MPin_Group']):
                group_status_suffix = group_status_map.get(row['MPin_Group'], "")

            # 1. Coverage (includes Group Status suffix)
            if pd.notna(row["Coverage"]) and row["Coverage"]:
                remarks.append(row["Coverage"] + group_status_suffix)

            # 2. Parallel Info
            if row["Is_Grouped"] == True and pd.notna(row["Parallel"]) and row["Parallel"]:
                remarks.append(row["Parallel"])

            # 3. Enhanced/Special Remarks (from the new configurable logic)
            if pd.notna(row.get('es_remarks')) and row.get('es_remarks').strip():
                # Use the clean 'es_remarks' string directly, possibly prefixing for clarity
                remarks.append(f"ES: {row['es_remarks'].strip()}")

            return ", ".join(remarks)

        # Append PC status to remarks for multi-pin components in the Config sheet
        config_df["Remarks"] = df.apply(combine_remarks, axis=1)

        return config_df

    def _populate_config_tab(self, parent_frame, all_filtered_df):
        f = ttk.Frame(parent_frame, padding=10)
        f.pack(fill=tk.BOTH, expand=True)

        config_df = self._calculate_config_data(all_filtered_df)

        if config_df.empty:
            ttk.Label(f, text="No data to display.").pack(padx=10, pady=10)
            return

        self._populate_treeview_tab(f, config_df)

    def get_coverage_text(self, classification):
        """Maps classification letter to full text."""
        if classification in ["P", "F"]:
            return "Covered"
        elif classification == "NC":
            return "Not Covered"
        else:
            return ""

    def get_current_classification(self, row):
        """Determines the effective classification (Override > Original Result)."""
        override = row.get("Override", pd.NA)
        if pd.notna(override):
            return override

        original_result = row.get("Result")
        if original_result == 'F':
            return 'F'
        if original_result == 'N':
            return 'NC'

        return original_result

    def _run_remarks_logic_update(self, parent_window):
        """Called by the Run button in the Remarks Config window."""
        if self.full_df is None:
            messagebox.showwarning("No Data", "Please load a data file before running the remarks logic.",
                                   parent=parent_window)
            return

        if self.current_view != "edited":
            messagebox.showwarning("Wrong View", "Please switch to the 'Edited file' view to make changes.",
                                   parent=parent_window)
            return

        if not messagebox.askyesno("Confirm Update",
                                   "This will apply the current Remark Configuration rules to the 'es_remarks' column in the edited file. Continue?",
                                   parent=parent_window):
            return

        self._set_buttons_state(tk.DISABLED)
        self.show_loading_screen("Applying remarks logic...")

        threading.Thread(
            target=self._thread_apply_remarks_logic,
            daemon=True
        ).start()

        # The result check will handle re-enabling buttons and refreshing treeview
        self.check_thread_queue()

    def _apply_remarks_logic(self):
        """Core logic to process and update the es_remarks column based on config."""

        # Compile rules for faster lookup: {special_name_upper: full_form}
        # Note: r[1] is the Full Form, as the Remarks column was removed.
        rules = {r[0].strip().upper(): r[1].strip()
                 for r in self.remarks_config_data if len(r) >= 2 and r[0].strip()}

        # Define delimiters for splitting the remark field
        delimiters_pattern = re.compile(r'[\s\/\,-]')

        def determine_es_remark(row):
            """Apply remark logic for a single row."""
            # Use original casing for 'Part' when constructing the output string
            part = str(row['Part']).strip()
            remark = str(row['Remark']).strip().upper()

            matched_remarks = []

            # Pre-split the remark into tokens for robust checking against multiple delimiters
            remark_tokens = [token for token in delimiters_pattern.split(remark) if token]

            # Iterate through all configured rules
            for special_name_upper, full_form in rules.items():

                if not special_name_upper:
                    continue

                # Rule 1: Check if the special name is the start of the part (Ref Des)
                # Case check is important here (part is original, special_name_upper is uppercase)
                if part.upper().startswith(special_name_upper):
                    # Output format: "C16 was not mounted"
                    matched_remarks.append(f"{part} was {full_form}")
                    continue

                # Rule 2: Check if the special name matches a token in the Remark
                if special_name_upper in remark_tokens:
                    # Output format: "C16 was not mounted"
                    matched_remarks.append(f"{part} was {full_form}")
                    continue

            # Use set to remove duplicates before joining, then sort for deterministic output
            unique_matches = sorted(list(set(matched_remarks)))
            return " | ".join(unique_matches).strip()

        # Apply the logic to the full DataFrame (updates the 'es_remarks' column)
        self.full_df['es_remarks'] = self.full_df.apply(determine_es_remark, axis=1)

    def _thread_apply_remarks_logic(self):
        """Worker thread to run the remarks logic."""
        try:
            self._apply_remarks_logic()
            self.thread_queue.put({
                'type': 'remarks_update',
                'status': 'success',
            })
        except Exception as e:
            self.thread_queue.put({
                'type': 'remarks_update',
                'status': 'error',
                'error': str(e)
            })

    def display_data(self, df):
        """Populates the main Treeview with the filtered data."""
        self.tree.delete(*self.tree.get_children())

        if df is None or df.empty:
            self.update_status_bar(0, 0)
            return

        # Include es_remarks in the displayed columns
        display_columns = self.trd_columns + ["Classification", "Parallel", "Coverage", "es_remarks"]

        df_cols = df.columns
        display_columns = [col for col in display_columns if col in df_cols]

        # Ensure minimal required columns exist for grouping/tagging
        if "Classification" not in df_cols: df["Classification"] = pd.NA
        if "Override" not in df_cols: df["Override"] = pd.NA
        if "MPin_Group" not in df_cols: df["MPin_Group"] = pd.NA
        if "es_remarks" not in df_cols: df["es_remarks"] = pd.NA

        self.tree["columns"] = display_columns

        self.tree.heading("#0", text="Component")
        self.tree.column("#0", width=150, anchor="w", stretch=tk.NO)

        # Dynamic column width calculation
        for col in display_columns:
            self.tree.heading(col, text=col)
            header_width = len(col) * 10 + 25

            data_width = 0
            if not df.empty:
                try:
                    # Estimate width based on data length
                    max_data_len = df[col].astype(str).str.len().max()
                    data_width = int((max_data_len or 0) * 8 + 20)
                except Exception:
                    pass

            col_width = max(100, header_width, data_width)
            self.tree.column(col, width=col_width, anchor="w", stretch=False)

        parent_items = {}

        # Sort by MPin_Group so parent rows are created correctly
        df_sorted = df.sort_values(by=["MPin_Group", "Part"], na_position='first')

        for index, row in df_sorted.iterrows():
            values = [row.get(col, "") for col in display_columns]

            classification_tag = str(row.get("Classification", "")).upper()

            if classification_tag not in ["P", "F", "NC"]:
                classification_tag = "default"
            tags_to_add = (classification_tag,)

            # Check for manual override
            override_val = row.get("Override", pd.NA)
            if pd.notna(override_val):
                if "Classification" in display_columns:
                    col_index = display_columns.index("Classification")
                    # Append (M) to visually indicate a manual override
                    values[col_index] = f"{values[col_index]} (M)"
                    tags_to_add = (classification_tag, "manual")

            mpin_group = row.get("MPin_Group")

            if pd.notna(mpin_group) and mpin_group.strip():
                # This is a child of a multi-pin component or a repeated 2-pin component group
                parent_iid = parent_items.get(mpin_group)

                if not parent_iid:
                    # Create the parent container row
                    parent_iid = f"parent_{mpin_group}"

                    # Prepare parent row values (mostly empty, except for key fields)
                    parent_values = [""] * len(display_columns)

                    try:
                        part_col_index = display_columns.index("Part")
                        parent_values[part_col_index] = mpin_group
                    except ValueError:
                        pass

                    try:
                        # Calculate summary info for the parent row dynamically
                        group_children_df = df_sorted[df_sorted["MPin_Group"] == mpin_group]
                        classifications = group_children_df["Classification"].unique()
                        group_count = len(group_children_df)

                        is_covered = any(c in ['P', 'F'] for c in classifications)
                        is_not_covered = 'NC' in classifications
                        parent_coverage_text = ""

                        if is_covered and is_not_covered:
                            parent_coverage_text = f"Covered, Not Covered ({group_count} items)"
                        elif is_covered:
                            parent_coverage_text = f"Covered ({group_count} items)"
                        elif is_not_covered:
                            parent_coverage_text = f"Not Covered ({group_count} items)"

                        if "Coverage" in display_columns:
                            coverage_col_index = display_columns.index("Coverage")
                            parent_values[coverage_col_index] = parent_coverage_text

                    except Exception:
                        pass

                    self.tree.insert("", "end", iid=parent_iid, text=mpin_group, values=parent_values,
                                     tags=("parent_row",), open=False)
                    parent_items[mpin_group] = parent_iid

                # Insert the child item
                self.tree.insert(parent_iid, "end", iid=index, text="", values=values, tags=tags_to_add)

            else:
                # This is a standalone 2-pin item
                self.tree.insert("", "end", iid=index, text="", values=values, tags=tags_to_add)

        # --- Status Bar Update ---
        source_df_in_board = self.full_df if self.current_view == "edited" else self.original_df
        if self.current_board_view != "All Boards" and source_df_in_board is not None:
            total_count = len(source_df_in_board[source_df_in_board["Board"] == self.current_board_view])
        elif source_df_in_board is not None:
            total_count = len(source_df_in_board)
        else:
            total_count = 0

        self.update_status_bar(len(df), total_count)
        # --- END ---

    def reset_tag_colors(self):
        """Configures the color tags for P, F, NC, Manual, and Parent rows based on the current theme."""
        if self.current_theme == "dark":
            self.tree.tag_configure("P", background="#004D00")
            self.tree.tag_configure("F", background="#6A0000")
            self.tree.tag_configure("NC", background="#6B5000")
            self.tree.tag_configure("default", background=self.colors["dark"]["bg_tree"],
                                    foreground=self.colors["dark"]["fg_tree"])
            self.tree.tag_configure("manual", foreground="#87CEFA")  # Light blue for manual override text
            self.tree.tag_configure("parent_row", background="#3E3E3E", foreground="#FFFFFF", font=("Arial", 9, "bold"))
        else:
            self.tree.tag_configure("P", background="lightgreen")
            self.tree.tag_configure("F", background="lightcoral")
            self.tree.tag_configure("NC", background="yellow")
            self.tree.tag_configure("default", background=self.colors["light"]["bg_tree"],
                                    foreground=self.colors["light"]["fg_tree"])
            self.tree.tag_configure("manual", foreground="#0000AA")  # Dark blue for manual override text
            self.tree.tag_configure("parent_row", background="#E0E0E0", foreground="#000000", font=("Arial", 9, "bold"))

    def highlight_rows(self, class_type):
        """Toggles the highlight/filter for a specific classification type (P, F, NC)."""
        if self.full_df is None:
            messagebox.showwarning("No Data", "Please open a file first.")
            return

        if self.current_highlight == class_type:
            self.current_highlight = None
        else:
            self.current_highlight = class_type

        self.refresh_treeview()

    def reset_highlight(self):
        """Resets all filters (Classification, Text, Pin Type, Group, MPin Group)."""
        self.current_highlight = None

        # Temporarily remove tracing to prevent multiple refresh_treeview calls
        try:
            self.style_filter.trace_remove("write", self.style_trace_id)
            self.part_filter.trace_remove("write", self.part_trace_id)
            self.remark_filter.trace_remove("write", self.remark_trace_id)
        except tk.TclError:
            pass

        self.style_filter.set("")
        self.part_filter.set("")
        self.remark_filter.set("")

        # Re-add tracing
        self.style_trace_id = self.style_filter.trace_add("write", self.on_text_filter_change)
        self.part_trace_id = self.part_filter.trace_add("write", self.on_text_filter_change)
        self.remark_trace_id = self.remark_filter.trace_add("write", self.on_text_filter_change)

        if self.board_selector['state'] != tk.DISABLED:
            self.current_board_view = "All Boards"
            self.board_selector.set("All Boards")

        self.current_pin_filter = "all"
        self.update_pin_filter_buttons()

        self.current_group_filter = "all"
        self.update_group_filter_buttons()

        self.current_mpin_group_filter = "All"
        if self.mpin_group_selector['state'] != tk.DISABLED:
            self.mpin_group_selector.set("All")

        # Reset global search query
        self.current_search_query = ""

        self.refresh_treeview()

    def on_text_filter_change(self, *args):
        """Triggers a refresh whenever a text filter entry changes."""
        self.refresh_treeview()

    def set_pin_filter(self, filter_type):
        """Sets the pin type filter and updates buttons."""
        self.current_pin_filter = filter_type
        self.update_pin_filter_buttons()
        self.refresh_treeview()

    def update_pin_filter_buttons(self):
        """Updates the visual state of the pin filter buttons."""
        if self.pin_filter_all_btn['state'] == tk.DISABLED:
            return

        # Ensure all buttons are configured correctly (selected/not selected state)
        self.pin_filter_all_btn.state(['selected'] if self.current_pin_filter == "all" else ['!selected'])
        self.pin_filter_2pin_btn.state(['selected'] if self.current_pin_filter == "2pin" else ['!selected'])
        self.pin_filter_mpin_btn.state(['selected'] if self.current_pin_filter == "mpin" else ['!selected'])

    def set_group_filter(self, filter_type):
        """Sets the grouped/single filter and updates buttons."""
        self.current_group_filter = filter_type
        self.update_group_filter_buttons()
        self.refresh_treeview()

    def update_group_filter_buttons(self):
        """Updates the visual state of the group filter buttons."""
        if self.group_filter_all_btn['state'] == tk.DISABLED:
            return

        # Ensure all buttons are configured correctly
        self.group_filter_all_btn.state(['selected'] if self.current_group_filter == "all" else ['!selected'])
        self.group_filter_single_btn.state(['selected'] if self.current_group_filter == "single" else ['!selected'])
        self.group_filter_grouped_btn.state(['selected'] if self.current_group_filter == "grouped" else ['!selected'])

    def refresh_treeview(self):
        """Applies all current filters and redraws the treeview."""

        source_df = self.full_df if self.current_view == "edited" else self.original_df

        if source_df is None:
            self.display_data(None)
            return

        df_to_show = source_df.copy()

        # Apply Board Filter
        if self.current_board_view != "All Boards":
            df_to_show = df_to_show[df_to_show["Board"] == self.current_board_view]

        # Apply Pin Type Filter
        if self.current_pin_filter == "2pin":
            df_to_show = df_to_show[df_to_show["Pin_Type"] == "2pin"]
        elif self.current_pin_filter == "mpin":
            df_to_show = df_to_show[df_to_show["Pin_Type"] == "mpin"]

        # Apply Group Filter
        if self.current_group_filter == "single":
            df_to_show = df_to_show[df_to_show["Is_Grouped"] == False]
        elif self.current_group_filter == "grouped":
            df_to_show = df_to_show[df_to_show["Is_Grouped"] == True]

        # Apply MPin Group Filter
        if self.current_mpin_group_filter != "All":
            df_to_show = df_to_show[df_to_show["MPin_Group"] == self.current_mpin_group_filter]

        # Apply Text Filters
        try:
            style_query = self.style_filter.get().upper()
            if style_query:
                df_to_show = df_to_show[df_to_show["Style"].astype(str).str.upper().str.startswith(style_query)]

            part_query = self.part_filter.get().upper()
            if part_query:
                df_to_show = df_to_show[df_to_show["Part"].astype(str).str.upper().str.startswith(part_query)]

            remark_query = self.remark_filter.get().upper()
            if remark_query:
                df_to_show = df_to_show[df_to_show["Remark"].astype(str).str.upper().str.startswith(remark_query)]
        except Exception as e:
            print(f"Error during text filtering: {e}")

        # --- Apply Ctrl+F Global Search Filter ---
        if self.current_search_query:
            query = self.current_search_query

            # Columns to search against (must handle NaNs by casting to string)
            search_cols = ['Part', 'Remark', 'Style', 'ID', 'MPin_Group']

            # Combine all search columns into one mask using OR (|) logic
            mask = False
            for col in search_cols:
                if col in df_to_show.columns:
                    # Use str.contains() for general search
                    mask |= df_to_show[col].astype(str).str.upper().str.contains(query, na=False)

            df_to_show = df_to_show[mask]
        # --- End Ctrl+F Filter ---

        # Apply Classification Highlight Filter
        if self.current_highlight:
            df_to_show = df_to_show[df_to_show["Classification"] == self.current_highlight]

        self.current_filtered_df = df_to_show

        self.display_data(df_to_show)

    def show_edited_file(self):
        """Switches to displaying the current (edited) dataframe."""
        self.current_view = "edited"
        self.view_menu.entryconfig("Show Edited File", state=tk.DISABLED)
        self.view_menu.entryconfig("Show Past File (Original)", state=tk.NORMAL)
        self.reload_logic_button.config(state=tk.NORMAL)
        self.refresh_treeview()

    def show_past_file(self):
        """Switches to displaying the original dataframe (read-only view)."""
        self.current_view = "past"
        self.view_menu.entryconfig("Show Edited File", state=tk.NORMAL)
        self.view_menu.entryconfig("Show Past File (Original)", state=tk.DISABLED)
        self.reload_logic_button.config(state=tk.DISABLED)
        self.refresh_treeview()

    def batch_update_classification(self, new_class_type):
        """Applies a manual override (P, F, NC, or Clear) to selected rows."""

        if self.current_view != "edited":
            messagebox.showwarning("Wrong View", "Please switch to the 'Edited file' view to make changes.")
            return

        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("No Selection", "Please select one or more rows to update.")
            return

        try:
            indices_to_update = []
            for item_id in selected_items:
                if "parent_" in str(item_id):
                    # If a parent is selected, update all its children
                    children = self.tree.get_children(item_id)
                    indices_to_update.extend([int(child_id) for child_id in children])
                else:
                    # If a child is selected, update it
                    try:
                        indices_to_update.append(int(item_id))
                    except ValueError:
                        pass  # Skip non-data rows (like the headers of grouped items if manually selected)

            unique_indices = list(set(indices_to_update))

            if not unique_indices:
                messagebox.showwarning("No Selection", "No valid data rows were selected.")
                return

            for row_index in unique_indices:
                if row_index not in self.full_df.index:
                    continue

                # Set the override value
                self.full_df.loc[row_index, "Override"] = new_class_type

                # Recalculate the effective classification and coverage
                new_classification = self.get_current_classification(self.full_df.loc[row_index])

                self.full_df.loc[row_index, "Classification"] = new_classification
                self.full_df.loc[row_index, "Coverage"] = self.get_coverage_text(new_classification)

            self.refresh_treeview()

            if pd.notna(new_class_type):
                messagebox.showinfo("Success", f"Updated {len(unique_indices)} rows to '{new_class_type}'.")
            else:
                messagebox.showinfo("Success", f"Cleared override on {len(unique_indices)} rows. Rows recalculated.")

        except Exception as e:
            messagebox.showerror("Error Updating", f"Could not update rows.\nError: {e}")

    def reload_original_logic(self):
        """Discards all manual overrides and restores Classification/Coverage from the original data."""
        if self.current_view != "edited":
            messagebox.showwarning("Wrong View", "Logic can only be reloaded in the 'Edited file' view.")
            return

        if self.full_df is None:
            messagebox.showwarning("No Data", "No data to reload.")
            return

        if not messagebox.askyesno("Confirm Reload",
                                   "This will discard ALL manual P/F/NC changes (overrides) and revert Classification/Coverage/es_remarks to the initial import state.\n\nThis action cannot be undone.\n\nAre you sure you want to continue?"):
            return

        try:
            if self.original_df is None:
                messagebox.showerror("Error", "Original data not found. Cannot reload.")
                return

            # Clear overrides and reset calculated columns based on the original data
            self.full_df["Override"] = pd.NA

            # Use original_df for all initial classification/remark columns
            self.full_df["Classification"] = self.original_df["Classification"]
            self.full_df["Coverage"] = self.original_df["Coverage"]
            self.full_df["Result"] = self.original_df["Result"]
            self.full_df["es_remarks"] = self.original_df["es_remarks"]

            self.refresh_treeview()

            messagebox.showinfo("Success",
                                "All manual overrides have been cleared and original logic has been reloaded.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while reloading logic:\n\n{e}")

    def _get_stats_from_df(self, df):
        """Calculates pin-level and component-level statistics from a given dataframe slice."""
        if df is None or df.empty:
            return 0, 0, 0, 0

        # --- Pin-level Counts ---
        classifications = df["Classification"].value_counts()
        count_p = classifications.get("P", 0)
        count_f = classifications.get("F", 0)
        count_nc = classifications.get("NC", 0)

        # --- Component-level Counts ---

        # 1. Start with Single/2-Pin (non-parallel) components
        qty_2pin_single = df[(df["Pin_Type"] == "2pin") & (df["Is_Grouped"] == False)]["Part"].nunique()

        # 2. Add Unique Multi-Pin Groups
        qty_mpin_groups = df[df["Pin_Type"] == "mpin"]["MPin_Group"].nunique()

        # 3. Add Unique Parallel Groups (if they aren't also multi-pin)
        # We assume the grouping by 'Parallel' field gives the unique component for parallel items
        qty_parallel_groups = df[(df["Is_Grouped"] == True) & (df["Pin_Type"] != "mpin")]["Parallel"].nunique()

        # Total Quantity of Unique Components (avoiding double counting between parallel splits and multi-pin pins)
        total_quantity = qty_2pin_single + qty_mpin_groups + qty_parallel_groups

        # Return Pin Counts (P, F, NC) and the unique Component Count
        return count_p, count_f, count_nc, total_quantity

    def calculate_stats(self):
        """Displays a detailed statistics window based on the current selection or view."""

        selected_items = self.tree.selection()
        base_df = self.full_df if self.current_view == "edited" else self.original_df

        if base_df is None:
            messagebox.showwarning("No Data", "No data to calculate statistics from.")
            return

        source_df = None
        data_source_text = ""

        # --- Determine Source Data (Selection vs. Filtered View) ---
        if selected_items:
            try:
                selected_indices = []
                for item_id in selected_items:
                    if "parent_" in str(item_id):
                        children = self.tree.get_children(item_id)
                        selected_indices.extend([int(child_id) for child_id in children])
                    else:
                        try:
                            selected_indices.append(int(item_id))
                        except ValueError:
                            pass

                unique_indices = list(set(selected_indices))

                if not unique_indices:
                    messagebox.showwarning("No Selection", "No valid data rows were selected.")
                    return

                # Filter the base DF using the selected indices
                # Check if all indices are valid before loc[]
                valid_indices = [idx for idx in unique_indices if idx in base_df.index]
                if not valid_indices:
                    messagebox.showwarning("No Data", "Selected rows do not correspond to valid data points.")
                    return

                data_source_text = f"{len(valid_indices)} selected rows"
                source_df = base_df.loc[valid_indices].copy()

            except Exception as e:
                messagebox.showerror("Stats Error", f"Could not read selected row data.\n\nError: {e}")
                return
        else:
            source_df = self.current_filtered_df
            if source_df is None or source_df.empty:
                messagebox.showwarning("No Data", "No data visible to calculate statistics from.")
                return
            data_source_text = f"{len(source_df)} visible rows (no selection)"

        if source_df is None or source_df.empty:
            messagebox.showwarning("No Data", "No data found in the current selection or view to calculate.")
            return

        try:
            # Recalculate stats for the source_df (handles filtering implicitly)
            count_p, count_f, count_nc, total_quantity = self._get_stats_from_df(source_df)

            result = ""
            total_rows = count_p + count_f + count_nc  # Total classified pins

            if total_rows == 0:
                result = "No P/F/NC data found"
            else:
                # Component-level status summary (simplified for the text box)
                has_covered = (count_p + count_f) > 0
                has_nc = count_nc > 0

                if has_covered and has_nc:
                    result = "Partially Covered (PC) or Mixed"
                elif has_covered:
                    result = "Covered (C)"
                elif has_nc:
                    result = "Not Covered (NC)"
                else:
                    result = "No P/F/NC data found"

            # --- Breakdown Statistics (using masks on the source_df) ---

            # Prepare breakdown DataFrames
            df_2pin_raw = source_df[source_df["Pin_Type"] == "2pin"]
            df_mpin_raw = source_df[source_df["Pin_Type"] == "mpin"]
            df_parallel_raw = source_df[source_df["Is_Grouped"] == True]

            # Calculate stats for the breakdowns (using pin counts and unique component counts)
            p_2pin, f_2pin, nc_2pin, qty_2pin = self._get_stats_from_df(df_2pin_raw)
            p_mpin, f_mpin, nc_mpin, qty_mpin = self._get_stats_from_df(df_mpin_raw)
            p_para, f_para, nc_para, qty_para = self._get_stats_from_df(df_parallel_raw)

            message = (
                f"Statistics for {data_source_text} (View: {self.current_view.capitalize()}):\n\n"
                f"--- Overall Pin Result ---\n"
                f"Result: {result}\n"
                f"Total Classified Pins: {total_rows}\n"
                f"Total P: {count_p} | Total F: {count_f} | Total NC: {count_nc}\n"
                f"Total Unique Components/Groups: {total_quantity}\n"
                f"----------------------------------------\n"
                f"--- 2-Pin Breakdown (by Pin Count) ---\n"
                f"P: {p_2pin} | F: {f_2pin} | NC: {nc_2pin}\n"
                f"Quantity (Unique 2-Pin Parts): {qty_2pin}\n"
                f"----------------------------------------\n"
                f"--- Multi-Pin Breakdown (by Pin Count) ---\n"
                f"P: {p_mpin} | F: {f_mpin} | NC: {nc_mpin}\n"
                f"Quantity (Unique MPin Groups): {qty_mpin}\n"
                f"----------------------------------------\n"
                f"--- Parallel (Grouped) Breakdown (by Pin Count) ---\n"
                f"P: {p_para} | F: {f_para} | NC: {nc_para}\n"
                f"Quantity (Unique Parallel Groups): {qty_para}\n"
            )

            self.show_stats_window(message)

        except Exception as e:
            messagebox.showerror("Stats Error", f"Could not calculate statistics.\n\nError: {e}")

    def show_stats_window(self, message):
        """Creates a modal Toplevel window to display the formatted statistics."""
        stats_win = tk.Toplevel(self.root)
        stats_win.title("Selection Statistics")
        stats_win.transient(self.root)
        stats_win.grab_set()

        # Apply theme to Toplevel frame
        c = self.colors[self.current_theme]
        frame = ttk.Frame(stats_win, padding=20, style="TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        label = ttk.Label(frame, text=message, justify=tk.LEFT, font=("Courier", 10), style="TLabel")
        label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(stats_win, style="TFrame")
        btn_frame.pack(pady=(0, 10))
        ok_btn = ttk.Button(btn_frame, text="OK", command=stats_win.destroy)
        ok_btn.pack()

        self.root.wait_window(stats_win)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelClassifierApp(root)
    root.mainloop()
