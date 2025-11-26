#!/usr/bin/env python3
"""
GUI application to browse Z21 locomotives and their details.
"""

import sys
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
from typing import Optional
import json
import re
import tempfile
import zipfile
import sqlite3
import uuid
import subprocess
import platform
import os

# Try to import PyObjC for macOS sharing
try:
    from AppKit import NSSharingService, NSURL, NSArray, NSWorkspace
    from Foundation import NSFileManager
    HAS_PYOBJC = True
except ImportError:
    HAS_PYOBJC = False

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from src.parser import Z21Parser
from src.data_models import Z21File, Locomotive, FunctionInfo


class Z21GUI:
    """Main GUI application for browsing Z21 locomotives."""

    def __init__(self, root, z21_file: Path):
        self.root = root
        self.z21_file = z21_file
        self.parser: Optional[Z21Parser] = None
        self.z21_data: Optional[Z21File] = None
        self.current_loco: Optional[Locomotive] = None
        self.current_loco_index: Optional[
            int] = None  # Index in z21_data.locomotives
        self.original_loco_address: Optional[
            int] = None  # Store original address for database lookup
        self.default_icon_path = Path(
            __file__).parent.parent / "icons" / "neutrals_normal.png"
        self.icon_cache = {}  # Cache for loaded icons
        self.icon_mapping = self.load_icon_mapping()  # Load icon mapping

        self.setup_ui()
        self.load_data()

    def load_icon_mapping(self):
        """Load icon mapping from JSON file."""
        mapping_file = Path(__file__).parent.parent / "icon_mapping.json"
        if mapping_file.exists():
            try:
                with open(mapping_file, 'r') as f:
                    data = json.load(f)
                    return data.get('matches', {})
            except Exception:
                return {}
        return {}

    def setup_ui(self):
        """Set up the user interface."""
        self.root.title("Z21 Locomotive Manager")
        self.root.geometry("1200x800")

        # Configure ttk styles for better visibility
        style = ttk.Style()
        # Configure Notebook tab colors for better visibility
        style.configure('TNotebook.Tab',
                        foreground='#000000',
                        background='#F0F0F0',
                        padding=[10, 5])
        style.map('TNotebook.Tab',
                  background=[('selected', '#E0E0E0')],
                  foreground=[('selected', '#000000')])

        # Create main paned window
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Left panel: Locomotive list
        left_frame = ttk.Frame(main_paned)
        main_paned.add(left_frame, weight=1)

        # Search box
        search_frame = ttk.Frame(left_frame)
        search_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.on_search)
        search_entry = ttk.Entry(search_frame,
                                 textvariable=self.search_var,
                                 width=20)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # New button to create a new locomotive
        new_button = ttk.Button(search_frame,
                                text="New",
                                command=self.create_new_locomotive)
        new_button.pack(side=tk.LEFT, padx=5)

        # Locomotive list
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        ttk.Label(list_frame, text="Locomotives:",
                  font=('Arial', 10, 'bold')).pack(anchor=tk.W)

        # Listbox with scrollbar
        listbox_frame = ttk.Frame(list_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.loco_listbox = tk.Listbox(listbox_frame,
                                       yscrollcommand=scrollbar.set,
                                       font=('Arial', 10))
        self.loco_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.loco_listbox.bind('<<ListboxSelect>>', self.on_loco_select)
        scrollbar.config(command=self.loco_listbox.yview)

        # Status label
        self.status_label = ttk.Label(left_frame,
                                      text="Loading...",
                                      relief=tk.SUNKEN)
        self.status_label.pack(fill=tk.X, padx=5, pady=5)

        # Right panel: Details
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=2)

        # Details notebook (tabs)
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Overview tab
        self.overview_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.overview_frame, text="Overview")
        self.setup_overview_tab()

        # Functions tab
        self.functions_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.functions_frame, text="Functions")
        self.setup_functions_tab()

    def setup_overview_tab(self):
        """Set up the overview tab."""
        # Create scrollable container for entire Overview tab content
        canvas = tk.Canvas(self.overview_frame, bg='#F0F0F0')
        scrollbar = ttk.Scrollbar(self.overview_frame,
                                  orient="vertical",
                                  command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#F0F0F0')

        def on_frame_configure(event):
            """Update scroll region when frame size changes."""
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(event):
            """Update frame width to match canvas."""
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)

        scrollable_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)

        canvas_window = canvas.create_window((0, 0),
                                             window=scrollable_frame,
                                             anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Store references for mouse wheel binding
        self.overview_canvas = canvas
        self.overview_scrollable_frame = scrollable_frame

        # Top frame for editable locomotive details
        details_frame = ttk.LabelFrame(scrollable_frame,
                                       text="Locomotive Details",
                                       padding=10)
        details_frame.pack(fill=tk.X, padx=5, pady=5)

        # Row 0: Image panel - width matches exactly from Name entry left to Address entry right
        self.loco_image_label = tk.Label(details_frame,
                                         text="No Image",
                                         bg='white',
                                         relief=tk.SUNKEN,
                                         anchor='center')
        # Match padding: Name entry has padx=(0, 19), Address entry has padx=(0, 10)
        # Image panel spans columns 1-3, matching the entry fields' width
        self.loco_image_label.grid(row=0,
                                   column=1,
                                   columnspan=3,
                                   padx=(0, 10),
                                   pady=5,
                                   sticky='ew')
        self.loco_image_label.image = None  # Keep a reference to prevent garbage collection

        # Row 1: Name and Address (two columns)
        ttk.Label(details_frame, text="Name:", width=10,
                  anchor='e').grid(row=1,
                                   column=0,
                                   padx=(5, 9),
                                   pady=2,
                                   sticky='e')
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(details_frame,
                                    textvariable=self.name_var,
                                    width=15)
        self.name_entry.grid(row=1,
                             column=1,
                             padx=(0, 19),
                             pady=2,
                             sticky='ew')

        ttk.Label(details_frame, text="Address:", width=10,
                  anchor='e').grid(row=1,
                                   column=2,
                                   padx=(0, 9),
                                   pady=2,
                                   sticky='e')
        self.address_var = tk.StringVar()
        self.address_entry = ttk.Entry(details_frame,
                                       textvariable=self.address_var,
                                       width=25)
        self.address_entry.grid(row=1,
                                column=3,
                                padx=(0, 10),
                                pady=2,
                                sticky='ew')

        # Row 2: Max Speed and Direction (two columns)
        ttk.Label(details_frame, text="Max Speed:", width=10,
                  anchor='e').grid(row=2,
                                   column=0,
                                   padx=(5, 9),
                                   pady=2,
                                   sticky='e')
        self.speed_var = tk.StringVar()
        self.speed_entry = ttk.Entry(details_frame,
                                     textvariable=self.speed_var,
                                     width=15)
        self.speed_entry.grid(row=2,
                              column=1,
                              padx=(0, 19),
                              pady=2,
                              sticky='ew')

        ttk.Label(details_frame, text="Direction:", width=10,
                  anchor='e').grid(row=2,
                                   column=2,
                                   padx=(0, 9),
                                   pady=2,
                                   sticky='e')
        self.direction_var = tk.StringVar()
        self.direction_combo = ttk.Combobox(details_frame,
                                            textvariable=self.direction_var,
                                            values=['Forward', 'Reverse'],
                                            state='readonly',
                                            width=25)
        self.direction_combo.grid(row=2,
                                  column=3,
                                  padx=(0, 10),
                                  pady=2,
                                  sticky='ew')

        # Additional Information Section
        row = 3
        ttk.Separator(details_frame, orient=tk.HORIZONTAL).grid(row=row,
                                                                column=0,
                                                                columnspan=6,
                                                                sticky='ew',
                                                                padx=5,
                                                                pady=5)
        row += 1

        # Note: Image column (4) spans rows 0-1, so additional fields start at row 3
        # Fields below separator fill the full width (columns 0-3, leaving column 4 for image)

        # Full Name field - spans full width (columns 0-5)
        ttk.Label(details_frame, text="Full Name:", width=10,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=(5, 1),
                                   pady=2,
                                   sticky='e')
        self.full_name_var = tk.StringVar()
        self.full_name_entry = ttk.Entry(details_frame,
                                         textvariable=self.full_name_var,
                                         width=60)
        self.full_name_entry.grid(row=row,
                                  column=1,
                                  columnspan=5,
                                  padx=(1, 5),
                                  pady=2,
                                  sticky='ew')
        row += 1

        # Two-column layout for remaining fields - fill entire width (columns 0-4)
        # Row 1: Railway and Article Number
        ttk.Label(details_frame, text="Railway:", width=10,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=(5, 1),
                                   pady=2,
                                   sticky='e')
        self.railway_var = tk.StringVar()
        self.railway_entry = ttk.Entry(details_frame,
                                       textvariable=self.railway_var,
                                       width=30)
        self.railway_entry.grid(row=row,
                                column=1,
                                padx=(1, 3),
                                pady=2,
                                sticky='ew')

        ttk.Label(details_frame, text="Article Number:", width=15,
                  anchor='e').grid(row=row,
                                   column=2,
                                   padx=(3, 1),
                                   pady=2,
                                   sticky='e')
        self.article_number_var = tk.StringVar()
        self.article_number_entry = ttk.Entry(
            details_frame, textvariable=self.article_number_var, width=15)
        self.article_number_entry.grid(row=row,
                                       column=3,
                                       padx=(1, 5),
                                       pady=2,
                                       sticky='ew')
        row += 1

        # Row 2: Decoder Type and Build Year
        ttk.Label(details_frame, text="Decoder Type:", width=10,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=(5, 1),
                                   pady=2,
                                   sticky='e')
        self.decoder_type_var = tk.StringVar()
        self.decoder_type_entry = ttk.Entry(details_frame,
                                            textvariable=self.decoder_type_var,
                                            width=30)
        self.decoder_type_entry.grid(row=row,
                                     column=1,
                                     padx=(1, 3),
                                     pady=2,
                                     sticky='ew')

        ttk.Label(details_frame, text="Build Year:", width=15,
                  anchor='e').grid(row=row,
                                   column=2,
                                   padx=(3, 1),
                                   pady=2,
                                   sticky='e')
        self.build_year_var = tk.StringVar()
        self.build_year_entry = ttk.Entry(details_frame,
                                          textvariable=self.build_year_var,
                                          width=15)
        self.build_year_entry.grid(row=row,
                                   column=3,
                                   padx=(1, 5),
                                   pady=2,
                                   sticky='ew')
        row += 1

        # Row 3: Model Buffer Length and Service Weight
        ttk.Label(details_frame, text="Buffer Length:", width=10,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=(5, 1),
                                   pady=2,
                                   sticky='e')
        self.model_buffer_length_var = tk.StringVar()
        self.model_buffer_length_entry = ttk.Entry(
            details_frame, textvariable=self.model_buffer_length_var, width=30)
        self.model_buffer_length_entry.grid(row=row,
                                            column=1,
                                            padx=(1, 3),
                                            pady=2,
                                            sticky='ew')

        ttk.Label(details_frame, text="Service Weight:", width=15,
                  anchor='e').grid(row=row,
                                   column=2,
                                   padx=(3, 1),
                                   pady=2,
                                   sticky='e')
        self.service_weight_var = tk.StringVar()
        self.service_weight_entry = ttk.Entry(
            details_frame, textvariable=self.service_weight_var, width=15)
        self.service_weight_entry.grid(row=row,
                                       column=3,
                                       padx=(1, 5),
                                       pady=2,
                                       sticky='ew')
        row += 1

        # Row 4: Model Weight and Minimum Radius
        ttk.Label(details_frame, text="Model Weight:", width=10,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=(5, 1),
                                   pady=2,
                                   sticky='e')
        self.model_weight_var = tk.StringVar()
        self.model_weight_entry = ttk.Entry(details_frame,
                                            textvariable=self.model_weight_var,
                                            width=30)
        self.model_weight_entry.grid(row=row,
                                     column=1,
                                     padx=(1, 3),
                                     pady=2,
                                     sticky='ew')

        ttk.Label(details_frame, text="Minimum Radius:", width=15,
                  anchor='e').grid(row=row,
                                   column=2,
                                   padx=(3, 1),
                                   pady=2,
                                   sticky='e')
        self.rmin_var = tk.StringVar()
        self.rmin_entry = ttk.Entry(details_frame,
                                    textvariable=self.rmin_var,
                                    width=15)
        self.rmin_entry.grid(row=row,
                             column=3,
                             padx=(1, 5),
                             pady=2,
                             sticky='ew')
        row += 1

        # Row 5: IP Address and Driver's Cab
        ttk.Label(details_frame, text="IP Address:", width=10,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=(5, 1),
                                   pady=2,
                                   sticky='e')
        self.ip_var = tk.StringVar()
        self.ip_entry = ttk.Entry(details_frame,
                                  textvariable=self.ip_var,
                                  width=30)
        self.ip_entry.grid(row=row, column=1, padx=(1, 3), pady=2, sticky='ew')

        ttk.Label(details_frame, text="Driver's Cab:", width=15,
                  anchor='e').grid(row=row,
                                   column=2,
                                   padx=(3, 1),
                                   pady=2,
                                   sticky='e')
        self.drivers_cab_var = tk.StringVar()
        self.drivers_cab_entry = ttk.Entry(details_frame,
                                           textvariable=self.drivers_cab_var,
                                           width=15)
        self.drivers_cab_entry.grid(row=row,
                                    column=3,
                                    padx=(1, 5),
                                    pady=2,
                                    sticky='ew')
        row += 1

        checkbox_frame = ttk.Frame(details_frame)
        checkbox_frame.grid(row=row, column=1, sticky='w', padx=(1, 3), pady=2)

        # Active Checkbox (直接在 checkbutton 里写 text)
        self.active_var = tk.BooleanVar()
        self.active_checkbox = ttk.Checkbutton(checkbox_frame,
                                               text="Active",
                                               variable=self.active_var)
        # side='left' 让它靠左排列
        self.active_checkbox.pack(side='left',
                                  padx=(0, 60))  # 增大两个checkbox之间的间距

        # Crane Checkbox
        self.crane_var = tk.BooleanVar()
        self.crane_checkbox = ttk.Checkbutton(checkbox_frame,
                                              text="Crane",
                                              variable=self.crane_var)
        # 紧接着 Active 排列
        self.crane_checkbox.pack(side='left')

        ttk.Label(details_frame, text="Speed Display:", width=15,
                  anchor='e').grid(row=row,
                                   column=2,
                                   padx=(3, 1),
                                   pady=2,
                                   sticky='e')
        self.speed_display_var = tk.StringVar()
        self.speed_display_combo = ttk.Combobox(
            details_frame,
            textvariable=self.speed_display_var,
            values=['km/h', 'Regulation Step', 'mph'],
            state='readonly',
            width=12)
        self.speed_display_combo.grid(row=row,
                                      column=3,
                                      padx=(1, 5),
                                      pady=2,
                                      sticky='ew')
        row += 1

        # Row 7: Vehicle Type and Reg Step (same row)
        ttk.Label(details_frame, text="Vehicle Type:", width=10,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=(5, 1),
                                   pady=2,
                                   sticky='e')
        self.rail_vehicle_type_var = tk.StringVar()
        self.rail_vehicle_type_combo = ttk.Combobox(
            details_frame,
            textvariable=self.rail_vehicle_type_var,
            values=['Loco', 'Wagon', 'Accessory'],
            state='readonly',
            width=12)
        self.rail_vehicle_type_combo.grid(row=row,
                                          column=1,
                                          padx=(1, 3),
                                          pady=2,
                                          sticky='ew')

        ttk.Label(details_frame, text="Reg Step:", width=10,
                  anchor='e').grid(row=row,
                                   column=2,
                                   padx=(3, 1),
                                   pady=2,
                                   sticky='e')
        self.regulation_step_var = tk.StringVar()
        self.regulation_step_combo = ttk.Combobox(
            details_frame,
            textvariable=self.regulation_step_var,
            values=['128', '28', '14'],
            state='readonly',
            width=12)
        self.regulation_step_combo.grid(row=row,
                                        column=3,
                                        padx=(1, 5),
                                        pady=2,
                                        sticky='ew')
        row += 1

        # Row 8: Categories and Have Since - two fields in one row
        ttk.Label(details_frame, text="Categories:", width=10,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=(5, 1),
                                   pady=2,
                                   sticky='e')
        self.categories_var = tk.StringVar()
        self.categories_entry = ttk.Entry(details_frame,
                                          textvariable=self.categories_var,
                                          width=30)
        self.categories_entry.grid(row=row,
                                   column=1,
                                   padx=(1, 3),
                                   pady=2,
                                   sticky='ew')

        ttk.Label(details_frame, text="In Stock Since:", width=10,
                  anchor='e').grid(row=row,
                                   column=2,
                                   padx=(3, 1),
                                   pady=2,
                                   sticky='e')
        self.in_stock_since_var = tk.StringVar()
        self.in_stock_since_entry = ttk.Entry(
            details_frame, textvariable=self.in_stock_since_var, width=15)
        self.in_stock_since_entry.grid(row=row,
                                       column=3,
                                       padx=(1, 5),
                                       pady=2,
                                       sticky='ew')
        row += 1

        # Description field (multiline) - spans full width (columns 0-5)
        ttk.Label(details_frame, text="Description:", width=10,
                  anchor='ne').grid(row=row,
                                    column=0,
                                    padx=(5, 1),
                                    pady=2,
                                    sticky='ne')
        self.description_text = scrolledtext.ScrolledText(details_frame,
                                                          wrap=tk.WORD,
                                                          width=60,
                                                          height=12,
                                                          font=('Arial', 11))
        self.description_text.grid(row=row,
                                   column=1,
                                   columnspan=5,
                                   padx=(1, 5),
                                   pady=2,
                                   sticky='ew')
        row += 1

        # Configure column weights for responsive layout
        # Column 0: fixed width for labels
        # Columns 1-4: expand equally for image panel (2/3 width centered)
        # Column 5: fixed width for spacing
        details_frame.grid_columnconfigure(0, weight=0)  # Label column fixed
        details_frame.grid_columnconfigure(1, weight=1)  # Image panel left
        details_frame.grid_columnconfigure(2,
                                           weight=1)  # Image panel center-left
        details_frame.grid_columnconfigure(
            3, weight=1)  # Image panel center-right
        details_frame.grid_columnconfigure(4, weight=1)  # Image panel right
        details_frame.grid_columnconfigure(5, weight=0)  # Right spacing fixed

        # Action buttons
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)

        self.export_button = ttk.Button(button_frame,
                                        text="Export Z21 Loco",
                                        command=self.export_z21_loco)
        self.export_button.pack(side=tk.LEFT, padx=5)
        
        self.share_button = ttk.Button(button_frame,
                                       text="Share with WIFI",
                                       command=self.share_with_airdrop)
        self.share_button.pack(side=tk.LEFT, padx=5)

        self.scan_button = ttk.Button(button_frame,
                                      text="Scan for Details",
                                      command=self.scan_for_details)
        self.scan_button.pack(side=tk.RIGHT, padx=5)

        self.save_button = ttk.Button(button_frame,
                                      text="Save Changes",
                                      command=self.save_locomotive_changes)
        self.save_button.pack(side=tk.RIGHT, padx=5)

        # Scrollable text area for function summary and CV values
        text_frame = ttk.Frame(scrollable_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.overview_text = scrolledtext.ScrolledText(text_frame,
                                                       wrap=tk.WORD,
                                                       font=('Courier', 10),
                                                       state=tk.DISABLED)
        self.overview_text.pack(fill=tk.BOTH, expand=True)

        # Add two-finger scrolling support for entire Overview tab
        def on_overview_mousewheel(event):
            """Handle mouse wheel scrolling for Overview tab (two-finger scroll on trackpad)."""
            # Check if we're in the overview tab (index 0)
            try:
                if self.notebook.index(self.notebook.select()) != 0:
                    return
            except:
                pass

            # Handle different platforms and event types
            scroll_amount = 0

            # macOS/Linux Button-4/5 (two-finger scroll)
            if event.num == 4:
                scroll_amount = -5
            elif event.num == 5:
                scroll_amount = 5
            # Windows/Linux with delta attribute
            elif hasattr(event, 'delta'):
                scroll_amount = -1 * (event.delta // 120)
                if scroll_amount == 0:
                    scroll_amount = -1 if event.delta > 0 else 1
            # macOS with deltaY attribute (newer tkinter)
            elif hasattr(event, 'deltaY'):
                scroll_amount = -1 * (event.deltaY // 120)
                if scroll_amount == 0:
                    scroll_amount = -1 if event.deltaY > 0 else 1

            if scroll_amount != 0:
                # Scroll the canvas (which contains all content)
                self.overview_canvas.yview_scroll(int(scroll_amount), "units")
                # Also scroll the text widget if it's focused (for nested scrolling)
                try:
                    if self.overview_text.winfo_containing(
                            event.x_root, event.y_root):
                        self.overview_text.yview_scroll(
                            int(scroll_amount), "units")
                except:
                    pass

            return "break"  # Prevent event propagation

        def bind_overview_mousewheel(widget):
            """Bind mouse wheel events to a widget and its children."""
            widget.bind("<MouseWheel>", on_overview_mousewheel, add='+')
            widget.bind("<Button-4>", on_overview_mousewheel, add='+')
            widget.bind("<Button-5>", on_overview_mousewheel, add='+')
            # Also bind to children
            for child in widget.winfo_children():
                try:
                    bind_overview_mousewheel(child)
                except:
                    pass

        # Bind mouse wheel events to canvas and all its contents
        canvas.bind("<MouseWheel>", on_overview_mousewheel, add='+')
        canvas.bind("<Button-4>", on_overview_mousewheel, add='+')
        canvas.bind("<Button-5>", on_overview_mousewheel, add='+')

        # Bind to scrollable frame and all its children
        scrollable_frame.bind("<MouseWheel>", on_overview_mousewheel, add='+')
        scrollable_frame.bind("<Button-4>", on_overview_mousewheel, add='+')
        scrollable_frame.bind("<Button-5>", on_overview_mousewheel, add='+')

        # Bind to overview_frame for comprehensive coverage
        self.overview_frame.bind("<MouseWheel>",
                                 on_overview_mousewheel,
                                 add='+')
        self.overview_frame.bind("<Button-4>", on_overview_mousewheel, add='+')
        self.overview_frame.bind("<Button-5>", on_overview_mousewheel, add='+')

        # Bind to notebook for overview tab (index 0)
        def overview_notebook_mousewheel(event):
            if self.notebook.index(self.notebook.select()) == 0:
                return on_overview_mousewheel(event)

        # Note: notebook bindings are handled in setup_functions_tab
        # but we can add additional bindings here if needed

        # Bind to root window for comprehensive trackpad support (macOS)
        root = self.root
        root.bind_all(
            "<MouseWheel>",
            lambda e: on_overview_mousewheel(e)
            if self.notebook.index(self.notebook.select()) == 0 else None,
            add='+')
        root.bind_all(
            "<Button-4>",
            lambda e: on_overview_mousewheel(e)
            if self.notebook.index(self.notebook.select()) == 0 else None,
            add='+')
        root.bind_all(
            "<Button-5>",
            lambda e: on_overview_mousewheel(e)
            if self.notebook.index(self.notebook.select()) == 0 else None,
            add='+')

    def setup_functions_tab(self):
        """Set up the functions tab."""
        # Scrollable frame for functions with grid layout
        canvas = tk.Canvas(self.functions_frame, bg='#F0F0F0')
        scrollbar = ttk.Scrollbar(self.functions_frame,
                                  orient="vertical",
                                  command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#F0F0F0')

        def on_frame_configure(event):
            """Update scroll region when frame size changes."""
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(event):
            """Update frame width to match canvas."""
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)

        def on_mousewheel(event):
            """Handle mouse wheel scrolling (two-finger scroll on trackpad)."""
            # Check if we're in the functions tab
            try:
                if self.notebook.index(self.notebook.select()) != 1:
                    return
            except:
                pass

            # Handle different platforms and event types
            scroll_amount = 0

            # macOS/Linux Button-4/5 (two-finger scroll)
            if event.num == 4:
                scroll_amount = -5
            elif event.num == 5:
                scroll_amount = 5
            # Windows/Linux with delta attribute
            elif hasattr(event, 'delta'):
                scroll_amount = -1 * (event.delta // 120)
                if scroll_amount == 0:
                    scroll_amount = -1 if event.delta > 0 else 1
            # macOS with deltaY attribute (newer tkinter)
            elif hasattr(event, 'deltaY'):
                scroll_amount = -1 * (event.deltaY // 120)
                if scroll_amount == 0:
                    scroll_amount = -1 if event.deltaY > 0 else 1

            if scroll_amount != 0:
                canvas.yview_scroll(int(scroll_amount), "units")

            return "break"  # Prevent event propagation

        def bind_mousewheel(widget):
            """Bind mouse wheel events to a widget."""
            widget.bind("<MouseWheel>", on_mousewheel)
            widget.bind("<Button-4>", on_mousewheel)
            widget.bind("<Button-5>", on_mousewheel)
            # Also try binding to children
            for child in widget.winfo_children():
                try:
                    bind_mousewheel(child)
                except:
                    pass

        scrollable_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)

        # Bind mouse wheel events for scrolling (two-finger scroll support)
        # Use add='+' to avoid overwriting existing bindings
        # Bind to canvas
        canvas.bind("<MouseWheel>", on_mousewheel, add='+')
        canvas.bind("<Button-4>", on_mousewheel, add='+')
        canvas.bind("<Button-5>", on_mousewheel, add='+')

        # Bind to scrollable frame
        scrollable_frame.bind("<MouseWheel>", on_mousewheel, add='+')
        scrollable_frame.bind("<Button-4>", on_mousewheel, add='+')
        scrollable_frame.bind("<Button-5>", on_mousewheel, add='+')

        # Bind to parent frame
        self.functions_frame.bind("<MouseWheel>", on_mousewheel, add='+')
        self.functions_frame.bind("<Button-4>", on_mousewheel, add='+')
        self.functions_frame.bind("<Button-5>", on_mousewheel, add='+')

        # Bind to notebook (only when functions tab is selected)
        def notebook_mousewheel(event):
            selected_tab = self.notebook.index(self.notebook.select())
            if selected_tab == 1:  # Functions tab
                return on_mousewheel(event)
            # Overview tab (index 0) is handled by its own bindings

        self.notebook.bind("<MouseWheel>", notebook_mousewheel, add='+')
        self.notebook.bind("<Button-4>", notebook_mousewheel, add='+')
        self.notebook.bind("<Button-5>", notebook_mousewheel, add='+')

        # Bind to root window for comprehensive trackpad support (macOS)
        # Note: Overview tab scrolling is handled by its own bindings
        root = self.root
        root.bind_all(
            "<MouseWheel>",
            lambda e: on_mousewheel(e)
            if self.notebook.index(self.notebook.select()) == 1 else None,
            add='+')
        root.bind_all(
            "<Button-4>",
            lambda e: on_mousewheel(e)
            if self.notebook.index(self.notebook.select()) == 1 else None,
            add='+')
        root.bind_all(
            "<Button-5>",
            lambda e: on_mousewheel(e)
            if self.notebook.index(self.notebook.select()) == 1 else None,
            add='+')

        canvas_window = canvas.create_window((0, 0),
                                             window=scrollable_frame,
                                             anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Enable focus for keyboard scrolling
        canvas.focus_set()

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.functions_frame_inner = scrollable_frame
        self.functions_canvas = canvas

        # Update bindings when frame is updated
        def update_bindings():
            """Update mouse wheel bindings for all widgets in scrollable frame."""
            bind_mousewheel(scrollable_frame)

        # Store update function for later use
        self.update_scroll_bindings = update_bindings

    def load_data(self):
        """Load Z21 file data."""
        self.status_label.config(text="Loading data...")
        self.root.update()

        try:
            self.parser = Z21Parser(self.z21_file)
            self.z21_data = self.parser.parse()

            self.populate_list()
            self.status_label.config(
                text=f"Loaded {len(self.z21_data.locomotives)} locomotives")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")
            self.status_label.config(text="Error loading file")

    def populate_list(self, filter_text: str = ""):
        """Populate the locomotive list."""
        if not self.z21_data:
            return

        self.loco_listbox.delete(0, tk.END)
        self.filtered_locos = []

        filter_lower = filter_text.lower()

        for loco in self.z21_data.locomotives:
            display_text = f"Address {loco.address:4d} - {loco.name}"

            if not filter_text or filter_lower in display_text.lower():
                self.loco_listbox.insert(tk.END, display_text)
                self.filtered_locos.append(loco)

    def on_search(self, *args):
        """Handle search text change."""
        filter_text = self.search_var.get()
        self.populate_list(filter_text)

    def on_loco_select(self, event):
        """Handle locomotive selection."""
        selection = self.loco_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        if index < len(self.filtered_locos):
            self.current_loco = self.filtered_locos[index]
            # Store original address for database lookup (in case user changes it)
            self.original_loco_address = self.current_loco.address
            # Find the locomotive index in z21_data.locomotives
            self.current_loco_index = None
            if self.z21_data:
                for i, loco in enumerate(self.z21_data.locomotives):
                    if loco.address == self.current_loco.address and loco.name == self.current_loco.name:
                        self.current_loco_index = i
                        break
            self.update_details()

    def create_new_locomotive(self):
        """Create a new locomotive with empty information."""
        if not self.z21_data:
            messagebox.showerror("Error", "No Z21 data loaded.")
            return

        # Find next available address
        used_addresses = {loco.address for loco in self.z21_data.locomotives}
        new_address = 1
        while new_address in used_addresses:
            new_address += 1
            if new_address > 9999:  # Safety limit
                messagebox.showerror(
                    "Error",
                    "Too many locomotives. Cannot find available address.")
                return

        # Create new locomotive with empty/default values
        new_loco = Locomotive()
        new_loco.address = new_address
        new_loco.name = f"New Locomotive {new_address}"
        new_loco.speed = 0
        new_loco.direction = True
        new_loco.functions = {}
        new_loco.function_details = {}
        new_loco.cvs = {}

        # Add to z21_data
        self.z21_data.locomotives.append(new_loco)
        self.current_loco_index = len(self.z21_data.locomotives) - 1

        # Update list and select the new locomotive
        self.populate_list(
            self.search_var.get() if hasattr(self, 'search_var') else "")

        # Find and select the new locomotive in the listbox
        for i in range(self.loco_listbox.size()):
            item_text = self.loco_listbox.get(i)
            if f"Address {new_address:4d}" in item_text:
                self.loco_listbox.selection_clear(0, tk.END)
                self.loco_listbox.selection_set(i)
                self.loco_listbox.see(i)
                break

        # Set as current locomotive and update details
        self.current_loco = new_loco
        self.original_loco_address = new_loco.address
        self.update_details()

        # Switch to overview tab
        self.notebook.select(0)

        # Focus on name field for easy editing
        self.root.after(100, lambda: self.name_entry.focus())

        messagebox.showinfo(
            "New Locomotive",
            f"Created new locomotive with address {new_address}.\n"
            f"You can now edit the details.")

    def update_details(self):
        """Update the details display."""
        if not self.current_loco:
            return

        self.update_overview()
        self.update_functions()

    def update_overview(self):
        """Update overview tab."""
        loco = self.current_loco

        # Update editable fields
        self.name_var.set(loco.name)
        self.address_var.set(str(loco.address))
        self.speed_var.set(str(loco.speed))
        self.direction_var.set('Forward' if loco.direction else 'Reverse')

        # Update additional fields
        self.full_name_var.set(loco.full_name)
        self.railway_var.set(loco.railway)
        self.article_number_var.set(loco.article_number)
        self.decoder_type_var.set(loco.decoder_type)
        self.build_year_var.set(loco.build_year)
        self.model_buffer_length_var.set(loco.model_buffer_length)
        self.service_weight_var.set(loco.service_weight)
        self.model_weight_var.set(loco.model_weight)
        self.rmin_var.set(loco.rmin)
        self.ip_var.set(loco.ip)
        self.drivers_cab_var.set(loco.drivers_cab)
        self.active_var.set(loco.active)
        # Speed Display: 0=km/h, 1=Regulation Step, 2=mph
        speed_display_map = {0: 'km/h', 1: 'Regulation Step', 2: 'mph'}
        self.speed_display_var.set(
            speed_display_map.get(loco.speed_display, 'km/h'))
        # Rail Vehicle Type: 0=Loco, 1=Wagon, 2=Accessory
        rail_type_map = {0: 'Loco', 1: 'Wagon', 2: 'Accessory'}
        self.rail_vehicle_type_var.set(
            rail_type_map.get(loco.rail_vehicle_type, 'Loco'))
        self.crane_var.set(loco.crane)
        # Regulation Step: 0=128, 1=28, 2=14
        regulation_step_map = {0: '128', 1: '28', 2: '14'}
        self.regulation_step_var.set(
            regulation_step_map.get(loco.regulation_step, '128'))
        # Categories: join list with comma
        self.categories_var.set(
            ', '.join(loco.categories) if loco.categories else '')

        # In Stock Since
        self.in_stock_since_var.set(getattr(loco, 'in_stock_since', '') or '')

        # Update description text
        self.description_text.delete(1.0, tk.END)
        self.description_text.insert(1.0, loco.description)

        # Load and display locomotive image (6cm wide x 2.5cm height ≈ 227px x 94px at 96 DPI)
        if loco.image_name:
            loco_image = self.load_locomotive_image(loco.image_name,
                                                    size=(227, 94))
            if loco_image:
                self.loco_image_label.config(image=loco_image, text='')
                self.loco_image_label.image = loco_image  # Keep a reference
            else:
                self.loco_image_label.config(image='',
                                             text=f'Image:\n{loco.image_name}')
                self.loco_image_label.image = None
        else:
            self.loco_image_label.config(image='', text='No Image')
            self.loco_image_label.image = None

        # Update scrollable text area with function summary and CV values
        self.overview_text.config(state=tk.NORMAL)
        self.overview_text.delete(1.0, tk.END)

        text = f"""
{'='*70}
FUNCTION SUMMARY
{'='*70}

Functions:         {len(loco.functions)} configured
Function Details:  {len(loco.function_details)} available

"""

        if loco.function_details:
            # List functions by function number order
            sorted_funcs = sorted(loco.function_details.items(),
                                  key=lambda x: x[1].function_number)

            text += "\n"
            for func_num, func_info in sorted_funcs:
                shortcut = f" [{func_info.shortcut}]" if func_info.shortcut else ""
                time_str = f" (time: {func_info.time}s)" if func_info.time != "0" else ""
                btn_type = func_info.button_type_name()
                text += f"  F{func_num:<3} - {func_info.image_name:<25} [{btn_type}] {shortcut}{time_str}\n"
            text += "\n"
        elif loco.functions:
            func_nums = sorted(loco.functions.keys())
            text += f"Function numbers: {', '.join(f'F{f}' for f in func_nums)}\n"

        if loco.cvs:
            text += f"\n{'='*70}\nCV VALUES\n{'='*70}\n"
            for cv_num, cv_value in sorted(loco.cvs.items()):
                text += f"CV{cv_num:3d} = {cv_value}\n"
        else:
            text += "\nNo CV values configured.\n"

        self.overview_text.insert(1.0, text)
        self.overview_text.config(state=tk.DISABLED)

    def scan_for_details(self):
        """Scan image or PDF for locomotive details and auto-fill fields."""
        if not self.current_loco:
            messagebox.showerror("Error", "Please select a locomotive first.")
            return

        # Open file dialog for image or PDF
        file_path = filedialog.askopenfilename(
            title="Select Image or PDF to Scan",
            filetypes=[("Image files",
                        "*.png *.jpg *.jpeg *.gif *.bmp *.tiff"),
                       ("PDF files", "*.pdf"), ("All files", "*.*")])

        if not file_path:
            return

        try:
            # Show progress
            self.status_label.config(text="Scanning document...")
            self.root.update()

            # Extract text using OCR
            extracted_text = self.extract_text_from_file(file_path)

            if not extracted_text:
                messagebox.showwarning(
                    "Warning", "No text could be extracted from the document.")
                self.status_label.config(
                    text=f"Loaded {len(self.z21_data.locomotives)} locomotives"
                )
                return

            # Parse and fill fields
            self.parse_and_fill_fields(extracted_text)

            messagebox.showinfo("Success",
                                "Details extracted and filled from document!")
            self.status_label.config(
                text=f"Loaded {len(self.z21_data.locomotives)} locomotives")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to scan document: {e}")
            self.status_label.config(
                text=f"Loaded {len(self.z21_data.locomotives)} locomotives")

    def extract_text_from_file(self, file_path: str) -> str:
        """Extract text from image or PDF using OCR."""
        file_path = Path(file_path)

        # Try to import OCR libraries
        try:
            import pytesseract
        except ImportError:
            messagebox.showerror(
                "Missing Dependency", "pytesseract is required for OCR.\n\n"
                "Install it with: pip install pytesseract\n"
                "Also install Tesseract OCR:\n"
                "  macOS: brew install tesseract\n"
                "  Linux: sudo apt-get install tesseract-ocr\n"
                "  Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki"
            )
            return ""

        try:
            if file_path.suffix.lower() == '.pdf':
                # Handle PDF files
                try:
                    from pdf2image import convert_from_path
                except ImportError:
                    messagebox.showerror(
                        "Missing Dependency",
                        "pdf2image is required for PDF processing.\n\n"
                        "Install it with: pip install pdf2image\n"
                        "Also install poppler:\n"
                        "  macOS: brew install poppler\n"
                        "  Linux: sudo apt-get install poppler-utils")
                    return ""

                # Convert PDF to images
                images = convert_from_path(str(file_path))
                # Extract text from all pages
                text_parts = []
                for image in images:
                    text = pytesseract.image_to_string(image)
                    text_parts.append(text)
                return "\n".join(text_parts)
            else:
                # Handle image files
                if not HAS_PIL:
                    messagebox.showerror(
                        "Error",
                        "PIL/Pillow is required for image processing.")
                    return ""

                image = Image.open(file_path)
                text = pytesseract.image_to_string(image)
                return text
        except Exception as e:
            raise Exception(f"OCR failed: {e}")

    def parse_and_fill_fields(self, text: str):
        """Parse extracted text and fill locomotive fields."""
        text = text.upper(
        )  # Convert to uppercase for case-insensitive matching

        # Extract Name (look for locomotive names like BR 103, BR 218, etc.)
        name_patterns = [
            r'\bBR\s*(\d+)\b',  # BR 103, BR 218
            r'\b(\d{4})\b',  # 4-digit numbers (could be locomotive numbers)
        ]
        if not self.name_var.get():
            for pattern in name_patterns:
                match = re.search(pattern, text)
                if match:
                    potential_name = match.group(0).strip()
                    if len(potential_name) <= 20:  # Reasonable name length
                        self.name_var.set(potential_name)
                        break

        # Extract Address (look for locomotive addresses)
        address_patterns = [
            r'\bADDRESS[:\s]+(\d+)\b',
            r'\bLOCO\s*ADDRESS[:\s]+(\d+)\b',
            r'\bADDR[:\s]+(\d+)\b',
        ]
        if not self.address_var.get():
            for pattern in address_patterns:
                match = re.search(pattern, text)
                if match:
                    self.address_var.set(match.group(1))
                    break

        # Extract Max Speed
        speed_patterns = [
            r'\bMAX\s*SPEED[:\s]+(\d+)\b',
            r'\bSPEED[:\s]+(\d+)\s*KM/H\b',
            r'\b(\d+)\s*KM/H\b',
            r'\bTOP\s*SPEED[:\s]+(\d+)\b',
        ]
        if not self.speed_var.get():
            for pattern in speed_patterns:
                match = re.search(pattern, text)
                if match:
                    speed = int(match.group(1))
                    if 0 < speed <= 300:  # Reasonable speed range
                        self.speed_var.set(str(speed))
                        break

        # Extract Railway/Company
        railway_patterns = [
            r'\bRAILWAY[:\s]+([A-Z][A-Z\s\.]+)\b',
            r'\bCOMPANY[:\s]+([A-Z][A-Z\s\.]+)\b',
            r'\b(K\.BAY\.STS\.B\.)\b',
            r'\b(DB|DR|SNCF|ÖBB)\b',
        ]
        if not self.railway_var.get():
            for pattern in railway_patterns:
                match = re.search(pattern, text)
                if match:
                    railway = match.group(1).strip()
                    if len(railway) <= 50:
                        self.railway_var.set(railway)
                        break

        # Extract Article Number
        article_patterns = [
            r'\bARTICLE[:\s]+(\d+)\b',
            r'\bART\.\s*NO[:\s]+(\d+)\b',
            r'\bPRODUCT[:\s]+(\d+)\b',
            r'\bITEM[:\s]+(\d+)\b',
        ]
        if not self.article_number_var.get():
            for pattern in article_patterns:
                match = re.search(pattern, text)
                if match:
                    self.article_number_var.set(match.group(1))
                    break

        # Extract Decoder Type
        decoder_patterns = [
            r'\bDECODER[:\s]+([A-Z0-9\s]+)\b',
            r'\b(NEM\s*\d+)\b',
            r'\b(DCC\s*DECODER)\b',
        ]
        if not self.decoder_type_var.get():
            for pattern in decoder_patterns:
                match = re.search(pattern, text)
                if match:
                    decoder = match.group(1).strip()
                    if len(decoder) <= 30:
                        self.decoder_type_var.set(decoder)
                        break

        # Extract Build Year
        year_patterns = [
            r'\bBUILD\s*YEAR[:\s]+(\d{4})\b',
            r'\bYEAR[:\s]+(\d{4})\b',
            r'\b(\d{4})\s*BUILD\b',
        ]
        if not self.build_year_var.get():
            for pattern in year_patterns:
                match = re.search(pattern, text)
                if match:
                    year = match.group(1)
                    if 1900 <= int(year) <= 2100:  # Reasonable year range
                        self.build_year_var.set(year)
                        break

        # Extract Weight
        weight_patterns = [
            r'\bWEIGHT[:\s]+(\d+(?:[.,]\d+)?)\s*(?:KG|G|T)?\b',
            r'\bSERVICE\s*WEIGHT[:\s]+(\d+(?:[.,]\d+)?)\b',
        ]
        if not self.service_weight_var.get():
            for pattern in weight_patterns:
                match = re.search(pattern, text)
                if match:
                    weight = match.group(1).replace(',', '.')
                    self.service_weight_var.set(weight)
                    break

        # Extract Minimum Radius
        radius_patterns = [
            r'\bMIN(?:IMUM)?\s*RADIUS[:\s]+(\d+(?:[.,]\d+)?)\b',
            r'\bRMIN[:\s]+(\d+(?:[.,]\d+)?)\b',
            r'\bRADIUS[:\s]+(\d+(?:[.,]\d+)?)\s*MM\b',
        ]
        if not self.rmin_var.get():
            for pattern in radius_patterns:
                match = re.search(pattern, text)
                if match:
                    radius = match.group(1).replace(',', '.')
                    self.rmin_var.set(radius)
                    break

        # Extract IP Address
        ip_pattern = r'\b(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b'
        if not self.ip_var.get():
            match = re.search(ip_pattern, text)
            if match:
                self.ip_var.set(match.group(1))

        # Extract Full Name (look for longer descriptive names)
        if not self.full_name_var.get():
            # Look for lines that might be full names (longer text, often at start)
            lines = text.split('\n')
            for line in lines[:10]:  # Check first 10 lines
                line = line.strip()
                if 20 <= len(line) <= 200 and not re.match(r'^\d+$', line):
                    # Check if it looks like a locomotive description
                    if any(keyword in line for keyword in
                           ['LOCOMOTIVE', 'LOCO', 'TRAIN', 'SET', 'MODEL']):
                        self.full_name_var.set(line)
                        break

        # Extract Description (collect longer text blocks)
        if not self.description_text.get(1.0, tk.END).strip():
            # Collect paragraphs that might be descriptions
            paragraphs = []
            for line in text.split('\n'):
                line = line.strip()
                if len(line) > 50:  # Longer lines are likely descriptions
                    paragraphs.append(line)

            if paragraphs:
                description = '\n\n'.join(
                    paragraphs[:5])  # Take first 5 paragraphs
                if len(description) > 100:  # Only if substantial content
                    self.description_text.delete(1.0, tk.END)
                    self.description_text.insert(
                        1.0, description[:2000])  # Limit to 2000 chars

    def scan_for_functions(self):
        """Scan image for function numbers and names, then auto-add functions."""
        if not self.current_loco:
            messagebox.showerror("Error", "Please select a locomotive first.")
            return

        # Open file dialog for image
        file_path = filedialog.askopenfilename(
            title="Select Image to Scan for Functions",
            filetypes=[("Image files",
                        "*.png *.jpg *.jpeg *.gif *.bmp *.tiff"),
                       ("All files", "*.*")])

        if not file_path:
            return

        try:
            # Show progress
            self.status_label.config(text="Scanning image for functions...")
            self.root.update()

            # Try AI-based extraction first, fallback to OCR
            functions = None
            try:
                functions = self.extract_functions_with_ai(file_path)
            except Exception as ai_error:
                # Fallback to OCR if AI fails
                self.status_label.config(
                    text="AI extraction failed, trying OCR...")
                self.root.update()
                try:
                    extracted_text = self.extract_text_from_file(file_path)
                    if extracted_text:
                        functions = self.parse_functions_from_text(
                            extracted_text)
                    else:
                        raise Exception(
                            f"AI extraction failed: {ai_error}. OCR also failed - no text extracted."
                        )
                except Exception as ocr_error:
                    raise Exception(
                        f"AI extraction failed: {ai_error}. OCR also failed: {ocr_error}"
                    )

            if not functions:
                messagebox.showwarning(
                    "Warning",
                    "No functions could be extracted from the image.")
                self.status_label.config(
                    text=f"Loaded {len(self.z21_data.locomotives)} locomotives"
                )
                return

            # Add functions to locomotive
            added_count = 0
            for func_num, func_name in functions:
                if func_num not in self.current_loco.function_details:
                    # Generate shortcut
                    shortcut = self.generate_shortcut(func_name)

                    # Match icon
                    icon_name = self.match_function_to_icon(func_name)

                    # Create FunctionInfo
                    func_info = FunctionInfo(
                        function_number=func_num,
                        image_name=icon_name,
                        shortcut=shortcut,
                        position=0,
                        time="0",
                        button_type=0,  # Default to switch
                        is_active=True)

                    # Add to locomotive
                    self.current_loco.function_details[func_num] = func_info
                    self.current_loco.functions[func_num] = True
                    added_count += 1

            if added_count > 0:
                messagebox.showinfo(
                    "Success",
                    f"Added {added_count} function(s) from scanned image!\n\n"
                    f"Functions added: {', '.join([f'F{num}' for num, _ in functions if num not in self.current_loco.function_details])}"
                )

                # Update functions display
                self.update_functions()
            else:
                messagebox.showinfo(
                    "Info", "All functions from the image already exist.")

            self.status_label.config(
                text=f"Loaded {len(self.z21_data.locomotives)} locomotives")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to scan functions: {e}")
            self.status_label.config(
                text=f"Loaded {len(self.z21_data.locomotives)} locomotives")

    def extract_functions_with_ai(self, image_path: str) -> list:
        """Extract function information from image using AI (OpenAI GPT-4 Vision).
        Returns list of tuples: [(function_number, function_name), ...]
        """
        try:
            import openai
        except ImportError:
            raise ImportError(
                "openai package is required for AI extraction.\n\n"
                "Install it with: pip install openai\n\n"
                "You also need to set your OpenAI API key:\n"
                "  - Set environment variable: export OPENAI_API_KEY='your-key'\n"
                "  - Or create a config file: ~/.openai/config.json")

        # Check for API key
        import os
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            # Try to read from config file
            config_path = Path.home() / '.openai' / 'config.json'
            if config_path.exists():
                try:
                    with open(config_path, 'r') as f:
                        config = json.load(f)
                        api_key = config.get('api_key')
                except:
                    pass

        if not api_key:
            raise ValueError("OpenAI API key not found.\n\n"
                             "Please set your API key:\n"
                             "  export OPENAI_API_KEY='your-key-here'\n\n"
                             "Or create ~/.openai/config.json with:\n"
                             '  {"api_key": "your-key-here"}')

        # Initialize OpenAI client
        client = openai.OpenAI(api_key=api_key)

        # Read image file
        if not HAS_PIL:
            raise ImportError("PIL/Pillow is required for image processing.")

        image = Image.open(image_path)

        # Convert to base64 for API
        import base64
        import io

        buffered = io.BytesIO()
        # Convert to RGB if necessary (for PNG with transparency)
        if image.mode in ('RGBA', 'LA', 'P'):
            rgb_image = Image.new('RGB', image.size, (255, 255, 255))
            if image.mode == 'P':
                image = image.convert('RGBA')
            rgb_image.paste(
                image,
                mask=image.split()[-1] if image.mode == 'RGBA' else None)
            image = rgb_image

        image.save(buffered, format="PNG")
        image_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')

        # Prepare prompt for AI
        prompt = """Analyze this image of a locomotive function list or manual page. 
Extract all function numbers and their corresponding names/descriptions.

Look for patterns like:
- F0: Light
- F1: Horn
- Function 0: Light
- F0 Light
- etc.

Return ONLY a JSON array of objects, each with "number" and "name" fields.
Example format:
[
  {"number": 0, "name": "Light"},
  {"number": 1, "name": "Horn"},
  {"number": 2, "name": "Bell"}
]

If you cannot find any functions, return an empty array [].
Be accurate and extract all visible functions."""

        # Call OpenAI Vision API
        self.status_label.config(
            text="Calling AI model to extract functions...")
        self.root.update()

        response = client.chat.completions.create(
            model="gpt-4o",  # or "gpt-4-vision-preview" for older models
            messages=[{
                "role":
                "user",
                "content": [{
                    "type": "text",
                    "text": prompt
                }, {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{image_base64}"
                    }
                }]
            }],
            max_tokens=1000)

        # Parse response
        response_text = response.choices[0].message.content.strip()

        # Try to extract JSON from response (might have markdown code blocks)
        import json
        json_match = re.search(r'\[.*\]', response_text, re.DOTALL)
        if json_match:
            response_text = json_match.group(0)

        try:
            functions_data = json.loads(response_text)
            functions = []
            for item in functions_data:
                if isinstance(item,
                              dict) and 'number' in item and 'name' in item:
                    func_num = int(item['number'])
                    func_name = str(item['name']).strip()
                    if 0 <= func_num <= 128 and func_name:
                        functions.append((func_num, func_name))
            return functions
        except json.JSONDecodeError as e:
            # Fallback: try to parse text response
            raise Exception(
                f"Failed to parse AI response as JSON: {e}\nResponse: {response_text}"
            )

    def parse_functions_from_text(self, text: str) -> list:
        """Parse function numbers and names from OCR text.
        Returns list of tuples: [(function_number, function_name), ...]
        """
        functions = []
        text_lines = text.split('\n')

        # Common patterns for function listings:
        # F0: Light, F1: Horn, F0 Light, F1 Horn, Function 0: Light, etc.
        patterns = [
            r'\bF(\d+)[:\s]+([A-Z][A-Z\s]+?)(?=\s+F\d+|\s*$|\n)',  # F0: Light
            r'\bF(\d+)\s+([A-Z][A-Z\s]+?)(?=\s+F\d+|\s*$|\n)',  # F0 Light
            r'\bFUNCTION\s+(\d+)[:\s]+([A-Z][A-Z\s]+?)(?=\s+FUNCTION|\s*$|\n)',  # Function 0: Light
            r'\bF(\d+)[:\s]+([A-Za-z][A-Za-z\s]+?)(?=\s+F\d+|\s*$|\n)',  # F0: light (lowercase)
        ]

        for line in text_lines:
            line = line.strip()
            if not line:
                continue

            # Try each pattern
            for pattern in patterns:
                matches = re.finditer(pattern, line, re.IGNORECASE)
                for match in matches:
                    func_num = int(match.group(1))
                    func_name = match.group(2).strip()

                    # Clean up function name
                    func_name = re.sub(r'\s+', ' ',
                                       func_name)  # Multiple spaces to single
                    func_name = func_name.strip()

                    if func_name and 0 <= func_num <= 128:
                        functions.append((func_num, func_name))

        # Also try to find function numbers followed by names on separate lines
        # Look for patterns like: "0" on one line, "Light" on next
        for i in range(len(text_lines) - 1):
            line1 = text_lines[i].strip()
            line2 = text_lines[i + 1].strip()

            # Check if line1 is just a number and line2 starts with a letter
            if re.match(r'^\d+$', line1) and re.match(r'^[A-Za-z]', line2):
                func_num = int(line1)
                func_name = line2.split()[0] if line2.split() else line2
                if 0 <= func_num <= 128:
                    functions.append((func_num, func_name))

        # Remove duplicates, keeping first occurrence
        seen = set()
        unique_functions = []
        for func_num, func_name in functions:
            if func_num not in seen:
                seen.add(func_num)
                unique_functions.append((func_num, func_name))

        return unique_functions

    def generate_shortcut(self, func_name: str) -> str:
        """Generate a keyboard shortcut for a function name."""
        func_name_lower = func_name.lower().strip()

        # Common shortcuts mapping
        shortcut_map = {
            'light': 'L',
            'horn': 'H',
            'bell': 'B',
            'whistle': 'W',
            'sound': 'S',
            'steam': 'S',
            'brake': 'B',
            'couple': 'C',
            'decouple': 'D',
            'door': 'D',
            'fan': 'F',
            'pump': 'P',
            'valve': 'V',
            'generator': 'G',
            'compressor': 'C',
            'neutral': 'N',
            'forward': 'F',
            'backward': 'B',
            'interior': 'I',
            'cabin': 'C',
            'cockpit': 'C',
        }

        # Check for exact matches first
        for key, shortcut in shortcut_map.items():
            if key in func_name_lower:
                return shortcut

        # Use first letter if no match
        if func_name_lower:
            first_char = func_name_lower[0]
            if first_char.isalpha():
                return first_char.upper()

        return ''

    def match_function_to_icon(self, func_name: str) -> str:
        """Match a function name to an icon using fuzzy matching."""
        func_name_lower = func_name.lower().strip()

        # Load icon mapping
        icon_names = list(self.icon_mapping.keys())

        # Direct keyword matching
        keyword_map = {
            'light': [
                'light', 'lamp', 'beam', 'sidelight', 'interior_light',
                'cabin_light'
            ],
            'horn': ['horn', 'horn_high', 'horn_low', 'horn_two_sound'],
            'bell': ['bell'],
            'whistle': ['whistle', 'whistle_long', 'whistle_short'],
            'sound': ['sound', 'sound1', 'sound2', 'sound3', 'sound4'],
            'steam': ['steam', 'dump_steam'],
            'brake': ['brake', 'brake_delay', 'sound_brake', 'handbrake'],
            'couple': ['couple'],
            'decouple': ['decouple'],
            'door': ['door', 'door_open', 'door_close'],
            'fan': ['fan', 'fan_strong', 'blower'],
            'pump': ['pump', 'feed_pump', 'air_pump'],
            'valve': ['valve', 'drain_valve'],
            'generator': ['generator', 'diesel_generator'],
            'compressor': ['compressor'],
            'neutral': ['neutral'],
            'forward': ['forward', 'forward_take_power'],
            'backward': ['backward', 'backward_take_power'],
            'interior': ['interior_light'],
            'cabin': ['cabin_light'],
            'cockpit': ['cockpit_light_left', 'cockpit_light_right'],
            'drain': ['drain', 'drainage', 'drain_mud', 'drain_valve'],
            'diesel': ['diesel', 'diesel_generator', 'diesel_regulation'],
            'rail': ['rail', 'rail_kick', 'rail_crossing'],
            'scoop': ['scoop', 'scoop_coal'],
            'firebox': ['firebox'],
            'injector': ['injector'],
            'preheat': ['preheat'],
            'mute': ['mute'],
            'louder': ['louder'],
            'quiter': ['quiter'],
        }

        # Try keyword matching
        for keyword, icon_candidates in keyword_map.items():
            if keyword in func_name_lower:
                # Find best match from candidates
                for candidate in icon_candidates:
                    if candidate in icon_names:
                        return candidate
                # If no exact match, try partial match
                for icon_name in icon_names:
                    for candidate in icon_candidates:
                        if candidate in icon_name:
                            return icon_name

        # Fuzzy matching: find icon names that contain words from function name
        func_words = set(re.findall(r'\b\w+\b', func_name_lower))
        best_match = None
        best_score = 0

        for icon_name in icon_names:
            icon_words = set(re.findall(r'\b\w+\b', icon_name.lower()))
            # Calculate overlap score
            overlap = len(func_words & icon_words)
            if overlap > best_score:
                best_score = overlap
                best_match = icon_name

        if best_match and best_score > 0:
            return best_match

        # Fallback: return first available icon or empty string
        return icon_names[0] if icon_names else ''

    def save_locomotive_changes(self):
        """Save changes to locomotive details."""
        if not self.current_loco or not self.z21_data or not self.parser:
            messagebox.showerror("Error",
                                 "No locomotive selected or data not loaded.")
            return

        try:
            # Update locomotive object with new values
            new_name = self.name_var.get()
            new_address = int(self.address_var.get())
            new_speed = int(self.speed_var.get())
            new_direction = (self.direction_var.get() == 'Forward')

            # Get additional field values
            new_full_name = self.full_name_var.get()
            new_railway = self.railway_var.get()
            new_article_number = self.article_number_var.get()
            new_decoder_type = self.decoder_type_var.get()
            new_build_year = self.build_year_var.get()
            new_model_buffer_length = self.model_buffer_length_var.get()
            new_service_weight = self.service_weight_var.get()
            new_model_weight = self.model_weight_var.get()
            new_rmin = self.rmin_var.get()
            new_ip = self.ip_var.get()
            new_drivers_cab = self.drivers_cab_var.get()
            new_description = self.description_text.get(1.0, tk.END).strip()
            new_active = self.active_var.get()
            # Speed Display: convert from text to int (0=km/h, 1=Regulation Step, 2=mph)
            speed_display_text = self.speed_display_var.get()
            speed_display_map = {'km/h': 0, 'Regulation Step': 1, 'mph': 2}
            new_speed_display = speed_display_map.get(speed_display_text, 0)
            # Rail Vehicle Type: convert from text to int (0=Loco, 1=Wagon, 2=Accessory)
            rail_type_text = self.rail_vehicle_type_var.get()
            rail_type_map = {'Loco': 0, 'Wagon': 1, 'Accessory': 2}
            new_rail_vehicle_type = rail_type_map.get(rail_type_text, 0)
            new_crane = self.crane_var.get()
            # Regulation Step: convert from display value to index (128=0, 28=1, 14=2)
            regulation_step_text = self.regulation_step_var.get()
            regulation_step_map = {'128': 0, '28': 1, '14': 2}
            new_regulation_step = regulation_step_map.get(
                regulation_step_text, 0)
            # Parse categories from comma-separated string
            categories_str = self.categories_var.get().strip()
            new_categories = [
                cat.strip() for cat in categories_str.split(',')
                if cat.strip()
            ] if categories_str else []

            # In Stock Since
            new_in_stock_since = self.in_stock_since_var.get().strip()

            # Update the locomotive in z21_data
            if self.current_loco_index is not None:
                loco = self.z21_data.locomotives[self.current_loco_index]
                loco.name = new_name
                loco.address = new_address
                loco.speed = new_speed
                loco.direction = new_direction

                # Update additional fields
                loco.full_name = new_full_name
                loco.railway = new_railway
                loco.article_number = new_article_number
                loco.decoder_type = new_decoder_type
                loco.build_year = new_build_year
                loco.model_buffer_length = new_model_buffer_length
                loco.service_weight = new_service_weight
                loco.model_weight = new_model_weight
                loco.rmin = new_rmin
                loco.ip = new_ip
                loco.drivers_cab = new_drivers_cab
                loco.description = new_description
                loco.active = new_active
                loco.speed_display = new_speed_display
                loco.rail_vehicle_type = new_rail_vehicle_type
                loco.crane = new_crane
                loco.regulation_step = new_regulation_step
                loco.categories = new_categories

                # Also update current_loco reference
                self.current_loco.name = new_name
                self.current_loco.address = new_address
                self.current_loco.speed = new_speed
                self.current_loco.direction = new_direction
                self.current_loco.full_name = new_full_name
                self.current_loco.railway = new_railway
                self.current_loco.article_number = new_article_number
                self.current_loco.decoder_type = new_decoder_type
                self.current_loco.build_year = new_build_year
                self.current_loco.model_buffer_length = new_model_buffer_length
                self.current_loco.service_weight = new_service_weight
                self.current_loco.model_weight = new_model_weight
                self.current_loco.rmin = new_rmin
                self.current_loco.ip = new_ip
                self.current_loco.drivers_cab = new_drivers_cab
                self.current_loco.description = new_description
                self.current_loco.active = new_active
                self.current_loco.speed_display = new_speed_display
                self.current_loco.rail_vehicle_type = new_rail_vehicle_type
                self.current_loco.crane = new_crane
                self.current_loco.regulation_step = new_regulation_step
                self.current_loco.categories = new_categories
                self.current_loco.in_stock_since = new_in_stock_since
            else:
                messagebox.showerror(
                    "Error", "Could not find locomotive in data structure.")
                return

            # Write changes back to file
            try:
                self.parser.write(self.z21_data, self.z21_file)
                messagebox.showinfo(
                    "Success",
                    "Locomotive details saved successfully to file!")
            except Exception as write_error:
                messagebox.showerror(
                    "Write Error",
                    f"Failed to write changes to file: {write_error}\n\n"
                    f"Changes have been saved in memory but not written to disk."
                )

            # Update the listbox to reflect name change
            self.populate_list(
                self.search_var.get() if hasattr(self, 'search_var') else "")

        except ValueError as e:
            messagebox.showerror(
                "Error",
                f"Invalid input: {e}\n\nPlease enter valid numbers for Address and Max Speed."
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save changes: {e}")

    def export_z21_loco(self):
        """Export current locomotive to z21_loco.z21loco format."""
        if not self.current_loco or not self.z21_data or not self.parser:
            messagebox.showerror("Error",
                                 "No locomotive selected or data not loaded.")
            return

        try:
            import uuid
            import shutil

            # Ask user for output file path
            output_file = filedialog.asksaveasfilename(
                title="Export Z21 Loco",
                defaultextension=".z21loco",
                filetypes=[("Z21 Loco files", "*.z21loco"),
                           ("All files", "*.*")])

            if not output_file:
                return  # User cancelled

            output_path = Path(output_file)

            # Generate UUID for export directory
            export_uuid = str(uuid.uuid4()).upper()
            export_dir = f"export/{export_uuid}"

            # Create temporary directory for export
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                export_path = temp_path / export_dir
                export_path.mkdir(parents=True, exist_ok=True)

                # Get original SQLite database to copy structure
                with zipfile.ZipFile(self.z21_file, 'r') as input_zip:
                    sqlite_files = [
                        f for f in input_zip.namelist()
                        if f.endswith('.sqlite')
                    ]
                    if not sqlite_files:
                        messagebox.showerror(
                            "Error",
                            "No SQLite database found in source file.")
                        return

                    sqlite_file = sqlite_files[0]
                    sqlite_data = input_zip.read(sqlite_file)

                    # Extract to temporary file
                    with tempfile.NamedTemporaryFile(delete=False,
                                                     suffix='.sqlite') as tmp:
                        tmp.write(sqlite_data)
                        source_db_path = tmp.name

                    try:
                        # Connect to source database
                        source_db = sqlite3.connect(source_db_path)
                        source_db.row_factory = sqlite3.Row
                        source_cursor = source_db.cursor()

                        # Create new database for single locomotive
                        new_db_path = export_path / "Loco.sqlite"
                        new_db = sqlite3.connect(str(new_db_path))
                        new_cursor = new_db.cursor()

                        # Copy all table schemas from source database
                        source_cursor.execute(
                            "SELECT name FROM sqlite_master WHERE type='table'"
                        )
                        tables = [row[0] for row in source_cursor.fetchall()]

                        for table in tables:
                            # Get CREATE TABLE statement
                            source_cursor.execute(
                                f"SELECT sql FROM sqlite_master WHERE type='table' AND name='{table}'"
                            )
                            create_sql = source_cursor.fetchone()
                            if create_sql and create_sql[0]:
                                new_cursor.execute(create_sql[0])

                        # Copy update_history if exists
                        if 'update_history' in tables:
                            source_cursor.execute(
                                "SELECT * FROM update_history")
                            for row in source_cursor.fetchall():
                                columns = ', '.join(row.keys())
                                placeholders = ', '.join(['?' for _ in row])
                                values = tuple(row)
                                new_cursor.execute(
                                    f"INSERT INTO update_history ({columns}) VALUES ({placeholders})",
                                    values)

                        # Get vehicle ID for current locomotive
                        vehicle_id = getattr(self.current_loco, '_vehicle_id',
                                             None)
                        if not vehicle_id:
                            # Try to find by address
                            source_cursor.execute(
                                "SELECT id FROM vehicles WHERE type = 0 AND address = ?",
                                (self.current_loco.address, ))
                            row = source_cursor.fetchone()
                            if row:
                                vehicle_id = row['id']

                        if vehicle_id:
                            # Copy vehicle data
                            source_cursor.execute(
                                "SELECT * FROM vehicles WHERE id = ?",
                                (vehicle_id, ))
                            vehicle_row = source_cursor.fetchone()
                            if vehicle_row:
                                # Get column names
                                columns = ', '.join(vehicle_row.keys())
                                placeholders = ', '.join(
                                    ['?' for _ in vehicle_row])
                                values = tuple(vehicle_row)
                                new_cursor.execute(
                                    f"INSERT INTO vehicles ({columns}) VALUES ({placeholders})",
                                    values)

                                # Export all functions from memory (current_loco.function_details)
                                # This ensures all functions are exported, including any modifications made in GUI
                                if 'functions' in tables:
                                    # First, get the functions table structure to understand columns
                                    source_cursor.execute(
                                        "PRAGMA table_info(functions)")
                                    func_columns_info = source_cursor.fetchall(
                                    )
                                    func_column_names = [
                                        col[1] for col in func_columns_info
                                    ]

                                    # Get a sample function row from source to understand all required columns
                                    source_cursor.execute(
                                        "SELECT * FROM functions LIMIT 1")
                                    sample_func = source_cursor.fetchone()
                                    if sample_func:
                                        all_func_columns = list(
                                            sample_func.keys())
                                    else:
                                        # If no sample, use column names from table info
                                        all_func_columns = func_column_names

                                    # Get the maximum id from existing functions in the new database
                                    new_cursor.execute(
                                        "SELECT MAX(id) FROM functions")
                                    max_id_result = new_cursor.fetchone()
                                    next_id = (
                                        max_id_result[0] + 1
                                    ) if max_id_result[0] is not None else 1

                                    # Delete any existing functions for this vehicle in the new database
                                    new_cursor.execute(
                                        "DELETE FROM functions WHERE vehicle_id = ?",
                                        (vehicle_id, ))

                                    # Export all functions from current_loco.function_details
                                    if self.current_loco and self.current_loco.function_details:
                                        for func_num, func_info in self.current_loco.function_details.items(
                                        ):
                                            # Build function row data with ALL columns from sample
                                            func_values = []

                                            for col in all_func_columns:
                                                if col == 'id':
                                                    # Generate id automatically based on existing functions
                                                    func_values.append(next_id)
                                                    next_id += 1
                                                elif col == 'vehicle_id':
                                                    func_values.append(
                                                        vehicle_id)
                                                elif col == 'function':
                                                    func_values.append(
                                                        func_num)
                                                elif col == 'position':
                                                    # Preserve the position from function_info
                                                    func_values.append(func_info.position)
                                                elif col == 'shortcut':
                                                    func_values.append(
                                                        func_info.shortcut
                                                        or '')
                                                elif col == 'time':
                                                    # Try to get original time format from source database first
                                                    source_cursor.execute(
                                                        "SELECT time FROM functions WHERE vehicle_id = ? AND function = ? LIMIT 1",
                                                        (vehicle_id, func_num))
                                                    orig_time_row = source_cursor.fetchone(
                                                    )

                                                    if orig_time_row and orig_time_row[
                                                            0] is not None:
                                                        # Use original format from source database
                                                        orig_time_str = str(
                                                            orig_time_row[0])
                                                        # Normalize to '0.000000' format to match back.z21loco (from Z21 export)
                                                        try:
                                                            time_float = float(
                                                                orig_time_str)
                                                            if time_float == 0:
                                                                # Use '0.000000' format to match back.z21loco
                                                                func_values.append(
                                                                    '0.000000')
                                                            else:
                                                                # For non-zero, preserve original format
                                                                func_values.append(
                                                                    orig_time_str
                                                                )
                                                        except (ValueError,
                                                                TypeError):
                                                            func_values.append(
                                                                '0.000000')
                                                    else:
                                                        # If not found in source, use function_info value
                                                        time_val = func_info.time or '0'
                                                        try:
                                                            time_float = float(
                                                                time_val)
                                                            if time_float == 0:
                                                                # For zero, use '0.000000' format to match back.z21loco
                                                                func_values.append(
                                                                    '0.000000')
                                                            else:
                                                                # For non-zero, preserve the format
                                                                func_values.append(
                                                                    str(time_val
                                                                        ))
                                                        except (ValueError,
                                                                TypeError):
                                                            func_values.append(
                                                                '0')
                                                elif col == 'image_name':
                                                    func_values.append(
                                                        func_info.image_name
                                                        or '')
                                                elif col == 'button_type':
                                                    func_values.append(
                                                        func_info.button_type)
                                                elif col == 'is_configured':
                                                    # Set to 0 to match Z21 export format (0 = configured/exported)
                                                    func_values.append(0)
                                                elif col == 'show_function_number':
                                                    func_values.append(1)
                                                else:
                                                    # For any other unknown columns, use None
                                                    func_values.append(None)

                                            # Insert function with all columns
                                            try:
                                                new_cursor.execute(
                                                    f"INSERT INTO functions ({', '.join(all_func_columns)}) VALUES ({', '.join(['?' for _ in all_func_columns])})",
                                                    tuple(func_values))
                                            except Exception as e:
                                                print(
                                                    f"Error inserting function {func_num}: {e}"
                                                )
                                                import traceback
                                                traceback.print_exc()
                                    else:
                                        # Fallback: copy from source database if no function_details in memory
                                        source_cursor.execute(
                                            "SELECT * FROM functions WHERE vehicle_id = ?",
                                            (vehicle_id, ))
                                        for func_row in source_cursor.fetchall(
                                        ):
                                            func_columns = ', '.join(
                                                func_row.keys())
                                            func_placeholders = ', '.join(
                                                ['?' for _ in func_row])
                                            func_values = tuple(func_row)
                                            new_cursor.execute(
                                                f"INSERT INTO functions ({func_columns}) VALUES ({func_placeholders})",
                                                func_values)

                                # Skip CVs - not exporting CVs as requested

                                # Skip traction_list - not exporting to match z21_loco.z21loco format

                                # Copy categories
                                source_cursor.execute(
                                    """
                                    SELECT vtc.* FROM vehicles_to_categories vtc
                                    WHERE vtc.vehicle_id = ?
                                """, (vehicle_id, ))
                                for cat_row in source_cursor.fetchall():
                                    # First ensure category exists
                                    source_cursor.execute(
                                        "SELECT * FROM categories WHERE id = ?",
                                        (cat_row['category_id'], ))
                                    cat_data = source_cursor.fetchone()
                                    if cat_data:
                                        # Insert category if not exists
                                        new_cursor.execute(
                                            "SELECT id FROM categories WHERE id = ?",
                                            (cat_data['id'], ))
                                        if not new_cursor.fetchone():
                                            cat_columns = ', '.join(
                                                cat_data.keys())
                                            cat_placeholders = ', '.join(
                                                ['?' for _ in cat_data])
                                            cat_values = tuple(cat_data)
                                            new_cursor.execute(
                                                f"INSERT INTO categories ({cat_columns}) VALUES ({cat_placeholders})",
                                                cat_values)

                                    # Insert vehicle_to_category link
                                    vtc_columns = ', '.join(cat_row.keys())
                                    vtc_placeholders = ', '.join(
                                        ['?' for _ in cat_row])
                                    vtc_values = tuple(cat_row)
                                    new_cursor.execute(
                                        f"INSERT INTO vehicles_to_categories ({vtc_columns}) VALUES ({vtc_placeholders})",
                                        vtc_values)

                        new_db.commit()
                        new_db.close()
                        source_db.close()

                        # Set text encoding to UTF-16le (16) for Z21 APP compatibility
                        # Text encoding is stored at offset 60-63 (4 bytes, big-endian)
                        with open(new_db_path, 'rb') as f:
                            sqlite_data = bytearray(f.read())
                        sqlite_data[60:64] = (16).to_bytes(4, 'big')
                        with open(new_db_path, 'wb') as f:
                            f.write(sqlite_data)

                        # Copy locomotive image if exists
                        if self.current_loco.image_name:
                            # Try to find image in original ZIP
                            image_found = False
                            for filename in input_zip.namelist():
                                if self.current_loco.image_name in filename or filename.endswith(
                                        f"lok_{self.current_loco.address}.png"
                                ):
                                    image_data = input_zip.read(filename)
                                    # Determine image filename
                                    if filename.endswith('.png'):
                                        image_filename = filename.split(
                                            '/')[-1]
                                    else:
                                        image_filename = f"lok_{self.current_loco.address}.png"
                                    (export_path /
                                     image_filename).write_bytes(image_data)
                                    image_found = True
                                    break

                        # Create ZIP file
                        with zipfile.ZipFile(
                                output_path, 'w',
                                zipfile.ZIP_DEFLATED) as output_zip:
                            # Add SQLite database
                            output_zip.write(new_db_path,
                                             f"{export_dir}/Loco.sqlite")

                            # Add image if found
                            if self.current_loco.image_name:
                                for img_file in export_path.glob("*.png"):
                                    output_zip.write(
                                        img_file,
                                        f"{export_dir}/{img_file.name}")

                        messagebox.showinfo(
                            "Success",
                            f"Locomotive exported successfully to:\n{output_path}"
                        )

                    finally:
                        # Clean up temporary files
                        Path(source_db_path).unlink()

        except Exception as e:
            messagebox.showerror("Export Error",
                                 f"Failed to export locomotive: {e}")
            import traceback
            traceback.print_exc()

    def share_with_airdrop(self):
        """Share z21loco file via AirDrop using NSSharingService (macOS)."""
        if not self.current_loco or not self.z21_data or not self.parser:
            messagebox.showerror("Error",
                                 "No locomotive selected or data not loaded.")
            return

        # Check if PyObjC is available (macOS only)
        if platform.system() != 'Darwin':
            messagebox.showerror("Error",
                                 "AirDrop sharing is only available on macOS.")
            return

        if not HAS_PYOBJC:
            messagebox.showerror(
                "Error",
                "PyObjC is required for AirDrop sharing.\n\n"
                "Please install it with:\n"
                "pip install pyobjc-framework-AppKit"
            )
            return

        try:
            # Create temporary file for sharing (NSSharingService requires a file path)
            # Use a descriptive filename based on locomotive name
            loco_name = self.current_loco.name.replace('/', '_').replace('\\', '_')
            if not loco_name:
                loco_name = f"locomotive_{self.current_loco.address}"
            
            # Create temporary file in system temp directory
            temp_dir = tempfile.gettempdir()
            temp_filename = f"{loco_name}_{uuid.uuid4().hex[:8]}.z21loco"
            output_path = Path(temp_dir) / temp_filename
            
            # Use the existing export_z21_loco logic but without showing success message
            # We'll call the export logic directly
            import shutil

            # Generate UUID for export directory
            export_uuid = str(uuid.uuid4()).upper()
            export_dir = f"export/{export_uuid}"

            # Create temporary directory for export
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                export_path = temp_path / export_dir
                export_path.mkdir(parents=True, exist_ok=True)

                # Get original SQLite database to copy structure
                with zipfile.ZipFile(self.z21_file, 'r') as input_zip:
                    sqlite_files = [
                        f for f in input_zip.namelist()
                        if f.endswith('.sqlite')
                    ]
                    if not sqlite_files:
                        messagebox.showerror(
                            "Error",
                            "No SQLite database found in source file.")
                        return

                    sqlite_file = sqlite_files[0]
                    sqlite_data = input_zip.read(sqlite_file)

                    # Extract to temporary file
                    with tempfile.NamedTemporaryFile(delete=False,
                                                     suffix='.sqlite') as tmp:
                        tmp.write(sqlite_data)
                        source_db_path = tmp.name

                    try:
                        # Connect to source database
                        source_db = sqlite3.connect(source_db_path)
                        source_db.row_factory = sqlite3.Row
                        source_cursor = source_db.cursor()

                        # Create new database for single locomotive
                        new_db_path = export_path / "Loco.sqlite"
                        new_db = sqlite3.connect(str(new_db_path))
                        new_cursor = new_db.cursor()

                        # Copy all table schemas from source database
                        source_cursor.execute(
                            "SELECT name FROM sqlite_master WHERE type='table'"
                        )
                        tables = [row[0] for row in source_cursor.fetchall()]

                        for table in tables:
                            source_cursor.execute(
                                f"SELECT sql FROM sqlite_master WHERE type='table' AND name='{table}'"
                            )
                            create_sql = source_cursor.fetchone()
                            if create_sql and create_sql[0]:
                                new_cursor.execute(create_sql[0])

                        # Copy update_history if exists
                        if 'update_history' in tables:
                            source_cursor.execute(
                                "SELECT * FROM update_history")
                            for row in source_cursor.fetchall():
                                columns = ', '.join(row.keys())
                                placeholders = ', '.join(['?' for _ in row])
                                values = tuple(row)
                                new_cursor.execute(
                                    f"INSERT INTO update_history ({columns}) VALUES ({placeholders})",
                                    values)

                        # Get vehicle ID for current locomotive
                        vehicle_id = getattr(self.current_loco, '_vehicle_id',
                                             None)
                        if not vehicle_id:
                            source_cursor.execute(
                                "SELECT id FROM vehicles WHERE type = 0 AND address = ?",
                                (self.current_loco.address, ))
                            row = source_cursor.fetchone()
                            if row:
                                vehicle_id = row['id']

                        if vehicle_id:
                            # Copy vehicle data
                            source_cursor.execute(
                                "SELECT * FROM vehicles WHERE id = ?",
                                (vehicle_id, ))
                            vehicle_row = source_cursor.fetchone()
                            if vehicle_row:
                                columns = ', '.join(vehicle_row.keys())
                                placeholders = ', '.join(
                                    ['?' for _ in vehicle_row])
                                values = tuple(vehicle_row)
                                new_cursor.execute(
                                    f"INSERT INTO vehicles ({columns}) VALUES ({placeholders})",
                                    values)

                                # Export all functions from memory
                                if 'functions' in tables:
                                    source_cursor.execute(
                                        "PRAGMA table_info(functions)")
                                    func_columns_info = source_cursor.fetchall()
                                    func_column_names = [
                                        col[1] for col in func_columns_info
                                    ]

                                    source_cursor.execute(
                                        "SELECT * FROM functions LIMIT 1")
                                    sample_func = source_cursor.fetchone()
                                    if sample_func:
                                        all_func_columns = list(
                                            sample_func.keys())
                                    else:
                                        all_func_columns = func_column_names

                                    new_cursor.execute(
                                        "SELECT MAX(id) FROM functions")
                                    max_id_result = new_cursor.fetchone()
                                    next_id = (
                                        max_id_result[0] + 1
                                    ) if max_id_result[0] is not None else 1

                                    new_cursor.execute(
                                        "DELETE FROM functions WHERE vehicle_id = ?",
                                        (vehicle_id, ))

                                    if self.current_loco and self.current_loco.function_details:
                                        for func_num, func_info in self.current_loco.function_details.items():
                                            func_values = []

                                            for col in all_func_columns:
                                                if col == 'id':
                                                    func_values.append(next_id)
                                                    next_id += 1
                                                elif col == 'vehicle_id':
                                                    func_values.append(vehicle_id)
                                                elif col == 'function':
                                                    func_values.append(func_num)
                                                elif col == 'position':
                                                    func_values.append(func_info.position)
                                                elif col == 'shortcut':
                                                    func_values.append(func_info.shortcut or '')
                                                elif col == 'time':
                                                    source_cursor.execute(
                                                        "SELECT time FROM functions WHERE vehicle_id = ? AND function = ? LIMIT 1",
                                                        (vehicle_id, func_num))
                                                    orig_time_row = source_cursor.fetchone()

                                                    if orig_time_row and orig_time_row[0] is not None:
                                                        orig_time_str = str(orig_time_row[0])
                                                        try:
                                                            time_float = float(orig_time_str)
                                                            if time_float == 0:
                                                                func_values.append('0.000000')
                                                            else:
                                                                func_values.append(orig_time_str)
                                                        except (ValueError, TypeError):
                                                            func_values.append('0.000000')
                                                    else:
                                                        time_val = func_info.time or '0'
                                                        try:
                                                            time_float = float(time_val)
                                                            if time_float == 0:
                                                                func_values.append('0.000000')
                                                            else:
                                                                func_values.append(str(time_val))
                                                        except (ValueError, TypeError):
                                                            func_values.append('0')
                                                elif col == 'image_name':
                                                    func_values.append(func_info.image_name or '')
                                                elif col == 'button_type':
                                                    func_values.append(func_info.button_type)
                                                elif col == 'is_configured':
                                                    func_values.append(0)
                                                elif col == 'show_function_number':
                                                    func_values.append(1)
                                                else:
                                                    func_values.append(None)

                                            try:
                                                new_cursor.execute(
                                                    f"INSERT INTO functions ({', '.join(all_func_columns)}) VALUES ({', '.join(['?' for _ in all_func_columns])})",
                                                    tuple(func_values))
                                            except Exception as e:
                                                print(f"Error inserting function {func_num}: {e}")

                        new_db.commit()
                        new_db.close()
                        source_db.close()

                        # Set text encoding to UTF-16le (16) for Z21 APP compatibility
                        with open(new_db_path, 'rb') as f:
                            sqlite_data = bytearray(f.read())
                        sqlite_data[60:64] = (16).to_bytes(4, 'big')
                        with open(new_db_path, 'wb') as f:
                            f.write(sqlite_data)

                        # Copy locomotive image if exists
                        if self.current_loco.image_name:
                            for filename in input_zip.namelist():
                                if self.current_loco.image_name in filename or filename.endswith(
                                        f"lok_{self.current_loco.address}.png"):
                                    image_data = input_zip.read(filename)
                                    if filename.endswith('.png'):
                                        image_filename = filename.split('/')[-1]
                                    else:
                                        image_filename = f"lok_{self.current_loco.address}.png"
                                    (export_path / image_filename).write_bytes(image_data)
                                    break

                        # Create ZIP file
                        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as output_zip:
                            output_zip.write(new_db_path, f"{export_dir}/Loco.sqlite")
                            if self.current_loco.image_name:
                                for img_file in export_path.glob("*.png"):
                                    output_zip.write(img_file, f"{export_dir}/{img_file.name}")

                    finally:
                        Path(source_db_path).unlink()

            # Share file via NSSharingService (AirDrop)
            try:
                # Convert Path to NSURL
                file_path_str = str(output_path.absolute())
                file_url = NSURL.fileURLWithPath_(file_path_str)
                
                # Create NSArray with file URL for sharing
                file_array = NSArray.arrayWithObject_(file_url)
                
                # Get AirDrop sharing service
                # AirDrop service may fail for several reasons:
                # 1. AirDrop is disabled in System Settings
                # 2. Bluetooth/WiFi is turned off
                # 3. The service name is incorrect for this macOS version
                # 4. AirDrop is not available (older macOS versions)
                
                sharing_service = None
                
                # Method 1: Try with the standard AirDrop service name
                try:
                    sharing_service = NSSharingService.sharingServiceNamed_(
                        'com.apple.share.AirDrop'
                    )
                    if sharing_service:
                        print(f"✓ Found AirDrop service via name: {sharing_service.title()}")
                except Exception as e:
                    print(f"✗ Method 1 failed: {e}")
                
                # Method 2: If that fails, try to find AirDrop from available services
                # This is more reliable as it queries what services are actually available
                if not sharing_service:
                    try:
                        # Get all available sharing services for the file
                        available_services = NSSharingService.sharingServicesForItems_(file_array)
                        print(f"Available sharing services: {[s.title() for s in available_services]}")
                        
                        for service in available_services:
                            # Check if this is AirDrop by title or identifier
                            service_title = service.title()
                            if 'AirDrop' in service_title or 'airdrop' in service_title.lower():
                                sharing_service = service
                                print(f"✓ Found AirDrop service via available services: {service_title}")
                                break
                    except Exception as e:
                        print(f"✗ Method 2 failed: {e}")
                
                # Method 3: Try alternative service name (for older macOS versions)
                if not sharing_service:
                    try:
                        # Some macOS versions use different identifiers
                        sharing_service = NSSharingService.sharingServiceNamed_(
                            'NSSharingServiceNameAirDrop'
                        )
                        if sharing_service:
                            print(f"✓ Found AirDrop service via alternative name")
                    except Exception as e:
                        print(f"✗ Method 3 failed: {e}")
                
                if sharing_service:
                    # Check if sharing service can perform with items
                    if sharing_service.canPerformWithItems_(file_array):
                        # Perform sharing - this will open the AirDrop share dialog
                        sharing_service.performWithItems_(file_array)
                        
                    else:
                        # Fallback: open AirDrop window and show file
                        subprocess.run(['open', 'airdrop://'], check=False)
                        subprocess.run(['open', '-R', file_path_str], check=True)
                        messagebox.showinfo(
                            "Share with AirDrop",
                            f"File exported to:\n{output_path}\n\n"
                            "AirDrop window opened. Drag the file to share."
                        )
                else:
                    # AirDrop service not available - possible reasons:
                    # 1. AirDrop is disabled in System Settings > General > AirDrop
                    # 2. Bluetooth/WiFi is turned off
                    # 3. macOS version doesn't support AirDrop sharing service
                    # 4. AirDrop is not available on this Mac
                    
                    print("⚠ AirDrop sharing service not available")
                    print("  Possible reasons:")
                    print("  1. AirDrop is disabled in System Settings")
                    print("  2. Bluetooth/WiFi is turned off")
                    print("  3. AirDrop not available on this Mac")
                    
                    # Fallback: open AirDrop window and show file
                    subprocess.run(['open', 'airdrop://'], check=False)
                    subprocess.run(['open', '-R', file_path_str], check=True)
                    messagebox.showwarning(
                        "AirDrop Service Not Available",
                        f"File exported to:\n{output_path}\n\n"
                        "AirDrop sharing service is not available.\n\n"
                        "Possible reasons:\n"
                        "• AirDrop is disabled in System Settings\n"
                        "• Bluetooth/WiFi is turned off\n"
                        "• AirDrop not available on this Mac\n\n"
                        "AirDrop window opened. Please drag the file manually to share."
                    )
                    
            except Exception as e:
                # Fallback: show file in Finder
                try:
                    subprocess.run(['open', '-R', str(output_path)], check=True)
                    subprocess.run(['open', 'airdrop://'], check=False)
                    messagebox.showinfo(
                        "File Ready",
                        f"File exported to:\n{output_path}\n\n"
                        "File shown in Finder. Drag to AirDrop to share."
                    )
                except Exception as e2:
                    messagebox.showinfo(
                        "Export Complete",
                        f"File exported to:\n{output_path}\n\n"
                        "Please manually share this file via AirDrop."
                    )
            finally:
                # Note: We don't delete the temp file immediately because
                # AirDrop sharing happens asynchronously. The file will be cleaned up
                # by the system's temp file cleanup, or we could add a cleanup mechanism later.
                pass

        except Exception as e:
            messagebox.showerror("Share Error",
                                 f"Failed to share locomotive: {e}")
            import traceback
            traceback.print_exc()

    def update_functions(self):
        """Update functions tab."""
        loco = self.current_loco

        # Clear existing widgets
        for widget in self.functions_frame_inner.winfo_children():
            widget.destroy()

        # Rebind mouse wheel events after clearing widgets
        if hasattr(self, 'update_scroll_bindings'):
            self.update_scroll_bindings()

        # Ensure canvas has focus for scrolling
        self.functions_canvas.focus_set()

        # Calculate grid layout based on available width
        # Each card is approximately 100 pixels wide (80 icon + padding)
        # Calculate columns based on canvas width
        self.functions_canvas.update_idletasks()  # Update to get actual width
        canvas_width = self.functions_canvas.winfo_width()
        if canvas_width < 100:
            canvas_width = 800  # Default width if not yet rendered

        card_width = 100  # Fixed card width (matches CARD_WIDTH in create_function_card)
        cols = max(1, (canvas_width - 40) //
                   card_width)  # Account for scrollbar and padding

        # Row 0: Title
        header_label = ttk.Label(self.functions_frame_inner,
                                 text=f"Functions for {loco.name}",
                                 font=('Arial', 14, 'bold'))
        header_label.grid(row=0,
                          column=0,
                          columnspan=cols,
                          sticky='ew',
                          padx=5,
                          pady=(10, 5))

        # Row 1: "Add New Function", "Scan for Functions", and "Save Changes" buttons
        # Always show these buttons even if no functions exist
        button_frame = ttk.Frame(self.functions_frame_inner)
        button_frame.grid(row=1,
                          column=0,
                          columnspan=cols,
                          sticky='ew',
                          padx=5,
                          pady=(0, 10))

        add_button = ttk.Button(button_frame,
                                text="+ Add New Function",
                                command=self.add_new_function)
        add_button.pack(side=tk.LEFT, padx=(0, 10))

        scan_functions_button = ttk.Button(button_frame,
                                           text="📷 Scan for Functions",
                                           command=self.scan_for_functions)
        scan_functions_button.pack(side=tk.LEFT, padx=(0, 10))

        save_button = ttk.Button(button_frame,
                                 text="💾 Save Changes",
                                 command=self.save_function_changes)
        save_button.pack(side=tk.LEFT)

        # Check if there are any functions
        if not loco.function_details:
            # Show message that no functions exist, but buttons are still available
            no_funcs_label = ttk.Label(
                self.functions_frame_inner,
                text=
                "No functions configured. Use 'Add New Function' or 'Scan for Functions' to add functions.",
                font=('Arial', 11),
                foreground='gray')
            no_funcs_label.grid(row=2,
                                column=0,
                                columnspan=cols,
                                sticky='ew',
                                padx=5,
                                pady=20)
            return

        # Sort functions by function number
        sorted_funcs = sorted(loco.function_details.items(),
                              key=lambda x: x[1].function_number)

        # Create function cards in a grid layout
        row = 2  # Start after title and button
        col = 0

        for func_num, func_info in sorted_funcs:
            card_frame = self.create_function_card(func_num, func_info)

            # Make card and all children clickable to edit function
            def make_clickable(widget, fn, fi):
                widget.bind("<Button-1>",
                            lambda e, fnum=fn, finfo=fi: self.edit_function(
                                fnum, finfo))
                widget.bind("<Enter>",
                            lambda e: e.widget.config(cursor="hand2"))
                widget.bind("<Leave>", lambda e: e.widget.config(cursor=""))
                for child in widget.winfo_children():
                    make_clickable(child, fn, fi)

            make_clickable(card_frame, func_num, func_info)

            # Place in grid
            card_frame.grid(row=row, column=col, padx=5, pady=5, sticky='nw')

            col += 1
            if col >= cols:
                col = 0
                row += 1

        # Configure grid columns to be equal width
        for i in range(cols):
            self.functions_frame_inner.grid_columnconfigure(i,
                                                            weight=0,
                                                            uniform='card')

    def save_function_changes(self):
        """Save all function changes to the Z21 file."""
        if not self.current_loco or not self.z21_data or not self.parser:
            messagebox.showerror("Error",
                                 "No locomotive selected or data not loaded.")
            return

        try:
            # Ensure locomotive is updated in z21_data
            if self.current_loco_index is not None:
                self.z21_data.locomotives[
                    self.current_loco_index] = self.current_loco

            # Write changes back to file
            self.parser.write(self.z21_data, self.z21_file)
            messagebox.showinfo(
                "Success", "All function changes saved successfully to file!")
        except Exception as write_error:
            messagebox.showerror(
                "Write Error",
                f"Failed to write changes to file: {write_error}\n\n"
                f"Changes have been saved in memory but not written to disk.")

    def get_next_unused_function_number(self):
        """Get the next unused function number for the current locomotive."""
        if not self.current_loco:
            return 0

        used_numbers = set(self.current_loco.function_details.keys())
        # Start from 0 and find first unused
        for i in range(128):  # DCC functions typically go up to F127
            if i not in used_numbers:
                return i
        return 128  # Fallback if all are used

    def get_available_icons(self):
        """Get list of available icon names from icon mapping."""
        icon_names = sorted(self.icon_mapping.keys())
        # Also add common icon names that might not be in mapping
        common_icons = [
            'light', 'bell', 'horn_two_sound', 'steam', 'whistle_long',
            'whistle_short', 'neutral', 'sound1', 'sound2', 'sound3', 'sound4'
        ]
        for icon in common_icons:
            if icon not in icon_names:
                icon_names.append(icon)
        return sorted(set(icon_names))

    def add_new_function(self):
        """Open dialog to add a new function."""
        if not self.current_loco:
            messagebox.showwarning("No Locomotive",
                                   "Please select a locomotive first.")
            return

        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Add New Function")
        dialog.transient(self.root)
        dialog.grab_set()

        # Variables
        icon_var = tk.StringVar()
        func_num_var = tk.StringVar(
            value=str(self.get_next_unused_function_number()))
        shortcut_var = tk.StringVar()
        button_type_var = tk.StringVar(value="switch")
        time_var = tk.StringVar(value="1.0")

        # Main container with padding
        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Top section: Icon preview (larger, centered)
        preview_frame = ttk.Frame(main_frame)
        preview_frame.pack(fill=tk.X, pady=(0, 15))

        icon_preview_label = ttk.Label(preview_frame,
                                       background='white',
                                       relief=tk.SUNKEN,
                                       borderwidth=2)
        icon_preview_label.pack()

        def update_icon_preview(*args):
            """Update icon preview when selection changes."""
            icon_name = icon_var.get()
            if icon_name:
                preview_image = self.load_icon_image(icon_name, (80, 80))
                if preview_image:
                    icon_preview_label.config(image=preview_image)
                    icon_preview_label.image = preview_image  # Keep a reference
                else:
                    # Clear preview if icon not found
                    icon_preview_label.config(image='', width=80, height=80)
            else:
                icon_preview_label.config(image='', width=80, height=80)

        icon_var.trace('w', update_icon_preview)

        # Form fields in two columns for better space usage
        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.BOTH, expand=True)

        # Left column
        left_col = ttk.Frame(form_frame)
        left_col.grid(row=0, column=0, padx=(0, 10), sticky='nsew')

        # Right column
        right_col = ttk.Frame(form_frame)
        right_col.grid(row=0, column=1, padx=(10, 0), sticky='nsew')

        form_frame.grid_columnconfigure(0, weight=1)
        form_frame.grid_columnconfigure(1, weight=1)
        form_frame.grid_rowconfigure(0, weight=1)

        # Left column fields
        row = 0

        # Icon selection
        ttk.Label(left_col, text="Icon:", width=12,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=3,
                                   pady=4,
                                   sticky='e')
        icon_combo = ttk.Combobox(left_col,
                                  textvariable=icon_var,
                                  width=20,
                                  state='readonly')
        icon_combo['values'] = self.get_available_icons()
        if icon_combo['values']:
            icon_combo.current(0)
            update_icon_preview()  # Initial preview
        icon_combo.grid(row=row, column=1, padx=3, pady=4, sticky='ew')
        row += 1

        # Function number
        ttk.Label(left_col, text="Function #:", width=12,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=3,
                                   pady=4,
                                   sticky='e')
        func_num_entry = ttk.Entry(left_col,
                                   textvariable=func_num_var,
                                   width=20)
        func_num_entry.grid(row=row, column=1, padx=3, pady=4, sticky='ew')
        row += 1

        # Right column fields
        row_right = 0

        # Shortcut
        ttk.Label(right_col, text="Shortcut:", width=12,
                  anchor='e').grid(row=row_right,
                                   column=0,
                                   padx=3,
                                   pady=4,
                                   sticky='e')
        shortcut_entry = ttk.Entry(right_col,
                                   textvariable=shortcut_var,
                                   width=20)
        shortcut_entry.grid(row=row_right,
                            column=1,
                            padx=3,
                            pady=4,
                            sticky='ew')
        row_right += 1

        # Button type
        ttk.Label(right_col, text="Button Type:", width=12,
                  anchor='e').grid(row=row_right,
                                   column=0,
                                   padx=3,
                                   pady=4,
                                   sticky='e')
        button_type_combo = ttk.Combobox(
            right_col,
            textvariable=button_type_var,
            values=['switch', 'push-button', 'time button'],
            state='readonly',
            width=17)
        button_type_combo.current(0)
        button_type_combo.grid(row=row_right,
                               column=1,
                               padx=3,
                               pady=4,
                               sticky='ew')
        row_right += 1

        # Time duration (only show for time button) - in right column
        time_label = ttk.Label(right_col,
                               text="Time (s):",
                               width=12,
                               anchor='e')
        time_entry = ttk.Entry(right_col, textvariable=time_var, width=20)

        def update_time_visibility(*args):
            """Show/hide time duration field based on button type."""
            if button_type_var.get() == 'time button':
                time_label.grid(row=row_right,
                                column=0,
                                padx=3,
                                pady=4,
                                sticky='e')
                time_entry.grid(row=row_right,
                                column=1,
                                padx=3,
                                pady=4,
                                sticky='ew')
            else:
                time_label.grid_remove()
                time_entry.grid_remove()

        button_type_var.trace('w', update_time_visibility)
        update_time_visibility()

        # Configure column weights
        left_col.grid_columnconfigure(1, weight=1)
        right_col.grid_columnconfigure(1, weight=1)

        # Buttons
        button_frame = ttk.Frame(main_frame, padding=(10, 15, 10, 10))
        button_frame.pack(fill=tk.X)

        # Calculate optimal window size
        dialog.update_idletasks()
        # Get natural size
        width = dialog.winfo_reqwidth()
        height = dialog.winfo_reqheight()
        # Set minimum size with some padding
        dialog.geometry(f"{max(480, width)}x{max(320, height)}")
        dialog.minsize(480, 320)

        def save_function():
            """Save the new function."""
            try:
                # Validate inputs
                icon_name = icon_var.get()
                if not icon_name:
                    messagebox.showerror("Error", "Please select an icon.")
                    return

                func_num = int(func_num_var.get())
                if func_num < 0 or func_num > 127:
                    messagebox.showerror(
                        "Error", "Function number must be between 0 and 127.")
                    return

                # Check if function number already exists
                if func_num in self.current_loco.function_details:
                    if not messagebox.askyesno(
                            "Overwrite?",
                            f"Function F{func_num} already exists. Overwrite it?"
                    ):
                        return

                shortcut = shortcut_var.get().strip()
                button_type_name = button_type_var.get()

                # Map button type name to integer
                button_type_map = {
                    'switch': 0,
                    'push-button': 1,
                    'time button': 2
                }
                button_type = button_type_map.get(button_type_name, 0)

                # Get time duration (only for time button, otherwise "0")
                if button_type == 2:  # time button
                    try:
                        time_value = float(time_var.get())
                        time_str = str(time_value)
                    except ValueError:
                        messagebox.showerror(
                            "Error", "Time duration must be a valid number.")
                        return
                else:
                    time_str = "0"

                # Find max position for ordering
                max_position = 0
                if self.current_loco.function_details:
                    max_position = max(
                        f.position
                        for f in self.current_loco.function_details.values())

                # Create new function info
                func_info = FunctionInfo(function_number=func_num,
                                         image_name=icon_name,
                                         shortcut=shortcut,
                                         position=max_position + 1,
                                         time=time_str,
                                         button_type=button_type,
                                         is_active=True)

                # Add to locomotive
                self.current_loco.function_details[func_num] = func_info
                self.current_loco.functions[func_num] = True

                # Update locomotive in z21_data
                if self.current_loco_index is not None:
                    self.z21_data.locomotives[
                        self.current_loco_index] = self.current_loco

                # Update display
                self.update_functions()

                # Close dialog
                dialog.destroy()

                messagebox.showinfo(
                    "Success", f"Function F{func_num} added successfully!")

            except ValueError as e:
                messagebox.showerror("Error", f"Invalid input: {e}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add function: {e}")

        ttk.Button(button_frame, text="Cancel",
                   command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Add Function",
                   command=save_function).pack(side=tk.RIGHT, padx=5)

    def edit_function(self, func_num: int, func_info: FunctionInfo):
        """Open dialog to edit an existing function."""
        if not self.current_loco:
            messagebox.showwarning("No Locomotive",
                                   "Please select a locomotive first.")
            return

        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit Function F{func_num}")
        dialog.transient(self.root)
        dialog.grab_set()

        # Variables - pre-populate with existing values
        icon_var = tk.StringVar(value=func_info.image_name)
        func_num_var = tk.StringVar(value=str(func_num))
        shortcut_var = tk.StringVar(value=func_info.shortcut or "")
        button_type_map = {0: "switch", 1: "push-button", 2: "time button"}
        button_type_var = tk.StringVar(
            value=button_type_map.get(func_info.button_type, "switch"))
        time_var = tk.StringVar(
            value=func_info.time if func_info.time != "0" else "1.0")

        # Main container with padding
        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Top section: Icon preview (larger, centered)
        preview_frame = ttk.Frame(main_frame)
        preview_frame.pack(fill=tk.X, pady=(0, 15))

        icon_preview_label = ttk.Label(preview_frame,
                                       background='white',
                                       relief=tk.SUNKEN,
                                       borderwidth=2)
        icon_preview_label.pack()

        def update_icon_preview(*args):
            """Update icon preview when selection changes."""
            icon_name = icon_var.get()
            if icon_name:
                preview_image = self.load_icon_image(icon_name, (80, 80))
                if preview_image:
                    icon_preview_label.config(image=preview_image)
                    icon_preview_label.image = preview_image  # Keep a reference
                else:
                    # Clear preview if icon not found
                    icon_preview_label.config(image='', width=80, height=80)
            else:
                icon_preview_label.config(image='', width=80, height=80)

        icon_var.trace('w', update_icon_preview)
        update_icon_preview()  # Initial preview

        # Form fields in two columns for better space usage
        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.BOTH, expand=True)

        # Left column
        left_col = ttk.Frame(form_frame)
        left_col.grid(row=0, column=0, padx=(0, 10), sticky='nsew')

        # Right column
        right_col = ttk.Frame(form_frame)
        right_col.grid(row=0, column=1, padx=(10, 0), sticky='nsew')

        form_frame.grid_columnconfigure(0, weight=1)
        form_frame.grid_columnconfigure(1, weight=1)
        form_frame.grid_rowconfigure(0, weight=1)

        # Left column fields
        row = 0

        # Icon selection
        ttk.Label(left_col, text="Icon:", width=12,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=3,
                                   pady=4,
                                   sticky='e')
        icon_combo = ttk.Combobox(left_col,
                                  textvariable=icon_var,
                                  width=20,
                                  state='readonly')
        icon_combo['values'] = self.get_available_icons()
        # Set current selection
        available_icons = self.get_available_icons()
        if func_info.image_name in available_icons:
            icon_combo.current(available_icons.index(func_info.image_name))
        elif available_icons:
            icon_combo.current(0)
        icon_combo.grid(row=row, column=1, padx=3, pady=4, sticky='ew')
        row += 1

        # Function number
        ttk.Label(left_col, text="Function #:", width=12,
                  anchor='e').grid(row=row,
                                   column=0,
                                   padx=3,
                                   pady=4,
                                   sticky='e')
        func_num_entry = ttk.Entry(left_col,
                                   textvariable=func_num_var,
                                   width=20)
        func_num_entry.grid(row=row, column=1, padx=3, pady=4, sticky='ew')
        row += 1

        # Right column fields
        row_right = 0

        # Shortcut
        ttk.Label(right_col, text="Shortcut:", width=12,
                  anchor='e').grid(row=row_right,
                                   column=0,
                                   padx=3,
                                   pady=4,
                                   sticky='e')
        shortcut_entry = ttk.Entry(right_col,
                                   textvariable=shortcut_var,
                                   width=20)
        shortcut_entry.grid(row=row_right,
                            column=1,
                            padx=3,
                            pady=4,
                            sticky='ew')
        row_right += 1

        # Button type
        ttk.Label(right_col, text="Button Type:", width=12,
                  anchor='e').grid(row=row_right,
                                   column=0,
                                   padx=3,
                                   pady=4,
                                   sticky='e')
        button_type_combo = ttk.Combobox(
            right_col,
            textvariable=button_type_var,
            values=['switch', 'push-button', 'time button'],
            state='readonly',
            width=17)
        button_type_combo.current(['switch', 'push-button',
                                   'time button'].index(button_type_var.get()))
        button_type_combo.grid(row=row_right,
                               column=1,
                               padx=3,
                               pady=4,
                               sticky='ew')
        row_right += 1

        # Time duration (only show for time button) - in right column
        time_label = ttk.Label(right_col,
                               text="Time (s):",
                               width=12,
                               anchor='e')
        time_entry = ttk.Entry(right_col, textvariable=time_var, width=20)

        def update_time_visibility(*args):
            """Show/hide time duration field based on button type."""
            if button_type_var.get() == 'time button':
                time_label.grid(row=row_right,
                                column=0,
                                padx=3,
                                pady=4,
                                sticky='e')
                time_entry.grid(row=row_right,
                                column=1,
                                padx=3,
                                pady=4,
                                sticky='ew')
            else:
                time_label.grid_remove()
                time_entry.grid_remove()

        button_type_var.trace('w', update_time_visibility)
        update_time_visibility()

        # Configure column weights
        left_col.grid_columnconfigure(1, weight=1)
        right_col.grid_columnconfigure(1, weight=1)

        # Buttons
        button_frame = ttk.Frame(main_frame, padding=(10, 15, 10, 10))
        button_frame.pack(fill=tk.X)

        def save_changes():
            """Save the edited function."""
            try:
                # Validate inputs
                icon_name = icon_var.get()
                if not icon_name:
                    messagebox.showerror("Error", "Please select an icon.")
                    return

                new_func_num = int(func_num_var.get())
                if new_func_num < 0 or new_func_num > 127:
                    messagebox.showerror(
                        "Error", "Function number must be between 0 and 127.")
                    return

                shortcut = shortcut_var.get().strip()
                button_type_name = button_type_var.get()

                # Map button type name to integer
                button_type_map = {
                    'switch': 0,
                    'push-button': 1,
                    'time button': 2
                }
                button_type = button_type_map.get(button_type_name, 0)

                # Get time duration (only for time button, otherwise "0")
                if button_type == 2:  # time button
                    try:
                        time_value = float(time_var.get())
                        time_str = str(time_value)
                    except ValueError:
                        messagebox.showerror(
                            "Error", "Time duration must be a valid number.")
                        return
                else:
                    time_str = "0"

                # If function number changed, check for conflicts
                if new_func_num != func_num and new_func_num in self.current_loco.function_details:
                    if not messagebox.askyesno(
                            "Overwrite?",
                            f"Function F{new_func_num} already exists. Overwrite it?"
                    ):
                        return
                    # Remove old function number entry
                    if new_func_num != func_num:
                        del self.current_loco.function_details[func_num]
                        del self.current_loco.functions[func_num]

                # Update function info
                func_info.image_name = icon_name
                func_info.function_number = new_func_num
                func_info.shortcut = shortcut
                func_info.button_type = button_type
                func_info.time = time_str

                # Update locomotive dictionaries
                if new_func_num != func_num:
                    # Function number changed, need to update dictionaries
                    self.current_loco.function_details[
                        new_func_num] = func_info
                    self.current_loco.functions[new_func_num] = True
                else:
                    # Same function number, just update the existing entry
                    self.current_loco.function_details[func_num] = func_info

                # Update locomotive in z21_data
                if self.current_loco_index is not None:
                    self.z21_data.locomotives[
                        self.current_loco_index] = self.current_loco

                # Update display
                self.update_functions()

                # Close dialog
                dialog.destroy()

                messagebox.showinfo(
                    "Success",
                    f"Function F{new_func_num} updated successfully!")

            except ValueError as e:
                messagebox.showerror("Error", f"Invalid input: {e}")
            except Exception as e:
                messagebox.showerror("Error",
                                     f"Failed to update function: {e}")

        def delete_function():
            """Delete the function from the locomotive."""
            func_display = f"F{func_num}"
            if func_info.image_name:
                func_display += f" ({func_info.image_name})"

            if not messagebox.askyesno(
                    "Confirm Delete",
                    f"Are you sure you want to delete function {func_display}?"
            ):
                return

            # Delete the function from locomotive
            if func_num in self.current_loco.function_details:
                del self.current_loco.function_details[func_num]
            if func_num in self.current_loco.functions:
                del self.current_loco.functions[func_num]

            # Update locomotive in z21_data
            if self.current_loco_index is not None:
                self.z21_data.locomotives[
                    self.current_loco_index] = self.current_loco

            # Update display
            self.update_functions()

            # Close dialog
            dialog.destroy()

            messagebox.showinfo("Success", f"Function {func_display} deleted.")

        ttk.Button(button_frame,
                   text="Delete Function",
                   command=delete_function).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel",
                   command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Save Changes",
                   command=save_changes).pack(side=tk.RIGHT, padx=5)

        # Calculate optimal window size
        dialog.update_idletasks()
        # Get natural size
        width = dialog.winfo_reqwidth()
        height = dialog.winfo_reqheight()
        # Set minimum size with some padding
        dialog.geometry(f"{max(480, width)}x{max(320, height)}")
        dialog.minsize(480, 320)

    def load_icon_image(self, icon_name: str = None, size: tuple = (80, 80)):
        """Load icon image with black foreground and white background."""
        project_root = Path(__file__).parent.parent
        icons_dir = project_root / "icons"

        def convert_to_black(img):
            """Convert icon to deep blue color on white background.
            White foreground icons become deep blue, dark foreground icons become black.
            """
            if not HAS_PIL:
                return img

            # Convert to RGBA if needed
            if img.mode != 'RGBA':
                img = img.convert('RGBA')

            original_pixels = img.load()

            # Convert to grayscale to detect foreground vs background
            gray = img.convert('L')
            gray_pixels = gray.load()

            # Detect if icon is white foreground by checking average intensity
            total_intensity = 0
            pixel_count = 0
            for x in range(img.size[0]):
                for y in range(img.size[1]):
                    alpha = original_pixels[x, y][3]
                    if alpha >= 30:  # Any visible pixel
                        total_intensity += gray_pixels[x, y]
                        pixel_count += 1

            # If average intensity > 140, likely white foreground - convert to deep blue
            avg_intensity = total_intensity / pixel_count if pixel_count > 0 else 128
            is_white_foreground = avg_intensity > 140

            # Deep blue color: RGB(0, 82, 204) or similar
            DEEP_BLUE = (0, 82, 204)

            # Create colored version
            colored_img = Image.new('RGBA', img.size)
            colored_pixels = colored_img.load()

            # Convert: preserve shape (alpha), convert color
            for x in range(img.size[0]):
                for y in range(img.size[1]):
                    r, g, b, alpha = original_pixels[x, y]
                    intensity = gray_pixels[x, y]

                    # Skip fully transparent pixels
                    if alpha < 5:
                        colored_pixels[x, y] = (0, 0, 0, 0)
                        continue

                    if is_white_foreground:
                        # White foreground icon: bright areas are the icon shape
                        # Convert to deep blue with high opacity
                        # Use intensity to determine opacity (bright = more opaque)
                        opacity = int(255 * (intensity / 255.0))
                        # Ensure minimum opacity for visibility
                        if opacity > 20:  # Any visible brightness
                            opacity = max(200, opacity)  # Make it very opaque
                            colored_pixels[x, y] = (*DEEP_BLUE, opacity)
                        else:
                            colored_pixels[x, y] = (0, 0, 0, 0)
                    else:
                        # Dark foreground icon: dark areas are the icon shape
                        # Convert to black with opacity based on how dark
                        opacity = int(255 * ((255 - intensity) / 255.0))
                        # Ensure minimum opacity for visibility
                        if opacity > 20:  # Any visible darkness
                            opacity = max(200, opacity)  # Make it very opaque
                            colored_pixels[x, y] = (0, 0, 0, opacity)
                        else:
                            colored_pixels[x, y] = (0, 0, 0, 0)

            return colored_img

        if icon_name:
            # First, try to use mapping file
            if icon_name in self.icon_mapping:
                mapped_file = self.icon_mapping[icon_name]
                icon_path = Path(mapped_file.get('path', ''))
                if icon_path.exists():
                    try:
                        if HAS_PIL:
                            img = Image.open(icon_path)
                            if img.mode != 'RGBA':
                                img = img.convert('RGBA')

                            # Convert to black color
                            img = convert_to_black(img)

                            # Create white background
                            white_bg = Image.new('RGB', size, color='white')

                            # Resize icon
                            icon_resized = img.resize(size, Image.LANCZOS)

                            # Paste icon on white background
                            if icon_resized.mode == 'RGBA':
                                white_bg.paste(icon_resized, (0, 0),
                                               icon_resized)
                            else:
                                white_bg.paste(icon_resized, (0, 0))

                            return ImageTk.PhotoImage(white_bg)
                    except Exception as e:
                        # Debug: print error (can be removed later)
                        print(
                            f"Error loading icon from mapping ({icon_name}): {e}"
                        )
                        pass

            # Fallback: Try multiple naming patterns for icons directory
            icon_patterns = [
                f"{icon_name}_normal.png",  # light_normal.png
                f"{icon_name}_Normal.png",  # light_Normal.png (actual pattern)
                f"{icon_name}.png",  # light.png
            ]

            for pattern in icon_patterns:
                icon_path = icons_dir / pattern
                if icon_path.exists():
                    try:
                        if HAS_PIL:
                            img = Image.open(icon_path)
                            # Convert to RGBA if needed
                            if img.mode != 'RGBA':
                                img = img.convert('RGBA')

                            # Convert to black color
                            img = convert_to_black(img)

                            # Create white background
                            white_bg = Image.new('RGB', size, color='white')

                            # Resize icon
                            icon_resized = img.resize(size, Image.LANCZOS)

                            # Paste icon on white background
                            if icon_resized.mode == 'RGBA':
                                white_bg.paste(icon_resized, (0, 0),
                                               icon_resized)
                            else:
                                white_bg.paste(icon_resized, (0, 0))

                            return ImageTk.PhotoImage(white_bg)
                    except Exception as e:
                        # Debug: print error (can be removed later)
                        print(
                            f"Error loading icon pattern {pattern} ({icon_name}): {e}"
                        )
                        continue

            # Try to load specific icon from extracted icons
            icon_path = project_root / "extracted_icons" / "icons_by_name" / icon_name / f"{icon_name}.png"
            if icon_path.exists():
                try:
                    if HAS_PIL:
                        img = Image.open(icon_path)
                        if img.mode != 'RGBA':
                            img = img.convert('RGBA')

                        # Convert to black color
                        img = convert_to_black(img)

                        white_bg = Image.new('RGB', size, color='white')
                        icon_resized = img.resize(size, Image.LANCZOS)

                        if icon_resized.mode == 'RGBA':
                            white_bg.paste(icon_resized, (0, 0), icon_resized)
                        else:
                            white_bg.paste(icon_resized, (0, 0))

                        return ImageTk.PhotoImage(white_bg)
                except Exception as e:
                    # Debug: print error (can be removed later)
                    print(
                        f"Error loading icon from extracted_icons ({icon_name}): {e}"
                    )
                    pass

        # Use default icon (neutrals_normal.png) with black color
        if self.default_icon_path.exists():
            try:
                if HAS_PIL:
                    img = Image.open(self.default_icon_path)
                    if img.mode != 'RGBA':
                        img = img.convert('RGBA')

                    # Convert to black color
                    img = convert_to_black(img)

                    white_bg = Image.new('RGB', size, color='white')
                    icon_resized = img.resize(size, Image.LANCZOS)

                    if icon_resized.mode == 'RGBA':
                        white_bg.paste(icon_resized, (0, 0), icon_resized)
                    else:
                        white_bg.paste(icon_resized, (0, 0))

                    return ImageTk.PhotoImage(white_bg)
            except Exception as e:
                # Debug: print error (can be removed later)
                print(f"Error loading default icon ({icon_name}): {e}")
                pass

        # Fallback: create a white square with black border
        if HAS_PIL:
            img = Image.new('RGB', size, color='white')
            return ImageTk.PhotoImage(img)

    def load_locomotive_image(self,
                              image_name: str = None,
                              size: tuple = (227, 94)):
        """Load locomotive image from Z21 ZIP file.
        
        Args:
            image_name: Name of the image file (UUID-based)
            size: Target size in pixels (width, height). Default is 6cm x 2.5cm (227px x 94px at 96 DPI)
        """
        if not image_name or not HAS_PIL:
            return None

        try:
            import zipfile

            # Open the Z21 file as ZIP
            with zipfile.ZipFile(self.z21_file, 'r') as zf:
                # Search for the image file in the ZIP
                image_path = None
                for filename in zf.namelist():
                    # Match by exact filename or UUID part
                    if image_name in filename or filename.endswith(image_name):
                        image_path = filename
                        break

                if image_path:
                    # Extract image data
                    image_data = zf.read(image_path)

                    # Load image with PIL
                    from io import BytesIO
                    img = Image.open(BytesIO(image_data))

                    # Resize while maintaining aspect ratio
                    img.thumbnail(size, Image.LANCZOS)

                    # Create a white background image
                    bg_img = Image.new('RGB', size, color='white')

                    # Center the image on white background
                    x_offset = (size[0] - img.size[0]) // 2
                    y_offset = (size[1] - img.size[1]) // 2
                    bg_img.paste(img, (x_offset, y_offset),
                                 img if img.mode == 'RGBA' else None)

                    # Convert to PhotoImage
                    return ImageTk.PhotoImage(bg_img)
        except Exception as e:
            # Silently fail if image can't be loaded
            pass

        return None

        return None

    def create_function_card(self, func_num: int, func_info):
        """Create a card widget for a function with consistent sizing and alignment."""
        # Fixed card dimensions for consistent alignment
        CARD_WIDTH = 100
        ICON_SIZE = 80
        CARD_PADDING = 5

        # Create card frame with fixed width
        card_frame = tk.Frame(self.functions_frame_inner,
                              relief=tk.RAISED,
                              borderwidth=2,
                              bg='white',
                              width=CARD_WIDTH)
        card_frame.pack_propagate(False)  # Prevent frame from resizing
        # Note: Don't pack here, will be placed in grid by caller

        # Use grid layout for precise alignment
        # Row 0: Icon (centered)
        icon_frame = tk.Frame(card_frame,
                              width=ICON_SIZE,
                              height=ICON_SIZE,
                              bg='white')
        icon_frame.grid(row=0,
                        column=0,
                        padx=CARD_PADDING,
                        pady=(CARD_PADDING, 2),
                        sticky='')
        icon_frame.pack_propagate(False)

        # Load and display icon with black color on white background
        icon_image = self.load_icon_image(func_info.image_name,
                                          (ICON_SIZE, ICON_SIZE))
        if icon_image:
            icon_label = tk.Label(icon_frame, image=icon_image, bg='white')
            icon_label.image = icon_image  # Keep a reference
            icon_label.pack(expand=True)
        else:
            # Fallback: show a visible placeholder with icon name
            # Create a canvas to draw a border and text
            fallback_canvas = tk.Canvas(icon_frame,
                                        width=ICON_SIZE,
                                        height=ICON_SIZE,
                                        bg='white',
                                        highlightthickness=0)
            fallback_canvas.pack(fill=tk.BOTH, expand=True)
            # Draw a black border rectangle
            fallback_canvas.create_rectangle(5,
                                             5,
                                             ICON_SIZE - 5,
                                             ICON_SIZE - 5,
                                             outline='#000000',
                                             width=2)
            # Draw icon name text (truncated if too long)
            icon_name_short = func_info.image_name[:
                                                   8] if func_info.image_name else "?"
            fallback_canvas.create_text(ICON_SIZE // 2,
                                        ICON_SIZE // 2,
                                        text=icon_name_short,
                                        fill='#666666',
                                        font=('Arial', 8))

        # Row 1: Function number (always present, centered)
        func_num_label = tk.Label(
            card_frame,
            text=f"F{func_num}",
            font=('Arial', 11, 'bold'),
            bg='white',
            fg='#333333'  # Dark gray
        )
        func_num_label.grid(row=1, column=0, pady=(0, 2), sticky='')

        # Row 2: Shortcut (always present, show placeholder if empty, centered)
        shortcut_text = func_info.shortcut if func_info.shortcut else "—"
        shortcut_label = tk.Label(
            card_frame,
            text=shortcut_text,
            font=('Arial', 9, 'bold'),
            bg='white',
            fg='#0066CC' if func_info.shortcut else '#CCCCCC')
        shortcut_label.grid(row=2, column=0, pady=(0, 2), sticky='')

        # Row 3: Button type and duration on the same line (always present, centered)
        button_type_colors = {
            'switch': '#4CAF50',
            'push-button': '#FF9800',
            'time button': '#2196F3'
        }
        btn_color = button_type_colors.get(func_info.button_type_name(),
                                           '#666666')

        # Create a frame to hold button type and time on the same line
        type_time_frame = tk.Frame(card_frame, bg='white')
        type_time_frame.grid(row=3,
                             column=0,
                             pady=(0, CARD_PADDING),
                             sticky='')

        # Button type
        button_type_label = tk.Label(type_time_frame,
                                     text=func_info.button_type_name(),
                                     font=('Arial', 8),
                                     bg='white',
                                     fg=btn_color)
        button_type_label.pack(side=tk.LEFT)

        # Time indicator (if available)
        if func_info.time and func_info.time != "0":
            time_label = tk.Label(type_time_frame,
                                  text=f" ⏱ {func_info.time}s",
                                  font=('Arial', 7),
                                  bg='white',
                                  fg='#666666')
            time_label.pack(side=tk.LEFT)

        # Configure column to center all elements
        card_frame.grid_columnconfigure(0, weight=1)

        return card_frame  # Return card frame for grid placement


def main():
    import argparse

    parser = argparse.ArgumentParser(description='Z21 Locomotive Browser GUI')
    parser.add_argument('file',
                        type=Path,
                        nargs='?',
                        default=Path('z21_new.z21'),
                        help='Z21 file to open (default: z21_new.z21)')

    args = parser.parse_args()

    if not args.file.exists():
        print(f"Error: File not found: {args.file}")
        print(f"Usage: {sys.argv[0]} <z21_file>")
        sys.exit(1)

    # Suppress macOS TSM warning messages
    import os
    os.environ['PYTHONUNBUFFERED'] = '1'
    # Redirect stderr to suppress TSM messages on macOS
    if sys.platform == 'darwin':
        # Save original stderr
        original_stderr = sys.stderr

        # Create a filter that removes TSM messages
        class TSMFilter:

            def __init__(self, original):
                self.original = original

            def write(self, text):
                if 'TSM AdjustCapsLockLED' not in text and 'TSM' not in text:
                    self.original.write(text)

            def flush(self):
                self.original.flush()

        sys.stderr = TSMFilter(original_stderr)

    root = tk.Tk()
    app = Z21GUI(root, args.file)
    root.mainloop()


if __name__ == '__main__':
    main()
