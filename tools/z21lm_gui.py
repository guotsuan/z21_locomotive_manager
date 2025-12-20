#!/usr/bin/env python3
"""
GUI application to browse Z21 locomotives and their details.
"""

import sys
import os
import warnings

# Suppress macOS-specific warnings that don't affect functionality
if sys.platform == 'darwin':
    # Suppress IMKCFRunLoopWakeUpReliable mach port warnings
    os.environ['PYTHONWARNINGS'] = 'ignore'
    # Filter specific macOS warnings
    warnings.filterwarnings('ignore',
                            category=RuntimeWarning,
                            message='.*IMKCFRunLoopWakeUpReliable.*')
    warnings.filterwarnings('ignore',
                            category=RuntimeWarning,
                            message='.*NSOpenPanel.*overrides.*method.*')
    warnings.filterwarnings(
        'ignore', message='.*The class.*NSOpenPanel.*overrides the method.*')

import customtkinter as ctk
from tkinter import messagebox, filedialog, scrolledtext
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
from tools.z21lm_gui_operations import Z21GUIOperationsMixin


class Z21GUI(Z21GUIOperationsMixin):
    """Main GUI application for browsing Z21 locomotives."""

    def __init__(self, root, z21_file: Path):
        self.root = root
        self.z21_file = z21_file
        self.parser: Optional[Z21Parser] = None
        self.z21_data: Optional[Z21File] = None
        self.current_loco: Optional[Locomotive] = None
        self.current_loco_index: Optional[int] = None
        self.current_filtered_index: Optional[
            int] = None  # Track selected index in filtered list
        self.original_loco_address: Optional[int] = None
        self.user_selected_loco: Optional[Locomotive] = None
        self.default_icon_path = Path(
            __file__).parent.parent / "icons" / "neutrals_normal.png"
        self.icon_cache = {}
        self.icon_mapping = self.load_icon_mapping()
        self.status_timeout_id = None
        self.default_status_text = "Loading..."
        self._mouse_over_function_icon = False
        self._mouse_leave_timeout_id = None  # Track mouse leave timeout to cancel if needed
        self._cursor_change_timeout_id = None  # Track cursor change timeout to prevent flicker
        self.selected_function_num = None  # Track selected function icon
        self.delete_function_button = None  # Reference to delete button
        self.function_card_frames = {
        }  # Store card frames for selection highlighting
        self.edit_function_dialog = None  # Track edit function dialog window
        self.setup_ui()
        self.load_data()

    def _set_mouse_over_function_icon(self, value: bool):
        """Set mouse over function icon flag and clear timeout ID."""
        self._mouse_over_function_icon = value
        self._mouse_leave_timeout_id = None

    def set_status_message(self, message: str, timeout: int = 5000):
        """Set status bar message and clear it after timeout (default 5 seconds)."""
        if self.status_timeout_id is not None:
            self.root.after_cancel(self.status_timeout_id)
            self.status_timeout_id = None
        self.status_label.configure(text=message)
        self.status_timeout_id = self.root.after(
            timeout,
            lambda: self.status_label.configure(text=self.default_status_text))

    def load_icon_mapping(self):
        """Load icon mapping from JSON file."""
        mapping_file = Path(__file__).parent.parent / "icon_mapping.json"
        if mapping_file.exists():
            try:
                with open(mapping_file, "r") as f:
                    data = json.load(f)
                    return data.get("matches", {})
            except Exception:
                return {}
        return {}

    def update_status_count(self):
        """Update the default status text with current locomotive count."""
        if self.z21_data:
            self.default_status_text = f"Loaded {len(self.z21_data.locomotives)} locomotives"
        else:
            self.default_status_text = "No data loaded"

    def setup_ui(self):
        """Set up the user interface."""
        self.root.title("Z21 Locomotive Manager")
        # Set initial window size and position
        self.root.update_idletasks()  # Ensure window is ready
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        # Set window width to 1020 (requested by user: left 340 + right 680 = 1020)
        # Note: Actual window width will be 1035 to accommodate sash (5) and padding (10)
        # But we use 1020 as the base and let the system handle the layout
        window_width = 1020
        # Set window height to 75% of screen height (min 600, max 1000)
        window_height = max(600, min(1000, int(screen_height * 0.75)))
        # Center the window on screen
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        # Set minimum window size
        # Allow window to shrink smaller: left panel (200) + right panel (200) + sash + padding ≈ 420
        # Set to 600 to allow more flexibility while ensuring usability
        self.root.minsize(600, 500)

        # Track window resize for function layout recalculation
        self._resize_timeout_id = None
        self._last_window_width = None
        self._last_window_height = None
        self.root.bind("<Configure>", self.on_window_resize)

        # Create main container with resizable paned window
        from tkinter import PanedWindow
        main_paned = PanedWindow(self.root,
                                 orient="horizontal",
                                 sashwidth=5,
                                 sashrelief="raised")
        main_paned.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Store reference to main_paned for resize handling
        self.main_paned = main_paned
        
        # Bind to sash movement to update Functions tab layout when panel sizes change
        def on_sash_moved(event):
            # Only update if we're on Functions tab
            if not (hasattr(self, 'notebook') and self.notebook.get() == "Functions" and self.current_loco):
                return
            # Debounce: cancel any pending update
            if hasattr(self, "_sash_update_timeout_id") and self._sash_update_timeout_id:
                self.root.after_cancel(self._sash_update_timeout_id)
            # Schedule update after a delay to avoid flickering
            def update_after_sash():
                if hasattr(self, "_cached_canvas_width"):
                    delattr(self, "_cached_canvas_width")
                if hasattr(self, "_cached_cols"):
                    delattr(self, "_cached_cols")
                if hasattr(self, 'notebook') and self.notebook.get() == "Functions" and self.current_loco:
                    self.update_functions(is_resize=True)
                self._sash_update_timeout_id = None
            self._sash_update_timeout_id = self.root.after(100, update_after_sash)
        
        main_paned.bind("<ButtonRelease-1>", on_sash_moved)

        # Left panel: Locomotive list
        left_frame = ctk.CTkFrame(main_paned)
        main_paned.add(left_frame, minsize=200, width=340)

        # Search box
        search_frame = ctk.CTkFrame(left_frame)
        search_frame.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(search_frame, text="Search:").pack(side="left", padx=5)
        self.search_var = ctk.StringVar()
        self.search_var.trace("w", lambda *args: self.on_search())
        search_entry = ctk.CTkEntry(search_frame,
                                    textvariable=self.search_var,
                                    width=140)
        search_entry.pack(side="left", fill="x", expand=True, padx=5)

        # Button container for New, Delete, and Import buttons
        button_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        button_frame.pack(fill="x", padx=5, pady=(0, 5))

        # Configure columns for equal distribution of 3 buttons
        for i in range(3):
            button_frame.grid_columnconfigure(i, weight=1, uniform="buttons")

        ctk.CTkButton(button_frame,
                      text="Import",
                      command=self.import_z21_loco).grid(row=0,
                                                         column=0,
                                                         padx=(0, 5),
                                                         pady=0,
                                                         sticky="ew")
        ctk.CTkButton(button_frame,
                      text="Delete",
                      command=self.delete_selected_locomotive).grid(
                          row=0, column=1, padx=(0, 5), pady=0, sticky="ew")
        ctk.CTkButton(button_frame,
                      text="New",
                      command=self.create_new_locomotive).grid(row=0,
                                                               column=2,
                                                               padx=(0, 0),
                                                               pady=0,
                                                               sticky="ew")

        # Locomotive list
        list_frame = ctk.CTkFrame(left_frame)
        list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        ctk.CTkLabel(list_frame,
                     text="Locomotives:",
                     font=ctk.CTkFont(size=12,
                                      weight="bold")).pack(anchor="w",
                                                           pady=(0, 5))

        self.loco_listbox_frame = ctk.CTkScrollableFrame(list_frame)
        self.loco_listbox_frame.pack(fill="both", expand=True)
        self.loco_listbox_buttons = []

        # Bind keyboard events for navigation
        self.loco_listbox_frame.bind("<Up>", self.on_arrow_up)
        self.loco_listbox_frame.bind("<Down>", self.on_arrow_down)
        self.loco_listbox_frame.bind("<KeyPress-Up>", self.on_arrow_up)
        self.loco_listbox_frame.bind("<KeyPress-Down>", self.on_arrow_down)
        # Make the frame focusable
        self.loco_listbox_frame.bind(
            "<Button-1>", lambda e: self.loco_listbox_frame.focus_set())
        # Also bind to root for global keyboard navigation
        self.root.bind(
            "<Up>", lambda e: self.on_arrow_up(e)
            if self.is_list_focused() else None)
        self.root.bind(
            "<Down>", lambda e: self.on_arrow_down(e)
            if self.is_list_focused() else None)

        # Status label
        self.status_label = ctk.CTkLabel(left_frame, text="Loading...")
        self.status_label.pack(fill="x", padx=5, pady=5)

        # Right panel: Details
        right_frame = ctk.CTkFrame(main_paned)
        # Set right panel width to 680, allow it to shrink smaller
        main_paned.add(right_frame, minsize=200, width=680)

        # Details notebook (tabs)
        self.notebook = ctk.CTkTabview(right_frame)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        self.overview_frame = self.notebook.add("Overview")
        self.setup_overview_tab()

        self.functions_frame = self.notebook.add("Functions")
        self.setup_functions_tab()
        
        # Track if Functions tab has been shown for the first time
        self._functions_tab_shown = False

    def setup_overview_tab(self):
        """Set up the overview tab."""
        scrollable_frame = ctk.CTkScrollableFrame(self.overview_frame,
                                                  fg_color="transparent")
        scrollable_frame.pack(fill="both", expand=True)
        self.overview_scrollable_frame = scrollable_frame

        # Top frame for editable locomotive details
        details_frame = ctk.CTkFrame(scrollable_frame)
        details_frame.pack(fill="x", padx=5, pady=5)

        # Row 0: Image panel
        self.loco_image_label = ctk.CTkLabel(details_frame,
                                             text="No Image",
                                             anchor="center")
        self.loco_image_label.grid(row=0,
                                   column=1,
                                   columnspan=4,
                                   padx=(0, 10),
                                   pady=5,
                                   sticky="ew")
        self.loco_image_label.image = None
        self.loco_image_label.bind("<Button-1>", self.on_image_click)

        # Row 1: Name and Address
        ctk.CTkLabel(details_frame, text="Name:", width=7,
                     anchor="e").grid(row=1,
                                      column=0,
                                      padx=(5, 9),
                                      pady=2,
                                      sticky="e")
        self.name_var = ctk.StringVar()
        self.name_entry = ctk.CTkEntry(details_frame,
                                       textvariable=self.name_var)
        self.name_entry.grid(row=1, column=1, padx=(1, 3), pady=2, sticky="ew")

        ctk.CTkLabel(details_frame, text="Address:", width=50,
                     anchor="e").grid(row=1,
                                      column=2,
                                      padx=(3, 1),
                                      pady=2,
                                      sticky="e")
        self.address_var = ctk.StringVar()
        self.address_entry = ctk.CTkEntry(details_frame,
                                          textvariable=self.address_var)
        self.address_entry.grid(row=1,
                                column=3,
                                padx=(1, 5),
                                pady=2,
                                sticky="ew")

        # Row 2: Max Speed and Direction
        ctk.CTkLabel(details_frame, text="Max Speed:", width=7,
                     anchor="e").grid(row=2,
                                      column=0,
                                      padx=(5, 9),
                                      pady=2,
                                      sticky="e")
        self.speed_var = ctk.StringVar()
        self.speed_entry = ctk.CTkEntry(details_frame,
                                        textvariable=self.speed_var)
        self.speed_entry.grid(row=2,
                              column=1,
                              padx=(1, 3),
                              pady=2,
                              sticky="ew")

        ctk.CTkLabel(details_frame, text="Direction:", width=50,
                     anchor="e").grid(row=2,
                                      column=2,
                                      padx=(3, 1),
                                      pady=2,
                                      sticky="e")
        self.direction_var = ctk.StringVar()
        self.direction_combo = ctk.CTkComboBox(details_frame,
                                               variable=self.direction_var,
                                               values=["Forward", "Reverse"],
                                               state="readonly",
                                               width=18)
        self.direction_combo.grid(row=2,
                                  column=3,
                                  padx=(1, 5),
                                  pady=2,
                                  sticky="ew")

        # Additional Information Section
        row = 3
        separator = ctk.CTkFrame(details_frame,
                                 height=2,
                                 fg_color=("gray70", "gray30"))
        separator.grid(row=row,
                       column=0,
                       columnspan=6,
                       sticky="ew",
                       padx=5,
                       pady=5)
        row += 1

        ctk.CTkLabel(details_frame, text="Full Name:", width=7,
                     anchor="e").grid(row=row,
                                      column=0,
                                      padx=(5, 1),
                                      pady=2,
                                      sticky="e")
        self.full_name_var = ctk.StringVar()
        self.full_name_entry = ctk.CTkEntry(details_frame,
                                            textvariable=self.full_name_var,
                                            width=28)
        self.full_name_entry.grid(row=row,
                                  column=1,
                                  columnspan=4,
                                  padx=(1, 5),
                                  pady=2,
                                  sticky="ew")
        row += 1

        # Detailed fields
        fields = [
            ("Railway:", "railway_var", "Article Number:",
             "article_number_var"),
            ("Decoder Type:", "decoder_type_var", "Build Year:",
             "build_year_var"),
            ("Buffer Length:", "model_buffer_length_var", "Service Weight:",
             "service_weight_var"),
            ("Model Weight:", "model_weight_var", "Minimum Radius:",
             "rmin_var"),
            ("IP Address:", "ip_var", "Driver's Cab:", "drivers_cab_var"),
        ]

        for label1, var1, label2, var2 in fields:
            ctk.CTkLabel(details_frame, text=label1, width=7,
                         anchor="e").grid(row=row,
                                          column=0,
                                          padx=(5, 1),
                                          pady=2,
                                          sticky="e")
            setattr(self, var1, ctk.StringVar())
            ctk.CTkEntry(details_frame,
                         textvariable=getattr(self, var1)).grid(row=row,
                                                                column=1,
                                                                padx=(1, 3),
                                                                pady=2,
                                                                sticky="ew")

            ctk.CTkLabel(details_frame, text=label2, width=7,
                         anchor="e").grid(row=row,
                                          column=2,
                                          padx=(3, 1),
                                          pady=2,
                                          sticky="e")
            setattr(self, var2, ctk.StringVar())
            ctk.CTkEntry(details_frame,
                         textvariable=getattr(self, var2)).grid(row=row,
                                                                column=3,
                                                                padx=(1, 5),
                                                                pady=2,
                                                                sticky="ew")
            row += 1

        # Checkboxes and Speed Display
        checkbox_frame = ctk.CTkFrame(details_frame, fg_color="transparent")
        checkbox_frame.grid(row=row,
                            column=1,
                            sticky="ew",
                            padx=(1, 3),
                            pady=2)

        self.active_var = ctk.BooleanVar()
        self.active_checkbox = ctk.CTkCheckBox(checkbox_frame,
                                               text="Active",
                                               variable=self.active_var)
        self.active_checkbox.pack(side="left", padx=(0, 5))

        self.crane_var = ctk.BooleanVar()
        self.crane_checkbox = ctk.CTkCheckBox(checkbox_frame,
                                              text="Crane",
                                              variable=self.crane_var)
        self.crane_checkbox.pack(side="right")

        ctk.CTkLabel(details_frame, text="Speed Display:", width=7,
                     anchor="e").grid(row=row,
                                      column=2,
                                      padx=(3, 1),
                                      pady=2,
                                      sticky="e")
        self.speed_display_var = ctk.StringVar()
        self.speed_display_combo = ctk.CTkComboBox(
            details_frame,
            variable=self.speed_display_var,
            values=["km/h", "Regulation Step", "mph"],
            state="readonly",
            width=7)
        self.speed_display_combo.grid(row=row,
                                      column=3,
                                      padx=(1, 5),
                                      pady=2,
                                      sticky="ew")
        row += 1

        # Vehicle Type and Reg Step
        ctk.CTkLabel(details_frame, text="Vehicle Type:", width=7,
                     anchor="e").grid(row=row,
                                      column=0,
                                      padx=(5, 1),
                                      pady=2,
                                      sticky="e")
        self.rail_vehicle_type_var = ctk.StringVar()
        self.rail_vehicle_type_combo = ctk.CTkComboBox(
            details_frame,
            variable=self.rail_vehicle_type_var,
            values=["Loco", "Wagon", "Accessory"],
            state="readonly",
            width=180)
        self.rail_vehicle_type_combo.grid(row=row,
                                          column=1,
                                          padx=(1, 3),
                                          pady=2,
                                          sticky="ew")

        ctk.CTkLabel(details_frame, text="Reg Step:", width=7,
                     anchor="e").grid(row=row,
                                      column=2,
                                      padx=(3, 1),
                                      pady=2,
                                      sticky="e")
        self.regulation_step_var = ctk.StringVar()
        self.regulation_step_combo = ctk.CTkComboBox(
            details_frame,
            variable=self.regulation_step_var,
            values=["128", "28", "14"],
            state="readonly",
            width=7)
        self.regulation_step_combo.grid(row=row,
                                        column=3,
                                        padx=(1, 5),
                                        pady=2,
                                        sticky="ew")
        row += 1

        # Categories and In Stock Since
        ctk.CTkLabel(details_frame, text="Categories:", width=7,
                     anchor="e").grid(row=row,
                                      column=0,
                                      padx=(5, 1),
                                      pady=2,
                                      sticky="e")
        self.categories_var = ctk.StringVar()
        self.categories_entry = ctk.CTkEntry(details_frame,
                                             textvariable=self.categories_var)
        self.categories_entry.grid(row=row,
                                   column=1,
                                   padx=(1, 3),
                                   pady=2,
                                   sticky="ew")

        ctk.CTkLabel(details_frame,
                     text="In Stock Since:",
                     width=7,
                     anchor="e").grid(row=row,
                                      column=2,
                                      padx=(3, 1),
                                      pady=2,
                                      sticky="e")
        self.in_stock_since_var = ctk.StringVar()
        self.in_stock_since_entry = ctk.CTkEntry(
            details_frame, textvariable=self.in_stock_since_var, width=7)
        self.in_stock_since_entry.grid(row=row,
                                       column=3,
                                       padx=(1, 5),
                                       pady=2,
                                       sticky="ew")
        row += 1

        # Description field
        ctk.CTkLabel(details_frame, text="Description: ", width=7,
                     anchor="e").grid(row=row,
                                      column=0,
                                      padx=(5, 1),
                                      pady=2,
                                      sticky="ne")
        self.description_text = ctk.CTkTextbox(details_frame,
                                               wrap="word",
                                               width=350,
                                               height=200,
                                               font=ctk.CTkFont(size=11))
        self.description_text.grid(row=row,
                                   column=1,
                                   columnspan=3,
                                   padx=(1, 5),
                                   pady=2,
                                   sticky="ew")
        row += 1

        details_frame.grid_columnconfigure(0, weight=0)
        details_frame.grid_columnconfigure(1, weight=1, uniform="input_group")
        details_frame.grid_columnconfigure(2, weight=0)
        details_frame.grid_columnconfigure(3, weight=1, uniform="input_group")

        # Action buttons - all in one row, evenly distributed in columns 1-3
        button_row = row
        # Create a button container that spans columns 1-3
        button_container = ctk.CTkFrame(details_frame, fg_color="transparent")
        button_container.grid(row=button_row,
                              column=1,
                              columnspan=3,
                              padx=(1, 5),
                              pady=5,
                              sticky="ew")

        # Configure button container columns for equal distribution of 4 buttons
        for i in range(4):
            button_container.grid_columnconfigure(i,
                                                  weight=1,
                                                  uniform="buttons")

        self.export_button = ctk.CTkButton(button_container,
                                           text="Export Z21 Loco",
                                           command=self.export_z21_loco)
        self.export_button.grid(row=0,
                                column=0,
                                padx=(0, 5),
                                pady=0,
                                sticky="ew")

        self.share_button = ctk.CTkButton(button_container,
                                          text="Share with WIFI",
                                          command=self.share_with_airdrop)
        self.share_button.grid(row=0,
                               column=1,
                               padx=(0, 5),
                               pady=0,
                               sticky="ew")

        self.scan_button = ctk.CTkButton(button_container,
                                         text="Import from JSON",
                                         command=self.scan_for_details)
        self.scan_button.grid(row=0,
                              column=2,
                              padx=(0, 5),
                              pady=0,
                              sticky="ew")

        self.save_button = ctk.CTkButton(button_container,
                                         text="Save Changes",
                                         command=self.save_locomotive_changes)
        self.save_button.grid(row=0,
                              column=3,
                              padx=(0, 0),
                              pady=0,
                              sticky="ew")
        row = button_row + 1

        # Overview text area - fills columns 1-3
        self.overview_text = ctk.CTkTextbox(details_frame,
                                            wrap="word",
                                            font=ctk.CTkFont(family="Courier",
                                                             size=12),
                                            state="disabled")
        self.overview_text.grid(row=row,
                                column=1,
                                columnspan=3,
                                padx=(1, 5),
                                pady=5,
                                sticky="nsew")

        # Configure row weight for overview text area to expand
        details_frame.grid_rowconfigure(row, weight=1)

        # Mousewheel binding logic for overview text area
        # This handler is bound directly to overview_text to prevent event bubbling
        def on_overview_text_mousewheel(event):
            """Handle mousewheel events on overview_text, preventing outer scroll."""
            try:
                if self.notebook.get() != "Overview":
                    return
            except:
                pass
            scroll_amount = 0
            if event.num == 4: scroll_amount = -5
            elif event.num == 5: scroll_amount = 5
            elif hasattr(event, "delta"):
                scroll_amount = -1 * (event.delta // 120)
                if scroll_amount == 0:
                    scroll_amount = -1 if event.delta > 0 else 1
            elif hasattr(event, "deltaY"):
                scroll_amount = -1 * (event.deltaY // 120)
                if scroll_amount == 0:
                    scroll_amount = -1 if event.deltaY > 0 else 1

            if scroll_amount != 0:
                try:
                    self.overview_text.yview_scroll(int(scroll_amount),
                                                    "units")
                except:
                    pass
            # Always return "break" to prevent event from bubbling to outer containers
            return "break"

        # Bind mousewheel events directly to overview_text to prevent outer scroll
        self.overview_text.bind("<MouseWheel>", on_overview_text_mousewheel)
        self.overview_text.bind("<Button-4>", on_overview_text_mousewheel)
        self.overview_text.bind("<Button-5>", on_overview_text_mousewheel)

        # Mousewheel binding logic for outer containers (only when NOT over overview_text)
        def on_overview_mousewheel(event):
            """Handle mousewheel events on outer containers, but skip if over overview_text."""
            try:
                if self.notebook.get() != "Overview":
                    return
            except:
                pass
            # Skip if mouse is over overview_text (let it handle its own scrolling)
            try:
                if self.overview_text.winfo_containing(event.x_root,
                                                       event.y_root):
                    return  # Let overview_text handle it, don't scroll outer container
            except:
                pass

            # For outer containers, allow normal scrolling behavior
            # (This will scroll the scrollable_frame if needed)
            return

        scrollable_frame.bind("<MouseWheel>", on_overview_mousewheel, add="+")
        scrollable_frame.bind("<Button-4>", on_overview_mousewheel, add="+")
        scrollable_frame.bind("<Button-5>", on_overview_mousewheel, add="+")
        self.overview_frame.bind("<MouseWheel>",
                                 on_overview_mousewheel,
                                 add="+")
        self.overview_frame.bind("<Button-4>", on_overview_mousewheel, add="+")
        self.overview_frame.bind("<Button-5>", on_overview_mousewheel, add="+")

        # Remove global bindings for Overview tab to prevent conflicts
        # (The frame-level bindings above should be sufficient)

    def setup_functions_tab(self):
        """Set up the functions tab."""
        scrollable_frame = ctk.CTkScrollableFrame(self.functions_frame,
                                                  fg_color="transparent")
        scrollable_frame.pack(fill="both", expand=True)
        self.functions_frame_inner = scrollable_frame

    def load_data(self):
        """Load Z21 file data."""
        self.status_label.configure(text="Loading data...")
        self.root.update()
        try:
            self.parser = Z21Parser(self.z21_file)
            self.z21_data = self.parser.parse()
            self.populate_list(auto_select_first=True)
            self.update_status_count()
            self.status_label.configure(text=self.default_status_text)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")
            self.set_status_message("Error loading file")

    def normalize_for_search(self, text: str) -> str:
        """Normalize text for fuzzy matching."""
        if not text:
            return ""
        normalized = text.lower()
        normalized = "".join(normalized.split())
        normalized = normalized.replace("-", "").replace("_",
                                                         "").replace(".", "")
        return normalized

    def populate_list(self,
                      filter_text: str = "",
                      preserve_selection: bool = False,
                      auto_select_first: bool = False):
        """Populate the locomotive list with fuzzy matching."""
        if not self.z21_data:
            return

        current_selection = None
        current_loco = self.current_loco if preserve_selection and self.current_loco else None

        for button in self.loco_listbox_buttons:
            button.destroy()
        self.loco_listbox_buttons = []
        self.filtered_locos = []

        filter_normalized = self.normalize_for_search(filter_text)
        for loco in self.z21_data.locomotives:
            display_text = f"Address {loco.address:4d} - {loco.name}"
            display_normalized = self.normalize_for_search(display_text)
            address_normalized = self.normalize_for_search(str(loco.address))
            name_normalized = self.normalize_for_search(loco.name)

            if not filter_text or (filter_normalized in display_normalized
                                   or filter_normalized in address_normalized
                                   or filter_normalized in name_normalized):
                button = ctk.CTkButton(
                    self.loco_listbox_frame,
                    text=display_text,
                    anchor="w",
                    text_color="black",
                    command=lambda idx=len(self.filtered_locos): self.
                    on_loco_button_click(idx),
                )
                button.pack(fill="x", padx=5, pady=2)
                self.loco_listbox_buttons.append(button)
                self.filtered_locos.append(loco)

        if preserve_selection and self.user_selected_loco:
            # Try to find the selected locomotive in the filtered list
            for i, loco in enumerate(self.filtered_locos):
                if loco.address == self.user_selected_loco.address and loco.name == self.user_selected_loco.name:
                    current_selection = i
                    break

        if current_selection is not None:
            self.current_filtered_index = current_selection
            self.highlight_button(current_selection)
            self.on_loco_select_by_index(current_selection)
        elif self.filtered_locos and auto_select_first:
            self.current_filtered_index = 0
            self.highlight_button(0)
            self.on_loco_select_by_index(0)

    def scroll_button_into_view(self, index: int):
        """Scroll the button at the given index into view if it's outside the visible area."""
        if index < 0 or index >= len(self.loco_listbox_buttons):
            return
        try:
            button = self.loco_listbox_buttons[index]
            self.root.update_idletasks()

            # Get the internal canvas of CTkScrollableFrame
            canvas = self.loco_listbox_frame._parent_canvas

            # Get button position relative to the scrollable frame
            button.update_idletasks()
            button_y = button.winfo_y()
            button_height = button.winfo_height()
            button_bottom = button_y + button_height

            # Get canvas dimensions
            canvas.update_idletasks()
            canvas_height = canvas.winfo_height()

            # Get current view position
            view_top_ratio, view_bottom_ratio = canvas.yview()

            # Get total scrollable height
            bbox = canvas.bbox("all")
            if not bbox:
                return
            total_height = max(bbox[3] - bbox[1], 1)

            # Calculate current visible area in pixels
            view_top_px = view_top_ratio * total_height
            view_bottom_px = view_bottom_ratio * total_height

            # Check if button is outside visible area and scroll if needed
            if button_y < view_top_px:
                # Button is above visible area, scroll up
                target_ratio = max(0.0, button_y / total_height)
                canvas.yview_moveto(target_ratio)
            elif button_bottom > view_bottom_px:
                # Button is below visible area, scroll down
                # Show button at the bottom of visible area
                target_ratio = max(
                    0.0,
                    min(1.0, (button_bottom - canvas_height) / total_height))
                canvas.yview_moveto(target_ratio)
        except Exception:
            pass  # Ignore scroll errors

    def highlight_button(self, index: int):
        """Highlight a button at the given index."""
        if index < 0 or index >= len(self.loco_listbox_buttons):
            return
        for i, button in enumerate(self.loco_listbox_buttons):
            if i == index:
                button.configure(
                    fg_color=("gray60", "gray40"))  # Darker color for selected
            else:
                button.configure(
                    fg_color=("gray85",
                              "gray18"))  # Lighter color for unselected
        self.current_filtered_index = index
        # Scroll the button into view if it's outside visible area
        self.scroll_button_into_view(index)

    def on_loco_button_click(self, index: int):
        self.highlight_button(index)
        self.on_loco_select_by_index(index)
        # Focus the list frame for keyboard navigation
        self.loco_listbox_frame.focus_set()

    def is_list_focused(self):
        """Check if the locomotive list has focus."""
        try:
            return self.loco_listbox_frame.focus_get(
            ) == self.loco_listbox_frame
        except:
            return False

    def on_arrow_up(self, event):
        """Handle up arrow key to navigate to previous locomotive."""
        if not self.filtered_locos:
            return
        if self.current_filtered_index is None:
            self.current_filtered_index = 0
        else:
            self.current_filtered_index = max(0,
                                              self.current_filtered_index - 1)
        self.highlight_button(self.current_filtered_index)
        self.on_loco_select_by_index(self.current_filtered_index)
        return "break"

    def on_arrow_down(self, event):
        """Handle down arrow key to navigate to next locomotive."""
        if not self.filtered_locos:
            return
        if self.current_filtered_index is None:
            self.current_filtered_index = 0
        else:
            self.current_filtered_index = min(
                len(self.filtered_locos) - 1, self.current_filtered_index + 1)
        self.highlight_button(self.current_filtered_index)
        self.on_loco_select_by_index(self.current_filtered_index)
        return "break"

    def on_loco_select_by_index(self, index: int):
        """Handle locomotive selection by index."""
        if index < len(self.filtered_locos):
            new_loco = self.filtered_locos[index]
            if self.current_loco is None or new_loco.address != self.current_loco.address or new_loco.name != self.current_loco.name:
                self.current_loco = new_loco
                self.original_loco_address = self.current_loco.address
                self.user_selected_loco = self.current_loco
                self.current_loco_index = None
                if self.z21_data:
                    for i, loco in enumerate(self.z21_data.locomotives):
                        if loco.address == self.current_loco.address and loco.name == self.current_loco.name:
                            self.current_loco_index = i
                            break
                self.update_details()

    def on_search(self, *args):
        filter_text = self.search_var.get()
        self.populate_list(filter_text, auto_select_first=True)

    def on_loco_select(self, event):
        pass

    def on_image_click(self, event):
        """Handle click on locomotive image to upload and crop new image."""
        if not self.current_loco:
            self.set_status_message("No locomotive selected.")
            return

        file_path = filedialog.askopenfilename(
            title="Select Locomotive Image",
            filetypes=[
                ("JPEG files", "*.jpg *.jpeg"),
                ("PNG files", "*.png"),
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff"),
                ("All files", "*.*")
            ],
        )
        if file_path:
            # Ensure file_path is a string (filedialog should return str, but be defensive)
            if not isinstance(file_path, str):
                file_path = str(file_path) if file_path else ""
            if file_path:
                self.open_image_crop_window(file_path)

    def update_details(self):
        """Update the details display."""
        if not self.current_loco:
            return
        self.update_overview()
        # Always update functions - it will handle first-time display correctly
        self.update_functions()

    def update_overview(self):
        """Update overview tab."""
        loco = self.current_loco
        self.name_var.set(loco.name)
        self.address_var.set(str(loco.address))
        self.speed_var.set(str(loco.speed))
        self.direction_var.set("Forward" if loco.direction else "Reverse")
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

        speed_display_map = {0: "km/h", 1: "Regulation Step", 2: "mph"}
        self.speed_display_var.set(
            speed_display_map.get(loco.speed_display, "km/h"))

        rail_type_map = {0: "Loco", 1: "Wagon", 2: "Accessory"}
        self.rail_vehicle_type_var.set(
            rail_type_map.get(loco.rail_vehicle_type, "Loco"))
        self.crane_var.set(loco.crane)

        regulation_step_map = {0: "128", 1: "28", 2: "14"}
        self.regulation_step_var.set(
            regulation_step_map.get(loco.regulation_step, "128"))

        self.categories_var.set(
            ", ".join(loco.categories) if loco.categories else "")
        self.in_stock_since_var.set(getattr(loco, "in_stock_since", "") or "")

        self.description_text.delete(1.0, "end")
        self.description_text.insert(1.0, loco.description)

        # Safely update image label with comprehensive error handling
        def safe_clear_image(label):
            """Safely clear image from label by accessing internal label directly."""
            try:
                # Clear our image reference
                if hasattr(label, 'image'):
                    label.image = None
                # Clear internal Tkinter label image directly
                try:
                    label._label.configure(image="")
                except Exception:
                    pass
            except Exception:
                pass

        def safe_configure_label(label, **kwargs):
            """Safely configure label, ignoring any image-related errors."""
            # First, always try to clear image before configuring
            if 'image' not in kwargs or kwargs.get('image') is None:
                safe_clear_image(label)

            try:
                label.configure(**kwargs)
            except Exception:
                # If configure fails, try to set only text via internal label
                if 'text' in kwargs:
                    try:
                        safe_clear_image(label)
                        label._text = kwargs['text']
                        if hasattr(label, '_draw'):
                            label._draw()
                    except Exception:
                        pass

        # Clear previous image reference and internal label image
        safe_clear_image(self.loco_image_label)

        # Now safely set new image or text
        if loco.image_name:
            loco_image = self.load_locomotive_image(loco.image_name,
                                                    size=(227, 94))
            if loco_image:
                try:
                    safe_configure_label(self.loco_image_label,
                                         image=loco_image,
                                         text="")
                    self.loco_image_label.image = loco_image
                except Exception:
                    # If setting image fails, just set text
                    try:
                        self.loco_image_label.image = None
                        safe_configure_label(self.loco_image_label,
                                             text=f"Image:\n{loco.image_name}")
                    except Exception:
                        pass
            else:
                try:
                    self.loco_image_label.image = None
                    safe_configure_label(self.loco_image_label,
                                         text=f"Image:\n{loco.image_name}")
                except Exception:
                    pass
        else:
            try:
                self.loco_image_label.image = None
                safe_configure_label(self.loco_image_label, text="No Image")
            except Exception:
                pass

        self.overview_text.configure(state="normal")
        self.overview_text.delete(1.0, "end")
        text = f"""
{'='*70}
FUNCTION SUMMARY
{'='*70}
Functions:         {len(loco.functions)} configured
Function Details:  {len(loco.function_details)} available
"""
        if loco.function_details:
            sorted_funcs = sorted(loco.function_details.items(),
                                  key=lambda x: x[1].function_number)
            text += "\n"
            for func_num, func_info in sorted_funcs:
                shortcut = f" [{func_info.shortcut}]" if func_info.shortcut else ""
                time_str = ""
                if func_info.button_type == 2 and func_info.time != "0":
                    time_str = f" (time: {func_info.time}s)"
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

        self.overview_text.insert(1.0, text)
        self.overview_text.configure(state="disabled")


    def update_functions(self, is_resize=False):
        """Update functions tab."""
        loco = self.current_loco
        if not loco:
            return
        
        # Check if this is the first time showing Functions tab
        is_first_show = not hasattr(self, "_functions_tab_shown") or not self._functions_tab_shown
        current_tab = self.notebook.get() if hasattr(self, 'notebook') else None
        
        # If first time showing Functions tab, ensure tab is active and wait for rendering
        if is_first_show and current_tab == "Functions":
            # Force tab to be visible and wait for rendering
            self.root.update_idletasks()
            self.functions_frame.update_idletasks()
            self.functions_frame_inner.update_idletasks()
            # Mark as shown
            self._functions_tab_shown = True
            # Clear any cached width to force recalculation
            if hasattr(self, "_cached_canvas_width"):
                delattr(self, "_cached_canvas_width")
            if hasattr(self, "_cached_cols"):
                delattr(self, "_cached_cols")
        
        # Calculate new column count BEFORE destroying widgets
        # This allows us to skip redraw if nothing changed
        self.root.update_idletasks()
        self.functions_frame.update_idletasks()
        self.functions_frame_inner.update_idletasks()
        
        # Get current width and calculate columns
        # Always get fresh width for resize operations to ensure accurate column count
        is_first_call = not hasattr(self, "_cached_canvas_width") or self._cached_canvas_width <= 100
        
        # For resize operations, always get fresh width to detect window size changes
        if is_resize:
            # Force fresh width measurement for resize
            current_canvas_width = self.functions_frame.winfo_width()
            if current_canvas_width <= 100:
                try:
                    canvas = self.functions_frame_inner._parent_canvas
                    canvas.update_idletasks()
                    canvas_width = canvas.winfo_width()
                    if canvas_width > 100:
                        current_canvas_width = canvas_width
                except:
                    pass
            if current_canvas_width <= 100:
                root_width = self.root.winfo_width()
                if root_width > 100:
                    left_panel_width = self.loco_listbox_frame.master.master.winfo_width() if hasattr(self, 'loco_listbox_frame') else 340
                    current_canvas_width = max(200, root_width - left_panel_width - 20)
                else:
                    current_canvas_width = 400
        elif hasattr(self, "_cached_canvas_width") and self._cached_canvas_width > 100:
            # For non-resize, try cached value but verify it's still valid
            actual_width = self.functions_frame.winfo_width()
            if actual_width > 100:
                current_canvas_width = actual_width
            else:
                current_canvas_width = self._cached_canvas_width
        else:
            # First call or no cache, get fresh width
            current_canvas_width = self.functions_frame.winfo_width()
            if current_canvas_width <= 100:
                try:
                    canvas = self.functions_frame_inner._parent_canvas
                    canvas.update_idletasks()
                    canvas_width = canvas.winfo_width()
                    if canvas_width > 100:
                        current_canvas_width = canvas_width
                except:
                    pass
            if current_canvas_width <= 100:
                root_width = self.root.winfo_width()
                if root_width > 100:
                    left_panel_width = self.loco_listbox_frame.master.master.winfo_width() if hasattr(self, 'loco_listbox_frame') else 340
                    current_canvas_width = max(200, root_width - left_panel_width - 20)
                else:
                    current_canvas_width = 400
        
        # Calculate column count based on available width
        # Use a minimum card width (100px) plus padding (4px total: 2px left + 2px right)
        # This is used only for calculating how many columns can fit
        min_card_width_with_padding = 104
        # Available width: total width minus frame padding and scrollbar space
        available_width = max(min_card_width_with_padding, current_canvas_width - 20)
        
        # Calculate how many complete columns can fit
        base_cols = available_width // min_card_width_with_padding
        
        # Calculate remaining space after fitting base columns
        remaining_space = available_width % min_card_width_with_padding
        
        # Determine final column count
        # If we have enough space for base_cols columns plus significant remaining space (>=50px),
        # we can fit one more column
        if remaining_space >= 50 and base_cols > 0:
            new_cols = base_cols + 1
        else:
            new_cols = base_cols
        
        # Ensure at least 1 column and reasonable maximum
        new_cols = max(1, new_cols)
        
        # Debug output (can be enabled for troubleshooting)
        # print(f"Cols calc: width={current_canvas_width}, available={available_width}, base={base_cols}, remaining={remaining_space}, final={new_cols}")
        
        # Check if column count actually changed OR if locomotive changed
        # We need to update even if column count is the same when locomotive changes
        old_cols = getattr(self, "_cached_cols", None)
        old_loco = getattr(self, "_cached_loco_for_functions", None)
        loco_changed = old_loco != loco
        
        # Skip redraw only if column count didn't change AND locomotive didn't change
        if old_cols == new_cols and not is_first_show and not is_resize and not loco_changed:
            # Nothing changed, skip redraw to prevent flickering
            return
        
        # Column count changed, locomotive changed, or first show - proceed with redraw
        for widget in self.functions_frame_inner.winfo_children():
            widget.destroy()
        
        # Use the already calculated values (no need to recalculate)
        cols = new_cols
        
        # Update cache including current locomotive reference
        self._cached_canvas_width = current_canvas_width
        self._cached_cols = cols
        self._cached_loco_for_functions = loco  # Store locomotive reference to detect changes

        header_label = ctk.CTkLabel(self.functions_frame_inner,
                                    text=f"Functions for {loco.name}",
                                    font=("Arial", 14, "bold"),
                                    anchor="w")
        header_label.grid(row=0,
                          column=0,
                          columnspan=cols,
                          sticky="w",
                          padx=5,
                          pady=(10, 5))

        button_frame = ctk.CTkFrame(self.functions_frame_inner,
                                    fg_color="transparent")
        button_frame.grid(row=1,
                          column=0,
                          columnspan=cols,
                          sticky="ew",
                          padx=5,
                          pady=(0, 10))
        self.functions_button_frame = button_frame  # Store reference for delete button placement
        ctk.CTkButton(button_frame,
                      text="+ Add New Function",
                      command=self.add_new_function).pack(side="left",
                                                          padx=(0, 10))
        ctk.CTkButton(button_frame,
                      text="📄 Scan from JSON",
                      command=self.scan_from_json).pack(side="left",
                                                        padx=(0, 10))
        ctk.CTkButton(button_frame,
                      text="💾 Save Changes",
                      command=self.save_function_changes).pack(side="left")

        if not loco.function_details:
            ctk.CTkLabel(self.functions_frame_inner,
                         text="No functions configured.",
                         font=("Arial", 11),
                         text_color="gray").grid(row=2,
                                                 column=0,
                                                 columnspan=cols,
                                                 sticky="ew",
                                                 padx=5,
                                                 pady=20)
            return

        # Clear previous selection and delete button
        self.selected_function_num = None
        self.function_card_frames = {}
        if self.delete_function_button:
            self.delete_function_button.destroy()
            self.delete_function_button = None

        # Configure grid columns BEFORE placing cards
        # Use weight=1 to allow columns to expand and fill available space evenly
        # No minsize constraint - let columns adapt to available width
        # This ensures columns can shrink when window is small (e.g., 3 columns)
        for i in range(cols):
            self.functions_frame_inner.grid_columnconfigure(i, weight=1)
        
        # Ensure columns beyond cols are not configured (in case cols decreased)
        # This prevents leftover column configurations from causing layout issues
        max_configured_cols = getattr(self, "_max_configured_cols", 0)
        if cols < max_configured_cols:
            # Clear configuration for columns that are no longer needed
            for i in range(cols, max_configured_cols):
                try:
                    self.functions_frame_inner.grid_columnconfigure(i, weight=0)
                except:
                    pass
        self._max_configured_cols = cols

        # Sort by position, then by function number (same as list_locomotives.py)
        sorted_funcs = sorted(loco.function_details.items(),
                              key=lambda x:
                              (x[1].position, x[1].function_number))
        row, col = 2, 0
        for func_num, func_info in sorted_funcs:
            card_frame = self.create_function_card(func_num, func_info)
            self.function_card_frames[
                func_num] = card_frame  # Store reference for selection highlighting

            def make_clickable(widget, fn, fi, card_ref):
                widget._click_pending = False

                def on_single_click(e):
                    # Mark that a click is pending
                    widget._click_pending = True

                    # Wait to see if it's a double-click (delay to allow double-click detection)
                    def do_select():
                        if widget._click_pending:
                            widget._click_pending = False
                            self.select_function(fn)

                    widget.after(300, do_select)

                def on_double_click(e):
                    # Cancel pending single click
                    widget._click_pending = False
                    # Double click: open edit window
                    self.edit_function(fn, fi)

                def on_enter(e):
                    # Cancel any pending mouse leave timeout
                    if self._mouse_leave_timeout_id is not None:
                        self.root.after_cancel(self._mouse_leave_timeout_id)
                        self._mouse_leave_timeout_id = None
                    # Cancel any pending cursor change to default
                    if self._cursor_change_timeout_id is not None:
                        self.root.after_cancel(self._cursor_change_timeout_id)
                        self._cursor_change_timeout_id = None
                    setattr(self, '_mouse_over_function_icon', True)
                    # Use debounced cursor change to prevent flicker when moving quickly
                    # Only change cursor after a short delay, avoiding rapid changes
                    def change_cursor():
                        try:
                            widget.configure(cursor="hand2")
                        except:
                            pass
                        self._cursor_change_timeout_id = None
                    # Delay cursor change to avoid flicker during rapid mouse movement
                    # Use after_idle to update when GUI is idle, reducing render impact
                    self._cursor_change_timeout_id = self.root.after_idle(change_cursor)

                def on_leave(e):
                    # Cancel any pending cursor change to hand
                    if self._cursor_change_timeout_id is not None:
                        self.root.after_cancel(self._cursor_change_timeout_id)
                        self._cursor_change_timeout_id = None
                    # Cancel any pending mouse leave timeout
                    if self._mouse_leave_timeout_id is not None:
                        self.root.after_cancel(self._mouse_leave_timeout_id)
                    # Reset cursor with a small delay to match enter behavior
                    def reset_cursor():
                        try:
                            widget.configure(cursor="")
                        except:
                            pass
                    self.root.after_idle(reset_cursor)
                    # Schedule mouse leave flag update with delay
                    self._mouse_leave_timeout_id = self.root.after(
                        100, lambda: self._set_mouse_over_function_icon(False))

                widget.bind("<Button-1>", on_single_click)
                widget.bind("<Double-Button-1>", on_double_click)
                widget.bind("<Enter>", on_enter)
                widget.bind("<Leave>", on_leave)

                for child in widget.winfo_children():
                    make_clickable(child, fn, fi, card_ref)
                    child._click_pending = False
                    child.bind("<Button-1>", on_single_click)
                    child.bind("<Double-Button-1>", on_double_click)

            make_clickable(card_frame, func_num, func_info, card_frame)
            # Use sticky="n" to allow horizontal centering but prevent vertical expansion
            # padx=2 ensures fixed spacing between columns (2px left + 2px right = 4px gap)
            # This keeps column spacing constant while allowing columns to adapt width
            card_frame.grid(row=row, column=col, padx=2, pady=5, sticky="n")
            col += 1
            if col >= cols:
                col = 0
                row += 1

    def select_function(self, func_num: int):
        """Select a function icon and show delete button."""
        # Clear previous selection visual feedback
        if self.selected_function_num is not None and self.selected_function_num in self.function_card_frames:
            prev_card = self.function_card_frames[self.selected_function_num]
            prev_card.configure(border_width=2,
                                border_color=("gray75", "gray25"))

        # Set new selection
        self.selected_function_num = func_num

        # Add visual feedback for selected card
        if func_num in self.function_card_frames:
            selected_card = self.function_card_frames[func_num]
            selected_card.configure(border_width=3,
                                    border_color=("blue", "light blue"))

        # Show or update delete button
        self.show_delete_button(func_num)

    def show_delete_button(self, func_num: int):
        """Show delete button for selected function."""
        # Remove existing delete button if any
        if self.delete_function_button:
            self.delete_function_button.destroy()
            self.delete_function_button = None

        # Create delete button in the button frame
        if hasattr(self, 'functions_button_frame'):
            self.delete_function_button = ctk.CTkButton(
                self.functions_button_frame,
                text="🗑️ Delete F" + str(func_num),
                command=lambda: self.delete_function(func_num),
                fg_color=("red4", "darkred"),
                hover_color=("red3", "red"))
            self.delete_function_button.pack(side="left", padx=(10, 0))

    def delete_function(self, func_num: int):
        """Delete the selected function."""
        if not self.current_loco:
            return

        if func_num not in self.current_loco.function_details:
            return

        # Confirm deletion
        func_info = self.current_loco.function_details[func_num]
        confirm_msg = f"Delete function F{func_num} ({func_info.image_name})?"
        confirmed = messagebox.askyesno("Confirm Delete", confirm_msg)
        # Return focus to main window after dialog closes
        self.root.focus_set()
        self.root.lift()
        if not confirmed:
            return

        # Delete function
        if func_num in self.current_loco.function_details:
            del self.current_loco.function_details[func_num]
        if func_num in self.current_loco.functions:
            del self.current_loco.functions[func_num]

        # Clear selection
        self.selected_function_num = None
        if self.delete_function_button:
            self.delete_function_button.destroy()
            self.delete_function_button = None

        # Refresh functions display
        self.update_functions()
        self.update_overview()
        # Switch focus back to Functions tab and return focus to main window
        if hasattr(self, 'notebook'):
            self.notebook.set("Functions")
        # Return focus to main window after deletion
        self.root.focus_set()
        self.root.lift()
        self.root.update()
        self.set_status_message(f"Function F{func_num} deleted successfully!")

    def recalculate_function_layout(self):
        """Recalculate function layout after window is fully rendered."""
        if hasattr(self, "_width_recalc_scheduled"):
            delattr(self, "_width_recalc_scheduled")
        # Force recalculation by clearing cache
        if hasattr(self, "_cached_canvas_width"):
            delattr(self, "_cached_canvas_width")
        if hasattr(self, "_cached_cols"):
            delattr(self, "_cached_cols")
        # Re-update functions to get correct layout
        if self.current_loco:
            self.update_functions()

    def on_window_resize(self, event):
        """Handle window resize events to recalculate function icon layout."""
        # Only handle root window resize events, not child widget events
        if event.widget != self.root:
            return

        # Check if window size actually changed
        current_width = self.root.winfo_width()
        current_height = self.root.winfo_height()

        # Ignore if size hasn't changed (might be just window movement)
        if (hasattr(self, "_last_window_width")
                and hasattr(self, "_last_window_height")
                and self._last_window_width == current_width
                and self._last_window_height == current_height):
            return

        # Update stored dimensions
        self._last_window_width = current_width
        self._last_window_height = current_height

        # Clear cached width so layout will recalculate with new window size
        # This ensures functions tab uses the updated right panel width
        if hasattr(self, "_cached_canvas_width"):
            delattr(self, "_cached_canvas_width")
        if hasattr(self, "_cached_cols"):
            delattr(self, "_cached_cols")

        # Use debouncing to avoid too frequent recalculations
        if hasattr(self, "_resize_timeout_id") and self._resize_timeout_id:
            self.root.after_cancel(self._resize_timeout_id)

        # Schedule recalculation after a short delay (debounce)
        self._resize_timeout_id = self.root.after(
            200, self._handle_resize_recalculation)

    def _handle_resize_recalculation(self):
        """Handle the actual recalculation after resize debounce delay."""
        self._resize_timeout_id = None
        # Only recalculate if we're on Functions tab and have a current locomotive
        if hasattr(self, 'notebook') and self.notebook.get(
        ) == "Functions" and self.current_loco:
            self.update_functions(
                is_resize=True)  # Mark as resize call to add extra column

    def save_function_changes(self):
        """Save all function changes to the Z21 file."""
        if not self.current_loco or not self.z21_data or not self.parser:
            self.set_status_message(
                "No locomotive selected or data not loaded.")
            return
        try:
            if self.current_loco_index is not None:
                self.z21_data.locomotives[
                    self.current_loco_index] = self.current_loco
            self.parser.write(self.z21_data, self.z21_file)
            self.set_status_message(
                "All function changes saved successfully to file!")
        except Exception as write_error:
            self.set_status_message(f"Failed to write changes: {write_error}")

    def get_next_unused_function_number(self):
        """Get the next unused function number for the current locomotive."""
        if not self.current_loco: return 0
        used_numbers = set(self.current_loco.function_details.keys())
        for i in range(128):
            if i not in used_numbers: return i
        return 128

    def get_available_icons(self):
        """Get list of available icon names - only from icon_mapping.json, ensuring files exist."""
        icon_names = []
        project_root = Path(__file__).parent.parent
        icons_dir = project_root / "icons"
        
        if not icons_dir.exists():
            return []
        
        # Only use icons from icon_mapping.json - do not scan file system
        for icon_key, icon_data in self.icon_mapping.items():
            if isinstance(icon_data, dict):
                # New format: {"path": "...", "filename": "..."}
                filename = icon_data.get("filename", "")
            else:
                # Old format: just the filename string
                filename = icon_data if isinstance(icon_data, str) else ""
            
            # Only add if the mapped file actually exists
            if filename:
                file_path = icons_dir / filename
                if file_path.exists():
                    icon_names.append(icon_key)
        
        # Return sorted list of only mapped icons that exist
        return sorted(icon_names)

    def load_icon_image(self, icon_name: str = None, size: tuple = (80, 80)):
        """Load icon image with black foreground and white background."""
        project_root = Path(__file__).parent.parent
        icons_dir = project_root / "icons"

        def convert_to_black(img):
            if not HAS_PIL: return img
            if img.mode != "RGBA": img = img.convert("RGBA")
            original_pixels = img.load()
            gray = img.convert("L")
            gray_pixels = gray.load()

            # Simple heuristic for white foreground
            avg_intensity = sum(gray_pixels[x, y] for x in range(img.size[0])
                                for y in range(img.size[1])
                                if original_pixels[x, y][3] > 30)
            pixel_count = sum(1 for x in range(img.size[0])
                              for y in range(img.size[1])
                              if original_pixels[x, y][3] > 30)
            avg_intensity = avg_intensity / pixel_count if pixel_count > 0 else 128
            is_white_foreground = avg_intensity > 140
            DEEP_BLUE = (0, 82, 204)

            colored_img = Image.new("RGBA", img.size)
            colored_pixels = colored_img.load()

            for x in range(img.size[0]):
                for y in range(img.size[1]):
                    r, g, b, alpha = original_pixels[x, y]
                    if alpha < 5:
                        colored_pixels[x, y] = (0, 0, 0, 0)
                        continue
                    intensity = gray_pixels[x, y]
                    if is_white_foreground:
                        opacity = int(255 * (intensity / 255.0))
                        if opacity > 20:
                            colored_pixels[x,
                                           y] = (*DEEP_BLUE, max(200, opacity))
                        else:
                            colored_pixels[x, y] = (0, 0, 0, 0)
                    else:
                        opacity = int(255 * ((255 - intensity) / 255.0))
                        if opacity > 20:
                            colored_pixels[x, y] = (0, 0, 0, max(200, opacity))
                        else:
                            colored_pixels[x, y] = (0, 0, 0, 0)
            return colored_img

        if icon_name:
            # First, try to load from icon_mapping.json
            if icon_name in self.icon_mapping:
                icon_data = self.icon_mapping[icon_name]
                if isinstance(icon_data, dict):
                    # New format: {"path": "...", "filename": "..."}
                    filename = icon_data.get("filename", "")
                else:
                    # Old format: just the filename string
                    filename = icon_data if isinstance(icon_data, str) else ""
                
                if filename:
                    icon_path = icons_dir / filename
                    if icon_path.exists() and HAS_PIL:
                        try:
                            img = Image.open(icon_path)
                            img = convert_to_black(img)
                            white_bg = Image.new("RGB", size, color="white")
                            icon_resized = img.resize(size, Image.LANCZOS)
                            white_bg.paste(
                                icon_resized, (0, 0), icon_resized
                                if icon_resized.mode == "RGBA" else None)
                            return ctk.CTkImage(light_image=white_bg, size=size)
                        except Exception:
                            pass  # Fall through to pattern matching
            
            # Fallback: try pattern matching if not in mapping or mapping failed
            icon_patterns = [
                icon_name, f"{icon_name}_normal.png",
                f"{icon_name}_Normal.png", f"{icon_name}.png"
            ]
            for pattern in icon_patterns:
                icon_path = icons_dir / pattern
                if icon_path.exists() and HAS_PIL:
                    try:
                        img = Image.open(icon_path)
                        img = convert_to_black(img)
                        white_bg = Image.new("RGB", size, color="white")
                        icon_resized = img.resize(size, Image.LANCZOS)
                        white_bg.paste(
                            icon_resized, (0, 0), icon_resized
                            if icon_resized.mode == "RGBA" else None)
                        return ctk.CTkImage(light_image=white_bg, size=size)
                    except Exception:
                        continue

        if self.default_icon_path.exists() and HAS_PIL:
            # ... (Load default) ...
            pass

        if HAS_PIL:
            return ctk.CTkImage(light_image=Image.new("RGB",
                                                      size,
                                                      color="white"),
                                size=size)

    def load_locomotive_image(self,
                              image_name: str = None,
                              size: tuple = (227, 94)):
        """Load locomotive image from Z21 ZIP file."""
        if not image_name or not HAS_PIL: return None
        try:
            with zipfile.ZipFile(self.z21_file, "r") as zf:
                image_path = next(
                    (f for f in zf.namelist() if f.endswith(image_name)), None)
                if image_path:
                    from io import BytesIO
                    img = Image.open(BytesIO(zf.read(image_path)))
                    img.thumbnail(size, Image.LANCZOS)
                    bg_img = Image.new("RGB", size, color="white")
                    x_off, y_off = (size[0] - img.size[0]) // 2, (
                        size[1] - img.size[1]) // 2
                    bg_img.paste(img, (x_off, y_off),
                                 img if img.mode == "RGBA" else None)
                    return ctk.CTkImage(light_image=bg_img, size=bg_img.size)
        except Exception as e:
            print(f"Error loading locomotive image '{image_name}': {e}")
            return None

    def create_function_card(self, func_num: int, func_info):
        """Create a card widget for a function."""
        ICON_SIZE, CARD_PADDING = 80, 5
        # Don't set fixed width - let card adapt to column width
        # The card will center its content and adapt to available column space
        card_frame = ctk.CTkFrame(self.functions_frame_inner,
                                  border_width=2,
                                  fg_color="white")
        # Allow card to adapt to column width
        card_frame.pack_propagate(True)

        icon_image = self.load_icon_image(func_info.image_name,
                                          (ICON_SIZE, ICON_SIZE))
        if icon_image:
            icon_label = ctk.CTkLabel(card_frame,
                                      image=icon_image,
                                      fg_color="white",
                                      text="")
            icon_label.image = icon_image
            icon_label.grid(row=0,
                            column=0,
                            padx=CARD_PADDING,
                            pady=(CARD_PADDING, 1),
                            sticky="")
        else:
            icon_name_short = func_info.image_name[:
                                                   8] if func_info.image_name else "?"
            ctk.CTkLabel(card_frame,
                         text=icon_name_short,
                         width=ICON_SIZE,
                         height=ICON_SIZE,
                         fg_color=("gray95", "gray20"),
                         corner_radius=4).grid(row=0,
                                               column=0,
                                               padx=CARD_PADDING,
                                               pady=(CARD_PADDING, 1),
                                               sticky="")

        ctk.CTkLabel(card_frame,
                     text=f"F{func_num}",
                     font=("Arial", 11, "bold"),
                     fg_color="white",
                     text_color="#333333",
                     height=13).grid(row=1, column=0, pady=(0, 5), sticky="")

        shortcut_text = func_info.shortcut if func_info.shortcut else "—"
        text_color = "#0066CC" if func_info.shortcut else "#CCCCCC"
        ctk.CTkLabel(card_frame,
                     text=shortcut_text,
                     font=("Arial", 9, "bold"),
                     fg_color="white",
                     text_color=text_color,
                     height=11).grid(row=2, column=0, pady=(0, 5), sticky="")

        type_time_frame = ctk.CTkFrame(card_frame, fg_color="white")
        type_time_frame.grid(row=3, column=0, pady=(0, 5), sticky="")

        button_type_colors = {
            "switch": "#4CAF50",
            "push-button": "#FF9800",
            "time button": "#2196F3"
        }
        btn_color = button_type_colors.get(func_info.button_type_name(),
                                           "#666666")
        ctk.CTkLabel(type_time_frame,
                     text=func_info.button_type_name(),
                     font=("Arial", 10, "bold"),
                     fg_color="white",
                     text_color=btn_color,
                     height=12).pack(side="left")

        # Only show time for time button (button_type == 2), not for push-button (button_type == 1)
        if func_info.button_type == 2 and func_info.time and func_info.time != "0":
            ctk.CTkLabel(type_time_frame,
                         text=f" ⏱ {func_info.time}s",
                         font=("Arial", 8),
                         fg_color="white",
                         text_color="#666666",
                         height=10).pack(side="left")

        card_frame.grid_columnconfigure(0, weight=1)
        return card_frame


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Z21 Locomotive Browser GUI")
    parser.add_argument("file",
                        type=Path,
                        nargs="?",
                        default=Path("z21_new.z21"),
                        help="Z21 file to open (default: z21_new.z21)")
    args = parser.parse_args()

    if not args.file.exists():
        print(f"Error: File not found: {args.file}")
        sys.exit(1)

    os.environ["PYTHONUNBUFFERED"] = "1"
    if sys.platform == "darwin":
        # Suppress macOS system logging warnings
        os.environ["OS_ACTIVITY_MODE"] = "disable"

        class macOSWarningFilter:
            """Filter macOS-specific stderr warnings that don't affect functionality."""

            def __init__(self, original):
                self.original = original
                import re
                # Patterns to filter out macOS warnings
                self.filter_patterns = [
                    re.compile(r'.*TSM.*', re.IGNORECASE),
                    re.compile(r'.*IMKCFRunLoopWakeUpReliable.*',
                               re.IGNORECASE),
                    re.compile(r'.*error messaging the mach port.*',
                               re.IGNORECASE),
                    re.compile(r'.*NSOpenPanel.*overrides the method.*',
                               re.IGNORECASE),
                    re.compile(
                        r".*The class 'NSOpenPanel' overrides the method.*",
                        re.IGNORECASE),
                ]

            def write(self, text):
                # Filter out macOS-specific warnings
                for pattern in self.filter_patterns:
                    if pattern.search(text):
                        return
                self.original.write(text)

            def flush(self):
                self.original.flush()

        sys.stderr = macOSWarningFilter(sys.stderr)

    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    app = Z21GUI(root, args.file)
    root.mainloop()


if __name__ == "__main__":
    main()
