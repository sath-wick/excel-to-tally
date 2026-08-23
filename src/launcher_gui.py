import os
import sys
import subprocess
import threading
import queue
import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext

# Ensure the root path and src root are in the sys.path
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
src_root = os.path.abspath(os.path.dirname(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)
if src_root not in sys.path:
    sys.path.insert(0, src_root)

from core.client_config import (
    extract_bank_code,
    get_client_bank_categories,
    get_client_names,
    is_import_enabled,
)

# Tally color palette
TALLY_BG = "#0B2527"          # Main background
TALLY_PANEL_BG = "#13383B"    # Panel background
TALLY_LIST_BG = "#0F2E31"     # Options List background
TALLY_TEXT_NORMAL = "#E6F2F2" # White text
TALLY_TEXT_MUTED = "#9FB1B2"  # Muted gray text
TALLY_HIGHLIGHT = "#EFE83E"   # Tally Yellow highlight
TALLY_CYAN = "#00E5FF"        # Tally Cyan for keys
TALLY_LOG_BG = "#051314"      # Dark console logs bg
TALLY_LOG_FG = "#CBE2E3"      # Console text
TALLY_WARN = "#FF5252"        # Alert red

class TallyImporterGUI:
    def __init__(self, root, on_exit_callback=None):
        self.root = root
        self.on_exit_callback = on_exit_callback
        self.root.title("Tally Import Generator")
        self.root.geometry("1150x700")
        self.root.configure(bg=TALLY_BG)
        
        # State variables
        self.running = False
        self.log_queue = queue.Queue()
        self.process = None
        self.selected_file_path = ""
        self.active_index = 0
        self.filter_text = ""
        self.bindings = []
        
        # File Entry Variable
        self.file_entry_var = tk.StringVar()
        
        # Form field model
        self.fields = [
            {"label": "Client Name", "var": tk.StringVar(), "widget_val": None},
            {"label": "Import Type", "var": tk.StringVar(value="Statements"), "widget_val": None},
            {"label": "Bank Category", "var": tk.StringVar(), "widget_val": None},
            {"label": "Bank / Ledger", "var": tk.StringVar(), "widget_val": None},
            {"label": "Input File", "var": self.file_entry_var, "widget_val": None},
            {"label": "", "var": tk.StringVar(value="[ Press Enter to Run Import ]"), "widget_val": None},
        ]
        
        # Initialize UI layout
        self.build_ui()
        
        # Set up keyboard hooks
        self.setup_keyboard_navigation()
        
        # Load clients
        self.load_clients()
        
    def build_ui(self):
        # Create a container frame to hold all Importer GUI elements for easy destruction/mounting
        self.main_container = tk.Frame(self.root, bg=TALLY_BG)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # 1. Tally Top Header Banner
        header_bar = tk.Frame(self.main_container, bg="#0E2F32", height=30)
        header_bar.pack(fill=tk.X)
        header_lbl = tk.Label(
            header_bar, 
            text=" Tally.ERP 9  |  Tally Import Generator  |  Keyboard Guided UI", 
            bg="#0E2F32", 
            fg=TALLY_CYAN, 
            font=("Segoe UI", 10, "bold")
        )
        header_lbl.pack(side=tk.LEFT, padx=10)
        
        # Main Workspace Container
        workspace = tk.Frame(self.main_container, bg=TALLY_BG)
        workspace.pack(fill=tk.BOTH, expand=True)
        
        # Vertical Button Bar on Right
        self.button_frame = tk.Frame(workspace, bg="#081E20", width=160)
        self.button_frame.pack(side=tk.RIGHT, fill=tk.Y)
        self.button_frame.pack_propagate(False)
        
        # Center Frame containing Left (Form) and Center (List)
        center_pane = tk.Frame(workspace, bg=TALLY_BG)
        center_pane.pack(fill=tk.BOTH, expand=True, side=tk.TOP)
        
        # Left Panel (Voucher Input Form)
        self.form_frame = tk.Frame(center_pane, bg=TALLY_PANEL_BG, bd=1, relief=tk.SOLID)
        self.form_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(12, 6), pady=12)
        
        form_title = tk.Label(
            self.form_frame, 
            text="Voucher Configuration", 
            bg="#0F2E31", 
            fg=TALLY_TEXT_NORMAL, 
            font=("Segoe UI", 11, "bold"), 
            pady=6
        )
        form_title.pack(fill=tk.X)
        
        # Right Panel (Dynamic Choices List)
        self.list_frame = tk.Frame(center_pane, bg=TALLY_PANEL_BG, bd=1, relief=tk.SOLID, width=350)
        self.list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(6, 12), pady=12)
        self.list_frame.pack_propagate(False)
        
        self.list_header = tk.Label(
            self.list_frame, 
            text="List of Options", 
            bg="#0F2E31", 
            fg=TALLY_HIGHLIGHT, 
            font=("Segoe UI", 11, "bold"), 
            pady=6
        )
        self.list_header.pack(fill=tk.X)
        
        # Listbox widget
        self.listbox = tk.Listbox(
            self.list_frame,
            bg=TALLY_LIST_BG,
            fg=TALLY_TEXT_NORMAL,
            selectbackground=TALLY_HIGHLIGHT,
            selectforeground="#000000",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            highlightthickness=0,
            relief=tk.FLAT
        )
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.listbox.bind("<Double-Button-1>", lambda e: self.on_enter_key())
        
        # Setup form rows
        self.field_frames = []
        self.cursor_labels = []
        self.label_widgets = []
        self.value_widgets = []
        
        for i, field in enumerate(self.fields):
            row_frame = tk.Frame(self.form_frame, bg=TALLY_PANEL_BG, pady=10)
            row_frame.pack(fill=tk.X, padx=15)
            self.field_frames.append(row_frame)
            
            cursor = tk.Label(
                row_frame, 
                text="  ", 
                bg=TALLY_PANEL_BG, 
                fg=TALLY_HIGHLIGHT, 
                font=("Consolas", 11, "bold"), 
                width=3
            )
            cursor.pack(side=tk.LEFT)
            self.cursor_labels.append(cursor)
            
            if field['label']:
                lbl = tk.Label(
                    row_frame, 
                    text=f"{field['label']:<18}:", 
                    bg=TALLY_PANEL_BG, 
                    fg=TALLY_TEXT_MUTED, 
                    font=("Segoe UI", 10, "bold"), 
                    anchor="w", 
                    width=20
                )
                lbl.pack(side=tk.LEFT)
                self.label_widgets.append(lbl)
                
                val = tk.Label(
                    row_frame, 
                    text="", 
                    bg=TALLY_PANEL_BG, 
                    fg=TALLY_TEXT_NORMAL, 
                    font=("Segoe UI", 10, "bold"), 
                    anchor="w", 
                    width=35, 
                    padx=6
                )
                val.pack(side=tk.LEFT, fill=tk.X, expand=True)
                self.value_widgets.append(val)
            else:
                self.label_widgets.append(None)
                val = tk.Label(
                    row_frame, 
                    text=field['var'].get(), 
                    bg=TALLY_PANEL_BG, 
                    fg=TALLY_TEXT_NORMAL, 
                    font=("Segoe UI", 10, "bold"), 
                    anchor="center"
                )
                val.pack(fill=tk.X, expand=True)
                self.value_widgets.append(val)
        
        # Dedicated File Entry Widget (Packed inside left pane row 4 when active)
        self.file_entry = tk.Entry(
            self.main_container, 
            textvariable=self.file_entry_var, 
            bg=TALLY_HIGHLIGHT, 
            fg="#000000", 
            insertbackground="#000000", 
            bd=0, 
            relief=tk.FLAT, 
            font=("Segoe UI", 10, "bold")
        )
        
        # Form alerts & status
        self.warning_label = tk.Label(
            self.form_frame, 
            text="", 
            bg=TALLY_PANEL_BG, 
            fg=TALLY_WARN, 
            wraplength=400, 
            justify="left", 
            font=('Segoe UI', 9, 'bold')
        )
        self.warning_label.pack(fill=tk.X, side=tk.BOTTOM, pady=10)
        
        status_frame = tk.Frame(self.form_frame, bg=TALLY_PANEL_BG)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(0, 10), padx=15)
        
        status_lbl = tk.Label(
            status_frame, 
            text="Status :", 
            bg=TALLY_PANEL_BG, 
            fg=TALLY_TEXT_MUTED, 
            font=("Segoe UI", 9, "bold")
        )
        status_lbl.pack(side=tk.LEFT)
        
        self.status_label = tk.Label(
            status_frame, 
            text="Ready", 
            bg=TALLY_PANEL_BG, 
            fg="#00FFCC", 
            font=("Segoe UI", 10, "bold")
        )
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        # Build Stacked Sidebar Buttons
        self.create_sidebar_button("Esc", "Quit", self.on_esc_key)
        self.create_sidebar_button("F1", "Select Client", lambda: self.jump_to_field(0))
        self.create_sidebar_button("F2", "Import Type", lambda: self.jump_to_field(1))
        self.create_sidebar_button("F3", "Category", lambda: self.jump_to_field(2))
        self.create_sidebar_button("F4", "Ledger", lambda: self.jump_to_field(3))
        self.create_sidebar_button("F5", "File", lambda: self.jump_to_field(4))
        
        self.run_btn = self.create_sidebar_button("F9", "Run Import", self.start_import_process)
        self.open_excel_btn = self.create_sidebar_button("F10", "Open Excel", self.open_output_excel, state=tk.DISABLED)
        self.open_folder_btn = self.create_sidebar_button("F11", "Open Folder", self.open_output_folder)
        self.clear_log_btn = self.create_sidebar_button("F12", "Delete Logs", self.clear_logs)
        
        # Bottom Console Log Frame
        self.log_frame = tk.Frame(workspace, bg=TALLY_BG, bd=1, relief=tk.SOLID)
        self.log_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=12, pady=(0, 12))
        
        log_title = tk.Label(
            self.log_frame, 
            text="Calculator / Diagnostic Console Logs", 
            bg="#081E20", 
            fg=TALLY_TEXT_MUTED, 
            font=("Segoe UI", 9, "bold"), 
            anchor="w", 
            padx=10, 
            pady=4
        )
        log_title.pack(fill=tk.X)
        
        self.console_text = scrolledtext.ScrolledText(
            self.log_frame, 
            bg=TALLY_LOG_BG, 
            fg=TALLY_LOG_FG, 
            insertbackground="#FFFFFF", 
            state=tk.DISABLED, 
            font=("Consolas", 10),
            bd=0,
            height=8,
            highlightthickness=0
        )
        self.console_text.pack(fill=tk.BOTH, expand=True)
        
        # Authentic Tally Quit Dialog Box (Centered Popup, initially hidden)
        self.quit_frame = tk.Frame(
            self.main_container, 
            bg=TALLY_PANEL_BG, 
            bd=2, 
            relief=tk.SOLID, 
            highlightbackground=TALLY_HIGHLIGHT, 
            highlightcolor=TALLY_HIGHLIGHT, 
            highlightthickness=1
        )
        
        quit_lbl = tk.Label(
            self.quit_frame, 
            text=" Quit ? ", 
            bg=TALLY_PANEL_BG, 
            fg=TALLY_TEXT_NORMAL, 
            font=("Segoe UI", 12, "bold")
        )
        quit_lbl.pack(pady=15)
        
        quit_btn_frame = tk.Frame(self.quit_frame, bg=TALLY_PANEL_BG)
        quit_btn_frame.pack(pady=(0, 15), padx=25)
        
        self.quit_yes_btn = tk.Button(
            quit_btn_frame, 
            text=" Yes (Y) ", 
            bg=TALLY_LIST_BG, 
            fg=TALLY_TEXT_NORMAL, 
            font=("Segoe UI", 10, "bold"), 
            bd=1, 
            relief=tk.SOLID, 
            command=self.root.destroy
        )
        self.quit_yes_btn.pack(side=tk.LEFT, padx=12)
        
        self.quit_no_btn = tk.Button(
            quit_btn_frame, 
            text=" No (N) ", 
            bg=TALLY_LIST_BG, 
            fg=TALLY_TEXT_NORMAL, 
            font=("Segoe UI", 10, "bold"), 
            bd=1, 
            relief=tk.SOLID, 
            command=self.hide_quit_dialog
        )
        self.quit_no_btn.pack(side=tk.RIGHT, padx=12)
        
    def create_sidebar_button(self, key_text, desc_text, command, state=tk.NORMAL):
        btn = tk.Button(
            self.button_frame,
            text=f"{key_text}\n{desc_text}",
            bg="#1D4C50",
            fg=TALLY_TEXT_NORMAL,
            activebackground=TALLY_HIGHLIGHT,
            activeforeground="#000000",
            font=("Segoe UI", 9, "bold"),
            bd=1,
            relief=tk.RAISED,
            padx=10,
            pady=5,
            command=command,
            state=state,
            anchor="w"
        )
        btn.pack(fill=tk.X, padx=5, pady=3)
        return btn
        
    def setup_keyboard_navigation(self):
        # We store the bindings to unbind them cleanly on destruction
        self.bindings = []
        
        def root_bind(event, handler):
            bind_id = self.root.bind(event, handler)
            self.bindings.append((event, bind_id))
            
        # Bind arrow keys for scrolling options
        root_bind("<Up>", self.on_arrow_key)
        root_bind("<Down>", self.on_arrow_key)
        
        # Enter / Esc bindings
        root_bind("<Return>", self.on_enter_key)
        root_bind("<Escape>", self.on_esc_key)
        
        # Alphanumeric keyboard hooks for filtering
        root_bind("<Key>", self.on_key_typed)
        
        # Sidebar function hotkeys
        root_bind("<F1>", lambda e: self.jump_to_field(0))
        root_bind("<F2>", lambda e: self.jump_to_field(1))
        root_bind("<F3>", lambda e: self.jump_to_field(2))
        root_bind("<F4>", lambda e: self.jump_to_field(3))
        root_bind("<F5>", lambda e: self.jump_to_field(4))
        root_bind("<F9>", lambda e: self.start_import_process())
        root_bind("<F10>", lambda e: self.open_output_excel())
        root_bind("<F11>", lambda e: self.open_output_folder())
        root_bind("<F12>", lambda e: self.clear_logs())
        
    def destroy(self):
        # Unbind root keyboard hooks
        for event, bind_id in self.bindings:
            try:
                self.root.unbind(event, bind_id)
            except:
                pass
        self.main_container.destroy()
        
    def jump_to_field(self, idx):
        self.active_index = idx
        self.filter_text = ""
        self.update_active_field()
        
    def show_quit_dialog(self):
        self.quit_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        self.quit_yes_btn.focus_set()
        
    def hide_quit_dialog(self):
        self.quit_frame.place_forget()
        self.root.focus_set()
        self.update_active_field()
        
    def on_arrow_key(self, event):
        curr = self.listbox.curselection()
        if not curr:
            if self.listbox.size() > 0:
                self.listbox.selection_set(0)
                self.listbox.see(0)
            return "break"
        idx = curr[0]
        if event.keysym == "Up":
            new_idx = max(0, idx - 1)
        else:
            new_idx = min(self.listbox.size() - 1, idx + 1)
            
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(new_idx)
        self.listbox.activate(new_idx)
        self.listbox.see(new_idx)
        return "break"
        
    def on_enter_key(self, event=None):
        if self.quit_frame.winfo_viewable():
            # If quit dialog has focus, Yes button command will handle exit
            return "break"
            
        curr = self.listbox.curselection()
        if not curr:
            if self.listbox.size() > 0:
                self.listbox.selection_set(0)
                curr = (0,)
            else:
                return "break"
                
        selected_val = self.listbox.get(curr[0])
        self.apply_selection(selected_val)
        return "break"
        
    def on_esc_key(self, event=None):
        if self.quit_frame.winfo_viewable():
            self.hide_quit_dialog()
            return "break"
            
        if self.active_index == 0:
            if self.on_exit_callback:
                self.on_exit_callback()
            else:
                self.show_quit_dialog()
        else:
            self.active_index -= 1
            self.filter_text = ""
            self.update_active_field()
        return "break"
            
    def on_key_typed(self, event):
        # If quit dialog is active
        if self.quit_frame.winfo_viewable():
            if event.keysym in ("y", "Y"):
                self.root.destroy()
            elif event.keysym in ("n", "N", "Escape"):
                self.hide_quit_dialog()
            return
            
        # Ignore if typing file path in text entry
        if self.root.focus_get() == self.file_entry:
            return
            
        # Ignore if Ctrl/Alt is held
        if event.state & 0x0004 or event.state & 0x0008:
            return
            
        if len(event.char) == 1:
            code = ord(event.char)
            if code >= 32 and event.keysym not in ("Escape", "Return", "Tab"):
                self.filter_text += event.char
                self.refresh_listbox()
        elif event.keysym == "BackSpace":
            self.filter_text = self.filter_text[:-1]
            self.refresh_listbox()
            
    def get_options_for_active_field(self):
        idx = self.active_index
        if idx == 0:  # Client
            try:
                return get_client_names()
            except:
                return []
        elif idx == 1:  # Import Type
            return ["Statements", "Sales", "Purchases"]
        elif idx == 2:  # Bank Category
            client = self.fields[0]['var'].get()
            if not client:
                return []
            try:
                return list(get_client_bank_categories(client).keys())
            except:
                return []
        elif idx == 3:  # Bank Ledger
            client = self.fields[0]['var'].get()
            category = self.fields[2]['var'].get()
            if not client or not category:
                return []
            try:
                return get_client_bank_categories(client).get(category, [])
            except:
                return []
        elif idx == 4:  # File
            return ["[Press Enter to Browse]", "[Clear Path]"]
        elif idx == 5:  # Run Import
            opts = ["[Press Enter to Run Import]"]
            if self.open_excel_btn['state'] == tk.NORMAL:
                opts.append("[Alt+O to Open Excel]")
            opts.append("[Alt+F to Open Folder]")
            opts.append("[Alt+D to Delete Logs]")
            return opts
        return []
        
    def get_active_field_name(self):
        return self.fields[self.active_index]['label'] or "Run Action"
        
    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        options = self.get_options_for_active_field()
        
        # Filter options
        filtered = [opt for opt in options if self.filter_text.lower() in opt.lower()]
        for opt in filtered:
            self.listbox.insert(tk.END, opt)
            
        if filtered:
            self.listbox.selection_set(0)
            self.listbox.activate(0)
            
        # Update dynamic header
        hdr = self.get_active_field_name()
        if self.filter_text:
            self.list_header.config(text=f"List of {hdr} ({self.filter_text})")
        else:
            self.list_header.config(text=f"List of {hdr}")
            
    def apply_selection(self, val):
        idx = self.active_index
        
        if idx == 4:  # File
            if val == "[Press Enter to Browse]":
                self.browse_file()
                return
            elif val == "[Clear Path]":
                self.fields[idx]['var'].set("")
                self.selected_file_path = ""
            else:
                self.fields[idx]['var'].set(val)
                self.selected_file_path = val
            self.active_index = 5
        elif idx == 5:  # Run Import
            if val == "[Press Enter to Run Import]":
                self.start_import_process()
                return
        else:
            self.fields[idx]['var'].set(val)
            if idx == 0:
                self.on_client_changed()
                self.active_index = 1
            elif idx == 1:
                self.on_import_type_changed()
                self.active_index = 2
            elif idx == 2:
                self.on_category_changed()
                self.active_index = 3
            elif idx == 3:
                self.on_bank_changed()
                self.active_index = 4
                
        self.filter_text = ""
        self.update_active_field()
        
    def update_active_field(self):
        for i in range(len(self.fields)):
            row_frame = self.field_frames[i]
            cursor = self.cursor_labels[i]
            lbl = self.label_widgets[i]
            val = self.value_widgets[i]
            
            if i == 4 and i == self.active_index:
                # Active Input File -> Show text Entry
                cursor.config(text="👉")
                lbl.config(fg=TALLY_TEXT_NORMAL)
                val.pack_forget()
                self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                self.file_entry.config(bg=TALLY_HIGHLIGHT, fg="#000000")
                self.file_entry.focus_set()
            else:
                # Other normal fields
                if i == 4:
                    self.file_entry.pack_forget()
                    val.pack(side=tk.LEFT, fill=tk.X, expand=True)
                    if self.root.focus_get() == self.file_entry:
                        self.root.focus_set()
                    
                if i == self.active_index:
                    cursor.config(text="👉")
                    if lbl:
                        lbl.config(fg=TALLY_TEXT_NORMAL)
                        val.config(bg=TALLY_HIGHLIGHT, fg="#000000", text=self.fields[i]['var'].get())
                    else:
                        val.config(bg=TALLY_HIGHLIGHT, fg="#000000")
                else:
                    cursor.config(text="  ")
                    if lbl:
                        lbl.config(fg=TALLY_TEXT_MUTED)
                        val.config(bg=TALLY_PANEL_BG, fg=TALLY_TEXT_NORMAL, text=self.fields[i]['var'].get())
                    else:
                        val.config(bg=TALLY_PANEL_BG, fg=TALLY_TEXT_NORMAL)
                        
        self.validate_settings()
        self.refresh_listbox()
        
    def load_clients(self):
        try:
            clients = get_client_names()
            if not clients:
                messagebox.showerror("Error", "No clients configured in ./config/clients.json")
                return
            
            # Default select default_client or first client
            from core.client_config import get_default_client_name
            default_client = get_default_client_name()
            if default_client in clients:
                self.fields[0]['var'].set(default_client)
            else:
                self.fields[0]['var'].set(clients[0])
                
            self.on_client_changed()
            self.update_active_field()
        except Exception as e:
            messagebox.showerror("Configuration Error", f"Failed to load client config:\n{str(e)}")
            
    def on_client_changed(self):
        client = self.fields[0]['var'].get()
        if not client:
            return
        try:
            bank_categories = get_client_bank_categories(client)
            categories = list(bank_categories.keys())
            if categories:
                self.fields[2]['var'].set(categories[0])
                self.on_category_changed()
            else:
                self.fields[2]['var'].set("")
                self.fields[3]['var'].set("")
        except:
            pass
            
    def on_category_changed(self):
        client = self.fields[0]['var'].get()
        category = self.fields[2]['var'].get()
        if not client or not category:
            return
        try:
            bank_categories = get_client_bank_categories(client)
            banks = bank_categories.get(category, [])
            if banks:
                self.fields[3]['var'].set(banks[0])
                self.on_bank_changed()
            else:
                self.fields[3]['var'].set("")
        except:
            pass
            
    def on_bank_changed(self):
        self.validate_settings()
        
    def on_import_type_changed(self):
        self.validate_settings()
        file_path = self.fields[4]['var'].get()
        if file_path:
            _, ext = os.path.splitext(file_path.lower())
            import_type = self.fields[1]['var'].get()
            if import_type == "Statements" and ext != ".pdf":
                self.fields[4]['var'].set("")
            elif import_type in ("Sales", "Purchases") and ext not in (".xlsx", ".xls"):
                self.fields[4]['var'].set("")
                
    def validate_settings(self):
        client = self.fields[0]['var'].get()
        import_type = self.fields[1]['var'].get()
        bank = self.fields[3]['var'].get()
        
        if not client or not bank:
            self.warning_label.config(text="")
            return True
            
        if import_type in ("Sales", "Purchases"):
            enabled = is_import_enabled(client, import_type, bank)
            if not enabled:
                bank_code = extract_bank_code(bank)
                self.warning_label.config(
                    text=f"⚠️ Warning: {import_type} is not configured/enabled\nfor bank '{bank_code}' under client '{client}'.",
                    fg=TALLY_WARN
                )
                return False
                
        self.warning_label.config(text="", fg=TALLY_TEXT_NORMAL)
        return True
        
    def browse_file(self):
        import_type = self.fields[1]['var'].get()
        if import_type == "Statements":
            filetypes = [("PDF Files", "*.pdf"), ("All Files", "*.*")]
            title = "Select Bank Statement PDF"
        else:
            filetypes = [("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
            title = f"Select {import_type} Excel File"
            
        file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if file_path:
            normalized_path = file_path.replace('\\', '/')
            self.fields[4]['var'].set(normalized_path)
            self.selected_file_path = normalized_path
            
            # Auto-advance to Run button
            self.active_index = 5
            self.filter_text = ""
            self.update_active_field()
            
    def clear_logs(self):
        self.console_text.config(state=tk.NORMAL)
        self.console_text.delete(1.0, tk.END)
        self.console_text.config(state=tk.DISABLED)
        
    def write_log(self, text):
        self.console_text.config(state=tk.NORMAL)
        self.console_text.insert(tk.END, text)
        self.console_text.see(tk.END)
        self.console_text.config(state=tk.DISABLED)
        
    def open_output_excel(self):
        import_type = self.fields[1]['var'].get()
        out_name = f"{import_type}_Import.xlsx"
        out_path = os.path.abspath(os.path.join(project_root, "output", out_name))
        
        if os.path.exists(out_path):
            try:
                os.startfile(out_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open output file:\n{str(e)}")
        else:
            messagebox.showerror("Error", f"Output file does not exist:\n{out_path}")
            
    def open_output_folder(self):
        out_dir = os.path.abspath(os.path.join(project_root, "output"))
        os.makedirs(out_dir, exist_ok=True)
        try:
            os.startfile(out_dir)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open output folder:\n{str(e)}")
            
    def start_import_process(self):
        if self.running:
            return
            
        client = self.fields[0]['var'].get()
        import_type = self.fields[1]['var'].get()
        bank = self.fields[3]['var'].get()
        file_path = self.fields[4]['var'].get().strip()
        
        if not client or not bank or not file_path:
            messagebox.showerror("Missing Information", "Please configure all settings and select an input file.")
            return
            
        if not os.path.exists(file_path):
            messagebox.showerror("File Error", f"The selected file does not exist:\n{file_path}")
            return
            
        if not self.validate_settings():
            ans = messagebox.askyesno(
                "Import Disabled", 
                f"{import_type} is not enabled for bank '{extract_bank_code(bank)}'.\n\nDo you want to run it anyway?"
            )
            if not ans:
                return

        self.running = True
        self.status_label.config(text="Running...", fg=TALLY_HIGHLIGHT)
        
        # Disable action buttons in vertical sidebar
        self.run_btn.config(state=tk.DISABLED)
        self.open_excel_btn.config(state=tk.DISABLED)
        
        if import_type == "Sales":
            script_name = "src/sales_main.py"
        elif import_type == "Purchases":
            script_name = "src/purchases_main.py"
        else:
            script_name = "src/main.py"
            
        self.clear_logs()
        self.write_log(f"==================================================\n")
        self.write_log(f" Starting Tally Import Generator\n")
        self.write_log(f"==================================================\n")
        self.write_log(f" Client:        {client}\n")
        self.write_log(f" Bank Ledger:   {bank}\n")
        self.write_log(f" Import Type:   {import_type}\n")
        self.write_log(f" Input File:    {file_path}\n")
        self.write_log(f" Script:        {script_name}\n")
        self.write_log(f"--------------------------------------------------\n\n")
        
        thread = threading.Thread(
            target=self.run_process_worker, 
            args=(script_name, file_path, bank, client),
            daemon=True
        )
        thread.start()
        
        self.root.after(100, self.poll_queue)
        
    def run_process_worker(self, script_name, file_path, bank, client):
        try:
            cmd = [sys.executable, "-u", script_name, file_path, bank, client]
            self.process = subprocess.Popen(
                cmd,
                cwd=project_root,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1
            )
            
            while True:
                line = self.process.stdout.readline()
                if not line:
                    break
                self.log_queue.put(line)
                
            self.process.wait()
            return_code = self.process.returncode
            self.log_queue.put(None)
            self.log_queue.put(return_code)
        except Exception as e:
            self.log_queue.put(f"\n[GUI Process Error]: {str(e)}\n")
            self.log_queue.put(None)
            self.log_queue.put(-1)
            
    def poll_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                if msg is None:
                    self.running = False
                    return_code = self.log_queue.get(timeout=1.0)
                    self.on_run_finished(return_code)
                    break
                if isinstance(msg, int):
                    self.running = False
                    self.on_run_finished(msg)
                    break
                    
                if msg.startswith("__GUI_CONFIRM__:"):
                    prompt = msg.split("__GUI_CONFIRM__:", 1)[1].strip()
                    self.root.attributes("-topmost", True)
                    self.root.attributes("-topmost", False)
                    choice = messagebox.askyesno("Confirmation Required", prompt)
                    self.write_log(f"[GUI Prompt]: {prompt}\n")
                    self.write_log(f"[GUI Answer]: {'Yes' if choice else 'No'}\n\n")
                    response = "yes\n" if choice else "no\n"
                    if self.process and self.process.stdin:
                        try:
                            self.process.stdin.write(response)
                            self.process.stdin.flush()
                        except Exception as e:
                            self.write_log(f"\n[GUI Error]: Failed to send confirmation response: {e}\n")
                    continue

                if msg.startswith("__GUI_ALERT__:"):
                    alert_msg = msg.split("__GUI_ALERT__:", 1)[1].strip()
                    self.root.attributes("-topmost", True)
                    self.root.attributes("-topmost", False)
                    messagebox.showinfo("Tally Import Generator - Notice", alert_msg)
                    self.write_log(f"[GUI Alert]: {alert_msg}\n\n")
                    if self.process and self.process.stdin:
                        try:
                            self.process.stdin.write("ok\n")
                            self.process.stdin.flush()
                        except Exception as e:
                            self.write_log(f"\n[GUI Error]: Failed to send alert acknowledgment: {e}\n")
                    continue
                    
                self.write_log(msg)
        except queue.Empty:
            pass
            
        if self.running:
            self.root.after(100, self.poll_queue)
            
    def on_run_finished(self, return_code):
        self.run_btn.config(state=tk.NORMAL)
        
        if return_code == 0:
            self.status_label.config(text="Success!", fg="#4CAF50")
            self.open_excel_btn.config(state=tk.NORMAL)
            self.write_log(f"\n[GUI]: Process completed successfully.\n")
        else:
            self.status_label.config(text="Failed!", fg="#F44336")
            self.write_log(f"\n[GUI]: Process finished with exit code {return_code}.\n")
            
        # Refresh current options panel to show Open Excel if enabled
        self.refresh_listbox()

if __name__ == "__main__":
    root = tk.Tk()
    app = TallyImporterGUI(root)
    root.mainloop()
