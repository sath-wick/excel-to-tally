import os
import sys
import json
import tkinter as tk
from tkinter import messagebox, filedialog

# Ensure the root path and src root are in the sys.path
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
src_root = os.path.abspath(os.path.dirname(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)
if src_root not in sys.path:
    sys.path.insert(0, src_root)

# Tally color palette
TALLY_BG = "#0B2527"          # Main background
TALLY_PANEL_BG = "#13383B"    # Panel background
TALLY_LIST_BG = "#0F2E31"     # Options List background
TALLY_TEXT_NORMAL = "#E6F2F2" # White text
TALLY_TEXT_MUTED = "#9FB1B2"  # Muted gray text
TALLY_HIGHLIGHT = "#EFE83E"   # Tally Yellow highlight
TALLY_CYAN = "#00E5FF"        # Tally Cyan for keys
TALLY_WARN = "#FF5252"        # Alert red

class TallyClientManagerGUI:
    def __init__(self, root, on_exit_callback=None):
        self.root = root
        self.on_exit_callback = on_exit_callback
        
        self.root.title("Tally Client Manager")
        self.root.geometry("1150x700")
        self.root.configure(bg=TALLY_BG)
        
        # State variables
        self.level = 0  # 0: Client list, 1: Category list, 2: Ledger list, 3: Ledger Editor
        self.selected_client = ""
        self.selected_category = ""
        self.selected_ledger = ""
        self.prompt_action = None  # Stores callback for submit prompt
        self.confirm_action = None  # Stores callback for confirmation
        
        # Bindings tracker
        self.bindings = []
        
        self.build_ui()
        self.setup_keyboard_navigation()
        self.load_config_data()
        self.show_level()
        
    def build_ui(self):
        # Container frame
        self.main_container = tk.Frame(self.root, bg=TALLY_BG)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # Header Banner
        header_bar = tk.Frame(self.main_container, bg="#0E2F32", height=30)
        header_bar.pack(fill=tk.X)
        header_lbl = tk.Label(
            header_bar, 
            text=" Tally.ERP 9  |  Manage Clients  |  Client Configurations", 
            bg="#0E2F32", 
            fg=TALLY_CYAN, 
            font=("Segoe UI", 10, "bold")
        )
        header_lbl.pack(side=tk.LEFT, padx=10)
        
        # Workspace Container
        workspace = tk.Frame(self.main_container, bg=TALLY_BG)
        workspace.pack(fill=tk.BOTH, expand=True)
        
        # Right Button Bar
        self.button_frame = tk.Frame(workspace, bg="#081E20", width=160)
        self.button_frame.pack(side=tk.RIGHT, fill=tk.Y)
        self.button_frame.pack_propagate(False)
        
        # Left Panel (Exploration & Forms)
        self.left_panel = tk.Frame(workspace, bg=TALLY_PANEL_BG, bd=1, relief=tk.SOLID)
        self.left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=12, pady=12)
        
        # Breadcrumbs Title Bar
        self.breadcrumbs_lbl = tk.Label(
            self.left_panel, 
            text="Clients Directory Explorer", 
            bg="#0F2E31", 
            fg=TALLY_TEXT_NORMAL, 
            font=("Segoe UI", 11, "bold"), 
            pady=6,
            anchor="w",
            padx=15
        )
        self.breadcrumbs_lbl.pack(fill=tk.X)
        
        # Exploration Subframe (Listbox + description text)
        self.explorer_frame = tk.Frame(self.left_panel, bg=TALLY_PANEL_BG)
        self.explorer_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Listbox widget
        self.listbox = tk.Listbox(
            self.explorer_frame,
            bg=TALLY_LIST_BG,
            fg=TALLY_TEXT_NORMAL,
            selectbackground=TALLY_HIGHLIGHT,
            selectforeground="#000000",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            highlightthickness=0,
            relief=tk.FLAT
        )
        self.listbox.pack(fill=tk.BOTH, expand=True)
        self.listbox.bind("<Double-Button-1>", lambda e: self.on_enter_key())
        
        # Form Editor Subframe (Level 3, initially hidden)
        self.form_frame = tk.Frame(self.left_panel, bg=TALLY_PANEL_BG)
        
        # Form Row 1 (Ledger Name)
        self.row1_frame = tk.Frame(self.form_frame, bg=TALLY_PANEL_BG, pady=10)
        self.row1_frame.pack(fill=tk.X, padx=15)
        self.cursor1 = tk.Label(self.row1_frame, text="  ", bg=TALLY_PANEL_BG, fg=TALLY_HIGHLIGHT, font=("Consolas", 11, "bold"), width=3)
        self.cursor1.pack(side=tk.LEFT)
        tk.Label(self.row1_frame, text=f"{'Ledger Name':<18}:", bg=TALLY_PANEL_BG, fg=TALLY_TEXT_MUTED, font=("Segoe UI", 10, "bold"), anchor="w", width=20).pack(side=tk.LEFT)
        
        self.ledger_name_var = tk.StringVar()
        self.ledger_name_entry = tk.Entry(self.row1_frame, textvariable=self.ledger_name_var, bg=TALLY_HIGHLIGHT, fg="#000000", insertbackground="#000000", bd=0, relief=tk.FLAT, font=("Segoe UI", 10, "bold"), width=40)
        self.ledger_name_entry.pack(side=tk.LEFT, padx=6)
        
        # Form Row 2 (Rule Path)
        self.row2_frame = tk.Frame(self.form_frame, bg=TALLY_PANEL_BG, pady=10)
        self.row2_frame.pack(fill=tk.X, padx=15)
        self.cursor2 = tk.Label(self.row2_frame, text="  ", bg=TALLY_PANEL_BG, fg=TALLY_HIGHLIGHT, font=("Consolas", 11, "bold"), width=3)
        self.cursor2.pack(side=tk.LEFT)
        tk.Label(self.row2_frame, text=f"{'Rule Path':<18}:", bg=TALLY_PANEL_BG, fg=TALLY_TEXT_MUTED, font=("Segoe UI", 10, "bold"), anchor="w", width=20).pack(side=tk.LEFT)
        
        self.rule_path_var = tk.StringVar()
        self.rule_path_entry = tk.Entry(self.row2_frame, textvariable=self.rule_path_var, bg=TALLY_PANEL_BG, fg=TALLY_TEXT_NORMAL, insertbackground="#FFFFFF", bd=0, relief=tk.FLAT, font=("Segoe UI", 10, "bold"), width=50)
        self.rule_path_entry.pack(side=tk.LEFT, padx=6)
        
        self.browse_btn = tk.Button(
            self.row2_frame,
            text=" Browse... ",
            bg=TALLY_LIST_BG,
            fg=TALLY_TEXT_NORMAL,
            font=("Segoe UI", 8, "bold"),
            bd=1,
            relief=tk.SOLID,
            command=self.browse_rule_path
        )
        self.browse_btn.pack(side=tk.LEFT, padx=6)
        
        # Bind Form field entries to transition keys
        self.ledger_name_entry.bind("<Return>", lambda e: self.focus_editor_row(1))
        self.ledger_name_entry.bind("<Escape>", lambda e: self.on_esc_key())
        self.rule_path_entry.bind("<Return>", lambda e: self.save_ledger_details())
        self.rule_path_entry.bind("<Escape>", lambda e: self.on_esc_key())
        self.rule_path_entry.bind("<Double-Button-1>", lambda e: self.browse_rule_path())
        
        # Dialog Popups
        # 1. Custom Prompt Frame for alphanumeric entries (Client, Category, Ledger Name)
        self.prompt_frame = tk.Frame(
            workspace, 
            bg=TALLY_PANEL_BG, 
            bd=2, 
            relief=tk.SOLID, 
            highlightbackground=TALLY_HIGHLIGHT, 
            highlightthickness=1
        )
        self.prompt_lbl = tk.Label(self.prompt_frame, text="Enter Value:", bg=TALLY_PANEL_BG, fg=TALLY_TEXT_NORMAL, font=("Segoe UI", 11, "bold"))
        self.prompt_lbl.pack(pady=(15, 5))
        
        self.prompt_var = tk.StringVar()
        self.prompt_entry = tk.Entry(self.prompt_frame, textvariable=self.prompt_var, bg=TALLY_HIGHLIGHT, fg="#000000", insertbackground="#000000", bd=0, relief=tk.FLAT, font=("Segoe UI", 11, "bold"), width=35)
        self.prompt_entry.pack(padx=20, pady=10)
        self.prompt_entry.bind("<Return>", lambda e: self.submit_prompt())
        self.prompt_entry.bind("<Escape>", lambda e: self.hide_prompt())
        
        # 2. Centered Quit/Confirmation popup
        self.confirm_frame = tk.Frame(
            workspace, 
            bg=TALLY_PANEL_BG, 
            bd=2, 
            relief=tk.SOLID, 
            highlightbackground=TALLY_HIGHLIGHT, 
            highlightthickness=1
        )
        self.confirm_lbl = tk.Label(self.confirm_frame, text="Are you sure?", bg=TALLY_PANEL_BG, fg=TALLY_TEXT_NORMAL, font=("Segoe UI", 12, "bold"))
        self.confirm_lbl.pack(pady=15)
        
        confirm_btn_frame = tk.Frame(self.confirm_frame, bg=TALLY_PANEL_BG)
        confirm_btn_frame.pack(pady=(0, 15), padx=25)
        
        self.confirm_yes_btn = tk.Button(confirm_btn_frame, text=" Yes (Y) ", bg=TALLY_LIST_BG, fg=TALLY_TEXT_NORMAL, font=("Segoe UI", 10, "bold"), bd=1, relief=tk.SOLID, command=self.submit_confirmation)
        self.confirm_yes_btn.pack(side=tk.LEFT, padx=12)
        
        self.confirm_no_btn = tk.Button(confirm_btn_frame, text=" No (N) ", bg=TALLY_LIST_BG, fg=TALLY_TEXT_NORMAL, font=("Segoe UI", 10, "bold"), bd=1, relief=tk.SOLID, command=self.hide_confirmation)
        self.confirm_no_btn.pack(side=tk.RIGHT, padx=12)
        
    def create_sidebar_button(self, key_text, desc_text, command):
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
            anchor="w"
        )
        btn.pack(fill=tk.X, padx=5, pady=3)
        return btn
        
    def setup_keyboard_navigation(self):
        self.bindings = []
        def root_bind(event, handler):
            bind_id = self.root.bind(event, handler)
            self.bindings.append((event, bind_id))
            
        root_bind("<Up>", self.on_arrow_key)
        root_bind("<Down>", self.on_arrow_key)
        root_bind("<Return>", self.on_enter_key)
        root_bind("<Escape>", self.on_esc_key)
        root_bind("<Key>", self.on_key_typed)
        
    def destroy(self):
        for event, bind_id in self.bindings:
            try:
                self.root.unbind(event, bind_id)
            except:
                pass
        self.main_container.destroy()
        
    # JSON Data Helpers
    def load_config_data(self):
        config_path = os.path.join(project_root, "config", "clients.json")
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                self.config = json.load(f)
        except Exception as e:
            messagebox.showerror("Config Error", f"Failed to load config/clients.json: {e}")
            self.config = {"clients": {}}
            
    def save_config_data(self):
        config_path = os.path.join(project_root, "config", "clients.json")
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            messagebox.showerror("Config Error", f"Failed to save config/clients.json: {e}")

    # UI Refresh Engine
    def show_level(self):
        self.listbox.selection_clear(0, tk.END)
        self.listbox.delete(0, tk.END)
        
        # Hide Form editor subframe if returning to list
        if self.level < 3:
            self.form_frame.pack_forget()
            self.explorer_frame.pack(fill=tk.BOTH, expand=True)
            self.listbox.focus_set()
            
        # Draw buttons on sidebar depending on level
        for w in self.button_frame.winfo_children():
            w.destroy()
            
        self.create_sidebar_button("Esc", "Back", self.on_esc_key)
        
        if self.level == 0:
            self.breadcrumbs_lbl.config(text="Clients Directory Explorer")
            clients = sorted(list(self.config.get("clients", {}).keys()))
            for c in clients:
                self.listbox.insert(tk.END, c)
            self.create_sidebar_button("C", "Create Client", lambda: self.show_prompt("Enter Client Name:", self.create_client))
            self.create_sidebar_button("D", "Delete Client", lambda: self.show_confirmation("Delete selected Client?", self.delete_client))
            
        elif self.level == 1:
            self.breadcrumbs_lbl.config(text=f"Client: {self.selected_client} > Categories")
            categories = sorted(list(self.config["clients"][self.selected_client].get("bank_categories", {}).keys()))
            for cat in categories:
                self.listbox.insert(tk.END, cat)
            self.create_sidebar_button("A", "Add Category", lambda: self.show_prompt("Enter Category Name:", self.add_category))
            self.create_sidebar_button("D", "Delete Category", lambda: self.show_confirmation("Delete selected Category?", self.delete_category))
            
        elif self.level == 2:
            self.breadcrumbs_lbl.config(text=f"Client: {self.selected_client} > Category: {self.selected_category} > Ledgers")
            ledgers = self.config["clients"][self.selected_client]["bank_categories"].get(self.selected_category, [])
            for led in ledgers:
                self.listbox.insert(tk.END, led)
            self.create_sidebar_button("B", "Add Ledger", lambda: self.show_prompt("Enter Bank/Ledger Name:", self.add_ledger))
            self.create_sidebar_button("D", "Delete Ledger", lambda: self.show_confirmation("Delete selected Ledger?", self.delete_ledger))
            self.create_sidebar_button("E", "Edit Ledger", self.open_ledger_editor)
            
        elif self.level == 3:
            self.breadcrumbs_lbl.config(text=f"Edit Ledger: {self.selected_ledger} under Category {self.selected_category}")
            self.explorer_frame.pack_forget()
            self.form_frame.pack(fill=tk.BOTH, expand=True, pady=15)
            
            # Load ledger details
            bank_info = self.config["clients"][self.selected_client].get("banks", {}).get(self.selected_ledger, {})
            self.ledger_name_var.set(self.selected_ledger)
            self.rule_path_var.set(bank_info.get("rule_path", ""))
            
            self.focus_editor_row(0)
            self.create_sidebar_button("Enter", "Save Details", self.save_ledger_details)
            
        if self.listbox.size() > 0 and self.level < 3:
            self.listbox.selection_set(0)
            self.listbox.activate(0)
            
    # Keyboard Navigation handlers
    def on_arrow_key(self, event):
        if self.prompt_frame.winfo_viewable() or self.confirm_frame.winfo_viewable():
            return "break"
        if self.level == 3:
            # Level 3: Up/Down arrow moves focus between form fields
            if event.keysym == "Up":
                self.focus_editor_row(0)
            else:
                self.focus_editor_row(1)
            return "break"
            
        # If the listbox has focus, let the default Tkinter handler do the work!
        # Do not manually move selection again.
        if self.root.focus_get() == self.listbox:
            return
            
        curr = self.listbox.curselection()
        if not curr:
            if self.listbox.size() > 0:
                self.listbox.selection_set(0)
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
        if self.prompt_frame.winfo_viewable() or self.confirm_frame.winfo_viewable():
            return
            
        if self.level == 3:
            # Hitting Enter inside the editor saves or moves focus
            focused = self.root.focus_get()
            if focused == self.ledger_name_entry:
                self.focus_editor_row(1)
            elif focused == self.rule_path_entry:
                self.save_ledger_details()
            else:
                self.save_ledger_details()
            return "break"
            
        curr = self.listbox.curselection()
        if not curr:
            return
        selected_val = self.listbox.get(curr[0])
        
        if self.level == 0:
            self.selected_client = selected_val
            self.level = 1
            self.show_level()
        elif self.level == 1:
            self.selected_category = selected_val
            self.level = 2
            self.show_level()
        elif self.level == 2:
            self.selected_ledger = selected_val
            self.open_ledger_editor()
            
    def on_esc_key(self, event=None):
        if self.prompt_frame.winfo_viewable():
            self.hide_prompt()
            return
        if self.confirm_frame.winfo_viewable():
            self.hide_confirmation()
            return
            
        if self.level == 0:
            if self.on_exit_callback:
                self.on_exit_callback()
        elif self.level == 1:
            self.level = 0
            self.show_level()
        elif self.level == 2:
            self.level = 1
            self.show_level()
        elif self.level == 3:
            self.level = 2
            self.show_level()
            
    def on_key_typed(self, event):
        if self.prompt_frame.winfo_viewable():
            return
        if self.confirm_frame.winfo_viewable():
            if event.keysym in ("y", "Y"):
                self.submit_confirmation()
            elif event.keysym in ("n", "N", "Escape"):
                self.hide_confirmation()
            return
            
        if self.level == 3:
            return
            
        # Hotkeys at list level
        if len(event.char) == 1:
            char = event.char.upper()
            if self.level == 0:
                if char == "C":
                    self.show_prompt("Enter Client Name:", self.create_client)
                elif char == "D":
                    self.show_confirmation("Delete selected Client?", self.delete_client)
            elif self.level == 1:
                if char == "A":
                    self.show_prompt("Enter Category Name:", self.add_category)
                elif char == "D":
                    self.show_confirmation("Delete selected Category?", self.delete_category)
            elif self.level == 2:
                if char == "B":
                    self.show_prompt("Enter Bank/Ledger Name:", self.add_ledger)
                elif char == "D":
                    self.show_confirmation("Delete selected Ledger?", self.delete_ledger)
                elif char == "E":
                    self.open_ledger_editor()

    # Form Editor Helpers (Level 3)
    def focus_editor_row(self, row_idx):
        if row_idx == 0:
            self.row1_frame.config(bg=TALLY_HIGHLIGHT)
            self.cursor1.config(text="👉", bg=TALLY_HIGHLIGHT, fg="#000000")
            self.ledger_name_entry.config(bg=TALLY_HIGHLIGHT, fg="#000000", insertbackground="#000000")
            
            self.row2_frame.config(bg=TALLY_PANEL_BG)
            self.cursor2.config(text="  ", bg=TALLY_PANEL_BG, fg=TALLY_HIGHLIGHT)
            self.rule_path_entry.config(bg=TALLY_PANEL_BG, fg=TALLY_TEXT_NORMAL, insertbackground="#FFFFFF")
            self.browse_btn.config(bg=TALLY_LIST_BG, fg=TALLY_TEXT_NORMAL)
            self.ledger_name_entry.focus_set()
        else:
            self.row1_frame.config(bg=TALLY_PANEL_BG)
            self.cursor1.config(text="  ", bg=TALLY_PANEL_BG, fg=TALLY_HIGHLIGHT)
            self.ledger_name_entry.config(bg=TALLY_PANEL_BG, fg=TALLY_TEXT_NORMAL, insertbackground="#FFFFFF")
            
            self.row2_frame.config(bg=TALLY_HIGHLIGHT)
            self.cursor2.config(text="👉", bg=TALLY_HIGHLIGHT, fg="#000000")
            self.rule_path_entry.config(bg=TALLY_HIGHLIGHT, fg="#000000", insertbackground="#000000")
            self.browse_btn.config(bg=TALLY_HIGHLIGHT, fg="#000000")
            self.rule_path_entry.focus_set()
            
    def browse_rule_path(self):
        initial_dir = os.path.join(project_root, "rules", self.selected_client)
        os.makedirs(initial_dir, exist_ok=True)
        file_path = filedialog.askopenfilename(
            title="Select Rules JSON File",
            initialdir=initial_dir,
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if file_path:
            # Convert to relative path if it is under project_root
            rel_path = os.path.relpath(file_path, project_root)
            if not rel_path.startswith(".."):
                # Use forward slashes
                rel_path = "./" + rel_path.replace("\\", "/")
                self.rule_path_var.set(rel_path)
            else:
                self.rule_path_var.set(file_path.replace("\\", "/"))
            
            # Refocus the rule path entry so they can save by pressing Enter
            self.focus_editor_row(1)
            
    def open_ledger_editor(self):
        curr = self.listbox.curselection()
        if not curr:
            return
        self.selected_ledger = self.listbox.get(curr[0])
        self.level = 3
        self.show_level()
        
    def save_ledger_details(self):
        new_name = self.ledger_name_var.get().strip()
        new_path = self.rule_path_var.get().strip()
        
        if not new_name:
            messagebox.showerror("Validation Error", "Ledger Name cannot be empty.")
            return
            
        client_dict = self.config["clients"][self.selected_client]
        
        # 1. Update in bank_categories
        ledgers = client_dict["bank_categories"].get(self.selected_category, [])
        if self.selected_ledger in ledgers:
            idx = ledgers.index(self.selected_ledger)
            ledgers[idx] = new_name
            
        # 2. Update in banks dictionary
        banks_dict = client_dict.setdefault("banks", {})
        info = banks_dict.pop(self.selected_ledger, {})
        info["ledger"] = new_name
        info["rule_path"] = new_path
        banks_dict[new_name] = info
        
        self.save_config_data()
        
        self.level = 2
        self.selected_ledger = new_name
        self.show_level()

    # Prompt Dialog System
    def show_prompt(self, label_text, action):
        self.prompt_var.set("")
        self.prompt_lbl.config(text=label_text)
        self.prompt_action = action
        
        self.prompt_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        self.prompt_entry.focus_set()
        
    def hide_prompt(self):
        self.prompt_frame.place_forget()
        self.listbox.focus_set()
        
    def submit_prompt(self):
        val = self.prompt_var.get().strip()
        if val and self.prompt_action:
            self.prompt_action(val)
        self.hide_prompt()
        
    # Confirmation Dialog System
    def show_confirmation(self, label_text, action):
        self.confirm_lbl.config(text=label_text)
        self.confirm_action = action
        
        self.confirm_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        self.confirm_yes_btn.focus_set()
        
    def hide_confirmation(self):
        self.confirm_frame.place_forget()
        self.listbox.focus_set()
        
    def submit_confirmation(self):
        if self.confirm_action:
            self.confirm_action()
        self.hide_confirmation()

    # Mutator Operations
    def create_client(self, client_name):
        if client_name in self.config.get("clients", {}):
            messagebox.showwarning("Duplicate Client", f"Client '{client_name}' already exists.")
            return
            
        self.config.setdefault("clients", {})
        self.config["clients"][client_name] = {
            "duplicate_json_path": f"./exports/{client_name}/Transactions.json",
            "ignore_json_path": f"./rules/{client_name}/ignored_descriptions.json",
            "bank_categories": {},
            "banks": {}
        }
        self.save_config_data()
        
        # Auto-create client directories
        exp_dir = os.path.join(project_root, "exports", client_name)
        rules_dir = os.path.join(project_root, "rules", client_name)
        os.makedirs(exp_dir, exist_ok=True)
        os.makedirs(rules_dir, exist_ok=True)
        
        messagebox.showinfo("Success", f"Client '{client_name}' created successfully.\nFolders created under exports/ and rules/.")
        self.show_level()
        
    def delete_client(self):
        curr = self.listbox.curselection()
        if not curr:
            return
        client_name = self.listbox.get(curr[0])
        
        self.config.get("clients", {}).pop(client_name, None)
        self.save_config_data()
        self.show_level()
        
    def add_category(self, cat_name):
        client_dict = self.config["clients"][self.selected_client]
        categories = client_dict.setdefault("bank_categories", {})
        if cat_name in categories:
            messagebox.showwarning("Duplicate Category", f"Category '{cat_name}' already exists.")
            return
            
        categories[cat_name] = []
        self.save_config_data()
        messagebox.showinfo("Success", f"Category '{cat_name}' added successfully under client '{self.selected_client}'.")
        self.show_level()
        
    def delete_category(self):
        curr = self.listbox.curselection()
        if not curr:
            return
        cat_name = self.listbox.get(curr[0])
        
        client_dict = self.config["clients"][self.selected_client]
        
        # Delete from categories
        client_dict.get("bank_categories", {}).pop(cat_name, None)
        self.save_config_data()
        self.show_level()
        
    def add_ledger(self, ledger_name):
        client_dict = self.config["clients"][self.selected_client]
        ledgers = client_dict["bank_categories"].setdefault(self.selected_category, [])
        
        if ledger_name in ledgers:
            messagebox.showwarning("Duplicate Ledger", f"Ledger '{ledger_name}' already exists.")
            return
            
        ledgers.append(ledger_name)
        
        # Add to banks
        banks_dict = client_dict.setdefault("banks", {})
        banks_dict[ledger_name] = {
            "ledger": ledger_name,
            "rule_path": f"./rules/{self.selected_client}/description_rules_{ledger_name.replace(' ', '_').lower()}.json"
        }
        
        self.save_config_data()
        messagebox.showinfo("Success", f"Ledger '{ledger_name}' added successfully under category '{self.selected_category}'.")
        self.show_level()
        
    def delete_ledger(self):
        curr = self.listbox.curselection()
        if not curr:
            return
        ledger_name = self.listbox.get(curr[0])
        
        client_dict = self.config["clients"][self.selected_client]
        
        # Remove from category list
        ledgers = client_dict["bank_categories"].get(self.selected_category, [])
        if ledger_name in ledgers:
            ledgers.remove(ledger_name)
            
        # Remove from banks dictionary
        client_dict.get("banks", {}).pop(ledger_name, None)
        
        self.save_config_data()
        self.show_level()
