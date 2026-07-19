import os
import sys
import subprocess
import tkinter as tk
from tkinter import messagebox

# Ensure the root path and src root are in the sys.path
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
src_root = os.path.abspath(os.path.dirname(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)
if src_root not in sys.path:
    sys.path.insert(0, src_root)

from core.client_config import get_client_names
from src.launcher_gui import TallyImporterGUI
from src.client_manager_gui import TallyClientManagerGUI

TALLY_BG = "#0B2527"
TALLY_PANEL_BG = "#13383B"
TALLY_LIST_BG = "#0F2E31"
TALLY_TEXT_NORMAL = "#E6F2F2"
TALLY_TEXT_MUTED = "#9FB1B2"
TALLY_HIGHLIGHT = "#EFE83E"
TALLY_CYAN = "#00E5FF"
TALLY_WARN = "#FF5252"

class GatewayGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Tally Import Suite")
        self.root.geometry("1150x700")
        self.root.configure(bg=TALLY_BG)
        
        self.active_index = 0
        self.importer_app = None
        self.client_manager_app = None
        
        self.menu_options = [
            {"label": "Import Generator", "hotkey": "I", "desc": "Import Generator", "action": self.open_importer},
            {"label": "Rule Studio", "hotkey": "R", "desc": "Rule Studio", "action": self.open_rule_studio},
            {"label": "Manage Clients", "hotkey": "M", "desc": "Manage Clients", "action": self.open_client_manager},
            {"label": "Quit", "hotkey": "Q", "desc": "Quit", "action": self.show_quit_dialog},
        ]
        
        self.build_ui()
        self.setup_keyboard_navigation()
        self.show_gateway_menu()
        
    def build_ui(self):
        # 1. Main Container Frame
        self.main_container = tk.Frame(self.root, bg=TALLY_BG)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # Header Banner
        header_bar = tk.Frame(self.main_container, bg="#0E2F32", height=30)
        header_bar.pack(fill=tk.X)
        header_lbl = tk.Label(
            header_bar, 
            text=" Tally.ERP 9  |  Gateway of Tally  |  Import & Rules Suite", 
            bg="#0E2F32", 
            fg=TALLY_CYAN, 
            font=("Segoe UI", 10, "bold")
        )
        header_lbl.pack(side=tk.LEFT, padx=10)
        
        # Workspace Frame
        self.workspace = tk.Frame(self.main_container, bg=TALLY_BG)
        self.workspace.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Right Button Panel (Sidebar)
        self.button_frame = tk.Frame(self.main_container, bg="#081E20", width=160)
        self.button_frame.pack(side=tk.RIGHT, fill=tk.Y)
        self.button_frame.pack_propagate(False)
        
        # Split Workspace into Left (Company Info) and Right (Menu Box)
        self.left_pane = tk.Frame(self.workspace, bg=TALLY_PANEL_BG, bd=1, relief=tk.SOLID)
        self.left_pane.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        self.right_pane = tk.Frame(self.workspace, bg=TALLY_PANEL_BG, bd=1, relief=tk.SOLID, width=400)
        self.right_pane.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False)
        self.right_pane.pack_propagate(False)
        
        # Left Panel content (Tally status info)
        info_header = tk.Frame(self.left_pane, bg="#0F2E31", height=25)
        info_header.pack(fill=tk.X)
        tk.Label(info_header, text="CURRENT PERIOD", bg="#0F2E31", fg=TALLY_TEXT_MUTED, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=15)
        tk.Label(info_header, text="CURRENT DATE", bg="#0F2E31", fg=TALLY_TEXT_MUTED, font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=15)
        
        dates_row = tk.Frame(self.left_pane, bg=TALLY_PANEL_BG, pady=10)
        dates_row.pack(fill=tk.X)
        tk.Label(dates_row, text="1-Apr-2025 to 31-Mar-2026", bg=TALLY_PANEL_BG, fg=TALLY_TEXT_NORMAL, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=15)
        tk.Label(dates_row, text="Sunday, 19-Jul-2026", bg=TALLY_PANEL_BG, fg=TALLY_TEXT_NORMAL, font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=15)
        
        # Separator line
        tk.Frame(self.left_pane, bg="#2A595D", height=1).pack(fill=tk.X, padx=10)
        
        company_header = tk.Frame(self.left_pane, bg=TALLY_PANEL_BG, pady=8)
        company_header.pack(fill=tk.X)
        tk.Label(company_header, text="List of Selected Companies", bg=TALLY_PANEL_BG, fg=TALLY_CYAN, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=15)
        tk.Label(company_header, text="Date of Last Entry", bg=TALLY_PANEL_BG, fg=TALLY_CYAN, font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=15)
        
        # Load real client list in company info
        try:
            clients = get_client_names()
        except:
            clients = ["Arjun Rao", "NSB Traders", "D Raju"]
            
        for client in clients[:3]:
            row = tk.Frame(self.left_pane, bg=TALLY_PANEL_BG, pady=5)
            row.pack(fill=tk.X)
            tk.Label(row, text=client, bg=TALLY_PANEL_BG, fg=TALLY_TEXT_NORMAL, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=15)
            tk.Label(row, text="19-Jul-2026", bg=TALLY_PANEL_BG, fg=TALLY_TEXT_NORMAL, font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=15)
            
        # Right Panel content (Gateway Menu Box)
        menu_title = tk.Label(self.right_pane, text="Gateway of Tally", bg="#0F2E31", fg=TALLY_TEXT_NORMAL, font=("Segoe UI", 11, "bold"), pady=6)
        menu_title.pack(fill=tk.X)
        
        menu_container = tk.Frame(self.right_pane, bg=TALLY_LIST_BG, bd=1, relief=tk.SOLID)
        menu_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self.menu_item_frames = []
        self.menu_item_labels = []
        
        # Create Menu Option Row widgets
        for i, opt in enumerate(self.menu_options):
            f = tk.Frame(menu_container, bg=TALLY_LIST_BG, pady=8)
            f.pack(fill=tk.X, padx=10, pady=5)
            self.menu_item_frames.append(f)
            
            # Hotkey is first character
            hot = opt["hotkey"]
            rest = opt["label"][1:]
            
            lbl_hot = tk.Label(f, text=hot, fg=TALLY_WARN, bg=TALLY_LIST_BG, font=("Segoe UI", 10, "bold"), underline=0)
            lbl_hot.pack(side=tk.LEFT, padx=(10, 0))
            
            lbl_rest = tk.Label(f, text=rest, fg=TALLY_TEXT_NORMAL, bg=TALLY_LIST_BG, font=("Segoe UI", 10, "bold"))
            lbl_rest.pack(side=tk.LEFT)
            
            self.menu_item_labels.append((lbl_hot, lbl_rest))
            
        # Stacked Sidebar Buttons
        self.create_sidebar_button("Esc", "Quit", self.on_esc_key)
        self.create_sidebar_button("I", "Import Gen", self.open_importer)
        self.create_sidebar_button("R", "Rule Studio", self.open_rule_studio)
        self.create_sidebar_button("M", "Clients", self.open_client_manager)
        
        # Authentic Tally Quit Dialog Box (Centered Popup, initially hidden)
        self.quit_frame = tk.Frame(
            self.workspace, 
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
        self.root.bind("<Up>", self.on_arrow_key)
        self.root.bind("<Down>", self.on_arrow_key)
        self.root.bind("<Return>", self.on_enter_key)
        self.root.bind("<Escape>", self.on_esc_key)
        self.root.bind("<Key>", self.on_key_typed)
        
    def show_gateway_menu(self):
        if self.importer_app:
            self.importer_app.destroy()
            self.importer_app = None
        if self.client_manager_app:
            self.client_manager_app.destroy()
            self.client_manager_app = None
            
        self.main_container.pack(fill=tk.BOTH, expand=True)
        self.setup_keyboard_navigation()
        self.root.focus_set()
        self.update_active_menu_item()
        
    def update_active_menu_item(self):
        for i in range(len(self.menu_options)):
            frame = self.menu_item_frames[i]
            lbl_hot, lbl_rest = self.menu_item_labels[i]
            
            if i == self.active_index:
                frame.config(bg=TALLY_HIGHLIGHT)
                lbl_hot.config(bg=TALLY_HIGHLIGHT, fg="#FF0000") # Dark red for hotkey on yellow highlight
                lbl_rest.config(bg=TALLY_HIGHLIGHT, fg="#000000")
            else:
                frame.config(bg=TALLY_LIST_BG)
                lbl_hot.config(bg=TALLY_LIST_BG, fg=TALLY_WARN)
                lbl_rest.config(bg=TALLY_LIST_BG, fg=TALLY_TEXT_NORMAL)
                
    def on_arrow_key(self, event):
        if self.quit_frame.winfo_viewable() or self.importer_app or self.client_manager_app:
            return
        if event.keysym == "Up":
            self.active_index = max(0, self.active_index - 1)
        else:
            self.active_index = min(len(self.menu_options) - 1, self.active_index + 1)
        self.update_active_menu_item()
        
    def on_enter_key(self, event=None):
        if self.quit_frame.winfo_viewable() or self.importer_app or self.client_manager_app:
            return
        self.menu_options[self.active_index]["action"]()
        
    def on_esc_key(self, event=None):
        if self.importer_app or self.client_manager_app:
            return
        if self.quit_frame.winfo_viewable():
            self.hide_quit_dialog()
        else:
            self.show_quit_dialog()
            
    def on_key_typed(self, event):
        if self.importer_app or self.client_manager_app:
            return
            
        if self.quit_frame.winfo_viewable():
            if event.keysym in ("y", "Y"):
                self.root.destroy()
            elif event.keysym in ("n", "N", "Escape"):
                self.hide_quit_dialog()
            return
            
        if len(event.char) == 1:
            char = event.char.upper()
            if char == "I":
                self.active_index = 0
                self.update_active_menu_item()
                self.open_importer()
            elif char == "R":
                self.active_index = 1
                self.update_active_menu_item()
                self.open_rule_studio()
            elif char == "M":
                self.active_index = 2
                self.update_active_menu_item()
                self.open_client_manager()
            elif char == "Q":
                self.active_index = 3
                self.update_active_menu_item()
                self.show_quit_dialog()
                
    def show_quit_dialog(self):
        self.quit_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        self.quit_yes_btn.focus_set()
        
    def hide_quit_dialog(self):
        self.quit_frame.place_forget()
        self.root.focus_set()
        
    def open_importer(self):
        # Hide Gateway UI container
        self.main_container.pack_forget()
        # Initialize Importer app, routing ESC on Client row to return to Gateway
        self.importer_app = TallyImporterGUI(self.root, on_exit_callback=self.show_gateway_menu)
        
    def open_client_manager(self):
        # Hide Gateway UI container
        self.main_container.pack_forget()
        # Initialize Client Manager app, routing ESC to return to Gateway
        self.client_manager_app = TallyClientManagerGUI(self.root, on_exit_callback=self.show_gateway_menu)
        
    def open_rule_studio(self):
        # Hide Gateway GUI
        self.root.withdraw()
        
        # Launch Rule Studio as a subprocess
        script_path = os.path.join(src_root, "tools", "rule_studio_gui.py")
        subprocess.run([sys.executable, script_path], cwd=project_root)
        
        # Restore Gateway GUI
        self.root.deiconify()
        self.root.focus_set()

if __name__ == "__main__":
    root = tk.Tk()
    app = GatewayGUI(root)
    root.mainloop()
