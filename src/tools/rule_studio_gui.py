import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
import json
from collections import defaultdict

# Add project root and src root to sys.path so it works when run from any folder
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
src_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if project_root not in sys.path:
    sys.path.insert(0, project_root)
if src_root not in sys.path:
    sys.path.insert(0, src_root)

from core.client_config import get_client_names, get_client, resolve_rule_path
from tools.rule_generator.chart_of_accounts import extract_chart_of_accounts
from tools.rule_generator.gemini_engine import generate_rules_with_gemini
from tools.rule_generator.tally_miner import mine_rules_from_history
from tools.rule_generator.rule_applier import append_rules
from tools.ai_rule_generator import load_unclassified_from_excel

class AutocompleteEntry(ttk.Entry):
    def __init__(self, autocompleteList, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.autocompleteList = autocompleteList
        self.var = self["textvariable"]
        if not self.var:
            self.var = tk.StringVar()
            self["textvariable"] = self.var
        self.var.trace_add("write", self.changed)
        self.bind("<Right>", self.selection)
        self.bind("<Up>", self.up)
        self.bind("<Down>", self.down)
        self.bind("<FocusOut>", self.hide_lb)
        self.lb_up = False

    def changed(self, name, index, mode):
        if self.var.get() == '':
            if self.lb_up:
                self.lb.destroy()
                self.lb_up = False
        else:
            words = self.comparison()
            if words:
                if not self.lb_up:
                    self.lb = tk.Listbox(width=self["width"])
                    self.lb.bind("<Double-Button-1>", self.selection)
                    self.lb.bind("<Button-1>", self.selection)
                    self.lb.place(x=self.winfo_x(), y=self.winfo_y() + self.winfo_height())
                    self.lb_up = True
                
                self.lb.delete(0, tk.END)
                for w in words:
                    self.lb.insert(tk.END, w)
            else:
                if self.lb_up:
                    self.lb.destroy()
                    self.lb_up = False

    def selection(self, event=None):
        if self.lb_up:
            self.var.set(self.lb.get(tk.ACTIVE))
            self.lb.destroy()
            self.lb_up = False
            self.icursor(tk.END)

    def up(self, event):
        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = int(self.lb.curselection()[0]) - 1
            if index != -1:
                self.lb.selection_clear(0, tk.END)
                self.lb.select_set(index)
                self.lb.activate(index)

    def down(self, event):
        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = int(self.lb.curselection()[0]) + 1
            if index != self.lb.size():
                self.lb.selection_clear(0, tk.END)
                self.lb.select_set(index)
                self.lb.activate(index)

    def comparison(self):
        pattern = self.var.get().lower()
        return [w for w in self.autocompleteList if pattern in w.lower()]

    def hide_lb(self, event):
        # Small delay so clicks on the listbox register before listbox is destroyed
        self.after(200, self._destroy_lb)

    def _destroy_lb(self):
        if self.lb_up:
            try:
                self.lb.destroy()
            except Exception:
                pass
            self.lb_up = False


class RuleStudioApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tally Framework — Rule Studio & AI Generator")
        self.geometry("1100x700")
        
        # Data states
        self.unclassified_txns = []
        self.coa = []
        self.clusters = {}
        self.proposals = []
        
        self.create_widgets()
        self.load_client_list()
        
    def create_widgets(self):
        # 1. Top Panel (Config & Load)
        top_frame = ttk.LabelFrame(self, text=" Configuration & Import Load ")
        top_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(top_frame, text="Client:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.client_cb = ttk.Combobox(top_frame, width=25, state="readonly")
        self.client_cb.grid(row=0, column=1, padx=5, pady=5)
        self.client_cb.bind("<<ComboboxSelected>>", self.on_client_selected)
        
        ttk.Label(top_frame, text="Bank:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.bank_cb = ttk.Combobox(top_frame, width=25, state="readonly")
        self.bank_cb.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(top_frame, text="Statement Excel:").grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        self.excel_var = tk.StringVar(value="./output/Statement_Import.xlsx")
        self.excel_entry = ttk.Entry(top_frame, textvariable=self.excel_var, width=35)
        self.excel_entry.grid(row=0, column=5, padx=5, pady=5)
        
        ttk.Button(top_frame, text="Browse...", command=self.browse_excel).grid(row=0, column=6, padx=5, pady=5)
        ttk.Button(top_frame, text="Load Data", command=self.load_data).grid(row=0, column=7, padx=5, pady=5)
        
        # 2. Main Middle Workspace
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Left Side Pane: Unclassified Clusters
        left_frame = ttk.LabelFrame(paned, text=" Unclassified Transaction Clusters ")
        paned.add(left_frame, weight=1)
        
        self.cluster_listbox = tk.Listbox(left_frame, font=("Courier", 9))
        self.cluster_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.cluster_listbox.bind("<<ListboxSelect>>", self.on_cluster_selected)
        
        # Right Side Pane: Rule Creation / Customizer
        right_frame = ttk.LabelFrame(paned, text=" Rule Editor ")
        paned.add(right_frame, weight=2)
        
        # Fields
        row = 0
        ttk.Label(right_frame, text="Pattern:").grid(row=row, column=0, padx=5, pady=5, sticky=tk.W)
        self.pattern_var = tk.StringVar()
        self.pattern_entry = ttk.Entry(right_frame, textvariable=self.pattern_var, width=50)
        self.pattern_entry.grid(row=row, column=1, padx=5, pady=5, columnspan=2, sticky=tk.W)
        
        row += 1
        ttk.Label(right_frame, text="Ledger:").grid(row=row, column=0, padx=5, pady=5, sticky=tk.W)
        self.ledger_entry = AutocompleteEntry([], right_frame, width=50)
        self.ledger_entry.grid(row=row, column=1, padx=5, pady=5, columnspan=2, sticky=tk.W)
        
        row += 1
        ttk.Label(right_frame, text="Priority:").grid(row=row, column=0, padx=5, pady=5, sticky=tk.W)
        self.priority_var = tk.StringVar(value="100")
        self.priority_entry = ttk.Entry(right_frame, textvariable=self.priority_var, width=10)
        self.priority_entry.grid(row=row, column=1, padx=5, pady=5, sticky=tk.W)
        
        row += 1
        # Optional Direction overrides
        ttk.Label(right_frame, text="Payment Ledger (OUT):").grid(row=row, column=0, padx=5, pady=5, sticky=tk.W)
        self.pay_ledger_entry = AutocompleteEntry([], right_frame, width=50)
        self.pay_ledger_entry.grid(row=row, column=1, padx=5, pady=5, columnspan=2, sticky=tk.W)
        
        row += 1
        ttk.Label(right_frame, text="Receipt Ledger (IN):").grid(row=row, column=0, padx=5, pady=5, sticky=tk.W)
        self.rec_ledger_entry = AutocompleteEntry([], right_frame, width=50)
        self.rec_ledger_entry.grid(row=row, column=1, padx=5, pady=5, columnspan=2, sticky=tk.W)

        # Amount split builder
        row += 1
        self.amount_split_var = tk.BooleanVar(value=False)
        self.amount_split_chk = ttk.Checkbutton(right_frame, text="Enable Amount Slabs / Tiers override", 
                                                variable=self.amount_split_var, command=self.toggle_amount_tiers)
        self.amount_split_chk.grid(row=row, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W)
        
        # Tiers input section
        row += 1
        self.tiers_frame = ttk.LabelFrame(right_frame, text=" Amount Tiers Configuration ")
        self.tiers_frame.grid(row=row, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        self.setup_tiers_frame()
        self.toggle_amount_tiers()
        
        # Save button
        row += 1
        self.save_btn = ttk.Button(right_frame, text="Save & Apply Rule", command=self.save_rule)
        self.save_btn.grid(row=row, column=1, padx=5, pady=10, sticky=tk.E)
        
        # 3. Bottom Panel (AI / Offline Proposals)
        bot_frame = ttk.LabelFrame(self, text=" Bulk AI Rule Proposals ")
        bot_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=5)
        
        btn_bar = ttk.Frame(bot_frame)
        btn_bar.pack(fill=tk.X, padx=5, pady=2)
        
        self.gemini_btn = ttk.Button(btn_bar, text="✨ Generate with Gemini Pro", command=self.run_gemini_autogen)
        self.gemini_btn.pack(side=tk.LEFT, padx=5, pady=2)
        
        self.miner_btn = ttk.Button(btn_bar, text="⛏️ Mine Tally History", command=self.run_history_miner)
        self.miner_btn.pack(side=tk.LEFT, padx=5, pady=2)

        self.suggest_changes_btn = ttk.Button(btn_bar, text="💡 Suggest Changes", command=self.open_suggest_changes_dialog, state=tk.DISABLED)
        self.suggest_changes_btn.pack(side=tk.LEFT, padx=5, pady=2)
        
        self.apply_proposals_btn = ttk.Button(btn_bar, text="Apply Checked Proposals", command=self.apply_proposals)
        self.apply_proposals_btn.pack(side=tk.RIGHT, padx=5, pady=2)
        
        # Treeview to show proposals
        self.tree = ttk.Treeview(bot_frame, columns=("check", "confidence", "pattern", "ledger", "reasoning"), 
                                 show="headings", height=5)
        self.tree.pack(fill=tk.X, padx=5, pady=5)
        self.tree.heading("check", text="Keep?")
        self.tree.heading("confidence", text="Conf")
        self.tree.heading("pattern", text="Pattern")
        self.tree.heading("ledger", text="Ledger")
        self.tree.heading("reasoning", text="Reasoning")
        
        self.tree.column("check", width=50, anchor=tk.CENTER)
        self.tree.column("confidence", width=80, anchor=tk.CENTER)
        self.tree.column("pattern", width=150)
        self.tree.column("ledger", width=250)
        self.tree.column("reasoning", width=400)
        self.tree.bind("<Double-Button-1>", self.toggle_tree_item)

    def setup_tiers_frame(self):
        # A simple grid of values for amount slabs
        # Tier 1
        ttk.Label(self.tiers_frame, text="Slab 1: Min:").grid(row=0, column=0, padx=2, pady=2)
        self.t1_min = ttk.Entry(self.tiers_frame, width=10)
        self.t1_min.grid(row=0, column=1, padx=2, pady=2)
        ttk.Label(self.tiers_frame, text="Max:").grid(row=0, column=2, padx=2, pady=2)
        self.t1_max = ttk.Entry(self.tiers_frame, width=10)
        self.t1_max.grid(row=0, column=3, padx=2, pady=2)
        ttk.Label(self.tiers_frame, text="Ledger:").grid(row=0, column=4, padx=2, pady=2)
        self.t1_led = AutocompleteEntry([], self.tiers_frame, width=30)
        self.t1_led.grid(row=0, column=5, padx=2, pady=2)
        
        # Tier 2
        ttk.Label(self.tiers_frame, text="Slab 2: Min:").grid(row=1, column=0, padx=2, pady=2)
        self.t2_min = ttk.Entry(self.tiers_frame, width=10)
        self.t2_min.grid(row=1, column=1, padx=2, pady=2)
        ttk.Label(self.tiers_frame, text="Max:").grid(row=1, column=2, padx=2, pady=2)
        self.t2_max = ttk.Entry(self.tiers_frame, width=10)
        self.t2_max.grid(row=1, column=3, padx=2, pady=2)
        ttk.Label(self.tiers_frame, text="Ledger:").grid(row=1, column=4, padx=2, pady=2)
        self.t2_led = AutocompleteEntry([], self.tiers_frame, width=30)
        self.t2_led.grid(row=1, column=5, padx=2, pady=2)

    def toggle_amount_tiers(self):
        state = tk.NORMAL if self.amount_split_var.get() else tk.DISABLED
        for child in self.tiers_frame.winfo_children():
            if isinstance(child, (ttk.Entry, AutocompleteEntry)):
                child.configure(state=state)

    def load_client_list(self):
        try:
            clients = get_client_names()
            self.client_cb["values"] = clients
            if clients:
                self.client_cb.current(0)
                self.on_client_selected()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load clients: {e}")

    def on_client_selected(self, event=None):
        client_name = self.client_cb.get()
        if not client_name:
            return
        
        # Load banks
        try:
            client = get_client(client_name)
            banks = list(client.get("banks", {}).keys())
            self.bank_cb["values"] = banks
            if banks:
                self.bank_cb.current(0)
            
            # Extract COA
            self.coa = extract_chart_of_accounts(client_name)
            self.ledger_entry.autocompleteList = self.coa
            self.pay_ledger_entry.autocompleteList = self.coa
            self.rec_ledger_entry.autocompleteList = self.coa
            self.t1_led.autocompleteList = self.coa
            self.t2_led.autocompleteList = self.coa
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load client details: {e}")

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.excel_var.set(path)

    def load_data(self):
        client = self.client_cb.get()
        bank = self.bank_cb.get()
        excel_path = self.excel_var.get()
        
        if not client or not bank or not excel_path:
            messagebox.showwarning("Warning", "Please configure all settings first.")
            return
            
        try:
            self.unclassified_txns = load_unclassified_from_excel(excel_path, bank)
            self.cluster_data()
            self.refresh_cluster_listbox()
            
            # Clear previous proposals on load
            self.tree.delete(*self.tree.get_children())
            self.proposals = []
            self.suggest_changes_btn.config(state=tk.DISABLED)
            messagebox.showinfo("Success", f"Loaded {len(self.unclassified_txns)} unclassified rows.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {e}")

    def cluster_data(self):
        self.clusters = defaultdict(list)
        # Helper to get keyword prefix
        for txn in self.unclassified_txns:
            desc = txn.get("description", "")
            # Simple keyword grouping
            words = [w for w in desc.replace("/", " ").replace("-", " ").split() if len(w) > 2]
            key = words[0].upper() if words else "UNKNOWN"
            self.clusters[key].append(txn)

    def refresh_cluster_listbox(self):
        self.cluster_listbox.delete(0, tk.END)
        # Sort by total transaction count descending
        sorted_keys = sorted(self.clusters.keys(), key=lambda k: len(self.clusters[k]), reverse=True)
        for key in sorted_keys:
            txns = self.clusters[key]
            total_amt = sum(t.get("amount", 0) for t in txns)
            self.cluster_listbox.insert(tk.END, f"{key:<15} ({len(txns):>3} rows, Total: ₹{total_amt:,.2f})")

    def on_cluster_selected(self, event=None):
        selection = self.cluster_listbox.curselection()
        if not selection:
            return
            
        selected_text = self.cluster_listbox.get(selection[0])
        key = selected_text.split()[0]
        
        # Populate Editor
        self.pattern_var.set(key)
        
        # Suggest baseline ledger based on prefix (or first match in COA)
        matched_ledger = ""
        for ledger in self.coa:
            if key.lower() in ledger.lower():
                matched_ledger = ledger
                break
        self.ledger_entry.var.set(matched_ledger)

    def run_history_miner(self):
        client = self.client_cb.get()
        if not self.unclassified_txns:
            messagebox.showwarning("Warning", "Please load transactions first.")
            return
            
        proposals = mine_rules_from_history(self.unclassified_txns, client)
        self.add_proposals_to_tree(proposals)

    def run_gemini_autogen(self):
        if not self.unclassified_txns:
            messagebox.showwarning("Warning", "Please load transactions first.")
            return
            
        # Call Gemini in background or direct (prompt user for API key if env is not set)
        api_key = os.environ.get("GEMINI_API_KEY")
        if not api_key:
            api_key = simpledialog.askstring("API Key Required", "Please enter your GEMINI_API_KEY:")
            if not api_key:
                return
                
        self.gemini_btn.configure(state=tk.DISABLED, text="Running...")
        self.update()
        
        client = self.client_cb.get()
        bank = self.bank_cb.get()
        feedback_reasonings = []
        try:
            rule_path = resolve_rule_path(bank, client)
            rules_dir = os.path.dirname(rule_path)
            feedback_path = os.path.join(rules_dir, "feedback_reasonings.json")
            if os.path.exists(feedback_path):
                with open(feedback_path, "r", encoding="utf-8") as f:
                    feedback_reasonings = json.load(f)
        except Exception as e:
            print(f"Error loading feedback reasonings: {e}")

        try:
            # Batch first 60 rows for speed and safety
            proposals = generate_rules_with_gemini(
                self.unclassified_txns[:60], 
                self.coa, 
                api_key=api_key,
                feedback_reasonings=feedback_reasonings
            )
            self.add_proposals_to_tree(proposals)
            messagebox.showinfo("Success", f"Received {len(proposals)} rule proposals from Gemini.")
        except Exception as e:
            messagebox.showerror("Error", f"Gemini API call failed: {e}")
        finally:
            self.gemini_btn.configure(state=tk.NORMAL, text="✨ Generate with Gemini Pro")

    def add_proposals_to_tree(self, proposals):
        for p in proposals:
            pat = ", ".join(p.get("pattern", []))
            led = p.get("ledger", "")
            conf = p.get("confidence", "HIGH")
            reason = p.get("reasoning", "")
            
            # Check if this proposal already exists in local list
            self.proposals.append(p)
            self.tree.insert("", tk.END, values=("[✓]", conf, pat, led, reason))
        
        if self.proposals:
            self.suggest_changes_btn.config(state=tk.NORMAL)

    def toggle_tree_item(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            val = self.tree.item(item, "values")
            new_check = "[ ]" if val[0] == "[✓]" else "[✓]"
            self.tree.item(item, values=(new_check, val[1], val[2], val[3], val[4]))

    def apply_proposals(self):
        client = self.client_cb.get()
        bank = self.bank_cb.get()
        rule_path = resolve_rule_path(bank, client)
        
        # Collect all checked proposals
        rules_to_save = []
        for i, item in enumerate(self.tree.get_children()):
            val = self.tree.item(item, "values")
            if val[0] == "[✓]":
                rules_to_save.append(self.proposals[i])
                
        if not rules_to_save:
            messagebox.showwarning("Warning", "No proposals are selected.")
            return
            
        try:
            added, updated = append_rules(rule_path, rules_to_save)
            messagebox.showinfo("Success", f"Applied to rules JSON.\nAdded: {added} new, Updated: {updated} existing.")
            # Clear treeview
            self.tree.delete(*self.tree.get_children())
            self.proposals = []
            self.suggest_changes_btn.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save rules: {e}")

    def save_rule(self):
        client = self.client_cb.get()
        bank = self.bank_cb.get()
        rule_path = resolve_rule_path(bank, client)
        
        pat = self.pattern_var.get().strip()
        led = self.ledger_entry.get().strip()
        priority = int(self.priority_var.get() or 100)
        
        if not pat or not led:
            messagebox.showwarning("Warning", "Pattern and Ledger fields are required.")
            return
            
        new_rule = {
            "pattern": [pat],
            "ledger": led,
            "priority": priority
        }
        
        # Direction specific overrides
        pay_led = self.pay_ledger_entry.get().strip()
        if pay_led:
            new_rule["payment_ledger"] = pay_led
            
        rec_led = self.rec_ledger_entry.get().strip()
        if rec_led:
            new_rule["receipt_ledger"] = rec_led
            
        # Amount tiers
        if self.amount_split_var.get():
            tiers = []
            # Slab 1
            if self.t1_led.get().strip():
                t1 = {"ledger": self.t1_led.get().strip()}
                if self.t1_min.get().strip():
                    t1["min"] = float(self.t1_min.get())
                if self.t1_max.get().strip():
                    t1["max"] = float(self.t1_max.get())
                tiers.append(t1)
            # Slab 2
            if self.t2_led.get().strip():
                t2 = {"ledger": self.t2_led.get().strip()}
                if self.t2_min.get().strip():
                    t2["min"] = float(self.t2_min.get())
                if self.t2_max.get().strip():
                    t2["max"] = float(self.t2_max.get())
                    
                tiers.append(t2)
            if tiers:
                new_rule["amount_tiers"] = tiers
                
        try:
            added, updated = append_rules(rule_path, [new_rule])
            messagebox.showinfo("Success", f"Rule saved to rules path.\nAdded: {added}, Updated: {updated}.")
            
            # Reset form
            self.pattern_var.set("")
            self.ledger_entry.var.set("")
            self.pay_ledger_entry.var.set("")
            self.rec_ledger_entry.var.set("")
            self.amount_split_var.set(False)
            self.toggle_amount_tiers()
            
            # Reload unclassified if loaded
            self.load_data()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save rule: {e}")

    def open_suggest_changes_dialog(self):
        client = self.client_cb.get()
        bank = self.bank_cb.get()
        if not self.proposals:
            messagebox.showwarning("Warning", "No active proposals available to edit.")
            return
        SuggestChangesDialog(self, self.proposals, self.coa, client, bank, self.on_dialog_applied)

    def on_dialog_applied(self, index, corrected_ledger, reasoning):
        p = self.proposals[index]
        p["ledger"] = corrected_ledger
        p["reasoning"] = reasoning
        
        children = self.tree.get_children()
        if index < len(children):
            item_iid = children[index]
            val = self.tree.item(item_iid, "values")
            self.tree.item(item_iid, values=(val[0], val[1], val[2], corrected_ledger, reasoning))


class SuggestChangesDialog(tk.Toplevel):
    def __init__(self, parent, proposals, coa, client_name, bank_ledger, on_apply_callback):
        super().__init__(parent)
        self.title("Suggest Changes & Correct Mappings")
        self.geometry("950x550")
        self.configure(bg="#1E1E1E")
        self.transient(parent)
        self.grab_set()
        
        self.proposals = proposals
        self.coa = coa
        self.client_name = client_name
        self.bank_ledger = bank_ledger
        self.on_apply_callback = on_apply_callback
        
        # Styling Toplevel (inherits ttk styles from parent, but Tk widgets need manual bg/fg setup)
        self.bg_main = "#1E1E1E"
        self.bg_card = "#252526"
        self.accent = "#007ACC"
        self.text_fg = "#E0E0E0"
        
        # Load feedback JSON path
        from core.client_config import resolve_rule_path
        try:
            rule_path = resolve_rule_path(self.bank_ledger, self.client_name)
            rules_dir = os.path.dirname(rule_path)
            self.feedback_path = os.path.join(rules_dir, "feedback_reasonings.json")
        except Exception:
            self.feedback_path = os.path.join(project_root, "rules", self.client_name, "feedback_reasonings.json")
            
        self.create_widgets()
        self.load_proposals()
        
    def create_widgets(self):
        # Master Grid
        self.grid_columnconfigure(0, weight=4) # Left side (List of proposals)
        self.grid_columnconfigure(1, weight=5) # Right side (Editor)
        self.grid_rowconfigure(0, weight=1)
        
        # Left Panel
        left_frame = tk.Frame(self, bg=self.bg_card, bd=1, relief=tk.SOLID)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        left_title = tk.Label(left_frame, text="Generated AI Proposals", bg=self.bg_card, fg=self.accent, font=("Segoe UI", 11, "bold"))
        left_title.pack(anchor=tk.W, padx=10, pady=10)
        
        # Treeview to display generated proposals in left panel
        self.tree = ttk.Treeview(left_frame, columns=("index", "pattern", "ledger"), show="headings", height=15)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.tree.heading("index", text="#")
        self.tree.heading("pattern", text="Pattern")
        self.tree.heading("ledger", text="AI Mapped Ledger")
        
        self.tree.column("index", width=30, anchor=tk.CENTER)
        self.tree.column("pattern", width=120)
        self.tree.column("ledger", width=200)
        self.tree.bind("<<TreeviewSelect>>", self.on_proposal_selected)
        
        # Right Panel (Editor Form)
        self.right_frame = tk.Frame(self, bg=self.bg_card, bd=1, relief=tk.SOLID)
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        
        right_title = tk.Label(self.right_frame, text="Modify Mapping & Reasoning", bg=self.bg_card, fg=self.accent, font=("Segoe UI", 11, "bold"))
        right_title.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=15, pady=10)
        
        # Form inputs
        tk.Label(self.right_frame, text="Pattern:", bg=self.bg_card, fg=self.text_fg, font=("Segoe UI", 10)).grid(row=1, column=0, sticky=tk.W, padx=15, pady=10)
        self.pattern_var = tk.StringVar()
        self.pattern_lbl = tk.Label(self.right_frame, textvariable=self.pattern_var, bg=self.bg_card, fg="#FFFFFF", font=("Segoe UI", 10, "bold"))
        self.pattern_lbl.grid(row=1, column=1, sticky=tk.W, padx=15, pady=10)
        
        tk.Label(self.right_frame, text="AI Mapped:", bg=self.bg_card, fg=self.text_fg, font=("Segoe UI", 10)).grid(row=2, column=0, sticky=tk.W, padx=15, pady=10)
        self.ai_ledger_var = tk.StringVar()
        self.ai_ledger_lbl = tk.Label(self.right_frame, textvariable=self.ai_ledger_var, bg=self.bg_card, fg="#999999", font=("Segoe UI", 10, "italic"))
        self.ai_ledger_lbl.grid(row=2, column=1, sticky=tk.W, padx=15, pady=10)
        
        tk.Label(self.right_frame, text="Correct Ledger:", bg=self.bg_card, fg=self.text_fg, font=("Segoe UI", 10)).grid(row=3, column=0, sticky=tk.W, padx=15, pady=10)
        
        self.ledger_entry = AutocompleteEntry(self.coa, self.right_frame, width=40)
        self.ledger_entry.grid(row=3, column=1, sticky=tk.W, padx=15, pady=10)
        
        tk.Label(self.right_frame, text="Actual Reasoning:\n(saved for future AI runs)", bg=self.bg_card, fg=self.text_fg, font=("Segoe UI", 9), justify=tk.LEFT).grid(row=4, column=0, sticky=tk.NW, padx=15, pady=10)
        
        self.reasoning_text = tk.Text(self.right_frame, width=40, height=5, bg="#333333", fg="#FFFFFF", insertbackground="#FFFFFF", bd=1, relief=tk.FLAT)
        self.reasoning_text.grid(row=4, column=1, sticky=tk.W, padx=15, pady=10)
        
        # Status Label
        self.status_var = tk.StringVar()
        self.status_lbl = tk.Label(self.right_frame, textvariable=self.status_var, bg=self.bg_card, fg="#4CAF50", font=("Segoe UI", 9, "bold"))
        self.status_lbl.grid(row=5, column=0, columnspan=2, pady=5)
        
        # Save Button
        self.apply_btn = ttk.Button(self.right_frame, text="💾 Apply & Save Correction", command=self.apply_selected_correction)
        self.apply_btn.grid(row=6, column=1, sticky=tk.E, padx=15, pady=10)
        self.apply_btn.config(state=tk.DISABLED)
        
        # Bottom Button Frame
        bottom_frame = tk.Frame(self, bg=self.bg_main)
        bottom_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=10)
        
        close_btn = ttk.Button(bottom_frame, text="Close Window", command=self.destroy)
        close_btn.pack(side=tk.RIGHT, padx=10)
        
    def load_proposals(self):
        # Clear
        self.tree.delete(*self.tree.get_children())
        for i, p in enumerate(self.proposals):
            pat = ", ".join(p.get("pattern", []))
            led = p.get("ledger", "")
            self.tree.insert("", tk.END, iid=str(i), values=(i+1, pat, led))
            
    def on_proposal_selected(self, event):
        selected = self.tree.selection()
        if not selected:
            return
            
        idx = int(selected[0])
        p = self.proposals[idx]
        
        pat = ", ".join(p.get("pattern", []))
        self.pattern_var.set(pat)
        
        ai_led = p.get("ledger", "")
        self.ai_ledger_var.set(ai_led)
        
        # Prepopulate correct ledger and reasoning (if they have already corrected it or AI reasoning)
        self.ledger_entry.delete(0, tk.END)
        self.ledger_entry.insert(0, ai_led)
        
        self.reasoning_text.delete("1.0", tk.END)
        self.reasoning_text.insert("1.0", p.get("reasoning", ""))
        
        self.status_var.set("")
        self.apply_btn.config(state=tk.NORMAL)
        
    def apply_selected_correction(self):
        selected = self.tree.selection()
        if not selected:
            return
            
        idx = int(selected[0])
        p = self.proposals[idx]
        pattern_str = p.get("pattern", [None])[0] or self.pattern_var.get().split(",")[0].strip()
        
        corrected_ledger = self.ledger_entry.get().strip()
        reasoning = self.reasoning_text.get("1.0", tk.END).strip()
        
        if not corrected_ledger:
            messagebox.showwarning("Warning", "Correct Ledger is required.")
            return
            
        if self.coa and (corrected_ledger not in self.coa):
            ans = messagebox.askyesno(
                "Unrecognized Ledger", 
                f"Ledger '{corrected_ledger}' is not found in the Whitelisted Chart of Accounts.\n\nDo you want to use it anyway?"
            )
            if not ans:
                return
                
        # Save correction to persistent JSON feedback file
        self.save_feedback_reasoning(pattern_str, corrected_ledger, reasoning)
        
        # Update selected in this dialog's Treeview
        self.tree.item(str(idx), values=(idx+1, ", ".join(p.get("pattern", [])), corrected_ledger))
        
        # Trigger parent update
        self.on_apply_callback(idx, corrected_ledger, reasoning)
        
        self.status_var.set("✓ Correction saved & applied to current proposals list.")
        
    def save_feedback_reasoning(self, pattern, ledger, reasoning):
        # Load existing feedback file
        os.makedirs(os.path.dirname(self.feedback_path), exist_ok=True)
        
        data = []
        if os.path.exists(self.feedback_path):
            try:
                with open(self.feedback_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception as e:
                print(f"Error loading feedback: {e}")
                
        # Update or add
        found = False
        for item in data:
            if item.get("pattern") == pattern:
                item["ledger"] = ledger
                item["reasoning"] = reasoning
                found = True
                break
                
        if not found:
            data.append({
                "pattern": pattern,
                "ledger": ledger,
                "reasoning": reasoning
            })
            
        # Write back
        try:
            with open(self.feedback_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save feedback on disk:\n{e}")

if __name__ == "__main__":
    app = RuleStudioApp()
    app.mainloop()
