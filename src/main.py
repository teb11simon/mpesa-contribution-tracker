import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path

# Ensure local modules are discoverable
sys.path.insert(0, str(Path(__file__).parent))

try:
    from mpesa_parser import MpesaParser, Transaction
    from excel_generator import ExcelGenerator
    from processor import ContributionProcessor
    from category_memory import CategoryMemory
    from member_registry import MemberRegistry
    from ocr_processor import OCRProcessor
    from settings_manager import SettingsManager
except ImportError as e:
    print(f"CRITICAL ERROR: A module is missing: {e}")

# ── Category Definitions ─────────────────────────────────────────────────
INCOME_CATEGORIES  = ["Contribution", "Benevolence", "Missions"]
EXPENSE_CATEGORIES = [
    "Worship & Admin & Welcome", "Facility Rental", "Transaction Charge",
    "Ministry/Assets", "Accommodation & Transport", "Mercy Day",
    "Lawyer/Regis", "Immigration & Flight",
    "Benevolence Expense", "Benevolence Expense Transaction Charge",
]
TRANSFER_CATEGORIES = ["Missions Transfer", "Benevolence Transfer", "Contribution Transfer"]

# ═════════════════════════════════════════════════════════════════════════
#  MEMBER MANAGEMENT
# ═════════════════════════════════════════════════════════════════════════
class MemberManagerWindow(tk.Toplevel):
    def __init__(self, master, processor, template_path=None):
        super().__init__(master)
        self.title("Manage Members (Master Ledger)")
        self.geometry("850x650")
        self.processor = processor
        self.template_path = template_path
        
        self.transient(master)
        self.grab_set()

        self._build()
        
        if self.template_path:
            self.processor.generator.load_template(self.template_path)
            self.lbl_tpl.config(text=f"Loaded: {Path(self.template_path).name}", foreground="green")
            self._refresh_list()
        else:
            self.after(100, self._prompt_template)

    def _build(self):
        # Top toolbar
        toolbar = ttk.Frame(self, padding=10)
        toolbar.pack(fill=tk.X)
        self.lbl_tpl = ttk.Label(toolbar, text="No Template Selected", foreground="red")
        self.lbl_tpl.pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Change Template", command=self._prompt_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Sync Template (Save & Sort)", command=self._sync).pack(side=tk.RIGHT, padx=5)

        # Treeview
        tree_f = ttk.Frame(self, padding=10)
        tree_f.pack(fill=tk.BOTH, expand=True)
        cols = ("Name", "Region", "Bible Talk", "Ministry", "Pledge")
        self.tree = ttk.Treeview(tree_f, columns=cols, show="headings", selectmode="browse")
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=150 if c == "Name" else 100)
        
        scroll = ttk.Scrollbar(tree_f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Controls (Add/Remove)
        ctrl_f = ttk.Frame(self, padding=10)
        ctrl_f.pack(fill=tk.X)
        
        add_f = ttk.LabelFrame(ctrl_f, text="Add New Member", padding=10)
        add_f.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        self.vars = {
            'First Name': tk.StringVar(), 'Last Name': tk.StringVar(),
            'Region': tk.StringVar(), 'Bible Talk': tk.StringVar(),
            'Ministry': tk.StringVar(value="Campus"), 'Pledge': tk.StringVar(value="0")
        }
        
        r, c = 0, 0
        for label, var in self.vars.items():
            ttk.Label(add_f, text=label).grid(row=r, column=c, padx=5, pady=2, sticky="e")
            if label == "Ministry":
                ttk.Combobox(add_f, textvariable=var, values=["Campus", "Singles", "Marrieds", "Teens"], width=13).grid(row=r, column=c+1, padx=5, pady=2)
            else:
                ttk.Entry(add_f, textvariable=var, width=15).grid(row=r, column=c+1, padx=5, pady=2)
            r += 1
            if r > 2:
                r, c = 0, c + 2
                
        ttk.Button(add_f, text="Add Member", command=self._add_member).grid(row=3, column=0, columnspan=4, pady=10)

        rem_f = ttk.LabelFrame(ctrl_f, text="Remove Member", padding=10)
        rem_f.pack(side=tk.RIGHT, fill=tk.Y, padx=5)
        
        ttk.Label(rem_f, text="First Name:").pack(anchor="w", pady=(0, 2))
        self.rem_name_var = tk.StringVar()
        ttk.Entry(rem_f, textvariable=self.rem_name_var, width=15).pack(anchor="w", pady=(0, 10))
        
        self.rem_type = tk.StringVar(value="fallaway")
        ttk.Radiobutton(rem_f, text="Fallaway", variable=self.rem_type, value="fallaway").pack(anchor="w")
        ttk.Radiobutton(rem_f, text="Moveaway", variable=self.rem_type, value="moveaway").pack(anchor="w")
        ttk.Button(rem_f, text="Remove by Name", command=self._remove_member).pack(pady=10)

    def _prompt_template(self):
        p = filedialog.askopenfilename(title="Select Master Ledger Template", filetypes=[("Excel files", "*.xlsx")])
        if p:
            self.template_path = p
            self.lbl_tpl.config(text=f"Loaded: {Path(p).name}", foreground="green")
            self.processor.generator.load_template(p)
            self._refresh_list()

    def _refresh_list(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        if not self.processor.generator.workbook: return
        members = self.processor.generator._get_members_from_combined()
        for m in members:
            full_name = f"{m.get('first_name', '')} {m.get('last_name', '')}".strip()
            self.tree.insert("", tk.END, values=(
                full_name, m.get('region', ''), 
                m.get('bible_talk', ''), m.get('ministry', ''), m.get('pledge', '')
            ), tags=(m.get('first_name', ''),)) # Tag with first name to identify easily

    def _on_ledger_updated(self, new_path):
        """Helper to sync the new working ledger across the UI and settings"""
        if not new_path: return
        self.template_path = new_path
        self.lbl_tpl.config(text=f"Loaded: {Path(new_path).name}", foreground="blue")
        
        # 1. Update the parent screen (Step3_ReportScreen)
        if hasattr(self.master, 'template_path'):
            self.master.template_path = new_path
            if hasattr(self.master, 'tpl_lbl'):
                self.master.tpl_lbl.config(text=Path(new_path).name, foreground="blue")
        
        # 2. Update global settings
        try:
            # self.master is Step3_ReportScreen, self.master.master is App
            self.master.master.settings.update("template_path", new_path)
        except:
            # Fallback for manual session handling
            pass

    def _add_member(self):
        if not self.template_path: return messagebox.showerror("Error", "Load template first.", parent=self)
        fn, ln = self.vars['First Name'].get().strip(), self.vars['Last Name'].get().strip()
        if not fn or not ln: return messagebox.showerror("Error", "First and Last names are required.", parent=self)
        
        try:
            new_path = self.processor.generator.add_member(
                self.vars['Region'].get(), self.vars['Bible Talk'].get(),
                self.vars['Ministry'].get(), fn, ln, self.vars['Pledge'].get()
            )
            self._on_ledger_updated(new_path)
            messagebox.showinfo("Success", f"Added {fn} {ln} to Combined sheet.\n\nOriginal file was preserved. Changes saved to:\n{Path(new_path).name if new_path else 'output folder'}", parent=self)
            # Clear fields
            for v in ['First Name', 'Last Name']: self.vars[v].set("")
            self._refresh_list()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add: {e}", parent=self)

    def _remove_member(self):
        if not self.template_path: return messagebox.showerror("Error", "Load template first.", parent=self)
        first_name = self.rem_name_var.get().strip()
        if not first_name: return messagebox.showwarning("Warning", "Please enter the First Name to remove.", parent=self)
        
        if messagebox.askyesno("Confirm Removal", f"Are you sure you want to remove '{first_name}' as a {self.rem_type.get()}?", parent=self):
            try:
                new_path = self.processor.generator.remove_member(first_name, self.rem_type.get())
                self._on_ledger_updated(new_path)
                messagebox.showinfo("Success", f"Removed '{first_name}'.\n\nOriginal preserved. Saved to:\n{Path(new_path).name if new_path else 'output folder'}", parent=self)
                self.rem_name_var.set("")
                self._refresh_list()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to remove: {e}", parent=self)

    def _sync(self):
        if not self.template_path: return
        try:
            new_path = self.processor.generator.sort_members()
            self._on_ledger_updated(new_path)
            messagebox.showinfo("Success", f"Master ledger has been successfully sorted and synced.\n\nChanges saved to:\n{Path(new_path).name if new_path else 'output folder'}", parent=self)
            self._refresh_list()
        except Exception as e:
            messagebox.showerror("Error", f"Sync failed: {e}", parent=self)

class ManualMatchDialog(tk.Toplevel):
    def __init__(self, master, unmatched_name, members, on_match):
        super().__init__(master)
        self.title(f"Match Name: {unmatched_name}")
        self.geometry("500x600")
        self.unmatched_name = unmatched_name
        self.members = members
        self.on_match = on_match
        
        self.transient(master)
        self.grab_set()

        self._build()

    def _build(self):
        main_f = ttk.Frame(self, padding=20)
        main_f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_f, text=f"Unmatched Name from M-Pesa/Cash:", font=("Helvetica", 10)).pack(anchor="w")
        ttk.Label(main_f, text=self.unmatched_name, font=("Helvetica", 12, "bold"), foreground="#2b4a7a").pack(anchor="w", pady=(0, 20))

        ttk.Label(main_f, text="Search Member to Match:", font=("Helvetica", 10)).pack(anchor="w")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self._filter_list())
        ttk.Entry(main_f, textvariable=self.search_var).pack(fill=tk.X, pady=(0, 10))

        # Listbox for members
        list_f = ttk.Frame(main_f)
        list_f.pack(fill=tk.BOTH, expand=True)
        
        self.listbox = tk.Listbox(list_f, font=("Helvetica", 10))
        scroll = ttk.Scrollbar(list_f, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scroll.set)
        
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._populate_list(self.members)

        btn_f = ttk.Frame(main_f, padding=(0, 20, 0, 0))
        btn_f.pack(fill=tk.X)
        
        ttk.Button(btn_f, text="Match & Save", command=self._confirm).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_f, text="Cancel", command=self.destroy).pack(side=tk.RIGHT)

    def _populate_list(self, members):
        self.listbox.delete(0, tk.END)
        for m in members:
            full_name = f"{m.get('first_name', '')} {m.get('last_name', '')}".strip()
            self.listbox.insert(tk.END, full_name)

    def _filter_list(self):
        query = self.search_var.get().lower()
        filtered = [
            m for m in self.members 
            if query in f"{m.get('first_name', '')} {m.get('last_name', '')}".lower()
        ]
        self._populate_list(filtered)

    def _confirm(self):
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showwarning("Selection Required", "Please select a member from the list.")
            return
        
        selected_name = self.listbox.get(selection[0])
        if messagebox.askyesno("Confirm Match", f"Are you sure you want to match '{self.unmatched_name}' to '{selected_name}'?\n\nThe app will remember this for future reports."):
            self.on_match(self.unmatched_name, selected_name)
            self.destroy()

# ═════════════════════════════════════════════════════════════════════════
#  STEP 1 — Upload Data
# ═════════════════════════════════════════════════════════════════════════
class Step1_UploadScreen(ttk.Frame):
    def __init__(self, master, on_next):
        super().__init__(master, padding=30)
        self.on_next = on_next
        self.pdf_path = None
        self.img_paths = {"Contribution": None, "Benevolence": None, "Missions": None}
        self.pdf_password = tk.StringVar()
        self._build()

    def _build(self):
        # Header frame for Title
        header_f = ttk.Frame(self)
        header_f.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(header_f, text="M-Pesa Tracker Pro", font=("Helvetica", 18, "bold")).pack(side=tk.LEFT)
        
        pdf_f = ttk.LabelFrame(self, text=" 1. M-Pesa Statement (PDF) ", padding=10)
        pdf_f.pack(fill=tk.X, pady=5)
        self.pdf_lbl = ttk.Label(pdf_f, text="No PDF selected")
        self.pdf_lbl.pack(side=tk.LEFT, padx=5)
        ttk.Button(pdf_f, text="Select PDF", command=self._add_pdf).pack(side=tk.RIGHT)

        pwd_f = ttk.Frame(self)
        pwd_f.pack(fill=tk.X, pady=5)
        ttk.Label(pwd_f, text="PDF Password:").pack(side=tk.LEFT)
        ttk.Entry(pwd_f, textvariable=self.pdf_password, show="*").pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        img_f = ttk.LabelFrame(self, text=" 2. Cash Record Images ", padding=10)
        img_f.pack(fill=tk.X, pady=10)
        
        for cat in self.img_paths.keys():
            row = ttk.Frame(img_f)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=f"{cat}:", width=15).pack(side=tk.LEFT)
            lbl = ttk.Label(row, text="Not uploaded", foreground="gray")
            lbl.pack(side=tk.LEFT, padx=5)
            setattr(self, f"lbl_{cat.lower()}", lbl)
            ttk.Button(row, text="Upload", command=lambda c=cat: self._add_img(c)).pack(side=tk.RIGHT)

        ttk.Button(self, text="Next: Review Data ➔", command=self._process).pack(pady=20)

    def _add_pdf(self):
        p = filedialog.askopenfilename(filetypes=[("PDF Statements", "*.pdf")])
        if p:
            self.pdf_path = p
            self.pdf_lbl.config(text=Path(p).name)

    def _add_img(self, category):
        p = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if p:
            self.img_paths[category] = p
            getattr(self, f"lbl_{category.lower()}").config(text=Path(p).name, foreground="black")

    def _process(self):
        if not self.pdf_path and not any(self.img_paths.values()):
            messagebox.showwarning("Missing Data", "Please upload at least one file.")
            return
        self.on_next(self.pdf_path, self.pdf_password.get(), self.img_paths)

# ═════════════════════════════════════════════════════════════════════════
#  STEP 2 — Categorize transactions (table view)
# ═════════════════════════════════════════════════════════════════════════
class Step2_CategorizeScreen(ttk.Frame):
    """
    Shows all transactions as a scrollable table.
    Rows = transactions.
    Columns = fixed info | one radio per category | Notes entry.
    Income rows only show income categories; expense rows only show expense categories.
    Both sets share a common Ignore column at the end.
    """

    # Shortened header labels so columns stay narrow
    INCOME_SHORT    = ["Contrib.", "Benev.", "Missions"]
    EXPENSE_SHORT   = ["Worship &\nAdmin", "Facility\nRental", "Trans.\nCharge",
                       "Ministry/\nAssets", "Accom. &\nTransport", "Mercy\nDay",
                       "Lawyer/\nRegis", "Immigr. &\nFlight",
                       "Benev.\nExpense", "Benev. Exp.\nTrans. Charge"]
    TRANSFER_SHORT  = ["Missions\nTransfer", "Benev.\nTransfer", "Contrib.\nTransfer"]

    def __init__(self, master, transactions, on_next, on_back):
        super().__init__(master, padding=10)
        self.transactions = transactions
        self.on_next      = on_next
        self.on_back      = on_back
        self.cat_vars     = [tk.StringVar() for _ in transactions]
        self.note_vars    = [tk.StringVar() for _ in transactions]

        # Pre-populate categories where known (cash uploads + M-Pesa keyword hints)
        for i, t in enumerate(transactions):
            pre = getattr(t, 'pre_category', None)
            if pre:
                self.cat_vars[i].set(pre)

        self._build()
        self._populate_table()

    # ── Column widths shared between sticky header and data rows ─────────
    COL_WIDTHS = [30, 40, 90, 90, 210,   # # Type Date Amount Details
                  70, 70, 70,             # income cats (Contrib, Benev, Missions)
                  70, 70, 70, 70, 70, 70, 70, 70, 70, 70,  # 10 expense cats
                  70, 70, 70,             # transfer cats (Missions, Benev, Contrib transfers)
                  55, 160]               # Ignore | Notes

    # ── layout ───────────────────────────────────────────────────────────
    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(top, text="← Back", command=self.on_back).pack(side=tk.LEFT)
        self.status_label = ttk.Label(top, text="", font=("Helvetica", 10))
        self.status_label.pack(side=tk.RIGHT)

        ttk.Label(self, text="Step 2 of 3  —  Categorize Transactions",
                  font=("Helvetica", 13, "bold")).pack(pady=(0, 2))
        ttk.Label(self,
                  text="Select a category for every row. "
                       "Green rows = money received  |  Red rows = payments made. "
                       "Use the Notes column to add expense detail.",
                  foreground="gray", font=("Helvetica", 9),
                  wraplength=700, justify=tk.LEFT).pack(pady=(0, 4))

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 4))

        # ── Outer container ───────────────────────────────────────────────
        outer = ttk.Frame(self)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.rowconfigure(1, weight=1)
        outer.columnconfigure(0, weight=1)

        # Sticky header canvas (scrolls horizontally in sync, never vertically)
        self.hdr_canvas = tk.Canvas(outer, borderwidth=0, highlightthickness=0)
        self.hdr_canvas.grid(row=0, column=0, columnspan=2, sticky=tk.EW)
        self.header_frame = tk.Frame(self.hdr_canvas, bg="#2b4a7a")
        self.hdr_window = self.hdr_canvas.create_window((0, 0), window=self.header_frame, anchor=tk.NW)
        self.header_frame.bind("<Configure>", self._on_header_configure)

        # Data canvas with both scrollbars
        self.canvas = tk.Canvas(outer, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient=tk.VERTICAL,   command=self.canvas.yview)
        hsb = ttk.Scrollbar(outer, orient=tk.HORIZONTAL, command=self._scroll_both_x)
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.canvas.grid(row=1, column=0, sticky=tk.NSEW)
        vsb.grid(row=1, column=1, sticky=tk.NS)
        hsb.grid(row=2, column=0, columnspan=2, sticky=tk.EW)

        self.table_frame = tk.Frame(self.canvas, bg="white")
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.table_frame, anchor=tk.NW)

        self.table_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>",      self._on_canvas_configure)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.bind_all("<MouseWheel>",    self._on_mousewheel)

        # ── Bottom bar ───────────────────────────────────────────────────
        bot = ttk.Frame(self)
        bot.pack(fill=tk.X, pady=(6, 0))
        self.finish_btn = ttk.Button(
            bot, text="✓  Finish & Continue →",
            command=self._finish, state=tk.DISABLED)
        self.finish_btn.pack(side=tk.RIGHT)
        ttk.Label(bot,
                  text="All rows must have a category selected before continuing.",
                  foreground="gray", font=("Helvetica", 9)).pack(side=tk.RIGHT, padx=10)

    def _on_header_configure(self, event=None):
        """Resize the header canvas height to fit its content."""
        self.hdr_canvas.configure(height=self.header_frame.winfo_reqheight())

    def _on_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        # Don't shrink table below its natural width
        min_w = self.table_frame.winfo_reqwidth()
        self.canvas.itemconfig(self.canvas_window,
                               width=max(event.width, min_w))

    def _scroll_both_x(self, *args):
        """Scroll data canvas and sticky header in sync horizontally."""
        self.canvas.xview(*args)
        self.hdr_canvas.xview(*args)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ── table population ─────────────────────────────────────────────────
    def _populate_table(self):
        h = self.header_frame   # sticky header (outside scroll)
        f = self.table_frame    # scrollable data rows

        HDR_BG  = "#2b4a7a"
        HDR_FG  = "white"
        INC_HDR = "#1a5c1a"
        EXP_HDR = "#7a1a1a"
        TRF_HDR = "#4a4a7a"    # dark blue-purple for transfers
        IGN_HDR = "#555555"

        # ── Column layout ────────────────────────────────────────────────
        N_FIXED   = 5
        N_INC     = len(INCOME_CATEGORIES)
        N_EXP     = len(EXPENSE_CATEGORIES)
        N_TRF     = len(TRANSFER_CATEGORIES)
        inc_start = N_FIXED
        exp_start = inc_start + N_INC
        trf_start = exp_start + N_EXP
        ign_col   = trf_start + N_TRF
        notes_col = ign_col + 1
        TOTAL_COLS = notes_col + 1

        ALL_CATS = INCOME_CATEGORIES + EXPENSE_CATEGORIES + TRANSFER_CATEGORIES + ["__IGNORE__"]

        # ── Apply shared column widths to both header and data frames ─────
        for col, w in enumerate(self.COL_WIDTHS):
            h.columnconfigure(col, minsize=w)
            f.columnconfigure(col, minsize=w)
        f.columnconfigure(4,         weight=2)
        f.columnconfigure(notes_col, weight=2)
        h.columnconfigure(4,         weight=2)
        h.columnconfigure(notes_col, weight=2)

        # ── Header row 0: group labels (into header_frame) ───────────────
        def _hdr(frame, row, col, text, bg, fg=HDR_FG, colspan=1):
            tk.Label(frame, text=text, bg=bg, fg=fg,
                     font=("Helvetica", 8, "bold"),
                     padx=4, pady=3, relief=tk.FLAT, anchor=tk.CENTER
                     ).grid(row=row, column=col, columnspan=colspan,
                            sticky=tk.EW, padx=1, pady=1)

        _hdr(h, 0, 0, "#",                    HDR_BG)
        _hdr(h, 0, 1, "Type",                 HDR_BG)
        _hdr(h, 0, 2, "Date",                 HDR_BG)
        _hdr(h, 0, 3, "Amount",               HDR_BG)
        _hdr(h, 0, 4, "Sender / Details",     HDR_BG)
        _hdr(h, 0, inc_start, "── INCOME CATEGORIES ──",    INC_HDR, colspan=N_INC)
        _hdr(h, 0, exp_start, "── EXPENSE CATEGORIES ──",   EXP_HDR, colspan=N_EXP)
        _hdr(h, 0, trf_start, "── TRANSFERS ──",            TRF_HDR, colspan=N_TRF)
        _hdr(h, 0, ign_col,   "Ignore",                     IGN_HDR)
        _hdr(h, 0, notes_col, "Notes (expense detail)",     HDR_BG)

        # ── Header row 1: individual category names ──────────────────────
        for j, short in enumerate(self.INCOME_SHORT):
            tk.Label(h, text=short, bg=INC_HDR, fg=HDR_FG,
                     font=("Helvetica", 7), padx=3, pady=2,
                     relief=tk.FLAT, anchor=tk.CENTER, justify=tk.CENTER
                     ).grid(row=1, column=inc_start + j, sticky=tk.EW, padx=1)

        for j, short in enumerate(self.EXPENSE_SHORT):
            tk.Label(h, text=short, bg=EXP_HDR, fg=HDR_FG,
                     font=("Helvetica", 7), padx=3, pady=2,
                     relief=tk.FLAT, anchor=tk.CENTER, justify=tk.CENTER
                     ).grid(row=1, column=exp_start + j, sticky=tk.EW, padx=1)

        for j, short in enumerate(self.TRANSFER_SHORT):
            tk.Label(h, text=short, bg=TRF_HDR, fg=HDR_FG,
                     font=("Helvetica", 7), padx=3, pady=2,
                     relief=tk.FLAT, anchor=tk.CENTER, justify=tk.CENTER
                     ).grid(row=1, column=trf_start + j, sticky=tk.EW, padx=1)

        tk.Label(h, text="⊘", bg=IGN_HDR, fg=HDR_FG,
                 font=("Helvetica", 9), padx=3, pady=2
                 ).grid(row=1, column=ign_col, sticky=tk.EW, padx=1)
        tk.Label(h, text="", bg=HDR_BG
                 ).grid(row=1, column=notes_col, sticky=tk.EW, padx=1)

        # ── Data rows (into table_frame) ──────────────────────────────────
        for i, t in enumerate(self.transactions):
            is_income = (t.transaction_type == "Paid In")
            row_bg  = "#eef7ee" if is_income else "#fdf0f0"
            alt_bg  = "#e2f0e2" if is_income else "#f5e5e5"
            bg      = row_bg if i % 2 == 0 else alt_bg
            grey    = "#d8d8d8"
            r       = i  # no header rows in table_frame

            def _cell(col, text, anchor=tk.W, bold=False, bg_=None):
                lbl = tk.Label(f, text=text, bg=bg_ or bg,
                               font=("Helvetica", 8, "bold" if bold else "normal"),
                               padx=5, pady=3, anchor=anchor, relief=tk.FLAT)
                lbl.grid(row=r, column=col, sticky=tk.EW, padx=1, pady=0)
                return lbl

            _cell(0, str(i + 1), anchor=tk.CENTER)
            tk.Label(f,
                     text="IN" if is_income else "OUT",
                     bg="#1a6e1a" if is_income else "#8b1a1a",
                     fg="white",
                     font=("Helvetica", 8, "bold"),
                     padx=6, pady=3).grid(row=r, column=1, sticky=tk.EW, padx=1)

            _cell(2, t.date.strftime("%Y-%m-%d\n%H:%M"), anchor=tk.CENTER)
            sign = "+" if is_income else "-"
            _cell(3, f"{sign}{abs(t.amount):,.2f}", anchor=tk.E, bold=True)

            sender   = t.sender_name or ""
            detail   = t.details or ""
            combined = f"{sender}\n{detail}".strip() if sender else detail
            tk.Label(f, text=combined, bg=bg,
                     font=("Helvetica", 8), padx=5, pady=3,
                     anchor=tk.W, relief=tk.FLAT,
                     wraplength=200, justify=tk.LEFT
                     ).grid(row=r, column=4, sticky=tk.EW, padx=1)

            # ── Tick-style category selectors ─────────────────────────────
            # indicatoron=0 turns the radiobutton into a flat toggle tile.
            # When selected it shows a green tick prefix; unselected is plain.
            for j, cat in enumerate(ALL_CATS):
                col = inc_start + j

                is_inc_cat = cat in INCOME_CATEGORIES
                is_exp_cat = cat in EXPENSE_CATEGORIES
                is_trf_cat = cat in TRANSFER_CATEGORIES
                is_ign     = (cat == "__IGNORE__")

                # Transfers show on both Paid In and Paid Out rows
                # Income cats only on Paid In rows
                # Expense cats only on Paid Out rows
                relevant = (
                    (is_income     and (is_inc_cat or is_trf_cat or is_ign)) or
                    (not is_income and (is_exp_cat or is_trf_cat or is_ign))
                )

                if not relevant:
                    tk.Label(f, text="", bg=grey
                             ).grid(row=r, column=col, sticky=tk.EW, padx=1)
                    continue

                # Colours for selected vs normal state
                if is_ign:
                    sel_color = "#c0392b"   # red when "ignore" is ticked
                    norm_bg   = "#e0e0e0"
                elif is_inc_cat:
                    sel_color = "#27ae60"   # green tick for income
                    norm_bg   = bg
                elif is_trf_cat:
                    sel_color = "#8b4589"   # purple tick for transfer
                    norm_bg   = bg
                else:  # expense
                    sel_color = "#2980b9"   # blue tick for expense
                    norm_bg   = bg

                # Short display label for the button face
                if is_ign:
                    btn_text = "✕ Ignore"
                elif is_inc_cat:
                    btn_text = "✓ " + self.INCOME_SHORT[INCOME_CATEGORIES.index(cat)].replace("\n", " ")
                elif is_trf_cat:
                    btn_text = "✓ " + self.TRANSFER_SHORT[TRANSFER_CATEGORIES.index(cat)].replace("\n", " ")
                else:
                    btn_text = "✓ " + self.EXPENSE_SHORT[EXPENSE_CATEGORIES.index(cat)].replace("\n", " ")

                rb = tk.Radiobutton(
                    f,
                    text=btn_text,
                    variable=self.cat_vars[i],
                    value=cat,
                    indicatoron=0,           # makes it a flat toggle tile, not a circle
                    selectcolor=sel_color,
                    bg=norm_bg,
                    fg="#333333",
                    activebackground=sel_color,
                    activeforeground="white",
                    font=("Helvetica", 7),
                    relief=tk.FLAT,
                    bd=1,
                    padx=3, pady=3,
                    anchor=tk.CENTER,
                    command=self._update_status,
                )
                rb.grid(row=r, column=col, sticky=tk.EW, padx=1, pady=1)

            # Notes entry
            ne = ttk.Entry(f, textvariable=self.note_vars[i], width=20)
            ne.grid(row=r, column=notes_col, sticky=tk.EW, padx=3, pady=2)

        self._update_status()

    # ── helpers ───────────────────────────────────────────────────────────
    def _update_status(self):
        total = len(self.transactions)
        done  = sum(1 for v in self.cat_vars if v.get())
        self.status_label.config(text=f"{done} of {total} categorized")
        self.progress["maximum"] = total
        self.progress["value"]   = done
        self.finish_btn.config(state=tk.NORMAL if done == total else tk.DISABLED)

    def _finish(self):
        income, expense, transfers = [], [], []
        ignored = 0
        for i, t in enumerate(self.transactions):
            cat  = self.cat_vars[i].get()
            note = self.note_vars[i].get().strip()
            if not cat:
                continue
            if cat == "__IGNORE__":
                ignored += 1
                continue
            row = {
                "receipt_no":   t.receipt_no,
                "date":         t.date,
                "details":      note if note else (t.details or ""),
                "amount":       t.amount,
                "type":         t.transaction_type,
                "name":         t.sender_name,   # used by create_report_with_matches for visitor list
                "sender_name":  t.sender_name,
                "sender_phone": t.sender_phone,
                "category":     cat,
            }
            # Route by the user-assigned category, NOT by M-Pesa type.
            # Transfers go to a separate list, income/expense routed as before
            if cat in INCOME_CATEGORIES:
                income.append(row)
            elif cat in TRANSFER_CATEGORIES:
                transfers.append(row)
            else:
                expense.append(row)

        if ignored > 0:
            messagebox.showinfo(
                "Ignored Transactions",
                f"{ignored} transaction(s) marked as ignored and excluded from the report.")
        # For now, pass transfers as part of expense list
        # (they won't appear in Income & Exp since filter checks category)
        # Future: handle transfers in a separate tab/logic
        self.on_next(income, expense + transfers)

# ═════════════════════════════════════════════════════════════════════════
#  STEP 3 — Final Report
# ═════════════════════════════════════════════════════════════════════════
class Step3_ReportScreen(ttk.Frame):
    def __init__(self, master, income_entries, expense_entries, processor, on_finish):
        super().__init__(master, padding=30)
        self.income_entries  = income_entries
        self.expense_entries = expense_entries
        self.processor       = processor
        self.on_finish       = on_finish
        self.template_path   = self.master.settings.get("template_path")
        self.output_path     = None

        self.men_var      = tk.StringVar(value="0")
        self.women_var    = tk.StringVar(value="0")
        self.children_var = tk.StringVar(value="0")
        self.report_date  = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        
        self.unmatched_names = []
        self._build()
        
        # Auto-load if we have a template from settings
        if self.template_path and os.path.exists(self.template_path):
            self.after(100, self._refresh_matches)

    def _build(self):
        ttk.Label(self, text="Finalize Sunday Report", font=("Helvetica", 18, "bold")).pack(pady=(0, 20))
        
        ref_f = ttk.LabelFrame(self, text=" 1. Last Sunday's Report (Required) ", padding=15)
        ref_f.pack(fill=tk.X, pady=10)
        
        tpl_name = Path(self.template_path).name if self.template_path else "⚠ Upload last week's report Excel to continue"
        tpl_color = "blue" if self.template_path else "red"
        
        self.tpl_lbl = ttk.Label(ref_f, text=tpl_name, foreground=tpl_color)
        self.tpl_lbl.pack(side=tk.LEFT, padx=5)
        ttk.Button(ref_f, text="⚙ Manage Members", command=self._open_member_manager).pack(side=tk.RIGHT, padx=5)
        ttk.Button(ref_f, text="Pick Last Report", command=self._browse_template).pack(side=tk.RIGHT)

        attn_f = ttk.LabelFrame(self, text=" 2. Attendance Stats ", padding=15)
        attn_f.pack(fill=tk.X, pady=10)
        row1 = ttk.Frame(attn_f)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text="Men:").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.men_var, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row1, text="Women:").pack(side=tk.LEFT, padx=(10,0))
        ttk.Entry(row1, textvariable=self.women_var, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row1, text="Children:").pack(side=tk.LEFT, padx=(10,0))
        ttk.Entry(row1, textvariable=self.children_var, width=8).pack(side=tk.LEFT, padx=5)

        cfg_f = ttk.LabelFrame(self, text=" 3. Report Date ", padding=15)
        cfg_f.pack(fill=tk.X, pady=5)
        ttk.Entry(cfg_f, textvariable=self.report_date).pack(fill=tk.X)

        match_f = ttk.LabelFrame(self, text=" 4. Review Member Matches (Optional) ", padding=15)
        match_f.pack(fill=tk.BOTH, expand=True, pady=10)
        
        ttk.Label(match_f, text="The following names from M-Pesa or Cash could not be matched to an existing member.", 
                  wraplength=700, foreground="gray", font=("Helvetica", 9)).pack(anchor="w", pady=(0, 10))

        # Treeview for unmatched names - with improved styling
        tree_f = ttk.Frame(match_f)
        tree_f.pack(fill=tk.BOTH, expand=True)
        
        # Configure Style for Treeview
        style = ttk.Style()
        style.configure("Custom.Treeview", rowheight=28, font=("Helvetica", 10))
        style.configure("Custom.Treeview.Heading", font=("Helvetica", 10, "bold"))
        
        self.match_tree = ttk.Treeview(tree_f, columns=("Name", "Amount"), show="headings", 
                                       height=8, style="Custom.Treeview")
        self.match_tree.heading("Name", text="Unmatched Name", anchor="w")
        self.match_tree.heading("Amount", text="Amount", anchor="e")
        self.match_tree.column("Name", width=550, anchor="w")
        self.match_tree.column("Amount", width=120, anchor="e")
        
        # Zebra striping tags
        self.match_tree.tag_configure('oddrow', background='#f7f9fc')
        self.match_tree.tag_configure('evenrow', background='white')
        
        match_scroll = ttk.Scrollbar(tree_f, orient="vertical", command=self.match_tree.yview)
        self.match_tree.configure(yscrollcommand=match_scroll.set)
        self.match_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        match_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        match_btn_f = ttk.Frame(match_f)
        match_btn_f.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(match_btn_f, text="Match Selected to Member", command=self._open_manual_match).pack(side=tk.LEFT)
        ttk.Button(match_btn_f, text="Refresh Matches", command=self._refresh_matches).pack(side=tk.RIGHT)

        self.btn_gen = ttk.Button(self, text="GENERATE FINAL REPORT", command=self._generate)
        self.btn_gen.pack(pady=20)

        self.post_frame = ttk.Frame(self)
        self.btn_open = ttk.Button(self.post_frame, text="📁 Open Excel", command=self._open_file)
        self.btn_reset = ttk.Button(self.post_frame, text="Start Over", command=self.on_finish)

    def _open_member_manager(self):
        if not getattr(self, 'template_path', None):
            messagebox.showwarning("Template Required", "Please select a Reference Excel file first.")
            return
        MemberManagerWindow(self, self.processor, self.template_path)

    def _browse_template(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if p:
            self.template_path = p
            self.tpl_lbl.config(text=Path(p).name, foreground="blue")
            self.master.settings.update("template_path", p)
            self._refresh_matches()

    def _refresh_matches(self):
        """Identifies which income entries are currently unmatched"""
        if not self.template_path: return
        
        try:
            # Temporarily prepare template to get members and matching engine
            self.processor.prepare_template(self.template_path)
            
            self.unmatched_names = []
            for item in self.match_tree.get_children(): self.match_tree.delete(item)
            
            seen_unmatched = set()
            count = 0
            for d in self.income_entries:
                sender_name = (d.get('sender_name') or d.get('name') or '').strip()
                if not sender_name: continue
                
                match, _ = self.processor.matching_engine.find_match(sender_name, d.get('sender_phone'), d['amount'])
                if not match:
                    if sender_name.lower() not in seen_unmatched:
                        tag = 'oddrow' if count % 2 == 0 else 'evenrow'
                        self.match_tree.insert("", tk.END, values=(sender_name, f"{d['amount']:,.2f}"), tags=(tag,))
                        seen_unmatched.add(sender_name.lower())
                        count += 1
            
            if not seen_unmatched:
                self.match_tree.insert("", tk.END, values=("(All names matched! No unmatched entries found)", ""))

        except Exception as e:
            messagebox.showerror("Match Error", f"Failed to refresh matches: {e}")

    def _open_manual_match(self):
        if not self.template_path:
            messagebox.showwarning("Template Required", "Please pick a Reference Excel file first.")
            return
            
        selection = self.match_tree.selection()
        if not selection:
            messagebox.showwarning("Selection Required", "Please select an unmatched name from the list.")
            return
            
        unmatched_name = self.match_tree.item(selection[0])['values'][0]
        if "(All names matched!" in unmatched_name: return

        members = self.processor.generator._get_members_from_combined()
        ManualMatchDialog(self, unmatched_name, members, self._save_manual_match)

    def _save_manual_match(self, unmatched_name, member_full_name):
        try:
            self.processor.matching_engine.save_alias(unmatched_name, member_full_name)
            messagebox.showinfo("Success", f"Mapping saved! '{unmatched_name}' will now match '{member_full_name}'.")
            self._refresh_matches()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save mapping: {e}")

    def _generate(self):
        if not self.template_path:
            messagebox.showerror("Error", "Please pick a Reference Excel file.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Report_{self.report_date.get()}.xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not save_path: return

        try:
            self.output_path = save_path
            attn = {"men": int(self.men_var.get() or 0), "women": int(self.women_var.get() or 0), "children": int(self.children_var.get() or 0)}
            report_dt = datetime.strptime(self.report_date.get(), "%Y-%m-%d")

            members = self.processor.prepare_template(self.template_path)

            # Map the verified amounts to the members for the Bible Talk Dashboard
            for m in members:
                m['amount'] = 0
            for d in self.income_entries:
                match, _ = self.processor.matching_engine.find_match(d['sender_name'], d.get('sender_phone'), d['amount'])
                if match:
                    for m in members:
                        if m['row_index'] == match['row_index']:
                            m['amount'] += d['amount']
                            break

            matched_mpesa, unmatched_mpesa = [], []
            matched_cash, unmatched_cash = [], []

            for d in self.income_entries:
                entry = dict(d)
                sender_name  = (d.get('sender_name') or d.get('name') or '').strip()
                sender_phone = d.get('sender_phone')
                match, _ = self.processor.matching_engine.find_match(sender_name, sender_phone, d['amount'])
                if match:
                    entry['member_row'] = match['row_index']
                    # Use the matched member's properly formatted first and last name
                    first = str(match.get('first_name', '')).strip()
                    last  = str(match.get('last_name',  '')).strip()
                    entry['name'] = f"{first} {last}".title() if (first or last) else sender_name
                    if d.get('receipt_no') == "CASH": matched_cash.append(entry)
                    else: matched_mpesa.append(entry)
                else:
                    # Preserve the original sender name for the visitor list
                    if not entry.get('name'):
                        entry['name'] = sender_name
                    if d.get('receipt_no') == "CASH": unmatched_cash.append(entry)
                    else: unmatched_mpesa.append(entry)

            # Generate the detailed Weekly Contribution tab
            self.processor.generator.create_report_with_matches(
                report_dt, self.output_path,
                matched_mpesa, unmatched_mpesa, matched_cash, unmatched_cash, members
            )

            # Finalize the other dashboards (Summary, Combined, Bible Talk, Income & Exp)
            # Before building the ledger list, ensure all income/mission entries have
            # converted names using the matching engine (some may have slipped through
            # as unmatched but should still get the properly formatted member name if matched)
            for entry in unmatched_mpesa + unmatched_cash + self.expense_entries:
                # Try to match and convert name for missions/benevolence/contribution/transfer transactions
                if (entry.get('category') or '').strip().lower() in {'missions', 'benevolence', 'contribution', 'missions transfer', 'benevolence transfer', 'contribution transfer'}:
                    sender_name = (entry.get('sender_name') or entry.get('name') or '').strip()
                    match, _ = self.processor.matching_engine.find_match(sender_name, entry.get('sender_phone'), entry.get('amount'))
                    if match:
                        first = str(match.get('first_name', '')).strip()
                        last  = str(match.get('last_name',  '')).strip()
                        entry['name'] = f"{first} {last}".title() if (first or last) else sender_name

            all_transactions_for_ledger = (
                matched_mpesa + unmatched_mpesa +
                matched_cash + unmatched_cash +
                self.expense_entries
            )
            self.processor.generator.finalize_report(
                report_dt,
                matched_mpesa + unmatched_mpesa,
                matched_cash + unmatched_cash,
                members,
                self.output_path,
                attendance=attn,
                all_transactions=all_transactions_for_ledger
            )

            self.btn_gen.config(state="disabled")
            self.post_frame.pack(pady=10)
            self.btn_open.pack(side=tk.LEFT, padx=5)
            self.btn_reset.pack(side=tk.LEFT, padx=5)
            messagebox.showinfo("Success", "Report Generated!")

        except Exception as e:
            messagebox.showerror("Error", f"Generation failed: {e}")

    def _open_file(self):
        if self.output_path and os.path.exists(self.output_path):
            if sys.platform == "win32": os.startfile(self.output_path)
            else: os.system(f'open "{self.output_path}"')
        else:
            messagebox.showerror("File Error", "File not found. Please generate it first.")

# ═════════════════════════════════════════════════════════════════════════
#  ROOT APP
# ═════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("M-Pesa Tracker Pro")
        self.geometry("900x850")
        self.memory = CategoryMemory()
        self.registry = MemberRegistry()
        self.processor = ContributionProcessor()
        self.settings = SettingsManager()
        self.ocr = OCRProcessor(ocr_backend=self.settings.get("ocr_backend", "easyocr"))
        self.parser = MpesaParser()
        self._show_step1()

    def _show_step1(self):
        for w in self.winfo_children(): w.destroy()
        Step1_UploadScreen(self, self._process_data).pack(fill=tk.BOTH, expand=True)

    def _process_data(self, pdf_path, password, img_paths):
        all_trans = []
        combined_breakdown = {}
        if pdf_path:
            try:
                parsed = self.parser.parse_pdf(pdf_path, password)
                for t in parsed:
                    # Pre-suggest category from details keywords
                    d = (t.details or '').lower()
                    if t.transaction_type == 'Paid In':
                        if 'miss' in d:          t.pre_category = 'Missions'
                        elif 'benev' in d:       t.pre_category = 'Benevolence'
                        else:                    t.pre_category = 'Contribution'
                    else:
                        if 'charge' in d:        t.pre_category = 'Transaction Charge'
                        elif 'bundle' in d or 'data' in d: t.pre_category = 'Worship & Admin & Welcome'
                    all_trans.append(t)
            except Exception as e: messagebox.showerror("PDF Error", str(e))

        for cat, path in img_paths.items():
            if path:
                try:
                    entries, breakdown = self.ocr.process_image(path)
                    for e in entries:
                        t = Transaction("CASH", datetime.now(), f"Cash: {cat}", e.amount, "Paid In", sender_name=e.name)
                        t.pre_category = cat   # pre-populate Step 2 radio button
                        all_trans.append(t)
                    for denom, total in breakdown.items():
                        combined_breakdown[denom] = combined_breakdown.get(denom, 0) + total
                except Exception as e: messagebox.showerror("OCR Error", str(e))

        self.processor.generator.cash_breakdown = combined_breakdown

        if not all_trans: return
        for w in self.winfo_children(): w.destroy()
        Step2_CategorizeScreen(self, all_trans, self._show_step3, self._show_step1).pack(fill=tk.BOTH, expand=True)

    def _show_step3(self, income, expense):
        for w in self.winfo_children(): w.destroy()
        Step3_ReportScreen(self, income, expense, self.processor, self._show_step1).pack(fill=tk.BOTH, expand=True)

if __name__ == "__main__":
    App().mainloop()