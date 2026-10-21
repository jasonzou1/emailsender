import tkinter as tk
import html
import re
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import pythoncom
import win32com.client as win32
import os
import time
import threading
import sys
import ctypes
import json
import random
from datetime import datetime, timedelta

CO_E_SERVER_EXEC_FAILURE = -2146959355  # 0x80080005
DISPID_SEND_USING_ACCOUNT = 64209

# Accepted header spellings (matched case-insensitively, ignoring spaces/underscores/hyphens)
EMAIL_HEADERS = ("email", "emailaddress", "e-mail")
FIRST_NAME_HEADERS = ("firstname", "first")
FULL_NAME_HEADERS = ("name", "fullname", "customername", "clientname", "contactname")

STATE_FILENAME = "email_sender_progress.json"
DEFAULT_BATCH_SIZE = 20
DEFAULT_DELAY_MIN = 30
DEFAULT_DELAY_MAX = 60
DEFAULT_BATCH_GAP_MINUTES = 30

# Customer status labels shown in the table
ST_SENT, ST_FAILED, ST_NEXT, ST_PENDING, ST_INVALID, ST_DUPLICATE = (
    "Sent", "Failed", "Next batch", "Pending", "Invalid email", "Duplicate")
STATUS_FILTERS = {
    "All": None,
    "Not sent yet": {ST_NEXT, ST_PENDING},
    "Sent": {ST_SENT},
    "Failed": {ST_FAILED},
    "Skipped": {ST_INVALID, ST_DUPLICATE},
}

# Matches {{Column}} (any Excel column) or {Column} (single braces, letters/digits only so CSS is left alone)
PLACEHOLDER_RE = re.compile(r'\{\{\s*([^{}<>]+?)\s*\}\}|\{([A-Za-z][A-Za-z0-9 _-]*)\}')


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def connect_outlook():
    """Attaches to the running Outlook first; only starts a new instance if none is found"""
    try:
        return win32.GetActiveObject('Outlook.Application')
    except Exception:
        return win32.Dispatch('Outlook.Application')


def explain_outlook_error(e):
    """Extra hint for the common 'Server execution failed' error"""
    if getattr(e, 'hresult', None) == CO_E_SERVER_EXEC_FAILURE or (e.args and e.args[0] == CO_E_SERVER_EXEC_FAILURE):
        if is_admin():
            return ("Cause: this program is running as Administrator but Outlook is not. "
                    "Close it and start it normally (not 'Run as administrator').")
        return ("Cause: Outlook seems to be running as Administrator, or is still starting/busy. "
                "Close Outlook completely (check Task Manager), reopen it normally, wait until it has loaded, then try again.")
    return None


def account_address(account):
    try:
        return account.SmtpAddress or account.DisplayName
    except Exception:
        return account.DisplayName


def get_accounts(outlook):
    """Returns (list of Outlook account objects, the default sending account)"""
    session = outlook.Session
    accounts = [session.Accounts.Item(i) for i in range(1, session.Accounts.Count + 1)]
    default = None
    try:
        default_store = session.DefaultStore.StoreID
        for acc in accounts:
            try:
                if acc.DeliveryStore.StoreID == default_store:
                    default = acc
                    break
            except Exception:
                pass
    except Exception:
        pass
    return accounts, default or (accounts[0] if accounts else None)


def set_send_account(mail, account):
    try:
        mail.SendUsingAccount = account
    except Exception:
        # Late-bound pywin32 sometimes refuses the property set; call it by DISPID instead
        mail._oleobj_.Invoke(DISPID_SEND_USING_ACCOUNT, 0, 8, 0, account)


# --- Helper function for script path ---
def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def clean_cell(value):
    """Returns a stripped string, or '' for empty/NaN cells"""
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_header(header):
    return str(header).strip().lower().replace(" ", "").replace("_", "").replace("-", "")


def find_column(columns, accepted):
    """Returns the first column whose normalized header is in accepted, or None"""
    accepted = {normalize_header(a) for a in accepted}
    for col in columns:
        if normalize_header(col) in accepted:
            return col
    return None


def first_name_from_full(full_name):
    """'Anna Smith' -> 'Anna', 'Smith, Anna' -> 'Anna'"""
    if "," in full_name:
        full_name = full_name.split(",", 1)[1]
    parts = full_name.split()
    return parts[0] if parts else ""


def build_values(row, columns, first_name):
    """Maps normalized column header -> cell text; {name}/{FirstName} always use the derived first name"""
    values = {normalize_header(col): clean_cell(row[col]) for col in columns}
    values["name"] = first_name
    values["firstname"] = first_name
    return values


def fill_placeholders(text, values, escape_html=False):
    """Replaces {{Column}} / {Column} with the row's values; unknown placeholders are left as-is"""
    def sub(match):
        key = normalize_header(match.group(1) or match.group(2))
        if key not in values:
            return match.group(0)
        return html.escape(values[key]) if escape_html else values[key]
    return PLACEHOLDER_RE.sub(sub, text)


def unknown_placeholders(text, known_keys):
    """Returns {{...}} placeholders in text that don't match any column"""
    found = []
    for match in PLACEHOLDER_RE.finditer(text):
        if match.group(1) and normalize_header(match.group(1)) not in known_keys and match.group(0) not in found:
            found.append(match.group(0))
    return found


def load_recipients(excel_file):
    """Reads the Excel list. Returns (columns, recipients, info) where each recipient is
    {'row': excel_row_number, 'email': str, 'name': str, 'values': dict}. Raises ValueError on bad columns."""
    df = pd.read_excel(excel_file)
    df.columns = [str(c).strip() for c in df.columns]

    email_col = find_column(df.columns, EMAIL_HEADERS)
    first_col = find_column(df.columns, FIRST_NAME_HEADERS)
    full_col = find_column(df.columns, FULL_NAME_HEADERS)
    if email_col is None or (first_col is None and full_col is None):
        raise ValueError("Excel needs an 'Email' column and a 'FirstName' (or full 'Name') column. "
                         f"Found: {', '.join(df.columns)}")

    if first_col is not None:
        info = f"Using columns: '{first_col}' (first name), '{email_col}' (email)"
    else:
        info = f"Using columns: '{full_col}' (full name, first word used), '{email_col}' (email)"

    recipients = []
    for pos, (_, row) in enumerate(df.iterrows(), start=2):  # row 1 is the header
        email = clean_cell(row[email_col])
        name = clean_cell(row[first_col]) if first_col is not None else ""
        if not name and full_col is not None:
            name = first_name_from_full(clean_cell(row[full_col]))
        recipients.append({'row': pos, 'email': email, 'name': name,
                           'values': build_values(row, df.columns, name)})
    return list(df.columns), recipients, info


def is_valid_email(email):
    return bool(email) and '@' in email


def compute_statuses(recipients, sent, failed, batch_size):
    """Returns [(recipient, status, detail)] in list order, plus counts per status"""
    rows, seen = [], set()
    counts = {s: 0 for s in (ST_SENT, ST_FAILED, ST_NEXT, ST_PENDING, ST_INVALID, ST_DUPLICATE)}
    next_left = batch_size
    for r in recipients:
        key = r['email'].lower()
        if not is_valid_email(r['email']):
            status, detail = ST_INVALID, ""
        elif key in seen:
            status, detail = ST_DUPLICATE, "Same address appears earlier in the list"
        elif key in sent:
            status, detail = ST_SENT, sent[key] or ""
        elif key in failed:
            status, detail = ST_FAILED, failed[key] or ""
        elif next_left > 0:
            status, detail = ST_NEXT, ""
            next_left -= 1
        else:
            status, detail = ST_PENDING, ""
        if is_valid_email(r['email']):
            seen.add(key)
        counts[status] += 1
        rows.append((r, status, detail))
    return rows, counts


class ProgressStore:
    """Remembers settings and which addresses were already sent, in a JSON file next to the program.
    Progress is kept per (Excel list + template) pair, so a new template starts the list from the top."""

    def __init__(self, path):
        self.path = path
        self.lock = threading.Lock()
        self.data = {"settings": {}, "campaigns": {}}
        try:
            with open(path, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
            self.data["settings"] = loaded.get("settings", {})
            self.data["campaigns"] = loaded.get("campaigns", {})
        except FileNotFoundError:
            pass
        except Exception:
            # Keep the unreadable file instead of silently overwriting the user's progress
            os.replace(path, path + ".broken")
        for c in self.data["campaigns"].values():
            # Older versions stored plain lists of addresses
            for field in ("sent", "failed"):
                if isinstance(c.get(field), list):
                    c[field] = {e: None for e in c[field]}

    @staticmethod
    def key(excel_file, template_file):
        norm = lambda p: os.path.normcase(os.path.abspath(p))
        return norm(excel_file) + "|" + norm(template_file)

    def _campaign(self, excel_file, template_file):
        return self.data["campaigns"].setdefault(self.key(excel_file, template_file), {
            "excel": excel_file, "template": template_file, "sent": {}, "failed": {}})

    def campaign(self, excel_file, template_file):
        """Returns copies of ({email: sent_time}, {email: error}) for this list + template"""
        with self.lock:
            c = self.data["campaigns"].get(self.key(excel_file, template_file))
            if not c:
                return {}, {}
            return dict(c["sent"]), dict(c["failed"])

    def mark(self, excel_file, template_file, email, ok, error=None):
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        with self.lock:
            c = self._campaign(excel_file, template_file)
            if ok:
                c["sent"][email.lower()] = now
                c["failed"].pop(email.lower(), None)
            else:
                c["failed"][email.lower()] = f"{now}: {error}"
            c["updated"] = now
        self.save()

    def reset(self, excel_file, template_file):
        with self.lock:
            self.data["campaigns"].pop(self.key(excel_file, template_file), None)
        self.save()

    def get_setting(self, name, default=None):
        with self.lock:
            return self.data["settings"].get(name, default)

    def set_settings(self, **values):
        with self.lock:
            self.data["settings"].update(values)
        self.save()

    def save(self):
        """Atomic write so a crash or power cut never leaves a half-written file"""
        with self.lock:
            tmp = self.path + ".tmp"
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, indent=2)
            os.replace(tmp, self.path)


class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("My Baskets Email Sender (HTML/OFT Supported)") # 软件标题
        self.root.geometry("820x720")
        self.root.minsize(700, 600)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Stored variables
        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.store = ProgressStore(os.path.join(get_script_dir(), STATE_FILENAME))
        self.subject_template = tk.StringVar(value=self.store.get_setting("subject", "Hello {{FirstName}} From My Baskets"))
        self.batch_size = tk.StringVar(value=str(self.store.get_setting("batch_size", DEFAULT_BATCH_SIZE)))
        self.delay_min = tk.StringVar(value=str(self.store.get_setting("delay_min", self.store.get_setting("delay_seconds", DEFAULT_DELAY_MIN))))
        self.delay_max = tk.StringVar(value=str(self.store.get_setting("delay_max", DEFAULT_DELAY_MAX)))
        self.batch_gap = tk.StringVar(value=str(self.store.get_setting("batch_gap_minutes", DEFAULT_BATCH_GAP_MINUTES)))
        self.auto_continue = tk.BooleanVar(value=self.store.get_setting("auto_continue", False))
        self.countdown_text = tk.StringVar(value="")
        self.next_batch_at = None  # datetime while waiting between batches in auto mode
        self.sender = tk.StringVar(value=self.store.get_setting("sender", ""))
        self.outlook_status = tk.StringVar(value="Not checked")
        self.status_text = tk.StringVar(value="Select a customer list and template.")
        self.list_filter = tk.StringVar(value="All")
        self.is_running = False
        self.is_checking = False
        self.recipients = []  # cached customer list for the table
        self.remaining_count = 0
        self._status_job = None

        # --- UI Layout ---

        # 0. Outlook connection + sending account
        frame_ol = tk.LabelFrame(root, text="Outlook Connection", padx=10, pady=5) # 标签
        frame_ol.pack(fill="x", padx=10, pady=3)
        tk.Label(frame_ol, text="Send from:").pack(side=tk.LEFT) # 标签
        self.sender_box = ttk.Combobox(frame_ol, textvariable=self.sender, state="readonly", width=34)
        self.sender_box.pack(side=tk.LEFT, padx=5)
        self.btn_check = tk.Button(frame_ol, text="Check Connection", command=self.check_connection) # 按钮
        self.btn_check.pack(side=tk.LEFT, padx=5)
        self.lbl_outlook = tk.Label(frame_ol, textvariable=self.outlook_status, fg="gray", font=("Arial", 9, "bold"))
        self.lbl_outlook.pack(side=tk.LEFT, padx=5)
        self.sender.trace_add("write", lambda *_: self.store.set_settings(sender=self.sender.get()))

        # 1. Files
        frame_files = tk.LabelFrame(root, text="1. Files", padx=10, pady=5) # 标签
        frame_files.pack(fill="x", padx=10, pady=3)
        frame_files.columnconfigure(1, weight=1)
        tk.Label(frame_files, text="Customer list (Excel):").grid(row=0, column=0, sticky="w") # 标签
        tk.Entry(frame_files, textvariable=self.excel_path).grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        tk.Button(frame_files, text="Browse...", command=self.select_excel).grid(row=0, column=2) # 按钮
        tk.Label(frame_files, text="Template (.oft / .html):").grid(row=1, column=0, sticky="w") # 标签
        tk.Entry(frame_files, textvariable=self.template_path).grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        tk.Button(frame_files, text="Browse...", command=self.select_template).grid(row=1, column=2) # 按钮

        # 2. Subject + batch settings
        frame_mail = tk.LabelFrame(root, text="2. Subject ({{FirstName}}, {{Company}}, any column) and Batch Settings", padx=10, pady=5) # 标签
        frame_mail.pack(fill="x", padx=10, pady=3)
        tk.Entry(frame_mail, textvariable=self.subject_template).pack(fill="x", pady=(0, 5))
        frame_batch = tk.Frame(frame_mail)
        frame_batch.pack(fill="x")
        tk.Label(frame_batch, text="Emails per batch:").pack(side=tk.LEFT) # 标签
        tk.Spinbox(frame_batch, from_=1, to=1000, textvariable=self.batch_size, width=5).pack(side=tk.LEFT, padx=(5, 20))
        tk.Label(frame_batch, text="Random wait between emails:").pack(side=tk.LEFT) # 标签
        tk.Spinbox(frame_batch, from_=0, to=600, textvariable=self.delay_min, width=5).pack(side=tk.LEFT, padx=5)
        tk.Label(frame_batch, text="to").pack(side=tk.LEFT) # 标签
        tk.Spinbox(frame_batch, from_=0, to=600, textvariable=self.delay_max, width=5).pack(side=tk.LEFT, padx=5)
        tk.Label(frame_batch, text="seconds").pack(side=tk.LEFT) # 标签
        frame_gap = tk.Frame(frame_mail)
        frame_gap.pack(fill="x", pady=(5, 0))
        tk.Label(frame_gap, text="Minutes between batches:").pack(side=tk.LEFT) # 标签
        tk.Spinbox(frame_gap, from_=0, to=1440, textvariable=self.batch_gap, width=5).pack(side=tk.LEFT, padx=(5, 20))
        tk.Checkbutton(frame_gap, text="Automatically send the next batch after the gap (until the list is done)",
                       variable=self.auto_continue, command=self.update_button_text).pack(side=tk.LEFT) # 标签
        self.batch_size.trace_add("write", lambda *_: self.on_batch_size_change())

        # 3. Progress + Action Buttons
        frame_btn = tk.Frame(root, pady=5)
        frame_btn.pack()
        tk.Label(frame_btn, textvariable=self.status_text, font=("Arial", 10, "bold")).pack() # 标签
        tk.Label(frame_btn, textvariable=self.countdown_text, fg="#1565c0").pack(pady=(0, 5)) # 标签
        self.btn_start = tk.Button(frame_btn, text="Send Next Batch", command=self.toggle_sending, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), width=18) # 按钮
        self.btn_start.pack(side=tk.LEFT, padx=5)
        self.btn_reset = tk.Button(frame_btn, text="Reset Progress", command=self.reset_progress) # 按钮
        self.btn_reset.pack(side=tk.LEFT, padx=5)

        # 4. Customers table + Execution Log (tabs)
        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True, padx=10, pady=5)

        tab_customers = tk.Frame(notebook)
        notebook.add(tab_customers, text="Customers") # 标签
        bar = tk.Frame(tab_customers)
        bar.pack(fill="x", pady=3)
        tk.Button(bar, text="Load / Refresh List", command=self.load_customer_list).pack(side=tk.LEFT) # 按钮
        tk.Label(bar, text="Show:").pack(side=tk.LEFT, padx=(15, 5)) # 标签
        filter_box = ttk.Combobox(bar, textvariable=self.list_filter, values=list(STATUS_FILTERS), state="readonly", width=14)
        filter_box.pack(side=tk.LEFT)
        filter_box.bind("<<ComboboxSelected>>", lambda _: self.refresh_status())

        table_frame = tk.Frame(tab_customers)
        table_frame.pack(fill="both", expand=True)
        columns = ("row", "name", "email", "status", "detail")
        self.table = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)
        for col, title, width, stretch in (("row", "Row", 50, False), ("name", "First Name", 120, False),
                                           ("email", "Email", 230, True), ("status", "Status", 100, False),
                                           ("detail", "Sent Time / Error", 220, True)):
            self.table.heading(col, text=title)
            self.table.column(col, width=width, stretch=stretch, anchor="w")
        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=scroll.set)
        self.table.pack(side=tk.LEFT, fill="both", expand=True)
        scroll.pack(side=tk.RIGHT, fill="y")
        self.table.tag_configure(ST_SENT, foreground="#2e7d32")
        self.table.tag_configure(ST_FAILED, foreground="#c62828")
        self.table.tag_configure(ST_NEXT, foreground="#1565c0")
        self.table.tag_configure(ST_INVALID, foreground="gray")
        self.table.tag_configure(ST_DUPLICATE, foreground="gray")

        tab_log = tk.Frame(notebook)
        notebook.add(tab_log, text="Execution Log") # 标签
        self.log_area = scrolledtext.ScrolledText(tab_log, height=10)
        self.log_area.pack(fill="both", expand=True)

        self.auto_load_files()
        self.excel_path.trace_add("write", lambda *_: self.schedule_list_reload())
        self.template_path.trace_add("write", lambda *_: self.schedule_list_reload())
        self.update_button_text()
        self.load_customer_list()
        if is_admin():
            self.log("Warning: running as Administrator. If Outlook is not also running as Administrator, "
                     "the connection will fail - start this program normally instead.") # 日志
        self.root.after(300, self.check_connection)
        self.tick_countdown()

    # --- Thread-safe UI helpers (Tkinter must only be touched from the main thread) ---

    def ui(self, func, *args):
        self.root.after(0, func, *args)

    def log(self, message):
        """Adds message to the log area (safe to call from any thread)"""
        if threading.current_thread() is threading.main_thread():
            self._append_log(message)
        else:
            self.ui(self._append_log, message)

    def _append_log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)

    def auto_load_files(self):
        """Reopens the files used last time, otherwise auto-detects files in the program's folder"""
        last_excel = self.store.get_setting("excel")
        last_template = self.store.get_setting("template")
        if last_excel and last_template and os.path.isfile(last_excel) and os.path.isfile(last_template):
            self.excel_path.set(last_excel)
            self.template_path.set(last_template)
            self.log(f"Reopened last Excel: {last_excel}") # 日志
            self.log(f"Reopened last Template: {last_template}") # 日志
            return

        cwd = get_script_dir()
        default_excel = os.path.join(cwd, "clients.xlsx")
        default_tmpl_oft = os.path.join(cwd, "template.oft")
        default_tmpl_html = os.path.join(cwd, "template.html")

        if os.path.exists(default_excel):
            self.excel_path.set(default_excel)
            self.log(f"Auto-loaded Excel: {default_excel}") # 日志

        if os.path.exists(default_tmpl_oft):
            self.template_path.set(default_tmpl_oft)
            self.log(f"Auto-loaded OFT Template: {default_tmpl_oft}") # 日志
        elif os.path.exists(default_tmpl_html):
            self.template_path.set(default_tmpl_html)
            self.log(f"Auto-loaded HTML Template: {default_tmpl_html}") # 日志

    def select_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            self.excel_path.set(filename)

    def select_template(self):
        filename = filedialog.askopenfilename(filetypes=[("Email Templates", "*.oft *.html *.htm")])
        if filename:
            self.template_path.set(filename)

    # --- Outlook connection check ---

    def check_connection(self):
        if self.is_checking or self.is_running:
            return
        self.is_checking = True
        self.btn_check.config(state="disabled")
        self.set_outlook_status("Checking...", "gray")
        threading.Thread(target=self._check_connection_worker, daemon=True).start()

    def _check_connection_worker(self):
        pythoncom.CoInitialize()
        try:
            outlook = connect_outlook()
            accounts, default = get_accounts(outlook)
            addresses = [account_address(a) for a in accounts]
            default_addr = account_address(default) if default else ""
            self.ui(self._on_connection_ok, addresses, default_addr)
        except Exception as e:
            self.log(f"Outlook check failed: {e}") # 日志
            hint = explain_outlook_error(e)
            if hint:
                self.log(hint) # 日志
            self.ui(self._on_connection_failed)
        finally:
            pythoncom.CoUninitialize()

    def _on_connection_ok(self, addresses, default_addr):
        self.is_checking = False
        self.btn_check.config(state="normal")
        self.sender_box.config(values=addresses)
        if not addresses:
            self.set_outlook_status("Connected, but no email accounts found", "#c62828")
            self.log("Outlook is running but has no email accounts set up.") # 日志
            return
        # Keep the remembered sender if it still exists, otherwise use Outlook's default account
        if self.sender.get().lower() not in [a.lower() for a in addresses]:
            self.sender.set(default_addr or addresses[0])
        self.set_outlook_status(f"Connected ({len(addresses)} account(s))", "#2e7d32")
        self.log(f"Outlook connected. Accounts: {', '.join(addresses)}. Default: {default_addr}") # 日志

    def _on_connection_failed(self):
        self.is_checking = False
        self.btn_check.config(state="normal")
        self.set_outlook_status("Not connected - see Execution Log", "#c62828")

    def set_outlook_status(self, text, color):
        self.outlook_status.set(text)
        self.lbl_outlook.config(fg=color)

    # --- Customer list + progress ---

    def get_batch_settings(self):
        """Returns (batch_size, delay_min, delay_max, gap_minutes) or None if the inputs aren't valid"""
        try:
            batch = int(self.batch_size.get())
            delay_min = float(self.delay_min.get())
            delay_max = float(self.delay_max.get())
            gap = float(self.batch_gap.get())
        except ValueError:
            return None
        if batch < 1 or delay_min < 0 or delay_max < delay_min or gap < 0:
            return None
        return batch, delay_min, delay_max, gap

    def settings_to_store(self, settings):
        batch, delay_min, delay_max, gap = settings
        return {"batch_size": batch, "delay_min": delay_min, "delay_max": delay_max,
                "batch_gap_minutes": gap, "auto_continue": self.auto_continue.get()}

    def update_button_text(self):
        if self.is_running:
            return
        settings = self.get_batch_settings()
        if not settings:
            text = "Send Next Batch"
        elif self.auto_continue.get():
            text = f"Start Auto ({settings[0]} per batch)"
        else:
            text = f"Send Next {settings[0]}"
        self.btn_start.config(text=text) # 按钮文本

    def gap_remaining(self, gap_minutes):
        """Time left before the gap after the last finished batch is over, or None"""
        last = self.store.get_setting("last_batch_end")
        if not last or gap_minutes <= 0:
            return None
        try:
            ready = datetime.fromisoformat(last) + timedelta(minutes=gap_minutes)
        except ValueError:
            return None
        left = ready - datetime.now()
        return left if left.total_seconds() > 0 else None

    @staticmethod
    def fmt_delta(delta):
        secs = int(delta.total_seconds())
        return f"{secs // 3600}:{secs % 3600 // 60:02d}:{secs % 60:02d}" if secs >= 3600 else f"{secs // 60}:{secs % 60:02d}"

    def tick_countdown(self):
        """Updates the 'next batch' line once a second"""
        if self.next_batch_at is not None:
            left = self.next_batch_at - datetime.now()
            self.countdown_text.set(f"Waiting between batches - next batch starts at "
                                    f"{self.next_batch_at:%H:%M} (in {self.fmt_delta(max(left, timedelta(0)))})")
        elif not self.is_running:
            settings = self.get_batch_settings()
            left = self.gap_remaining(settings[3]) if settings and self.remaining_count else None
            show = left is not None and left.total_seconds() >= 1
            self.countdown_text.set(f"Batch gap: next batch recommended in {self.fmt_delta(left)}" if show else "")
        else:
            self.countdown_text.set("")
        self.root.after(1000, self.tick_countdown)

    def on_batch_size_change(self):
        self.update_button_text()
        if not self.is_running:
            self.refresh_status()

    def schedule_list_reload(self):
        # Wait until the user stops typing a path before re-reading the Excel file
        if self._status_job is not None:
            self.root.after_cancel(self._status_job)
        self._status_job = self.root.after(500, self.load_customer_list)

    def load_customer_list(self):
        """Reads the Excel file into the Customers table (main thread only)"""
        self._status_job = None
        excel_file = self.excel_path.get().strip()
        self.recipients = []
        if not os.path.isfile(excel_file):
            self.status_text.set("Select a customer list and template.")
            self.table.delete(*self.table.get_children())
            return
        try:
            _, self.recipients, _ = load_recipients(excel_file)
        except Exception as e:
            self.status_text.set("Can't read the customer list - see Execution Log")
            self.log(f"Error reading customer list: {e}") # 日志
            self.table.delete(*self.table.get_children())
            return
        self.refresh_status()

    def refresh_status(self):
        """Re-marks every customer as Sent / Failed / Next batch / Pending and updates the summary"""
        excel_file = self.excel_path.get().strip()
        template_file = self.template_path.get().strip()
        if not self.recipients:
            return
        if os.path.isfile(template_file):
            sent, failed = self.store.campaign(excel_file, template_file)
        else:
            sent, failed = {}, {}
        settings = self.get_batch_settings()
        rows, counts = compute_statuses(self.recipients, sent, failed, settings[0] if settings else DEFAULT_BATCH_SIZE)

        show = STATUS_FILTERS.get(self.list_filter.get())
        self.table.delete(*self.table.get_children())
        for r, status, detail in rows:
            if show is None or status in show:
                self.table.insert("", tk.END, values=(r['row'], r['name'], r['email'], status, detail), tags=(status,))

        total = counts[ST_SENT] + counts[ST_FAILED] + counts[ST_NEXT] + counts[ST_PENDING]
        remaining = counts[ST_NEXT] + counts[ST_PENDING]
        self.remaining_count = remaining
        skipped = counts[ST_INVALID] + counts[ST_DUPLICATE]
        text = f"Sent {counts[ST_SENT]} of {total}  |  Failed {counts[ST_FAILED]}  |  Remaining {remaining}"
        if skipped:
            text += f"  |  Skipped {skipped}"
        if not os.path.isfile(template_file):
            text += "  (select a template to see progress)"
        self.status_text.set(text)

    def reset_progress(self):
        excel_file = self.excel_path.get().strip()
        template_file = self.template_path.get().strip()
        if self.is_running or not excel_file or not template_file:
            return
        if messagebox.askyesno("Reset Progress",
                               "Forget which customers were already emailed for this list + template?\n\n"
                               "The next batch will start again from the top of the list."): # 提示框
            self.store.reset(excel_file, template_file)
            self.log("Progress reset. Next batch starts from the top of the list.") # 日志
            self.refresh_status()

    # --- Sending ---

    def toggle_sending(self):
        if self.is_running:
            self.is_running = False
            self.btn_start.config(state="disabled", text="Stopping...") # 按钮文本
        else:
            self.start_thread()

    def start_thread(self):
        """Starts the sending process in a separate thread"""
        excel_file = self.excel_path.get().strip()
        template_file = self.template_path.get().strip()

        if not excel_file or not template_file:
            messagebox.showwarning("Warning", "Please select both Excel and Template files!") # 提示框
            return
        for path in (excel_file, template_file):
            if not os.path.isfile(path):
                messagebox.showwarning("Warning", f"File not found:\n{path}") # 提示框
                return
        if self.is_checking:
            messagebox.showwarning("Warning", "Still checking the Outlook connection, please wait a moment.") # 提示框
            return
        sender = self.sender.get().strip()
        if not sender:
            messagebox.showwarning("Warning", "No sending account selected. Click 'Check Connection' first.") # 提示框
            return

        settings = self.get_batch_settings()
        if settings is None:
            messagebox.showwarning("Warning", "Please check the batch settings:\n\n"
                                   "- Emails per batch: 1 or more\n"
                                   "- Random wait: numbers 0 or more, and the second must not be smaller than the first\n"
                                   "- Minutes between batches: 0 or more") # 提示框
            return
        batch_size, delay_min, delay_max, gap = settings
        auto = self.auto_continue.get()

        left = self.gap_remaining(gap)
        if left and not messagebox.askyesno(
                "Batch Gap", f"The last batch finished less than {gap:g} minutes ago.\n"
                             f"The gap ends in {self.fmt_delta(left)}.\n\nSend anyway?"): # 提示框
            return

        wait_text = f"{delay_min:g}-{delay_max:g} seconds" if delay_max > delay_min else f"{delay_min:g} seconds"
        if auto:
            question = (f"Send ALL remaining emails from:\n\n{sender}\n\n"
                        f"in batches of {batch_size}, waiting {wait_text} between emails "
                        f"and {gap:g} minutes between batches.\n\nKeep this program and Outlook open. Continue?")
        else:
            question = (f"Send the next {batch_size} email(s) from:\n\n{sender}\n\n"
                        f"waiting {wait_text} between emails.\n\nContinue?")
        if not messagebox.askyesno("Confirm", question): # 提示框
            return

        # Read Tk variables here, in the main thread
        subject_tmpl = self.subject_template.get()
        self.store.set_settings(excel=excel_file, template=template_file, subject=subject_tmpl,
                                sender=sender, **self.settings_to_store(settings))

        self.is_running = True
        self.btn_start.config(text="Stop", bg="#f44336") # 按钮文本
        self.btn_reset.config(state="disabled")
        self.btn_check.config(state="disabled")
        thread = threading.Thread(target=self.run_sending_process,
                                  args=(excel_file, template_file, subject_tmpl, batch_size,
                                        delay_min, delay_max, gap, auto, sender))
        thread.daemon = True
        thread.start()

    def run_sending_process(self, excel_file, template_file, subject_tmpl, batch_size,
                            delay_min, delay_max, gap_minutes, auto, sender):
        # COM must be initialized on every thread that uses it
        pythoncom.CoInitialize()
        try:
            self.log("-" * 30)
            self.log("Reading Excel file...") # 日志
            try:
                columns, recipients, info = load_recipients(excel_file)
            except ValueError as e:
                self.log(f"Error: {e}") # 日志
                return
            self.log(info) # 日志

            # --- Load HTML content early (if using HTML template) ---
            html_content = None
            if template_file.lower().endswith(('.html', '.htm')):
                self.log("HTML template detected. Loading content...") # 日志
                try:
                    with open(template_file, 'r', encoding='utf-8-sig') as f:
                        html_content = f.read()
                except Exception as e:
                    self.log(f"Error: Could not read HTML template file. {e}") # 日志
                    return

            try:
                outlook = connect_outlook()
                accounts, _ = get_accounts(outlook)
            except Exception as e:
                self.log(f"Error: Could not connect to Outlook. Ensure Classic Outlook is open. ({e})") # 日志
                hint = explain_outlook_error(e)
                if hint:
                    self.log(hint) # 日志
                return

            account = next((a for a in accounts if account_address(a).lower() == sender.lower()), None)
            if account is None:
                self.log(f"Error: The account '{sender}' was not found in Outlook. Click 'Check Connection' and pick an account.") # 日志
                return
            self.log(f"Sending from: {sender}") # 日志

            # --- Make sure every {{placeholder}} matches an Excel column before sending anything ---
            body_text = html_content if html_content is not None else (outlook.CreateItemFromTemplate(template_file).HTMLBody or "")
            known_keys = {normalize_header(c) for c in columns} | {"name", "firstname"}
            unknown = unknown_placeholders(subject_tmpl + body_text, known_keys)
            if unknown:
                self.log(f"Error: No Excel column for placeholder(s): {', '.join(unknown)}. Found columns: {', '.join(columns)}") # 日志
                return

            batch_no = 0
            total_sent = total_failed = 0
            while self.is_running:
                batch, remaining_total = self.next_batch(recipients, excel_file, template_file, batch_size)
                if not batch:
                    if batch_no == 0:
                        self.log("Nothing left to send: every address in this list has already been emailed with this template.") # 日志
                        self.log("Use 'Reset Progress' to send to the whole list again.") # 日志
                        self.ui(messagebox.showinfo, "All Done", "Every customer in this list has already been emailed.") # 提示框
                    break
                batch_no += 1
                self.log(f"Batch {batch_no}: {len(batch)} email(s), {remaining_total} remaining in the list.") # 日志

                sent_n, failed_n = self.send_batch(outlook, account, batch, excel_file, template_file,
                                                   subject_tmpl, html_content, delay_min, delay_max)
                self.store.set_settings(last_batch_end=datetime.now().isoformat(timespec="seconds"))
                total_sent += sent_n
                total_failed += failed_n
                left = remaining_total - sent_n - failed_n
                summary = f"Sent: {sent_n}, Failed: {failed_n}. Remaining in list: {left}"
                self.log(f"Batch {batch_no} finished. {summary}") # 日志

                if not self.is_running:
                    break
                if not auto or left <= 0:
                    if batch_no > 1:
                        summary = (f"{batch_no} batches: Sent {total_sent}, Failed {total_failed}. "
                                   f"Remaining in list: {left}")
                    if left <= 0:
                        self.ui(messagebox.showinfo, "All Done", f"Every customer in the list has been emailed!\n{summary}") # 提示框
                    else:
                        self.ui(messagebox.showinfo, "Batch Complete", f"Batch finished!\n{summary}") # 提示框
                    break

                # Auto mode: wait out the gap, then loop to the next batch
                self.next_batch_at = datetime.now() + timedelta(minutes=gap_minutes)
                self.log(f"Waiting {gap_minutes:g} minute(s); next batch starts at {self.next_batch_at:%H:%M}.") # 日志
                self.wait(gap_minutes * 60)
                self.next_batch_at = None

            if not self.is_running:
                self.log("Stopped. Progress is saved; the next batch continues from here.") # 日志

        except Exception as e:
            self.log(f"Critical Error: {e}") # 日志
            self.ui(messagebox.showerror, "Error", str(e)) # 提示框

        finally:
            self.next_batch_at = None
            pythoncom.CoUninitialize()
            self.ui(self.reset_ui)

    def next_batch(self, recipients, excel_file, template_file, batch_size):
        """Valid addresses not already sent/failed, first occurrence only. Returns (batch, remaining_total)"""
        sent, failed = self.store.campaign(excel_file, template_file)
        done = set(sent) | set(failed)
        batch, seen = [], set()
        for r in recipients:
            key = r['email'].lower()
            if not is_valid_email(r['email']) or key in done or key in seen:
                continue
            seen.add(key)
            batch.append(r)
            if len(batch) >= batch_size:
                break
        remaining_total = len({r['email'].lower() for r in recipients if is_valid_email(r['email'])} - done)
        return batch, remaining_total

    def send_batch(self, outlook, account, batch, excel_file, template_file, subject_tmpl, html_content,
                   delay_min, delay_max):
        """Sends one batch with a random wait between emails. Returns (sent, failed)"""
        success_count = 0
        failed_count = 0
        total = len(batch)
        for pos, r in enumerate(batch, start=1):
            if not self.is_running:
                break

            email, name, values = r['email'], r['name'], r['values']
            try:
                # Create Mail Item
                if html_content is not None:
                    mail = outlook.CreateItem(0)
                    mail.HTMLBody = fill_placeholders(html_content, values, escape_html=True)
                else:
                    mail = outlook.CreateItemFromTemplate(template_file)
                    if mail.HTMLBody:
                        mail.HTMLBody = fill_placeholders(mail.HTMLBody, values, escape_html=True)

                set_send_account(mail, account)
                mail.To = email
                mail.Subject = fill_placeholders(subject_tmpl, values)
                mail.Send()

                self.store.mark(excel_file, template_file, email, ok=True)
                self.log(f"[{pos}/{total}] Sent to: {name} ({email}) - Excel row {r['row']}") # 日志
                success_count += 1

            except Exception as e:
                self.store.mark(excel_file, template_file, email, ok=False, error=str(e))
                self.log(f"[{pos}/{total}] Failed to send ({email}) - Excel row {r['row']}: {e}") # 日志
                failed_count += 1

            self.ui(self.refresh_status)
            if pos < total and self.is_running:
                pause = random.uniform(delay_min, delay_max)
                self.log(f"    waiting {pause:.0f}s...") # 日志
                self.wait(pause)
        return success_count, failed_count

    def wait(self, seconds):
        """Sleeps in small steps so Stop takes effect right away"""
        end = time.time() + seconds
        while self.is_running and time.time() < end:
            time.sleep(min(0.2, end - time.time()))

    def reset_ui(self):
        self.is_running = False
        self.btn_start.config(state="normal", bg="#4CAF50")
        self.btn_reset.config(state="normal")
        self.btn_check.config(state="normal")
        self.update_button_text()
        self.load_customer_list()

    def on_close(self):
        if self.is_running and not messagebox.askyesno("Quit", "Emails are still being sent. Stop and quit?"):
            return
        self.is_running = False
        settings = self.get_batch_settings()
        self.store.set_settings(excel=self.excel_path.get().strip(), template=self.template_path.get().strip(),
                                subject=self.subject_template.get(), sender=self.sender.get(),
                                **(self.settings_to_store(settings) if settings else {}))
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
