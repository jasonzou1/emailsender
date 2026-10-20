import tkinter as tk
import html
import re
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import pythoncom
import win32com.client as win32
import os
import time
import threading
import sys
import ctypes
import json
from datetime import datetime

CO_E_SERVER_EXEC_FAILURE = -2146959355  # 0x80080005

# Accepted header spellings (matched case-insensitively, ignoring spaces/underscores/hyphens)
EMAIL_HEADERS = ("email", "emailaddress", "e-mail")
FIRST_NAME_HEADERS = ("firstname", "first")
FULL_NAME_HEADERS = ("name", "fullname", "customername", "clientname", "contactname")

STATE_FILENAME = "email_sender_progress.json"
DEFAULT_BATCH_SIZE = 20
DEFAULT_DELAY_SECONDS = 5

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

    @staticmethod
    def key(excel_file, template_file):
        norm = lambda p: os.path.normcase(os.path.abspath(p))
        return norm(excel_file) + "|" + norm(template_file)

    def campaign(self, excel_file, template_file):
        with self.lock:
            c = self.data["campaigns"].setdefault(self.key(excel_file, template_file), {
                "excel": excel_file, "template": template_file, "sent": [], "failed": []})
            return set(c["sent"]), set(c["failed"])

    def mark(self, excel_file, template_file, email, ok):
        with self.lock:
            c = self.data["campaigns"].setdefault(self.key(excel_file, template_file), {
                "excel": excel_file, "template": template_file, "sent": [], "failed": []})
            target = c["sent"] if ok else c["failed"]
            if email.lower() not in target:
                target.append(email.lower())
            c["updated"] = datetime.now().isoformat(timespec="seconds")
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
        self.root.geometry("620x680")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Stored variables
        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.store = ProgressStore(os.path.join(get_script_dir(), STATE_FILENAME))
        self.subject_template = tk.StringVar(value=self.store.get_setting("subject", "Hello {{FirstName}} From My Baskets"))
        self.batch_size = tk.StringVar(value=str(self.store.get_setting("batch_size", DEFAULT_BATCH_SIZE)))
        self.delay_seconds = tk.StringVar(value=str(self.store.get_setting("delay_seconds", DEFAULT_DELAY_SECONDS)))
        self.status_text = tk.StringVar(value="Select a customer list and template.")
        self.is_running = False
        self._status_job = None

        # --- UI Layout ---

        # 1. Select Excel File
        frame_excel = tk.LabelFrame(root, text="1. Select Customer List (Excel)", padx=10, pady=10) # 标签
        frame_excel.pack(fill="x", padx=10, pady=5)

        tk.Entry(frame_excel, textvariable=self.excel_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_excel, text="Browse...", command=self.select_excel).pack(side=tk.LEFT) # 按钮

        # 2. Select Outlook Template
        frame_tmpl = tk.LabelFrame(root, text="2. Select Email Template (.oft or .html)", padx=10, pady=10) # 标签
        frame_tmpl.pack(fill="x", padx=10, pady=5)

        tk.Entry(frame_tmpl, textvariable=self.template_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_tmpl, text="Browse...", command=self.select_template).pack(side=tk.LEFT) # 按钮

        # 3. Set Subject Line
        frame_subj = tk.LabelFrame(root, text="3. Set Email Subject ({{FirstName}}, {{Company}}, any column)", padx=10, pady=10) # 标签
        frame_subj.pack(fill="x", padx=10, pady=5)
        tk.Entry(frame_subj, textvariable=self.subject_template, width=60).pack(padx=5)

        # 4. Batch settings
        frame_batch = tk.LabelFrame(root, text="4. Batch Settings", padx=10, pady=10) # 标签
        frame_batch.pack(fill="x", padx=10, pady=5)
        tk.Label(frame_batch, text="Emails per batch:").pack(side=tk.LEFT) # 标签
        tk.Spinbox(frame_batch, from_=1, to=1000, textvariable=self.batch_size, width=6,
                   command=self.update_button_text).pack(side=tk.LEFT, padx=(5, 20))
        tk.Label(frame_batch, text="Seconds between emails:").pack(side=tk.LEFT) # 标签
        tk.Spinbox(frame_batch, from_=0, to=600, textvariable=self.delay_seconds, width=6).pack(side=tk.LEFT, padx=5)
        self.batch_size.trace_add("write", lambda *_: self.update_button_text())

        # 5. Progress + Action Buttons
        frame_btn = tk.Frame(root, pady=10)
        frame_btn.pack()
        tk.Label(frame_btn, textvariable=self.status_text, font=("Arial", 10, "bold")).pack(pady=(0, 8)) # 标签
        self.btn_start = tk.Button(frame_btn, text="Send Next Batch", command=self.toggle_sending, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), width=18) # 按钮
        self.btn_start.pack(side=tk.LEFT, padx=5)
        self.btn_reset = tk.Button(frame_btn, text="Reset Progress", command=self.reset_progress) # 按钮
        self.btn_reset.pack(side=tk.LEFT, padx=5)

        # 6. Execution Log
        tk.Label(root, text="Execution Log:").pack(anchor="w", padx=10) # 标签
        self.log_area = scrolledtext.ScrolledText(root, height=15)
        self.log_area.pack(fill="both", expand=True, padx=10, pady=5)

        self.auto_load_files()
        self.excel_path.trace_add("write", lambda *_: self.schedule_status_refresh())
        self.template_path.trace_add("write", lambda *_: self.schedule_status_refresh())
        self.update_button_text()
        self.refresh_status()
        if is_admin():
            self.log("Warning: running as Administrator. If Outlook is not also running as Administrator, "
                     "the connection will fail - start this program normally instead.") # 日志

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

    def get_batch_settings(self):
        """Returns (batch_size, delay_seconds) or None if the inputs aren't valid numbers"""
        try:
            batch = int(self.batch_size.get())
            delay = float(self.delay_seconds.get())
        except ValueError:
            return None
        if batch < 1 or delay < 0:
            return None
        return batch, delay

    def update_button_text(self):
        if self.is_running:
            return
        settings = self.get_batch_settings()
        self.btn_start.config(text=f"Send Next {settings[0]}" if settings else "Send Next Batch") # 按钮文本

    def schedule_status_refresh(self):
        # Wait until the user stops typing a path before re-reading the Excel file
        if self._status_job is not None:
            self.root.after_cancel(self._status_job)
        self._status_job = self.root.after(500, self.refresh_status)

    def refresh_status(self):
        """Shows how far through the current list we are (main thread only)"""
        self._status_job = None
        excel_file = self.excel_path.get().strip()
        template_file = self.template_path.get().strip()
        if not (os.path.isfile(excel_file) and os.path.isfile(template_file)):
            self.status_text.set("Select a customer list and template.")
            return
        try:
            _, recipients, _ = load_recipients(excel_file)
        except Exception as e:
            self.status_text.set(f"Can't read list: {e}")
            return
        sent, failed = self.store.campaign(excel_file, template_file)
        valid = {r['email'].lower() for r in recipients if is_valid_email(r['email'])}
        done_sent = len(valid & sent)
        done_failed = len(valid & failed - sent)
        remaining = len(valid) - done_sent - done_failed
        self.status_text.set(f"Sent {done_sent} of {len(valid)}  |  Failed {done_failed}  |  Remaining {remaining}")

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

        settings = self.get_batch_settings()
        if settings is None:
            messagebox.showwarning("Warning", "Emails per batch must be 1 or more, and seconds between emails 0 or more.") # 提示框
            return
        batch_size, delay = settings

        # Read Tk variables here, in the main thread
        subject_tmpl = self.subject_template.get()
        self.store.set_settings(excel=excel_file, template=template_file, subject=subject_tmpl,
                                batch_size=batch_size, delay_seconds=delay)

        self.is_running = True
        self.btn_start.config(text="Stop", bg="#f44336") # 按钮文本
        self.btn_reset.config(state="disabled")
        thread = threading.Thread(target=self.run_sending_process,
                                  args=(excel_file, template_file, subject_tmpl, batch_size, delay))
        thread.daemon = True
        thread.start()

    def run_sending_process(self, excel_file, template_file, subject_tmpl, batch_size, delay):
        # COM must be initialized on every thread that uses it
        pythoncom.CoInitialize()
        success_count = 0
        failed_count = 0
        try:
            self.log("-" * 30)
            self.log("Reading Excel file...") # 日志
            try:
                columns, recipients, info = load_recipients(excel_file)
            except ValueError as e:
                self.log(f"Error: {e}") # 日志
                return
            self.log(info) # 日志

            # Pick the next batch: valid addresses not already sent/failed, first occurrence only
            sent, failed = self.store.campaign(excel_file, template_file)
            done = sent | failed
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

            if not batch:
                self.log("Nothing left to send: every address in this list has already been emailed with this template.") # 日志
                self.log("Use 'Reset Progress' to send to the whole list again.") # 日志
                self.ui(messagebox.showinfo, "All Done", "Every customer in this list has already been emailed.") # 提示框
                return
            self.log(f"This batch: {len(batch)} email(s), {remaining_total} remaining in the list, "
                     f"{delay:g}s between emails.") # 日志

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
            except Exception as e:
                self.log(f"Error: Could not connect to Outlook. Ensure Classic Outlook is open. ({e})") # 日志
                if getattr(e, 'hresult', None) == CO_E_SERVER_EXEC_FAILURE or (e.args and e.args[0] == CO_E_SERVER_EXEC_FAILURE):
                    if is_admin():
                        self.log("Cause: this program is running as Administrator but Outlook is not. "
                                 "Close it and start it normally (not 'Run as administrator').") # 日志
                    else:
                        self.log("Cause: Outlook seems to be running as Administrator, or is still starting/busy. "
                                 "Close Outlook completely (check Task Manager), reopen it normally, wait until it has loaded, then try again.") # 日志
                return

            # --- Make sure every {{placeholder}} matches an Excel column before sending anything ---
            body_text = html_content if html_content is not None else (outlook.CreateItemFromTemplate(template_file).HTMLBody or "")
            known_keys = {normalize_header(c) for c in columns} | {"name", "firstname"}
            unknown = unknown_placeholders(subject_tmpl + body_text, known_keys)
            if unknown:
                self.log(f"Error: No Excel column for placeholder(s): {', '.join(unknown)}. Found columns: {', '.join(columns)}") # 日志
                return

            total = len(batch)
            for pos, r in enumerate(batch, start=1):
                if not self.is_running:
                    self.log("Stopped. Progress is saved; the next batch continues from here.") # 日志
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

                    mail.To = email
                    mail.Subject = fill_placeholders(subject_tmpl, values)
                    mail.Send()

                    self.store.mark(excel_file, template_file, email, ok=True)
                    self.log(f"[{pos}/{total}] Sent to: {name} ({email}) - Excel row {r['row']}") # 日志
                    success_count += 1

                except Exception as e:
                    self.store.mark(excel_file, template_file, email, ok=False)
                    self.log(f"[{pos}/{total}] Failed to send ({email}) - Excel row {r['row']}: {e}") # 日志
                    failed_count += 1

                self.ui(self.refresh_status)
                if pos < total:
                    self.wait(delay)

            left = remaining_total - success_count - failed_count
            summary = f"Sent: {success_count}, Failed: {failed_count}. Remaining in list: {left}"
            self.log(f"Batch finished. {summary}") # 日志
            if self.is_running:
                self.ui(messagebox.showinfo, "Batch Complete", f"Batch finished!\n{summary}") # 提示框

        except Exception as e:
            self.log(f"Critical Error: {e}") # 日志
            self.ui(messagebox.showerror, "Error", str(e)) # 提示框

        finally:
            pythoncom.CoUninitialize()
            self.ui(self.reset_ui)

    def wait(self, seconds):
        """Sleeps in small steps so Stop takes effect right away"""
        end = time.time() + seconds
        while self.is_running and time.time() < end:
            time.sleep(min(0.2, end - time.time()))

    def reset_ui(self):
        self.is_running = False
        self.btn_start.config(state="normal", bg="#4CAF50")
        self.btn_reset.config(state="normal")
        self.update_button_text()
        self.refresh_status()

    def on_close(self):
        if self.is_running and not messagebox.askyesno("Quit", "Emails are still being sent. Stop and quit?"):
            return
        self.is_running = False
        settings = self.get_batch_settings()
        self.store.set_settings(excel=self.excel_path.get().strip(), template=self.template_path.get().strip(),
                                subject=self.subject_template.get(),
                                **({"batch_size": settings[0], "delay_seconds": settings[1]} if settings else {}))
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
