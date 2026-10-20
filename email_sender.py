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

CO_E_SERVER_EXEC_FAILURE = -2146959355  # 0x80080005

# Accepted header spellings (matched case-insensitively, ignoring spaces/underscores/hyphens)
EMAIL_HEADERS = ("email", "emailaddress", "e-mail")
FIRST_NAME_HEADERS = ("firstname", "first")
FULL_NAME_HEADERS = ("name", "fullname", "customername", "clientname", "contactname")

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


class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("My Baskets Email Sender (HTML/OFT Supported)") # 软件标题
        self.root.geometry("600x550")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Stored variables
        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.subject_template = tk.StringVar(value="Hello {name} From My Baskets")
        self.is_running = False

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

        # 4. Action Button (Start / Stop)
        frame_btn = tk.Frame(root, pady=10)
        frame_btn.pack()
        self.btn_start = tk.Button(frame_btn, text="Start Sending", command=self.toggle_sending, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), width=15) # 按钮
        self.btn_start.pack()

        # 5. Execution Log
        tk.Label(root, text="Execution Log:").pack(anchor="w", padx=10) # 标签
        self.log_area = scrolledtext.ScrolledText(root, height=15)
        self.log_area.pack(fill="both", expand=True, padx=10, pady=5)

        self.auto_load_files()
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
        """Auto-detects files in the current directory"""
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

        # Read Tk variables here, in the main thread
        subject_tmpl = self.subject_template.get()

        self.is_running = True
        self.btn_start.config(text="Stop", bg="#f44336") # 按钮文本
        thread = threading.Thread(target=self.run_sending_process,
                                  args=(excel_file, template_file, subject_tmpl))
        thread.daemon = True
        thread.start()

    def run_sending_process(self, excel_file, template_file, subject_tmpl):
        # COM must be initialized on every thread that uses it
        pythoncom.CoInitialize()
        success_count = 0
        skipped_count = 0
        failed_count = 0
        try:
            self.log("-" * 30)
            self.log("Reading Excel file...") # 日志
            df = pd.read_excel(excel_file)
            df.columns = [str(c).strip() for c in df.columns]

            email_col = find_column(df.columns, EMAIL_HEADERS)
            first_col = find_column(df.columns, FIRST_NAME_HEADERS)
            full_col = find_column(df.columns, FULL_NAME_HEADERS)

            if email_col is None or (first_col is None and full_col is None):
                self.log("Error: Excel needs an 'Email' column and a 'FirstName' (or full 'Name') column. "
                         f"Found: {', '.join(df.columns)}") # 日志
                return
            if first_col is not None:
                self.log(f"Using columns: '{first_col}' (first name), '{email_col}' (email)") # 日志
            else:
                self.log(f"Using columns: '{full_col}' (full name, first word used), '{email_col}' (email)") # 日志

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
            known_keys = {normalize_header(c) for c in df.columns} | {"name", "firstname"}
            unknown = unknown_placeholders(subject_tmpl + body_text, known_keys)
            if unknown:
                self.log(f"Error: No Excel column for placeholder(s): {', '.join(unknown)}. Found columns: {', '.join(df.columns)}") # 日志
                return

            total = len(df)
            for pos, (_, row) in enumerate(df.iterrows(), start=1):
                if not self.is_running:
                    self.log("Task interrupted.") # 日志
                    break

                email = clean_cell(row[email_col])
                name = clean_cell(row[first_col]) if first_col is not None else ""
                if not name and full_col is not None:
                    name = first_name_from_full(clean_cell(row[full_col]))
                values = build_values(row, df.columns, name)

                if not email or '@' not in email:
                    self.log(f"[{pos}/{total}] Skipped: invalid or empty email '{email}'") # 日志
                    skipped_count += 1
                    continue

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

                    self.log(f"[{pos}/{total}] Successfully sent to: {name} ({email})") # 日志
                    success_count += 1

                    time.sleep(1)

                except Exception as e:
                    self.log(f"[{pos}/{total}] Failed to send ({email}): {e}") # 日志
                    failed_count += 1

            summary = f"Sent: {success_count}, Failed: {failed_count}, Skipped: {skipped_count}"
            self.log(f"Task finished. {summary}") # 日志
            self.ui(messagebox.showinfo, "Complete", f"Task finished!\n{summary}") # 提示框

        except Exception as e:
            self.log(f"Critical Error: {e}") # 日志
            self.ui(messagebox.showerror, "Error", str(e)) # 提示框

        finally:
            pythoncom.CoUninitialize()
            self.ui(self.reset_ui)

    def reset_ui(self):
        self.is_running = False
        self.btn_start.config(state="normal", text="Start Sending", bg="#4CAF50") # 按钮文本

    def on_close(self):
        if self.is_running and not messagebox.askyesno("Quit", "Emails are still being sent. Stop and quit?"):
            return
        self.is_running = False
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
