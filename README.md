# My Baskets Email Sender

Windows desktop tool that sends personalized emails through **Classic Outlook** to a list of customers from Excel.

## Setup

```
pip install -r requirements.txt
python email_sender.py
```

## Usage

1. **Excel** – needs an `Email` column plus either:
   - a `FirstName` column, or
   - a full-name column (`Name`, `Full Name`, `Customer Name`, ...) – the first word is used, so `Anna Smith` → `Anna` (and `Smith, Anna` → `Anna`).

   Headers are matched ignoring case, spaces, `_` and `-` (e.g. `first name`, `E-mail`, `email_address` all work).
   If both exist, `FirstName` wins; an empty `FirstName` cell falls back to the full name.
2. **Template** – an Outlook `.oft` file or an `.html` file.
3. **Subject** – placeholders work here too.

### Placeholders

Use `{{ColumnName}}` for **any** column in your Excel, e.g. `{{FirstName}}`, `{{Company}}`, `{{City}}`.

- Matching ignores case, spaces, `_` and `-`: `{{company}}`, `{{ Company }}` and `{{Company_Name}}` ↔ column `Company Name` all work.
- `{{FirstName}}` / `{{name}}` always give the first name (taken from the full name if there's no FirstName column).
- Single braces (`{FirstName}`, `{Company}`) still work too.
- An empty cell becomes blank text.
- If the template or subject contains a `{{placeholder}}` with no matching column, **nothing is sent** and the log lists the unknown placeholder(s), so typos never reach customers.
- In `.oft` templates, type each placeholder in one go: if Outlook formats part of it differently (e.g. bold only `First`), the placeholder gets split and won't be replaced.

4. **Batch settings** – how many emails per batch (default 20) and seconds to wait between emails (default 5).
5. Click **Send Next 20**. It sends one batch and stops. Click it again whenever you're ready for the next batch. **Stop** halts mid-batch.

### Remembering where you are

- Every address is saved to `email_sender_progress.json` (next to the program) **right after it is sent**, so closing the app, a crash or a restart never loses your place and never re-sends to someone.
- Progress is kept per **Excel list + template** pair: a new template starts the same list from the top.
- The app reopens the last Excel/template, subject and batch settings automatically.
- Progress is tracked by email address, not row number, so you can add, remove or re-sort rows in the Excel safely. Duplicate addresses are only emailed once.
- Failed addresses are not retried automatically (shown as "Failed" in the progress line, with the reason in the log).
- **Reset Progress** forgets what was sent for the current list + template, so the next batch starts from the top again.
- Keep `email_sender_progress.json` private: it contains customer email addresses.


If `clients.xlsx`, `template.oft` or `template.html` sit next to the script/exe, they are loaded automatically.

## Troubleshooting

**"Could not connect to Outlook ... Server execution failed" (-2146959355)**

The sender and Outlook are running with different permission levels. Windows blocks them from talking to each other.

1. Close the email sender and close Outlook completely (check Task Manager for leftover `OUTLOOK.EXE`).
2. Open Outlook normally (not "Run as administrator") and wait until it has finished loading.
3. Start the email sender normally too: don't use "Run as administrator", don't launch it from an admin Command Prompt or an IDE running as admin, and make sure the `.exe`'s Properties → Compatibility → "Run this program as an administrator" is unchecked.
