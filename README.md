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

4. Click **Start Sending** (click again to **Stop**).

If `clients.xlsx`, `template.oft` or `template.html` sit next to the script/exe, they are loaded automatically.
