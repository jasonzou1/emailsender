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
2. **Template** – an Outlook `.oft` file or an `.html` file. `{name}` and `{FirstName}` are replaced with the customer's first name.
3. **Subject** – `{name}` / `{FirstName}` are replaced too.
4. Click **Start Sending** (click again to **Stop**).

If `clients.xlsx`, `template.oft` or `template.html` sit next to the script/exe, they are loaded automatically.
