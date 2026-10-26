# My Baskets Email Sender

Windows desktop tool that sends personalized emails through **Classic Outlook** to a list of customers from Excel.

## Setup

```
pip install -r requirements.txt
python email_sender.py
```

## Usage

0. **Outlook Connection** – the app checks Outlook when it opens (or click **Check Connection**). The status turns green and **Send from** lists every email account in Outlook; pick the one to send from. Your choice is remembered, and you're asked to confirm the sending address before every batch.

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
- If a customer's cell is **empty** for any placeholder the template or subject uses, that customer is **skipped** and shown as **Missing data** (orange) in the Customers tab, naming the empty field – so nobody gets "you sent  to your team". Fill in the cell and click **Load / Refresh List**.
- Numbers show as typed (`25`, not `25.0`). Excel dates: a date on the 1st shows as a month (`November 2025`), any other date as `December 5, 2025`. For just a month name ("November"), type it as text.
- If the template or subject contains a `{{placeholder}}` with no matching column, **nothing is sent** and the log lists the unknown placeholder(s), so typos never reach customers.
- In `.oft` templates, type each placeholder in one go: if Outlook formats part of it differently (e.g. bold only `First`), the placeholder gets split and won't be replaced.

4. **Batch settings**
   - **Emails per batch** (default 20).
   - **Random wait between emails**: e.g. `30` to `60` seconds – each email waits a random time in that range, so sending doesn't look robotic. Set both to the same number for a fixed wait.
   - **Minutes between batches** (default 30).
   - **Automatically send the next batch after the gap**:
     - **Off** (default): click **Send Next 20** for each batch. If you click before the gap is over, the app warns you and shows how long is left (you can still send anyway). The time left is also shown under the progress line, and it's remembered even after closing the app.
     - **On**: click **Start Auto** once; the app sends a batch, waits the gap (with a countdown showing when the next batch starts), sends the next, and so on until the whole list is done. Keep the app and Outlook open. **Stop** works at any time, including during the wait.
5. Click **Send Next 20** / **Start Auto**. **Stop** halts immediately; progress is saved.

### Customers tab

Shows every row of the Excel list with its status:

| Status | Meaning |
|---|---|
| **Sent** (green) | Already emailed – the time it was sent is shown |
| **Failed** (red) | Outlook refused it – the error is shown; not retried automatically |
| **Next batch** (blue) | Will be sent when you click **Send Next** |
| **Pending** | Not sent yet, in a later batch |
| **Invalid email** / **Duplicate** (gray) | Skipped |
| **Missing data** (orange) | Skipped – a cell used by the template is empty |

Use **Show** to filter (All / Not sent yet / Sent / Failed / Skipped) and **Load / Refresh List** after editing the Excel file. The table updates live while a batch is sending. The full step-by-step log is in the **Execution Log** tab.

### Remembering where you are

- Every address is saved to `email_sender_progress.json` (next to the program) **right after it is sent**, so closing the app, a crash or a restart never loses your place and never re-sends to someone.
- Progress is kept per **Excel list + template** pair: a new template starts the same list from the top.
- The app reopens the last Excel/template, subject and batch settings automatically.
- Progress is tracked by email address, not row number, so you can add, remove or re-sort rows in the Excel safely. Duplicate addresses are only emailed once.
- Failed addresses are not retried automatically (shown as "Failed" with the reason). Click **Retry Failed** to move them all back to "not sent yet".
- If Outlook is closed, restarts or the PC sleeps mid-run, the app reconnects automatically (4 tries, 15 s apart) and carries on. If Outlook doesn't come back it stops and tells you; customers it couldn't reach stay "not sent yet" – they are never marked Failed because of an Outlook connection problem.
- **Reset Progress** forgets what was sent for the current list + template, so the next batch starts from the top again.
- Keep `email_sender_progress.json` private: it contains customer email addresses.


If `clients.xlsx`, `template.oft` or `template.html` sit next to the script/exe, they are loaded automatically.

## Troubleshooting

**"Could not connect to Outlook ... Server execution failed" (-2146959355)**

The sender and Outlook are running with different permission levels. Windows blocks them from talking to each other.

1. Close the email sender and close Outlook completely (check Task Manager for leftover `OUTLOOK.EXE`).
2. Open Outlook normally (not "Run as administrator") and wait until it has finished loading.
3. Start the email sender normally too: don't use "Run as administrator", don't launch it from an admin Command Prompt or an IDE running as admin, and make sure the `.exe`'s Properties → Compatibility → "Run this program as an administrator" is unchecked.

**"The RPC server is unavailable" (-2147023174)**

Outlook went away mid-run (closed, restarted by an Office update, crashed, or the PC went to sleep). The app now reconnects automatically. To keep it from happening: keep Outlook open, and set Windows *Settings → System → Power → Sleep* to "Never" while a long auto run is going.

## Templates

`templates/holiday-reorder-2026.html` – 2026 holiday re-order email. Starter Excel with the right headers: `templates/holiday-reorder-2026-clients.xlsx`. Excel columns needed: `Email`, `Name` (or `FirstName`), `Order Month`, `Quantity`, `Product`, `List`.

`templates/winback-2026.html` – 2026 win-back email with Early Bird pricing and CASL footer. Excel columns needed: `Email`, `Name` (first name, or "there"), `Sentence` (one full sentence about their last order, e.g. "In December 2024 you sent 187 gift baskets to your team and clients, including the Medium Charcuterie Xmas Gift."). Use a subject without `{{Name}}`, since some rows use "there".
