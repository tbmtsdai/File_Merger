# Daily Auto-Merge — Operations Guide

`daily_auto_merge.py` runs unattended each morning on the office laptop,
picks up the day's ZOHO / ERP Pending Calls email from Outlook, saves the
attachment, and unions it into a rolling `merged_output_auto.xlsx` in the
`Raw Files` folder.

It is **local only** — nothing about this touches Streamlit Cloud, the
`master` branch, or `merged_output.xlsx` (your manual workflow file).

---

## What gets installed

| File                          | Purpose                                                        |
| ----------------------------- | -------------------------------------------------------------- |
| `merge_core.py`               | Pure-pandas merge / clean / dedup helpers (shared with the app) |
| `daily_auto_merge.py`         | The automation script                                          |
| `install_scheduled_task.ps1`  | One-time Task Scheduler registration                           |
| `requirements-auto.txt`       | Extra deps (`pywin32`, `winotify`)                             |

Also created on first run, under `Raw Files\`:

| File                               | Purpose                                                                       |
| ---------------------------------- | ----------------------------------------------------------------------------- |
| `merged_output_auto.xlsx`          | The rolling merged output (this script's own file)                            |
| `column_map.json`                  | Learned canonical column mapping (`{"source name": "canonical name", ...}`)   |
| `column_audit.jsonl`               | Datewise audit trail — one JSON line per `seeded` / `merged` / `pause` event  |
| `auto_merge.log`                   | Rolling log (5 MB rotation, 3 backups kept)                                   |
| `.auto_merge_state.json`           | Idempotency marker (which day was last processed)                             |

---

## One-time install (run once)

Open **PowerShell as your normal user** (not Administrator) and run:

```powershell
cd "C:\Users\k.buch\Documents\file merger and join app"

# 1. Install the automation-only Python deps into your Anaconda base env.
C:\Users\k.buch\AppData\Local\anaconda3\python.exe -m pip install -r requirements-auto.txt

# 2. Register the scheduled task (hourly Mon-Sat 11:00-15:00).
powershell -ExecutionPolicy Bypass -File .\install_scheduled_task.ps1

# 3. Optional: run once now to prime today's merge.
Start-ScheduledTask -TaskName ZOHO_Daily_AutoMerge
```

Outlook desktop must be signed in and running while the task fires;
otherwise the script logs "Outlook COM failed" and retries next hour.

---

## How it decides what to process

Each run, in order:

1. **Outlook Inbox scan** — restrict to the last **30 days** of Inbox,
   then keep messages where:
   - Subject contains **"Pending Call"** (case-insensitive) — matches both
     the ZOHO and ERP subject variants.
   - Attachment count is **exactly 1** — filters out the evening mail that
     has 2 attachments (Excel + PPT).
2. **File-date extraction** — parse the `DD-Mon-YYYY` date embedded in
   the **attachment's filename** (e.g. `ERP Pending Calls as on
   05-Sep-2026.xls` → 2026-09-05). Handles all the sender's variants
   (short/full month names, 2-/4-digit years, spaces or hyphens).
3. **Floor guard** — skip any mail whose file-date is **≤ the last
   processed file-date** (`state.last_processed_date`). This does three
   things at once:
   - Dedupes same-file resends: if the sender sends the 05-Sep file
     again on Sep 7, the resend is skipped.
   - Ignores pre-install backlog: mails older than the last date you
     manually merged are treated as "already handled by hand".
   - Skips accidental duplicates within a hourly cadence.
4. **Earliest per file-date** — if the sender resent the same-dated file
   in multiple mails, the **earliest-received** mail in that group wins.
5. **Oldest file-date first** — across file-dates, the **oldest**
   unprocessed date is processed first. This makes catch-up automatic:
   if you miss Monday and log in Tuesday, Tuesday's 11:00 run picks
   Monday's mail, 12:00 picks Tuesday's, etc.
6. **Attachment saved as-is** under the sender's original filename in
   `Raw Files\` (e.g. `ERP Pending Calls as on 07-Sep-2026.xls`).
   No renaming, no format conversion — the folder view stays identical
   to your manual practice.
7. **Column mapping applied** from `Raw Files\column_map.json`.
8. **Union onto `merged_output_auto.xlsx`** using the exact same
   Union All + clean-dtypes + XlsxWriter code path the app uses.
9. **Windows toast** + **rolling log line** + **`column_audit.jsonl`**
   entry (`merged` event with `file_date`).
10. **`last_processed_date` advanced** to this file-date — the floor
    moves forward monotonically.
11. **Processed mail moved** from Inbox to `Inbox\TSD\CC` (best-effort;
    a failure here writes `move_failed` to the audit trail but does not
    undo the merge).

`Raw Files\merged_output.xlsx` is never opened or modified.
Mail is moved only after the merge itself succeeded; on a paused day
(unknown columns) the mail stays in Inbox for the next hourly retry.

### What happens if you miss a day

Missed Monday, back at your desk Tuesday morning:

1. **Task Scheduler:** `-StartWhenAvailable` fires the missed Monday
   trigger once when you log in, alongside Tuesday's normal 11:00 trigger.
2. **Script logic:** the Inbox scan finds both Monday's and Tuesday's
   mails, both with file-dates > last_processed_date. It picks the
   **oldest** (Monday), merges it, moves it to CC. The next hourly run
   picks Tuesday. Backlog clears one mail per hourly tick.

This works for any gap ≤ 30 days (the LOOKBACK_DAYS scan window).

---

## The column-mapping learning loop

The script **never** guesses column names. `merged_output.xlsx`'s
columns are the canonical set: every incoming column must either match
one of them exactly (whitespace / hyphen / case included — strict) or
be listed as a source key in `column_map.json`.

If the day's file has **any** unknown column, the script:

- Logs the unknown column names.
- Fires a toast: *"Auto-merge paused — new columns detected. Open the app
  once to teach the mapping."*
- Writes a `pause` line to `column_audit.jsonl` naming the offending
  columns so you can review later.
- Does **not** mark the day as processed — after you teach the mapping,
  the next hourly run picks it up cleanly.

**To teach a mapping:** open the app (`run_app.bat` → localhost:8501),
load today's file plus at least one older file that already merges, and
use the Column Alignment step to assign canonical names. Submitting the
form writes the learned mapping to `column_map.json` automatically.

---

## Audit trail (`column_audit.jsonl`)

One append-only JSON line per event, so you can see who paused when.
Open it in any text editor, or import into Excel as JSON.

Events:

| event          | Fields (in addition to `timestamp`, `date`, `event`)                                       |
| -------------- | ------------------------------------------------------------------------------------------ |
| `seeded`       | `source_file` (= `merged_output.xlsx`), `canonical_column_count`                           |
| `merged`       | `source_file`, `rows_added`, `total_rows`, `output_file`                                   |
| `pause`        | `source_file`, `unknown_columns` (array), `known_column_count`                             |
| `moved`        | `source_file`, `destination` (e.g. `Inbox\TSD\CC`)                                         |
| `move_failed`  | `source_file`, `destination`, `error` (Outlook COM exception summary — merge still stood)  |

Example:

```json
{"timestamp": "2026-09-07T08:03:12", "date": "2026-09-07", "event": "pause", "source_file": "ZOHO_Pending_Calls_2026-09-07.xlsx", "unknown_columns": ["Contract End Dt"], "known_column_count": 37}
{"timestamp": "2026-09-07T10:14:55", "date": "2026-09-07", "event": "merged", "source_file": "ZOHO_Pending_Calls_2026-09-07.xlsx", "rows_added": 1104, "total_rows": 116558, "output_file": "merged_output_auto.xlsx"}
```

Never truncated — keep the full history.

---

## Checking on it

```powershell
# See the last few log entries
Get-Content "C:\Users\k.buch\OneDrive - Transasia Bio Medicals Ltd\TBM 2026 Onwards\Pending Calls\Raw Files\auto_merge.log" -Tail 30

# Scheduled-task status
Get-ScheduledTaskInfo -TaskName ZOHO_Daily_AutoMerge

# Fire once by hand (idempotent — no-op if today already done)
Start-ScheduledTask -TaskName ZOHO_Daily_AutoMerge
```

Successful run's log line looks like:

```
2026-09-05 08:03:12  INFO     OK  date=2026-09-05  source=ZOHO_Pending_Calls_2026-09-05.xlsx  rows_added=1247  total_rows=53219  output=merged_output_auto.xlsx
```

---

## Pausing / removing

```powershell
# Pause (task stays registered but won't fire)
Disable-ScheduledTask -TaskName ZOHO_Daily_AutoMerge

# Re-enable
Enable-ScheduledTask -TaskName ZOHO_Daily_AutoMerge

# Remove entirely
Unregister-ScheduledTask -TaskName ZOHO_Daily_AutoMerge -Confirm:$false
```

You can also open **Task Scheduler** (Win+R → `taskschd.msc`) and find
`ZOHO_Daily_AutoMerge` under the top-level Task Scheduler Library.

---

## Troubleshooting

| Symptom                                          | What to check                                                                     |
| ------------------------------------------------ | --------------------------------------------------------------------------------- |
| Log says *"No qualifying mail in today's Inbox"* | Mail hasn't arrived yet. Wait for the next hourly retry.                          |
| Log says *"Outlook COM failed"*                  | Outlook desktop isn't open. Open Outlook, then `Start-ScheduledTask …`.           |
| Toast: *"new columns detected"*                  | Open the app, teach the mapping (see above), then re-run the task.                |
| Log says *"Already processed 2026-…-…"*          | Normal on a re-run. Delete `.auto_merge_state.json` if you need to force a re-run.|
| `pip install` complains about pywin32            | Run `python -m pip install --upgrade pip` first, then retry.                      |

The script is written to **never touch `merged_output.xlsx`**, **never
scan the whole folder**, and **never guess** column names — if any of
those assumptions get violated, prefer opening an issue over patching.
