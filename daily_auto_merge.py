"""
daily_auto_merge.py — Windows-scheduled unattended daily ZOHO / ERP Pending
Calls auto-merge.

Runs hourly Mon-Sat 11:00-15:00 (see install_scheduled_task.ps1;
the 08:00 and 10:00 slots are used by other project automations).
Each run:

    1. Outlook COM scan    — walk the last LOOKBACK_DAYS days of Inbox, keep
                             mails whose subject contains "Pending Call".
                             (Attachment count is NOT used — evening mails
                             may have 1 or 2 attachments; we identify by
                             subject + parseable Excel file-date instead.)
    2. File-date dedupe    — for each mail pick the first Excel attachment
                             whose filename yields a parseable "DD-Mon-YYYY"
                             date; skip mails whose file-date is <=
                             state.last_processed_date (already processed
                             or covered by the user's manual workflow
                             before install).
    3. Earliest-per-date   — if the sender resent the same-dated file
                             across multiple mails, keep the earliest one.
                             Across dates, pick the OLDEST file-date first,
                             so hourly runs naturally drain backlog after
                             a missed day.
    4. Save attachment     — Raw Files\\<sender's original filename>. No
                             renaming, no conversion. Matches the user's
                             existing manual practice.
    5. Apply column_map    — abort with a toast if any column is unknown;
                             user teaches it in the Streamlit app and the
                             next hourly run succeeds.
    6. Union onto rolling  — first-run seed = merged_output.xlsx; subsequent
       merged_output_auto     runs = previous merged_output_auto + today's
                             file only.
    7. XlsxWriter output   — mirrors the app's OOM-safe writer.
    8. Notify + log        — winotify toast, rolling log at Raw Files\\auto_merge.log,
                             structured audit trail at column_audit.jsonl.
    9. Sweep-move to CC    — after merge success, SWEEP every Inbox
                             Pending-Calls mail whose Excel file-date is
                             <= the new floor into Inbox\\TSD\\CC. Catches
                             morning + evening (1 or 2 attachments) +
                             resends + stragglers. Best-effort; failure
                             is audited but never fails the run.

    Never touches Raw Files\\merged_output.xlsx (user's manual workflow).
    Never creates any file the user wasn't already saving by hand.
    Never pushes anywhere. Never scans the whole Raw Files folder.
    Never moves mail unless the merge succeeded.
    Idempotent: last_processed_date is monotonic; a later resend of an
    older file-date is naturally skipped by the floor check.

Runs entirely on-device: no network. Requires Outlook desktop client to be
open (it uses the same COM session the user's Outlook session provides).
"""

from __future__ import annotations

import json
import logging
import os
import re
import sys
import traceback
from datetime import date, datetime, time as dtime, timedelta
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

# merge_core.py sits next to this script — reuse the exact same union / clean
# / dedup / xlsxwriter code path the Streamlit app uses.
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from merge_core import (  # noqa: E402
    TRACKING_COLS,
    _dedup_columns,
    apply_column_map,
    clean_dtypes,
    do_union_all,
    load_column_map,
    read_from_path,
    to_excel_file,
    unknown_columns,
)


# ─── Config ──────────────────────────────────────────────────────────────────
RAW_FILES_DIR = Path(
    r"C:\Users\k.buch\OneDrive - Transasia Bio Medicals Ltd"
    r"\TBM 2026 Onwards\Pending Calls\Raw Files"
)
LOG_PATH             = RAW_FILES_DIR / "auto_merge.log"
STATE_PATH           = RAW_FILES_DIR / ".auto_merge_state.json"
AUDIT_PATH           = RAW_FILES_DIR / "column_audit.jsonl"          # datewise audit
COLUMN_MAP_PATH      = RAW_FILES_DIR / "column_map.json"
MERGED_MANUAL_PATH   = RAW_FILES_DIR / "merged_output.xlsx"          # DO NOT TOUCH
MERGED_AUTO_PATH     = RAW_FILES_DIR / "merged_output_auto.xlsx"     # ours

SUBJECT_MATCH        = "pending call"   # case-insensitive substring
LOG_MAX_BYTES        = 5 * 1024 * 1024  # 5 MB rotation
LOG_BACKUP_COUNT     = 3

APP_ID               = "ZOHO Auto-Merge"

# After a successful merge, move the processed mail to
# Inbox\<CC_FOLDER_PATH[0]>\<CC_FOLDER_PATH[1]>\... — matches the user's
# manual "file it under TSD\CC" habit and keeps the Inbox clean.
CC_FOLDER_PATH       = ("TSD", "CC")

# Outlook scan window: how many days back from today we search the Inbox for
# unprocessed mails. Covers backlog when the user misses one or more days.
LOOKBACK_DAYS        = 30

# Mapping of month names (short + full, lowercased) to month numbers, used
# by extract_file_date. Handles the sender's variety: "Sep", "September",
# "SEPT", "MAY", "June", etc.
_MONTH_MAP: Dict[str, int] = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7,
    "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
    "january": 1, "february": 2, "march": 3, "april": 4, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11,
    "december": 12,
}

# Matches "05-Sep-2026", "13-MAY-26", "25-Apr 2026", "02-June-2026" —
# day + separator + month name + separator + 2- or 4-digit year.
_FILE_DATE_RE = re.compile(
    r"(\d{1,2})\s*[-\s]\s*([A-Za-z]{3,})\s*[-\s]\s*(\d{2,4})",
    re.IGNORECASE,
)


def extract_file_date(filename: str) -> Optional[date]:
    """Parse the day-of-report from a Pending-Calls attachment filename.

    Returns None if no valid date is found. Two-digit years are treated
    as 20xx (so "26" → 2026); invalid combinations (e.g. Feb 30) return
    None rather than raising.
    """
    m = _FILE_DATE_RE.search(filename)
    if not m:
        return None
    try:
        day = int(m.group(1))
    except ValueError:
        return None
    month = _MONTH_MAP.get(m.group(2).lower())
    if month is None:
        return None
    try:
        year = int(m.group(3))
    except ValueError:
        return None
    if year < 100:
        year += 2000
    try:
        return date(year, month, day)
    except ValueError:
        return None


# ─── Logging + toast ─────────────────────────────────────────────────────────
def _init_logger() -> logging.Logger:
    RAW_FILES_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("zoho_auto_merge")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        h = RotatingFileHandler(
            LOG_PATH, maxBytes=LOG_MAX_BYTES,
            backupCount=LOG_BACKUP_COUNT, encoding="utf-8")
        h.setFormatter(logging.Formatter(
            "%(asctime)s  %(levelname)-7s  %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"))
        logger.addHandler(h)
    return logger


log = _init_logger()


def _toast(title: str, body: str) -> None:
    """Best-effort Windows toast via winotify. Absence is not an error."""
    try:
        from winotify import Notification  # local import: dep is optional-ish
        Notification(app_id=APP_ID, title=title, msg=body).show()
    except Exception as e:  # noqa: BLE001
        log.warning("toast failed (%s): %s", type(e).__name__, e)


def _audit(event: str, **fields) -> None:
    """Append one JSON line to Raw Files\\column_audit.jsonl.

    Each line records: timestamp, date, event (pause | merged | seeded),
    source_file, and event-specific fields (unknown_columns, rows_added, ...).
    Never raises — audit failure must not break the pipeline.
    """
    try:
        record = {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "date": date.today().isoformat(),
            "event": event,
            **fields,
        }
        AUDIT_PATH.parent.mkdir(parents=True, exist_ok=True)
        with AUDIT_PATH.open("a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception as e:  # noqa: BLE001
        log.warning("audit write failed: %s", e)


# ─── State (idempotency) ─────────────────────────────────────────────────────
def _read_state() -> Dict[str, str]:
    if not STATE_PATH.exists():
        return {}
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _write_state(state: Dict[str, str]) -> None:
    STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")


def _today_iso() -> str:
    return date.today().isoformat()


# ─── Outlook COM helpers ─────────────────────────────────────────────────────
def _pick_excel_with_date(msg) -> Optional[Tuple[str, date]]:
    """Return (attachment_name, file_date) for the FIRST Excel attachment
    on msg whose filename yields a parseable date. None if there isn't one.

    Iterating (rather than assuming attachment #1) is what lets us handle
    evening mails that pair the Excel with a PPT: the Excel might sit at
    any 1-indexed position.
    """
    try:
        atts = msg.Attachments
    except Exception:  # noqa: BLE001
        return None
    if atts.Count < 1:
        return None
    for i in range(1, atts.Count + 1):
        try:
            att  = atts.Item(i)
            name = str(getattr(att, "FileName", "") or "")
        except Exception:  # noqa: BLE001
            continue
        if not name.lower().endswith((".xls", ".xlsx", ".xlsm")):
            continue
        fd = extract_file_date(name)
        if fd is None:
            continue
        return name, fd
    return None


# ─── Outlook COM: find the next unprocessed mail ─────────────────────────────
def _find_next_mail(
    after: Optional[date],
) -> Optional[Tuple[object, str, date]]:
    """Find the OLDEST-file-date unprocessed Pending-Calls mail in Inbox.

    A candidate is a mail whose subject contains SUBJECT_MATCH and that
    carries exactly ONE Excel attachment whose filename yields a parseable
    file-date (via extract_file_date). Mails whose file-date is <= `after`
    are skipped — this is how we dedupe resends of the same-dated file
    (already processed) and how we ignore old Inbox mail from before the
    automation started (`after` = last_processed_date on install).

    If the sender resent the same-dated file across several mails, the
    EARLIEST-received mail in that group wins. Across groups, the OLDEST
    file-date wins. This makes catch-up after missed days automatic — the
    hourly runs drain the backlog one file-date at a time.

    Returns (MailItem, attachment_filename, file_date) or None.
    """
    import pythoncom            # noqa: F401  (needed to init COM in Task Scheduler)
    import win32com.client as win32

    outlook   = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox     = namespace.GetDefaultFolder(6)   # 6 = olFolderInbox

    items = inbox.Items
    items.Sort("[ReceivedTime]", False)         # oldest first

    # We do NOT use items.Restrict — its date-string parsing is locale-
    # sensitive (US format on an en-IN client can silently return zero rows
    # without raising). At ~1000s of items this iteration is trivial.
    cutoff = date.today() - timedelta(days=LOOKBACK_DAYS)

    # For each file-date encountered, keep the EARLIEST-received message.
    # Identification is subject-based (SUBJECT_MATCH), NOT attachment-count
    # based — evening mails can carry 1 OR 2 attachments, so we pick the
    # first Excel attachment whose filename parses to a file-date.
    by_date: Dict[date, Tuple[datetime, object, str]] = {}
    for msg in items:
        try:
            received = getattr(msg, "ReceivedTime", None)
            if received is None:
                continue
            try:
                if received.date() < cutoff:
                    continue                  # older than LOOKBACK_DAYS
            except Exception:  # noqa: BLE001
                pass
            subj = str(getattr(msg, "Subject", "") or "")
            if SUBJECT_MATCH not in subj.lower():
                continue
            picked = _pick_excel_with_date(msg)
            if picked is None:
                continue
            att_name, file_dt = picked
            if after is not None and file_dt <= after:
                continue        # dedupe resends + skip pre-install backlog
            existing = by_date.get(file_dt)
            if existing is None or received < existing[0]:
                by_date[file_dt] = (received, msg, att_name)
        except Exception as e:  # noqa: BLE001
            log.warning("skipping a mail (%s): %s", type(e).__name__, e)
            continue

    if not by_date:
        return None
    oldest_file_dt = min(by_date.keys())
    _received, msg, att_name = by_date[oldest_file_dt]
    return msg, att_name, oldest_file_dt


def _try_sweep_processed_to_cc(floor: date) -> None:
    """Sweep every Inbox Pending-Calls mail whose Excel-attachment file-date
    is <= floor into Inbox\\TSD\\CC.

    Catches:
      - The morning mail we just processed.
      - The evening mail for the same file-date (with 1 OR 2 attachments —
        we identify it by subject + parseable Excel file-date, not count).
      - Same-date resends that arrived later.
      - Stragglers from earlier days that never got moved (e.g., a Day-1
        evening mail that landed after the Day-1 fire had already run).

    Best-effort: any COM failure here is logged + audited but never breaks
    the pipeline. The merge has already committed by the time we get here.
    """
    import pythoncom            # noqa: F401  (COM init under Task Scheduler)
    import win32com.client as win32

    dest_display = "Inbox\\" + "\\".join(CC_FOLDER_PATH)
    try:
        outlook   = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox     = namespace.GetDefaultFolder(6)   # 6 = olFolderInbox
        target    = inbox
        for name in CC_FOLDER_PATH:
            target = target.Folders[name]
    except Exception as e:  # noqa: BLE001
        log.warning("Sweep: cannot open %s (%s: %s).",
                    dest_display, type(e).__name__, e)
        _audit("move_failed", destination=dest_display,
               error=f"open target: {type(e).__name__}: {e}")
        return

    items = inbox.Items
    items.Sort("[ReceivedTime]", False)
    cutoff = date.today() - timedelta(days=LOOKBACK_DAYS)

    # Collect first — moving items while enumerating the same collection
    # skips entries (COM iterator invalidation). Iterate everything and
    # filter in Python; items.Restrict on [ReceivedTime] is locale-flaky.
    candidates: List[Tuple[object, str, date]] = []
    for msg in items:
        try:
            received = getattr(msg, "ReceivedTime", None)
            if received is not None:
                try:
                    if received.date() < cutoff:
                        continue
                except Exception:  # noqa: BLE001
                    pass
            subj = str(getattr(msg, "Subject", "") or "")
            if SUBJECT_MATCH not in subj.lower():
                continue
            picked = _pick_excel_with_date(msg)
            if picked is None:
                continue
            att_name, fd = picked
            if fd > floor:
                continue        # not yet processed — leave in Inbox
            candidates.append((msg, att_name, fd))
        except Exception as e:  # noqa: BLE001
            log.warning("Sweep: skipping mail (%s: %s)",
                        type(e).__name__, e)

    moved = 0
    for msg, att_name, fd in candidates:
        try:
            msg.Move(target)
            moved += 1
            _audit("moved",
                   source_file=att_name,
                   file_date=fd.isoformat(),
                   destination=dest_display)
        except Exception as e:  # noqa: BLE001
            log.warning("Sweep: move failed for '%s' (%s: %s).",
                        att_name, type(e).__name__, e)
            _audit("move_failed",
                   source_file=att_name,
                   file_date=fd.isoformat(),
                   destination=dest_display,
                   error=f"{type(e).__name__}: {e}")
    log.info("Sweep moved %d mail(s) with file-date <= %s to %s.",
             moved, floor.isoformat(), dest_display)


def _save_attachment(msg, att_name: str, dest_path: Path) -> None:
    """Save the attachment named att_name from msg to dest_path AS-IS.

    Iterates msg.Attachments to find the right one by filename — do NOT
    assume it's at index 1, because the evening mail may pair a PPT with
    the Excel in either order. No renaming, no format conversion. Deletes
    the destination first if it already exists (SaveAsFile refuses to
    overwrite).
    """
    try:
        if dest_path.exists():
            dest_path.unlink()
    except OSError:
        pass
    atts = msg.Attachments
    for i in range(1, atts.Count + 1):
        att = atts.Item(i)
        if str(getattr(att, "FileName", "") or "") == att_name:
            att.SaveAsFile(str(dest_path))
            return
    raise ValueError(f"Attachment named {att_name!r} not found on message.")


# ─── Data pipeline ───────────────────────────────────────────────────────────
def _read_all_sheets_stacked(path: Path, source_label: str) -> pd.DataFrame:
    """Read every sheet of an Excel file and stack them into one DataFrame,
    then dedup column names + set the Source File column (preserved if
    already present, matching the app's invariant)."""
    sheets = read_from_path(str(path))
    frames: List[pd.DataFrame] = []
    for _sname, df in sheets.items():
        d = _dedup_columns(df.copy())
        if "Source File" not in d.columns:
            d.insert(0, "Source File", source_label)
        frames.append(d)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, join="outer")


def _load_prev_merged() -> pd.DataFrame:
    """Load the previous rolling merge into a single DataFrame.
    On the first run this returns an empty frame and the seed step handles
    priming from merged_output.xlsx."""
    if not MERGED_AUTO_PATH.exists():
        return pd.DataFrame()
    return _read_all_sheets_stacked(MERGED_AUTO_PATH, source_label="merged_output_auto.xlsx")


def _seed_from_manual() -> pd.DataFrame:
    """First-run seed: the user's existing manual merge, as-is."""
    if not MERGED_MANUAL_PATH.exists():
        log.warning("Manual merged_output.xlsx not found — seeding empty.")
        return pd.DataFrame()
    return _read_all_sheets_stacked(MERGED_MANUAL_PATH, source_label="merged_output.xlsx")


def _union_and_write(prev: pd.DataFrame,
                     today_df: pd.DataFrame,
                     today_source: str) -> Tuple[int, int]:
    """Union prev + today_df (Union All), clean dtypes, write to
    MERGED_AUTO_PATH via XlsxWriter. Returns (rows_added, total_rows)."""
    # Enforce Source File on today's frame (preserved if the file already has it).
    if "Source File" not in today_df.columns:
        today_df = today_df.copy()
        today_df.insert(0, "Source File", today_source)

    # Union All — matches app default. Empty prev is fine.
    parts = [d for d in (prev, today_df) if d is not None and not d.empty]
    if not parts:
        combined = pd.DataFrame()
    else:
        combined, _audit = do_union_all(parts)

    # Auto-clean data types ON, same as the app default.
    combined, _report = clean_dtypes(combined)

    total  = len(combined)
    added  = total - len(prev)

    # Write single-sheet multi-sheet-shaped output (one sheet = "Merged").
    to_excel_file({"Merged": combined}, str(MERGED_AUTO_PATH))
    return added, total


# ─── Main ────────────────────────────────────────────────────────────────────
def _load_floor_date(state: Dict[str, object]) -> Optional[date]:
    """Return the max file-date we've already processed (the 'floor').

    Any incoming file-date <= floor is treated as already-handled — this
    dedupes same-day resends and, on install, treats every mail older than
    the last manually-merged date as already-covered by the user's manual
    workflow.
    """
    lp = state.get("last_processed_date")
    if not lp:
        return None
    try:
        return date.fromisoformat(str(lp))
    except ValueError:
        return None


def _process_next(dry_run_attachment: Optional[Path] = None) -> int:
    """Process the OLDEST-file-date unprocessed Pending-Calls mail.

    On the live path this scans Outlook Inbox for the oldest mail whose
    attachment file-date hasn't been processed yet, saves it under the
    sender's original name, merges, and moves the mail to Inbox\\TSD\\CC.
    On dry-run, the given file is used directly.

    Exit codes:
      0 = success OR nothing to do
      2 = Outlook COM failure
      3 = save-attachment failure
      4 = read-attachment failure
      5 = paused (unknown columns — user must teach mapping)
      6 = union / write failure
    """
    RAW_FILES_DIR.mkdir(parents=True, exist_ok=True)
    state = _read_state()
    floor = _load_floor_date(state)   # None on very first run ever.

    # 1. Find the next mail to process (or use the dry-run one) ──────────────
    #    Save it under the sender's ORIGINAL filename — matches the user's
    #    manual practice. No canonical rename, no format conversion.
    msg = None    # live Outlook mail; kept around so step 8 can move it.
    if dry_run_attachment is not None:
        source_path  = dry_run_attachment
        source_label = dry_run_attachment.name
        file_date    = extract_file_date(source_label) or date.today()
        if floor is not None and file_date <= floor:
            log.info("Dry-run: file-date %s <= floor %s — already processed.",
                     file_date.isoformat(), floor.isoformat())
            return 0
        log.info("Dry-run: reading %s directly (file-date %s).",
                 source_label, file_date.isoformat())
    else:
        log.info("Scanning Outlook Inbox for oldest unprocessed Pending-Call mail…")
        try:
            found = _find_next_mail(floor)
        except Exception as e:  # noqa: BLE001
            log.error("Outlook COM failed: %s\n%s", e, traceback.format_exc())
            _toast("Auto-merge failed",
                   f"Outlook COM error: {type(e).__name__}. See auto_merge.log.")
            return 2
        if not found:
            log.info("No unprocessed Pending-Call mail in Inbox — nothing to do.")
            return 0
        msg, att_name, file_date = found
        source_path  = RAW_FILES_DIR / att_name
        try:
            _save_attachment(msg, att_name, source_path)
            log.info("Saved attachment as-is: %s (file-date %s)",
                     source_path.name, file_date.isoformat())
            source_label = source_path.name
        except Exception as e:  # noqa: BLE001
            log.error("Save attachment failed: %s\n%s", e, traceback.format_exc())
            _toast("Auto-merge failed",
                   f"Save attachment error: {type(e).__name__}. See auto_merge.log.")
            return 3

    # 2. Load the file + apply saved column map ──────────────────────────────
    try:
        today_df = _read_all_sheets_stacked(source_path, source_label=source_label)
    except Exception as e:  # noqa: BLE001
        log.error("Reading file failed: %s", e)
        _toast("Auto-merge failed",
               f"Couldn't read {source_path.name}: {type(e).__name__}.")
        return 4

    column_map = load_column_map(str(COLUMN_MAP_PATH))
    today_df   = apply_column_map(today_df, column_map)

    # 3. First-run seed vs steady-state ──────────────────────────────────────
    if not MERGED_AUTO_PATH.exists():
        log.info("First run — seeding from merged_output.xlsx (canonical columns).")
        prev = _seed_from_manual()
        prev = apply_column_map(prev, column_map) if not prev.empty else prev
        _audit("seeded",
               source_file=MERGED_MANUAL_PATH.name,
               canonical_column_count=len(prev.columns) if not prev.empty else 0)
    else:
        prev = _load_prev_merged()

    # 4. Unknown-column guard — ALWAYS enforced. merged_output.xlsx's columns
    #    (as loaded into prev) are the canonical set. Any incoming column not
    #    already canonical AND not listed as a source key in column_map.json
    #    triggers a pause + toast + audit-log entry. Strict: even whitespace,
    #    hyphen, or case differences are treated as "new column".
    known_canonical = set(prev.columns) if not prev.empty else set()
    unknown = unknown_columns(today_df, column_map, known_canonical)
    if unknown:
        log.warning(
            "New columns detected in %s — aborting merge until mapping is taught. "
            "Unknown columns: %s",
            source_label, ", ".join(unknown))
        _audit("pause",
               source_file=source_label,
               file_date=file_date.isoformat(),
               unknown_columns=list(unknown),
               known_column_count=len(known_canonical))
        _toast(
            "Auto-merge paused — new columns detected",
            f"{len(unknown)} new column(s) in {source_label}. "
            "Open the app once to teach the mapping.")
        # Do NOT mark this file-date processed — after the user teaches the
        # mapping via the Streamlit app, the next hourly run picks it up.
        return 5

    # 5. Union + write ───────────────────────────────────────────────────────
    try:
        added, total = _union_and_write(prev, today_df, today_source=source_label)
    except Exception as e:  # noqa: BLE001
        log.error("Merge/write failed: %s\n%s", e, traceback.format_exc())
        _toast("Auto-merge failed",
               f"Write error: {type(e).__name__}. See auto_merge.log.")
        return 6

    log.info(
        "OK  file_date=%s  source=%s  rows_added=%d  total_rows=%d  output=%s",
        file_date.isoformat(), source_label, added, total, MERGED_AUTO_PATH.name)
    _audit("merged",
           source_file=source_label,
           file_date=file_date.isoformat(),
           rows_added=added,
           total_rows=total,
           output_file=MERGED_AUTO_PATH.name)
    _toast(
        "Auto-merge complete",
        f"{added:,} new rows for {file_date.isoformat()} · total {total:,} rows.")

    # 6. Advance the floor. last_processed_date is monotonic: it only moves
    #    forward, so a later resend of an older file is naturally skipped.
    state["last_processed_date"]   = file_date.isoformat()
    state["last_processed_source"] = source_label
    state["last_processed_rows"]   = total
    state["last_processed_added"]  = added
    _write_state(state)

    # 7. Best-effort: sweep every Inbox Pending-Calls mail whose Excel
    #    file-date is <= the new floor (= file_date we just processed) into
    #    Inbox\TSD\CC. Catches morning + evening + resends + stragglers.
    #    Only in live Outlook mode (msg is None during dry-run).
    if msg is not None:
        _try_sweep_processed_to_cc(file_date)

    return 0


def main() -> int:
    # Minimal CLI: `--dry-run-attachment PATH` bypasses Outlook and uses the
    # given file directly (used by verification).
    dry_path: Optional[Path] = None
    args = sys.argv[1:]
    if args and args[0] == "--dry-run-attachment" and len(args) >= 2:
        dry_path = Path(args[1]).expanduser().resolve()
        if not dry_path.exists():
            print(f"--dry-run-attachment path does not exist: {dry_path}",
                  file=sys.stderr)
            return 2

    try:
        return _process_next(dry_run_attachment=dry_path)
    except Exception as e:  # noqa: BLE001
        log.critical("Unhandled error: %s\n%s", e, traceback.format_exc())
        _toast("Auto-merge crashed",
               f"{type(e).__name__}: {e}. See auto_merge.log.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
