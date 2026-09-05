"""
daily_auto_merge.py — Windows-scheduled unattended daily ZOHO / ERP Pending
Calls auto-merge.

Runs hourly Mon-Sat 08:00-12:00 (see install_scheduled_task.ps1). Each run:

    1. Idempotency check   — if today's mail was already processed, exit quiet.
    2. Outlook COM scan    — walk today's Inbox, keep mails whose subject
                             contains "Pending Call" (case-insensitive) AND
                             that carry exactly ONE attachment. Pick earliest.
    3. Save attachment     — Raw Files\\ZOHO_Pending_Calls_YYYY-MM-DD.xlsx.
    4. Apply column_map    — abort with a toast if any column of the new file
                             is unknown; user teaches it in the Streamlit app
                             and tomorrow's run succeeds.
    5. Union onto rolling  — first-run seed = merged_output.xlsx + today's
       merged_output_auto     file; subsequent runs = previous merged_output_auto
                             + today's file only.
    6. XlsxWriter output   — mirrors the app's OOM-safe writer.
    7. Notify + log        — winotify toast, rolling log at Raw Files\\auto_merge.log.

    Never touches Raw Files\\merged_output.xlsx (user's manual workflow).
    Never pushes anywhere. Never scans the whole Raw Files folder.

Runs entirely on-device: no network. Requires Outlook desktop client to be
open (it uses the same COM session the user's Outlook session provides).
"""

from __future__ import annotations

import json
import logging
import os
import shutil
import sys
import traceback
from datetime import date, datetime, time as dtime
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
COLUMN_MAP_PATH      = RAW_FILES_DIR / "column_map.json"
MERGED_MANUAL_PATH   = RAW_FILES_DIR / "merged_output.xlsx"          # DO NOT TOUCH
MERGED_AUTO_PATH     = RAW_FILES_DIR / "merged_output_auto.xlsx"     # ours

SUBJECT_MATCH        = "pending call"   # case-insensitive substring
LOG_MAX_BYTES        = 5 * 1024 * 1024  # 5 MB rotation
LOG_BACKUP_COUNT     = 3

APP_ID               = "ZOHO Auto-Merge"


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


def _canonical_download_name(d: date) -> str:
    return f"ZOHO_Pending_Calls_{d.isoformat()}.xlsx"


# ─── Outlook COM: find today's mail ──────────────────────────────────────────
def _find_todays_mail() -> Optional[Tuple[object, str]]:
    """
    Walk the default Inbox for messages received today whose subject matches
    SUBJECT_MATCH and that carry exactly one attachment. Return
    (MailItem, attachment_filename) for the earliest such message, or None.
    """
    import pythoncom            # noqa: F401  (needed to init COM in Task Scheduler context)
    import win32com.client as win32

    outlook   = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox     = namespace.GetDefaultFolder(6)   # 6 = olFolderInbox

    items = inbox.Items
    items.Sort("[ReceivedTime]", False)         # oldest first

    # Restrict to today's mail server-side — much faster than scanning all.
    today       = date.today()
    start_str   = datetime.combine(today, dtime.min).strftime("%m/%d/%Y %I:%M %p")
    end_str     = datetime.combine(today, dtime.max).strftime("%m/%d/%Y %I:%M %p")
    restriction = f"[ReceivedTime] >= '{start_str}' AND [ReceivedTime] <= '{end_str}'"
    try:
        items = items.Restrict(restriction)
    except Exception as e:  # noqa: BLE001
        log.warning("Outlook Restrict failed (%s) — scanning full inbox instead", e)

    candidates: List[Tuple[datetime, object, str]] = []
    for msg in items:
        try:
            subj = str(getattr(msg, "Subject", "") or "")
            if SUBJECT_MATCH not in subj.lower():
                continue
            atts = msg.Attachments
            if atts.Count != 1:
                continue        # filters the evening 2-attachment mail
            att = atts.Item(1)  # 1-indexed
            att_name = str(getattr(att, "FileName", "") or "")
            if not att_name.lower().endswith((".xls", ".xlsx", ".xlsm")):
                continue
            received = getattr(msg, "ReceivedTime", None)
            if received is None:
                continue
            candidates.append((received, msg, att_name))
        except Exception as e:  # noqa: BLE001
            log.warning("skipping a mail (%s): %s", type(e).__name__, e)
            continue

    if not candidates:
        return None
    candidates.sort(key=lambda t: t[0])
    _received, msg, att_name = candidates[0]
    return msg, att_name


def _save_attachment(msg, att_name: str, dest_path: Path) -> None:
    """Save the (single) attachment of msg to dest_path, converting .xls → .xlsx
    if needed by round-tripping through pandas."""
    tmp_path = dest_path.with_suffix(Path(att_name).suffix.lower())
    # Delete any stale temp with the same name so SaveAsFile doesn't refuse.
    try:
        if tmp_path.exists():
            tmp_path.unlink()
    except OSError:
        pass
    msg.Attachments.Item(1).SaveAsFile(str(tmp_path))

    if tmp_path.suffix.lower() == ".xlsx":
        if tmp_path != dest_path:
            shutil.move(str(tmp_path), str(dest_path))
        return

    # .xls → .xlsx: read every sheet, write via merge_core (XlsxWriter).
    sheets = read_from_path(str(tmp_path))
    from merge_core import to_excel_file as _to_excel_file
    _to_excel_file(sheets, str(dest_path))
    try:
        tmp_path.unlink()
    except OSError:
        pass


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
def _process_today(today: date, dry_run_attachment: Optional[Path] = None) -> int:
    """Main pipeline for one run. Returns process exit code.

    dry_run_attachment: if given, skip the Outlook step entirely and treat
    this .xls/.xlsx file as the attachment already saved for today. Used by
    the local dry-run harness.
    """
    RAW_FILES_DIR.mkdir(parents=True, exist_ok=True)
    state       = _read_state()
    today_iso   = today.isoformat()
    dest_path   = RAW_FILES_DIR / _canonical_download_name(today)

    # 1. Idempotency ─────────────────────────────────────────────────────────
    if state.get("last_processed_date") == today_iso:
        log.info("Nothing to do — already processed %s.", today_iso)
        return 0

    # 2. Find + save today's attachment (or use the dry-run one) ─────────────
    if dry_run_attachment is not None:
        if not dest_path.exists():
            # Convert / copy the dry-run source into the canonical name.
            if dry_run_attachment.suffix.lower() == ".xlsx":
                shutil.copyfile(dry_run_attachment, dest_path)
            else:
                sheets = read_from_path(str(dry_run_attachment))
                to_excel_file(sheets, str(dest_path))
        source_label = dry_run_attachment.name
        log.info("Dry-run: using %s as today's attachment → %s",
                 dry_run_attachment.name, dest_path.name)
    else:
        if dest_path.exists():
            log.info("Today's file already downloaded (%s) — skipping mail scan.",
                     dest_path.name)
            source_label = dest_path.name
        else:
            log.info("Scanning Outlook Inbox for today's Pending-Call mail…")
            try:
                found = _find_todays_mail()
            except Exception as e:  # noqa: BLE001
                log.error("Outlook COM failed: %s\n%s", e, traceback.format_exc())
                _toast("Auto-merge failed",
                       f"Outlook COM error: {type(e).__name__}. See auto_merge.log.")
                return 2
            if not found:
                log.info("No qualifying mail in today's Inbox — will retry next hour.")
                return 0
            msg, att_name = found
            try:
                _save_attachment(msg, att_name, dest_path)
                log.info("Saved attachment '%s' → %s", att_name, dest_path.name)
                source_label = dest_path.name
            except Exception as e:  # noqa: BLE001
                log.error("Save attachment failed: %s\n%s", e, traceback.format_exc())
                _toast("Auto-merge failed",
                       f"Save attachment error: {type(e).__name__}. See auto_merge.log.")
                return 3

    # 3. Load today's file + apply saved column map ──────────────────────────
    try:
        today_df = _read_all_sheets_stacked(dest_path, source_label=source_label)
    except Exception as e:  # noqa: BLE001
        log.error("Reading today's file failed: %s", e)
        _toast("Auto-merge failed",
               f"Couldn't read {dest_path.name}: {type(e).__name__}.")
        return 4

    column_map = load_column_map(str(COLUMN_MAP_PATH))
    today_df   = apply_column_map(today_df, column_map)

    # 4. First-run seed vs steady-state ──────────────────────────────────────
    seeded = False
    if not MERGED_AUTO_PATH.exists():
        log.info("First run — seeding from merged_output.xlsx.")
        prev = _seed_from_manual()
        # Ensure the seed's columns also pass through the column map so it
        # lines up with today_df.
        prev = apply_column_map(prev, column_map) if not prev.empty else prev
        seeded = True
    else:
        prev = _load_prev_merged()

    # 5. Unknown-column guard (skip on first-run seed — user's baseline is
    #    treated as authoritative) ──────────────────────────────────────────
    if not seeded:
        known_canonical = set(prev.columns) if not prev.empty else set()
        unknown = unknown_columns(today_df, column_map, known_canonical)
        if unknown:
            log.warning(
                "New columns detected in %s — aborting merge until mapping is taught. "
                "Unknown columns: %s",
                dest_path.name, ", ".join(unknown))
            _toast(
                "Auto-merge paused — new columns detected",
                f"{len(unknown)} new column(s) in {dest_path.name}. "
                "Open the app once to teach the mapping.")
            # Do NOT mark today as processed — tomorrow's mapping fix should
            # let a subsequent run pick this up cleanly.
            return 5

    # 6. Union + write ───────────────────────────────────────────────────────
    try:
        added, total = _union_and_write(prev, today_df, today_source=source_label)
    except Exception as e:  # noqa: BLE001
        log.error("Merge/write failed: %s\n%s", e, traceback.format_exc())
        _toast("Auto-merge failed",
               f"Write error: {type(e).__name__}. See auto_merge.log.")
        return 6

    log.info(
        "OK  date=%s  source=%s  rows_added=%d  total_rows=%d  output=%s",
        today_iso, source_label, added, total, MERGED_AUTO_PATH.name)
    _toast(
        "Auto-merge complete",
        f"{added:,} new rows for {today_iso} · total {total:,} rows.")

    # 7. Mark today processed ────────────────────────────────────────────────
    state["last_processed_date"]   = today_iso
    state["last_processed_source"] = source_label
    state["last_processed_rows"]   = total
    state["last_processed_added"]  = added
    _write_state(state)
    return 0


def main() -> int:
    # Minimal CLI: `--dry-run-attachment PATH` bypasses Outlook and uses the
    # given file as if it were today's attachment (used by verification).
    dry_path: Optional[Path] = None
    args = sys.argv[1:]
    if args and args[0] == "--dry-run-attachment" and len(args) >= 2:
        dry_path = Path(args[1]).expanduser().resolve()
        if not dry_path.exists():
            print(f"--dry-run-attachment path does not exist: {dry_path}",
                  file=sys.stderr)
            return 2

    try:
        return _process_today(date.today(), dry_run_attachment=dry_path)
    except Exception as e:  # noqa: BLE001
        log.critical("Unhandled error: %s\n%s", e, traceback.format_exc())
        _toast("Auto-merge crashed",
               f"{type(e).__name__}: {e}. See auto_merge.log.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
