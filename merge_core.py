"""
merge_core.py — pure-pandas helpers shared by DataMerge Studio and the daily
auto-merge script.

Nothing in this module imports Streamlit, so it is safe to import from a plain
Python script (e.g. `daily_auto_merge.py`) or from within the app.

Extracted from file_merger_app.py on 2026-09-05 as part of the daily
auto-merge feature. Behavior is preserved 1:1 — the app continues to import
these names and use them exactly as before.
"""

from __future__ import annotations

import io
import json
import os
import re
from typing import Dict, Iterable, List, Tuple

import pandas as pd


# ─── Constants (shared with the app) ──────────────────────────────────────────
TRACKING_COLS   = {"Source File", "Source Date", "Source Sheet"}
DATE_KEYWORDS   = {"date", "time", "created", "modified", "due", "updated",
                   "closed", "opened", "reported", "raised", "logged"}
CITY_KEYWORDS   = {"city", "town"}
REGION_KEYWORDS = {"region", "zone", "area", "territory", "state", "country"}

EXCEL_ROW_LIMIT = 1_048_576


# ─── Column helpers ───────────────────────────────────────────────────────────
def _col_words(name: str) -> set:
    return set(re.split(r"[\s_\-/()+]+", name.lower()))


def _dedup_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Return df with any duplicate column names made unique (col → col_2, col_3…)."""
    if not df.columns.duplicated().any():
        return df
    seen: dict = {}
    new_cols = []
    for c in df.columns:
        if c not in seen:
            seen[c] = 0
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c] + 1}")
    df = df.copy()
    df.columns = new_cols
    return df


def col_sig(df: pd.DataFrame) -> frozenset:
    return frozenset(c for c in df.columns if c not in TRACKING_COLS)


def group_sheets(file_sheet_dfs):
    buckets = {}
    for fname, sheet, df in file_sheet_dfs:
        buckets.setdefault(col_sig(df), []).append((fname, sheet, df))
    return sorted(buckets.items(), key=lambda x: -len(x[0]))


# ─── Data-type cleaning ───────────────────────────────────────────────────────
def clean_dtypes(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """
    Auto-clean column types in-place.
    Returns (cleaned_df, report_list).
      - Date-sounding columns: try pd.to_datetime (dayfirst=True)
      - City/region columns:   strip whitespace + title-case
      - All other str columns: strip whitespace

    Defensive: skips non-string column names, duplicate column names, and any
    column that raises unexpectedly. Identical to the app's implementation so
    the auto-merge output matches what the app produces.
    """
    df = df.copy()
    report: List[str] = []
    duped = set(df.columns[df.columns.duplicated(keep=False)])

    for col in df.columns:
        if col in TRACKING_COLS or col in duped:
            continue
        if not isinstance(col, str):
            continue

        try:
            series = df[col]
            if isinstance(series, pd.DataFrame):
                continue
        except Exception:
            continue

        words = _col_words(col)

        if words & DATE_KEYWORDS and series.dtype == object:
            conv  = pd.to_datetime(series, errors="coerce", dayfirst=True)
            total = int(series.notna().sum())
            hit   = int(conv.notna().sum())
            if total > 0 and hit / total >= 0.5:
                df[col] = conv
                report.append(f"'{col}' → datetime  ({hit}/{total} values parsed)")
                continue

        if series.dtype == object:
            before  = series.fillna("").copy()
            df[col] = series.str.strip()
            if words & (CITY_KEYWORDS | REGION_KEYWORDS):
                df[col] = df[col].str.title()
                report.append(f"'{col}' → stripped + title-cased (city/region)")
            elif (df[col].fillna("") != before).any():
                report.append(f"'{col}' → stripped whitespace")

    return df, report


# ─── Duplicate audit ──────────────────────────────────────────────────────────
def dedup_with_audit(df: pd.DataFrame, check_cols: Iterable[str]):
    """Remove duplicates (keep first) and return (clean_df, audit_df)."""
    valid = [c for c in check_cols if c in df.columns]
    if not valid:
        return df.copy(), pd.DataFrame()
    mask    = df.duplicated(subset=valid, keep="first")
    removed = df[mask].copy()
    if not removed.empty:
        removed.insert(0, "Removed Row# (Excel)", [i + 2 for i in df[mask].index])
    return df[~mask].reset_index(drop=True), removed


# ─── Merge / join strategies (pure pandas) ────────────────────────────────────
def do_union_all(dfs, key=None, excl=None):
    """Stack all rows vertically. Columns aligned by name (outer join on columns)."""
    return pd.concat(dfs, ignore_index=True, join="outer"), pd.DataFrame()


def do_union_distinct(dfs, key=None, excl=None):
    """Stack vertically, then drop rows where every checked column is identical."""
    combined = pd.concat(dfs, ignore_index=True, join="outer")
    check = [c for c in combined.columns
             if c not in (excl or set()) and c not in TRACKING_COLS]
    return dedup_with_audit(combined, check)


def _chain_join(dfs, key, how):
    """Chain pd.merge left-to-right. All dfs must contain the key column."""
    missing = [i + 1 for i, d in enumerate(dfs) if key not in d.columns]
    if missing:
        raise ValueError(
            f"Key column '{key}' is missing in file/sheet #{missing}. "
            f"Add it (or remap a column to '{key}') before joining.")
    result = dfs[0].copy()
    for i, d in enumerate(dfs[1:], start=2):
        result = pd.merge(result, d, on=key, how=how, suffixes=("", f"__t{i}"))
    return result


def do_inner_join(dfs, key, excl=None):
    if not key:
        raise ValueError("Inner Join requires a key column.")
    if len(dfs) == 1:
        return dfs[0].copy(), pd.DataFrame()
    return _chain_join(dfs, key, "inner"), pd.DataFrame()


def do_left_join(dfs, key, excl=None):
    if not key:
        raise ValueError("Left Join requires a key column.")
    if len(dfs) == 1:
        return dfs[0].copy(), pd.DataFrame()
    return _chain_join(dfs, key, "left"), pd.DataFrame()


def do_right_join(dfs, key, excl=None):
    if not key:
        raise ValueError("Right Join requires a key column.")
    if len(dfs) == 1:
        return dfs[0].copy(), pd.DataFrame()
    return _chain_join(dfs, key, "right"), pd.DataFrame()


def do_full_outer_join(dfs, key, excl=None):
    if not key:
        raise ValueError("Full Outer Join requires a key column.")
    if len(dfs) == 1:
        return dfs[0].copy(), pd.DataFrame()
    return _chain_join(dfs, key, "outer"), pd.DataFrame()


def do_cross_join(dfs, key=None, excl=None):
    """Cartesian product of all files. Pre-checks against Excel row limit."""
    if len(dfs) == 1:
        return dfs[0].copy(), pd.DataFrame()
    rows = 1
    for d in dfs:
        rows *= len(d)
    if rows > EXCEL_ROW_LIMIT:
        raise ValueError(
            f"Cross join would produce {rows:,} rows — exceeds Excel's "
            f"{EXCEL_ROW_LIMIT:,} row limit. Reduce inputs first.")
    result = dfs[0].copy()
    for i, d in enumerate(dfs[1:], start=2):
        result = pd.merge(result, d, how="cross", suffixes=("", f"__t{i}"))
    return result, pd.DataFrame()


# ─── I/O helpers ──────────────────────────────────────────────────────────────
def read_from_path(path: str) -> Dict[str, pd.DataFrame]:
    """Read a CSV or Excel file. Returns {sheet_name: DataFrame}."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return {"Sheet1": pd.read_csv(path)}
    engines = ["openpyxl", "xlrd"] if ext == ".xlsx" else ["xlrd", "openpyxl"]
    last_err = None
    for eng in engines:
        try:
            xl = pd.ExcelFile(path, engine=eng)
            return {s: xl.parse(s) for s in xl.sheet_names}
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Cannot read {path}: {last_err}")


def _make_excel_writer(buf):
    """Prefer XlsxWriter over openpyxl (streams to temp file → much lower RAM)."""
    try:
        return pd.ExcelWriter(buf, engine="xlsxwriter")
    except (ImportError, ModuleNotFoundError, ValueError):
        return pd.ExcelWriter(buf, engine="openpyxl")


def _safe_sheet_name(name: str) -> str:
    return name[:31].translate(str.maketrans(r'\/[]*?:', '_______'))


def to_excel_bytes(sheet_dict: Dict[str, pd.DataFrame]) -> bytes:
    """Serialize {sheet_name: df} to a multi-sheet .xlsx byte-string."""
    buf = io.BytesIO()
    with _make_excel_writer(buf) as w:
        for name, df in sheet_dict.items():
            df.to_excel(w, sheet_name=_safe_sheet_name(name), index=False)
    return buf.getvalue()


def to_excel_file(sheet_dict: Dict[str, pd.DataFrame], path: str) -> None:
    """Write {sheet_name: df} straight to an .xlsx file on disk (XlsxWriter)."""
    with _make_excel_writer(path) as w:
        for name, df in sheet_dict.items():
            df.to_excel(w, sheet_name=_safe_sheet_name(name), index=False)


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


# ─── Column-map persistence (used by both the app and the auto-merge script) ─
def load_column_map(path: str) -> Dict[str, str]:
    """
    Load canonical column mapping from JSON on disk.
    Shape: {"source column name": "canonical column name", ...}
    Returns {} if the file doesn't exist or is empty/invalid.
    """
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {}
        return {str(k): str(v) for k, v in data.items() if k and v}
    except (OSError, json.JSONDecodeError):
        return {}


def save_column_map(path: str, mapping: Dict[str, str]) -> None:
    """
    Persist canonical column mapping to JSON on disk (pretty-printed, sorted).
    Overwrites any existing file. Callers that want a merge-in behavior should
    load, update, and pass the merged dict.
    """
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tidy = {str(k): str(v) for k, v in mapping.items() if k and v}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(tidy, f, indent=2, sort_keys=True, ensure_ascii=False)


def flatten_renames_map(renames_map: Dict[str, Dict[str, str]]) -> Dict[str, str]:
    """
    Convert the app's per-file rename map {fname: {orig: canonical}}
    into the flat auto-merge shape {orig: canonical}.

    On conflict (same orig mapped to different canonical values across files),
    the LAST value wins — matches "apply the newest confirmed mapping" intent.
    """
    flat: Dict[str, str] = {}
    for _fname, m in (renames_map or {}).items():
        for orig, canonical in (m or {}).items():
            if orig and canonical:
                flat[str(orig)] = str(canonical)
    return flat


def merge_and_save_column_map(path: str,
                              new_flat: Dict[str, str]) -> Dict[str, str]:
    """Load the on-disk map, merge new_flat on top, save, and return the result."""
    current = load_column_map(path)
    current.update({k: v for k, v in (new_flat or {}).items() if k and v})
    save_column_map(path, current)
    return current


def apply_column_map(df: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
    """Rename df's columns according to `mapping` (source → canonical).
    Columns not in the mapping are left as-is."""
    if not mapping:
        return df
    rn = {c: mapping[c] for c in df.columns if c in mapping}
    return df.rename(columns=rn) if rn else df


def unknown_columns(df: pd.DataFrame,
                    mapping: Dict[str, str],
                    known_canonical: Iterable[str] = ()) -> List[str]:
    """
    Return columns in df that are neither:
      - already a known canonical column (present in `known_canonical`), nor
      - listed as a source key in `mapping`, nor
      - a tracking column (Source File / Source Date / Source Sheet).

    Used by the daily auto-merge script to decide whether to abort and ask the
    user to teach the mapping before unioning today's file.
    """
    known = set(known_canonical) | set((mapping or {}).values()) | TRACKING_COLS
    src_keys = set((mapping or {}).keys())
    out = []
    for c in df.columns:
        if c in known or c in src_keys:
            continue
        out.append(c)
    return out
