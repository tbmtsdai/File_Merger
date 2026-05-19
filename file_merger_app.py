"""
DataMerge Studio — Streamlit application for merging and joining CSV/Excel files.
Implements all standard SQL set operations (Union All, Union All Distinct,
Inner / Left / Right / Full Outer / Cross Join) with column mapping across files
with different column names, data-type cleaning, duplicate-audit reporting,
and an interactive dashboard with offline-capable HTML export.
"""

import io, os, re
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from matplotlib_venn import venn2, venn2_circles

try:
    import tkinter as tk
    from tkinter import filedialog as _tkfd
    HAS_TKINTER = True
except Exception:
    HAS_TKINTER = False

try:
    import plotly.express as px
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

st.set_page_config(page_title="DataMerge Studio", page_icon="🔀",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
.badge {display:inline-block;padding:3px 10px;border-radius:12px;
        font-size:.78em;font-weight:600;margin-bottom:8px;}
.badge-blue   {background:#dbeafe;color:#1d4ed8;}
.badge-green  {background:#dcfce7;color:#166534;}
.badge-orange {background:#fef3c7;color:#92400e;}
.badge-red    {background:#fee2e2;color:#991b1b;}
.badge-purple {background:#ede9fe;color:#5b21b6;}
.badge-teal   {background:#ccfbf1;color:#0f766e;}
.group-box {background:#f8fafc;border:1px solid #e2e8f0;
            border-radius:8px;padding:12px;margin-bottom:10px;}
</style>""", unsafe_allow_html=True)


# ─── Constants ─────────────────────────────────────────────────────────────────
TRACKING_COLS   = {"Source File", "Source Date", "Source Sheet"}
DATE_KEYWORDS   = {"date","time","created","modified","due","updated","closed",
                   "opened","reported","raised","logged"}
CITY_KEYWORDS   = {"city","town"}
REGION_KEYWORDS = {"region","zone","area","territory","state","country"}


# ─── Dummy data (Learn tab) ────────────────────────────────────────────────────
@st.cache_data
def get_dummy_join():
    """Demo data for join strategies — File A = Employees, File B = Payroll."""
    fa = pd.DataFrame({
        "Employee ID": ["E001", "E002", "E003", "E004", "E005"],
        "Name":        ["Alice", "Bob", "Charlie", "Diana", "Eve"],
        "Department":  ["Engineering", "Marketing", "Engineering", "HR", "Finance"],
    })
    fb = pd.DataFrame({
        "Employee ID": ["E003", "E004", "E005", "E006", "E007"],
        "Salary":      [90000, 71500, 78000, 88000, 76000],
        "Manager":     ["Tom", "Raj", "Sara", "Raj", "Sara"],
    })
    return fa, fb


@st.cache_data
def get_dummy_union():
    """Demo data for union strategies — same columns, some duplicate rows."""
    fa = pd.DataFrame({
        "Order ID":   ["O001", "O002", "O003"],
        "Customer":   ["Alice", "Bob", "Charlie"],
        "Amount":     [120, 350, 80],
    })
    fb = pd.DataFrame({
        "Order ID":   ["O003", "O004", "O005"],
        "Customer":   ["Charlie", "Diana", "Eve"],
        "Amount":     [80, 200, 150],     # O003 row identical to File A
    })
    return fa, fb


# Backwards-compat alias (used by render_dashboard etc.)
get_dummy = get_dummy_join


# ═══════════════════════════════════════════════════════════════════════════════
# DATA TYPE CLEANING
# ═══════════════════════════════════════════════════════════════════════════════

def _col_words(name):
    return set(re.split(r"[\s_\-/()+]+", name.lower()))


def _dedup_columns(df):
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


def clean_dtypes(df):
    """
    Auto-clean column types in-place.
    Returns (cleaned_df, report_list).
      - Date-sounding columns: try pd.to_datetime (dayfirst=True)
      - City/region columns:   strip whitespace + title-case
      - All other str columns: strip whitespace

    Defensive: skips non-string column names, duplicate column names
    (where df[col] would return a DataFrame rather than a Series), and
    any column that raises unexpectedly.
    """
    df = df.copy()
    report = []
    # Build a set of column names that appear more than once — accessing
    # df[col] for a duplicate name returns a DataFrame, not a Series,
    # which causes AttributeError on .dtype.  Skip all copies of such names.
    duped = set(df.columns[df.columns.duplicated(keep=False)])

    for col in df.columns:
        # Skip tracking cols, duplicates, and non-string names
        if col in TRACKING_COLS or col in duped:
            continue
        if not isinstance(col, str):
            continue

        try:
            series = df[col]
            # Extra guard: if somehow still a DataFrame, skip
            if isinstance(series, pd.DataFrame):
                continue
        except Exception:
            continue

        words = _col_words(col)

        # ── Date columns ──────────────────────────────────────────────────
        if words & DATE_KEYWORDS and series.dtype == object:
            conv  = pd.to_datetime(series, errors="coerce", dayfirst=True)
            total = int(series.notna().sum())
            hit   = int(conv.notna().sum())
            if total > 0 and hit / total >= 0.5:
                df[col] = conv
                report.append(f"'{col}' → datetime  ({hit}/{total} values parsed)")
                continue

        # ── String cleanup ────────────────────────────────────────────────
        if series.dtype == object:
            before  = series.fillna("").copy()
            df[col] = series.str.strip()
            if words & (CITY_KEYWORDS | REGION_KEYWORDS):
                df[col] = df[col].str.title()
                report.append(f"'{col}' → stripped + title-cased (city/region)")
            elif (df[col].fillna("") != before).any():
                report.append(f"'{col}' → stripped whitespace")

    return df, report


# ═══════════════════════════════════════════════════════════════════════════════
# DUPLICATE AUDIT
# ═══════════════════════════════════════════════════════════════════════════════

def dedup_with_audit(df, check_cols):
    """
    Remove duplicates (keep first occurrence) and return (clean_df, audit_df).

    audit_df contains every removed row plus a 'Removed Row# (Excel)' column
    showing the 1-based row number in the pre-dedup combined frame (row 2 = first
    data row, matching Excel's header-on-row-1 convention).

    To plug in a different dedup strategy: write a new function with signature
        fn(dfs, key, excl) -> (result_df, audit_df)
    and add it to STRATEGIES below.
    """
    valid = [c for c in check_cols if c in df.columns]
    if not valid:
        return df.copy(), pd.DataFrame()
    mask    = df.duplicated(subset=valid, keep="first")
    removed = df[mask].copy()
    if not removed.empty:
        removed.insert(0, "Removed Row# (Excel)", [i + 2 for i in df[mask].index])
    return df[~mask].reset_index(drop=True), removed


def show_audit(all_audits):
    """Render the combined duplicate-audit section after a merge."""
    non_empty = [a for a in all_audits if not a.empty]
    if not non_empty:
        return
    audit_df = pd.concat(non_empty, ignore_index=True)
    n = len(audit_df)
    with st.expander(f"Duplicate Audit — {n:,} row(s) removed", expanded=True):
        st.caption(
            "'Removed Row# (Excel)' is the 1-based row number in the combined "
            "pre-dedup dataset (row 1 = header, row 2 = first data row).")
        st.dataframe(audit_df, use_container_width=True, hide_index=True)
        st.download_button(
            "Download duplicate audit (.csv)",
            data=audit_df.to_csv(index=False).encode("utf-8"),
            file_name="duplicate_audit.csv",
            mime="text/csv")


# ═══════════════════════════════════════════════════════════════════════════════
# MERGE / JOIN FUNCTIONS  — all return (result_df, audit_df)
#
# Standard SQL set operations implemented on top of pandas:
#   • Union All              → pd.concat
#   • Union All (Distinct)   → pd.concat → drop_duplicates
#   • Inner / Left / Right / Full Outer / Cross  → chained pd.merge
#
# Multi-file joins are chained left-to-right:  ((A ⨝ B) ⨝ C) ⨝ D ...
# All operations are pure pandas — no custom logic, fully standard.
# ═══════════════════════════════════════════════════════════════════════════════

EXCEL_ROW_LIMIT = 1_048_576


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
    # Validate key column is present everywhere
    missing = [i + 1 for i, d in enumerate(dfs) if key not in d.columns]
    if missing:
        raise ValueError(
            f"Key column '{key}' is missing in file/sheet #{missing}. "
            f"Add it (or remap a column to '{key}') before joining.")

    result = dfs[0].copy()
    for i, d in enumerate(dfs[1:], start=2):
        result = pd.merge(
            result, d, on=key, how=how,
            suffixes=("", f"__t{i}"))
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
        result = pd.merge(
            result, d, how="cross",
            suffixes=("", f"__t{i}"))
    return result, pd.DataFrame()


STRATEGIES = {
    "Union All": dict(
        fn=do_union_all, needs_key=False, allows_excl=False,
        merge_all_groups=True, family="union",
        icon="➕", badge="badge-blue", badge_lbl="No key needed",
        head="Stack all rows vertically. Every row kept, including duplicates.",
        detail="Equivalent to SQL `UNION ALL`. Columns are matched by name; "
               "any column missing in one file becomes blank for that file's rows.",
        best="Combining periodic exports (e.g. daily/monthly files) that don't overlap.",
    ),
    "Union All (Distinct)": dict(
        fn=do_union_distinct, needs_key=False, allows_excl=True,
        merge_all_groups=True, family="union",
        icon="🧹", badge="badge-green", badge_lbl="No key · exclude-cols optional",
        head="Stack all rows vertically, then remove exact-duplicate rows.",
        detail="Equivalent to SQL `UNION` (which deduplicates by default). "
               "A row is a duplicate only if EVERY checked column is identical. "
               "You can optionally exclude specific columns from the duplicate "
               "check (e.g. an 'Assigned To' column where minor differences "
               "shouldn't prevent two rows being treated as the same).",
        best="Two exports of the same data with minor irrelevant field differences.",
    ),
    "Inner Join": dict(
        fn=do_inner_join, needs_key=True, allows_excl=False,
        merge_all_groups=True, family="join",
        icon="🔍", badge="badge-purple", badge_lbl="Key column required",
        head="Keep only rows whose key exists in EVERY file. Columns combined.",
        detail="Equivalent to SQL `INNER JOIN`. For each matching key, columns "
               "from all files are placed side-by-side. Rows whose key appears "
               "in only some files are dropped.",
        best="Records confirmed in every source — e.g. customers active in both Jan and Feb.",
    ),
    "Left Join": dict(
        fn=do_left_join, needs_key=True, allows_excl=False,
        merge_all_groups=True, family="join",
        icon="⬅️", badge="badge-teal", badge_lbl="Key column required",
        head="Keep ALL rows from the FIRST file. Match rows from later files.",
        detail="Equivalent to SQL `LEFT JOIN`. Every row from File 1 is kept; "
               "matching data from File 2/3/... is added on the right (NULL "
               "where no match). Rows unique to later files are dropped.",
        best="Enriching a master list (File 1) with extra info from another file.",
    ),
    "Right Join": dict(
        fn=do_right_join, needs_key=True, allows_excl=False,
        merge_all_groups=True, family="join",
        icon="➡️", badge="badge-orange", badge_lbl="Key column required",
        head="Keep ALL rows from the LAST file. Match rows from earlier files.",
        detail="Equivalent to SQL `RIGHT JOIN`. Mirror image of Left Join — "
               "every row from the last file is kept; matching data from earlier "
               "files is added on the left (NULL where no match).",
        best="When the latest file is the source of truth.",
    ),
    "Full Outer Join": dict(
        fn=do_full_outer_join, needs_key=True, allows_excl=False,
        merge_all_groups=True, family="join",
        icon="🔗", badge="badge-red", badge_lbl="Key column required",
        head="Keep EVERY row from EVERY file. Pair up where the key matches.",
        detail="Equivalent to SQL `FULL OUTER JOIN`. Nothing is dropped. "
               "Rows that share a key in multiple files are combined into one row. "
               "Rows unique to any file appear with NULLs for the other files' columns.",
        best="Building a complete master list from disparate sources.",
    ),
    "Cross Join": dict(
        fn=do_cross_join, needs_key=False, allows_excl=False,
        merge_all_groups=True, family="join",
        icon="✖️", badge="badge-red", badge_lbl="No key · cartesian product",
        head="Pair every row of File 1 with every row of File 2 (×3, ×4...).",
        detail="Equivalent to SQL `CROSS JOIN`. Produces N₁ × N₂ × ... rows. "
               "Use sparingly — easily exceeds Excel's 1,048,576 row limit. "
               "App pre-checks the row count and refuses if too large.",
        best="Generating all combinations (e.g. every product × every region).",
    ),
}


# ─── Row fingerprint (Learn tab annotate) ──────────────────────────────────────
def _row_sig(row):
    parts = []
    for v in row:
        try:
            if pd.isna(v):
                parts.append("__NaN__"); continue
        except (TypeError, ValueError):
            pass
        try:
            f = float(v)
            parts.append(str(int(f)) if f == int(f) else str(round(f, 8)))
        except (ValueError, TypeError):
            parts.append(str(v).strip())
    return tuple(parts)


def annotate(fa, fb, result, key):
    r = result.copy()
    if key and key in r.columns:
        a_k, b_k = set(fa[key].dropna()), set(fb[key].dropna())
        r.insert(0, "Source",
                 ["🟡 Both"    if v in a_k and v in b_k
                  else "🟢 File A" if v in a_k
                  else "🔵 File B" for v in r[key]])
    else:
        a_sigs = {_row_sig(row) for row in fa.values}
        b_sigs = {_row_sig(row) for row in fb.values}
        r.insert(0, "Source", [
            "🟡 Both files"    if _row_sig(row) in a_sigs and _row_sig(row) in b_sigs
            else "🟢 File A only" if _row_sig(row) in a_sigs
            else "🔵 File B only"
            for row in result.values])
    return r


# ─── Venn diagrams (Learn tab) ─────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def make_venn_fig(strategy_name):
    C = dict(a="#4ade80", b="#60a5fa", both="#fbbf24", fill="#c084fc", gray="#e5e7eb")
    CFGS = {
        "Union All": dict(
            c10=C["a"], c11=C["both"], c01=C["b"],
            l10="O001\nO002", l11="O003\n(×2 copies)", l01="O004\nO005",
            title="Union All — vertical stack",
            desc="All rows from all files kept (duplicates included)",
            rows="6 rows out (3 + 3, nothing removed)",
        ),
        "Union All (Distinct)": dict(
            c10=C["a"], c11=C["both"], c01=C["b"],
            l10="O001\nO002", l11="O003\n(1 copy)", l01="O004\nO005",
            title="Union All (Distinct) — stack then dedupe",
            desc="Stack vertically, then drop rows identical in every column",
            rows="5 rows out (O003 collapsed; differing rows kept)",
        ),
        "Inner Join": dict(
            c10=C["gray"], c11=C["both"], c01=C["gray"],
            l10="E001,E002\ndropped", l11="E003\nE004\nE005", l01="E006,E007\ndropped",
            title="Inner Join — intersection",
            desc="Only keys present in BOTH files survive; columns combined",
            rows="3 rows out (E003, E004, E005 with combined columns)",
        ),
        "Left Join": dict(
            c10=C["a"], c11=C["both"], c01=C["gray"],
            l10="E001,E002\n(no match)", l11="E003,E004\nE005\n(matched)",
            l01="E006,E007\ndropped",
            title="Left Join — keep File A",
            desc="All File A rows kept; File B columns added where key matches",
            rows="5 rows out (E001/E002 have NULLs for File B columns)",
        ),
        "Right Join": dict(
            c10=C["gray"], c11=C["both"], c01=C["b"],
            l10="E001,E002\ndropped", l11="E003,E004\nE005\n(matched)",
            l01="E006,E007\n(no match)",
            title="Right Join — keep File B",
            desc="All File B rows kept; File A columns added where key matches",
            rows="5 rows out (E006/E007 have NULLs for File A columns)",
        ),
        "Full Outer Join": dict(
            c10=C["a"], c11=C["both"], c01=C["b"],
            l10="E001,E002", l11="E003,E004\nE005", l01="E006,E007",
            title="Full Outer Join — keep everything",
            desc="Every row from every file kept; matched rows combined",
            rows="7 rows out (5 + 5 minus 3 matches; NULLs where no match)",
        ),
        "Cross Join": dict(
            c10=C["a"], c11=C["fill"], c01=C["b"],
            l10="A1×B1,A1×B2,…", l11="every A\n×\nevery B",
            l01="A5×B1,A5×B2,…",
            title="Cross Join — cartesian product",
            desc="Every row of File A paired with every row of File B",
            rows="N₁ × N₂ rows (5 × 5 = 25 rows here)",
        ),
    }
    cfg = CFGS.get(strategy_name)
    if cfg is None:
        return None
    fig, ax = plt.subplots(figsize=(5, 3.8))
    fig.patch.set_facecolor("#f9fafb")
    ax.set_facecolor("#f9fafb")
    v = venn2(subsets=(2, 2, 3), set_labels=("File A", "File B"), ax=ax)
    venn2_circles(subsets=(2, 2, 3), ax=ax, color="#9ca3af", linewidth=1.5)
    for rid, col in [("10", cfg["c10"]), ("11", cfg["c11"]), ("01", cfg["c01"])]:
        p = v.get_patch_by_id(rid)
        if p:
            p.set_facecolor(col); p.set_alpha(0.82)
    for rid, txt in [("10", cfg["l10"]), ("11", cfg["l11"]), ("01", cfg["l01"])]:
        lbl = v.get_label_by_id(rid)
        if lbl:
            lbl.set_text(txt); lbl.set_fontsize(7)
    for sid, col in [("A", "#166534"), ("B", "#1d4ed8")]:
        lbl = v.get_label_by_id(sid)
        if lbl:
            lbl.set_fontsize(10); lbl.set_fontweight("bold"); lbl.set_color(col)
    ax.set_title(cfg["title"], fontsize=10, fontweight="bold", color="#111827", pad=8)
    fig.text(0.5, 0.13, cfg["desc"], ha="center", fontsize=7.5, color="#374151")
    fig.text(0.5, 0.04, cfg["rows"],  ha="center", fontsize=7, color="#6b7280",
             style="italic")
    plt.tight_layout(rect=[0, 0.16, 1, 1])
    return fig


# ═══════════════════════════════════════════════════════════════════════════════
# FILE READING HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def read_all_sheets_cached(file_bytes, filename):
    ext = os.path.splitext(filename)[1].lower()
    buf = io.BytesIO(file_bytes)
    if ext == ".csv":
        return {"Sheet1": pd.read_csv(buf)}
    engines = ["openpyxl","xlrd"] if ext == ".xlsx" else ["xlrd","openpyxl"]
    last_err = None
    for eng in engines:
        try:
            buf.seek(0)
            xl = pd.ExcelFile(buf, engine=eng)
            result = {}
            for sheet in xl.sheet_names:
                try:
                    result[sheet] = xl.parse(sheet)
                except Exception as e:
                    result[sheet] = pd.DataFrame({"ERROR": [str(e)]})
            return result
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Cannot read {filename}: {last_err}")


def read_from_path(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return {"Sheet1": pd.read_csv(path)}
    engines = ["openpyxl","xlrd"] if ext == ".xlsx" else ["xlrd","openpyxl"]
    last_err = None
    for eng in engines:
        try:
            xl = pd.ExcelFile(path, engine=eng)
            return {s: xl.parse(s) for s in xl.sheet_names}
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Cannot read {path}: {last_err}")


def col_sig(df):
    return frozenset(c for c in df.columns if c not in TRACKING_COLS)


def group_sheets(file_sheet_dfs):
    buckets = {}
    for fname, sheet, df in file_sheet_dfs:
        buckets.setdefault(col_sig(df), []).append((fname, sheet, df))
    return sorted(buckets.items(), key=lambda x: -len(x[0]))


def to_excel_bytes(sheet_dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheet_dict.items():
            safe = name[:31].translate(str.maketrans(r'\/[]*?:', '_______'))
            df.to_excel(w, sheet_name=safe, index=False)
    return buf.getvalue()


def to_csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8")


# ═══════════════════════════════════════════════════════════════════════════════
# COLUMN MAPPING UI
# ═══════════════════════════════════════════════════════════════════════════════

def render_column_mapping(all_triples, tab_key="upload"):
    """
    Detects column-name mismatches across files and shows a mapping form.
    The user assigns a canonical name to columns that mean the same thing.
    Returns renames_map: {filename: {original_col: canonical_col}}
    Persists in st.session_state per tab_key across re-runs.
    tab_key must be unique per call site to avoid Streamlit DuplicateWidgetID errors.
    """
    ss_key   = f"col_renames_{tab_key}"
    form_key = f"col_mapping_form_{tab_key}"

    file_cols = {}
    # Also keep a sample-values map: {fname: {col: [first 10 non-null values]}}
    file_samples = {}
    for fname, _sname, df in all_triples:
        file_cols.setdefault(fname, set()).update(
            c for c in df.columns if c not in TRACKING_COLS)
        for c in df.columns:
            if c in TRACKING_COLS:
                continue
            try:
                series = df[c]
                if isinstance(series, pd.DataFrame):
                    continue
                vals = series.dropna().astype(str).head(10).tolist()
            except Exception:
                vals = []
            # Only set the first time we see this (file, col) pair
            file_samples.setdefault(fname, {}).setdefault(c, vals)

    fnames = list(file_cols.keys())
    if len(fnames) < 2:
        return st.session_state.get(ss_key, {})

    all_unique = set().union(*file_cols.values())
    common     = set.intersection(*file_cols.values())
    unmatched  = sorted(all_unique - common)

    # ── ALWAYS show a "data sample preview" panel for ALL columns ─────────
    # This is the key UX upgrade — user can verify column meaning without
    # opening files separately.
    with st.expander(
            f"🔍 Preview column contents — {len(all_unique)} unique column(s) across all files",
            expanded=False):
        st.caption(
            "First 10 non-null sample values from each file. "
            "Use this to verify two differently-named columns "
            "(e.g. 'SR No' vs 'Ser Num') actually contain the same kind of data.")
        for col in sorted(all_unique):
            present_in = [f for f in fnames if col in file_cols[f]]
            st.markdown(f"**`{col}`**  — in {len(present_in)} of {len(fnames)} file(s)")
            for f in present_in:
                samples = file_samples.get(f, {}).get(col, [])
                sample_str = ", ".join(samples[:10]) if samples else "_(all null)_"
                st.caption(f"  📄 `{os.path.basename(f)[:35]}`: {sample_str}")
            st.markdown("")

    if not unmatched:
        st.success("✅ All files share identical column names — no mapping needed.")
        st.session_state[ss_key] = {}
        return {}

    st.warning(
        f"**{len(unmatched)} column name(s)** don't appear in every file. "
        "Give the **same Canonical Name** to columns that represent the same data. "
        "Leave a name unchanged to keep that column separate in the output.")

    # Show current applied renames as a reminder
    existing = st.session_state.get(ss_key, {})
    if existing:
        total_applied = sum(len(v) for v in existing.values())
        st.info(f"Currently applied: {total_applied} rename(s) from a previous mapping. "
                "Re-submit the form below to change them.")

    with st.form(form_key):
        st.markdown("**Column Alignment — assign canonical names:**")
        st.caption(
            "💡 Tip: expand the *Preview column contents* panel above to see actual "
            "sample values for each column before deciding which ones should be merged.")
        rows_data = []
        for col in unmatched:
            present_in = [os.path.basename(f)[:28] for f in fnames if col in file_cols[f]]
            absent_in  = [os.path.basename(f)[:28] for f in fnames if col not in file_cols[f]]
            label = (f"`{col}`  ✅ {', '.join(present_in)}"
                     + (f"  ❌ missing: {', '.join(absent_in)}" if absent_in else ""))
            # Pre-fill with previously saved canonical name if available
            prev_canonical = col
            for f in fnames:
                if col in file_cols[f] and col in existing.get(f, {}):
                    prev_canonical = existing[f][col]
                    break
            canonical = st.text_input(label, value=prev_canonical,
                                      key=f"cmap_{tab_key}_{col}")
            # Inline samples for this column from the file(s) that have it
            sample_lines = []
            for f in fnames:
                if col in file_cols[f]:
                    smp = file_samples.get(f, {}).get(col, [])
                    if smp:
                        sample_lines.append(
                            f"  📄 `{os.path.basename(f)[:30]}`: "
                            + ", ".join(smp[:10]))
            if sample_lines:
                st.caption("\n".join(sample_lines))
            rows_data.append((col, [f for f in fnames if col in file_cols[f]], canonical.strip()))

        submitted = st.form_submit_button(
            "Apply Column Mapping", type="primary", use_container_width=True)

    if submitted:
        renames_map = {}
        for col, files_with_col, canonical in rows_data:
            if canonical and canonical != col:
                for f in files_with_col:
                    renames_map.setdefault(f, {})[col] = canonical
        st.session_state[ss_key] = renames_map
        n = sum(len(v) for v in renames_map.values())
        if n:
            st.success(f"Mapping applied: {n} rename(s). Sheets re-grouped below.")
        else:
            st.info("No renames applied — all canonical names match originals.")

    return st.session_state.get(ss_key, {})


def apply_renames_to_triples(triples, renames_map):
    result = []
    for fname, sname, df in triples:
        rn = renames_map.get(fname, {})
        result.append((fname, sname, df.rename(columns=rn) if rn else df))
    return result


# ═══════════════════════════════════════════════════════════════════════════════
# SHARED SETTINGS WIDGET
# ═══════════════════════════════════════════════════════════════════════════════

def render_settings(all_cols, tab_key, mapped_triples=None):
    """
    Render the Merge Settings widget AND a 5-row output preview at the bottom.
    mapped_triples (optional): list of (fname, sname, df) used to generate
    a small preview of what the chosen operation would produce.
    """
    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown("**Operation**  — pick how to combine the files")
        # Group by family with a divider for clarity
        strat_list = list(STRATEGIES.keys())
        chosen = st.selectbox(
            "Merge / Join operation",
            strat_list,
            key=f"strat_{tab_key}",
            help="Union All / Distinct = stack rows vertically.  "
                 "Inner / Left / Right / Full Outer Join = combine rows side-by-side "
                 "on a shared key column.  Cross Join = every-pair cartesian product.")
        cfg = STRATEGIES[chosen]
        st.markdown(f'<span class="badge {cfg["badge"]}">{cfg["badge_lbl"]}</span>',
                    unsafe_allow_html=True)
        st.caption(cfg["head"])
        family_lbl = "🔵 Vertical stacking (Union)" if cfg["family"] == "union" \
                     else "🔗 Horizontal combination (Join)"
        st.caption(f"_{family_lbl}_")
    with col_r:
        key_col, excl_cols = None, []
        if cfg["needs_key"] and all_cols:
            # Use a "— pick a column —" placeholder so we don't auto-run a join
            # with a wrong default key, which can crash on large data
            KEY_PLACEHOLDER = "— pick a key column —"
            key_choice = st.selectbox(
                "Key column  — must exist in every file (use column mapping above to align names)",
                [KEY_PLACEHOLDER] + sorted(all_cols),
                key=f"key_{tab_key}")
            key_col = None if key_choice == KEY_PLACEHOLDER else key_choice
        if cfg["allows_excl"] and all_cols:
            excl_cols = st.multiselect(
                "Columns to IGNORE during duplicate check (optional)",
                sorted(all_cols), key=f"excl_{tab_key}",
                help="e.g. exclude 'Service Request Owner' so engineer-name "
                     "differences don't prevent two rows being called duplicates.")
        clean_types = st.checkbox(
            "Auto-clean data types (dates, city/region, whitespace)",
            value=True, key=f"clean_{tab_key}")
        add_src = st.checkbox("Add 'Source File' column", value=True,
                              key=f"src_{tab_key}")
        out_fmt = st.radio("Download format",
                           ["Excel (.xlsx) — multi-sheet", "CSV (.csv) — first sheet only"],
                           key=f"fmt_{tab_key}", horizontal=True)

    # ── 5-row output preview ───────────────────────────────────────────────
    # Only auto-runs once you've picked an operation + (for joins) an explicit
    # key column. We cap inputs to 10 rows × 3 sheets so even a many-to-many
    # join can't blow up memory on Streamlit Cloud's small container.
    if mapped_triples:
        with st.expander(
                f"👀 Preview — what '{chosen}' will produce (first 5 rows)",
                expanded=True):
            if cfg["needs_key"] and not key_col:
                st.info("👆 Pick a key column above to see a live preview.")
            else:
                try:
                    PREVIEW_ROWS_PER_SHEET = 10   # small to keep cloud memory safe
                    PREVIEW_MAX_SHEETS     = 3
                    sample_dfs = []
                    for _, _, df in mapped_triples[:PREVIEW_MAX_SHEETS]:
                        d = _dedup_columns(df.copy().head(PREVIEW_ROWS_PER_SHEET))
                        sample_dfs.append(d)

                    # Defensive row-count estimate before running the merge —
                    # protects against many-to-many key explosions
                    if cfg["family"] == "join" and key_col:
                        est_rows = 1
                        for d in sample_dfs:
                            if key_col in d.columns:
                                est_rows *= max(1, len(d))
                        if est_rows > 50_000:
                            st.warning(
                                f"⚠️ Preview skipped — even on a 10-row sample, "
                                f"the join on key `{key_col}` would produce "
                                f"≥{est_rows:,} rows (many-to-many explosion). "
                                f"This usually means `{key_col}` isn't a unique "
                                f"identifier in your files. Try a column with "
                                f"unique values per row, or run the full merge "
                                f"to see the real result.")
                            return cfg, key_col, excl_cols, clean_types, add_src, out_fmt

                    preview, _ = cfg["fn"](
                        sample_dfs,
                        key=key_col,
                        excl=set(excl_cols) if excl_cols else None)
                    st.dataframe(preview.head(5), use_container_width=True,
                                 hide_index=True)
                    st.caption(
                        f"Preview built from first {PREVIEW_ROWS_PER_SHEET} rows "
                        f"of up to {PREVIEW_MAX_SHEETS} sheet(s) → "
                        f"{len(preview):,} preview-rows produced.  "
                        f"The actual run uses your full data.")
                    if len(preview) == 0:
                        st.warning(
                            "⚠️ Preview is empty. This usually means the join "
                            "key doesn't match between files. Check that the "
                            "key column has matching values, or revisit the "
                            "Column Alignment step above.")
                except Exception as e:
                    st.warning(f"Preview unavailable: {type(e).__name__}: {e}")

    return cfg, key_col, excl_cols, clean_types, add_src, out_fmt


# ═══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

def render_dashboard():
    """Full EDA dashboard — reads merged data from session state."""
    sheets = st.session_state.get("merged_sheets")
    if not sheets:
        st.info("Run a merge in **Merge & Download** or **Folder Mode** first, "
                "then come back here.")
        if not HAS_PLOTLY:
            st.warning("For interactive charts install plotly:  `pip install plotly`")
        return

    sheet_names = list(sheets.keys())
    sel = (st.selectbox("Sheet to analyse", sheet_names, key="dash_sheet")
           if len(sheet_names) > 1 else sheet_names[0])
    df = _dedup_columns(sheets[sel].copy())

    # ── Summary metrics ───────────────────────────────────────────────────────
    st.subheader("Summary")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Rows",          f"{len(df):,}")
    c2.metric("Columns",       len(df.columns))
    src_n = df["Source File"].nunique() if "Source File" in df.columns else "—"
    c3.metric("Source files",  src_n)
    complete = int(df.dropna().shape[0])
    c4.metric("Complete rows", f"{complete:,}")
    c5.metric("Avg missing",   f"{df.isnull().mean().mean():.1%}")
    st.divider()

    # ── Auto-detect column types ──────────────────────────────────────────────
    def _safe_col(c):
        try:
            s = df[c]
            return None if isinstance(s, pd.DataFrame) else s
        except Exception:
            return None

    date_cols, num_cols, cat_cols, geo_cols = [], [], [], []
    for c in df.columns:
        if c in TRACKING_COLS:
            continue
        s = _safe_col(c)
        if s is None:
            continue
        words = _col_words(c)
        if pd.api.types.is_datetime64_any_dtype(s):
            date_cols.append(c)
        elif pd.api.types.is_numeric_dtype(s):
            num_cols.append(c)
        elif s.dtype == object:
            nu = s.nunique()
            if words & (REGION_KEYWORDS | CITY_KEYWORDS):
                geo_cols.append(c)
                if 1 < nu <= 200:
                    cat_cols.append(c)
            elif 1 < nu <= 80:          # wider than original 40
                cat_cols.append(c)

    if not (date_cols or num_cols or cat_cols or geo_cols or
            ("Source File" in df.columns)):
        st.info(
            "📊 **No charts produced — your data doesn't have enough variability.**  \n"
            "DataMerge Studio looks for: date columns, numeric columns, categorical "
            "columns (≤80 unique values), geographic columns, or a 'Source File' column.  \n"
            "The summary metrics above are still valid.")
        return

    figs_for_export = []   # (title, plotly_fig)

    # ══════════════════════════════════════════════════════════════════════════
    # 1 ▸ NUMERICAL ANALYSIS
    # ══════════════════════════════════════════════════════════════════════════
    if num_cols:
        st.subheader("📊 Numerical Analysis")

        # ── Descriptive statistics table ──────────────────────────────────────
        with st.expander("Descriptive Statistics (mean, std, min, quartiles, max)",
                         expanded=True):
            st.dataframe(df[num_cols].describe().T.round(2),
                         use_container_width=True)

        # ── Histogram with controls ───────────────────────────────────────────
        st.markdown("**Distribution — Histogram**")
        ha, hb, hc, hd = st.columns([2, 1, 1, 1])
        with ha:
            hist_col = st.selectbox("Column", num_cols, key="dash_hist_col")
        with hb:
            nbins = st.slider("Bins", 5, 100, 30, key="dash_hist_bins")
        with hc:
            show_box = st.checkbox("Show box plot", value=True, key="dash_hist_box")
        with hd:
            hist_color = st.selectbox("Color by", ["(none)"] + cat_cols,
                                      key="dash_hist_color")
        if HAS_PLOTLY:
            color_arg  = None if hist_color == "(none)" else hist_color
            marginal   = "box" if show_box else None
            fig = px.histogram(df, x=hist_col, nbins=nbins, color=color_arg,
                               marginal=marginal,
                               title=f"Distribution of {hist_col}", height=380,
                               opacity=0.85)
            fig.update_layout(bargap=0.05)
            st.plotly_chart(fig, use_container_width=True)
            figs_for_export.append((f"Histogram — {hist_col}", fig))
        else:
            st.bar_chart(df[hist_col].value_counts().sort_index())

        # ── Correlation heatmap (≥ 2 numeric cols) ────────────────────────────
        if len(num_cols) >= 2:
            with st.expander("Correlation Heatmap", expanded=False):
                corr = df[num_cols].corr().round(2)
                if HAS_PLOTLY:
                    fig = px.imshow(corr, text_auto=True, aspect="auto",
                                    color_continuous_scale="RdBu_r",
                                    zmin=-1, zmax=1,
                                    title="Pearson Correlation Matrix", height=420)
                    st.plotly_chart(fig, use_container_width=True)
                    figs_for_export.append(("Correlation Heatmap", fig))
                else:
                    st.dataframe(corr, use_container_width=True)

        st.divider()

    # ══════════════════════════════════════════════════════════════════════════
    # 2 ▸ TIME TRENDS
    # ══════════════════════════════════════════════════════════════════════════
    if date_cols:
        st.subheader("📈 Time Trends")
        ta, tb, tc = st.columns([2, 1, 2])
        with ta:
            dcol    = st.selectbox("Date column", date_cols, key="dash_dc")
        with tb:
            agg_lbl = st.selectbox("Group by", ["Day", "Week", "Month"],
                                   key="dash_agg")
        with tc:
            metric_opts = (["Count of records"]
                           + [f"Sum of {c}"  for c in num_cols]
                           + [f"Mean of {c}" for c in num_cols])
            metric_ch   = st.selectbox("Metric", metric_opts, key="dash_metric")

        freq = {"Day": "D", "Week": "W",
                "Month": "ME" if pd.__version__ >= "2.2" else "M"}[agg_lbl]
        try:
            base = df.dropna(subset=[dcol]).set_index(dcol)
            if metric_ch == "Count of records":
                ts   = base.resample(freq).size().rename("Value").reset_index()
                ylab = "Count"
            elif metric_ch.startswith("Sum of "):
                mc   = metric_ch[7:]
                ts   = base[mc].resample(freq).sum().rename("Value").reset_index()
                ylab = f"Sum of {mc}"
            else:
                mc   = metric_ch[8:]
                ts   = base[mc].resample(freq).mean().rename("Value").reset_index()
                ylab = f"Mean of {mc}"

            if HAS_PLOTLY:
                fig = px.line(ts, x=dcol, y="Value", markers=True, height=350,
                              title=f"{ylab} over time  ({agg_lbl})",
                              labels={"Value": ylab})
                fig.update_traces(line_color="#2563eb", marker_size=5)
                fig.update_layout(hovermode="x unified")
                st.plotly_chart(fig, use_container_width=True)
                figs_for_export.append((f"Trend — {ylab}", fig))
            else:
                st.line_chart(ts.set_index(dcol)["Value"])
        except Exception as e:
            st.warning(f"Time trend could not be rendered: {e}")

        st.divider()

    # ══════════════════════════════════════════════════════════════════════════
    # 3 ▸ CATEGORICAL BREAKDOWNS
    # ══════════════════════════════════════════════════════════════════════════
    if cat_cols:
        st.subheader("📋 Categorical Breakdowns")
        ca, cb, cc, cd = st.columns([3, 1, 1, 1])
        with ca:
            sel_cats = st.multiselect(
                "Columns to chart  (select any / all)",
                cat_cols,
                default=cat_cols[:min(6, len(cat_cols))],
                key="dash_cats")
        with cb:
            top_n = st.slider("Top N values", 5, 30, 15, key="dash_topn")
        with cc:
            chart_type = st.radio("Chart type", ["Bar", "Pie"],
                                  key="dash_cat_type", horizontal=True)
        with cd:
            sort_by = st.radio("Sort", ["Count ↓", "A–Z"],
                               key="dash_cat_sort", horizontal=True)

        if sel_cats:
            grid_cols = 2
            rows = (len(sel_cats) + grid_cols - 1) // grid_cols
            grid = [st.columns(grid_cols) for _ in range(rows)]
            flat = [cell for row in grid for cell in row]
            for i, col in enumerate(sel_cats):
                with flat[i]:
                    vc = (df[col].value_counts().head(top_n)
                            .rename_axis(col).reset_index(name="Count"))
                    if sort_by == "A–Z":
                        vc = vc.sort_values(col)
                    if HAS_PLOTLY:
                        if chart_type == "Bar":
                            fig = px.bar(vc, x="Count", y=col, orientation="h",
                                         title=col,
                                         height=max(300, len(vc) * 26 + 80),
                                         color="Count",
                                         color_continuous_scale="Teal")
                            fig.update_layout(
                                yaxis={"categoryorder": "total ascending"},
                                showlegend=False,
                                coloraxis_showscale=False,
                                title_font_size=13,
                                margin=dict(l=10, r=10, t=40, b=10))
                        else:
                            fig = px.pie(vc, names=col, values="Count",
                                         title=col, height=350,
                                         hole=0.35)
                            fig.update_traces(textposition="inside",
                                              textinfo="percent+label")
                            fig.update_layout(title_font_size=13,
                                              showlegend=True)
                        st.plotly_chart(fig, use_container_width=True)
                        figs_for_export.append((f"Category — {col}", fig))
                    else:
                        st.write(f"**{col}**")
                        st.bar_chart(vc.set_index(col))
        else:
            st.info("Select at least one column above to see charts.")

        st.divider()

    # ══════════════════════════════════════════════════════════════════════════
    # 4 ▸ GEOGRAPHIC / REGION DISTRIBUTION
    # ══════════════════════════════════════════════════════════════════════════
    geo_candidates = geo_cols if geo_cols else [
        c for c in cat_cols
        if any(kw in c.lower() for kw in
               {"city","town","state","region","zone","area","territory",
                "country","district","branch","location","place"})]

    if geo_candidates:
        st.subheader("🗺️ Geographic / Region Distribution")
        ga, gb, gc = st.columns([2, 1, 1])
        with ga:
            map_col = st.selectbox("Region / location column",
                                   geo_candidates, key="dash_mapcol")
        with gb:
            map_topn = st.slider("Top N regions", 5, 60, 25, key="dash_mapn")
        with gc:
            map_style = st.radio("Chart style", ["Horizontal Bar", "Choropleth Map"],
                                 key="dash_mapstyle", horizontal=False)

        rc = (df[map_col].value_counts().head(map_topn)
                .rename_axis("Region").reset_index(name="Count"))

        if HAS_PLOTLY:
            rendered_map = False
            if map_style == "Choropleth Map":
                try:
                    fig = px.choropleth(
                        rc, locations="Region",
                        locationmode="country names",
                        color="Count", hover_name="Region",
                        color_continuous_scale="Blues",
                        title=f"Geographic Distribution — {map_col}",
                        height=460)
                    fig.update_layout(geo=dict(showframe=False,
                                               showcoastlines=True))
                    st.plotly_chart(fig, use_container_width=True)
                    figs_for_export.append((f"Map — {map_col}", fig))
                    rendered_map = True
                except Exception:
                    st.caption("⚠️ Choropleth requires standard country names or "
                               "ISO codes. Falling back to bar chart.")

            if not rendered_map:
                fig = px.bar(
                    rc.sort_values("Count"),
                    x="Count", y="Region", orientation="h",
                    title=f"Region Distribution — {map_col}",
                    height=max(350, len(rc) * 24 + 80),
                    color="Count", color_continuous_scale="Blues")
                fig.update_layout(
                    yaxis={"categoryorder": "total ascending"},
                    coloraxis_showscale=False,
                    margin=dict(l=10, r=10, t=40, b=10))
                st.plotly_chart(fig, use_container_width=True)
                figs_for_export.append((f"Region — {map_col}", fig))
                if map_style == "Choropleth Map":
                    st.caption(
                        "💡 For a true world map, ensure your column contains "
                        "standard country names (e.g. 'India', 'United States').")
        else:
            st.bar_chart(rc.set_index("Region"))

        st.divider()

    # ══════════════════════════════════════════════════════════════════════════
    # 5 ▸ RECORDS BY SOURCE FILE  (keep as-is)
    # ══════════════════════════════════════════════════════════════════════════
    if "Source File" in df.columns:
        st.subheader("Records by Source File")
        src = (df["Source File"].value_counts()
                 .rename_axis("Source File").reset_index(name="Count"))
        if HAS_PLOTLY:
            fig = px.bar(src, x="Source File", y="Count", color="Count",
                         color_continuous_scale="Blues", height=300)
            fig.update_layout(showlegend=False, coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)
            figs_for_export.append(("Records by Source File", fig))
        else:
            st.bar_chart(src.set_index("Source File"))
        st.divider()

    # ══════════════════════════════════════════════════════════════════════════
    # 6 ▸ MISSING VALUES %  (keep as-is)
    # ══════════════════════════════════════════════════════════════════════════
    st.subheader("Missing Values % by Column")
    null_s = (df.isnull().sum() / len(df) * 100).round(1).sort_values(ascending=False)
    null_s = null_s[null_s > 0]
    null_s.name = "Missing %"
    if null_s.empty:
        st.success("No missing values — perfect data quality!")
    elif HAS_PLOTLY:
        fig = px.bar(null_s.reset_index(), x="index", y="Missing %",
                     labels={"index": "Column"},
                     color="Missing %", color_continuous_scale="Reds", height=300)
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)
        figs_for_export.append(("Missing Values %", fig))
    else:
        st.bar_chart(null_s)
    st.divider()

    # ══════════════════════════════════════════════════════════════════════════
    # 7 ▸ DATA EXPLORER  (keep as-is)
    # ══════════════════════════════════════════════════════════════════════════
    st.subheader("Data Explorer")
    with st.expander("Filter & browse merged data", expanded=False):
        fc1, fc2 = st.columns(2)
        with fc1:
            fcol = st.selectbox("Filter by column",
                                ["(none)"] + list(df.columns), key="dash_fcol")
        view_df = df
        if fcol != "(none)":
            with fc2:
                if df[fcol].dtype == object:
                    fvals = st.multiselect(
                        "Include values",
                        sorted(df[fcol].dropna().unique().tolist()),
                        key="dash_fvals")
                    if fvals:
                        view_df = df[df[fcol].isin(fvals)]
                elif pd.api.types.is_numeric_dtype(df[fcol]):
                    mn = float(df[fcol].min())
                    mx = float(df[fcol].max())
                    rng = st.slider("Range", mn, mx, (mn, mx), key="dash_rng")
                    view_df = df[df[fcol].between(*rng)]
        st.caption(f"Showing {len(view_df):,} of {len(df):,} rows")
        st.dataframe(view_df, use_container_width=True, height=400, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════════
    # 8 ▸ EXPORT STANDALONE HTML
    # ══════════════════════════════════════════════════════════════════════════
    st.subheader("Export Dashboard")
    if not HAS_PLOTLY:
        st.info("Install plotly to enable HTML export:  `pip install plotly`")
    elif figs_for_export:
        if st.button("Generate standalone HTML dashboard (works offline)",
                     type="secondary", use_container_width=True):
            parts = [
                "<html><head><meta charset='utf-8'>",
                f"<title>EDA Dashboard — {sel}</title>",
                "<style>body{font-family:sans-serif;margin:32px;background:#f9fafb;}"
                "h1{color:#1e3a5f;}h2{color:#374151;margin-top:32px;border-top:"
                "1px solid #e2e8f0;padding-top:16px;}"
                "p{color:#6b7280;}</style></head><body>",
                "<h1>DataMerge Studio — EDA Dashboard</h1>",
                f"<p><b>Sheet:</b> {sel} &nbsp;|&nbsp; "
                f"<b>Rows:</b> {len(df):,} &nbsp;|&nbsp; "
                f"<b>Columns:</b> {len(df.columns)}</p>",
            ]
            first = True
            for title, fig in figs_for_export:
                parts.append(f"<h2>{title}</h2>")
                parts.append(fig.to_html(
                    full_html=False,
                    include_plotlyjs="cdn" if first else False))
                first = False
            parts.append("</body></html>")
            st.download_button(
                "⬇️ Download dashboard.html",
                data="\n".join(parts).encode("utf-8"),
                file_name="dashboard.html",
                mime="text/html",
                use_container_width=True,
                type="primary")
            st.caption("Fully self-contained HTML — no internet or Streamlit needed to view it.")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE
# ═══════════════════════════════════════════════════════════════════════════════
st.title("🔀 DataMerge Studio")
st.markdown(
    "Combine **CSV or Excel** files — including **multi-sheet** workbooks — "
    "using all standard SQL set operations: **Union All**, **Inner / Left / Right / "
    "Full Outer / Cross Join**. Includes column mapping, data-type cleaning, "
    "duplicate audit, and an interactive dashboard.")
st.divider()

# Folder Mode tab is only shown when running locally (requires filesystem access)
_tab_names = ["📚 Learn Merge & Joins", "📁 Upload Files",
              "🔀 Merge & Download", "📊 Dashboard"]
if HAS_TKINTER:
    _tab_names.insert(3, "📂 Folder Mode")

_tabs      = st.tabs(_tab_names)
tab_learn  = _tabs[0]
tab_upload = _tabs[1]
tab_merge  = _tabs[2]
if HAS_TKINTER:
    tab_folder = _tabs[3]
    tab_dash   = _tabs[4]
else:
    tab_folder = None   # never rendered on cloud
    tab_dash   = _tabs[3]


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — LEARN
# ══════════════════════════════════════════════════════════════════════════════
with tab_learn:
    st.subheader("Merge vs. Join — what's the difference?")
    st.markdown(
        "**Merge (Union)** stacks rows **vertically** — combines files into a "
        "longer list. Use when files have the same kind of records.  \n"
        "**Join** combines rows **horizontally** by matching on a shared key column — "
        "puts related data side-by-side. Use when each file has different details "
        "about the same entities.")
    st.divider()

    st.markdown("### 🔵 Union family — vertical stacking")
    st.markdown(
        "Use the **Orders** demo data below. Both files have the same columns.")
    fau, fbu = get_dummy_union()
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**File A** — Orders (Jan)")
        st.dataframe(fau, use_container_width=True, hide_index=True)
    with c2:
        st.markdown("**File B** — Orders (Feb)")
        st.dataframe(fbu, use_container_width=True, hide_index=True)
    st.info(
        "• **O003** — identical row appears in both files (true duplicate)  \n"
        "• **O001, O002** unique to File A   •   **O004, O005** unique to File B")
    st.divider()

    st.markdown("### 🔗 Join family — horizontal combination")
    st.markdown(
        "Use the **Employees** demo data below. Both files share `Employee ID`. "
        "File A has names/departments; File B has salaries/managers — joins put them side-by-side.")
    fa, fb = get_dummy_join()
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**File A** — Employees (5 rows)")
        st.dataframe(fa, use_container_width=True, hide_index=True)
    with c2:
        st.markdown("**File B** — Payroll (5 rows)")
        st.dataframe(fb, use_container_width=True, hide_index=True)
    st.info(
        "• Keys **E003, E004, E005** appear in both files (matches)  \n"
        "• **E001, E002** only in File A   •   **E006, E007** only in File B")
    st.divider()

    st.markdown("### Detailed walk-through — click any operation to expand")

    for name, cfg in STRATEGIES.items():
        with st.expander(f"{cfg['icon']}  **{name}**", expanded=False):
            st.markdown(f'<span class="badge {cfg["badge"]}">{cfg["badge_lbl"]}</span>',
                        unsafe_allow_html=True)
            st.markdown(f"**{cfg['head']}**\n\n{cfg['detail']}")
            st.markdown(f"*Best for: {cfg['best']}*")

            # Pick the right demo data + key per family
            if cfg["family"] == "union":
                dfa, dfb, key = fau.copy(), fbu.copy(), None
                excl_demo = None
            else:  # join family
                dfa, dfb, key = fa.copy(), fb.copy(), "Employee ID"
                excl_demo = None

            try:
                if name == "Cross Join":
                    # Small subset to avoid 25-row display
                    raw, _ = cfg["fn"]([dfa.head(2), dfb.head(2)])
                    cap = "Showing first 2 rows of each → 2×2 = 4 output rows for clarity."
                else:
                    raw, _ = cfg["fn"]([dfa, dfb], key=key, excl=excl_demo)
                    cap = ""

                col_v, col_t = st.columns([1, 1.4])
                with col_v:
                    fig = make_venn_fig(name)
                    if fig:
                        st.pyplot(fig, use_container_width=True)
                        plt.close(fig)
                with col_t:
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Input (A+B)", len(dfa) + len(dfb))
                    m2.metric("Output rows", len(raw))
                    m3.metric("Output cols", len(raw.columns))
                    if cap:
                        st.caption(cap)
                    st.dataframe(raw, use_container_width=True, hide_index=True)
            except Exception as e:
                st.warning(f"Example error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with tab_upload:
    st.subheader("Upload Your Files")
    st.markdown(
        "Upload **2 or more** CSV or Excel files. "
        "**Multi-sheet workbooks are fully supported** — all sheets are read.")

    # ── New Session button — clears all session state with confirmation ─────
    _has_run = bool(st.session_state.get("_last_run"))
    _hdr_l, _hdr_r = st.columns([4, 1])
    with _hdr_r:
        if _has_run:
            if st.session_state.get("_confirm_new_session"):
                st.warning("Sure? This clears all current results.")
                _yes, _no = st.columns(2)
                if _yes.button("✅ Yes, reset", key="confirm_yes",
                               use_container_width=True):
                    # Clear everything related to a run
                    for _k in ("uf", "_last_run", "_upload_dl", "_folder_dl",
                               "merged_sheets", "_folder_read_key",
                               "_folder_triples", "_folder_read_errors",
                               "_confirm_new_session"):
                        st.session_state.pop(_k, None)
                    # Also clear col mapping & widget state
                    for _k in list(st.session_state.keys()):
                        if _k.startswith(("col_renames_", "cmap_", "strat_",
                                          "key_", "excl_", "clean_", "src_",
                                          "fmt_", "fc_")):
                            st.session_state.pop(_k, None)
                    # Rotate the uploader key — forces Streamlit to render a
                    # completely fresh file_uploader widget (the only reliable
                    # way to clear file chips from the browser UI)
                    st.session_state["_uploader_key"] = (
                        st.session_state.get("_uploader_key", 0) + 1)
                    st.rerun()
                if _no.button("Cancel", key="confirm_no",
                              use_container_width=True):
                    st.session_state.pop("_confirm_new_session", None)
                    st.rerun()
            else:
                if st.button("🔄 New Session", key="new_session_btn",
                             use_container_width=True,
                             help="Clear all uploaded files, mappings, and results."):
                    st.session_state["_confirm_new_session"] = True
                    st.rerun()
    _uploader_key = f"uploader_{st.session_state.get('_uploader_key', 0)}"
    uploaded = st.file_uploader(
        "Drag & drop files here, or click Browse",
        type=["csv", "xlsx", "xls"], accept_multiple_files=True,
        key=_uploader_key)
    if uploaded:
        st.session_state["uf"] = uploaded
        # Warn if new files differ from the last merge — user might lose output
        _new_fnames = sorted(u.name for u in uploaded)
        _old_fnames = st.session_state.get("_last_run", {}).get("file_names", [])
        if _old_fnames and _new_fnames != _old_fnames:
            st.warning(
                "⚠️ You have results from a previous merge session. "
                "Running a new merge will replace them. "
                "**Download your previous output first** if you haven't already — "
                "go to the **Merge & Download** tab.")
        st.success(f"{len(uploaded)} file(s) ready. Go to **Merge & Download**.")
        for i, uf in enumerate(uploaded):
            with st.expander(f"File {i+1}: {uf.name}", expanded=True):
                try:
                    data = uf.read(); uf.seek(0)
                    sheets = read_all_sheets_cached(data, uf.name)
                    st.markdown(f"**{len(sheets)} sheet(s) found:**")
                    for sname, df in sheets.items():
                        c1, c2, c3 = st.columns(3)
                        c1.metric(f"'{sname}' rows", f"{len(df):,}")
                        c2.metric("Columns", len(df.columns))
                        c3.metric("File size", f"{uf.size/1024:.1f} KB")
                        st.dataframe(df.head(5), use_container_width=True, hide_index=True)
                        if len(df) > 5:
                            st.caption(f"Showing 5 of {len(df):,} rows.")
                        st.divider()
                except Exception as e:
                    st.error(f"Could not read: {e}")
    else:
        st.info("No files uploaded yet.")
        st.session_state.pop("uf", None)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — MERGE & DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════════
with tab_merge:
    st.subheader("Configure and Run Merge")
    files = st.session_state.get("uf", [])
    if len(files) < 2:
        st.warning("Upload at least 2 files in the **Upload Files** tab first.")
    else:
        # Collect raw triples
        all_triples = []
        for uf in files:
            try:
                data = uf.read(); uf.seek(0)
                sheets = read_all_sheets_cached(data, uf.name)
                for sname, df in sheets.items():
                    all_triples.append((uf.name, sname, df))
            except Exception as e:
                st.error(f"Error reading {uf.name}: {e}")

        # ── Step 1: Column mapping ─────────────────────────────────────────
        st.markdown("### Step 1 — Column Alignment")
        renames_map    = render_column_mapping(all_triples, tab_key="upload")
        mapped_triples = apply_renames_to_triples(all_triples, renames_map)

        st.divider()

        # ── Step 2: Settings (with live preview) ──────────────────────────
        st.markdown("### Step 2 — Merge Settings")
        all_cols_flat = sorted(
            set(c for _, _, df in mapped_triples for c in df.columns) - TRACKING_COLS)
        cfg, key_col, excl_cols, clean_types, add_src, out_fmt = \
            render_settings(all_cols_flat, "upload", mapped_triples=mapped_triples)

        st.divider()

        # ── Step 3: Sheet groups ──────────────────────────────────────────
        # All standard ops use merge_all_groups=True → every sheet from every
        # file participates in ONE combined operation (chained join, or stack).
        groups = group_sheets(mapped_triples)
        if cfg.get("merge_all_groups") and len(groups) > 1:
            _all_entries = [e for _, entries in groups for e in entries]
            _union_cols  = frozenset(c for _, _, df in _all_entries
                                     for c in df.columns if c not in TRACKING_COLS)
            groups     = [(_union_cols, _all_entries)]
            group_label = (f"### Step 3 — Sheet Groups  "
                           f"(1 group — all {len(_all_entries)} sheet(s) participate "
                           f"in one combined output)")
        else:
            group_label = f"### Step 3 — Sheet Groups  ({len(groups)} detected)"

        st.markdown(group_label)
        for i, (cols_key, entries) in enumerate(groups):
            file_names  = sorted({e[0] for e in entries})
            sheet_names = sorted({e[1] for e in entries})
            total_rows  = sum(len(e[2]) for e in entries)
            bg     = "#f0fdf4" if len(entries) > 1 else "#fefce8"
            action = (f"MERGE {len(entries)} sheets"
                      if len(entries) > 1 else "1 sheet (no merge needed)")
            st.markdown(
                f'<div class="group-box" style="background:{bg}">'
                f'<b>Group {i+1}</b> — {len(cols_key)} columns — '
                f'{total_rows:,} rows — {action}<br>'
                f'<small>Sheet names: {", ".join(sheet_names)}<br>'
                f'Files: {", ".join(f[:40] for f in file_names)}</small></div>',
                unsafe_allow_html=True)

        st.divider()

        if st.button("Run Merge", type="primary", use_container_width=True):
            if cfg["needs_key"] and not key_col:
                st.error("Please select a key column.")
            else:
                output_sheets    = {}
                all_audits       = []
                all_type_reports = []
                total_in = total_out = 0

                with st.spinner("Merging..."):
                    for i, (cols_key, entries) in enumerate(groups):
                        out_name = " + ".join(sorted({e[1] for e in entries}))[:31]
                        dfs = []
                        for fname, sname, df in entries:
                            d = _dedup_columns(df.copy())
                            if add_src:
                                if "Source File" not in d.columns:
                                    d.insert(0, "Source File", fname)
                            if clean_types:
                                d, trpt = clean_dtypes(d)
                                all_type_reports.extend(trpt)
                            dfs.append(d)

                        try:
                            result, audit = cfg["fn"](
                                dfs, key=key_col,
                                excl=set(excl_cols) if excl_cols else None)
                        except Exception as e:
                            st.warning(f"Group {i+1} merge failed: {e}")
                            result = pd.concat(dfs, ignore_index=True)
                            audit  = pd.DataFrame()

                        n_in       = sum(len(d) for d in dfs)
                        total_in  += n_in
                        total_out += len(result)
                        output_sheets[out_name] = result
                        all_audits.append(audit)

                # Build download bytes
                if out_fmt.startswith("Excel"):
                    dl_data = to_excel_bytes(output_sheets)
                    dl_name = "merged_output.xlsx"
                    dl_mime = ("application/vnd.openxmlformats-"
                               "officedocument.spreadsheetml.sheet")
                else:
                    first_df = next(iter(output_sheets.values()))
                    dl_data  = to_csv_bytes(first_df)
                    dl_name  = "merged_output.csv"
                    dl_mime  = "text/csv"

                # Store everything — rendered persistently below, outside this block
                st.session_state["merged_sheets"] = output_sheets
                st.session_state["_upload_dl"]    = {
                    "data": dl_data, "name": dl_name, "mime": dl_mime}
                st.session_state["_last_run"]     = {
                    "file_names":    sorted(f.name for f in files),
                    "metrics":       (len(output_sheets), total_in, total_out),
                    "type_reports":  sorted(set(all_type_reports)),
                    "audits":        all_audits,
                    "output_sheets": output_sheets,
                    "operation":     list(STRATEGIES.keys())[
                        list(STRATEGIES.values()).index(cfg)],
                    "key_col":       key_col,
                }

                # Zero-row warning — most commonly a join key mismatch
                if total_out == 0 and cfg["family"] == "join":
                    st.error(
                        f"⚠️ Your **{list(STRATEGIES.keys())[list(STRATEGIES.values()).index(cfg)]}** "
                        f"produced **0 rows** — this almost always means the key column "
                        f"'`{key_col}`' has **no matching values across files**.  \n"
                        f"Check that:\n"
                        f"- The values in '{key_col}' actually overlap between files "
                        f"(e.g. `E001` vs `e001` won't match — case matters)\n"
                        f"- You haven't mapped two genuinely different columns to the same name\n"
                        f"- Try `Full Outer Join` first to see all rows side by side")

    # ── Persistent results — shown until user explicitly runs a new merge ─
    # Lives OUTSIDE the if/else so it survives every widget interaction.
    _lr = st.session_state.get("_last_run")
    if _lr:
        _cur_fnames = sorted(f.name for f in files)
        if _cur_fnames != _lr["file_names"]:
            st.warning(
                "⚠️ The files loaded above have changed since this merge was run. "
                "Results below are from your **previous session**. "
                "Download now if you still need this output, then run Merge again "
                "to process the new files.")

        _gn, _ti, _to = _lr["metrics"]
        st.success("✅ Merge complete!")
        _m1, _m2, _m3, _m4 = st.columns(4)
        _m1.metric("Groups merged",     _gn)
        _m2.metric("Total input rows",  f"{_ti:,}")
        _m3.metric("Total output rows", f"{_to:,}")
        _m4.metric("Rows removed",      f"{_ti - _to:,}")

        if _lr["type_reports"]:
            with st.expander(
                    f"Data Type Cleaning — "
                    f"{len(_lr['type_reports'])} change(s) applied"):
                for _line in _lr["type_reports"]:
                    st.markdown(f"- {_line}")

        show_audit(_lr["audits"])

        for _sname, _df in _lr["output_sheets"].items():
            with st.expander(f"Output: '{_sname}'  ({len(_df):,} rows)",
                             expanded=True):
                st.dataframe(_df.head(100), use_container_width=True,
                             hide_index=True)
                if len(_df) > 100:
                    st.caption(f"Preview: first 100 of {len(_df):,} rows.")

        if "_upload_dl" in st.session_state:
            _dl = st.session_state["_upload_dl"]
            st.download_button(
                f"⬇️ Download {_dl['name']}",
                data=_dl["data"], file_name=_dl["name"], mime=_dl["mime"],
                use_container_width=True, type="primary",
                key="upload_dl_persistent")

        st.info("Go to the **📊 Dashboard** tab to explore and export charts.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — FOLDER MODE
# ══════════════════════════════════════════════════════════════════════════════
if HAS_TKINTER:
    with tab_folder:
        st.subheader("Folder Mode")
        st.markdown(
            "Point to a folder — the app reads all CSV/Excel files automatically, "
            "groups their sheets by column structure, and merges matching groups. "
            "Supports **incremental append** when new files arrive.")

        st.info(
            "🖥️ **Local-only feature** — Folder Mode reads files directly from your "
            "computer's file system, so the app must be running locally (via "
            "`run_app.bat`). It cannot access your PC's folders when hosted on the cloud.  \n"
            "📁 **Upload Files** and **Merge & Download** tabs work everywhere — "
            "locally and on the cloud.", icon="ℹ️")

        # ── Browse button (only shown when tkinter is available) ──────────────────
        if HAS_TKINTER:
            if st.button("📁 Browse for folder…", key="folder_browse_btn"):
                _root = tk.Tk()
                _root.withdraw()
                _root.wm_attributes("-topmost", 1)
                _picked = _tkfd.askdirectory(title="Select folder containing your files")
                _root.destroy()
                if _picked:
                    st.session_state["_folder_path_input"] = os.path.normpath(_picked)
                st.rerun()
        else:
            st.caption(
                "💡 Folder browser dialog not available in this Python environment — "
                "just paste or type your folder path directly in the box below. "
                "All merge features work normally without it.")

        f_col, o_col = st.columns([3, 2])
        with f_col:
            _default_path = st.session_state.get("_folder_path_input", "")
            folder_path = st.text_input(
                "Folder path" + (" (or use Browse button above)" if HAS_TKINTER else ""),
                value=_default_path,
                key="_folder_path_text")
        with o_col:
            output_name = st.text_input(
                "Output filename",
                value="MERGED_output.xlsx",
                help="💡 Use .csv extension for a smaller file — CSV is typically "
                     "3–5× smaller than Excel because it stores only raw data "
                     "with no formatting, styles, or workbook metadata.")
            st.caption("📦 Tip: rename to `MERGED_output.csv` to reduce file size "
                       "(select CSV in Download Format below too).")

        # Sync typed value back to session state so Browse and typing both work
        st.session_state["_folder_path_input"] = folder_path

        # Strip whitespace and surrounding quotes (common when copy-pasting from Explorer)
        folder_path = folder_path.strip().strip('"').strip("'").strip()
        output_name = output_name.strip().strip('"').strip("'").strip()

        # Try to list the folder
        _folder_ok = False
        _folder_list_err = None
        if folder_path:
            try:
                os.listdir(folder_path)
                _folder_ok = True
            except Exception as _e:
                _folder_list_err = str(_e)

        if not folder_path:
            st.info("Enter a folder path above, or click Browse to pick one.")
        elif not _folder_ok:
            st.error(
                f"Folder not accessible. Path the app is checking:\n\n"
                f"`{folder_path}`\n\n"
                + (f"OS error: `{_folder_list_err}`\n\n" if _folder_list_err else "")
                + "Try using the **Browse** button instead of typing the path."
            )
        else:
            output_path    = os.path.join(folder_path, output_name)
            # Also check the alternate-extension variant so switching
            # between Excel and CSV doesn't lose the existing-output detection.
            _base_name     = os.path.splitext(output_name)[0]
            _alt_name      = (_base_name + ".csv"  if output_name.lower().endswith(".xlsx")
                              else _base_name + ".xlsx")
            _output_names  = {output_name.lower(), _alt_name.lower()}
            all_data_files = sorted([
                f for f in os.listdir(folder_path)
                if f.lower().endswith((".csv", ".xlsx", ".xls"))
                and f.lower() not in _output_names   # exclude both format variants
            ])

            _existing_path = (output_path if os.path.exists(output_path)
                              else os.path.join(folder_path, _alt_name))
            st.caption(f"Looking for existing output at: `{output_path}`  "
                       f"{'✅ found' if os.path.exists(output_path) else '— not found yet'}")

            if not all_data_files:
                st.warning("No CSV or Excel files found in this folder.")
            else:
                # Detect already-processed files from existing output
                already_done       = set()
                existing_row_count = 0
                has_existing       = (os.path.exists(output_path) or
                                      os.path.exists(_existing_path))
                ex_sheets          = {}
                if has_existing:
                    _read_path = (output_path if os.path.exists(output_path)
                                  else _existing_path)
                    try:
                        ex_sheets = read_from_path(_read_path)
                        for df in ex_sheets.values():
                            if "Source File" in df.columns:
                                already_done.update(df["Source File"].dropna().unique())
                        existing_row_count = sum(len(d) for d in ex_sheets.values())
                    except Exception as e:
                        st.warning(f"Could not read existing output: {e}")

                new_files = [f for f in all_data_files if f not in already_done]

                s1, s2, s3 = st.columns(3)
                s1.metric("Files in folder",          len(all_data_files))
                s2.metric("Already in merged output", len(all_data_files) - len(new_files))
                s3.metric("New files detected",       len(new_files))

                if has_existing and already_done:
                    st.info(f"Existing **{output_name}** found "
                            f"({existing_row_count:,} rows across {len(ex_sheets)} sheet(s)). "
                            f"{len(new_files)} new file(s) detected.")

                st.divider()

                if has_existing and already_done and new_files:
                    mode = st.radio("Mode",
                                    ["Append only new files", "Re-merge ALL files from scratch"],
                                    horizontal=True)
                elif has_existing and already_done and not new_files:
                    st.success("Merged file is up to date — no new files detected.")
                    mode = "Re-merge ALL files from scratch"
                    st.caption("You can still re-merge all files to refresh.")
                else:
                    mode = "Re-merge ALL files from scratch"

                files_to_show = new_files if mode.startswith("Append") else all_data_files
                if not files_to_show:
                    st.info("No files to process under the current mode.")
                else:
                    st.markdown(f"**Files to process ({len(files_to_show)}) — uncheck to skip:**")
                    selected = []
                    cols_row = st.columns(2)
                    for i, f in enumerate(files_to_show):
                        with cols_row[i % 2]:
                            if st.checkbox(f, value=True, key=f"fc_{f}"):
                                selected.append(f)

                    if not selected:
                        st.warning("No files selected.")
                    else:
                        st.divider()

                        # ── Read all selected files upfront (needed for column mapping) ──
                        read_key = (folder_path, tuple(sorted(selected)))
                        if st.session_state.get("_folder_read_key") != read_key:
                            raw_triples, read_errors = [], []
                            prog = st.progress(0, text="Scanning files...")
                            for i, fname in enumerate(selected):
                                try:
                                    fsheets = read_from_path(
                                        os.path.join(folder_path, fname))
                                    for sname, df in fsheets.items():
                                        raw_triples.append((fname, sname, df))
                                except Exception as e:
                                    read_errors.append(f"{fname}: {e}")
                                prog.progress((i + 1) / len(selected),
                                              text=f"Scanning: {fname}")
                            prog.empty()
                            st.session_state["_folder_read_key"]    = read_key
                            st.session_state["_folder_triples"]     = raw_triples
                            st.session_state["_folder_read_errors"] = read_errors

                        raw_triples = st.session_state.get("_folder_triples", [])
                        for err in st.session_state.get("_folder_read_errors", []):
                            st.warning(f"Could not read: {err}")

                        if not raw_triples:
                            st.error("No sheets could be read from the selected files.")
                        else:
                            # ── Step 1: Column alignment ───────────────────────────
                            st.markdown("### Step 1 — Column Alignment")
                            folder_renames = render_column_mapping(raw_triples, tab_key="folder")
                            mapped_triples = apply_renames_to_triples(
                                raw_triples, folder_renames)

                            st.divider()

                            # ── Step 2: Merge settings ─────────────────────────────
                            st.markdown("### Step 2 — Merge Settings")
                            all_folder_cols = sorted(
                                set().union(*[set(e[2].columns)
                                              for e in mapped_triples]) - TRACKING_COLS)
                            cfg, key_col, excl_cols, clean_types, add_src, out_fmt = \
                                render_settings(all_folder_cols, "folder",
                                                mapped_triples=mapped_triples)
                            st.divider()

                            if st.button("Run Folder Merge", type="primary",
                                         use_container_width=True):
                                if cfg["needs_key"] and not key_col:
                                    st.error("Please select a key column.")
                                else:
                                    # Apply source columns + type cleaning
                                    all_triples = []
                                    all_type_reports = []
                                    for fname, sname, df in mapped_triples:
                                        d = _dedup_columns(df.copy())
                                        if add_src:
                                            if "Source File" not in d.columns:
                                                d.insert(0, "Source File", fname)
                                        if clean_types:
                                            d, trpt = clean_dtypes(d)
                                            all_type_reports.extend(trpt)
                                        all_triples.append((fname, sname, d))

                                    with st.spinner("Merging..."):
                                        groups = group_sheets(all_triples)
                                        if cfg.get("merge_all_groups") and len(groups) > 1:
                                            _ae = [e for _, entries in groups for e in entries]
                                            _uc = frozenset(c for _, _, df in _ae
                                                            for c in df.columns
                                                            if c not in TRACKING_COLS)
                                            groups = [(_uc, _ae)]
                                        # Simple Append: collapse all groups → one output sheet
                                        if cfg.get("merge_all_groups") and len(groups) > 1:
                                            _ae = [e for _, entries in groups for e in entries]
                                            _uc = frozenset(
                                                c for _, _, df in _ae
                                                for c in df.columns if c not in TRACKING_COLS)
                                            groups = [(_uc, _ae)]
                                        new_output_sheets = {}
                                        all_audits        = []
                                        total_in = total_out = 0

                                        for cols_key, entries in groups:
                                            out_name = " + ".join(
                                                sorted({e[1] for e in entries}))[:31]
                                            dfs = [e[2] for e in entries]
                                            try:
                                                result, audit = cfg["fn"](
                                                    dfs, key=key_col,
                                                    excl=set(excl_cols) if excl_cols else None)
                                            except Exception as e:
                                                st.warning(f"Merge failed for '{out_name}': {e}")
                                                result = pd.concat(dfs, ignore_index=True)
                                                audit  = pd.DataFrame()
                                            n_in       = sum(len(d) for d in dfs)
                                            total_in  += n_in
                                            total_out += len(result)
                                            new_output_sheets[out_name] = result
                                            all_audits.append(audit)

                                        if mode.startswith("Append") and has_existing:
                                            final_sheets = dict(ex_sheets)
                                            for sname, new_df in new_output_sheets.items():
                                                if sname in final_sheets:
                                                    final_sheets[sname] = pd.concat(
                                                        [final_sheets[sname], new_df],
                                                        ignore_index=True)
                                                else:
                                                    final_sheets[sname] = new_df
                                        else:
                                            final_sheets = new_output_sheets

                                    st.session_state["merged_sheets"] = final_sheets

                                    n_total = sum(len(d) for d in final_sheets.values())
                                    st.success(f"Done! {len(final_sheets)} sheet(s), "
                                               f"{n_total:,} total rows.")

                                    m1, m2, m3, m4 = st.columns(4)
                                    m1.metric("Groups",       len(final_sheets))
                                    m2.metric("Input rows",   f"{total_in:,}")
                                    m3.metric("Output rows",  f"{total_out:,}")
                                    m4.metric("Rows removed", f"{total_in - total_out:,}")

                                    if all_type_reports:
                                        with st.expander(
                                                f"Data Type Cleaning — "
                                                f"{len(set(all_type_reports))} change(s)"):
                                            for line in sorted(set(all_type_reports)):
                                                st.markdown(f"- {line}")

                                    show_audit(all_audits)

                                    for sname, df in final_sheets.items():
                                        with st.expander(
                                                f"Sheet: '{sname}'  ({len(df):,} rows)",
                                                expanded=False):
                                            st.dataframe(df.head(50),
                                                         use_container_width=True,
                                                         hide_index=True)

                                    try:
                                        if out_fmt.startswith("Excel"):
                                            _save_bytes = to_excel_bytes(final_sheets)
                                            _save_name  = (output_name
                                                           if output_name.lower().endswith(".xlsx")
                                                           else _base_name + ".xlsx")
                                        else:
                                            _first_df   = next(iter(final_sheets.values()))
                                            _save_bytes = to_csv_bytes(_first_df)
                                            _save_name  = (output_name
                                                           if output_name.lower().endswith(".csv")
                                                           else _base_name + ".csv")
                                        _save_path = os.path.join(folder_path, _save_name)
                                        with open(_save_path, "wb") as fh:
                                            fh.write(_save_bytes)
                                        _sz_kb = len(_save_bytes) / 1024
                                        _sz_str = (f"{_sz_kb/1024:.2f} MB"
                                                   if _sz_kb > 1024
                                                   else f"{_sz_kb:.0f} KB")
                                        st.info(f"✅ Saved: `{_save_path}`  "
                                                f"({_sz_str})")
                                    except Exception as e:
                                        st.warning(f"Could not save to folder: {e}")

                                    if out_fmt.startswith("Excel"):
                                        dl_data = to_excel_bytes(final_sheets)
                                        dl_name = output_name
                                        dl_mime = ("application/vnd.openxmlformats-"
                                                   "officedocument.spreadsheetml.sheet")
                                    else:
                                        first_df = next(iter(final_sheets.values()))
                                        dl_data  = to_csv_bytes(first_df)
                                        dl_name  = output_name.replace(".xlsx", ".csv")
                                        dl_mime  = "text/csv"

                                    # Store for persistence
                                    st.session_state["_folder_dl"] = {
                                        "data": dl_data, "name": dl_name,
                                        "mime": dl_mime}

                                    st.info("Go to the **📊 Dashboard** tab to explore "
                                            "and export charts.")

                            # ── Persistent download — survives widget re-runs ──
                            if "_folder_dl" in st.session_state:
                                _dl = st.session_state["_folder_dl"]
                                st.download_button(
                                    f"⬇️ Download {_dl['name']}",
                                    data=_dl["data"], file_name=_dl["name"],
                                    mime=_dl["mime"],
                                    use_container_width=True, type="primary",
                                    key="folder_dl_persistent",
                                    help="Your last folder merge output. "
                                         "Re-click 'Run Folder Merge' to regenerate.")


        # ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
with tab_dash:
    st.subheader("Dashboard")
    render_dashboard()


st.divider()
st.caption("DataMerge Studio · column mapping · type cleaning · dedup audit · "
           "interactive dashboard · pd.concat() + drop_duplicates()")
