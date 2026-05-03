"""
File Merger Pro — Streamlit application for merging large CSV / Excel files.
Supports multi-sheet Excel files, column mapping across files with different
column names, data-type cleaning, duplicate-audit reporting, and an
interactive dashboard with offline-capable HTML export.
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

st.set_page_config(page_title="File Merger Pro", page_icon="🔀",
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
def get_dummy():
    fa = pd.DataFrame({
        "Employee ID": ["E001","E002","E003","E004","E005"],
        "Name":        ["Alice","Bob","Charlie","Diana","Eve"],
        "Department":  ["Engineering","Marketing","Engineering","HR","Finance"],
        "Salary":      [85000,72000,90000,68000,78000],
        "Manager":     ["Tom","Sara","Tom","Raj","Sara"],
    })
    fb = pd.DataFrame({
        "Employee ID": ["E003","E004","E005","E006","E007"],
        "Name":        ["Charlie","DIANA","Eve","Frank","Grace"],
        "Department":  ["Engineering","HR","Finance","Finance","Marketing"],
        "Salary":      [90000,71500,None,88000,76000],
        "Manager":     ["Tom","Raj","Sara","Raj","Sara"],
    })
    return fa, fb


# ═══════════════════════════════════════════════════════════════════════════════
# DATA TYPE CLEANING
# ═══════════════════════════════════════════════════════════════════════════════

def _col_words(name):
    return set(re.split(r"[\s_\-/()+]+", name.lower()))


def clean_dtypes(df):
    """
    Auto-clean column types in-place.
    Returns (cleaned_df, report_list).
      - Date-sounding columns: try pd.to_datetime (dayfirst=True)
      - City/region columns:   strip whitespace + title-case
      - All other str columns: strip whitespace
    """
    df = df.copy()
    report = []
    for col in df.columns:
        if col in TRACKING_COLS:
            continue
        words = _col_words(col)

        # ── Date columns ──────────────────────────────────────────────────
        if words & DATE_KEYWORDS and df[col].dtype == object:
            conv  = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
            total = int(df[col].notna().sum())
            hit   = int(conv.notna().sum())
            if total > 0 and hit / total >= 0.5:
                df[col] = conv
                report.append(f"'{col}' → datetime  ({hit}/{total} values parsed)")
                continue

        # ── String cleanup ────────────────────────────────────────────────
        if df[col].dtype == object:
            before  = df[col].fillna("").copy()
            df[col] = df[col].str.strip()
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
# MERGE STRATEGY FUNCTIONS  — all return (result_df, audit_df)
# To add a new strategy: write a function here, then add an entry to STRATEGIES.
# ═══════════════════════════════════════════════════════════════════════════════

def do_append(dfs, key=None, excl=None):
    return pd.concat(dfs, ignore_index=True), pd.DataFrame()

def do_exact_dedup(dfs, key=None, excl=None):
    combined = pd.concat(dfs, ignore_index=True)
    check = [c for c in combined.columns
             if c not in (excl or set()) and c not in TRACKING_COLS]
    return dedup_with_audit(combined, check)

def do_key_first(dfs, key, excl=None):
    combined = pd.concat(dfs, ignore_index=True)
    return dedup_with_audit(combined, [key])

def do_key_last(dfs, key, excl=None):
    combined = pd.concat(dfs, ignore_index=True)
    rev      = combined.iloc[::-1].reset_index(drop=True)
    clean, audit = dedup_with_audit(rev, [key])
    return clean.iloc[::-1].reset_index(drop=True), audit

def do_smart_fill(dfs, key, excl=None):
    combined = pd.concat(dfs, ignore_index=True)
    result   = combined.groupby(key, sort=False, as_index=False).first()
    cols     = [c for c in combined.columns if c in result.columns]
    return result[cols].reset_index(drop=True), pd.DataFrame()

def do_intersection(dfs, key, excl=None):
    sets   = [set(df[key].dropna()) for df in dfs]
    common = sets[0].intersection(*sets[1:])
    comb   = pd.concat(dfs, ignore_index=True)
    sub    = comb[comb[key].isin(common)].reset_index(drop=True)
    return dedup_with_audit(sub, [key])


STRATEGIES = {
    "1 — Simple Append": dict(
        fn=do_append, needs_key=False, allows_excl=False,
        icon="➕", badge="badge-blue", badge_lbl="No key needed",
        head="Stack all rows as-is. Every row kept, including exact duplicates.",
        detail="No filtering at all. Use when files cover completely different records.",
        best="Daily exports with no expected overlap.",
    ),
    "2 — Remove Exact Duplicates": dict(
        fn=do_exact_dedup, needs_key=False, allows_excl=True,
        icon="🧹", badge="badge-green", badge_lbl="No key · exclude-cols optional",
        head="Remove rows where every checked column is identical.",
        detail=("ALL columns must match for a row to be a duplicate. "
                "Optionally exclude specific columns (e.g. engineer name) from the check — "
                "so the same call assigned to two engineers is still treated as one record."),
        best="Two system exports of the same data with minor field differences.",
    ),
    "3 — Key Dedup: First File Wins": dict(
        fn=do_key_first, needs_key=True, allows_excl=False,
        icon="1️⃣", badge="badge-green", badge_lbl="Key column required",
        head="When a key repeats across files, keep the FIRST file's version.",
        detail="All unique keys are kept; only the first-uploaded version survives for duplicates.",
        best="First file is the authoritative master source.",
    ),
    "4 — Key Dedup: Last File Wins": dict(
        fn=do_key_last, needs_key=True, allows_excl=False,
        icon="🔄", badge="badge-orange", badge_lbl="Key column required",
        head="When a key repeats across files, keep the LAST file's version.",
        detail="Same as option 3 but the most recently uploaded file's version is kept.",
        best="Latest file has the freshest data.",
    ),
    "5 — Smart Fill (Best of Both)": dict(
        fn=do_smart_fill, needs_key=True, allows_excl=False,
        icon="🧩", badge="badge-purple", badge_lbl="Key column required",
        head="For duplicate keys, fill empty cells from the other file.",
        detail="First non-empty value per column wins. Upload authoritative file first.",
        best="Two systems each export partial data for the same records.",
    ),
    "6 — Intersection Only": dict(
        fn=do_intersection, needs_key=True, allows_excl=False,
        icon="🔍", badge="badge-red", badge_lbl="Key column required",
        head="Keep ONLY records whose key exists in EVERY uploaded file.",
        detail="Records unique to any single file are excluded entirely.",
        best="Records confirmed across all sources.",
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
        "1 — Simple Append": dict(
            c10=C["a"], c11=C["both"], c01=C["b"],
            l10="E001\nE002", l11="E003\nE004\nE005\n(×2 copies)", l01="E006\nE007",
            title="Simple Append",
            desc="All rows kept — every duplicate appears twice",
            rows="10 rows out (5 + 5, nothing removed)",
        ),
        "2 — Remove Exact Duplicates": dict(
            c10=C["a"], c11=C["both"], c01=C["b"],
            l10="E001\nE002",
            l11="E003 → 1 copy\n(exact dup)\nE004 & E005\n→ both kept\n(data differs)",
            l01="E006\nE007",
            title="Remove Exact Duplicates",
            desc="Only rows identical in EVERY column are collapsed to one copy",
            rows="9 rows out (E003 collapsed; E004 & E005 differ so both kept)",
        ),
        "3 — Key Dedup: First File Wins": dict(
            c10=C["a"], c11=C["a"], c01=C["b"],
            l10="E001\nE002", l11="File A\nwins\nE003,E004\nE005", l01="E006\nE007",
            title="Key Dedup — First File Wins",
            desc="Shared keys → File A's row is kept; File B's version is discarded",
            rows="7 rows out (1 row per unique key; intersection = File A version)",
        ),
        "4 — Key Dedup: Last File Wins": dict(
            c10=C["a"], c11=C["b"], c01=C["b"],
            l10="E001\nE002", l11="File B\nwins\nE003,E004\nE005", l01="E006\nE007",
            title="Key Dedup — Last File Wins",
            desc="Shared keys → File B's row is kept; File A's version is discarded",
            rows="7 rows out (1 row per unique key; intersection = File B version)",
        ),
        "5 — Smart Fill (Best of Both)": dict(
            c10=C["a"], c11=C["fill"], c01=C["b"],
            l10="E001\nE002", l11="Merged\nE003,E004\nE005\n(gaps filled)", l01="E006\nE007",
            title="Smart Fill (Best of Both)",
            desc="Shared keys → row assembled from first non-empty value across files",
            rows="7 rows out (E005 salary: None→78000 filled from File A)",
        ),
        "6 — Intersection Only": dict(
            c10=C["gray"], c11=C["both"], c01=C["gray"],
            l10="E001,E002\nexcluded", l11="E003\nE004\nE005", l01="E006,E007\nexcluded",
            title="Intersection Only",
            desc="Only records whose key exists in ALL files are kept",
            rows="3 rows out (only E003, E004, E005 appear in both files)",
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
    for fname, _sname, df in all_triples:
        file_cols.setdefault(fname, set()).update(
            c for c in df.columns if c not in TRACKING_COLS)

    fnames = list(file_cols.keys())
    if len(fnames) < 2:
        return st.session_state.get(ss_key, {})

    all_unique = set().union(*file_cols.values())
    common     = set.intersection(*file_cols.values())
    unmatched  = sorted(all_unique - common)

    if not unmatched:
        st.success("All files share identical column names — no mapping needed.")
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

def render_settings(all_cols, tab_key):
    col_l, col_r = st.columns(2)
    with col_l:
        chosen = st.selectbox("Merge strategy", list(STRATEGIES.keys()),
                              key=f"strat_{tab_key}")
        cfg = STRATEGIES[chosen]
        st.markdown(f'<span class="badge {cfg["badge"]}">{cfg["badge_lbl"]}</span>',
                    unsafe_allow_html=True)
        st.caption(cfg["head"])
    with col_r:
        key_col, excl_cols = None, []
        if cfg["needs_key"] and all_cols:
            key_col = st.selectbox("Key column", sorted(all_cols),
                                   key=f"key_{tab_key}")
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
    return cfg, key_col, excl_cols, clean_types, add_src, out_fmt


# ═══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

def render_dashboard():
    """Interactive dashboard — reads merged data from session state."""
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
    df = sheets[sel].copy()

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

    # Auto-detect column types
    date_cols = [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]
    num_cols  = [c for c in df.columns
                 if pd.api.types.is_numeric_dtype(df[c]) and c not in TRACKING_COLS]
    cat_cols  = [c for c in df.columns
                 if df[c].dtype == object and 1 < df[c].nunique() <= 40
                 and c not in TRACKING_COLS]

    figs_for_export = []   # (title, plotly_fig)

    # ── Source breakdown ──────────────────────────────────────────────────────
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

    # ── Date trends ───────────────────────────────────────────────────────────
    if date_cols:
        st.subheader("Date Trends")
        dc1, dc2 = st.columns([2, 1])
        with dc1:
            dcol = st.selectbox("Date column", date_cols, key="dash_dc")
        with dc2:
            agg_lbl = st.selectbox("Group by", ["Day", "Week", "Month"], key="dash_agg")
        freq = {"Day": "D", "Week": "W", "Month": "ME"}[agg_lbl]
        ts = (df.set_index(dcol).resample(freq).size()
                .rename("Count").reset_index())
        if HAS_PLOTLY:
            fig = px.line(ts, x=dcol, y="Count", markers=True, height=300)
            fig.update_traces(line_color="#2563eb")
            st.plotly_chart(fig, use_container_width=True)
            figs_for_export.append((f"Trend by {dcol}", fig))
        else:
            st.line_chart(ts.set_index(dcol))

    # ── Categorical breakdowns ────────────────────────────────────────────────
    if cat_cols:
        st.subheader("Categorical Breakdowns")
        sel_cats = st.multiselect(
            "Columns to chart", cat_cols,
            default=cat_cols[:min(4, len(cat_cols))], key="dash_cats")
        if sel_cats:
            ncols = min(len(sel_cats), 2)
            rows  = (len(sel_cats) + ncols - 1) // ncols
            grid  = [st.columns(ncols) for _ in range(rows)]
            flat  = [cell for row in grid for cell in row]
            for i, col in enumerate(sel_cats):
                with flat[i]:
                    vc = (df[col].value_counts().head(15)
                            .rename_axis(col).reset_index(name="Count"))
                    if HAS_PLOTLY:
                        fig = px.bar(vc, x="Count", y=col, orientation="h",
                                     title=col, height=350, color="Count",
                                     color_continuous_scale="Teal")
                        fig.update_layout(
                            yaxis={"categoryorder": "total ascending"},
                            showlegend=False, coloraxis_showscale=False,
                            title_font_size=13)
                        st.plotly_chart(fig, use_container_width=True)
                        figs_for_export.append((col, fig))
                    else:
                        st.write(f"**{col}**")
                        st.bar_chart(vc.set_index(col))

    # ── Numeric summary ───────────────────────────────────────────────────────
    if num_cols:
        st.subheader("Numeric Column Summary")
        st.dataframe(df[num_cols].describe().T.round(2), use_container_width=True)

    # ── Data quality ──────────────────────────────────────────────────────────
    st.subheader("Missing Values % by Column")
    null_s = (df.isnull().sum() / len(df) * 100).round(1).sort_values(ascending=False)
    null_s = null_s[null_s > 0]
    if null_s.empty:
        st.success("No missing values — perfect data quality!")
    elif HAS_PLOTLY:
        fig = px.bar(null_s.reset_index(), x="index", y=null_s.name or 0,
                     labels={"index": "Column", null_s.name or 0: "Missing %"},
                     color=null_s.name or 0, color_continuous_scale="Reds", height=300)
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)
        figs_for_export.append(("Missing Values %", fig))
    else:
        st.bar_chart(null_s)

    # ── Data explorer ─────────────────────────────────────────────────────────
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

    # ── Export standalone HTML ────────────────────────────────────────────────
    st.subheader("Export Dashboard")
    if not HAS_PLOTLY:
        st.info("Install plotly to enable HTML export:  `pip install plotly`")
    elif figs_for_export:
        if st.button("Generate standalone HTML dashboard (works offline)",
                     type="secondary", use_container_width=True):
            parts = [
                "<html><head><meta charset='utf-8'>",
                f"<title>Dashboard — {sel}</title>",
                "<style>body{font-family:sans-serif;margin:32px;background:#f9fafb;}"
                "h1{color:#1e3a5f;}h2{color:#374151;margin-top:32px;}"
                "p{color:#6b7280;}</style></head><body>",
                f"<h1>File Merger Pro — Dashboard</h1>",
                f"<p><b>Sheet:</b> {sel} &nbsp;|&nbsp; "
                f"<b>Rows:</b> {len(df):,} &nbsp;|&nbsp; "
                f"<b>Columns:</b> {len(df.columns)}</p>",
            ]
            first = True
            for title, fig in figs_for_export:
                parts.append(f"<h2>{title}</h2>")
                parts.append(fig.to_html(
                    full_html=False,
                    include_plotlyjs=True if first else False))
                first = False
            parts.append("</body></html>")
            st.download_button(
                "Download dashboard.html",
                data="\n".join(parts).encode("utf-8"),
                file_name="dashboard.html",
                mime="text/html",
                use_container_width=True,
                type="primary")
            st.caption("Fully self-contained HTML — no internet or Streamlit needed to view it.")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE
# ═══════════════════════════════════════════════════════════════════════════════
st.title("File Merger Pro")
st.markdown(
    "Merge large **CSV or Excel** files — including **multi-sheet** workbooks. "
    "Sheets with identical column structures are merged; those with different "
    "structures become separate output sheets automatically.")
st.divider()

tab_learn, tab_upload, tab_merge, tab_folder, tab_dash = st.tabs(
    ["Learn Merge Types", "Upload Files", "Merge & Download", "Folder Mode", "📊 Dashboard"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — LEARN
# ══════════════════════════════════════════════════════════════════════════════
with tab_learn:
    st.subheader("Which merge type do you need?")
    fa, fb = get_dummy()
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**File A** — HR System (5 rows)")
        st.dataframe(fa, use_container_width=True, hide_index=True)
    with c2:
        st.markdown("**File B** — Payroll System (5 rows)")
        st.dataframe(fb, use_container_width=True, hide_index=True)
    st.info(
        "• **E003/Charlie** — exact duplicate (every column identical in both files)\n"
        "• **E004/Diana** — same ID, but name case and salary differ — NOT a duplicate\n"
        "• **E005/Eve** — same ID, File B has missing salary — NOT a duplicate\n"
        "• **E001,E002** unique to File A   •   **E006,E007** unique to File B")
    st.divider()

    for name, cfg in STRATEGIES.items():
        with st.expander(f"{cfg['icon']}  **{name}**", expanded=False):
            st.markdown(f'<span class="badge {cfg["badge"]}">{cfg["badge_lbl"]}</span>',
                        unsafe_allow_html=True)
            st.markdown(f"**{cfg['head']}**\n\n{cfg['detail']}")
            st.markdown(f"*Best for: {cfg['best']}*")
            excl_demo = ["Manager"] if cfg["allows_excl"] else []
            key = "Employee ID" if cfg["needs_key"] else None
            try:
                raw, _ = cfg["fn"]([fa.copy(), fb.copy()],
                                   key=key, excl=excl_demo or None)
                show = annotate(fa, fb, raw, key)
                col_v, col_t = st.columns([1, 1.4])
                with col_v:
                    fig = make_venn_fig(name)
                    if fig:
                        st.pyplot(fig, use_container_width=True)
                        plt.close(fig)
                with col_t:
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Input (A+B)", len(fa) + len(fb))
                    m2.metric("Output", len(raw))
                    m3.metric("Removed", len(fa) + len(fb) - len(raw))
                    if excl_demo:
                        st.caption(f"Example excludes **Manager** from duplicate check.")
                    st.dataframe(show, use_container_width=True, hide_index=True)
                    st.caption("🟢 File A only   🔵 File B only   🟡 Both files")
            except Exception as e:
                st.warning(f"Example error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with tab_upload:
    st.subheader("Upload Your Files")
    st.markdown(
        "Upload **2 or more** CSV or Excel files. "
        "**Multi-sheet workbooks are fully supported** — all sheets are read and "
        "automatically grouped by column structure.")
    uploaded = st.file_uploader(
        "Drag & drop files here, or click Browse",
        type=["csv", "xlsx", "xls"], accept_multiple_files=True, key="uploader")
    if uploaded:
        st.session_state["uf"] = uploaded
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

        # ── Step 2: Sheet groups ───────────────────────────────────────────
        groups = group_sheets(mapped_triples)
        st.markdown(f"### Step 2 — Sheet Groups  ({len(groups)} detected)")
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

        # ── Step 3: Settings ───────────────────────────────────────────────
        st.markdown("### Step 3 — Merge Settings")
        all_cols = (sorted(set.union(*[set(k) for k, _ in groups]) - TRACKING_COLS)
                    if groups else [])
        cfg, key_col, excl_cols, clean_types, add_src, out_fmt = render_settings(
            all_cols, "upload")
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
                            d = df.copy()
                            if add_src:
                                d.insert(0, "Source File",  fname)
                                d.insert(1, "Source Sheet", sname)
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

                # Save for dashboard
                st.session_state["merged_sheets"] = output_sheets

                st.success("Merge complete!")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Groups merged",     len(output_sheets))
                m2.metric("Total input rows",  f"{total_in:,}")
                m3.metric("Total output rows", f"{total_out:,}")
                m4.metric("Rows removed",      f"{total_in - total_out:,}")

                # Type-cleaning report
                if all_type_reports:
                    with st.expander(
                            f"Data Type Cleaning — "
                            f"{len(set(all_type_reports))} change(s) applied"):
                        for line in sorted(set(all_type_reports)):
                            st.markdown(f"- {line}")

                # Duplicate audit
                show_audit(all_audits)

                # Preview
                for sname, df in output_sheets.items():
                    with st.expander(f"Output: '{sname}'  ({len(df):,} rows)",
                                     expanded=True):
                        st.dataframe(df.head(100), use_container_width=True,
                                     hide_index=True)
                        if len(df) > 100:
                            st.caption(f"Preview: first 100 of {len(df):,} rows.")

                # Download
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

                st.download_button(
                    f"Download {dl_name}",
                    data=dl_data, file_name=dl_name, mime=dl_mime,
                    use_container_width=True, type="primary")

                st.info("Go to the **📊 Dashboard** tab to explore and export charts.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — FOLDER MODE
# ══════════════════════════════════════════════════════════════════════════════
with tab_folder:
    st.subheader("Folder Mode")
    st.markdown(
        "Point to a folder — the app reads all CSV/Excel files automatically, "
        "groups their sheets by column structure, and merges matching groups. "
        "Supports **incremental append** when new files arrive.")

    if not HAS_TKINTER:
        st.error(
            "**Folder Mode requires the app to be running locally on your computer.** "
            "This feature needs access to your file system, which is not possible "
            "on the cloud-hosted version.\n\n"
            "To use Folder Mode:\n"
            "1. Download the app files from GitHub\n"
            "2. Double-click **`run_app.bat`** on your PC\n"
            "3. The app opens at **http://localhost:8501** — Folder Mode will work there\n\n"
            "The **Upload Files** and **Merge & Download** tabs work on both local and cloud.")
        st.stop()

    # ── Browse button (native Windows folder picker) ──────────────────────────
    if HAS_TKINTER and st.button("📁 Browse for folder...", key="folder_browse_btn"):
        _root = tk.Tk()
        _root.withdraw()
        _root.wm_attributes("-topmost", 1)
        _picked = _tkfd.askdirectory(title="Select folder containing your files")
        _root.destroy()
        if _picked:
            # tkinter returns forward-slash paths; normalise for Windows
            st.session_state["_folder_path_input"] = os.path.normpath(_picked)
        st.rerun()

    f_col, o_col = st.columns([3, 2])
    with f_col:
        _default_path = st.session_state.get(
            "_folder_path_input",
            r"C:\Users\Kshitij Buch\OneDrive\Documents\TBM 2026 Onwards\Pending Calls\Raw Files")
        folder_path = st.text_input(
            "Folder path (type or use Browse button above)",
            value=_default_path,
            key="_folder_path_text")
    with o_col:
        output_name = st.text_input("Output filename", value="MERGED_output.xlsx")

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
        all_data_files = sorted([
            f for f in os.listdir(folder_path)
            if f.lower().endswith((".csv", ".xlsx", ".xls")) and f != output_name
        ])

        if not all_data_files:
            st.warning("No CSV or Excel files found in this folder.")
        else:
            # Detect already-processed files from existing output
            already_done       = set()
            existing_row_count = 0
            has_existing       = os.path.exists(output_path)
            ex_sheets          = {}
            if has_existing:
                try:
                    ex_sheets = read_from_path(output_path)
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
                            render_settings(all_folder_cols, "folder")
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
                                    d = df.copy()
                                    if add_src:
                                        d.insert(0, "Source File",  fname)
                                        d.insert(1, "Source Sheet", sname)
                                    if clean_types:
                                        d, trpt = clean_dtypes(d)
                                        all_type_reports.extend(trpt)
                                    all_triples.append((fname, sname, d))

                                with st.spinner("Merging..."):
                                    groups            = group_sheets(all_triples)
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
                                    xl_bytes = to_excel_bytes(final_sheets)
                                    with open(output_path, "wb") as fh:
                                        fh.write(xl_bytes)
                                    st.info(f"Saved: `{output_path}`")
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

                                st.download_button(
                                    f"Download {dl_name}", data=dl_data,
                                    file_name=dl_name, mime=dl_mime,
                                    use_container_width=True, type="primary")

                                st.info("Go to the **📊 Dashboard** tab to explore "
                                        "and export charts.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
with tab_dash:
    st.subheader("Dashboard")
    render_dashboard()


st.divider()
st.caption("File Merger Pro · column mapping · type cleaning · dedup audit · "
           "interactive dashboard · pd.concat() + drop_duplicates()")
