"""
File Merger Pro — Streamlit application for merging large CSV / Excel files.
Supports multi-sheet Excel files: sheets with identical column sets are merged;
sheets with different column sets get separate output sheets automatically.
Run with:  streamlit run file_merger_app.py
"""

import io, os
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from matplotlib_venn import venn2, venn2_circles

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


# ─── Tracking columns added by this app ───────────────────────────────────────
TRACKING_COLS = {"Source File", "Source Date", "Source Sheet"}

# ─── Dummy data ────────────────────────────────────────────────────────────────
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


# ─── Merge implementations ─────────────────────────────────────────────────────
def do_append(dfs, key=None, excl=None):
    return pd.concat(dfs, ignore_index=True)

def do_exact_dedup(dfs, key=None, excl=None):
    combined = pd.concat(dfs, ignore_index=True)
    check = [c for c in combined.columns
             if c not in (excl or set()) and c not in TRACKING_COLS]
    return combined.drop_duplicates(subset=check, ignore_index=True)

def do_key_first(dfs, key, excl=None):
    combined = pd.concat(dfs, ignore_index=True)
    return combined.drop_duplicates(subset=[key], keep="first", ignore_index=True)

def do_key_last(dfs, key, excl=None):
    combined = pd.concat(dfs, ignore_index=True)
    return combined.drop_duplicates(subset=[key], keep="last", ignore_index=True)

def do_smart_fill(dfs, key, excl=None):
    combined = pd.concat(dfs, ignore_index=True)
    result = combined.groupby(key, sort=False, as_index=False).first()
    cols = [c for c in combined.columns if c in result.columns]
    return result[cols].reset_index(drop=True)

def do_intersection(dfs, key, excl=None):
    sets = [set(df[key].dropna()) for df in dfs]
    common = sets[0].intersection(*sets[1:])
    combined = pd.concat(dfs, ignore_index=True)
    return (combined[combined[key].isin(common)]
            .drop_duplicates(subset=[key], keep="first")
            .reset_index(drop=True))

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

def _row_sig(row):
    """Robust row fingerprint: normalises int/float differences and NaN."""
    parts = []
    for v in row:
        try:
            if pd.isna(v):
                parts.append("__NaN__")
                continue
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
        a_k = set(fa[key].dropna())
        b_k = set(fb[key].dropna())
        r.insert(0, "Source",
                 [("🟡 Both"    if v in a_k and v in b_k
                   else "🟢 File A" if v in a_k
                   else "🔵 File B")
                  for v in r[key]])
    else:
        # Use robust fingerprinting so int/float mismatches don't mislead
        a_sigs = {_row_sig(row) for row in fa.values}
        b_sigs = {_row_sig(row) for row in fb.values}
        sources = []
        for row in result.values:
            sig = _row_sig(row)
            in_a, in_b = sig in a_sigs, sig in b_sigs
            if in_a and in_b:
                sources.append("🟡 Both files")
            elif in_a:
                sources.append("🟢 File A only")
            else:
                sources.append("🔵 File B only")
        r.insert(0, "Source", sources)
    return r


# ─── Venn diagram per strategy ────────────────────────────────────────────────
# Dummy data key membership:
#   File A keys: E001 E002 E003 E004 E005  (A-only: E001,E002  |  shared: E003,E004,E005)
#   File B keys: E003 E004 E005 E006 E007  (B-only: E006,E007  |  shared: E003,E004,E005)
#   E003/Charlie is an exact duplicate; E004 and E005 differ in at least one column.

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
            l10="E001\nE002", l11="E003 → 1 copy\n(exact dup)\nE004 & E005\n→ both kept\n(data differs)", l01="E006\nE007",
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
            p.set_facecolor(col)
            p.set_alpha(0.82)

    for rid, txt in [("10", cfg["l10"]), ("11", cfg["l11"]), ("01", cfg["l01"])]:
        lbl = v.get_label_by_id(rid)
        if lbl:
            lbl.set_text(txt)
            lbl.set_fontsize(7)

    for sid, col in [("A", "#166534"), ("B", "#1d4ed8")]:
        lbl = v.get_label_by_id(sid)
        if lbl:
            lbl.set_fontsize(10)
            lbl.set_fontweight("bold")
            lbl.set_color(col)

    ax.set_title(cfg["title"], fontsize=10, fontweight="bold", color="#111827", pad=8)
    fig.text(0.5, 0.13, cfg["desc"], ha="center", fontsize=7.5, color="#374151")
    fig.text(0.5, 0.04, cfg["rows"],  ha="center", fontsize=7,   color="#6b7280",
             style="italic")

    plt.tight_layout(rect=[0, 0.16, 1, 1])
    return fig


# ─── File reading helpers ──────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def read_all_sheets_cached(file_bytes, filename):
    """Returns dict: sheet_name -> DataFrame. CSV returns {"Sheet1": df}."""
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
    """Frozenset of column names excluding tracking cols — used as group key."""
    return frozenset(c for c in df.columns if c not in TRACKING_COLS)


def group_sheets(file_sheet_dfs):
    """
    file_sheet_dfs: list of (filename, sheet_name, df)
    Returns: list of groups, each group = (col_frozenset, [(filename, sheet_name, df)])
    Groups sorted by descending column count (richest first).
    """
    buckets = {}
    for fname, sheet, df in file_sheet_dfs:
        key = col_sig(df)
        buckets.setdefault(key, []).append((fname, sheet, df))
    return sorted(buckets.items(), key=lambda x: -len(x[0]))


def to_excel_bytes(sheet_dict):
    """sheet_dict: {sheet_name: df} -> bytes of Excel file."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheet_dict.items():
            safe = name[:31].translate(str.maketrans(r'\/[]*?:', '_______'))
            df.to_excel(w, sheet_name=safe, index=False)
    return buf.getvalue()


def to_csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8")


# ─── Shared merge-settings widget ─────────────────────────────────────────────
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
        add_src  = st.checkbox("Add 'Source File' column", value=True,
                               key=f"src_{tab_key}")
        out_fmt  = st.radio("Download format",
                            ["Excel (.xlsx) — multi-sheet", "CSV (.csv) — first sheet only"],
                            key=f"fmt_{tab_key}", horizontal=True)
    return cfg, key_col, excl_cols, add_src, out_fmt


# ══════════════════════════════════════════════════════════════════════════════
# PAGE
# ══════════════════════════════════════════════════════════════════════════════
st.title("File Merger Pro")
st.markdown(
    "Merge large **CSV or Excel** files — including **multi-sheet** workbooks. "
    "Sheets with identical column structures are merged; those with different "
    "structures become separate output sheets automatically.")
st.divider()

tab_learn, tab_upload, tab_merge, tab_folder = st.tabs(
    ["Learn Merge Types", "Upload Files", "Merge & Download", "Folder Mode"])


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
                raw = cfg["fn"]([fa.copy(), fb.copy()], key=key, excl=excl_demo or None)
                show = annotate(fa, fb, raw, key)

                col_v, col_t = st.columns([1, 1.4])
                with col_v:
                    fig = make_venn_fig(name)
                    if fig:
                        st.pyplot(fig, use_container_width=True)
                        plt.close(fig)
                with col_t:
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Input (A+B)", len(fa)+len(fb))
                    m2.metric("Output", len(raw))
                    m3.metric("Removed", len(fa)+len(fb)-len(raw))
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
        type=["csv","xlsx","xls"], accept_multiple_files=True, key="uploader")

    if uploaded:
        st.session_state["uf"] = uploaded
        st.success(f"{len(uploaded)} file(s) ready. Go to **Merge & Download**.")

        for i, uf in enumerate(uploaded):
            with st.expander(f"File {i+1}: {uf.name}", expanded=(i==0)):
                try:
                    data = uf.read(); uf.seek(0)
                    sheets = read_all_sheets_cached(data, uf.name)
                    st.markdown(f"**{len(sheets)} sheet(s) found:**")
                    for sname, df in sheets.items():
                        c1,c2,c3 = st.columns(3)
                        c1.metric(f"'{sname}' rows", f"{len(df):,}")
                        c2.metric("Columns", len(df.columns))
                        c3.metric("File size", f"{uf.size/1024:.1f} KB")
                        st.dataframe(df.head(5), use_container_width=True,
                                     hide_index=True)
                        if len(df) > 5:
                            st.caption(f"Showing 5 of {len(df):,} rows.")
                        st.divider()
                except Exception as e:
                    st.error(f"Could not read: {e}")
    else:
        st.info("No files uploaded yet.")
        st.session_state.pop("uf", None)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — MERGE
# ══════════════════════════════════════════════════════════════════════════════
with tab_merge:
    st.subheader("Configure and Run Merge")
    files = st.session_state.get("uf", [])
    if len(files) < 2:
        st.warning("Upload at least 2 files in the **Upload Files** tab first.")
    else:
        # Collect all (filename, sheet_name, df) triples
        all_triples = []
        for uf in files:
            try:
                data = uf.read(); uf.seek(0)
                sheets = read_all_sheets_cached(data, uf.name)
                for sname, df in sheets.items():
                    all_triples.append((uf.name, sname, df))
            except Exception as e:
                st.error(f"Error reading {uf.name}: {e}")

        groups = group_sheets(all_triples)

        # ── Show detected groups ──────────────────────────────────────────
        st.markdown(f"**{len(groups)} sheet group(s) detected** "
                    f"(sheets with identical column sets will be merged together):")
        for i, (cols_key, entries) in enumerate(groups):
            file_names  = sorted({e[0] for e in entries})
            sheet_names = sorted({e[1] for e in entries})
            total_rows  = sum(len(e[2]) for e in entries)
            bg = "#f0fdf4" if len(entries) > 1 else "#fefce8"
            action = f"MERGE {len(entries)} sheets" if len(entries)>1 else "1 sheet (no merge needed)"
            st.markdown(
                f'<div class="group-box" style="background:{bg}">'
                f'<b>Group {i+1}</b> — {len(cols_key)} columns — {total_rows:,} rows — {action}<br>'
                f'<small>Sheet names: {", ".join(sheet_names)}<br>'
                f'Files: {", ".join(f[:40] for f in file_names)}</small></div>',
                unsafe_allow_html=True)

        st.divider()

        # ── Common cols across ALL groups (for settings widgets) ──────────
        all_cols = sorted(set.union(*[set(k) for k,_ in groups]) - TRACKING_COLS) \
                   if groups else []

        cfg, key_col, excl_cols, add_src, out_fmt = render_settings(all_cols, "upload")
        st.divider()

        if st.button("Run Merge", type="primary", use_container_width=True):
            if cfg["needs_key"] and not key_col:
                st.error("Please select a key column.")
            else:
                output_sheets = {}
                total_in = total_out = 0

                with st.spinner("Merging..."):
                    for i, (cols_key, entries) in enumerate(groups):
                        # Build suggested output name from sheet names in group
                        sheet_names_set = sorted({e[1] for e in entries})
                        out_name = " + ".join(sheet_names_set)[:31]

                        dfs = []
                        for fname, sname, df in entries:
                            d = df.copy()
                            if add_src:
                                d.insert(0, "Source File",  fname)
                                d.insert(1, "Source Sheet", sname)
                            dfs.append(d)

                        try:
                            result = cfg["fn"](dfs, key=key_col,
                                               excl=set(excl_cols) if excl_cols else None)
                        except Exception as e:
                            st.warning(f"Group {i+1} merge failed: {e}")
                            result = pd.concat(dfs, ignore_index=True)

                        n_in  = sum(len(d) for d in dfs)
                        total_in  += n_in
                        total_out += len(result)
                        output_sheets[out_name] = result

                st.success("Merge complete!")
                m1,m2,m3,m4 = st.columns(4)
                m1.metric("Groups merged", len(output_sheets))
                m2.metric("Total input rows", f"{total_in:,}")
                m3.metric("Total output rows", f"{total_out:,}")
                m4.metric("Rows removed", f"{total_in-total_out:,}")

                # Preview each output sheet
                for sname, df in output_sheets.items():
                    with st.expander(f"Output sheet: '{sname}'  ({len(df):,} rows)", expanded=True):
                        st.dataframe(df.head(100), use_container_width=True, hide_index=True)
                        if len(df) > 100:
                            st.caption(f"Preview: first 100 of {len(df):,} rows.")

                # Download
                if out_fmt.startswith("Excel"):
                    dl_data  = to_excel_bytes(output_sheets)
                    dl_name  = "merged_output.xlsx"
                    dl_mime  = ("application/vnd.openxmlformats-"
                                "officedocument.spreadsheetml.sheet")
                else:
                    first_df = next(iter(output_sheets.values()))
                    dl_data  = to_csv_bytes(first_df)
                    dl_name  = "merged_output.csv"
                    dl_mime  = "text/csv"

                st.download_button(
                    f"Download merged file ({dl_name})",
                    data=dl_data, file_name=dl_name, mime=dl_mime,
                    use_container_width=True, type="primary")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — FOLDER MODE
# ══════════════════════════════════════════════════════════════════════════════
with tab_folder:
    st.subheader("Folder Mode")
    st.markdown(
        "Point to a folder — the app reads all CSV/Excel files automatically, "
        "groups their sheets by column structure, and merges matching groups. "
        "Supports **incremental append** when new files arrive.")

    f_col, o_col = st.columns([3,2])
    with f_col:
        folder_path = st.text_input(
            "Folder path",
            value=r"C:\Users\Kshitij Buch\OneDrive\Documents\TBM 2026 Onwards\Pending Calls")
    with o_col:
        output_name = st.text_input("Output filename", value="MERGED_output.xlsx")

    if not folder_path:
        st.info("Enter a folder path above.")
        st.stop()
    if not os.path.isdir(folder_path):
        st.error(f"Folder not found: `{folder_path}`")
        st.stop()

    output_path = os.path.join(folder_path, output_name)

    all_data_files = sorted([
        f for f in os.listdir(folder_path)
        if f.lower().endswith((".csv",".xlsx",".xls")) and f != output_name
    ])
    if not all_data_files:
        st.warning("No CSV or Excel files found in this folder.")
        st.stop()

    # Detect already-processed files via Source File column in existing output
    already_done = set()
    existing_row_count = 0
    has_existing = os.path.exists(output_path)
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

    s1,s2,s3 = st.columns(3)
    s1.metric("Files in folder",           len(all_data_files))
    s2.metric("Already in merged output",  len(all_data_files)-len(new_files))
    s3.metric("New files detected",        len(new_files))

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
        st.stop()

    st.markdown(f"**Files to process ({len(files_to_show)}) — uncheck to skip:**")
    selected = []
    cols_row = st.columns(2)
    for i, f in enumerate(files_to_show):
        with cols_row[i % 2]:
            if st.checkbox(f, value=True, key=f"fc_{f}"):
                selected.append(f)

    if not selected:
        st.warning("No files selected.")
        st.stop()

    st.divider()
    st.markdown("**Merge Settings**")

    # Sample first selected file to get column names for dropdowns
    try:
        sample_sheets = read_from_path(os.path.join(folder_path, selected[0]))
        sample_cols   = sorted(set().union(*[set(d.columns) for d in sample_sheets.values()]) - TRACKING_COLS)
    except Exception as e:
        st.error(f"Could not read {selected[0]}: {e}")
        sample_cols = []

    cfg, key_col, excl_cols, add_src, out_fmt = render_settings(sample_cols, "folder")
    st.divider()

    if st.button("Run Folder Merge", type="primary", use_container_width=True):
        if cfg["needs_key"] and not key_col:
            st.error("Please select a key column.")
            st.stop()

        prog = st.progress(0, text="Reading files...")
        all_triples = []
        errors = []

        for i, fname in enumerate(selected):
            fpath = os.path.join(folder_path, fname)
            try:
                sheets = read_from_path(fpath)
                for sname, df in sheets.items():
                    d = df.copy()
                    if add_src:
                        d.insert(0, "Source File",  fname)
                        d.insert(1, "Source Sheet", sname)
                    all_triples.append((fname, sname, d))
            except Exception as e:
                errors.append(f"{fname}: {e}")
            prog.progress((i+1)/len(selected), text=f"Reading: {fname}")

        prog.empty()
        for err in errors:
            st.warning(f"Could not read: {err}")

        if not all_triples:
            st.error("No sheets could be read.")
            st.stop()

        with st.spinner("Grouping and merging..."):
            groups = group_sheets(all_triples)
            new_output_sheets = {}

            for cols_key, entries in groups:
                sheet_names = sorted({e[1] for e in entries})
                out_name = " + ".join(sheet_names)[:31]
                dfs = [e[2] for e in entries]
                try:
                    result = cfg["fn"](dfs, key=key_col,
                                       excl=set(excl_cols) if excl_cols else None)
                except Exception as e:
                    st.warning(f"Merge failed for group {out_name}: {e}")
                    result = pd.concat(dfs, ignore_index=True)
                new_output_sheets[out_name] = result

            if mode.startswith("Append") and has_existing:
                # Load existing sheets and append new rows
                existing_sheets = read_from_path(output_path)
                final_sheets = dict(existing_sheets)
                for sname, new_df in new_output_sheets.items():
                    if sname in final_sheets:
                        final_sheets[sname] = pd.concat(
                            [final_sheets[sname], new_df], ignore_index=True)
                    else:
                        final_sheets[sname] = new_df
            else:
                final_sheets = new_output_sheets

        n_total = sum(len(d) for d in final_sheets.values())
        st.success(f"Done! {len(final_sheets)} sheet(s), {n_total:,} total rows.")

        for sname, df in final_sheets.items():
            with st.expander(f"Sheet: '{sname}'  ({len(df):,} rows)", expanded=False):
                st.dataframe(df.head(50), use_container_width=True, hide_index=True)

        # Save to folder
        try:
            xl_bytes = to_excel_bytes(final_sheets)
            with open(output_path, "wb") as fh:
                fh.write(xl_bytes)
            st.info(f"Saved: `{output_path}`")
        except Exception as e:
            st.warning(f"Could not save to folder: {e}")

        # Download
        if out_fmt.startswith("Excel"):
            dl_data = to_excel_bytes(final_sheets)
            dl_name, dl_mime = output_name, ("application/vnd.openxmlformats-"
                                             "officedocument.spreadsheetml.sheet")
        else:
            first_df = next(iter(final_sheets.values()))
            dl_data  = to_csv_bytes(first_df)
            dl_name  = output_name.replace(".xlsx",".csv")
            dl_mime  = "text/csv"

        st.download_button(f"Download {dl_name}", data=dl_data,
                           file_name=dl_name, mime=dl_mime,
                           use_container_width=True, type="primary")

st.divider()
st.caption("File Merger Pro · multi-sheet aware · "
           "pd.concat() for stacking · drop_duplicates() for dedup")
