# Claude_Context.md — DataMerge Studio handoff

> **⚠️ 2026-07-11 UPDATE — READ `CLAUDE.md` FIRST.** The repo has since moved
> OUT of OneDrive to `C:\Users\k.buch\Documents\file merger and join app`, all
> commits are pushed to `origin`, and the live cloud URL is now
> `https://tbmfilemergejoin.streamlit.app`. `CLAUDE.md` (auto-loaded) is the
> current source of truth. Sections below on **paths, git state, remotes,
> "action items", and blockers are HISTORICAL (as of 2026-05-11)** and may be
> stale — but the **architecture, design decisions, and code snippets
> (Sections 3, 7, 9) remain accurate and useful.**

> **Purpose:** detailed architecture reference. For a quick current summary,
> see `CLAUDE.md`.

---

## 1. Project

**Name:** DataMerge Studio (formerly "File Merger Pro")
**What it is:** A Streamlit web app that merges & joins CSV / Excel files using
all standard SQL set operations, with column mapping, data-type cleaning,
duplicate audit, and an interactive dashboard.

**Working directory (Windows):**
```
C:\Users\Kshitij Buch\OneDrive\Documents\TBM 2026 Onwards\Pending Calls
```

**Hosted demo (Streamlit Cloud):**
```
https://filemerger-kb.streamlit.app
```

**GitHub remote — confirmed via verification on 2026-05-11:**
```
origin → git@github-company:tbmtsdai/File_Merger.git   ← LIVE source (cloud deploys from here)
```
The Streamlit Cloud app at `https://filemerger-kb.streamlit.app` is wired to
the **office `tbmtsdai/File_Merger` repo**, not the user's personal account.
The URL slug `kshitijbuch-file_merger` in the logs is just the deployer's
username — misleading; it does **not** indicate which repo is being deployed.

A secondary mirror exists at `github.com/kshitijbuch/File_Merger` (personal
account, remote name `personal`). It was sync'd as a backup on 2026-05-11.
**Cloud does not deploy from it.** Don't waste time pushing to it unless
explicitly asked.

> ### ⚠️ ACTION REQUIRED ON OFFICE LAPTOP — DO THIS FIRST
> As of 2026-05-11, **`origin` (the office repo, the one cloud actually
> deploys from) is 2 commits behind** `personal`. These 2 commits exist
> ONLY in this personal-OneDrive working copy and on the `personal` mirror —
> they have **never reached `origin`**:
> ```
> c7b3161  Update Claude_Context.md — cloud confirmed deploying from office tbmtsdai repo
> 78328a0  Add Claude_Context.md handoff doc for cross-session continuity
> ```
> Both are just this handoff-doc file — **no app code is at risk** — but push
> them so `origin` has the full history before this personal OneDrive folder
> is deleted. From the office laptop / wherever the `github-company` SSH
> alias resolves, run:
> ```bash
> cd "<office OneDrive path to this project>"
> git fetch origin
> git log origin/master..HEAD          # confirm these 2 commits are the only diff
> git push origin master
> ```
> If the office laptop's working copy is a **different clone** (not this
> exact personal-OneDrive folder), you'll need to either (a) re-create
> `Claude_Context.md` there manually using this file as the source, or
> (b) add this personal-OneDrive path as a remote on the office clone and
> pull from it once, e.g. `git remote add winhome <path-or-share>` then
> `git fetch winhome && git merge winhome/master`.
> **Once confirmed pushed, this entire warning block can be deleted.**

**Pushing from this Windows machine fails:** `origin` uses SSH alias
`github-company` which is not in this machine's `~/.ssh/config`. Push from
the office machine (where the alias is configured) or fix the SSH config
locally — see Section 6, Next Step 1.

---

## 2. Current state — git

**Branch:** `master`
**HEAD commit:** `f3cf588 Add Deduplicate tab: remove duplicates from a single file with column exclusions`
**Working tree:** clean (no uncommitted changes)
**Remote:** up to date with `origin/master`

**Recent commits (newest first):**
```
f3cf588  Add Deduplicate tab — single-file dedup with column exclusions
195d32f  Fix dark-mode contrast — group-boxes and badges invisible on dark BG
dd7eb04  Add 'Drop columns from output' feature in Merge Settings
3168a3b  Hide Folder Mode tab entirely on cloud — only visible when local
7683da5  Folder Mode: remove tkinter gate, friendlier message, Browse-optional UX
e2db529  Fix file uploader not clearing after New Session reset
2cfbc8e  Add CSV save support in Folder Mode — respects format choice
888c8e9  Upgrade dashboard to full EDA — histograms, trends, geo map, categoricals
53dc72b  Fix dashboard never rendering — remove st.stop() from folder tab
9222f07  Prevent preview crash on cloud when joining with auto-picked key
75e4b49  Rename to DataMerge Studio + replace custom strategies with SQL ops
b7b7f30  Fix AttributeError in clean_dtypes + persist results in session state
0f10c1a  Fix crash when re-merging a file that already has Source File column
171af6f  Remove hardcoded personal path from Folder Mode default
```

**Untracked files (intentionally not committed):**
- `Feedbacks to App/` — user feedback collection folder
- `Pending Call Analysis.pbix` — Power BI report (binary, kept out of git)
- `Prompt.docx` — drafting doc

---

## 3. Architecture & key decisions

### App tab structure
1. **📚 Learn Merge & Joins** — interactive examples with Venn diagrams
2. **📁 Upload Files** — drag-drop, has "🔄 New Session" button top-right
3. **🔀 Merge & Download** — 3-step flow: Column Alignment → Merge Settings → Sheet Groups
4. **📂 Folder Mode** — local-only (hidden on cloud); points at a folder of files
5. **📊 Dashboard** — auto-EDA from merged data, exportable as offline HTML
6. **🧹 Deduplicate** — single-file dedup (added in `f3cf588`)

### The 7 operations (replaced custom strategies entirely)

| Family | Operation | Pandas impl |
|--------|-----------|-------------|
| Union  | Union All | `pd.concat(join="outer")` |
| Union  | Union All (Distinct) | concat + `drop_duplicates` (with optional exclude-cols) |
| Join   | Inner Join | chained `pd.merge(how="inner")` |
| Join   | Left Join | chained `pd.merge(how="left")` |
| Join   | Right Join | chained `pd.merge(how="right")` |
| Join   | Full Outer Join | chained `pd.merge(how="outer")` |
| Join   | Cross Join | chained `pd.merge(how="cross")` with pre-check against 1,048,576 Excel row limit |

Multi-file joins chain left-to-right: `((A ⨝ B) ⨝ C) ⨝ D`.
Suffixes for collisions: `("", "__t2", "__t3", ...)`.

### Critical design decisions (don't undo these)
1. **No Smart Fill / no custom dedup strategies** — user explicitly rejected the old
   `Smart Fill (Best of Both)` as non-standard. Stay with pure pandas SQL operations.
2. **No "Source Sheet" column** ever inserted. Only `Source File` is added (if user ticks the box).
3. **`Source File` is preserved if already present** — when re-merging a previous output,
   the existing per-row source names are NOT overwritten with the new filename.
4. **Key column dropdown starts with placeholder** `— pick a key column —`.
   Preview will not auto-run for joins until user makes an explicit choice.
   This prevents many-to-many cartesian explosions OOM-killing the cloud container.
5. **`_dedup_columns()` is called on every input df** before `clean_dtypes()` runs.
   Defends against duplicate column names that make `df[col]` return a DataFrame
   instead of a Series (the `'DataFrame' has no attribute dtype'` bug).
6. **No fuzzy column matching** — user rejected it. Rely on the 10-row sample previews
   in the column-mapping UI instead.
7. **All merge results stored in `st.session_state["_last_run"]`** and rendered
   OUTSIDE the Run-button block so they survive widget interactions
   (downloading audit, etc.).
8. **Folder Mode is hidden entirely on cloud** (no `HAS_TKINTER`) so cloud users
   don't see a feature that can't work.

---

## 4. All file paths in this project

```
Pending Calls\
├── file_merger_app.py           # Main Streamlit app — 2,242 lines, in git
├── requirements.txt             # 7 lines, in git
├── run_app.bat                  # Windows launcher — 32 lines, in git
├── README.md                    # 194 lines, in git
├── Claude_Context.md            # THIS FILE, in git (commit before handing off)
├── merge_files.py               # Personal CLI script, gitignored
├── save_to_sql.py               # Personal SQL Server push, gitignored
├── Raw Files\                   # Source Excel files, gitignored
│   ├── ZOHO Pending Call Report as on 22-Apr-2026.xlsx
│   ├── ZOHO Pending Call Report as on 23-Apr-2026.xlsx
│   ├── ... (one per day, ~20 files)
│   └── ZOHO Pending Call Report as on 08-May-2026.xlsx
├── Feedbacks to App\            # User feedback, untracked
├── Pending Call Analysis.pbix   # Power BI, untracked
└── Prompt.docx                  # Drafting doc, untracked
```

---

## 5. Active blocker — RESOLVED (2026-05-11)

**Previous symptom:** `https://filemerger-kb.streamlit.app` was showing the
generic "Oh no. Error running app" page in the user's normal browser tab.

**Actual root cause:** **Browser cache.** The cloud deployment was healthy
and serving the latest code (Deduplicate tab visible in incognito and via
share.streamlit.io). The user's normal browser had cached the redacted "Oh no"
response from an earlier real crash (since fixed by `9222f07`) and was never
re-fetching from the server.

**Verification chain:**
- Yesterday: Deduplicate tab visible via share.streamlit.io and direct app URL → deploy was fine.
- Today: Deduplicate tab also visible in incognito → confirmed it's a per-session cache issue.
- Conclusion: there was never a deployment problem. Only browser caching.

**If this re-occurs in future:** tell the user to either
(a) hard-reload with `Ctrl + Shift + R` / `Ctrl + F5`, or
(b) clear site data: padlock icon → Site settings → Clear data, or
(c) test in an incognito window first to rule out cache before assuming server problem.

**Historical fixes that DID happen** and shipped in earlier commits this session:
- `9222f07` — key column placeholder + sample-size caps prevent OOM cartesian explosions on joins
- `b7b7f30` — defensive `clean_dtypes` + `_dedup_columns` prevent AttributeError on duplicate column names
- `0f10c1a` — Source File column preserved when re-merging a previous output

These are all live on the office repo and currently deployed to cloud.

---

## 6. Exact next steps (numbered, in priority order)

1. **Clean up `use_container_width` deprecations.** The cloud logs have ~hundreds
   of warnings like `Please replace use_container_width with width`. Streamlit
   said this becomes an error after 2025-12-31 — today is 2026-05-11, 5 months past.
   - Find all occurrences: `Grep "use_container_width" file_merger_app.py`
   - Replace `use_container_width=True` → `width="stretch"`
   - Replace `use_container_width=False` → `width="content"`
   - Bump `requirements.txt` to `streamlit>=1.46` (width param introduced in 1.46)
   - Test locally with `run_app.bat` before pushing.
   - Push from the **office machine** (where `github-company` SSH alias works);
     pushes from this Windows PC fail with `Could not resolve hostname github-company`.

2. **Address the `FutureWarning` about `pd.concat` with empty/all-NA entries.**
   Line 243 (inside `do_union_distinct`):
   ```python
   combined = pd.concat(dfs, ignore_index=True, join="outer")
   ```
   In pandas 3.x this will change behavior. Filter out empty dfs before concat:
   ```python
   dfs = [d for d in dfs if not d.empty]
   if not dfs: return pd.DataFrame(), pd.DataFrame()
   ```

3. **(Optional) Switch date parsing away from `pd.to_datetime(dayfirst=True)`** —
   the log shows `UserWarning: Could not infer format, falling back to dateutil`.
   Currently in `clean_dtypes` at line ~118. If files have mixed date formats,
   parsing is slow. Consider explicit format detection per column or letting
   pandas use ISO format only.

4. **(Optional) Verify the Deduplicate tab (commit `f3cf588`)** works end-to-end —
   it was added in a session I don't have memory of. Spot-check that:
   - It correctly handles files with `Source File` already present
   - It uses `_dedup_columns()` like the other merge paths
   - Excluding columns from the dedup check works

5. **(Optional) Fix local SSH so pushes from this Windows PC work.** Add to
   `~/.ssh/config`:
   ```
   Host github-company
       HostName github.com
       User git
       IdentityFile ~/.ssh/id_rsa
   ```
   After this, plain `git push` from this machine will reach the office repo
   without needing to log in to the office laptop.

---

## 7. Critical code snippets — common edit hotspots

### Adding `_dedup_columns()` defensively (helper)
Located near line 100 in `file_merger_app.py`:
```python
def _dedup_columns(df):
    """Return df with duplicate column names made unique (col → col_2, col_3...)."""
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
```
**Call it on every input df** before any column-by-column processing.

### The Source File preservation guard (used in two merge loops)
```python
d = _dedup_columns(df.copy())
if add_src:
    if "Source File" not in d.columns:   # ← critical: preserves existing values
        d.insert(0, "Source File", fname)
```

### Key-column placeholder pattern (Step 2 settings)
```python
if cfg["needs_key"] and all_cols:
    KEY_PLACEHOLDER = "— pick a key column —"
    key_choice = st.selectbox(
        "Key column", [KEY_PLACEHOLDER] + sorted(all_cols),
        key=f"key_{tab_key}")
    key_col = None if key_choice == KEY_PLACEHOLDER else key_choice
```

### Chained multi-file join
```python
def _chain_join(dfs, key, how):
    missing = [i + 1 for i, d in enumerate(dfs) if key not in d.columns]
    if missing:
        raise ValueError(f"Key column '{key}' missing in file/sheet #{missing}.")
    result = dfs[0].copy()
    for i, d in enumerate(dfs[1:], start=2):
        result = pd.merge(result, d, on=key, how=how, suffixes=("", f"__t{i}"))
    return result
```

### Preview row-count pre-check (prevents OOM)
Located inside `render_settings`:
```python
if cfg["family"] == "join" and key_col:
    est_rows = 1
    for d in sample_dfs:
        if key_col in d.columns:
            est_rows *= max(1, len(d))
    if est_rows > 50_000:
        st.warning(...)   # bail out, don't run the merge
        return cfg, key_col, excl_cols, clean_types, add_src, out_fmt
```

### Defensive `clean_dtypes`
Skips duplicate-named cols and non-string col names:
```python
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
    ...
```

---

## 8. How to verify everything is working

**Local sanity check (Windows):**
```
cd "C:\Users\Kshitij Buch\OneDrive\Documents\TBM 2026 Onwards\Pending Calls"
run_app.bat
```
Browser opens at `http://localhost:8501`. Upload 2 files, try each operation.

**Smoke test for merge functions (no Streamlit needed):**
```bash
python -c "
import pandas as pd, sys
sys.path.insert(0, '.')
# Replicate functions in isolation if module-level streamlit calls block import
# See conversation transcript for full snippet
"
```

**Streamlit Cloud verification:**
1. Open `https://filemerger-kb.streamlit.app` in incognito
2. Confirm the title reads **🔀 DataMerge Studio** (not "File Merger Pro")
3. Confirm the operation dropdown shows exactly: Union All, Union All (Distinct),
   Inner / Left / Right / Full Outer / Cross Join (NOT the old "1 — Simple Append" etc.)
4. Confirm a Left Join with no key column selected shows
   *"Pick a key column above to see a live preview"* instead of crashing

---

## 9. Things explicitly NOT to do

- **Don't add Smart Fill back.** User rejected it as non-standard.
- **Don't add fuzzy column matching (thefuzz / difflib).** User chose to rely
  on the 10-row sample previews instead.
- **Don't insert a `Source Sheet` column.** Removed per user feedback.
- **Don't auto-pick the first column as default key** for joins. Always require
  explicit user selection — the cartesian-explosion crash will return otherwise.
- **Don't downgrade `streamlit` below 1.46** if you migrate to `width="stretch"`.
- **Don't commit anything from `Raw Files/`** — `.gitignore` keeps personal data out.
- **Don't run `git push --force` to master.** Always create new commits.

---

## 10. User context (for tone & defaults)

- User is using the app for **TAT analysis of Pending Calls** (daily ZOHO + ERP exports).
- Multiple files per day in `Raw Files/`, named like `ZOHO Pending Call Report as on DD-MMM-YYYY.xlsx`.
- User prefers concise, action-oriented responses. Diagnose first, then fix.
- User runs **locally on Windows** but the hosted demo is for sharing with colleagues.
- User isn't a developer — explain trade-offs in plain language, not jargon.
- When asked to "build it" or "go ahead", do the full implementation + commit + push in one go.
- Always confirm what was pushed and where to verify it.

---

*End of handoff. Paste this entire file into a new Claude session and say
"Resume from Claude_Context.md" to pick up exactly where we left off.*
