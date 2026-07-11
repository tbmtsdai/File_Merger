# CLAUDE.md — File Merger & Join App

Auto-loaded project context. Read this first; see `Claude_Context.md` for deep
architecture detail and code hotspots.

## What this is
**DataMerge Studio** — a Streamlit web app (`file_merger_app.py`, ~2,300 lines,
Python + pandas) that merges/joins CSV & Excel files using the 7 standard SQL
set operations (Union All, Union All Distinct, Inner/Left/Right/Full Outer/Cross
Join), with column mapping, data-type cleaning, a Deduplicate tab, and an
auto-EDA dashboard. The user (non-developer) uses it for **TAT analysis of daily
ZOHO Pending Call reports**, but the app itself is generic.

## Where things live (IMPORTANT)
- **Repo (code + git):** `C:\Users\k.buch\Documents\file merger and join app`.
  Run locally via `run_app.bat` → http://localhost:8501.
- **Raw Files (Excel sources + merged output):** kept in OneDrive at
  `C:\Users\k.buch\OneDrive - Transasia Bio Medicals Ltd\TBM 2026 Onwards\Pending Calls\Raw Files`.
  Gitignored — never part of the repo. The app doesn't need it (upload/browse
  any folder); nothing in the app hardcodes it.
- **`docs/`** (gitignored): user's Power BI report (`.pbix`) + draft docs. Local only.

## 🔴 CRITICAL: never move this repo back into OneDrive
OneDrive dehydrates `.git` into cloud-only placeholders, which corrupts git:
read failures ("cloud file provider is not running"), `mmap failed` on push, and
on 2026-07-11 it silently **reverted `master` and dropped a commit**. Keep the
repo on a non-synced local path (Documents is NOT redirected to OneDrive here).
Do NOT restart OneDrive to "fix" hydration — that restart caused the revert.

## Cloud deployment
- Live app: **https://tbmfilemergejoin.streamlit.app**
- Deploys from office repo **`origin = git@github-company:tbmtsdai/File_Merger.git`**.
  The `github-company` SSH alias **works from this PC** — `git push origin master`
  pushes and triggers auto-redeploy. (`personal = github.com:kshitijbuch/File_Merger`
  is a backup mirror; cloud does NOT deploy from it.)
- **Python version:** set to **3.12** in Streamlit Cloud → Manage app → Settings →
  General → Python. Do NOT switch to 3.14 (no wheels for pinned pandas/numpy → source
  builds time out; also involved in the 2026-07-11 pandas 3.0 segfault).
- **Streamlit Cloud caps RAM at ~1 GB.** OOM manifests as `Killed` in logs.
  Large-merge Excel export uses XlsxWriter, not openpyxl (`_make_excel_writer`).
- **Pinned dependency ceilings** in `requirements.txt` — DO NOT relax without testing:
  `pandas<3`, `numpy<2.3`, `pyarrow<20`, plus matching ceilings elsewhere. Without
  these, uv pulls latest-of-everything (pandas 3.0, pyarrow 25, etc.) which caused
  two separate cloud crashes on 2026-07-11.

## 🔴 Distinguishing OOM from segfault in cloud logs
Both fail the app but need different fixes — do NOT conflate them.
- **OOM** → log line contains **`Killed`** (SIGKILL). Fix: reduce peak memory
  (that's what XlsxWriter did — 944 MB peak → 307 MB).
- **Segfault** → log line contains **`Segmentation fault`** (SIGSEGV). This is a
  native crash in a C extension, NOT memory. Fix: pin the offending native library
  to an older stable version. On 2026-07-11, `pyarrow 25.0.0` segfaulted right
  after `st.dataframe(...)` render of the merged DataFrame; pinning `pyarrow<20`
  resolved it (Streamlit uses pyarrow to serialize DataFrames for display).

## Design decisions — do NOT undo
1. No "Smart Fill" / no custom dedup strategies (user rejected as non-standard).
2. No fuzzy column matching (rely on 10-row sample previews).
3. Never insert a `Source Sheet` column; only `Source File`, and preserve it if
   already present (don't overwrite per-row source when re-merging output).
4. Never auto-pick a join key — always require explicit selection (prevents
   cartesian-explosion OOM). Key dropdown starts with a placeholder.
5. `_dedup_columns()` runs on every input df before `clean_dtypes()`.
6. Folder Mode is gated on `HAS_TKINTER`: visible/works locally, hidden on cloud.
7. Excel export goes through `_make_excel_writer` (XlsxWriter, openpyxl fallback).
8. Never commit anything from `Raw Files/`. Never force-push master.

## Recent work (2026-07-11)
- **Fixed cloud OOM crash** on large merges: XlsxWriter instead of openpyxl
  (commit 387d7b8). Also makes downloads smaller — the old 16 MB→10 MB shrink on
  Excel re-save is gone; XlsxWriter writes the lean file directly (~9.5 MB).
- **Dedup tab enhancement** (commit 4ec08f0): for each removed duplicate, shows
  which kept row it matches, both rows' Source File, a same/different-file flag,
  and a source-file pairing summary — so the user can tell if duplicates are the
  same dated file uploaded twice vs. genuinely repeated records across files.
- **Repo moved out of OneDrive** → `Documents\file merger and join app`.
- **Two cloud segfaults debugged (same day):**
  1. `requirements.txt` had no upper bounds → uv pulled `pandas==3.0.3` on
     Python 3.14. Fixed by adding ceilings (commit f35a94b) and setting
     Streamlit Cloud's Python to 3.12.
  2. That surfaced the real culprit — `pyarrow==25.0.0` still segfaulted after
     merge. Fixed by pinning `pyarrow<20` (commit 36a0ac7).
- All fixes pushed to origin; cloud confirmed working after (2).

## Known noise in cloud logs (not errors)
Suppress the urge to "fix" these mid-debug — they're pre-existing and irrelevant
to any live incident:
- Many `Please replace use_container_width with width` deprecation lines —
  Streamlit deprecation, becomes error eventually. Cleanup deferred (see
  Claude_Context.md §6).
- `FutureWarning: The behavior of DataFrame concatenation with empty or all-NA
  entries is deprecated` at line 258 (`do_union_distinct`) — pandas 3.x will
  change behavior; filter empty dfs before concat when this is addressed.
- `UserWarning: Could not infer format ... falling back to dateutil` from
  `clean_dtypes` — slow but correct; only worth fixing if merge time becomes
  a complaint.

## User context
Non-developer. Prefers concise, action-oriented replies — diagnose first, then
fix. Works locally on Windows; hosted demo is for sharing. When told "go ahead",
do the full implementation + commit + push, then confirm what was pushed and
where to verify.
