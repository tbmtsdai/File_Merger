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
- Streamlit Cloud caps RAM at **~1 GB**. Large merges (tens of thousands of rows)
  OOM-crash if Excel is built with openpyxl → **use XlsxWriter** (`_make_excel_writer`).

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
- Repo moved out of OneDrive; all pushed to origin.

## User context
Non-developer. Prefers concise, action-oriented replies — diagnose first, then
fix. Works locally on Windows; hosted demo is for sharing. When told "go ahead",
do the full implementation + commit + push, then confirm what was pushed and
where to verify.
