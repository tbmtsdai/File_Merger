# File Merger Pro

A local web application for merging large CSV and Excel files — with full control over
how duplicates are identified and handled. Works with any files, any column structure.

---

## Who Is This For?

This is a **local application** — it runs entirely on your computer in a browser window.
No data is uploaded to any server. No internet connection is required after the first setup.
Anyone who has Python installed can run it — it is freely usable by any user for any files.

---

## Quick Start (Windows)

1. Make sure **Python 3.9+** is installed. Download from https://www.python.org/downloads/
   During install, check ✅ "Add Python to PATH".

2. Double-click **`run_app.bat`** — it installs the required packages and launches the app.

3. Your browser opens at **http://localhost:8501**

To stop the app, press `Ctrl+C` in the terminal window.

---

## Manual Setup (if run_app.bat does not work)

```
pip install pandas openpyxl xlrd streamlit
streamlit run file_merger_app.py
```

---

## Supported File Formats

| Format | Extension | Notes |
|--------|-----------|-------|
| Excel (modern) | `.xlsx` | Recommended |
| Excel (legacy) | `.xls` | Fully supported |
| CSV | `.csv` | Any delimiter; auto-detected |

Any number of columns, any column names — the app adapts to whatever files you upload.

---

## Application Tabs

### 1. Learn Merge Types
Live interactive examples using dummy employee data. Shows what each merge strategy
produces before you commit to running it on your real files.

### 2. Upload Files
Drag and drop 2–20 files. Preview the first 10 rows of each file and see row/column counts.

### 3. Merge & Download
Configure your merge strategy, run it, preview the result, and download as Excel or CSV.

### 4. Folder Mode
Point to a folder on your computer. The app auto-detects all CSV/Excel files in it.
Supports **incremental append** — if you ran a merge before, only new files are added
(see *Folder Mode* section below).

---

## Merge Strategies

### 1 — Simple Append
Stacks all files row-by-row. No filtering. Every row from every file is included,
including exact duplicates.

**Use when:** Files cover completely different records with no expected overlap
(e.g. different regions, different months with no carry-over).

---

### 2 — Remove Exact Duplicates  ← Default for this use case

Stacks all files, then removes rows where **every column has the exact same value**.

**Default behaviour:** ALL columns must match for a row to be considered a duplicate.
One column differing (even by a single character) means both rows are kept.

**Optional:** You can select specific columns to *exclude from the duplicate check*.
For example: exclude "Service Request Owner" so that the same pending call assigned to
two different engineers is still treated as one record (the engineer name difference
does not make it a new row).

**Use when:** Two system exports (e.g. ERP + CRM) may have the same events but minor
variations in non-key fields like assigned user, timestamps, or internal IDs.

---

### 3 — Key Dedup: First File Wins
When the same key value (e.g. SR Number, Employee ID, Invoice No.) appears in multiple
files, the version from the **first-uploaded** file is kept. All unique keys from all
files are included.

**Use when:** Your first file is the authoritative master source.

---

### 4 — Key Dedup: Last File Wins
Same as option 3 but the **last-uploaded** file's version wins for duplicate keys.

**Use when:** Your newest file has the most up-to-date data (e.g. updated status,
revised salary) and should overwrite older entries.

---

### 5 — Smart Fill (Best of Both)
For duplicate keys, assembles the output row by taking the **first non-empty value**
for each column across all files. Fills gaps in one file with data from another.
Upload your most authoritative file first.

**Use when:** Two systems each export partial data for the same records
(e.g. ERP has contract info, CRM has contact details). You want the most complete row.

---

### 6 — Intersection Only
Keeps **only** records whose key exists in every single uploaded file.
Records unique to any one file are excluded.

**Use when:** You need records confirmed across all sources — e.g. only calls that
appear in both ERP and CRM.

---

## Folder Mode — Appending New Files

When new files arrive in your folder over time, you do not need to re-merge everything.
The app tracks which files have already been processed via the **"Source File"** column
in the merged output.

**How it works:**
1. Run a full merge (all files, any strategy). The output file gets a "Source File" column.
2. Next week, new files arrive in the folder.
3. Open the app → Folder Mode tab → point to the same folder.
4. The app shows: "X files already processed, Y new files detected."
5. Choose **"Append only new files"** → only the new files are read and stacked below the
   existing merged output. The old data is untouched.
6. The merged file is updated and saved back to the same folder.

**Requirement:** The merged output must have a "Source File" column (the app adds this
automatically when "Add Source File column" is checked — which is the default).

---

## Pending Calls Files — Specific Script (merge_files.py)

For the specific set of ZOHO/ERP Pending Call files in this folder, a dedicated script
is provided: **`merge_files.py`**

Run it from the terminal:
```
python merge_files.py
```

**What it does differently from the general app:**

| Feature | General App | merge_files.py |
|---------|-------------|----------------|
| Duplicate logic | All columns same (user-configurable exclusions) | All common columns same, explicitly ignores "Service Request Owner" |
| April 25th | Choose your strategy | Auto-handled: ZOHO file (34 cols) as base, ERP file (30 cols) supplementary |
| Output | Download via browser | Saves directly to folder as MERGED_All_Pending_Calls.xlsx |
| Column normalization | Not needed (general) | Built-in: normalizes column name variants across files |

**Why the different duplicate logic for April 25th:**
The ERP file has 30 columns and the ZOHO file has 34 columns (4 extra contract-related
columns). When both files contain the same service call, the ERP row will have empty
contract columns. The script checks only the 29 columns that exist in both files
(excluding Service Request Owner), so it correctly identifies and removes 239 genuine
duplicates while keeping the ZOHO version which has richer data.

**April 25th result:** 629 (ZOHO) + 647 (ERP) − 239 duplicates = **1,037 unique records**

---

## Cross-File Duplicate Analysis (Pending Calls)

These files are **daily snapshots** of all open/pending service calls. A call that
remains unresolved will appear in every daily file until it is closed.

**Key findings from the duplicate check:**

| What it means | Count |
|---|---|
| Calls pending AND unchanged across multiple days | Appear as exact duplicate rows in multiple files |
| Calls pending but with status/data changes | Same SR Number, different row content — these are NOT duplicates |

**Exact identical rows found between files (sample):**

| File pair | Identical rows | Meaning |
|-----------|---------------|---------|
| 23-Apr-A × 23-Apr-B | 285 | Same day, same calls in both reports |
| 28-Apr × 29-Apr | 422 | 422 calls were pending with no change overnight |
| 24-Apr × 25-Apr-ZOHO | 357 | 357 calls were stuck from Apr 24 to Apr 25 |
| 28-Apr × 30-Apr | 276 | 276 calls unchanged across 2 days |

**This is normal for daily snapshot reports.** The `merge_files.py` script does NOT
remove these cross-date duplicates by design — each daily row is a valid data point
showing the call was still pending on that date. If you want a single-row-per-call
view showing the latest status, use the app with **Option 4 (Last File Wins)** on
SR Number as the key.

---

## SQL Server Integration (Personal Use)

To push the merged file into SQL Server, use the included `save_to_sql.py` script.

**One-time setup:**
```
pip install sqlalchemy pyodbc
```

**Configure the script** by editing the top section of `save_to_sql.py`:
- Server name
- Database name
- Table name
- Input file path

**Run:**
```
python save_to_sql.py
```

Alternatively, use the SSMS Import/Export Wizard:
`SSMS → Right-click database → Tasks → Import Data → Excel source`

---

## Files in This Package

| File | Purpose |
|------|---------|
| `file_merger_app.py` | Main Streamlit web application |
| `merge_files.py` | Specific script for the Pending Calls Excel files |
| `save_to_sql.py` | Push merged data to SQL Server (personal use) |
| `requirements.txt` | Python package dependencies |
| `run_app.bat` | Windows launcher — double-click to start the app |
| `README.md` | This file |

---

## Requirements

```
pandas >= 2.0
openpyxl >= 3.1
xlrd >= 2.0
streamlit >= 1.32
```

For SQL Server integration additionally:
```
sqlalchemy
pyodbc
```
