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

2. Double-click **`run_app.bat`** — it installs all required packages and launches the app.

3. Your browser opens at **http://localhost:8501**

To stop the app, press `Ctrl+C` in the terminal window.

---

## Manual Setup (if run_app.bat does not work)

```
pip install -r requirements.txt
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
produces — with a Venn diagram and a before/after table — before you commit to running
it on your real files.

### 2. Upload Files
Drag and drop 2–20 files. Preview the first rows of each sheet and see row/column counts.

### 3. Merge & Download
Three-step process:
- **Step 1 — Column Alignment:** if files have different column names, a mapping form
  appears so you can assign a canonical name to columns that mean the same thing
  (e.g. "Customer Name" and "Client Name" → both become "Customer Name").
- **Step 2 — Sheet Groups:** sheets with identical column sets are automatically grouped
  for merging; sheets with different structures get separate output sheets.
- **Step 3 — Merge Settings:** choose your strategy, run the merge, preview results,
  download as Excel or CSV.

After every merge you also see:
- A **Data Type Cleaning report** listing every date parsed, whitespace stripped, or
  city/region column title-cased.
- A **Duplicate Audit** panel showing the exact row numbers of every removed duplicate,
  with a downloadable CSV report.

### 4. Folder Mode
Point to a folder on your computer. The app auto-detects all CSV/Excel files in it.
Supports **incremental append** — if you ran a merge before, only new files are added.
Has the same column mapping, type cleaning, and duplicate audit as the Upload tab.

### 5. Dashboard
Auto-generates charts from the merged data immediately after any merge:
- Records by source file (bar chart)
- Date trends (line chart — any detected date column, groupable by day/week/month)
- Categorical breakdowns (horizontal bar charts for city, region, zone, etc.)
- Numeric column summary table
- Missing value % by column
- Filterable data explorer

**Export as HTML:** generates a fully self-contained `.html` file that works offline
in any browser — no Streamlit, no internet needed. Share it as a report.

---

## Column Alignment — How It Works

When you upload files where the same data has different column names
(e.g. "SR No" in one file, "Ticket ID" in another), the app detects the mismatch
and shows a mapping form:

```
`SR No`     ✅ File A   ❌ missing: File B    Canonical name: [SR No      ]
`Ticket ID` ✅ File B   ❌ missing: File A    Canonical name: [SR No      ]  ← user types same name
```

Type the same canonical name for both → they are renamed and treated as one column
before grouping and merging. Leave names unchanged to keep them as separate columns.

---

## Data Type Cleaning

Enabled by default (checkbox in Merge Settings). Applied automatically during merge:

| Column type | What happens |
|-------------|-------------|
| Column name contains: `date`, `time`, `created`, `due`, `closed`, `updated`, etc. | Parsed to datetime using `pd.to_datetime` (dayfirst) — only if ≥50% of values parse successfully |
| Column name contains: `city`, `town`, `region`, `zone`, `area`, `state`, `country` | Whitespace stripped + title-cased (`mumbai` → `Mumbai`) |
| Any other string column | Whitespace stripped only |

A report after the merge shows exactly which columns were changed and how many values parsed.

---

## Duplicate Audit

After every merge with a dedup strategy, the app shows:
- **Count** of rows removed
- **Table** of every removed row, with a `Removed Row# (Excel)` column — the 1-based row
  number in the combined pre-dedup dataset (row 2 = first data row, matching Excel's
  header-on-row-1 convention)
- **Download button** for the full audit as a CSV (`duplicate_audit.csv`)

---

## Merge Strategies

### 1 — Simple Append
Stacks all files row-by-row. No filtering. Every row from every file is included,
including exact duplicates.

**Use when:** Files cover completely different records with no expected overlap.

---

### 2 — Remove Exact Duplicates
Stacks all files, then removes rows where **every column has the exact same value**.

**Optional:** Exclude specific columns from the duplicate check. For example: exclude
"Service Request Owner" so the same pending call assigned to two different engineers
is still treated as one record.

**Use when:** Two system exports may have the same events but minor field differences.

---

### 3 — Key Dedup: First File Wins
When the same key value appears in multiple files, the version from the **first-uploaded**
file is kept. All unique keys from all files are included.

**Use when:** Your first file is the authoritative master source.

---

### 4 — Key Dedup: Last File Wins
Same as option 3 but the **last-uploaded** file's version wins for duplicate keys.

**Use when:** Your newest file has the most up-to-date data.

---

### 5 — Smart Fill (Best of Both)
For duplicate keys, assembles the output row by taking the **first non-empty value**
for each column across all files. Fills gaps in one file with data from another.

**Use when:** Two systems each export partial data for the same records.

---

### 6 — Intersection Only
Keeps **only** records whose key exists in every single uploaded file.
Records unique to any one file are excluded.

**Use when:** You need records confirmed across all sources.

---

## Folder Mode — Appending New Files

When new files arrive in your folder over time, you do not need to re-merge everything.
The app tracks which files have already been processed via the **"Source File"** column
in the merged output.

1. Run a full merge (all files). The output gets a "Source File" column.
2. New files arrive in the folder next week.
3. Open the app → Folder Mode tab → point to the same folder.
4. The app shows: "X files already processed, Y new files detected."
5. Choose **"Append only new files"** — only new files are stacked below the existing data.
6. The merged file is updated and saved back to the same folder.

**Requirement:** "Add Source File column" must be checked (it is by default).

---

## Pending Calls — Specific Script (`merge_files.py`)

For the specific set of ZOHO/ERP Pending Call files in this folder, a dedicated
command-line script is also provided.

Run from terminal:
```
python merge_files.py
```

| Feature | General App | merge_files.py |
|---------|-------------|----------------|
| Duplicate logic | User-configurable | Hardcoded: ignores "Service Request Owner" |
| ERP file handling | Separate group auto-detected | Explicitly labelled `_ERP` sheets |
| Output | Download via browser | Saves directly to folder |
| Column normalization | Column Alignment UI | Built-in alias map |

---

## SQL Server Integration (`save_to_sql.py`)

Push any merged Excel/CSV file into a SQL Server table.

**One-time setup:**
```
pip install sqlalchemy pyodbc
```

Edit the configuration block at the top of `save_to_sql.py` (server name, database,
table name, input file path), then run:
```
python save_to_sql.py
```

Alternatively, use the SSMS Import/Export Wizard:
`SSMS → Right-click database → Tasks → Import Data → Excel source`

---

## Files in This Package

| File | Purpose | In git? |
|------|---------|---------|
| `file_merger_app.py` | Main Streamlit web application | ✅ Yes |
| `requirements.txt` | Python package dependencies | ✅ Yes |
| `run_app.bat` | Windows launcher — double-click to start | ✅ Yes |
| `README.md` | This file | ✅ Yes |
| `merge_files.py` | Specific script for Pending Calls files (has personal paths) | ❌ Gitignored |
| `save_to_sql.py` | Push merged data to SQL Server (has personal config) | ❌ Gitignored |
| `Raw Files/` | Source Excel files | ❌ Gitignored |
| `MERGED_*.xlsx` | Generated output files | ❌ Gitignored |
| `__pycache__/` | Python bytecode cache | ❌ Gitignored |

---

## Requirements

```
pandas >= 2.0
openpyxl >= 3.1
xlrd >= 2.0
streamlit >= 1.32
matplotlib >= 3.7
matplotlib-venn >= 0.11
plotly >= 5.0        ← for interactive dashboard charts and HTML export
```

For SQL Server integration additionally:
```
sqlalchemy
pyodbc
```
