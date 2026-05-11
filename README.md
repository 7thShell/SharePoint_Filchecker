# SharePoint FileChecker

> PowerShell tool that scans folders for OneDrive/SharePoint naming and path compliance, generating an interactive HTML report.

---

## Overview

Before migrating files to SharePoint Online or OneDrive, paths and filenames must meet strict naming rules. SharePoint FileChecker scans any folder recursively and flags every file that would cause a migration failure — along with early warnings for files approaching the limits.

Results are written to a self-contained, interactive HTML report that opens automatically in your default browser.

---

## Compliance Checks

| # | Rule | Threshold | Status |
|---|------|-----------|--------|
| 1 | Full path length | > 255 characters | FAIL |
| 2 | Full path length approaching limit | 201 – 255 characters | WARN |
| 3 | Filename length | > 128 characters | FAIL |
| 4 | Folder name length | > 128 characters | FAIL |
| 5 | Forbidden characters in filename or folder | `* : " < > ? / \ \|` | FAIL |
| 6 | Reserved Windows device names | CON, PRN, AUX, NUL, COM1-9, LPT1-9 | FAIL |
| 7 | Filename starts or ends with space or period | | FAIL |
| 8 | Folder name starts or ends with space or period | | FAIL |

---

## Report Features

- **Three status levels** — OK (green), WARN (yellow), FAIL (red)
- **Summary cards** showing total files, compliant, warnings, and issues found
- **Filter buttons** — show All / OK / Warnings / Issues only
- **Live search** — filter by filename or full path
- **Sortable columns** — click any column header
- **Dark / light mode** — follows your OS preference
- **Self-contained HTML** — no server or internet connection required
- **Sequential naming** — reports are saved as `SharePoint_FileCheck_001.html`, `_002.html`, etc.

---

## Requirements

- Windows PowerShell 5.1 or later
- Windows 10 / 11

---

## Usage

### Option 1 — Double-click launcher

Run `SharePoint_FileChecker.cmd`. A dialog opens where you select the folder to scan.

### Option 2 — PowerShell with a path

```powershell
.\SharePoint_FileChecker.ps1 -FolderPath "C:\Users\Alice\Documents"
```

### Option 3 — PowerShell interactive

```powershell
.\SharePoint_FileChecker.ps1
```

A folder-browser dialog appears for you to select the target folder.

The HTML report is saved next to the script and opens automatically when the scan completes.

---

## Status Levels

| Badge | Meaning |
|-------|---------|
| ✅ OK | File meets all SharePoint requirements |
| ⚠️ WARN | Path is between 201–255 characters — consider shortening before migration |
| ❌ FAIL | File has one or more violations that must be resolved before migration |
