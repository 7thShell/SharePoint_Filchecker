#Requires -Version 5.1
<#
.SYNOPSIS
    Scans a folder for OneDrive/SharePoint file compliance and outputs an interactive HTML report.

.DESCRIPTION
    Check-CloudCompliance.ps1 recursively scans a given folder and checks every file against
    the naming and path rules enforced by OneDrive and SharePoint. Violations detected:

      - Full path length exceeding 255 characters
      - Filename length exceeding 128 characters
      - Folder name length exceeding 128 characters
      - Forbidden characters (* : " < > ? / \ |) in file or folder names along the path
      - Reserved Windows device names (CON, PRN, AUX, NUL, COM1-COM9, LPT1-LPT9)
      - File or folder names that start or end with a space or period

    Results are written to a fully self-contained, interactive HTML report saved next to this
    script. The report supports dark/light mode, column sorting, live search, and filter buttons.

.PARAMETER FolderPath
    The root folder to scan. If omitted a Windows folder-picker dialog is shown.

.EXAMPLE
    .\Check-CloudCompliance.ps1 -FolderPath "C:\Users\Alice\Documents"

    Scans C:\Users\Alice\Documents recursively and opens the HTML report in the default browser.

.EXAMPLE
    .\Check-CloudCompliance.ps1

    Opens a folder-browser dialog, then scans and opens the report.
#>
[CmdletBinding()]
param (
    [Parameter(Position = 0)]
    [string]$FolderPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
Add-Type -AssemblyName System.Drawing       -ErrorAction Stop

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Show-LauncherGui {
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # ---- Form ----
    $form                  = New-Object System.Windows.Forms.Form
    $form.Text             = 'Cloud Compliance Scanner'
    $form.ClientSize       = [System.Drawing.Size]::new(520, 152)
    $form.StartPosition    = 'CenterScreen'
    $form.FormBorderStyle  = 'FixedDialog'
    $form.MaximizeBox      = $false
    $form.MinimizeBox      = $false

    # ---- Label ----
    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = 'Select folder to scan for OneDrive / SharePoint compliance:'
    $lbl.Location = [System.Drawing.Point]::new(16, 18)
    $lbl.Size     = [System.Drawing.Size]::new(488, 18)
    $lbl.Font     = New-Object System.Drawing.Font('Segoe UI', 9)

    # ---- Path textbox ----
    $txt           = New-Object System.Windows.Forms.TextBox
    $txt.Location  = [System.Drawing.Point]::new(16, 42)
    $txt.Size      = [System.Drawing.Size]::new(386, 26)
    $txt.ReadOnly  = $true
    $txt.BackColor = [System.Drawing.SystemColors]::Window
    $txt.ForeColor = [System.Drawing.Color]::Gray
    $txt.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    $txt.Text      = 'No folder selected'

    # ---- Browse button ----
    $btnBrowse          = New-Object System.Windows.Forms.Button
    $btnBrowse.Text     = 'Browse...'
    $btnBrowse.Location = [System.Drawing.Point]::new(410, 41)
    $btnBrowse.Size     = [System.Drawing.Size]::new(94, 28)
    $btnBrowse.Font     = New-Object System.Drawing.Font('Segoe UI', 9)

    # ---- Scan button ----
    $btnScan                           = New-Object System.Windows.Forms.Button
    $btnScan.Text                      = 'Scan'
    $btnScan.Location                  = [System.Drawing.Point]::new(195, 96)
    $btnScan.Size                      = [System.Drawing.Size]::new(130, 36)
    $btnScan.Font                      = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $btnScan.Enabled                   = $false
    $btnScan.BackColor                 = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnScan.ForeColor                 = [System.Drawing.Color]::White
    $btnScan.FlatStyle                 = 'Flat'
    $btnScan.FlatAppearance.BorderSize = 0

    # ---- Events ----
    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description         = 'Select the folder to scan'
        $dlg.ShowNewFolderButton = $false
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txt.Text      = $dlg.SelectedPath
            $txt.ForeColor = [System.Drawing.SystemColors]::WindowText
            $btnScan.Enabled = $true
        }
        $dlg.Dispose()
    })

    $btnScan.Add_Click({
        $form.Tag          = $txt.Text
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $form.Controls.AddRange(@($lbl, $txt, $btnBrowse, $btnScan))
    $form.AcceptButton = $btnScan

    $res      = $form.ShowDialog()
    $selected = $form.Tag
    $form.Dispose()

    if ($res -ne [System.Windows.Forms.DialogResult]::OK -or -not $selected) {
        Write-Host 'No folder selected - exiting.' -ForegroundColor Yellow
        exit 0
    }
    return $selected
}

function Test-ReservedName {
    param ([string]$Name)
    $reserved = @('CON','PRN','AUX','NUL',
                  'COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9',
                  'LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9')
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Name).ToUpperInvariant()
    return $reserved -contains $baseName
}

function Get-ComplianceIssues {
    param (
        [System.IO.FileInfo]$File,
        [string]$RootPath
    )

    $issues   = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $fullPath = $File.FullName
    $fileName = $File.Name

    # 1. Full path length
    if ($fullPath.Length -gt 255) {
        $issues.Add("Path length $($fullPath.Length) > 255 chars")
    } elseif ($fullPath.Length -gt 200) {
        $warnings.Add("Path length $($fullPath.Length) approaching limit (> 200 chars)")
    }

    # 2. Filename length
    if ($fileName.Length -gt 128) {
        $issues.Add("Filename length $($fileName.Length) > 128 chars")
    }

    # 3. Forbidden characters - check filename AND every folder segment
    $forbiddenPattern = '[*:"<>?/\\|]'
    if ($fileName -match $forbiddenPattern) {
        $issues.Add("Forbidden character in filename")
    }

    # Build folder segments between root and file
    $relativePath  = $fullPath.Substring($RootPath.TrimEnd('\').Length).TrimStart('\')
    $segments      = $relativePath -split '\\' | Select-Object -SkipLast 1  # exclude filename
    foreach ($seg in $segments) {
        if ($seg -match $forbiddenPattern) {
            $issues.Add("Forbidden character in folder '$seg'")
            break
        }
    }

    # 4. Reserved Windows names (filename)
    if (Test-ReservedName -Name $fileName) {
        $issues.Add("Reserved Windows filename '$([System.IO.Path]::GetFileNameWithoutExtension($fileName))'")
    }

    # 5. Leading/trailing space or period - filename
    if ($fileName -match '^\s|^\.|\.(\s*)$|\s$') {
        $issues.Add("Filename starts or ends with space/period")
    }

    # 5b. Leading/trailing space or period - folder segments
    foreach ($seg in $segments) {
        if ($seg -match '^\s|^\.|\.(\s*)$|\s$') {
            $issues.Add("Folder '$seg' starts or ends with space/period")
            break
        }
    }

    # 6. Folder name length exceeding 128 characters
    foreach ($seg in $segments) {
        if ($seg.Length -gt 128) {
            $issues.Add("Folder name length $($seg.Length) > 128 chars: '$($seg.Substring(0,40))...'")
        }
    }

    return [PSCustomObject]@{
        Issues   = $issues
        Warnings = $warnings
    }
}

function ConvertTo-HtmlEncoded {
    param ([string]$Text)
    return [System.Web.HttpUtility]::HtmlEncode($Text)
}

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if (-not $FolderPath) {
    $FolderPath = Show-LauncherGui
} else {
    $FolderPath = $FolderPath.Trim('"').Trim("'")
    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        Write-Error "Folder not found: $FolderPath"
        exit 1
    }
}

$FolderPath = (Resolve-Path -LiteralPath $FolderPath).Path
$scanStart  = Get-Date

Write-Host "`nScanning: $FolderPath" -ForegroundColor Cyan

# Collect files
try {
    $allFiles = @(Get-ChildItem -LiteralPath $FolderPath -Recurse -File -ErrorAction SilentlyContinue -ErrorVariable enumErrors)
} catch {
    Write-Error "Cannot access folder: $_"
    exit 1
}

if ($enumErrors) {
    foreach ($e in $enumErrors) {
        Write-Warning "Access denied (skipped): $($e.TargetObject)"
    }
}

if ($allFiles.Count -eq 0) {
    Write-Warning "No files found in $FolderPath"
}

# Load System.Web for HtmlEncode (available in .NET Framework / PS 5.1)
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

$results   = [System.Collections.Generic.List[PSCustomObject]]::new()
$total     = $allFiles.Count
$index     = 0

foreach ($file in $allFiles) {
    $index++
    Write-Progress -Activity "Checking compliance" `
                   -Status "$index of $total - $($file.Name)" `
                   -PercentComplete ([Math]::Round($index / [Math]::Max($total,1) * 100))

    $issues   = @()
    $warnings = @()
    try {
        $result   = Get-ComplianceIssues -File $file -RootPath $FolderPath
        $issues   = @($result.Issues)
        $warnings = @($result.Warnings)
    } catch {
        $issues = @("Error reading file: $_")
    }

    $status = if ($issues.Count -gt 0) { 'FAIL' } elseif ($warnings.Count -gt 0) { 'WARN' } else { 'OK' }

    $results.Add([PSCustomObject]@{
        FullPath       = $file.FullName
        FileName       = $file.Name
        FullPathLength = $file.FullName.Length
        FileNameLength = $file.Name.Length
        Status         = $status
        Issues         = ($issues -join '; ')
        Warnings       = ($warnings -join '; ')
    })
}

Write-Progress -Activity "Checking compliance" -Completed

$okCount   = @($results | Where-Object { $_.Status -eq 'OK'   }).Count
$warnCount = @($results | Where-Object { $_.Status -eq 'WARN' }).Count
$failCount = @($results | Where-Object { $_.Status -eq 'FAIL' }).Count
$scanEnd   = Get-Date
$duration  = ($scanEnd - $scanStart).TotalSeconds

Write-Host "Scan complete in $([Math]::Round($duration,1))s - $total files, $okCount OK, $warnCount warnings, $failCount issues" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Build HTML
# ---------------------------------------------------------------------------

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path

$existingNums = Get-ChildItem -LiteralPath $scriptDir -Filter 'SharePoint_FileCheck_*.html' |
    ForEach-Object { if ($_.BaseName -match '_(\d+)$') { [int]$Matches[1] } } |
    Measure-Object -Maximum
$nextNum    = if ($existingNums.Count -gt 0) { $existingNums.Maximum + 1 } else { 1 }
$reportPath = Join-Path $scriptDir ("SharePoint_FileCheck_{0:D3}.html" -f $nextNum)

# Build table rows
$rowsHtml = [System.Text.StringBuilder]::new()
foreach ($r in $results) {
    $statusClass  = switch ($r.Status) { 'OK' { 'ok' } 'WARN' { 'warn' } default { 'fail' } }
    $escapedPath  = [System.Web.HttpUtility]::HtmlEncode($r.FullPath)
    $escapedName  = [System.Web.HttpUtility]::HtmlEncode($r.FileName)
    $escapedIssue = [System.Web.HttpUtility]::HtmlEncode($r.Issues)
    $escapedWarn  = [System.Web.HttpUtility]::HtmlEncode($r.Warnings)

    $issueCell = if ($r.Status -eq 'FAIL') {
        "<td class='issues-cell'><span class='issues-text'>$escapedIssue</span></td>"
    } elseif ($r.Status -eq 'WARN') {
        "<td class='issues-cell'><span class='issues-warn'>$escapedWarn</span></td>"
    } else {
        "<td class='issues-cell'><span class='issues-none'>&mdash;</span></td>"
    }

    [void]$rowsHtml.AppendLine(
        "<tr class='row-$statusClass' data-path='$escapedPath' data-name='$escapedName'>" +
        "<td class='col-name' title='$escapedPath'>$escapedName</td>" +
        "<td class='col-path path-cell'>$escapedPath</td>" +
        "<td class='col-pathlen num'>$($r.FullPathLength)</td>" +
        "<td class='col-namelen num'>$($r.FileNameLength)</td>" +
        "<td class='col-status'><span class='badge badge-$statusClass'>$($r.Status)</span></td>" +
        $issueCell +
        "</tr>"
    )
}

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Cloud Compliance Report</title>
<style>
/* ===== Reset & tokens ===== */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --bg:          #f4f6fb;
  --surface:     #ffffff;
  --surface2:    #f0f2f8;
  --border:      #dde1eb;
  --text:        #1a1d2e;
  --text-muted:  #6b7280;
  --accent:      #4f6ef7;
  --ok-bg:       #ecfdf5;
  --ok-text:     #065f46;
  --ok-badge:    #d1fae5;
  --warn-bg:     #fffbeb;
  --warn-text:   #92400e;
  --warn-badge:  #fef3c7;
  --fail-bg:     #fff1f2;
  --fail-text:   #9f1239;
  --fail-badge:  #ffe4e6;
  --hover:       #eef0fb;
  --shadow:      0 2px 12px rgba(0,0,0,.08);
  --radius:      10px;
  --font:        'Segoe UI', system-ui, -apple-system, sans-serif;
}

@media (prefers-color-scheme: dark) {
  :root {
    --bg:          #0f1117;
    --surface:     #1a1d2a;
    --surface2:    #22263a;
    --border:      #2e3248;
    --text:        #e8eaf6;
    --text-muted:  #8b91a8;
    --accent:      #7b93ff;
    --ok-bg:       #052e1e;
    --ok-text:     #6ee7b7;
    --ok-badge:    #064e33;
    --warn-bg:     #2d1f02;
    --warn-text:   #fcd34d;
    --warn-badge:  #451a03;
    --fail-bg:     #2d0a10;
    --fail-text:   #fca5a5;
    --fail-badge:  #4c1525;
    --hover:       #252a40;
    --shadow:      0 2px 16px rgba(0,0,0,.4);
  }
}

body {
  background: var(--bg);
  color: var(--text);
  font-family: var(--font);
  font-size: 14px;
  line-height: 1.5;
  padding: 24px 16px 48px;
}

/* ===== Header ===== */
.header {
  max-width: 1400px;
  margin: 0 auto 24px;
}
.header h1 {
  font-size: 1.6rem;
  font-weight: 700;
  letter-spacing: -.5px;
  color: var(--accent);
  margin-bottom: 4px;
}
.header .subtitle {
  color: var(--text-muted);
  font-size: .85rem;
}

/* ===== Summary cards ===== */
.summary {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
  gap: 14px;
  max-width: 1400px;
  margin: 0 auto 24px;
}
.card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 16px 20px;
  box-shadow: var(--shadow);
}
.card .label {
  font-size: .72rem;
  text-transform: uppercase;
  letter-spacing: .08em;
  color: var(--text-muted);
  margin-bottom: 6px;
}
.card .value {
  font-size: 1.75rem;
  font-weight: 700;
  line-height: 1;
}
.card.ok   .value { color: var(--ok-text);   }
.card.warn .value { color: var(--warn-text); }
.card.fail .value { color: var(--fail-text); }
.card .meta {
  font-size: .78rem;
  color: var(--text-muted);
  margin-top: 4px;
  word-break: break-all;
}

/* ===== Toolbar ===== */
.toolbar {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  align-items: center;
  max-width: 1400px;
  margin: 0 auto 14px;
}
.filter-group { display: flex; gap: 6px; }
.btn {
  padding: 6px 14px;
  border-radius: 6px;
  border: 1px solid var(--border);
  background: var(--surface);
  color: var(--text);
  cursor: pointer;
  font-size: .82rem;
  font-family: var(--font);
  transition: background .15s, color .15s, border-color .15s;
}
.btn:hover { background: var(--hover); }
.btn.active {
  background: var(--accent);
  color: #fff;
  border-color: var(--accent);
}
#btnWarn.active {
  background: var(--warn-text);
  border-color: var(--warn-text);
}
.search-wrap { flex: 1; min-width: 200px; }
.search-wrap input {
  width: 100%;
  padding: 7px 12px;
  border-radius: 6px;
  border: 1px solid var(--border);
  background: var(--surface);
  color: var(--text);
  font-family: var(--font);
  font-size: .85rem;
  outline: none;
  transition: border-color .15s;
}
.search-wrap input:focus { border-color: var(--accent); }
.count-label {
  color: var(--text-muted);
  font-size: .82rem;
  white-space: nowrap;
}

/* ===== Table ===== */
.table-wrap {
  max-width: 1400px;
  margin: 0 auto;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  overflow-x: auto;
}
table {
  width: 100%;
  border-collapse: collapse;
  table-layout: auto;
}
thead th {
  background: var(--surface2);
  padding: 10px 12px;
  text-align: left;
  font-size: .78rem;
  text-transform: uppercase;
  letter-spacing: .06em;
  color: var(--text-muted);
  cursor: pointer;
  user-select: none;
  white-space: nowrap;
  border-bottom: 1px solid var(--border);
  position: sticky;
  top: 0;
  z-index: 1;
}
thead th:hover { color: var(--text); }
thead th .sort-icon { margin-left: 4px; opacity: .4; }
thead th.asc  .sort-icon::after { content: ' \25B2'; opacity: 1; }
thead th.desc .sort-icon::after { content: ' \25BC'; opacity: 1; }
thead th:not(.asc):not(.desc) .sort-icon::after { content: ' \21C5'; }

tbody tr {
  border-bottom: 1px solid var(--border);
  transition: background .1s;
}
tbody tr:last-child { border-bottom: none; }
tbody tr:hover { background: var(--hover); }

tbody tr.row-ok   { background: var(--ok-bg);   }
tbody tr.row-ok:hover   { filter: brightness(.97); }
tbody tr.row-warn { background: var(--warn-bg); }
tbody tr.row-warn:hover { filter: brightness(.97); }
tbody tr.row-fail { background: var(--fail-bg); }
tbody tr.row-fail:hover { filter: brightness(.97); }

td {
  padding: 8px 12px;
  vertical-align: top;
  font-size: .84rem;
}
.col-name  { max-width: 260px; word-break: break-word; font-weight: 500; }
.col-path  { max-width: 440px; word-break: break-all; font-size: .78rem; color: var(--text-muted); }
.num       { text-align: right; font-variant-numeric: tabular-nums; }

.badge {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 20px;
  font-size: .75rem;
  font-weight: 600;
  letter-spacing: .04em;
}
.badge-ok   { background: var(--ok-badge);   color: var(--ok-text);   }
.badge-warn { background: var(--warn-badge); color: var(--warn-text); }
.badge-fail { background: var(--fail-badge); color: var(--fail-text); }

.issues-cell { max-width: 360px; }
.issues-text {
  display: block;
  font-size: .78rem;
  color: var(--fail-text);
  word-break: break-word;
  line-height: 1.4;
}
.issues-warn {
  display: block;
  font-size: .78rem;
  color: var(--warn-text);
  word-break: break-word;
  line-height: 1.4;
}
.issues-none { color: var(--text-muted); font-size: .78rem; }

/* hidden rows */
tr.hidden { display: none; }

/* ===== No-results message ===== */
.no-results {
  display: none;
  text-align: center;
  padding: 32px;
  color: var(--text-muted);
  font-size: .95rem;
}

/* ===== Responsive ===== */
@media (max-width: 700px) {
  .col-path { display: none; }
  .col-pathlen { display: none; }
}
</style>
</head>
<body>

<div class="header">
  <h1>&#9729; Cloud Compliance Report</h1>
  <p class="subtitle">OneDrive / SharePoint naming &amp; path compliance scan</p>
</div>

<div class="summary">
  <div class="card">
    <div class="label">Total Files</div>
    <div class="value" id="cTotal">$total</div>
    <div class="meta">files scanned</div>
  </div>
  <div class="card ok">
    <div class="label">Compliant</div>
    <div class="value" id="cOk">$okCount</div>
    <div class="meta">no issues</div>
  </div>
  <div class="card warn">
    <div class="label">Warnings</div>
    <div class="value" id="cWarn">$warnCount</div>
    <div class="meta">approaching limits</div>
  </div>
  <div class="card fail">
    <div class="label">Issues Found</div>
    <div class="value" id="cFail">$failCount</div>
    <div class="meta">need attention</div>
  </div>
  <div class="card" style="grid-column: span 2;">
    <div class="label">Scan Path</div>
    <div class="meta" style="font-size:.9rem;color:var(--text);word-break:break-all;">$([System.Web.HttpUtility]::HtmlEncode($FolderPath))</div>
    <div class="meta" style="margin-top:6px;">Scanned: $($scanStart.ToString('yyyy-MM-dd HH:mm:ss')) &nbsp;|&nbsp; Duration: $([Math]::Round($duration,1))s</div>
  </div>
</div>

<div class="toolbar">
  <div class="filter-group">
    <button class="btn active" id="btnAll"  onclick="setFilter('all')">All ($total)</button>
    <button class="btn"        id="btnOk"   onclick="setFilter('ok')">OK ($okCount)</button>
    <button class="btn"        id="btnWarn" onclick="setFilter('warn')">Warnings ($warnCount)</button>
    <button class="btn"        id="btnFail" onclick="setFilter('fail')">Issues ($failCount)</button>
  </div>
  <div class="search-wrap">
    <input type="search" id="searchBox" placeholder="Search filename or path..." oninput="applySearch()" autocomplete="off">
  </div>
  <span class="count-label" id="visibleCount">Showing $total rows</span>
</div>

<div class="table-wrap">
  <table id="mainTable">
    <thead>
      <tr>
        <th onclick="sortTable(0)" data-col="0">Filename<span class="sort-icon"></span></th>
        <th onclick="sortTable(1)" data-col="1" class="col-path">Full Path<span class="sort-icon"></span></th>
        <th onclick="sortTable(2)" data-col="2" style="text-align:right">Path Len<span class="sort-icon"></span></th>
        <th onclick="sortTable(3)" data-col="3" style="text-align:right">Name Len<span class="sort-icon"></span></th>
        <th onclick="sortTable(4)" data-col="4">Status<span class="sort-icon"></span></th>
        <th onclick="sortTable(5)" data-col="5">Issues<span class="sort-icon"></span></th>
      </tr>
    </thead>
    <tbody id="tableBody">
$($rowsHtml.ToString())    </tbody>
  </table>
  <div class="no-results" id="noResults">No matching files found.</div>
</div>

<script>
// ---- State ----
var currentFilter = 'all';
var currentSort   = { col: -1, asc: true };

// ---- Filter ----
function setFilter(f) {
  currentFilter = f;
  document.getElementById('btnAll').classList.toggle('active',  f === 'all');
  document.getElementById('btnOk').classList.toggle('active',   f === 'ok');
  document.getElementById('btnWarn').classList.toggle('active', f === 'warn');
  document.getElementById('btnFail').classList.toggle('active', f === 'fail');
  applySearch();
}

// ---- Search + Filter combined ----
function applySearch() {
  var q    = document.getElementById('searchBox').value.toLowerCase().trim();
  var rows = document.querySelectorAll('#tableBody tr');
  var vis  = 0;

  rows.forEach(function(row) {
    var isOk   = row.classList.contains('row-ok');
    var isWarn = row.classList.contains('row-warn');
    var isFail = row.classList.contains('row-fail');

    var filterMatch = currentFilter === 'all'
      || (currentFilter === 'ok'   && isOk)
      || (currentFilter === 'warn' && isWarn)
      || (currentFilter === 'fail' && isFail);

    var path = row.getAttribute('data-path') || '';
    var name = row.getAttribute('data-name') || '';
    var searchMatch = !q || name.toLowerCase().indexOf(q) !== -1 || path.toLowerCase().indexOf(q) !== -1;

    if (filterMatch && searchMatch) {
      row.classList.remove('hidden');
      vis++;
    } else {
      row.classList.add('hidden');
    }
  });

  document.getElementById('visibleCount').textContent = 'Showing ' + vis + ' rows';
  document.getElementById('noResults').style.display = vis === 0 ? 'block' : 'none';
}

// ---- Sort ----
function sortTable(colIdx) {
  var tbody  = document.getElementById('tableBody');
  var rows   = Array.from(tbody.querySelectorAll('tr'));
  var ths    = document.querySelectorAll('thead th');

  var asc = currentSort.col === colIdx ? !currentSort.asc : true;
  currentSort = { col: colIdx, asc: asc };

  ths.forEach(function(th, i) {
    th.classList.remove('asc', 'desc');
    if (i === colIdx) th.classList.add(asc ? 'asc' : 'desc');
  });

  rows.sort(function(a, b) {
    var aCell = a.querySelectorAll('td')[colIdx];
    var bCell = b.querySelectorAll('td')[colIdx];
    var aVal  = aCell ? aCell.textContent.trim() : '';
    var bVal  = bCell ? bCell.textContent.trim() : '';

    // numeric columns
    if (colIdx === 2 || colIdx === 3) {
      return asc ? (+aVal - +bVal) : (+bVal - +aVal);
    }
    return asc ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
  });

  rows.forEach(function(r) { tbody.appendChild(r); });
}
</script>
</body>
</html>
"@

# Write report
try {
    [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
    Write-Host "Report saved: $reportPath" -ForegroundColor Green
} catch {
    Write-Error "Failed to write report: $_"
    exit 1
}

# Open in default browser
try {
    Start-Process $reportPath
} catch {
    Write-Warning "Could not auto-open report: $_"
    Write-Host "Open manually: $reportPath"
}
