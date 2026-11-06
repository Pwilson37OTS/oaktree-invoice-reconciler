# Refresh the bundled data/ files from OneDrive and push to GitHub so the
# Streamlit Cloud deploy picks up the latest QBO and Bullhorn exports.
#
# Usage:
#   .\refresh_data.ps1
# or just double-click the file in Explorer (it'll open in a PowerShell window).
#
# Safe to run repeatedly — if nothing changed since last push, it skips the commit.

$ErrorActionPreference = "Stop"

# Validate that a copied file is a complete, readable .xlsx (OOXML zip).
# A truncated/mid-sync copy keeps the "PK" header but is missing the zip
# central directory, so openpyxl fails with "File is not a zip file" on the
# cloud. Opening it with .NET ZipFile reads that central directory, so a
# truncated file throws here and we catch it before committing.
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Test-XlsxValid {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $false }
    if ((Get-Item $Path).Length -lt 1024) { return $false }
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        foreach ($e in $zip.Entries) {
            if ($e.FullName -eq '[Content_Types].xml') { return $true }
        }
        return $false
    } catch {
        return $false
    } finally {
        if ($zip) { $zip.Dispose() }
    }
}

# Resolve repo paths from the script's own location
$repo = $PSScriptRoot
$ops_root = Split-Path (Split-Path $repo -Parent) -Parent   # OakTree Operations
$src_qbo = Join-Path $ops_root "_shared\quickbooks-data"
$src_ats = Join-Path $ops_root "_shared\bullhorn-data"
$dst_dir = Join-Path $repo "data"

Set-Location $repo

if (-not (Test-Path $dst_dir)) {
    New-Item -ItemType Directory -Force $dst_dir | Out-Null
}

$files = @(
    @{ src = "$src_qbo\audit_rev_ots.xlsx";          dst = "$dst_dir\audit_rev_ots.xlsx" },
    @{ src = "$src_qbo\audit_rev_cts.xlsx";          dst = "$dst_dir\audit_rev_cts.xlsx" },
    @{ src = "$src_ats\billable_charges.xlsx";       dst = "$dst_dir\billable_charges.xlsx" },
    @{ src = "$src_ats\billable_charges_CTS.xlsx";   dst = "$dst_dir\billable_charges_CTS.xlsx" }
)

Write-Host "Copying source files from OneDrive..." -ForegroundColor Cyan
$maxAttempts = 3
$failed = @()
foreach ($f in $files) {
    if (-not (Test-Path $f.src)) {
        Write-Error "Source file missing: $($f.src)"
    }
    $leaf = Split-Path $f.src -Leaf
    $ok = $false
    for ($i = 1; $i -le $maxAttempts; $i++) {
        Copy-Item $f.src $f.dst -Force
        if (Test-XlsxValid $f.dst) { $ok = $true; break }
        Write-Host "  $leaf  attempt $i produced a truncated/invalid file (OneDrive mid-sync?) - retrying in 5s" -ForegroundColor Yellow
        Start-Sleep -Seconds 5
    }
    if ($ok) {
        Write-Host "  $leaf  (validated)"
    } else {
        Write-Host "  $leaf  INVALID after $maxAttempts attempts" -ForegroundColor Red
        $failed += $leaf
        # Restore the last-known-good committed copy so we never stage a corrupt file.
        git checkout -- "data/$leaf" 2>$null
    }
}

if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Error ("Aborting without commit. Invalid copies (likely OneDrive still syncing): {0}. Nothing pushed; the cloud keeps its last good data. Re-run once OneDrive has finished syncing." -f ($failed -join ', '))
}

# Stage and check whether anything actually changed
git add data/ | Out-Null
$status = git status --porcelain data/
if (-not $status) {
    Write-Host ""
    Write-Host "No changes since last push. Cloud app already has the latest." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "Changes to push:" -ForegroundColor Cyan
$status | ForEach-Object { Write-Host "  $_" }

$stamp = Get-Date -Format "yyyy-MM-dd HH:mm"
git commit -m "Refresh data $stamp" | Out-Null
git push

Write-Host ""
Write-Host "Pushed. Streamlit Cloud will auto-redeploy in ~30 seconds." -ForegroundColor Green
