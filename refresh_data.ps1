# Refresh the bundled data/ files from OneDrive and push to GitHub so the
# Streamlit Cloud deploy picks up the latest QBO and Bullhorn exports.
#
# Usage:
#   .\refresh_data.ps1
# or just double-click the file in Explorer (it'll open in a PowerShell window).
#
# Safe to run repeatedly — if nothing changed since last push, it skips the commit.

$ErrorActionPreference = "Stop"

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
foreach ($f in $files) {
    if (-not (Test-Path $f.src)) {
        Write-Error "Source file missing: $($f.src)"
    }
    Copy-Item $f.src $f.dst -Force
    Write-Host "  $($f.src | Split-Path -Leaf)"
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
