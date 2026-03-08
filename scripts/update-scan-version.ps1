# external-scan-poc.html の SCAN_PAGE_VERSION を現在日時に更新する。
# 使い方: リポジトリルート（gas-project）で実行する。
#   .\scripts\update-scan-version.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not (Test-Path (Join-Path $root 'inventory-app\external-scan-poc.html'))) {
  $root = Get-Location
}
$path = Join-Path $root 'inventory-app\external-scan-poc.html'
if (-not (Test-Path $path)) {
  Write-Error "File not found: $path (run from gas-project root)"
  exit 1
}
$now = Get-Date -Format 'yyyy-MM-dd HH:mm'
$content = Get-Content $path -Raw -Encoding UTF8
$content = $content -replace "var SCAN_PAGE_VERSION = '[^']*';", "var SCAN_PAGE_VERSION = '$now';"
Set-Content $path $content -NoNewline -Encoding UTF8
Write-Host "SCAN_PAGE_VERSION updated to: $now"
