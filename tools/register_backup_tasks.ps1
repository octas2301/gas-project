# Register / refresh Windows scheduled tasks for G1 (git) and G3 (all-clasp snapshots).
# Run once in PowerShell (current user). Does not require clasp push.
param(
  [string]$RepoRoot = 'C:\Users\takuy\Desktop\gas-project'
)

$ErrorActionPreference = 'Stop'
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

function Register-DailyTask([string]$name, [string]$scriptPath, [string]$time) {
  if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "missing script: $scriptPath"
  }
  $tr = "$ps -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
  # Replace if exists (ignore "not found")
  cmd /c "schtasks /Delete /TN `"$name`" /F >NUL 2>&1" | Out-Null
  & schtasks /Create /TN $name /SC DAILY /ST $time /TR $tr /RL LIMITED /F
  if ($LASTEXITCODE -ne 0) {
    throw "schtasks create failed for $name"
  }
  Write-Output "OK $name at $time -> $scriptPath"
}

Register-DailyTask 'gas-project-git-daily-backup' (Join-Path $RepoRoot 'tools\git_daily_backup.ps1') '08:05'
Register-DailyTask 'gas-project-clasp-all-snapshot' (Join-Path $RepoRoot 'tools\backup_all_clasp_snapshots.ps1') '08:00'

# Keep existing listing-only snapshot task if present; do not delete it.
Write-Output 'NOTE: existing gas-project-gas-snapshot (listing-only) left as-is if present.'
Write-Output 'Done.'
