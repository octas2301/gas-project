# G1: daily commit + push to origin/main (Asia/Tokyo morning task).
# Skips when clean. Never force-adds ignored secrets. Logs under _local_backup/git_daily/.
param(
  [switch]$DryRun,
  [string]$RepoRoot = 'C:\Users\takuy\Desktop\gas-project'
)

$ErrorActionPreference = 'Stop'
$tzStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logRoot = Join-Path $RepoRoot '_local_backup\git_daily'
New-Item -ItemType Directory -Force -Path $logRoot | Out-Null
$logFile = Join-Path $logRoot ("{0}.log" -f $tzStamp)
$lastFile = Join-Path $logRoot '_last_result.txt'

function Write-Log([string]$msg) {
  $line = "{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg
  Add-Content -LiteralPath $logFile -Value $line -Encoding utf8
  Write-Output $line
}

function Set-Last([string]$status, [string]$detail) {
  $body = @(
    "taken=$tzStamp"
    "status=$status"
    "detail=$detail"
  ) -join "`r`n"
  Set-Content -LiteralPath $lastFile -Value $body -Encoding utf8
}

try {
  Set-Location -LiteralPath $RepoRoot
  if (-not (Test-Path (Join-Path $RepoRoot '.git'))) {
    throw "not a git repo: $RepoRoot"
  }

  & git -C $RepoRoot rev-parse --abbrev-ref HEAD | Out-Null
  $branch = (& git -C $RepoRoot rev-parse --abbrev-ref HEAD).Trim()
  if ($branch -ne 'main') {
    Write-Log "SKIP branch=$branch (G1 only runs on main)"
    Set-Last 'skip_branch' $branch
    exit 0
  }

  & git -C $RepoRoot add -A
  $porcelain = & git -C $RepoRoot status --porcelain
  if (-not $porcelain) {
    Write-Log 'OK no changes'
    Set-Last 'ok_clean' 'no changes'
    exit 0
  }

  # Safety: refuse if a secrets path somehow became staged
  $staged = & git -C $RepoRoot diff --cached --name-only
  $bad = @(
    $staged | Where-Object {
      $_ -match '(^|/|\\)secrets(/|\\)|credentials\.json$|token\.json$|client_secret|\.pem$|\.p12$|api_key' 
    }
  )
  if ($bad.Count -gt 0) {
    & git -C $RepoRoot reset HEAD -- | Out-Null
    $detail = ($bad -join ' | ')
    Write-Log "FAIL refused secrets in staging: $detail"
    Set-Last 'fail_secrets' $detail
    exit 3
  }

  $ymd = Get-Date -Format 'yyyy-MM-dd'
  $msg = "chore(backup): daily auto $ymd"
  if ($DryRun) {
    Write-Log "DRYRUN would commit $($staged.Count) files: $msg"
    & git -C $RepoRoot reset HEAD -- | Out-Null
    Set-Last 'dry_run' ("files={0}" -f @($staged).Count)
    exit 0
  }

  & git -C $RepoRoot commit -m $msg
  if ($LASTEXITCODE -ne 0) {
    throw "git commit failed exit=$LASTEXITCODE"
  }

  & git -C $RepoRoot push origin main
  if ($LASTEXITCODE -ne 0) {
    throw "git push failed exit=$LASTEXITCODE"
  }

  $head = (& git -C $RepoRoot rev-parse --short HEAD).Trim()
  Write-Log "OK pushed head=$head files=$(@($staged).Count)"
  Set-Last 'ok_pushed' "head=$head files=$(@($staged).Count)"
  exit 0
}
catch {
  Write-Log ("FAIL {0}" -f $_.Exception.Message)
  Set-Last 'fail' $_.Exception.Message
  exit 1
}
