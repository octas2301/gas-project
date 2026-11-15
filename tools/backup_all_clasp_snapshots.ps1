# G3: local snapshot for every clasp project root (insurance beside GitHub G1).
# Rejects a project copy when a previously good large file shrinks >20%.
param(
  [switch]$Force,
  [string]$RepoRoot = 'C:\Users\takuy\Desktop\gas-project'
)

$ErrorActionPreference = 'Stop'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$destRoot = Join-Path $RepoRoot '_local_backup\clasp_all'
New-Item -ItemType Directory -Force -Path $destRoot | Out-Null

$projects = @(
  @{ Name = 'listing'; Path = $RepoRoot },
  @{ Name = 'purchase_research'; Path = (Join-Path $RepoRoot 'purchase_research') },
  @{ Name = 'competitor_jitta'; Path = (Join-Path $RepoRoot 'competitor_jitta') },
  @{ Name = 'ads'; Path = (Join-Path $RepoRoot '広告運用GAS') },
  @{ Name = 'inventory'; Path = (Join-Path $RepoRoot 'inventory-app') },
  @{ Name = 'export_tax'; Path = (Join-Path $RepoRoot 'export-tax-list') }
)

function Get-SnapshotFiles([string]$projPath) {
  $files = @()
  if (-not (Test-Path -LiteralPath $projPath)) { return $files }
  $files += @(Get-ChildItem -LiteralPath $projPath -File -ErrorAction SilentlyContinue | Where-Object {
    ($_.Extension -in '.js', '.gs', '.json') -and
    ($_.Name -notmatch 'bak|local_backup|broken|keep_|merged_|check-rakuten')
  })
  return $files
}

function Get-GoodDirs([string]$pjDest) {
  if (-not (Test-Path -LiteralPath $pjDest)) { return @() }
  @(Get-ChildItem -LiteralPath $pjDest -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match '^\d{8}_\d{6}$' } |
    Sort-Object Name -Descending)
}

$summary = New-Object System.Collections.Generic.List[string]
$fail = 0

foreach ($p in $projects) {
  $name = $p.Name
  $path = $p.Path
  $pjDest = Join-Path $destRoot $name
  New-Item -ItemType Directory -Force -Path $pjDest | Out-Null

  if (-not (Test-Path -LiteralPath (Join-Path $path '.clasp.json'))) {
    $summary.Add("$name SKIP no .clasp.json")
    continue
  }

  $files = @(Get-SnapshotFiles $path)
  if ($files.Count -eq 0) {
    $summary.Add("$name SKIP no js/gs/json")
    continue
  }

  $prev = Get-GoodDirs $pjDest | Select-Object -First 1
  $reasons = New-Object System.Collections.Generic.List[string]
  if ($prev -and -not $Force) {
    $prevFiles = @(Get-ChildItem -LiteralPath $prev.FullName -File | Where-Object { $_.Name -ne '_note.txt' })
    $newMap = @{}
    foreach ($f in $files) { $newMap[$f.Name] = [int64]$f.Length }
    foreach ($old in $prevFiles) {
      $oldLen = [int64]$old.Length
      if (-not $newMap.ContainsKey($old.Name)) {
        $reasons.Add("missing $($old.Name)")
        continue
      }
      $newLen = [int64]$newMap[$old.Name]
      if ($oldLen -ge 50000 -and $newLen -lt [int64]($oldLen * 0.8)) {
        $reasons.Add("shrink $($old.Name) $oldLen -> $newLen")
      }
    }
  }

  $head = ''
  try { $head = (& git -C $RepoRoot rev-parse --short HEAD) } catch { $head = '' }

  if ($reasons.Count -gt 0) {
    $rej = Join-Path $pjDest '_rejected'
    $dest = Join-Path $rej $stamp
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    foreach ($f in $files) {
      Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $dest $f.Name) -Force
    }
    $note = @(
      "taken=$stamp"
      'status=rejected'
      "gitHead=$head"
      'reason=' + ($reasons -join ' | ')
      "keptGood=$($prev.Name)"
    ) -join "`r`n"
    Set-Content -LiteralPath (Join-Path $dest '_note.txt') -Value $note -Encoding utf8
    $summary.Add("$name REJECTED $($reasons[0])")
    $fail++
    continue
  }

  $dest = Join-Path $pjDest $stamp
  New-Item -ItemType Directory -Force -Path $dest | Out-Null
  foreach ($f in $files) {
    Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $dest $f.Name) -Force
  }
  $note = @(
    "taken=$stamp"
    'status=ok'
    "gitHead=$head"
    "files=$($files.Count)"
    $(if ($Force) { 'force=1' } else { 'force=0' })
  ) -join "`r`n"
  Set-Content -LiteralPath (Join-Path $dest '_note.txt') -Value $note -Encoding utf8

  $old = @(Get-GoodDirs $pjDest)
  if ($old.Count -gt 14) {
    $old | Select-Object -Skip 14 | Remove-Item -Recurse -Force
  }
  $summary.Add("$name OK files=$($files.Count) -> $dest")
}

$summaryText = ($summary -join "`r`n")
Set-Content -LiteralPath (Join-Path $destRoot '_last_result.txt') -Value $summaryText -Encoding utf8
Write-Output $summaryText
if ($fail -gt 0) { exit 2 }
exit 0
