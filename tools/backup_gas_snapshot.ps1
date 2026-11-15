# Snapshot of listing GAS files that clasp push sends from the repo root.
# A new copy is kept only when files are not missing and not sharply smaller
# than the previous good copy. Pass -Force to keep a copy anyway.
param([switch]$Force)

$ErrorActionPreference = 'Stop'
$root = 'C:\Users\takuy\Desktop\gas-project'
$destRoot = Join-Path $root '_local_backup\gas'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
New-Item -ItemType Directory -Force -Path $destRoot | Out-Null

$files = @(Get-ChildItem -Path $root -File | Where-Object {
    ($_.Extension -eq '.js' -or $_.Name -eq 'appsscript.json') -and
    ($_.Name -notmatch 'bak|local_backup|check-rakuten-32byte')
})

function Get-GoodDirs {
    @(Get-ChildItem -LiteralPath $destRoot -Directory |
        Where-Object { $_.Name -match '^\d{8}_\d{6}$' } |
        Sort-Object Name -Descending)
}

$prev = Get-GoodDirs | Select-Object -First 1
$reasons = New-Object System.Collections.Generic.List[string]
if ($prev -and -not $Force) {
    $prevFiles = @(Get-ChildItem -LiteralPath $prev.FullName -File |
        Where-Object { $_.Name -ne '_note.txt' })
    $newMap = @{}
    foreach ($f in $files) { $newMap[$f.Name] = [int64]$f.Length }
    foreach ($p in $prevFiles) {
        $oldLen = [int64]$p.Length
        if (-not $newMap.ContainsKey($p.Name)) {
            $reasons.Add("missing $($p.Name)")
            continue
        }
        $newLen = [int64]$newMap[$p.Name]
        if ($oldLen -ge 50000 -and $newLen -lt [int64]($oldLen * 0.8)) {
            $reasons.Add("shrink $($p.Name) $oldLen -> $newLen")
        }
    }
}

$head = ''
try { $head = (& git -C $root rev-parse --short HEAD) } catch { $head = '' }

if ($reasons.Count -gt 0) {
    $rejRoot = Join-Path $destRoot '_rejected'
    $dest = Join-Path $rejRoot $stamp
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    foreach ($f in $files) {
        Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $dest $f.Name)
    }
    $note = @(
        "taken=$stamp"
        'status=rejected'
        "gitHead=$head"
        'reason=' + ($reasons -join ' | ')
        "keptGood=$($prev.Name)"
    ) -join "`r`n"
    Set-Content -LiteralPath (Join-Path $dest '_note.txt') -Value $note -Encoding utf8
    Set-Content -LiteralPath (Join-Path $destRoot '_last_result.txt') -Value $note -Encoding utf8
    $oldRej = @(Get-ChildItem -LiteralPath $rejRoot -Directory | Sort-Object Name -Descending)
    if ($oldRej.Count -gt 5) {
        $oldRej | Select-Object -Skip 5 | Remove-Item -Recurse -Force
    }
    Write-Output "REJECTED $($prev.Name) stays. $($reasons -join ' | ')"
    exit 2
}

$dest = Join-Path $destRoot $stamp
New-Item -ItemType Directory -Force -Path $dest | Out-Null
foreach ($f in $files) {
    Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $dest $f.Name)
}
$note = @(
    "taken=$stamp"
    'status=ok'
    "gitHead=$head"
    "files=$($files.Count)"
    $(if ($Force) { 'force=1' } else { 'force=0' })
) -join "`r`n"
Set-Content -LiteralPath (Join-Path $dest '_note.txt') -Value $note -Encoding utf8
Set-Content -LiteralPath (Join-Path $destRoot '_last_result.txt') -Value $note -Encoding utf8

$old = @(Get-GoodDirs)
if ($old.Count -gt 30) {
    $old | Select-Object -Skip 30 | Remove-Item -Recurse -Force
}

Write-Output $dest
