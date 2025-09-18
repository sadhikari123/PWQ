# PowerShell script to auto-increment version and update VERSION_HISTORY.md on publish
# Reads version from .csproj, increments patch, updates csproj and version history

$ErrorActionPreference = 'Stop'

$csproj = "PWQ.csproj"
$versionHistory = "VERSION_HISTORY.md"

function Write-DebugLine($msg) {
    Write-Host "[update_version_history] $msg"
}

# Ensure VERSION_HISTORY.md exists and has a simple header if missing
if (-not (Test-Path $versionHistory)) {
    Write-DebugLine "$versionHistory not found. Creating with header."
    @"
| Version | Date | Changes |
|---|---|---|
"@ | Out-File -Encoding UTF8 $versionHistory
}

# Read csproj and extract version
if (-not (Test-Path $csproj)) {
    Write-DebugLine "Could not find $csproj in working directory ($(Get-Location)). Exiting."
    exit 1
}

$csprojContent = Get-Content $csproj -Raw
$versionMatch = [regex]::Match($csprojContent, '<Version>(.*?)</Version>')
if (-not $versionMatch.Success) {
    Write-DebugLine "No <Version> found in $csproj. Exiting."
    exit 1
}
$version = $versionMatch.Groups[1].Value.Trim()
Write-DebugLine "Version read from csproj: '$version'"

# Increment patch version (expecting semantic 3-part version)
if ($version -match '^([0-9]+)\.([0-9]+)\.([0-9]+)$') {
    $major = [int]$matches[1]
    $minor = [int]$matches[2]
    $patch = [int]$matches[3] + 1
    $newVersion = "$major.$minor.$patch"
} else {
    Write-DebugLine "Version format not recognized: $version. Expected 'MAJOR.MINOR.PATCH'. Exiting."
    exit 1
}

# Update csproj with new version
$newCsprojContent = [regex]::Replace($csprojContent, '<Version>.*?</Version>', "<Version>$newVersion</Version>")
Set-Content -Encoding UTF8 $csproj $newCsprojContent
Write-DebugLine "Updated $csproj to version $newVersion"

# Get today's date
$date = Get-Date -Format 'yyyy-MM-dd'

# Read current version history lines
$lines = Get-Content -Raw $versionHistory -Encoding UTF8 -ErrorAction Stop
$linesArr = $lines -split "`r?`n"

# Determine header line index (0-based). Find a line containing both 'Version' and 'Date' as the header
$headerIndex = -1
for ($i = 0; $i -lt $linesArr.Count; $i++) {
    if ($linesArr[$i] -match '(?i)\|.*\bVersion\b.*\|.*\bDate\b.*\|') {
        $headerIndex = $i
        break
    }
}

if ($headerIndex -eq -1) {
    # Prepend a header
    Write-DebugLine "Header row not found in $versionHistory; adding default header."
    $newHeader = @('| Version | Date | Changes |','|---|---|---|')
    $linesArr = $newHeader + $linesArr
    $headerIndex = 0
}

$insertIndex = $headerIndex + 1

# Try to gather today's git commit messages; if git not available or no commits, use '-'
$changes = "-"
try {
    $gitCmd = "git"
    $since = "$date 00:00"
    $until = "$date 23:59"
    $gitLogRaw = & $gitCmd log --since="$since" --until="$until" --pretty=format:"%s" 2>$null
    if ($gitLogRaw -and $gitLogRaw.Trim() -ne "") {
        if ($gitLogRaw -is [array]) { $gitMsgs = $gitLogRaw } else { $gitMsgs = $gitLogRaw -split "`n" }
        $gitMsgs = $gitMsgs | Where-Object { $_ -and $_.Trim() -ne "" }
        if ($gitMsgs.Count -gt 0) {
            $changes = ($gitMsgs -join '; ')
        }
    }
} catch {
    Write-DebugLine "git not available or failed to run; using '-' for changes."
}

# Remove any existing entry for today's date (lines matching | x.y.z | yyyy-mm-dd |)
$filtered = @()
for ($i = 0; $i -lt $linesArr.Count; $i++) {
    $line = $linesArr[$i]
    if ($line -match '^\|\s*\d+\.\d+\.\d+\s*\|\s*' + [regex]::Escape($date) + '\s*\|') {
        continue
    }
    $filtered += $line
}

# Format new entry (pad to match existing table spacing)
$newEntry = "| $newVersion | $date | $changes |"

# Insert new entry after header (or replace existing entry with same version)
$replaced = $false
$versionPattern = '^\|\s*' + [regex]::Escape($newVersion) + '\s*\|'
for ($i = 0; $i -lt $filtered.Count; $i++) {
    if ($filtered[$i] -match $versionPattern) {
        # Replace existing version line
        $filtered[$i] = $newEntry
        $replaced = $true
        break
    }
}

if (-not $replaced) {
    $updated = @()
    for ($i = 0; $i -lt $filtered.Count; $i++) {
        $updated += $filtered[$i]
        if ($i -eq $insertIndex) { $updated += $newEntry }
    }
    if ($insertIndex -ge $filtered.Count) { $updated += $newEntry }
} else {
    $updated = $filtered
}

# Deduplicate exact lines while preserving order
$seen = @{}
$final = @()
foreach ($ln in $updated) {
    if (-not $seen.ContainsKey($ln)) {
        $final += $ln
        $seen[$ln] = $true
    }
}

# Write back to VERSION_HISTORY.md
$final -join "`n" | Set-Content -Encoding UTF8 $versionHistory
Write-DebugLine "Updated $versionHistory (added or replaced version $newVersion)."
