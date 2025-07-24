# PowerShell script to auto-increment version and update VERSION_HISTORY.md on publish
# Reads version from .csproj, increments patch, updates csproj and version history

$ErrorActionPreference = 'Stop'

$csproj = "PWQ.csproj"
$versionHistory = "VERSION_HISTORY.md"

# Get version from csproj using regex
$csprojContent = Get-Content $csproj -Raw
$versionMatch = [regex]::Match($csprojContent, '<Version>(.*?)</Version>')
if (-not $versionMatch.Success) {
    Write-Host "No <Version> found in $csproj. Exiting."
    exit 1
}
$version = $versionMatch.Groups[1].Value.Trim()
Write-Host "[DEBUG] Version read from csproj: '$version'"

# Increment patch version
if ($version -match '^(\d+)\.(\d+)\.(\d+)$') {
    $major = [int]$matches[1]
    $minor = [int]$matches[2]
    $patch = [int]$matches[3] + 1
    $newVersion = "$major.$minor.$patch"
} else {
    Write-Host "Version format not recognized: $version"
    exit 1
}

# Update csproj with new version using regex replace
$newCsprojContent = [regex]::Replace($csprojContent, '<Version>.*?</Version>', "<Version>$newVersion</Version>")
Set-Content $csproj $newCsprojContent
Write-Host "Updated $csproj to version $newVersion"

# Get today's date
$date = Get-Date -Format 'yyyy-MM-dd'

# Read current version history
$lines = Get-Content $versionHistory

# Prompt for change description
$changes = Read-Host "Enter change description for version $newVersion"
if (-not $changes) { $changes = "-" }

# Find the line after the header to insert new entry
$headerIndex = ($lines | Select-String -Pattern "^\| Version ").LineNumber
if (-not $headerIndex) { $headerIndex = 1 }
$insertIndex = $headerIndex

# Format new entry
$newEntry = "| $newVersion   | $date | $changes        |"

# Insert new entry after header
$updated = @()
for ($i = 0; $i -lt $lines.Count; $i++) {
    $updated += $lines[$i]
    if ($i -eq $insertIndex) {
        $updated += $newEntry
    }
}

# Write back to file
$updated | Set-Content $versionHistory
Write-Host "Added version $newVersion to $versionHistory."
