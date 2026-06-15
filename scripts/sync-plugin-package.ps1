Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$Source = Join-Path $RepoRoot "skills\onenote-desktop"
$Destination = Join-Path $RepoRoot "plugins\onenote-desktop\skills\onenote-desktop"

if (-not (Test-Path -LiteralPath $Source -PathType Container)) {
    throw "Source skill directory not found: $Source"
}

$PluginSkillsRoot = Join-Path $RepoRoot "plugins\onenote-desktop\skills"
if (-not (Test-Path -LiteralPath $PluginSkillsRoot -PathType Container)) {
    New-Item -ItemType Directory -Path $PluginSkillsRoot | Out-Null
}

$ResolvedPluginSkillsRoot = (Resolve-Path -LiteralPath $PluginSkillsRoot).Path
$DestinationParent = Split-Path -Parent $Destination
$ResolvedDestinationParent = (Resolve-Path -LiteralPath $DestinationParent).Path

if (-not $ResolvedDestinationParent.StartsWith($ResolvedPluginSkillsRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to sync outside plugin skills root: $Destination"
}

if (Test-Path -LiteralPath $Destination) {
    Remove-Item -LiteralPath $Destination -Recurse -Force
}

Copy-Item -LiteralPath $Source -Destination $Destination -Recurse
Write-Host "Synced $Source -> $Destination"
