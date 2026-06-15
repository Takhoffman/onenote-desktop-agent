Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$Source = Join-Path $RepoRoot "skills\onenote-desktop"
$Destination = Join-Path $RepoRoot "plugins\onenote-desktop\skills\onenote-desktop"

if (-not (Test-Path -LiteralPath $Source -PathType Container)) {
    throw "Source skill directory not found: $Source"
}

if (-not (Test-Path -LiteralPath $Destination -PathType Container)) {
    throw "Packaged skill directory not found: $Destination"
}

function Get-RelativeFileHash {
    param(
        [Parameter(Mandatory = $true)][string]$Root
    )

    $ResolvedRoot = (Resolve-Path -LiteralPath $Root).Path.TrimEnd("\")
    Get-ChildItem -LiteralPath $ResolvedRoot -Recurse -File |
        Sort-Object FullName |
        ForEach-Object {
            [PSCustomObject]@{
                Path = $_.FullName.Substring($ResolvedRoot.Length + 1)
                Hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
            }
        }
}

$SourceFiles = Get-RelativeFileHash -Root $Source
$DestinationFiles = Get-RelativeFileHash -Root $Destination

$Differences = Compare-Object `
    -ReferenceObject $SourceFiles `
    -DifferenceObject $DestinationFiles `
    -Property Path, Hash `
    -PassThru

if ($Differences) {
    Write-Error "Packaged skill differs from source skill. Run scripts\sync-plugin-package.ps1 and review the result."
}

Write-Host "Packaged skill matches source skill."
