param(
    [int]$MaxFilesPerRoot = 200,
    [switch]$IncludeHashes,
    [string]$OutputDir = '',
    [string[]]$ExtraRoot = @()
)

$ErrorActionPreference = 'Stop'

function Write-Json {
    param([object]$Value)
    $Value | ConvertTo-Json -Depth 100
}

function Get-FileSignature {
    param([string]$Path)

    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
    try {
        $readLength = [Math]::Min(32, [int]$stream.Length)
        $bytes = New-Object byte[] $readLength
        if ($readLength -gt 0) {
            [void]$stream.Read($bytes, 0, $readLength)
        }
        $fileLength = $stream.Length
    }
    finally {
        $stream.Dispose()
    }

    $headLength = [Math]::Min(32, $bytes.Length)
    $headBytes = if ($headLength -gt 0) { $bytes[0..($headLength - 1)] } else { @() }
    $headHex = ($headBytes | ForEach-Object { $_.ToString('X2') }) -join ' '
    $headAscii = -join ($headBytes | ForEach-Object {
        if ($_ -ge 32 -and $_ -le 126) { [char]$_ } else { '.' }
    })

    $kind = 'unknown'
    if ($bytes.Length -ge 4) {
        $first4 = ($bytes[0..3] | ForEach-Object { $_.ToString('X2') }) -join ' '
        switch ($first4) {
            'E4 52 5C 7B' { $kind = 'onenote-section-or-toc' }
            '50 4B 03 04' { $kind = 'zip-or-onepkg' }
            'D0 CF 11 E0' { $kind = 'ole-compound-file' }
            '30 26 B2 75' { $kind = 'asf-wma-media' }
            default {}
        }
    }

    [pscustomobject]@{
        headHex = $headHex
        headAscii = $headAscii
        detectedKind = $kind
        readable = $true
        readError = ''
        length = $fileLength
    }
}

function Add-Root {
    param(
        [System.Collections.ArrayList]$Roots,
        [string]$Name,
        [string]$Path,
        [string]$Role,
        [bool]$Scan = $true
    )

    [void]$Roots.Add([pscustomobject]@{
        name = $Name
        path = $Path
        role = $Role
        scan = $Scan
        exists = [bool](Test-Path -LiteralPath $Path)
    })
}

$roots = [System.Collections.ArrayList]::new()
Add-Root $roots 'office-onenote-local' (Join-Path $env:LOCALAPPDATA 'Microsoft\OneNote') 'Office desktop OneNote local data root' $false
Add-Root $roots 'office-onenote-version' (Join-Path $env:LOCALAPPDATA 'Microsoft\OneNote\16.0') 'Office desktop OneNote versioned data root' $false
Add-Root $roots 'office-onenote-cache' (Join-Path $env:LOCALAPPDATA 'Microsoft\OneNote\16.0\cache') 'volatile desktop cache blocks'
Add-Root $roots 'office-onenote-backup' (Join-Path $env:LOCALAPPDATA 'Microsoft\OneNote\16.0\Backup') 'desktop backup sections'
Add-Root $roots 'office-onenote-audio-cache' (Join-Path $env:LOCALAPPDATA 'Microsoft\OneNote\16.0\Audio Cache') 'desktop audio recording cache'
Add-Root $roots 'office-onenote-search-index' (Join-Path $env:LOCALAPPDATA 'Microsoft\OneNote\16.0\FullTextSearchIndex') 'desktop full text search index'
Add-Root $roots 'office-onenote-master-index' (Join-Path $env:LOCALAPPDATA 'Microsoft\OneNote\16.0\MasterIndex') 'desktop master index'
Add-Root $roots 'documents-onenote-notebooks' (Join-Path $env:USERPROFILE 'Documents\OneNote Notebooks') 'default local notebook folder'
Add-Root $roots 'uwp-onenote-localcache' (Join-Path $env:LOCALAPPDATA 'Packages\Microsoft.Office.OneNote_8wekyb3d8bbwe\LocalCache') 'legacy/store OneNote local cache root'
for ($i = 0; $i -lt $ExtraRoot.Count; $i++) {
    Add-Root $roots "extra-root-$i" $ExtraRoot[$i] 'caller-supplied root'
}

$extensions = @('.one', '.onetoc2', '.onepkg', '.bin', '.dat', '.idx', '.wma', '.m4a', '.mp3', '.mp4')
$files = @()
foreach ($root in $roots | Where-Object { $_.exists -and $_.scan }) {
    $found = @(Get-ChildItem -LiteralPath $root.path -Recurse -Force -File -ErrorAction SilentlyContinue |
        Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object Length -Descending |
        Select-Object -First $MaxFilesPerRoot)

    foreach ($file in $found) {
        try {
            $sig = Get-FileSignature $file.FullName
        }
        catch {
            $sig = [pscustomobject]@{
                headHex = ''
                headAscii = ''
                detectedKind = 'unreadable'
                readable = $false
                readError = $_.Exception.Message
                length = $file.Length
            }
        }
        $hash = $null
        if ($IncludeHashes) {
            $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        }

        $files += [pscustomobject]@{
            root = $root.name
            path = $file.FullName
            extension = $file.Extension
            bytes = $file.Length
            lastWriteTime = $file.LastWriteTime.ToString('o')
            detectedKind = $sig.detectedKind
            headHex = $sig.headHex
            headAscii = $sig.headAscii
            readable = $sig.readable
            readError = $sig.readError
            sha256 = $hash
        }
    }
}

$summary = $files |
    Group-Object root, extension, detectedKind |
    ForEach-Object {
        [pscustomobject]@{
            group = $_.Name
            count = $_.Count
            totalBytes = ($_.Group | Measure-Object bytes -Sum).Sum
        }
    }

$result = [pscustomobject]@{
    ok = $true
    generatedAt = (Get-Date).ToString('o')
    note = 'Read-only storage inventory. Do not modify OneNote cache files through this probe.'
    roots = $roots
    summary = @($summary)
    files = @($files)
}

if ($OutputDir) {
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
    $outPath = Join-Path $OutputDir 'onenote-storage-probe.json'
    Write-Json $result | Set-Content -LiteralPath $outPath -Encoding UTF8
}

Write-Json $result
