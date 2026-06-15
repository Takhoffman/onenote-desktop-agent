param(
    [ValidateSet('check-install', 'hierarchy', 'get-page', 'create-page', 'update-page-xml', 'find-pages', 'find-meta', 'get-parent', 'get-hyperlink', 'get-special-location', 'navigate', 'publish', 'get-binary', 'demo-write', 'control-page', 'extract-media')]
    [string]$Operation = 'check-install',

    [string]$StartNodeId = '',

    [ValidateSet('self', 'children', 'notebooks', 'sections', 'pages')]
    [string]$Scope = 'notebooks',

    [string]$PageId = '',
    [string]$SectionId = '',
    [string]$Title = 'Untitled',
    [string]$Text = '',
    [string]$XmlPath = '',
    [string]$Query = '',
    [string]$SectionPath = '',
    [string]$OutputDir = '',
    [string]$ObjectId = '',
    [string]$TargetPath = '',
    [ValidateSet('one', 'onepkg', 'mhtml', 'pdf', 'xps', 'word', 'emf')]
    [string]$PublishFormat = 'pdf',
    [ValidateSet('backup-folder', 'unfiled-notes-section', 'default-notebook-folder')]
    [string]$SpecialLocation = 'default-notebook-folder',
    [string]$CallbackId = '',
    [switch]$Display,
    [switch]$IncludeUnindexedPages,
    [switch]$NewWindow
)

$ErrorActionPreference = 'Stop'

function Write-Json {
    param([object]$Value)
    $Value | ConvertTo-Json -Depth 100
}

function Get-OneNoteApp {
    New-Object -ComObject OneNote.Application
}

function Get-ScopeValue {
    param([string]$Name)
    switch ($Name) {
        'self' { 0 }
        'children' { 1 }
        'notebooks' { 2 }
        'sections' { 3 }
        'pages' { 4 }
    }
}

function Get-PublishFormatValue {
    param([string]$Name)
    switch ($Name) {
        'one' { 0 }
        'onepkg' { 1 }
        'mhtml' { 2 }
        'pdf' { 3 }
        'xps' { 4 }
        'word' { 5 }
        'emf' { 6 }
    }
}

function Get-SpecialLocationValue {
    param([string]$Name)
    switch ($Name) {
        'backup-folder' { 0 }
        'unfiled-notes-section' { 1 }
        'default-notebook-folder' { 2 }
    }
}

function Convert-OneNoteXml {
    param([string]$Xml)
    if ([string]::IsNullOrWhiteSpace($Xml)) {
        return @{
            xml = $Xml
            nodes = @()
        }
    }

    [xml]$doc = $Xml
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('one', 'http://schemas.microsoft.com/office/onenote/2013/onenote')

    $nodes = @()
    foreach ($node in $doc.SelectNodes('//*[@ID]', $ns)) {
        $nodes += [pscustomobject]@{
            kind = $node.LocalName
            name = $node.GetAttribute('name')
            id = $node.GetAttribute('ID')
            path = $node.GetAttribute('path')
            lastModifiedTime = $node.GetAttribute('lastModifiedTime')
        }
    }

    @{
        xml = $Xml
        nodes = $nodes
    }
}

function New-OneNotePageXml {
    param(
        [string]$ExistingXml,
        [string]$NewTitle,
        [string]$BodyText
    )

    [xml]$doc = $ExistingXml
    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('one', 'http://schemas.microsoft.com/office/onenote/2013/onenote')
    $ns = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

    $page = $doc.SelectSingleNode('/one:Page', $nsm)
    if ($null -eq $page) {
        throw 'Page XML did not contain one:Page.'
    }
    $page.SetAttribute('dateTime', (Get-Date).ToString('s'))

    $titleNode = $doc.SelectSingleNode('/one:Page/one:Title/one:OE/one:T', $nsm)
    if ($null -eq $titleNode) {
        $title = $doc.CreateElement('one', 'Title', $ns)
        $oe = $doc.CreateElement('one', 'OE', $ns)
        $titleNode = $doc.CreateElement('one', 'T', $ns)
        [void]$oe.AppendChild($titleNode)
        [void]$title.AppendChild($oe)
        [void]$page.PrependChild($title)
    }
    $titleNode.InnerText = $NewTitle

    if (-not [string]::IsNullOrWhiteSpace($BodyText)) {
        $outline = $doc.CreateElement('one', 'Outline', $ns)
        $oeChildren = $doc.CreateElement('one', 'OEChildren', $ns)
        $oe = $doc.CreateElement('one', 'OE', $ns)
        $t = $doc.CreateElement('one', 'T', $ns)
        $t.InnerText = $BodyText
        [void]$oe.AppendChild($t)
        [void]$oeChildren.AppendChild($oe)
        [void]$outline.AppendChild($oeChildren)
        [void]$page.AppendChild($outline)
    }

    $doc.OuterXml
}

function Add-OneNoteOutline {
    param(
        [xml]$Document,
        [System.Xml.XmlElement]$Page,
        [string]$Content,
        [double]$X,
        [double]$Y,
        [double]$Width = 640.0
    )

    $ns = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
    $outline = $Document.CreateElement('one', 'Outline', $ns)
    $position = $Document.CreateElement('one', 'Position', $ns)
    $position.SetAttribute('x', ([string]$X))
    $position.SetAttribute('y', ([string]$Y))
    $position.SetAttribute('z', '0')
    $size = $Document.CreateElement('one', 'Size', $ns)
    $size.SetAttribute('width', ([string]$Width))
    $size.SetAttribute('height', '90.0')
    $oeChildren = $Document.CreateElement('one', 'OEChildren', $ns)
    $oe = $Document.CreateElement('one', 'OE', $ns)
    $t = $Document.CreateElement('one', 'T', $ns)
    [void]$t.AppendChild($Document.CreateCDataSection($Content))
    [void]$oe.AppendChild($t)
    [void]$oeChildren.AppendChild($oe)
    [void]$outline.AppendChild($position)
    [void]$outline.AppendChild($size)
    [void]$outline.AppendChild($oeChildren)
    [void]$Page.AppendChild($outline)
}

function New-OneNoteControlPageXml {
    param(
        [string]$ExistingXml,
        [string]$NewTitle,
        [string[]]$Blocks
    )

    [xml]$doc = $ExistingXml
    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('one', 'http://schemas.microsoft.com/office/onenote/2013/onenote')
    $page = $doc.SelectSingleNode('/one:Page', $nsm)
    if ($null -eq $page) {
        throw 'Page XML did not contain one:Page.'
    }

    $titleNode = $doc.SelectSingleNode('/one:Page/one:Title/one:OE/one:T', $nsm)
    if ($null -ne $titleNode) {
        $titleNode.RemoveAll()
        [void]$titleNode.AppendChild($doc.CreateCDataSection($NewTitle))
    }

    $y = 86.4
    foreach ($block in $Blocks) {
        Add-OneNoteOutline -Document $doc -Page $page -Content $block -X 36.0 -Y $y
        $lineCount = (($block -split "`n").Count)
        $y += [Math]::Max(72.0, 22.0 * ($lineCount + 1))
    }

    $doc.OuterXml
}

function Get-SafeFileName {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $safe = -join ($Name.ToCharArray() | ForEach-Object {
        if ($invalid -contains $_) { '_' } else { $_ }
    })
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return 'onenote-media.bin'
    }
    $safe
}

try {
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        throw 'Run this script with powershell.exe -NoProfile -STA.'
    }

    switch ($Operation) {
        'check-install' {
            $appPath = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\ONENOTE.EXE' -ErrorAction SilentlyContinue
            $exe = if ($appPath) { Join-Path $appPath.Path 'ONENOTE.EXE' } else { $null }
            $version = if ($exe -and (Test-Path $exe)) { (Get-Item $exe).VersionInfo.ProductVersion } else { $null }
            $app = Get-OneNoteApp
            Write-Json ([pscustomobject]@{
                ok = $true
                apartment = [Threading.Thread]::CurrentThread.GetApartmentState().ToString()
                exe = $exe
                version = $version
                comCreated = ($null -ne $app)
            })
        }
        'hierarchy' {
            $app = Get-OneNoteApp
            $xml = ''
            $app.GetHierarchy($StartNodeId, (Get-ScopeValue $Scope), [ref]$xml)
            Write-Json (Convert-OneNoteXml $xml)
        }
        'get-page' {
            if (-not $PageId) { throw '-PageId is required.' }
            $app = Get-OneNoteApp
            $xml = ''
            $app.GetPageContent($PageId, [ref]$xml)
            Write-Json (Convert-OneNoteXml $xml)
        }
        'create-page' {
            if (-not $SectionId) { throw '-SectionId is required.' }
            $app = Get-OneNoteApp
            $newPageId = ''
            $app.CreateNewPage($SectionId, [ref]$newPageId)
            $xml = ''
            $app.GetPageContent($newPageId, [ref]$xml)
            $updated = New-OneNotePageXml -ExistingXml $xml -NewTitle $Title -BodyText $Text
            $app.UpdatePageContent($updated)
            Write-Json ([pscustomobject]@{
                ok = $true
                pageId = $newPageId
                title = $Title
            })
        }
        'update-page-xml' {
            if (-not $XmlPath) { throw '-XmlPath is required.' }
            if (-not (Test-Path -LiteralPath $XmlPath)) { throw "XML path not found: $XmlPath" }
            $app = Get-OneNoteApp
            $xml = Get-Content -LiteralPath $XmlPath -Raw
            $app.UpdatePageContent($xml)
            Write-Json ([pscustomobject]@{
                ok = $true
                xmlPath = (Resolve-Path -LiteralPath $XmlPath).Path
            })
        }
        'find-pages' {
            if (-not $Query) { throw '-Query is required.' }
            $app = Get-OneNoteApp
            $xml = ''
            $app.FindPages($StartNodeId, $Query, [ref]$xml, [bool]$IncludeUnindexedPages, [bool]$Display)
            Write-Json (Convert-OneNoteXml $xml)
        }
        'find-meta' {
            if (-not $Query) { throw '-Query is required.' }
            $app = Get-OneNoteApp
            $xml = ''
            $app.FindMeta($StartNodeId, $Query, [ref]$xml, [bool]$IncludeUnindexedPages)
            Write-Json (Convert-OneNoteXml $xml)
        }
        'get-parent' {
            $id = if ($ObjectId) { $ObjectId } elseif ($PageId) { $PageId } elseif ($SectionId) { $SectionId } else { '' }
            if (-not $id) { throw '-ObjectId, -PageId, or -SectionId is required.' }
            $app = Get-OneNoteApp
            $parentId = ''
            $app.GetHierarchyParent($id, [ref]$parentId)
            Write-Json ([pscustomobject]@{
                ok = $true
                objectId = $id
                parentId = $parentId
            })
        }
        'get-hyperlink' {
            $hierarchyId = if ($ObjectId) { $ObjectId } elseif ($PageId) { $PageId } elseif ($SectionId) { $SectionId } else { '' }
            if (-not $hierarchyId) { throw '-ObjectId, -PageId, or -SectionId is required.' }
            $app = Get-OneNoteApp
            $link = ''
            $app.GetHyperlinkToObject($hierarchyId, $CallbackId, [ref]$link)
            Write-Json ([pscustomobject]@{
                ok = $true
                hierarchyId = $hierarchyId
                objectId = $CallbackId
                hyperlink = $link
            })
        }
        'get-special-location' {
            $app = Get-OneNoteApp
            $path = ''
            $app.GetSpecialLocation((Get-SpecialLocationValue $SpecialLocation), [ref]$path)
            Write-Json ([pscustomobject]@{
                ok = $true
                specialLocation = $SpecialLocation
                path = $path
            })
        }
        'navigate' {
            $hierarchyId = if ($ObjectId) { $ObjectId } elseif ($PageId) { $PageId } elseif ($SectionId) { $SectionId } else { '' }
            if (-not $hierarchyId) { throw '-ObjectId, -PageId, or -SectionId is required.' }
            $app = Get-OneNoteApp
            $app.NavigateTo($hierarchyId, $CallbackId, [bool]$NewWindow)
            Write-Json ([pscustomobject]@{
                ok = $true
                hierarchyId = $hierarchyId
                objectId = $CallbackId
                newWindow = [bool]$NewWindow
            })
        }
        'publish' {
            $hierarchyId = if ($ObjectId) { $ObjectId } elseif ($PageId) { $PageId } elseif ($SectionId) { $SectionId } else { '' }
            if (-not $hierarchyId) { throw '-ObjectId, -PageId, or -SectionId is required.' }
            if (-not $TargetPath) { throw '-TargetPath is required.' }
            $targetDir = Split-Path -Parent $TargetPath
            if ($targetDir) {
                New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
            }
            $app = Get-OneNoteApp
            $app.Publish($hierarchyId, $TargetPath, (Get-PublishFormatValue $PublishFormat), '')
            Write-Json ([pscustomobject]@{
                ok = $true
                hierarchyId = $hierarchyId
                targetPath = (Resolve-Path -LiteralPath $TargetPath).Path
                publishFormat = $PublishFormat
                bytes = if (Test-Path -LiteralPath $TargetPath) { (Get-Item -LiteralPath $TargetPath).Length } else { 0 }
            })
        }
        'get-binary' {
            if (-not $PageId) { throw '-PageId is required.' }
            if (-not $CallbackId) { throw '-CallbackId is required. Use a binary object callbackID from page XML.' }
            $app = Get-OneNoteApp
            $base64 = ''
            $app.GetBinaryPageContent($PageId, $CallbackId, [ref]$base64)
            Write-Json ([pscustomobject]@{
                ok = $true
                pageId = $PageId
                callbackId = $CallbackId
                base64 = $base64
                length = $base64.Length
            })
        }
        'demo-write' {
            $app = Get-OneNoteApp
            $targetSectionPath = if ($SectionPath) {
                $SectionPath
            }
            else {
                Join-Path (Join-Path $env:TEMP 'OneNoteDesktopSkill') 'DirectComTest.one'
            }

            $targetDir = Split-Path -Parent $targetSectionPath
            if ($targetDir) {
                New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
            }

            $sectionId = ''
            $app.OpenHierarchy($targetSectionPath, '', [ref]$sectionId, 3)
            $pageId = ''
            $app.CreateNewPage($sectionId, [ref]$pageId)
            $xml = ''
            $app.GetPageContent($pageId, [ref]$xml)
            $body = if ($Text) { $Text } else { 'Desktop OneNote direct COM write succeeded.' }
            $pageTitle = if ($Title) { $Title } else { 'OneNote direct COM test' }
            $updated = New-OneNotePageXml -ExistingXml $xml -NewTitle $pageTitle -BodyText $body
            $app.UpdatePageContent($updated)
            Write-Json ([pscustomobject]@{
                ok = $true
                route = 'desktop-com'
                sectionPath = $targetSectionPath
                sectionId = $sectionId
                pageId = $pageId
                title = $pageTitle
                text = $body
            })
        }
        'control-page' {
            $app = Get-OneNoteApp
            $targetSectionPath = if ($SectionPath) {
                $SectionPath
            }
            else {
                Join-Path (Join-Path $env:TEMP 'OneNoteDesktopSkill') 'AgentControlSurface.one'
            }

            $targetDir = Split-Path -Parent $targetSectionPath
            if ($targetDir) {
                New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
            }

            $sectionId = ''
            $app.OpenHierarchy($targetSectionPath, '', [ref]$sectionId, 3)

            $hierarchyXml = ''
            $app.GetHierarchy('', 4, [ref]$hierarchyXml)
            [xml]$hierarchyDoc = $hierarchyXml
            $ns = New-Object System.Xml.XmlNamespaceManager($hierarchyDoc.NameTable)
            $ns.AddNamespace('one', 'http://schemas.microsoft.com/office/onenote/2013/onenote')
            $sectionCount = $hierarchyDoc.SelectNodes('//one:Section', $ns).Count
            $pageCount = $hierarchyDoc.SelectNodes('//one:Page', $ns).Count
            $recentPages = @($hierarchyDoc.SelectNodes('//one:Page', $ns) |
                Sort-Object { $_.lastModifiedTime } -Descending |
                Select-Object -First 8 |
                ForEach-Object { "- $($_.name) [$($_.ID)]" })

            $appPath = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\ONENOTE.EXE' -ErrorAction SilentlyContinue
            $exe = if ($appPath) { Join-Path $appPath.Path 'ONENOTE.EXE' } else { 'not found' }
            $version = if (($exe -ne 'not found') -and (Test-Path $exe)) { (Get-Item $exe).VersionInfo.ProductVersion } else { 'unknown' }
            $now = Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'

            $blocks = @(
                "STATUS`nDesktop OneNote COM control is live.`nCreated: $now`nPowerShell apartment: $([Threading.Thread]::CurrentThread.GetApartmentState())`nOneNote: $exe`nVersion: $version",
                "CURRENT AUTOMATION SURFACE`nOpen sections visible through COM: $sectionCount`nPages visible through COM: $pageCount`nControl section: $targetSectionPath`nControl section ID: $sectionId",
                "RECENT PAGES`n$(if ($recentPages.Count) { $recentPages -join "`n" } else { "- No pages were visible before this control page was created." })",
                "AGENT COMMANDS TO TRY NEXT`n1. List open notebooks: Invoke-OneNoteCom.ps1 -Operation hierarchy -Scope notebooks`n2. Search this workspace: Invoke-OneNoteCom.ps1 -Operation find-pages -Query OneNote`n3. Read this page: Invoke-OneNoteCom.ps1 -Operation get-page -PageId <page-id>`n4. Patch page XML with update-page-xml after preserving IDs/namespaces.",
                "WHY THIS IS INTERESTING`nThis page was generated without UI automation, without a localhost server, and without a OneNote add-in. The agent created a local .one section, queried OneNote hierarchy, composed page XML, and asked desktop OneNote to materialize the result through COM."
            )

            $pageId = ''
            $app.CreateNewPage($sectionId, [ref]$pageId)
            $xml = ''
            $app.GetPageContent($pageId, [ref]$xml)
            $pageTitle = if ($Title -and $Title -ne 'Untitled') { $Title } else { 'OneNote COM Control Surface' }
            $updated = New-OneNoteControlPageXml -ExistingXml $xml -NewTitle $pageTitle -Blocks $blocks
            $app.UpdatePageContent($updated)

            Write-Json ([pscustomobject]@{
                ok = $true
                route = 'desktop-com'
                operation = 'control-page'
                sectionPath = $targetSectionPath
                sectionId = $sectionId
                pageId = $pageId
                title = $pageTitle
                sectionCount = $sectionCount
                pageCountBeforeCreate = $pageCount
            })
        }
        'extract-media' {
            if (-not $PageId) { throw '-PageId is required.' }
            $targetOutputDir = if ($OutputDir) {
                $OutputDir
            }
            else {
                Join-Path (Get-Location).Path 'onenote-media-export'
            }
            New-Item -ItemType Directory -Force -Path $targetOutputDir | Out-Null

            $app = Get-OneNoteApp
            $xml = ''
            $app.GetPageContent($PageId, [ref]$xml)
            [xml]$doc = $xml
            $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
            $ns.AddNamespace('one', 'http://schemas.microsoft.com/office/onenote/2013/onenote')
            $mediaNodes = $doc.SelectNodes('//one:MediaFile', $ns)

            $items = @()
            $index = 0
            foreach ($node in $mediaNodes) {
                $index += 1
                $preferredName = $node.GetAttribute('preferredName')
                $pathSource = $node.GetAttribute('pathSource')
                $pathCache = $node.GetAttribute('pathCache')
                $mediaRef = $node.SelectSingleNode('one:MediaReference', $ns)
                $mediaId = if ($mediaRef) { $mediaRef.GetAttribute('mediaID') } else { '' }
                $sourcePath = if ($pathSource -and (Test-Path -LiteralPath $pathSource)) {
                    $pathSource
                }
                elseif ($pathCache -and (Test-Path -LiteralPath $pathCache)) {
                    $pathCache
                }
                else {
                    ''
                }

                $safeName = Get-SafeFileName $preferredName
                if (-not $safeName) {
                    $safeName = "onenote-media-$index.bin"
                }
                $destination = if ($sourcePath) { Join-Path $targetOutputDir $safeName } else { '' }
                if ($sourcePath) {
                    Copy-Item -LiteralPath $sourcePath -Destination $destination -Force
                }

                $items += [pscustomobject]@{
                    mediaId = $mediaId
                    preferredName = $preferredName
                    pathSource = $pathSource
                    sourceExists = [bool]($pathSource -and (Test-Path -LiteralPath $pathSource))
                    pathCache = $pathCache
                    cacheExists = [bool]($pathCache -and (Test-Path -LiteralPath $pathCache))
                    copiedFrom = $sourcePath
                    copiedTo = $destination
                    bytes = if ($destination -and (Test-Path -LiteralPath $destination)) { (Get-Item -LiteralPath $destination).Length } else { 0 }
                }
            }

            Write-Json ([pscustomobject]@{
                ok = $true
                pageId = $PageId
                outputDir = (Resolve-Path -LiteralPath $targetOutputDir).Path
                count = $items.Count
                media = $items
            })
        }
    }
}
catch {
    Write-Json ([pscustomobject]@{
        ok = $false
        error = $_.Exception.Message
        hresult = ('{0:X8}' -f ($_.Exception.HResult -band 0xffffffff))
    })
    exit 1
}
