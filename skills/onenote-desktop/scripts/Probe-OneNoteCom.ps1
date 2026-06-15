param(
    [switch]$SmokeTest,
    [string]$OutputDir = '',
    [string]$ProbeSectionPath = ''
)

$ErrorActionPreference = 'Stop'

function Write-Json {
    param([object]$Value)
    $Value | ConvertTo-Json -Depth 100
}

function Invoke-ProbeStep {
    param(
        [string]$Name,
        [scriptblock]$Script
    )

    try {
        $value = & $Script
        [pscustomobject]@{
            name = $Name
            ok = $true
            result = $value
        }
    }
    catch {
        [pscustomobject]@{
            name = $Name
            ok = $false
            error = $_.Exception.Message
            hresult = ('{0:X8}' -f ($_.Exception.HResult -band 0xffffffff))
        }
    }
}

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    throw 'Run this script with powershell.exe -NoProfile -STA.'
}

Add-Type -AssemblyName Microsoft.Office.Interop.OneNote
$assembly = [Microsoft.Office.Interop.OneNote.Application].Assembly
$app = New-Object -ComObject OneNote.Application

$types = @($assembly.GetExportedTypes() | Sort-Object FullName | ForEach-Object {
    [pscustomobject]@{
        fullName = $_.FullName
        isEnum = $_.IsEnum
        isInterface = $_.IsInterface
        isClass = $_.IsClass
    }
})

$enums = @($assembly.GetExportedTypes() | Where-Object { $_.IsEnum } | Sort-Object FullName | ForEach-Object {
    $enumType = $_
    [pscustomobject]@{
        name = $enumType.FullName
        values = @([enum]::GetValues($enumType) | ForEach-Object {
            [pscustomobject]@{
                name = $_.ToString()
                value = [int]$_
            }
        })
    }
})

$methods = @([Microsoft.Office.Interop.OneNote.IApplication].GetMethods() | Sort-Object Name | ForEach-Object {
    [pscustomobject]@{
        name = $_.Name
        returnType = $_.ReturnType.FullName
        parameters = @($_.GetParameters() | ForEach-Object {
            [pscustomobject]@{
                name = $_.Name
                type = $_.ParameterType.FullName
                isOut = $_.IsOut
                isOptional = $_.IsOptional
            }
        })
    }
})

$smoke = @()
if ($SmokeTest) {
    $targetOutputDir = if ($OutputDir) {
        $OutputDir
    }
    else {
        Join-Path (Get-Location).Path 'onenote-com-probe'
    }
    New-Item -ItemType Directory -Force -Path $targetOutputDir | Out-Null

    $targetSectionPath = if ($ProbeSectionPath) {
        $ProbeSectionPath
    }
    else {
        Join-Path (Join-Path $env:TEMP 'OneNoteDesktopSkill') 'ComProbe.one'
    }
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $targetSectionPath) | Out-Null

    $state = [ordered]@{}

    $smoke += Invoke-ProbeStep 'GetHierarchy notebooks' {
        $xml = ''
        $app.GetHierarchy('', [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsNotebooks, [ref]$xml)
        @{ length = $xml.Length; hasXml = $xml.StartsWith('<?xml') }
    }

    $smoke += Invoke-ProbeStep 'GetSpecialLocation default notebook folder' {
        $path = ''
        $app.GetSpecialLocation([Microsoft.Office.Interop.OneNote.SpecialLocation]::slDefaultNotebookFolder, [ref]$path)
        @{ path = $path }
    }

    $smoke += Invoke-ProbeStep 'OpenHierarchy create local section' {
        $sectionId = ''
        $app.OpenHierarchy($targetSectionPath, '', [ref]$sectionId, [Microsoft.Office.Interop.OneNote.CreateFileType]::cftSection)
        $state.sectionId = $sectionId
        @{ sectionPath = $targetSectionPath; sectionId = $sectionId }
    }

    $smoke += Invoke-ProbeStep 'CreateNewPage' {
        $pageId = ''
        $app.CreateNewPage($state.sectionId, [ref]$pageId, [Microsoft.Office.Interop.OneNote.NewPageStyle]::npsBlankPageWithTitle)
        $state.pageId = $pageId
        @{ pageId = $pageId }
    }

    $smoke += Invoke-ProbeStep 'GetPageContent basic' {
        $xml = ''
        $app.GetPageContent($state.pageId, [ref]$xml, [Microsoft.Office.Interop.OneNote.PageInfo]::piBasic)
        $state.pageXml = $xml
        @{ length = $xml.Length; hasPage = $xml -like '*<one:Page*' }
    }

    $smoke += Invoke-ProbeStep 'UpdatePageContent title/body' {
        [xml]$doc = $state.pageXml
        $ns = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
        $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $nsm.AddNamespace('one', $ns)
        $title = $doc.SelectSingleNode('/one:Page/one:Title/one:OE/one:T', $nsm)
        if ($title) {
            $title.RemoveAll()
            [void]$title.AppendChild($doc.CreateCDataSection('OneNote COM Probe'))
        }
        $page = $doc.SelectSingleNode('/one:Page', $nsm)
        $outline = $doc.CreateElement('one', 'Outline', $ns)
        $pos = $doc.CreateElement('one', 'Position', $ns)
        $pos.SetAttribute('x', '36.0')
        $pos.SetAttribute('y', '86.4')
        $pos.SetAttribute('z', '0')
        $children = $doc.CreateElement('one', 'OEChildren', $ns)
        $oe = $doc.CreateElement('one', 'OE', $ns)
        $t = $doc.CreateElement('one', 'T', $ns)
        [void]$t.AppendChild($doc.CreateCDataSection('Smoke test page created by Probe-OneNoteCom.ps1.'))
        [void]$oe.AppendChild($t)
        [void]$children.AppendChild($oe)
        [void]$outline.AppendChild($pos)
        [void]$outline.AppendChild($children)
        [void]$page.AppendChild($outline)
        $app.UpdatePageContent($doc.OuterXml, [DateTime]::MinValue)
        @{ updated = $true }
    }

    $smoke += Invoke-ProbeStep 'FindPages' {
        $xml = ''
        $app.FindPages($state.sectionId, 'OneNote COM Probe', [ref]$xml, $true, $false)
        @{ length = $xml.Length; foundProbe = $xml -like '*OneNote COM Probe*' }
    }

    $smoke += Invoke-ProbeStep 'FindMeta' {
        $xml = ''
        $app.FindMeta($state.sectionId, 'name', [ref]$xml, $true)
        @{ length = $xml.Length }
    }

    $smoke += Invoke-ProbeStep 'GetHierarchyParent' {
        $parentId = ''
        $app.GetHierarchyParent($state.pageId, [ref]$parentId)
        @{ pageId = $state.pageId; parentId = $parentId }
    }

    $smoke += Invoke-ProbeStep 'GetHyperlinkToObject' {
        $link = ''
        $app.GetHyperlinkToObject($state.pageId, '', [ref]$link)
        @{ hyperlink = $link }
    }

    $smoke += Invoke-ProbeStep 'Publish PDF page' {
        $target = Join-Path $targetOutputDir 'onenote-com-probe.pdf'
        $app.Publish($state.pageId, $target, [Microsoft.Office.Interop.OneNote.PublishFormat]::pfPDF, '')
        @{ targetPath = $target; exists = (Test-Path -LiteralPath $target); bytes = if (Test-Path -LiteralPath $target) { (Get-Item -LiteralPath $target).Length } else { 0 } }
    }
}

Write-Json ([pscustomobject]@{
    ok = $true
    generatedAt = (Get-Date).ToString('o')
    apartment = [Threading.Thread]::CurrentThread.GetApartmentState().ToString()
    assembly = [pscustomobject]@{
        fullName = $assembly.FullName
        location = $assembly.Location
    }
    types = $types
    enums = $enums
    methods = $methods
    smokeTests = $smoke
})
