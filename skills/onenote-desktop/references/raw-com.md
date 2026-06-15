# Raw OneNote COM

Use this reference when the helper script does not expose the needed capability. Raw COM means writing a direct STA PowerShell script against `OneNote.Application` and the XML returned by the OneNote COM API.

## Rules

- Run via `powershell.exe -NoProfile -STA`, not `pwsh`.
- Return JSON for anything another agent may chain.
- Prefer small scripts that do one operation and print one JSON result.
- Treat COM errors as useful probes. Capture `Exception.Message` and HRESULT.
- If an enum conversion fails, load `Microsoft.Office.Interop.OneNote` and inspect enum names/values.

## Minimal Pattern

```powershell
powershell.exe -NoProfile -STA -Command '
$ErrorActionPreference = "Stop"
$app = New-Object -ComObject OneNote.Application
$xml = ""
$app.GetHierarchy("", 4, [ref]$xml)
[pscustomobject]@{ ok = $true; xml = $xml } | ConvertTo-Json -Depth 20
'
```

## Useful Calls

Create the COM object:

```powershell
$app = New-Object -ComObject OneNote.Application
```

Get hierarchy:

```powershell
$xml = ""
$app.GetHierarchy("", 4, [ref]$xml) # 4 = pages
```

Read a page:

```powershell
$xml = ""
$app.GetPageContent($pageId, [ref]$xml)
```

Create a page:

```powershell
$pageId = ""
$app.CreateNewPage($sectionId, [ref]$pageId)
```

Open or create a local section:

```powershell
$sectionId = ""
$app.OpenHierarchy($sectionPath, "", [ref]$sectionId, 3) # 3 = cftSection
```

Update a page:

```powershell
$app.UpdatePageContent($pageXml)
```

Navigate OneNote:

```powershell
$app.NavigateTo($pageId)
```

Search:

```powershell
$xml = ""
$app.FindPages("", "meeting notes", [ref]$xml)
```

Publish/export:

```powershell
$app.Publish($pageId, ".\exports\page.pdf", 3, "") # 3 = pfPDF
```

Get a OneNote link:

```powershell
$link = ""
$app.GetHyperlinkToObject($pageId, "", [ref]$link)
```

Get a parent object:

```powershell
$parentId = ""
$app.GetHierarchyParent($pageId, [ref]$parentId)
```

Get a special location:

```powershell
$path = ""
$app.GetSpecialLocation(2, [ref]$path) # 2 = default notebook folder
```

Fetch binary page content by callback ID:

```powershell
$base64 = ""
$app.GetBinaryPageContent($pageId, $callbackId, [ref]$base64)
```

## Probe Suite

Run the probe when adapting to a new OneNote installation or before expanding wrappers:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\scripts\Probe-OneNoteCom.ps1
```

Run safe smoke tests against a temporary local section:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\scripts\Probe-OneNoteCom.ps1 -SmokeTest -OutputDir .\probe-output
```

The smoke tests are intentionally non-destructive: they create a temporary local section/page, update that page, search it, generate a hyperlink, and export it to PDF. They do not delete, close notebooks, or alter existing user notebooks.

## Installed Interop Surface

The OneNote desktop interop assembly exposes these `IApplication` methods:

- `CloseNotebook`
- `CreateNewPage`
- `DeleteHierarchy`
- `DeletePageContent`
- `FindMeta`
- `FindPages`
- `GetBinaryPageContent`
- `GetHierarchy`
- `GetHierarchyParent`
- `GetHyperlinkToObject`
- `GetPageContent`
- `GetSpecialLocation`
- `NavigateTo`
- `OpenHierarchy`
- `OpenPackage`
- `Publish`
- `UpdateHierarchy`
- `UpdatePageContent`

Wrap stable, commonly useful, non-destructive methods in `Invoke-OneNoteCom.ps1`. Keep destructive or uncommon methods as raw COM unless a user explicitly asks for them and the calling pattern is proven.

## Enum Inspection

```powershell
Add-Type -AssemblyName Microsoft.Office.Interop.OneNote
[enum]::GetValues([Microsoft.Office.Interop.OneNote.HierarchyScope]) |
  ForEach-Object { [pscustomobject]@{ Name = $_.ToString(); Value = [int]$_ } }
```

Known hierarchy scopes used successfully:

- `0`: self
- `1`: children
- `2`: notebooks
- `3`: sections
- `4`: pages

Known create file types used successfully:

- `0`: none
- `1`: notebook
- `2`: folder
- `3`: section

Known page info values:

- `0`: basic
- `1`: binary data
- `2`: selection
- `3`: binary data plus selection

Known publish formats:

- `0`: OneNote section
- `1`: OneNote package
- `2`: MHTML
- `3`: PDF
- `4`: XPS
- `5`: Word
- `6`: EMF

Known special locations:

- `0`: backup folder
- `1`: unfiled notes section
- `2`: default notebook folder

Known new page styles:

- `0`: default
- `1`: blank page with title
- `2`: blank page without title

## Wrapper Boundary

Do not try to turn every COM method into a permanent helper operation. The helper should cover repeatable operations that are easy to validate. For everything else, write a task-specific raw COM script using the debug pattern below, run it once, and only promote it into `Invoke-OneNoteCom.ps1` after it proves broadly useful.

## Debug Pattern

```powershell
try {
  $app = New-Object -ComObject OneNote.Application
  # COM call here
} catch {
  [pscustomobject]@{
    ok = $false
    error = $_.Exception.Message
    hresult = ("{0:X8}" -f ($_.Exception.HResult -band 0xffffffff))
  } | ConvertTo-Json
  exit 1
}
```
