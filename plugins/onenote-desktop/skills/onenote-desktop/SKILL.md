---
name: onenote-desktop
description: Control Microsoft OneNote desktop through direct Windows PowerShell COM automation, including raw one-off COM scripts. Use when an agent needs to inspect OneNote installation state, list notebooks/sections/pages, read or write page XML, create or update pages, navigate OneNote, search content, extract audio/media recordings, export data, or automate desktop OneNote without browser automation, localhost servers, or Office.js add-ins.
version: 0.1.0
---

# OneNote Desktop

Use desktop OneNote COM directly. Prefer `scripts/Invoke-OneNoteCom.ps1` for known operations, but do not treat it as a limit: when the task needs a capability the helper lacks, write a focused one-off PowerShell COM script and run it with Windows PowerShell STA.

Always use this shell shape for COM:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -Command '<OneNote COM code>'
```

or:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\scripts\Invoke-OneNoteCom.ps1 <args>
```

Do not use `pwsh` for OneNote COM unless you just smoke-tested the exact call. Office COM is apartment-sensitive; `powershell.exe -STA` is the reliable baseline.

## Quick Start

Check desktop OneNote and COM:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\scripts\Invoke-OneNoteCom.ps1 -Operation check-install
```

List visible notebooks, sections, and pages:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\scripts\Invoke-OneNoteCom.ps1 -Operation hierarchy -Scope pages
```

Read a page:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\scripts\Invoke-OneNoteCom.ps1 -Operation get-page -PageId "<page-id>"
```

Export a page to PDF:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\scripts\Invoke-OneNoteCom.ps1 -Operation publish -PageId "<page-id>" -TargetPath ".\exports\page.pdf" -PublishFormat pdf
```

Create a page:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\scripts\Invoke-OneNoteCom.ps1 -Operation create-page -SectionId "<section-id>" -Title "Agent note" -Text "Written by an agent through OneNote COM."
```

Extract audio/media from a page:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\scripts\Invoke-OneNoteCom.ps1 -Operation extract-media -PageId "<page-id>" -OutputDir ".\media-export"
```

## Workflow

1. Run `check-install`.
2. Run `hierarchy -Scope pages` and identify the target notebook/section/page.
3. If hierarchy is empty, COM is working but no notebooks are open in this OneNote profile/session. Ask the user to open the notebook or create a local section with `demo-write`.
4. Use helper operations for routine tasks.
5. Use raw COM for anything not yet wrapped. Read `references/raw-com.md` before writing new COM code.
6. For page edits, read existing XML first and preserve namespaces, IDs, timestamps, object IDs, and unknown elements.

## References

Read `references/raw-com.md` when writing direct `New-Object -ComObject OneNote.Application` scripts, probing methods, calling enums, navigating OneNote, publishing/exporting, or doing anything not already wrapped by `Invoke-OneNoteCom.ps1`.

Read `references/page-xml.md` before modifying page XML, generating multi-outline pages, preserving objects, or adding rich content.

Read `references/media.md` when extracting OneNote audio recordings or other embedded media from `MediaFile` elements.

Read `references/storage.md` for last-resort raw storage/cache inventory, `.one`/`.onepkg`/cache file signatures, and boundaries for probing storage without mutating OneNote internals.

## Script Operations

`check-install`: report the detected desktop OneNote executable, version, STA state, and COM creation status.

`hierarchy`: call `GetHierarchy` and return raw XML plus parsed nodes. Scopes are `self`, `children`, `notebooks`, `sections`, and `pages`.

`get-page`: call `GetPageContent` for a page ID and return raw XML plus parsed nodes.

`create-page`: create a page in an existing section, set a title, and optionally add body text.

`update-page-xml`: update a page from a prepared XML file. Use only after preserving OneNote XML structure.

`find-pages`: call OneNote search with `-Query`.

`find-meta`: call OneNote metadata search with `-Query`.

`get-parent`: return the parent hierarchy ID for a page, section, or object ID.

`get-hyperlink`: return a OneNote hyperlink for a hierarchy object and optional page content object.

`get-special-location`: return OneNote special folders such as default notebook folder, backup folder, or unfiled notes section.

`navigate`: navigate OneNote to a hierarchy object and optional page content object.

`publish`: export a notebook/section/page to `one`, `onepkg`, `mhtml`, `pdf`, `xps`, `word`, or `emf`.

`get-binary`: fetch binary page content by callback ID from page XML.

`demo-write`: create/open a local `.one` section, create a page, and write test content.

`control-page`: create/open a local `.one` section and write a multi-outline control page with install status, hierarchy counts, recent page IDs, and next commands.

`extract-media`: scan a page for `MediaFile` XML elements, copy audio/media from OneNote source/cache paths into `-OutputDir`, and return copied file paths plus media IDs.

`scripts/Probe-OneNoteCom.ps1`: enumerate the installed OneNote interop assembly and optionally run safe smoke tests. Use it before adding wrappers for less common COM methods.

`scripts/Probe-OneNoteStorage.ps1`: inventory likely OneNote storage/cache locations read-only, report signatures, sizes, and optional hashes, and support `-ExtraRoot` for local notebook folders.

## Release Notes

- Target desktop OneNote for Windows, not OneNote on the web.
- Detect the installed executable and version with `check-install`; do not assume a fixed Office path or version.
- The desktop app may be branded "Microsoft OneNote" even when it is the Office16 desktop product formerly known as OneNote 2016.
- Audio/media extraction depends on OneNote exposing `MediaFile` paths in page XML; use `extract-media` and report when neither `pathSource` nor `pathCache` exists locally.
- Raw `.one` and cache probing is a last resort. Prefer COM/page XML first, then use `Probe-OneNoteStorage.ps1` for read-only discovery.
- Do not use a localhost HTTP server for v1 desktop automation.
