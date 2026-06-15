<p align="center">
  <img src="assets/readme-banner.svg" alt="OneNote Desktop Agent - Claude Code and Codex automation for Microsoft OneNote desktop" width="100%" />
</p>

# OneNote Desktop Agent

OneNote has years of notes, meetings, research, and recordings.

Agents should be able to use that context without asking you to copy and paste it by hand.

**OneNote Desktop Agent** connects Claude Code and Codex to the real Microsoft OneNote desktop app on Windows. It uses the local OneNote COM API through Windows PowerShell. No Microsoft Graph. No browser bridge. No Office.js add-in.

The plugin and skill name is `onenote-desktop`.

## Install

### Claude Code

```powershell
claude plugin marketplace add takhoffman/onenote-desktop-agent
claude plugin install onenote-desktop@takhoffman
```

Then ask Claude Code:

```text
Use the onenote-desktop skill to check my OneNote desktop COM install.
```

For local development from this repository:

```powershell
claude plugin install .
```

### Codex

```powershell
codex plugin marketplace add https://github.com/takhoffman/onenote-desktop-agent
codex plugin add onenote-desktop@takhoffman
```

Then ask Codex:

```text
Use the onenote-desktop skill to list my visible OneNote notebooks.
```

For Codex Desktop, open Plugins, add this marketplace URL, then install `onenote-desktop`:

```text
https://github.com/takhoffman/onenote-desktop-agent
```

## What It Gives Agents

- A map of visible notebooks, section groups, sections, and pages.
- Page XML, including IDs, timestamps, tags, media references, author metadata, and unknown elements.
- Page creation and targeted page updates.
- Search over indexed page text and metadata.
- Durable OneNote links to pages and objects.
- Desktop navigation to a page or object.
- Export to PDF, XPS, Word, MHTML, EMF, `.one`, or `.onepkg`.
- Embedded audio and media extraction when OneNote exposes source/cache paths.

## Why Desktop

This project is for **Microsoft OneNote desktop for Windows**.

That matters because many useful OneNote workflows are local: desktop notebooks, local cache, embedded recordings, exported pages, and the exact app state the user is looking at. Cloud APIs are useful, but they are not the same thing as controlling the desktop app that already has the notebook open.

OneNote Desktop Agent keeps that boundary clear. It automates the local app. It does not create a hosted service, a browser session, a localhost bridge, or a Microsoft Graph permission flow.

## Requirements

- Windows.
- Microsoft OneNote desktop with COM automation available.
- Windows PowerShell via `powershell.exe`.
- Claude Code or Codex if installing as an agent plugin.

OneNote COM calls should run under single-threaded apartment PowerShell:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\skills\onenote-desktop\scripts\Invoke-OneNoteCom.ps1 -Operation check-install
```

## Use It Directly

You can run the PowerShell helper without Claude Code or Codex.

Check OneNote desktop and COM:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\skills\onenote-desktop\scripts\Invoke-OneNoteCom.ps1 -Operation check-install
```

List visible notebooks, sections, and pages:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\skills\onenote-desktop\scripts\Invoke-OneNoteCom.ps1 -Operation hierarchy -Scope pages
```

Read a page:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\skills\onenote-desktop\scripts\Invoke-OneNoteCom.ps1 -Operation get-page -PageId "<page-id>"
```

Create a page:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\skills\onenote-desktop\scripts\Invoke-OneNoteCom.ps1 -Operation create-page -SectionId "<section-id>" -Title "Agent note" -Text "Written through OneNote COM."
```

Export a page to PDF:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\skills\onenote-desktop\scripts\Invoke-OneNoteCom.ps1 -Operation publish -PageId "<page-id>" -TargetPath ".\exports\page.pdf" -PublishFormat pdf
```

Extract embedded media from a page:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\skills\onenote-desktop\scripts\Invoke-OneNoteCom.ps1 -Operation extract-media -PageId "<page-id>" -OutputDir ".\media-export"
```

## COM Operations

The helper wraps the repeatable, lower-risk OneNote desktop COM operations:

- `GetHierarchy`: inventory notebooks, section groups, sections, and pages.
- `GetPageContent`: read page XML for inspection, extraction, and targeted edits.
- `CreateNewPage` and `UpdatePageContent`: create pages and update prepared XML.
- `FindPages` and `FindMeta`: search indexed page text and metadata.
- `GetHierarchyParent`: climb from a page or object ID back to its parent.
- `GetHyperlinkToObject`: create durable OneNote links for pages or page objects.
- `GetSpecialLocation`: resolve default notebook, backup, and unfiled-notes locations.
- `NavigateTo`: bring OneNote desktop to a target page or object.
- `Publish`: export notebooks, sections, or pages to supported desktop formats.
- `GetBinaryPageContent`: retrieve binary content referenced by callback IDs in page XML.
- `OpenHierarchy`: open or create local notebook/section files for controlled tests.

The desktop interop also exposes higher-risk calls such as `CloseNotebook`, `DeleteHierarchy`, `DeletePageContent`, `OpenPackage`, and `UpdateHierarchy`. Those are intentionally not promoted as default workflows. Use raw COM for them only when the target IDs are verified and the user explicitly asks for that operation.

## Safety And Privacy

Runtime output can include private notebook names, page IDs, local file paths, note text, exported documents, and extracted media. This repository ignores built packages, probe output, OneNote section files, exported documents, and media output so private notebook data is not staged accidentally.

Normal automation should prefer OneNote COM and page XML over raw cache or `.one` file inspection.

For page edits, read existing XML first, preserve namespaces and IDs, update only the intended content, then read the page again to verify.

## Package Layout

- `.claude-plugin/plugin.json`: Claude Code plugin manifest.
- `.claude-plugin/marketplace.json`: Claude Code marketplace descriptor.
- `.agents/plugins/marketplace.json`: Codex marketplace descriptor.
- `.codex-plugin/plugin.json`: root Codex plugin manifest.
- `plugins/onenote-desktop/`: installable Codex marketplace package.
- `skills/onenote-desktop/SKILL.md`: shared skill instructions.
- `skills/onenote-desktop/scripts/Invoke-OneNoteCom.ps1`: main OneNote COM helper.
- `skills/onenote-desktop/scripts/Probe-OneNoteCom.ps1`: COM method and enum probe.
- `skills/onenote-desktop/scripts/Probe-OneNoteStorage.ps1`: read-only OneNote storage/cache inventory probe.
- `skills/onenote-desktop/references/`: notes for raw COM, page XML, media extraction, and storage boundaries.

The root `skills/onenote-desktop/` directory is the source of truth. The Codex marketplace package under `plugins/onenote-desktop/` carries a synced copy for package installation. After changing the root skill, run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\sync-plugin-package.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\check-plugin-package.ps1
```

## Status

This is an early public package. The helper scripts cover common non-destructive OneNote COM operations. Destructive operations such as deleting notebooks, sections, pages, or page content are intentionally not wrapped as default workflows.

## Search Keywords

Microsoft OneNote desktop automation, OneNote COM API, OneNote PowerShell automation, Claude Code plugin, Codex plugin, OneNote page XML, local OneNote automation, Windows OneNote scripting, export OneNote to PDF, extract OneNote audio recordings.

## License

MIT License. See `LICENSE`.
