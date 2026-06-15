<p align="center">
  <img src="assets/readme-banner.svg" alt="OneNote Desktop Agent - Claude Code and Codex automation for Microsoft OneNote desktop" width="100%" />
</p>

# OneNote Desktop Agent

Your notes are already in OneNote.

Your AI assistant should be able to work with them.

**OneNote Desktop Agent** connects Claude Code and Codex to Microsoft OneNote desktop on Windows. It uses the local OneNote COM API through Windows PowerShell, so an agent can see notebooks, read pages, create notes, export content, search, navigate, and extract media without a browser, without Microsoft Graph, and without a OneNote add-in.

It is built for **Microsoft OneNote desktop for Windows**. Not OneNote web. Not a cloud sync API. The real desktop app.

## The Problem

OneNote is full of useful information, but to an AI assistant it is usually invisible.

You can paste a page into chat. You can export a PDF by hand. You can click through notebooks and copy IDs one at a time. But that is not automation. That is you doing the integration work.

OneNote Desktop Agent gives agents a clean local path into OneNote desktop:

- discover the notebook structure
- read page XML
- create and update pages
- search page text and metadata
- export pages and sections
- open OneNote at the exact object being discussed
- extract embedded audio and media when OneNote exposes the paths

The stable plugin and skill name is `onenote-desktop`. The repository is `onenote-desktop-agent`.

## What It Makes Possible

### Give agents real notebook context

List notebooks, sections, section groups, and pages without clicking through the OneNote UI. Build a machine-readable map before cleanup, migration, archiving, review, or export.

### Work with page XML directly

Read OneNote pages as XML so an agent can inspect structure instead of guessing from screenshots or copied text. Preserve namespaces, IDs, timestamps, tags, media references, author metadata, and unknown elements while making targeted edits.

### Create useful notes from scripts

Generate pages for meeting notes, research summaries, logs, checklists, status pages, or agent output. Create a local test `.one` section when you want a reproducible workflow that does not touch existing notebooks.

### Search and link precisely

Search page text, search metadata, resolve parent hierarchy IDs, and generate durable OneNote hyperlinks. When an agent finds something, it can navigate the desktop app to that exact page or object.

### Export without the clipboard

Publish notebooks, sections, or pages to PDF, XPS, Word, MHTML, EMF, `.one`, or `.onepkg`. Collect page IDs first, then export the exact pages you need.

### Pull out recordings and media

Inspect page XML for `MediaFile` entries and copy embedded audio or media from OneNote source/cache paths when available. Useful for transcription, review, archiving, and recovery.

### Debug local OneNote safely

Check whether OneNote desktop is installed, whether COM can be created, whether PowerShell is running in STA mode, which notebooks are visible, where OneNote stores default notebooks/backups/unfiled notes, and what cache or backup files exist. The storage probe is read-only.

### Keep private notes local

This is local desktop automation. There is no hosted service, no browser session, no localhost bridge, no Office.js add-in, and no Microsoft Graph permission flow.

## What The COM Surface Unlocks

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

## Install In Claude Code

```powershell
claude plugin marketplace add takhoffman/onenote-desktop-agent
claude plugin install onenote-desktop@takhoffman
```

Test it in a fresh Claude Code session:

```text
Use the onenote-desktop skill to check my OneNote desktop COM install.
```

For local development from the repository root:

```powershell
claude plugin install .
```

## Install In Codex

```powershell
codex plugin marketplace add https://github.com/takhoffman/onenote-desktop-agent
codex plugin add onenote-desktop@takhoffman
```

In a new Codex session, ask:

```text
Use the onenote-desktop skill to list my visible OneNote notebooks.
```

For Codex Desktop, open Plugins, add this marketplace URL, then install `onenote-desktop`:

```text
https://github.com/takhoffman/onenote-desktop-agent
```

## Use It Directly

You can also run the PowerShell helper without Claude Code or Codex.

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

## Requirements

- Windows.
- Microsoft OneNote desktop with COM automation available.
- Windows PowerShell via `powershell.exe`.
- Claude Code or Codex if installing as an agent plugin.

OneNote COM calls should run under single-threaded apartment PowerShell:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\skills\onenote-desktop\scripts\Invoke-OneNoteCom.ps1 -Operation check-install
```

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

## Safety And Privacy

Runtime output can include private notebook names, page IDs, local file paths, note text, exported documents, and extracted media. This repository ignores built packages, probe output, OneNote section files, exported documents, and media output so private notebook data is not staged accidentally.

Normal automation should prefer OneNote COM and page XML over raw cache or `.one` file inspection.

For page edits, read existing XML first, preserve namespaces and IDs, update only the intended content, then read the page again to verify.

## Search Keywords

Microsoft OneNote desktop automation, OneNote COM API, OneNote PowerShell automation, Claude Code plugin, Codex plugin, OneNote page XML, local OneNote automation, Windows OneNote scripting, export OneNote to PDF, extract OneNote audio recordings.

## Status

This is an early public package. The helper scripts cover common non-destructive OneNote COM operations; destructive operations such as deleting notebooks, sections, pages, or page content are intentionally not wrapped as default workflows.

## License

MIT License. See `LICENSE`.
