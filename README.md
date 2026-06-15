# OneNote Desktop Agent

OneNote Desktop Agent is a Claude Code and Codex plugin that teaches agents how to automate Microsoft OneNote desktop on Windows through the OneNote COM API and Windows PowerShell.

It is for local desktop OneNote, not OneNote on the web. The skill can inspect installation state, list notebooks and pages, read and update page XML, create pages, navigate OneNote, export pages, and extract embedded audio/media when OneNote exposes media file paths.

## Naming

Recommended public repository name: `onenote-desktop-agent`.

Stable plugin and skill name: `onenote-desktop`.

The repo name describes the package as a cross-agent helper. The skill name stays short because users and agents will invoke it directly.

## What It Includes

- `skills/onenote-desktop/SKILL.md`: shared skill instructions for Claude Code and Codex.
- `skills/onenote-desktop/scripts/Invoke-OneNoteCom.ps1`: main PowerShell COM helper.
- `skills/onenote-desktop/scripts/Probe-OneNoteCom.ps1`: COM surface and smoke-test probe.
- `skills/onenote-desktop/scripts/Probe-OneNoteStorage.ps1`: read-only storage/cache inventory helper.
- `skills/onenote-desktop/references/`: notes for raw COM, page XML, media extraction, and storage boundaries.
- `.claude-plugin/plugin.json`: Claude Code plugin manifest.
- `.codex-plugin/plugin.json`: Codex plugin manifest.

## Requirements

- Windows.
- Microsoft OneNote desktop with COM automation available.
- Windows PowerShell via `powershell.exe`.
- For OneNote COM calls, use STA PowerShell:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\skills\onenote-desktop\scripts\Invoke-OneNoteCom.ps1 -Operation check-install
```

## Claude Code Install

After publishing this repository to GitHub:

```powershell
claude plugin marketplace add takhoffman/onenote-desktop-agent
claude plugin install onenote-desktop@takhoffman
```

For local testing from the repository root:

```powershell
claude plugin install .
```

Start a fresh Claude Code session and ask:

```text
Use the onenote-desktop skill to check my OneNote desktop COM install.
```

## Codex Install

After publishing this repository to GitHub, add it as a Codex plugin marketplace/source using the Codex plugin UI or CLI, then install `onenote-desktop`.

For local desktop testing, copy or archive the committed package into the Codex plugin cache and restart Codex if needed:

```powershell
$target = "$env:USERPROFILE\.codex\plugins\cache\personal\onenote-desktop\0.1.0"
New-Item -ItemType Directory -Force -Path $target | Out-Null
git archive --format=tar HEAD | tar -x -C $target
```

## Direct Helper Usage

Check OneNote and COM:

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

Export a page to PDF:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\skills\onenote-desktop\scripts\Invoke-OneNoteCom.ps1 -Operation publish -PageId "<page-id>" -TargetPath ".\exports\page.pdf" -PublishFormat pdf
```

Extract media from a page:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\skills\onenote-desktop\scripts\Invoke-OneNoteCom.ps1 -Operation extract-media -PageId "<page-id>" -OutputDir ".\media-export"
```

## Safety And Privacy

This plugin is designed for local desktop automation. It does not require a localhost server, browser automation, or an Office.js add-in.

Runtime output can include notebook names, page IDs, local paths, note text, exported documents, and media files. The repository intentionally ignores built packages, probe output, notebook section files, exported documents, and extracted media so private notebook data is not staged accidentally.

The storage probe is read-only. Normal workflows should prefer OneNote COM and page XML over raw cache or `.one` file inspection.

## Status

This is an early public package. The helper scripts cover common non-destructive OneNote COM operations, but OneNote XML edits still require care: read the existing page XML first, preserve IDs and namespaces, update only the intended content, and verify by reading the page again.
