# OneNote Desktop Codex Plugin

This plugin teaches Codex agents how to automate Microsoft OneNote desktop on Windows through the OneNote COM API and Windows PowerShell.

It provides one skill, `onenote-desktop`, with helper scripts for checking OneNote COM availability, listing notebooks/sections/pages, reading and updating page XML, exporting pages, navigating OneNote, and extracting embedded media when OneNote exposes media file paths.

The repository intentionally ignores built packages, runtime probes, notebook section files, exported documents, and extracted media because those outputs can include notebook names, page IDs, local paths, or note content.
