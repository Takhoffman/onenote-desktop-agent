# OneNote Storage And Cache

Use this reference only for last-resort discovery, recovery, forensics, and debugging. Prefer COM and page XML for normal automation.

## Storage Layers

- Local notebooks are folders containing `.one` section files and often `.onetoc2` table-of-contents files.
- OneNote desktop keeps local data under `%LOCALAPPDATA%\Microsoft\OneNote\<version>\`.
- Common desktop subfolders include `cache`, `Backup`, `Audio Cache`, `FullTextSearchIndex`, `MasterIndex`, and `ServerListings`.
- OneNote audio recordings may be copied from page XML `MediaFile` paths; use `references/media.md` first.
- OneNote for Windows/Store builds may use a package-local cache under `%LOCALAPPDATA%\Packages\Microsoft.Office.OneNote_8wekyb3d8bbwe\LocalCache`.

## Raw Format

Microsoft documents the `.one` file format as MS-ONE: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-one/73d22548-a613-4350-8c23-07d15576be50. It is a binary persistence format for sections/pages and can contain text, images, tables, note tags, and other content. It is powerful but far more complex than COM page XML, so treat direct parsing as a last resort.

Observed signatures:

- `.one` / `.onetoc2` may begin with `E4 52 5C 7B`.
- `.onepkg` may be ZIP-like and begin with `50 4B 03 04`.
- WMA/ASF recordings may begin with `30 26 B2 75`.
- Cache `.bin` files are implementation details; inventory them read-only.

## Probe

Run a read-only storage inventory:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\Probe-OneNoteStorage.ps1 -OutputDir .\storage-probe
```

The probe lists likely OneNote roots, file extensions, sizes, timestamps, and file signatures. It does not parse or modify cache files.

Probe an additional local notebook or temporary section folder:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\Probe-OneNoteStorage.ps1 -ExtraRoot "<local-notebook-root>" -OutputDir .\storage-probe
```

Add `-IncludeHashes` when comparing snapshots:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\Probe-OneNoteStorage.ps1 -IncludeHashes -OutputDir .\storage-probe
```

## Boundaries

- Do not write to cache files.
- Do not delete cache files from an agent workflow.
- Do not rely on cache naming as stable API.
- Use `Publish`, `GetPageContent`, `GetBinaryPageContent`, and `MediaFile` extraction before raw file parsing.
- If direct `.one` parsing is required, base work on the MS-ONE specification and keep it as a separate parser/probe rather than mixing it into routine COM operations.
