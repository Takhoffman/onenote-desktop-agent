# OneNote Media And Audio

OneNote audio recordings can appear in page XML as `one:MediaFile` elements. The raw audio may live in `pathSource` or `pathCache`.

## Extraction Pattern

Use the helper first:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\scripts\Invoke-OneNoteCom.ps1 -Operation extract-media -PageId "<page-id>" -OutputDir ".\onenote-media-export"
```

The operation:

- Reads page XML.
- Finds `//one:MediaFile`.
- Captures `preferredName`, `pathSource`, `pathCache`, and nested `mediaID`.
- Copies from `pathSource` if it exists, otherwise from `pathCache`.
- Returns copied file paths and byte counts as JSON.

## Observed Audio XML

```xml
<one:MediaFile
  pathCache="%TEMP%\{GUID}.bin"
  pathSource="%LOCALAPPDATA%\Microsoft\OneNote\<version>\Audio Cache\Page name.wma"
  preferredName="Page name.wma">
  <one:MediaReference mediaID="{GUID}" />
</one:MediaFile>
```

## Conversion

Use ffmpeg when available:

```powershell
ffmpeg -y -i ".\onenote-media-export\Page name.wma" -ar 16000 -ac 1 ".\onenote-media-export\Page name.wav"
```

OneNote desktop recordings may be WMA/ASF files using codecs such as `wmavoice`. Convert to WAV before sending to transcription tools unless the transcriber accepts the extracted format directly.
