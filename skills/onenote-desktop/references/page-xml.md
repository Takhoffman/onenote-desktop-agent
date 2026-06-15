# OneNote Page XML

OneNote COM reads and writes page content as XML. Always inspect a target page with `get-page` before modifying it.

## Preservation Rules

- Preserve the `one` namespace: `http://schemas.microsoft.com/office/onenote/2013/onenote`.
- Preserve page IDs, object IDs, author metadata, timestamps, media references, tags, and unknown elements unless the task explicitly requires changing them.
- Add new outlines instead of rewriting the whole page when possible.
- Use CDATA for text in `one:T`.
- After editing XML, call `UpdatePageContent` and then read the page again to verify.

## Common Structure

```xml
<one:Page ...>
  <one:Title>
    <one:OE>
      <one:T><![CDATA[Title text]]></one:T>
    </one:OE>
  </one:Title>
  <one:Outline>
    <one:Position x="36.0" y="86.4" z="0" />
    <one:Size width="468.0" height="40.0" />
    <one:OEChildren>
      <one:OE>
        <one:T><![CDATA[Body text]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>
```

## Multi-Outline Pages

Create separate `one:Outline` elements with different `one:Position` y values. This is useful for dashboards, transcripts, extracted media summaries, and agent logs.

## Update Flow

1. Read page XML with `GetPageContent`.
2. Parse as XML with a namespace manager.
3. Append or modify targeted nodes.
4. Save XML to a temp file if using `update-page-xml`.
5. Call `UpdatePageContent`.
6. Re-read and verify text or element presence.
