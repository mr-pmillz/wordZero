# CLAUDE.md — pkg/document

Core library for creating, reading, and modifying Word (.docx) documents.

## Key Types

- `Document` — Central type. Created via `New()` or `Open(path)`. Holds `Body`, relationships, parts, style manager, footnote manager.
- `Body` — Contains `Elements []interface{}` (Paragraphs and Tables). Has custom `MarshalXML` for element ordering.
- `Paragraph` — Has `Properties *ParagraphProperties` and `Runs []Run`.
- `Run` — Smallest text unit. Has custom `MarshalXML` — new fields must be added there too.
- `Table` / `TableRow` / `TableCell` — Table structures with merging support.

## Adding New XML Elements to Run

The `Run` struct uses custom `MarshalXML`. When adding a new field:
1. Add the field to the `Run` struct
2. Add serialization in `Run.MarshalXML()` in the correct OOXML order
3. Define the element type with proper `xml:"w:elementName"` tag

## Relationship Architecture

Two separate relationship collections — using the wrong one causes Word corruption:
- `d.relationships` → `_rels/.rels` (root level, only officeDocument ref)
- `d.documentRelationships` → `word/_rels/document.xml.rels` (footnotes, endnotes, images, headers, footers, settings)

## State Management

Per-document state avoids cross-document leaks:
- `footnoteManager` — tracks footnote/endnote IDs and content per Document
- `nextImageID` — image counter per Document
- `parts map[string][]byte` — raw XML parts for the ZIP archive

## Logging

Uses `messages.go` message keys for bilingual output. For new log calls:
- No args: `DebugMsg(MsgSomeKey)` / `InfoMsg(MsgSomeKey)`
- With args: `DebugMsgf(MsgSomeKey, arg1, arg2)` / `InfoMsgf(MsgSomeKey, arg)`
- Add new keys to `messages.go` in both `messagesZH` and `messagesEN` maps

## File Layout

- `document.go` — Document, Body, Paragraph, Run, RunProperties, XML types, MarshalXML, core methods (~4000 lines)
- `table.go` — Table CRUD, merging, styling, iterators, nested tables, cell content
- `template.go` + `template_engine.go` — Template system (`{{var}}`, `{{#if}}`, `{{#each}}`, `{{extends}}`)
- `footnotes.go` — Footnotes/endnotes with proper OOXML references, settings management
- `image.go` — Image embedding, floating images, positioning
- `header_footer.go` — Headers/footers, content types helper
- `messages.go` — Bilingual log message catalog (MsgKey constants + ZH/EN maps)
- `logger.go` — Logger with language switching support
