# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

WordZero is a Go library for creating, reading, and modifying Word (.docx) documents following the Office Open XML (OOXML) specification. Minimal dependencies — only goldmark for Markdown support.

Module: `github.com/mr-pmillz/wordZero`
Go version: 1.19+ (toolchain go1.24.11)

## Build & Test Commands

```bash
go build ./...                          # Build all
go test ./...                           # Run all tests
go test ./pkg/document ./pkg/style ./test  # Test specific packages
go test -v ./test -run TestTemplateName # Run a single test
go test -cover ./...                    # Coverage report
go test -short ./...                    # Skip long-running tests
go vet ./...                            # Static analysis
go fmt ./...                            # Format code
```

## Architecture

Three packages under `pkg/`:

- **`pkg/document`** — Core library. `Document` is the central type, created via `document.New()` or `document.Open(path)`. Contains paragraphs, tables, images, headers/footers, footnotes, TOC, numbering, template engine, and page settings. Key files:
  - `document.go` — Document struct, Body, Paragraph, Run, BodyElement interface, core CRUD methods
  - `table.go` — Table creation, cell merging, styling, iterators
  - `template.go` + `template_engine.go` — Template system with `{{variable}}`, `{{#if}}`, `{{#each}}`, `{{extends}}`, `{{#block}}`
  - `image.go` — Image embedding with relationship management
  - `header_footer.go` — Headers/footers (default, first page, even pages)
  - `footnotes.go` — Footnotes and endnotes
  - `toc.go` — Table of contents with bookmarks
  - `numbering.go` — Ordered/unordered/multi-level lists
  - `page.go` — Page size, orientation, margins, section properties
  - `errors.go` — Custom error types with `WrapError`/`WrapErrorWithContext`

- **`pkg/style`** — Style management. `StyleManager` handles 18 predefined styles plus custom styles with inheritance (BasedOn, Next chain).

- **`pkg/markdown`** — Markdown-to-Word conversion via goldmark. `MarkdownConverter` interface with `ConvertString`, `ConvertFile`, `BatchConvert`. `WordRenderer` translates goldmark AST to Document. Supports GFM, footnotes, LaTeX math.

## Key Design Decisions

- **Fluent API**: Methods return receivers for chaining (e.g., `para.SetAlignment(...).SetSpacing(...).SetStyle(...)`)
- **Custom XML marshaling**: OOXML requires strict element ordering. `Body`, `Run`, `Relationships`, etc. implement custom `MarshalXML`. Never rely on default struct marshaling for these types.
- **BodyElement interface**: Implemented by `Paragraph` and `Table`. Don't manipulate `Body.Elements` directly — use the provided Document methods.
- **Relationship management**: Images, headers, footers use relationship IDs stored in `DocumentRelationships`. Adding media requires creating a relationship entry.
- **OOXML units**: Measurements are in TWIPs (1 point = 20 TWIPs). Some APIs accept millimeters and convert internally.
- **Two relationship collections**: `d.relationships` → root `_rels/.rels` (package-level). `d.documentRelationships` → `word/_rels/document.xml.rels` (document-level). Footnotes, endnotes, settings, images, headers/footers go in `documentRelationships`. Only the officeDocument relationship goes in root.
- **Bilingual logging**: `messages.go` defines `MsgKey` constants with ZH/EN maps. Use `DebugMsg`/`InfoMsgf` etc. for new log calls. `SetGlobalLanguage(LogLanguageEN)` switches to English. Default is English.

## OOXML Gotchas

- **Debugging DOCX output**: Unzip the `.docx` and inspect XML. Compare against Word-repaired output to find issues: `unzip -o file.docx -d /tmp/debug/`
- **Footnote separators**: Must use `<w:separator/>` (id=-1) + `<w:continuationSeparator/>` (id=0), not `<w:footnoteRef/>`
- **Footnote body references**: Only set `<w:rStyle>` on reference runs in document.xml — do NOT add `<w:vertAlign>` (the style provides superscript)
- **Footnote content**: Each footnote paragraph needs `<w:pStyle val="FootnoteText"/>`, a self-ref run with `<w:footnoteRef/>`, then a text run with `xml:space="preserve"`
- **Per-document state**: `FootnoteManager` is per-Document (not global). Each Document tracks its own footnote IDs independently.

## Code Conventions

- **Comments in English**, variable/function names in English
- PascalCase for exported, camelCase for unexported
- Config structs for complex operations (`TableConfig`, `SpacingConfig`, `PageSettings`)
- OOXML namespace prefixes: `w:` (WordML), `r:` (relationships), `a:` (DrawingML)
- Test output goes to `test_output/` directories — always defer cleanup with `os.RemoveAll`

## Test Organization

- **Unit tests**: `pkg/document/*_test.go`, `pkg/style/*_test.go`
- **Integration tests**: `test/*_test.go` (20 files covering tables, templates, markdown, page settings, etc.)
- **Examples**: `examples/` directory (27 runnable examples)
- **Benchmarks**: `benchmark/` directory
