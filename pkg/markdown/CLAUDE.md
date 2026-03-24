# CLAUDE.md — pkg/markdown

Markdown-to-Word and Word-to-Markdown conversion using the goldmark parser.

## Key Types

- `MarkdownConverter` — Interface with `ConvertString`, `ConvertBytes`, `ConvertFile`, `BatchConvert`.
- `Converter` — Implementation. Created via `NewConverter(options)`.
- `WordRenderer` — Goldmark AST renderer that produces Word `Document` output.
- `Exporter` — Word-to-Markdown conversion.
- `BidirectionalConverter` — Auto-detects direction and converts.

## Dependencies

- `github.com/yuin/goldmark` — Markdown parser with extensions
- `github.com/litao91/goldmark-mathjax` — LaTeX math formula support

## Goldmark Extensions

Enabled via `ConvertOptions`:
- GFM (tables, strikethrough, autolinks, task lists)
- Footnotes
- Math formulas (`$...$` inline, `$$...$$` block) — rendered with Cambria Math font

## Conversion Mapping

- Markdown headings → Heading1-9 styles
- Code blocks → CodeBlock paragraph style
- Inline code → CodeChar character style
- Lists → Word numbering system
- Tables → Word tables with cell content
- Math → Unicode math symbols (not all LaTeX operators supported)

## Limitations

- Word-to-Markdown loses Word-specific formatting (field codes, complex styles)
- Image export during word→markdown requires separate file handling
- Complex table layouts may need manual adjustment after conversion
