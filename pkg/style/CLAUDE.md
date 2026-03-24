# CLAUDE.md — pkg/style

Style management system for Word documents. Independent package but integrates with `Document.GetStyleManager()`.

## Key Types

- `StyleManager` — Central manager. Created via `NewStyleManager()`, auto-loads predefined styles. Methods: `GetStyle`, `AddStyle`, `GetAllStyles`, `GetStyleWithInheritance`.
- `Style` — Style definition with `StyleID`, `Type`, `Name`, `BasedOn`, `Next`, `ParagraphPr`, `RunPr`.
- `StyleType` — `paragraph`, `character`, `table`, `numbering`.

## Predefined Styles

Auto-registered on `NewStyleManager()` creation:
- Normal, Heading1-9, Title, Subtitle, Emphasis, Strong, Quote, CodeBlock, CodeChar, ListParagraph
- TOC styles (toc 1-9, TOCHeading)
- FootnoteReference (character), FootnoteText (paragraph), EndnoteReference, EndnoteText

## Style Inheritance

`BasedOn` field creates inheritance chains. `GetStyleWithInheritance()` merges parent properties automatically. When modifying styles, preserve inheritance — don't flatten.

## Adding New Predefined Styles

Add to `addSpecialStyles()` in `style.go`. Follow existing pattern:
```go
myStyle := &Style{
    Type:    string(StyleTypeParagraph),
    StyleID: "MyStyle",
    Name:    &StyleName{Val: "my style"},
    BasedOn: &BasedOn{Val: "Normal"},
    // ...properties
}
sm.AddStyle(myStyle)
```

## Conventions

- Colors: hex strings without `#` prefix (e.g., `"FF0000"`)
- Font sizes: string values in half-points (e.g., `"24"` = 12pt, `"20"` = 10pt)
- `RunProperties` here mirrors but differs slightly from `document.RunProperties` — both packages can be used independently
