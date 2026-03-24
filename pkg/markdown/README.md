# WordZero Markdown Conversion Package

The `pkg/markdown` package provides bidirectional conversion between Markdown and Word documents.

## Features

### Markdown to Word Conversion
- Based on the goldmark parsing engine
- Supports GitHub Flavored Markdown (GFM)
- Supports headings, formatted text, lists, tables, images, links, and more
- Configurable conversion options

### Word to Markdown Conversion
- Export Word documents back to Markdown
- Preserves document structure and formatting
- Supports image export
- Multiple export configuration options

## Basic Usage

### Word to Markdown Conversion

```go
package main

import (
    "fmt"
    "github.com/mr-pmillz/wordZero/pkg/markdown"
)

func main() {
    // Create an exporter
    exporter := markdown.NewExporter(markdown.DefaultExportOptions())

    // Export a Word document to Markdown
    err := exporter.ExportToFile("document.docx", "output.md", nil)
    if err != nil {
        fmt.Printf("Export failed: %v\n", err)
        return
    }

    fmt.Println("Word document successfully converted to Markdown!")
}
```

### Markdown to Word Conversion

```go
package main

import (
    "fmt"
    "github.com/mr-pmillz/wordZero/pkg/markdown"
)

func main() {
    // Create a converter
    converter := markdown.NewConverter(markdown.DefaultOptions())

    // Convert Markdown to a Word document
    err := converter.ConvertFile("input.md", "output.docx", nil)
    if err != nil {
        fmt.Printf("Conversion failed: %v\n", err)
        return
    }

    fmt.Println("Markdown successfully converted to Word document!")
}
```

### Bidirectional Converter

```go
package main

import (
    "fmt"
    "github.com/mr-pmillz/wordZero/pkg/markdown"
)

func main() {
    // Create a bidirectional converter
    converter := markdown.NewBidirectionalConverter(
        markdown.DefaultOptions(),      // Markdown to Word options
        markdown.DefaultExportOptions(), // Word to Markdown options
    )

    // Automatically detect file type and convert
    err := converter.AutoConvert("input.docx", "output.md")
    if err != nil {
        fmt.Printf("Conversion failed: %v\n", err)
        return
    }

    fmt.Println("Document conversion complete!")
}
```

## Advanced Configuration

### Word to Markdown Export Options

```go
options := &markdown.ExportOptions{
    UseGFMTables:      true,  // Use GitHub Flavored Markdown tables
    ExtractImages:     true,  // Export image files
    ImageOutputDir:    "images/", // Image output directory
    PreserveFootnotes: true,  // Preserve footnotes
    UseSetext:         true,  // Use Setext-style headings
    IncludeMetadata:   true,  // Include document metadata
    ProgressCallback: func(current, total int) {
        fmt.Printf("Progress: %d/%d\n", current, total)
    },
}

exporter := markdown.NewExporter(options)
```

### Markdown to Word Conversion Options

```go
options := &markdown.ConvertOptions{
    EnableGFM:         true,     // Enable GitHub Flavored Markdown
    EnableFootnotes:   true,     // Enable footnote support
    EnableTables:      true,     // Enable table support
    EnableMath:        true,     // Enable math formula support (LaTeX syntax)
    DefaultFontFamily: "Calibri", // Default font
    DefaultFontSize:   11.0,     // Default font size
    GenerateTOC:       true,     // Generate table of contents
    TOCMaxLevel:       3,        // Maximum TOC depth
}

converter := markdown.NewConverter(options)
```

## Supported Conversion Mappings

### Word to Markdown

| Word Element | Markdown Syntax | Description |
|--------------|----------------|-------------|
| Heading1-6 | `# Heading` | Heading levels map directly |
| Bold | `**bold**` | Text formatting |
| Italic | `*italic*` | Text formatting |
| Strikethrough | `~~strikethrough~~` | Text formatting |
| Code | `` `code` `` | Inline code |
| Code Block | ```` code block ```` | Code block |
| Hyperlink | `[link](url)` | Link conversion |
| Image | `![image](src)` | Image reference |
| Table | `\| table \|` | GFM table |
| List | `- item` | List item |

### Markdown to Word

| Markdown Syntax | Word Element | Implementation |
|----------------|--------------|----------------|
| `# Heading` | Heading1 style | `AddHeadingParagraph()` |
| `**bold**` | Bold formatting | `RunProperties.Bold` |
| `*italic*` | Italic formatting | `RunProperties.Italic` |
| `` `code` `` | Code style | Monospace font |
| `[link](url)` | Hyperlink | `AddHyperlink()` |
| `![image](src)` | Image | `AddImageFromFile()` |
| `\| table \|` | Word table | `AddTable()` |
| `- list` | Bullet list | `AddBulletList()` |
| `$formula$` | Math formula | Cambria Math font |
| `$$formula$$` | Block-level math formula | Centered display |

## Batch Conversion

```go
// Batch Markdown to Word
converter := markdown.NewConverter(markdown.DefaultOptions())
inputs := []string{"doc1.md", "doc2.md", "doc3.md"}
err := converter.BatchConvert(inputs, "output/", nil)

// Batch Word to Markdown
exporter := markdown.NewExporter(markdown.DefaultExportOptions())
inputs := []string{"doc1.docx", "doc2.docx", "doc3.docx"}
err := exporter.BatchExport(inputs, "markdown/", nil)
```

## Error Handling

```go
options := &markdown.ExportOptions{
    StrictMode: true,  // Strict mode
    IgnoreErrors: false, // Do not ignore errors
    ErrorCallback: func(err error) {
        fmt.Printf("Conversion error: %v\n", err)
    },
}
```

## Compatibility Notes

- This package is fully compatible with the existing `pkg/document` package
- Does not modify any existing APIs
- Integrates seamlessly with existing code
- Supports all existing Word document operations

## Important Considerations

1. Word-to-Markdown conversion will lose certain Word-specific formatting information
2. Complex table layouts may require manual adjustment
3. Images need to be handled separately during export
4. Some Word styles have no direct Markdown equivalent
5. Math formula conversion uses Unicode characters and the Cambria Math font, supporting common LaTeX syntax

## Math Formula Support

### Inline Formulas
Wrap with single dollar signs: `$E = mc^2$`

### Block-Level Formulas
Wrap with double dollar signs:
```
$$
x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}
$$
```

### Supported LaTeX Syntax
- Greek letters: `\alpha`, `\beta`, `\gamma`, `\pi`, `\sigma`, etc.
- Operators: `\times`, `\div`, `\pm`, `\leq`, `\geq`, `\neq`, etc.
- Superscripts and subscripts: `x^2`, `x_i`, `x^{n+1}`, `x_{i,j}`, etc.
- Fractions: `\frac{a}{b}`
- Square roots: `\sqrt{x}`, `\sqrt[3]{x}`
- Special symbols: `\infty`, `\sum`, `\int`, `\partial`, `\nabla`, etc.
- Arrows: `\rightarrow`, `\leftarrow`, `\Rightarrow`, etc.

## Roadmap

- [x] Math formula support
- [ ] Mermaid diagram conversion
- [ ] Better nested list support
- [ ] Custom style mapping
- [ ] Command-line tool
