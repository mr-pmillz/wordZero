package markdown

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"github.com/mr-pmillz/wordZero/pkg/document"
)

// ExportOptions configures export options.
type ExportOptions struct {
	// Basic configuration
	UseGFMTables       bool // use GFM table syntax
	PreserveFootnotes  bool // preserve footnotes
	PreserveLineBreaks bool // preserve line breaks
	WrapLongLines      bool // auto-wrap long lines
	MaxLineLength      int  // maximum line length

	// Image handling
	ExtractImages     bool   // whether to export image files
	ImageOutputDir    string // image output directory
	ImageNamePattern  string // image naming pattern
	ImageRelativePath bool   // use relative paths for image references

	// Link handling
	PreserveBookmarks bool // preserve bookmarks as anchor links
	ConvertHyperlinks bool // convert hyperlinks

	// Code block handling
	PreserveCodeStyle bool   // preserve code style
	DefaultCodeLang   string // default code language identifier

	// Style mapping
	CustomStyleMap      map[string]string // custom style mapping
	IgnoreUnknownStyles bool              // ignore unknown styles

	// Content handling
	PreserveTOC     bool // preserve table of contents
	IncludeMetadata bool // include document metadata
	StripComments   bool // remove comments

	// Formatting options
	UseSetext        bool   // use Setext-style headings
	BulletListMarker string // bullet list marker
	EmphasisMarker   string // emphasis marker

	// Error handling
	StrictMode    bool        // strict mode
	IgnoreErrors  bool        // ignore conversion errors
	ErrorCallback func(error) // error callback

	// Progress reporting
	ProgressCallback func(int, int) // progress callback
}

// MarkdownWriter outputs content in Markdown format.
type MarkdownWriter struct {
	opts      *ExportOptions
	doc       *document.Document
	output    strings.Builder
	imageNum  int
	footnotes []string
}

// Write generates Markdown content.
func (w *MarkdownWriter) Write() ([]byte, error) {
	// Process document metadata
	if w.opts.IncludeMetadata {
		w.writeMetadata()
	}

	// Iterate over document paragraphs
	if w.doc.Body != nil {
		for _, para := range w.doc.Body.GetParagraphs() {
			err := w.writeParagraph(para)
			if err != nil {
				if w.opts.ErrorCallback != nil {
					w.opts.ErrorCallback(err)
				}
				if !w.opts.IgnoreErrors {
					return nil, err
				}
			}
		}

		// Process tables
		for _, table := range w.doc.Body.GetTables() {
			err := w.writeTable(table)
			if err != nil {
				if w.opts.ErrorCallback != nil {
					w.opts.ErrorCallback(err)
				}
				if !w.opts.IgnoreErrors {
					return nil, err
				}
			}
		}
	}

	// Add footnotes
	if w.opts.PreserveFootnotes && len(w.footnotes) > 0 {
		w.writeFootnotes()
	}

	return []byte(w.output.String()), nil
}

// writeMetadata writes document metadata.
func (w *MarkdownWriter) writeMetadata() {
	w.output.WriteString("---\n")
	w.output.WriteString("title: \"Document\"\n")
	w.output.WriteString("---\n\n")
}

// writeParagraph writes a paragraph.
func (w *MarkdownWriter) writeParagraph(para *document.Paragraph) error {
	if para == nil {
		return nil
	}

	// Check paragraph style
	style := w.getParagraphStyle(para)

	switch {
	case strings.HasPrefix(style, "Heading"):
		return w.writeHeading(para, style)
	case style == "Quote":
		return w.writeQuote(para)
	case style == "CodeBlock":
		return w.writeCodeBlock(para)
	case w.isListParagraph(para):
		return w.writeListItem(para)
	default:
		return w.writeNormalParagraph(para)
	}
}

// writeHeading writes a heading.
func (w *MarkdownWriter) writeHeading(para *document.Paragraph, style string) error {
	level := w.getHeadingLevel(style)
	if level > 6 {
		level = 6
	}

	text := w.extractParagraphText(para)
	if strings.TrimSpace(text) == "" {
		return nil
	}

	if w.opts.UseSetext && level <= 2 {
		// Use Setext style
		w.output.WriteString(text + "\n")
		if level == 1 {
			w.output.WriteString(strings.Repeat("=", len(text)) + "\n\n")
		} else {
			w.output.WriteString(strings.Repeat("-", len(text)) + "\n\n")
		}
	} else {
		// Use ATX style
		w.output.WriteString(strings.Repeat("#", level) + " " + text + "\n\n")
	}

	return nil
}

// writeQuote writes a blockquote.
func (w *MarkdownWriter) writeQuote(para *document.Paragraph) error {
	text := w.extractParagraphText(para)
	if strings.TrimSpace(text) == "" {
		return nil
	}

	lines := strings.Split(text, "\n")
	for _, line := range lines {
		w.output.WriteString("> " + line + "\n")
	}
	w.output.WriteString("\n")

	return nil
}

// writeCodeBlock writes a code block.
func (w *MarkdownWriter) writeCodeBlock(para *document.Paragraph) error {
	text := w.extractParagraphText(para)
	if strings.TrimSpace(text) == "" {
		return nil
	}

	lang := w.opts.DefaultCodeLang
	w.output.WriteString("```" + lang + "\n")
	w.output.WriteString(text + "\n")
	w.output.WriteString("```\n\n")

	return nil
}

// writeListItem writes a list item.
func (w *MarkdownWriter) writeListItem(para *document.Paragraph) error {
	text := w.extractParagraphText(para)
	if strings.TrimSpace(text) == "" {
		return nil
	}

	// Simple list item handling
	marker := w.opts.BulletListMarker
	if w.isNumberedList(para) {
		marker = "1."
	}

	w.output.WriteString(marker + " " + text + "\n")

	return nil
}

// writeNormalParagraph writes a normal paragraph.
func (w *MarkdownWriter) writeNormalParagraph(para *document.Paragraph) error {
	text := w.extractParagraphText(para)
	if strings.TrimSpace(text) == "" {
		w.output.WriteString("\n")
		return nil
	}

	// Handle long line wrapping
	if w.opts.WrapLongLines && len(text) > w.opts.MaxLineLength {
		text = w.wrapText(text, w.opts.MaxLineLength)
	}

	w.output.WriteString(text + "\n\n")

	return nil
}

// writeTable writes a table.
func (w *MarkdownWriter) writeTable(table *document.Table) error {
	if table == nil || len(table.Rows) == 0 {
		return nil
	}

	if !w.opts.UseGFMTables {
		return w.writeSimpleTable(table)
	}

	rows := table.Rows

	// Write header row
	if len(rows) > 0 {
		headerRow := rows[0]
		w.output.WriteString("|")
		for _, cell := range headerRow.Cells {
			text := w.extractCellText(&cell)
			w.output.WriteString(" " + text + " |")
		}
		w.output.WriteString("\n")

		// Write separator row
		w.output.WriteString("|")
		for i := 0; i < len(headerRow.Cells); i++ {
			w.output.WriteString("-----|")
		}
		w.output.WriteString("\n")

		// Write data rows
		for i := 1; i < len(rows); i++ {
			w.output.WriteString("|")
			for _, cell := range rows[i].Cells {
				text := w.extractCellText(&cell)
				w.output.WriteString(" " + text + " |")
			}
			w.output.WriteString("\n")
		}
	}

	w.output.WriteString("\n")

	return nil
}

// writeSimpleTable writes a simple table format.
func (w *MarkdownWriter) writeSimpleTable(table *document.Table) error {
	for i, row := range table.Rows {
		if i == 0 {
			w.output.WriteString("**")
		}
		for j, cell := range row.Cells {
			if j > 0 {
				w.output.WriteString(" | ")
			}
			text := w.extractCellText(&cell)
			w.output.WriteString(text)
		}
		if i == 0 {
			w.output.WriteString("**")
		}
		w.output.WriteString("\n")
	}
	w.output.WriteString("\n")

	return nil
}

// writeFootnotes writes footnotes.
func (w *MarkdownWriter) writeFootnotes() {
	w.output.WriteString("\n---\n\n")
	for i, footnote := range w.footnotes {
		fmt.Fprintf(&w.output, "[^%d]: %s\n", i+1, footnote)
	}
}

// extractParagraphText extracts text from a paragraph.
func (w *MarkdownWriter) extractParagraphText(para *document.Paragraph) string {
	if para == nil {
		return ""
	}

	var result strings.Builder

	for _, run := range para.Runs {
		text := w.formatRunText(&run)
		result.WriteString(text)
	}

	return result.String()
}

// formatRunText formats a text run.
func (w *MarkdownWriter) formatRunText(run *document.Run) string {
	if run == nil {
		return ""
	}

	text := run.Text.Content
	if text == "" {
		return ""
	}

	// Check formatting properties
	if run.Properties != nil {
		// Check for bold
		if run.Properties.Bold != nil {
			if run.Properties.Italic != nil {
				text = "***" + text + "***" // bold italic
			} else {
				text = "**" + text + "**" // bold
			}
		} else if run.Properties.Italic != nil {
			text = w.opts.EmphasisMarker + text + w.opts.EmphasisMarker // italic
		}

		// Check for strikethrough
		if run.Properties.Strike != nil {
			text = "~~" + text + "~~" // strikethrough
		}

		// Handle code style
		if w.isCodeStyle(run.Properties) {
			text = "`" + text + "`"
		}
	}

	return text
}

// extractCellText extracts text from a table cell.
func (w *MarkdownWriter) extractCellText(cell *document.TableCell) string {
	if cell == nil {
		return ""
	}

	var result strings.Builder

	for _, para := range cell.Paragraphs {
		text := w.extractParagraphText(&para)
		result.WriteString(text)
	}

	// Clean up line breaks in table cell text
	text := result.String()
	text = strings.ReplaceAll(text, "\n", " ")
	text = strings.TrimSpace(text)

	return text
}

// getParagraphStyle returns the paragraph style name.
func (w *MarkdownWriter) getParagraphStyle(para *document.Paragraph) string {
	if para.Properties != nil && para.Properties.ParagraphStyle != nil {
		return para.Properties.ParagraphStyle.Val
	}
	return "Normal"
}

// getHeadingLevel returns the heading level from a style name.
func (w *MarkdownWriter) getHeadingLevel(style string) int {
	// Extract digit
	re := regexp.MustCompile(`\d+`)
	matches := re.FindString(style)
	if matches != "" {
		if level, err := strconv.Atoi(matches); err == nil {
			return level
		}
	}
	return 1
}

// isListParagraph checks whether the paragraph is a list item.
func (w *MarkdownWriter) isListParagraph(para *document.Paragraph) bool {
	if para.Properties == nil {
		return false
	}
	return para.Properties.NumberingProperties != nil
}

// isNumberedList checks whether the paragraph is a numbered list item.
func (w *MarkdownWriter) isNumberedList(para *document.Paragraph) bool {
	// Simple implementation; should check numbering format in practice
	return false
}

// isCodeStyle checks whether the run properties indicate a code style.
func (w *MarkdownWriter) isCodeStyle(props *document.RunProperties) bool {
	if props.FontFamily != nil {
		font := props.FontFamily.ASCII
		// Check if it is a monospace font
		codefonts := []string{"Consolas", "Courier New", "Monaco", "Menlo", "Source Code Pro"}
		for _, codefont := range codefonts {
			if strings.Contains(font, codefont) {
				return true
			}
		}
	}
	return false
}

// wrapText wraps text to the specified maximum line length.
func (w *MarkdownWriter) wrapText(text string, maxLength int) string {
	if len(text) <= maxLength {
		return text
	}

	var result strings.Builder
	words := strings.Fields(text)
	var line strings.Builder

	for _, word := range words {
		if line.Len()+len(word)+1 > maxLength {
			if line.Len() > 0 {
				result.WriteString(line.String() + "\n")
				line.Reset()
			}
		}
		if line.Len() > 0 {
			line.WriteString(" ")
		}
		line.WriteString(word)
	}

	if line.Len() > 0 {
		result.WriteString(line.String())
	}

	return result.String()
}
