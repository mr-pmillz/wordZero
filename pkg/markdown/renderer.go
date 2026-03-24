package markdown

import (
	"path/filepath"
	"regexp"
	"strings"

	"github.com/zerx-lab/wordZero/pkg/document"
	"github.com/yuin/goldmark/ast"

	// goldmark extension AST node support
	extast "github.com/yuin/goldmark/extension/ast"

	// math formula support
	mathjax "github.com/litao91/goldmark-mathjax"
)

// WordRenderer renders Markdown AST nodes into a Word document.
type WordRenderer struct {
	doc       *document.Document
	opts      *ConvertOptions
	source    []byte
	listLevel int // current list nesting level
}

// Render walks the AST and renders it into a Word document.
func (r *WordRenderer) Render(doc ast.Node) error {
	return ast.Walk(doc, func(node ast.Node, entering bool) (ast.WalkStatus, error) {
		if !entering {
			return ast.WalkContinue, nil
		}

		switch n := node.(type) {
		case *ast.Document:
			// Document root node, continue processing children
			return ast.WalkContinue, nil

		case *ast.Heading:
			return r.renderHeading(n)

		case *ast.Paragraph:
			return r.renderParagraph(n)

		case *ast.List:
			return r.renderList(n)

		case *ast.ListItem:
			return r.renderListItem(n)

		case *ast.Blockquote:
			return r.renderBlockquote(n)

		case *ast.FencedCodeBlock:
			return r.renderCodeBlock(n)

		case *ast.CodeBlock:
			return r.renderCodeBlock(n)

		case *ast.ThematicBreak:
			return r.renderThematicBreak(n)

		case *ast.Text:
			// Text nodes are handled by the parent node
			return ast.WalkSkipChildren, nil

		case *ast.Emphasis:
			// Emphasis nodes are handled by the parent node
			return ast.WalkSkipChildren, nil

		case *ast.Link:
			// Link nodes are handled by the parent node
			return ast.WalkSkipChildren, nil

		case *ast.Image:
			return r.renderImage(n)

		// Table support
		case *extast.Table:
			if r.opts.EnableTables {
				return r.renderTable(n)
			}
			return ast.WalkContinue, nil

		case *extast.TableRow:
			// TableRow nodes are handled by Table
			return ast.WalkSkipChildren, nil

		case *extast.TableCell:
			// TableCell nodes are handled by Table
			return ast.WalkSkipChildren, nil

		// Task list support
		case *extast.TaskCheckBox:
			if r.opts.EnableTaskList {
				return r.renderTaskCheckBox(n)
			}
			return ast.WalkContinue, nil

		default:
			// Check if this is a math formula node
			if r.opts.EnableMath {
				// Check for block-level math formula
				if node.Kind() == mathjax.KindMathBlock {
					return r.renderMathBlock(node)
				}
				// Check for inline math formula
				if node.Kind() == mathjax.KindInlineMath {
					return r.renderInlineMath(node)
				}
			}
			// For unsupported node types, log the error but continue processing
			if r.opts.ErrorCallback != nil {
				r.opts.ErrorCallback(NewConversionError("UnsupportedNode", "unsupported markdown node type", 0, 0, nil))
			}
			return ast.WalkContinue, nil
		}
	})
}

// renderHeading renders a heading node.
func (r *WordRenderer) renderHeading(node *ast.Heading) (ast.WalkStatus, error) {
	text := r.extractTextContent(node)
	level := node.Level

	// Clamp heading level
	if level > 6 {
		level = 6
	}

	// Use the existing API for compatibility
	if r.opts.GenerateTOC && level <= r.opts.TOCMaxLevel {
		// Reuse the existing AddHeadingWithBookmark method
		r.doc.AddHeadingWithBookmark(text, level, "")
	} else {
		// Reuse the existing AddHeadingParagraph method
		r.doc.AddHeadingParagraph(text, level)
	}

	return ast.WalkSkipChildren, nil
}

// renderParagraph renders a paragraph node.
func (r *WordRenderer) renderParagraph(node *ast.Paragraph) (ast.WalkStatus, error) {
	// Check if the paragraph is empty
	if !node.HasChildren() {
		return ast.WalkSkipChildren, nil
	}

	// Create paragraph
	para := r.doc.AddParagraph("")

	// Process paragraph content
	r.renderInlineContent(node, para)

	return ast.WalkSkipChildren, nil
}

// renderInlineContent renders inline content (text, emphasis, links, etc.).
func (r *WordRenderer) renderInlineContent(node ast.Node, para *document.Paragraph) {
	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		switch n := child.(type) {
		case *ast.Text:
			text := string(n.Segment.Value(r.source))
			para.AddFormattedText(text, nil)
			
			// Handle soft line breaks (single \n)
			// goldmark parses a single \n into multiple Text nodes, with the first node's SoftLineBreak set to true
			// In Markdown, soft line breaks are typically rendered as spaces
			if n.SoftLineBreak() {
				para.AddFormattedText(" ", nil)
			}

		case *ast.Emphasis:
			text := r.extractTextContent(n)
			// In goldmark, level=1 is italic, level=2 is bold
			if n.Level == 2 {
				// Apply bold formatting
				format := &document.TextFormat{Bold: true}
				para.AddFormattedText(text, format)
			} else {
				// Apply italic formatting
				format := &document.TextFormat{Italic: true}
				para.AddFormattedText(text, format)
			}

		case *ast.CodeSpan:
			text := r.extractTextContent(n)
			// Apply CodeChar style formatting
			format := &document.TextFormat{
				FontFamily: "Consolas",
				FontColor:  "D73A49", // GitHub-style red
			}
			para.AddFormattedText(text, format)

		case *ast.Link:
			text := r.extractTextContent(n)
			// Simple link handling; can be extended to hyperlinks later
			format := &document.TextFormat{
				FontColor: "0000FF", // blue
			}
			para.AddFormattedText(text, format)

		case *ast.Image:
			r.renderImageInline(n, para)
		case *extast.Strikethrough:
			// Handle strikethrough
			text := r.extractTextContent(n)
			format := &document.TextFormat{
				Strike: true,
			}
			para.AddFormattedText(text, format)

		default:
			// Check if this is an inline math formula
			if r.opts.EnableMath && child.Kind() == mathjax.KindInlineMath {
				r.renderInlineMathToParagraph(child, para)
				continue
			}
			// For other types, try to extract text content
			text := r.extractTextContent(n)
			if text != "" {
				para.AddFormattedText(text, nil)
			}
		}
	}
}

// renderList renders a list node.
func (r *WordRenderer) renderList(node *ast.List) (ast.WalkStatus, error) {
	r.listLevel++
	defer func() { r.listLevel-- }()

	// Process list items
	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		if listItem, ok := child.(*ast.ListItem); ok {
			r.renderListItem(listItem)
		}
	}

	return ast.WalkSkipChildren, nil
}

// renderListItem renders a list item node.
func (r *WordRenderer) renderListItem(node *ast.ListItem) (ast.WalkStatus, error) {
	// Check if the item contains a task checkbox
	hasTaskCheckBox := false
	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		if _, ok := child.(*extast.TaskCheckBox); ok {
			hasTaskCheckBox = true
			break
		}
	}

	// If the item contains a task checkbox and task lists are enabled, let the TaskCheckBox node handle it
	if hasTaskCheckBox && r.opts.EnableTaskList {
		// Task list items will be handled by the TaskCheckBox node
		return ast.WalkContinue, nil
	}

	// Normal list item processing
	text := r.extractTextContent(node)

	// Simple list item handling; can be extended to proper list formatting later
	// For now, indentation and bullet symbols are used to simulate lists
	indent := strings.Repeat("  ", r.listLevel-1)
	bulletText := "• " + text

	r.doc.AddParagraph(indent + bulletText)

	return ast.WalkSkipChildren, nil
}

// renderBlockquote renders a blockquote node.
func (r *WordRenderer) renderBlockquote(node *ast.Blockquote) (ast.WalkStatus, error) {
	text := r.extractTextContent(node)

	// Create a blockquote paragraph using the Quote style
	para := r.doc.AddParagraph(text)
	para.SetStyle("Quote")

	return ast.WalkSkipChildren, nil
}

// renderCodeBlock renders a code block node.
func (r *WordRenderer) renderCodeBlock(node ast.Node) (ast.WalkStatus, error) {
	// Process code block content line by line, preserving formatting
	lines := r.extractCodeBlockLines(node)

	// Create a paragraph for each code line, preserving line breaks and indentation
	for _, line := range lines {
		// Handle empty lines
		if strings.TrimSpace(line) == "" {
			para := r.doc.AddParagraph(" ") // Represent empty lines with a space
			para.SetStyle("CodeBlock")
			r.applyCodeBlockFormatting(para)
			continue
		}

		// Create a code line paragraph
		para := r.doc.AddParagraph(line)
		para.SetStyle("CodeBlock")
		r.applyCodeBlockFormatting(para)
	}

	return ast.WalkSkipChildren, nil
}

// extractCodeBlockLines extracts code block text line by line, preserving formatting.
func (r *WordRenderer) extractCodeBlockLines(node ast.Node) []string {
	var lines []string

	for i := 0; i < node.Lines().Len(); i++ {
		line := node.Lines().At(i)
		lineText := string(line.Value(r.source))
		// Preserve original formatting, including spaces and tabs
		lines = append(lines, lineText)
	}

	return lines
}

// applyCodeBlockFormatting applies code block formatting to a paragraph.
func (r *WordRenderer) applyCodeBlockFormatting(para *document.Paragraph) {
	// Apply additional code block formatting
	if para.Properties == nil {
		para.Properties = &document.ParagraphProperties{}
	}

	// Set left indentation (matching the code_template style)
	para.Properties.Indentation = &document.Indentation{
		Left: "360", // 0.25 inch left indent, consistent with code_template
	}

	// Set spacing (before and after paragraph)
	para.Properties.Spacing = &document.Spacing{
		Before: "60", // 3pt before spacing (reduced to avoid excessive gaps between code lines)
		After:  "60", // 3pt after spacing
	}

	// Set alignment to left
	para.Properties.Justification = &document.Justification{
		Val: "left",
	}
}

// renderThematicBreak renders a thematic break (horizontal rule).
func (r *WordRenderer) renderThematicBreak(node *ast.ThematicBreak) (ast.WalkStatus, error) {
	// Create an empty paragraph to display the horizontal rule
	para := r.doc.AddParagraph("")

	// Set horizontal rule style
	// Single line style, medium weight, black
	para.SetHorizontalRule(document.BorderStyleSingle, 12, "000000")

	// Set paragraph spacing to make the horizontal rule more visually distinct
	para.SetSpacing(&document.SpacingConfig{
		BeforePara: 6, // 6pt before spacing
		AfterPara:  6, // 6pt after spacing
	})

	return ast.WalkSkipChildren, nil
}

// renderImage renders an image node.
func (r *WordRenderer) renderImage(node *ast.Image) (ast.WalkStatus, error) {
	// Get image path
	src := string(node.Destination)
	alt := r.extractTextContent(node)

	// Handle relative paths
	if !filepath.IsAbs(src) && r.opts.ImageBasePath != "" {
		src = filepath.Join(r.opts.ImageBasePath, src)
	}

	// Try to add the image; fall back to alt text on failure
	// Image handling logic can be improved later
	if alt != "" {
		r.doc.AddParagraph("[Image: " + alt + "]")
	} else {
		r.doc.AddParagraph("[Image: " + src + "]")
	}

	return ast.WalkSkipChildren, nil
}

// renderImageInline renders an inline image node.
func (r *WordRenderer) renderImageInline(node *ast.Image, para *document.Paragraph) {
	src := string(node.Destination)
	alt := r.extractTextContent(node)

	// Handle relative paths
	if !filepath.IsAbs(src) && r.opts.ImageBasePath != "" {
		src = filepath.Join(r.opts.ImageBasePath, src)
	}

	// Inline images are temporarily represented as text
	if alt != "" {
		para.AddFormattedText("[Image: "+alt+"]", nil)
	} else {
		para.AddFormattedText("[Image: "+src+"]", nil)
	}
}

// extractTextContent extracts the text content of a node.
func (r *WordRenderer) extractTextContent(node ast.Node) string {
	var buf strings.Builder
	r.extractTextContentRecursive(node, &buf)
	return buf.String()
}

// extractTextContentRecursive recursively extracts text content.
func (r *WordRenderer) extractTextContentRecursive(node ast.Node, buf *strings.Builder) {
	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		switch n := child.(type) {
		case *ast.Text:
			buf.Write(n.Segment.Value(r.source))
		default:
			r.extractTextContentRecursive(child, buf)
		}
	}
}

// cleanText cleans up text content by collapsing whitespace.
func (r *WordRenderer) cleanText(text string) string {
	// Remove excess whitespace characters
	re := regexp.MustCompile(`\s+`)
	text = re.ReplaceAllString(text, " ")
	return strings.TrimSpace(text)
}

// renderTable renders a table node.
func (r *WordRenderer) renderTable(node *extast.Table) (ast.WalkStatus, error) {
	// Collect table data
	var tableData [][]string
	var alignments []extast.Alignment
	var emphases [][]int

	// Iterate over table headers
	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		if row, ok := child.(*extast.TableHeader); ok {
			var rowData []string
			var rowEmphasis []int
			// Iterate over header cells
			for cellChild := row.FirstChild(); cellChild != nil; cellChild = cellChild.NextSibling() {
				if cell, ok := cellChild.(*extast.TableCell); ok {
					cellText := r.extractTextContent(cell)
					rowData = append(rowData, cellText)
					// Headers default to bold
					rowEmphasis = append(rowEmphasis, 2)
				}
			}
			tableData = append(tableData, rowData)
			emphases = append(emphases, rowEmphasis)
		}
	}

	// Iterate over table rows
	for child := node.FirstChild(); child != nil; child = child.NextSibling() {
		if row, ok := child.(*extast.TableRow); ok {
			var rowData []string
			var rowEmphasis []int
			if len(alignments) == 0 {
				// Get alignment from the first row
				alignments = row.Alignments
			}

			// Iterate over cells
			for cellChild := row.FirstChild(); cellChild != nil; cellChild = cellChild.NextSibling() {
				if cell, ok := cellChild.(*extast.TableCell); ok {
					cellText := r.extractTextContent(cell)
					rowData = append(rowData, cellText)
					emphasis := extractCellEmphasis(cell)
					rowEmphasis = append(rowEmphasis, emphasis)
				}
			}
			tableData = append(tableData, rowData)
			emphases = append(emphases, rowEmphasis)
		}
	}

	// Skip if there is no data
	if len(tableData) == 0 {
		return ast.WalkSkipChildren, nil
	}

	// Calculate column count
	cols := 0
	for _, row := range tableData {
		if len(row) > cols {
			cols = len(row)
		}
	}

	// Create table configuration
	config := &document.TableConfig{
		Rows:     len(tableData),
		Cols:     cols,
		Width:    9000, // default width (points)
		Data:     tableData,
		Emphases: emphases,
	}

	// Add table to document
	table, err := r.doc.AddTable(config)
	if err != nil && r.opts.ErrorCallback != nil {
		r.opts.ErrorCallback(NewConversionError("AddTable", err.Error(), 0, 0, err))
	}
	if table != nil {
		// Set header style (if applicable)
		if len(tableData) > 0 {
			// Set the first row as a header
			err := table.SetRowAsHeader(0, true)
			if err != nil && r.opts.ErrorCallback != nil {
				r.opts.ErrorCallback(NewConversionError("TableHeader", "failed to set table header", 0, 0, err))
			}
		}

		// Set cell alignment based on column alignments
		for rowIdx, row := range tableData {
			for colIdx := range row {
				if colIdx < len(alignments) {
					var align document.CellAlignment
					switch alignments[colIdx] {
					case extast.AlignLeft:
						align = document.CellAlignLeft
					case extast.AlignCenter:
						align = document.CellAlignCenter
					case extast.AlignRight:
						align = document.CellAlignRight
					default:
						align = document.CellAlignLeft
					}

					format := &document.CellFormat{
						HorizontalAlign: align,
					}
					err := table.SetCellFormat(rowIdx, colIdx, format)
					if err != nil && r.opts.ErrorCallback != nil {
						r.opts.ErrorCallback(NewConversionError("CellFormat", "failed to set cell format", rowIdx, colIdx, err))
					}
				}
			}
		}
	}

	return ast.WalkSkipChildren, nil
}

// extractCellEmphasis extracts the emphasis level from a table cell.
func extractCellEmphasis(cell *extast.TableCell) int {
	format := 0 // 0 means no formatting
	// Iterate over cell content
	ast.Walk(cell, func(n ast.Node, entering bool) (ast.WalkStatus, error) {
		if !entering {
			return ast.WalkContinue, nil
		}

		switch node := n.(type) {
		case *ast.Emphasis:
			// Handle emphasis text (bold or italic)
			format = node.Level
		}

		return ast.WalkContinue, nil
	})

	return format
}

// renderTaskCheckBox renders a task list checkbox.
func (r *WordRenderer) renderTaskCheckBox(node *extast.TaskCheckBox) (ast.WalkStatus, error) {
	// Get checkbox state
	checked := node.IsChecked

	// Select symbol based on state
	var checkSymbol string
	if checked {
		checkSymbol = "☑" // checked checkbox
	} else {
		checkSymbol = "☐" // unchecked checkbox
	}

	// Create a paragraph to contain the checkbox
	para := r.doc.AddParagraph("")

	// Add checkbox symbol
	para.AddFormattedText(checkSymbol+" ", nil)

	// Process task item text (usually other content in the parent ListItem)
	// Note: TaskCheckBox is typically the first child element of a ListItem
	parent := node.Parent()
	if parent != nil {
		// Extract text content other than the TaskCheckBox
		r.renderTaskItemContent(parent, para, node)
	}

	return ast.WalkSkipChildren, nil
}

// renderTaskItemContent renders task item content (text excluding the checkbox).
func (r *WordRenderer) renderTaskItemContent(parent ast.Node, para *document.Paragraph, skipNode ast.Node) {
	for child := parent.FirstChild(); child != nil; child = child.NextSibling() {
		// Skip the checkbox node itself
		if child == skipNode {
			continue
		}

		switch n := child.(type) {
		case *ast.Text:
			text := string(n.Segment.Value(r.source))
			para.AddFormattedText(text, nil)

			// Handle soft line breaks (single \n)
			if n.SoftLineBreak() {
				para.AddFormattedText(" ", nil)
			}
		case *ast.Emphasis:
			text := r.extractTextContent(n)
			if n.Level == 2 {
				format := &document.TextFormat{Bold: true}
				para.AddFormattedText(text, format)
			} else {
				format := &document.TextFormat{Italic: true}
				para.AddFormattedText(text, format)
			}
		case *ast.CodeSpan:
			text := r.extractTextContent(n)
			format := &document.TextFormat{
				FontFamily: "Consolas",
			}
			para.AddFormattedText(text, format)
		case *ast.Link:
			text := r.extractTextContent(n)
			format := &document.TextFormat{
				FontColor: "0000FF", // blue
			}
			para.AddFormattedText(text, format)
		default:
			// Check if this is an inline math formula
			if r.opts.EnableMath && child.Kind() == mathjax.KindInlineMath {
				r.renderInlineMathToParagraph(child, para)
				continue
			}
			// For other types, try to extract text content
			text := r.extractTextContent(n)
			if text != "" {
				para.AddFormattedText(text, nil)
			}
		}
	}
}

// renderInlineMathToParagraph renders an inline math formula into a paragraph.
func (r *WordRenderer) renderInlineMathToParagraph(node ast.Node, para *document.Paragraph) {
	latex := r.extractMathContent(node)
	para.AddFormattedText(latex, &document.TextFormat{
		FontFamily: "Cambria Math",
	})
}

// renderMathBlock renders a block-level math formula.
func (r *WordRenderer) renderMathBlock(node ast.Node) (ast.WalkStatus, error) {
	// Extract LaTeX content
	latex := r.extractMathContent(node)

	// Create a paragraph containing the formula
	para := r.doc.AddParagraph("")
	para.SetAlignment(document.AlignCenter)

	// Add formula content (using Cambria Math font for better math symbol rendering)
	para.AddFormattedText(latex, &document.TextFormat{
		FontFamily: "Cambria Math",
		FontSize:   12,
	})

	return ast.WalkSkipChildren, nil
}

// renderInlineMath renders an inline math formula.
func (r *WordRenderer) renderInlineMath(node ast.Node) (ast.WalkStatus, error) {
	// Inline formulas are usually handled by the parent node (in renderInlineContent)
	// If this method is called directly, create a new paragraph
	para := r.doc.AddParagraph("")
	r.renderInlineMathToParagraph(node, para)

	return ast.WalkSkipChildren, nil
}

// extractMathContent extracts LaTeX content from a math node.
func (r *WordRenderer) extractMathContent(node ast.Node) string {
	var content strings.Builder

	// Check if this is a block-level node (MathBlock)
	if node.Kind() == mathjax.KindMathBlock {
		// MathBlock is a block-level node; Lines() can be safely accessed
		if blockNode, ok := node.(ast.Node); ok {
			lines := blockNode.Lines()
			if lines != nil {
				for i := 0; i < lines.Len(); i++ {
					line := lines.At(i)
					content.Write(line.Value(r.source))
				}
			}
		}
	}

	// If there is no line content (possibly an inline formula), try extracting from child nodes
	if content.Len() == 0 {
		for child := node.FirstChild(); child != nil; child = child.NextSibling() {
			if text, ok := child.(*ast.Text); ok {
				content.Write(text.Segment.Value(r.source))
			}
		}
	}

	latex := strings.TrimSpace(content.String())

	// Convert common LaTeX commands to Unicode characters for better display
	latex = convertLaTeXToDisplay(latex)

	return latex
}

// convertLaTeXToDisplay converts LaTeX commands to displayable Unicode characters.
func convertLaTeXToDisplay(latex string) string {
	// Handle fractions: \frac{a}{b} -> a/b or (a)/(b)
	fracPattern := regexp.MustCompile(`\\frac\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}`)
	for fracPattern.MatchString(latex) {
		latex = fracPattern.ReplaceAllStringFunc(latex, func(match string) string {
			parts := fracPattern.FindStringSubmatch(match)
			if len(parts) == 3 {
				num := convertLaTeXToDisplay(parts[1])
				den := convertLaTeXToDisplay(parts[2])
				// Use fraction slash notation to represent fractions
				return "(" + num + ")/(" + den + ")"
			}
			return match
		})
	}

	// Handle square roots: \sqrt{x} -> √x, \sqrt[n]{x} -> ⁿ√x
	sqrtPattern := regexp.MustCompile(`\\sqrt\s*(?:\[([^\]]*)\])?\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}`)
	for sqrtPattern.MatchString(latex) {
		latex = sqrtPattern.ReplaceAllStringFunc(latex, func(match string) string {
			parts := sqrtPattern.FindStringSubmatch(match)
			if len(parts) == 3 {
				deg := parts[1]
				content := convertLaTeXToDisplay(parts[2])
				if deg == "" {
					return "√(" + content + ")"
				}
				// Convert the root index to superscript
				degSup := convertToSuperscript(deg)
				return degSup + "√(" + content + ")"
			}
			return match
		})
	}

	// Handle superscripts: x^{n} -> xⁿ or x^n -> xⁿ
	supBracePattern := regexp.MustCompile(`\^\{([^{}]*)\}`)
	latex = supBracePattern.ReplaceAllStringFunc(latex, func(match string) string {
		parts := supBracePattern.FindStringSubmatch(match)
		if len(parts) == 2 {
			return convertToSuperscript(parts[1])
		}
		return match
	})
	supSimplePattern := regexp.MustCompile(`\^([a-zA-Z0-9])`)
	latex = supSimplePattern.ReplaceAllStringFunc(latex, func(match string) string {
		parts := supSimplePattern.FindStringSubmatch(match)
		if len(parts) == 2 {
			return convertToSuperscript(parts[1])
		}
		return match
	})

	// Handle subscripts: x_{n} -> xₙ or x_n -> xₙ
	subBracePattern := regexp.MustCompile(`_\{([^{}]*)\}`)
	latex = subBracePattern.ReplaceAllStringFunc(latex, func(match string) string {
		parts := subBracePattern.FindStringSubmatch(match)
		if len(parts) == 2 {
			return convertToSubscript(parts[1])
		}
		return match
	})
	subSimplePattern := regexp.MustCompile(`_([a-zA-Z0-9])`)
	latex = subSimplePattern.ReplaceAllStringFunc(latex, func(match string) string {
		parts := subSimplePattern.FindStringSubmatch(match)
		if len(parts) == 2 {
			return convertToSubscript(parts[1])
		}
		return match
	})

	// LaTeX command replacement map
	replacements := map[string]string{
		// Greek letters (lowercase)
		`\alpha`:   "α",
		`\beta`:    "β",
		`\gamma`:   "γ",
		`\delta`:   "δ",
		`\epsilon`: "ε",
		`\zeta`:    "ζ",
		`\eta`:     "η",
		`\theta`:   "θ",
		`\iota`:    "ι",
		`\kappa`:   "κ",
		`\lambda`:  "λ",
		`\mu`:      "μ",
		`\nu`:      "ν",
		`\xi`:      "ξ",
		`\pi`:      "π",
		`\rho`:     "ρ",
		`\sigma`:   "σ",
		`\tau`:     "τ",
		`\upsilon`: "υ",
		`\phi`:     "φ",
		`\chi`:     "χ",
		`\psi`:     "ψ",
		`\omega`:   "ω",

		// Greek letters (uppercase)
		`\Alpha`:   "Α",
		`\Beta`:    "Β",
		`\Gamma`:   "Γ",
		`\Delta`:   "Δ",
		`\Epsilon`: "Ε",
		`\Zeta`:    "Ζ",
		`\Eta`:     "Η",
		`\Theta`:   "Θ",
		`\Iota`:    "Ι",
		`\Kappa`:   "Κ",
		`\Lambda`:  "Λ",
		`\Mu`:      "Μ",
		`\Nu`:      "Ν",
		`\Xi`:      "Ξ",
		`\Pi`:      "Π",
		`\Rho`:     "Ρ",
		`\Sigma`:   "Σ",
		`\Tau`:     "Τ",
		`\Upsilon`: "Υ",
		`\Phi`:     "Φ",
		`\Chi`:     "Χ",
		`\Psi`:     "Ψ",
		`\Omega`:   "Ω",

		// Operators
		`\times`:  "×",
		`\div`:    "÷",
		`\pm`:     "±",
		`\mp`:     "∓",
		`\cdot`:   "·",
		`\ast`:    "∗",
		`\star`:   "⋆",
		`\circ`:   "∘",
		`\bullet`: "•",
		`\oplus`:  "⊕",
		`\ominus`: "⊖",
		`\otimes`: "⊗",

		// Relational symbols
		`\leq`:      "≤",
		`\le`:       "≤",
		`\geq`:      "≥",
		`\ge`:       "≥",
		`\neq`:      "≠",
		`\ne`:       "≠",
		`\approx`:   "≈",
		`\equiv`:    "≡",
		`\sim`:      "∼",
		`\simeq`:    "≃",
		`\cong`:     "≅",
		`\propto`:   "∝",
		`\ll`:       "≪",
		`\gg`:       "≫",
		`\subset`:   "⊂",
		`\supset`:   "⊃",
		`\subseteq`: "⊆",
		`\supseteq`: "⊇",
		`\in`:       "∈",
		`\notin`:    "∉",
		`\ni`:       "∋",

		// Arrows
		`\rightarrow`:     "→",
		`\leftarrow`:      "←",
		`\leftrightarrow`: "↔",
		`\Rightarrow`:     "⇒",
		`\Leftarrow`:      "⇐",
		`\Leftrightarrow`: "⇔",
		`\uparrow`:        "↑",
		`\downarrow`:      "↓",
		`\to`:             "→",
		`\gets`:           "←",
		`\mapsto`:         "↦",

		// Miscellaneous symbols
		`\infty`:      "∞",
		`\partial`:    "∂",
		`\nabla`:      "∇",
		`\forall`:     "∀",
		`\exists`:     "∃",
		`\nexists`:    "∄",
		`\emptyset`:   "∅",
		`\varnothing`: "∅",
		`\neg`:        "¬",
		`\lnot`:       "¬",
		`\land`:       "∧",
		`\lor`:        "∨",
		`\cap`:        "∩",
		`\cup`:        "∪",
		`\int`:        "∫",
		`\iint`:       "∬",
		`\iiint`:      "∭",
		`\oint`:       "∮",
		`\sum`:        "∑",
		`\prod`:       "∏",
		`\coprod`:     "∐",

		// Ellipsis
		`\ldots`: "…",
		`\cdots`: "⋯",
		`\vdots`: "⋮",
		`\ddots`: "⋱",

		// Spaces
		`\quad`:  " ",
		`\qquad`: "  ",
		`\,`:     " ",
		`\;`:     " ",
		`\:`:     " ",
		`\ `:     " ",

		// Brackets
		`\{`:      "{",
		`\}`:      "}",
		`\lbrace`: "{",
		`\rbrace`: "}",
		`\langle`: "⟨",
		`\rangle`: "⟩",
		`\lceil`:  "⌈",
		`\rceil`:  "⌉",
		`\lfloor`: "⌊",
		`\rfloor`: "⌋",
		`\left`:   "",
		`\right`:  "",
	}

	// Apply replacements - sorted by length in descending order to ensure longer commands are replaced first
	// This avoids conflicts between commands like \neq and \ne
	type replacement struct {
		cmd     string
		unicode string
	}
	sortedReplacements := make([]replacement, 0, len(replacements))
	for cmd, u := range replacements {
		sortedReplacements = append(sortedReplacements, replacement{cmd, u})
	}
	// Sort by command length in descending order
	for i := 0; i < len(sortedReplacements); i++ {
		for j := i + 1; j < len(sortedReplacements); j++ {
			if len(sortedReplacements[j].cmd) > len(sortedReplacements[i].cmd) {
				sortedReplacements[i], sortedReplacements[j] = sortedReplacements[j], sortedReplacements[i]
			}
		}
	}
	for _, r := range sortedReplacements {
		latex = strings.ReplaceAll(latex, r.cmd, r.unicode)
	}

	// Clean up excess curly braces
	latex = strings.ReplaceAll(latex, "{", "")
	latex = strings.ReplaceAll(latex, "}", "")

	return latex
}

// convertToSuperscript converts a string to superscript form.
func convertToSuperscript(s string) string {
	superscripts := map[rune]rune{
		'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
		'5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
		'+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾',
		'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'd': 'ᵈ', 'e': 'ᵉ',
		'f': 'ᶠ', 'g': 'ᵍ', 'h': 'ʰ', 'i': 'ⁱ', 'j': 'ʲ',
		'k': 'ᵏ', 'l': 'ˡ', 'm': 'ᵐ', 'n': 'ⁿ', 'o': 'ᵒ',
		'p': 'ᵖ', 'r': 'ʳ', 's': 'ˢ', 't': 'ᵗ', 'u': 'ᵘ',
		'v': 'ᵛ', 'w': 'ʷ', 'x': 'ˣ', 'y': 'ʸ', 'z': 'ᶻ',
	}

	var result strings.Builder
	for _, r := range s {
		if sup, ok := superscripts[r]; ok {
			result.WriteRune(sup)
		} else {
			result.WriteRune(r)
		}
	}
	return result.String()
}

// convertToSubscript converts a string to subscript form.
func convertToSubscript(s string) string {
	subscripts := map[rune]rune{
		'0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
		'5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
		'+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
		'a': 'ₐ', 'e': 'ₑ', 'h': 'ₕ', 'i': 'ᵢ', 'j': 'ⱼ',
		'k': 'ₖ', 'l': 'ₗ', 'm': 'ₘ', 'n': 'ₙ', 'o': 'ₒ',
		'p': 'ₚ', 'r': 'ᵣ', 's': 'ₛ', 't': 'ₜ', 'u': 'ᵤ',
		'v': 'ᵥ', 'x': 'ₓ',
	}

	var result strings.Builder
	for _, r := range s {
		if sub, ok := subscripts[r]; ok {
			result.WriteRune(sub)
		} else {
			result.WriteRune(r)
		}
	}
	return result.String()
}
