/*
Package document provides a Go library for creating, editing, and manipulating Microsoft Word documents.

WordZero focuses on the modern Office Open XML (OOXML) format (.docx files),
providing a simple and easy-to-use API for creating and modifying Word documents.

# Main Features

## Basic Features
- Create new Word documents
- Open and parse existing .docx files
- Add and format text content
- Set paragraph styles and alignment
- Configure fonts, colors, and text formatting
- Set line spacing and paragraph spacing
- Error handling and logging

## Advanced Features
- Headers and footers: support for different headers/footers on default, first, and even pages
- Table of contents generation: automatically generate TOC based on heading styles, with hyperlinks and page numbers
- Footnotes and endnotes: full footnote and endnote functionality with multiple numbering formats
- List numbering: unordered and ordered lists with multi-level nesting support
- Page setup: complete page property settings including page size, orientation, and margins
- Tables: powerful table creation, formatting, and styling features
- Style system: 18 predefined styles plus custom style support

# Quick Start

Create a simple document:

	doc := document.New()
	doc.AddParagraph("Hello, World!")
	err := doc.Save("hello.docx")

Create a formatted document:

	doc := document.New()

	// Add a formatted title
	titleFormat := &document.TextFormat{
		Bold:      true,
		FontSize:  18,
		FontColor: "FF0000", // Red
		FontName:  "Arial",
	}
	title := doc.AddFormattedParagraph("Document Title", titleFormat)
	title.SetAlignment(document.AlignCenter)

	// Add a body paragraph
	para := doc.AddParagraph("This is the body content...")
	para.SetSpacing(&document.SpacingConfig{
		LineSpacing:     1.5, // 1.5x line spacing
		BeforePara:      12,  // 12pt before paragraph
		AfterPara:       6,   // 6pt after paragraph
		FirstLineIndent: 24,  // 24pt first line indent
	})

	err := doc.Save("formatted.docx")

Open an existing document:

	doc, err := document.Open("existing.docx")
	if err != nil {
		log.Fatal(err)
	}

	// Read paragraph content
	for i, para := range doc.Body.Paragraphs {
		fmt.Printf("Paragraph %d: ", i+1)
		for _, run := range para.Runs {
			fmt.Print(run.Text.Content)
		}
		fmt.Println()
	}

# Advanced Feature Examples

## Headers and Footers

	// Add a header
	doc.AddHeader(document.HeaderFooterTypeDefault, "This is the header")

	// Add a footer with page number
	doc.AddFooterWithPageNumber(document.HeaderFooterTypeDefault, "Page", true)

	// Set different first page
	doc.SetDifferentFirstPage(true)

## Table of Contents

	// Add headings with bookmarks
	doc.AddHeadingWithBookmark("Chapter 1: Overview", 1, "chapter1")
	doc.AddHeadingWithBookmark("1.1 Background", 2, "section1_1")

	// Generate table of contents
	tocConfig := document.DefaultTOCConfig()
	tocConfig.Title = "Table of Contents"
	tocConfig.MaxLevel = 3
	doc.GenerateTOC(tocConfig)

## Footnotes and Endnotes

	// Add a footnote
	doc.AddFootnote("This is the body text", "This is the footnote content")

	// Add an endnote
	doc.AddEndnote("Additional notes", "This is the endnote content")

	// Custom footnote configuration
	footnoteConfig := &document.FootnoteConfig{
		NumberFormat: document.FootnoteFormatLowerRoman,
		StartNumber:  1,
		RestartEach:  document.FootnoteRestartEachPage,
		Position:     document.FootnotePositionPageBottom,
	}
	doc.SetFootnoteConfig(footnoteConfig)

## Lists

	// Unordered list
	doc.AddBulletList("List item 1", 0, document.BulletTypeDot)
	doc.AddBulletList("Sub-item", 1, document.BulletTypeCircle)

	// Ordered list
	doc.AddNumberedList("First item", 0, document.ListTypeDecimal)
	doc.AddNumberedList("Second item", 0, document.ListTypeDecimal)

	// Multi-level list
	items := []document.ListItem{
		{Text: "Level 1 item", Level: 0, Type: document.ListTypeDecimal},
		{Text: "Level 2 item", Level: 1, Type: document.ListTypeLowerLetter},
		{Text: "Level 3 item", Level: 2, Type: document.ListTypeLowerRoman},
	}
	doc.CreateMultiLevelList(items)

## Page Setup

	// Set page to A4 landscape
	doc.SetPageOrientation(document.OrientationLandscape)

	// Set page margins (in millimeters)
	doc.SetPageMargins(25, 25, 25, 25)

	// Full page settings
	pageSettings := &document.PageSettings{
		Size:           document.PageSizeLetter,
		Orientation:    document.OrientationPortrait,
		MarginTop:      30,
		MarginRight:    20,
		MarginBottom:   30,
		MarginLeft:     20,
		HeaderDistance: 15,
		FooterDistance: 15,
		GutterWidth:    0,
	}
	doc.SetPageSettings(pageSettings)

## Tables

	// Create a table
	table := doc.CreateTable(&document.TableConfig{
		Rows:  3,
		Cols:  3,
		Width: 5000,
	})

	// Set cell text
	table.SetCellText(0, 0, "Title")

	// Apply table style
	table.ApplyTableStyle(&document.TableStyleConfig{
		HeaderRow:    true,
		FirstColumn:  true,
		BandedRows:   true,
		BandedCols:   false,
	})

# Error Handling

The library provides a unified error handling mechanism:

	doc, err := document.Open("nonexistent.docx")
	if err != nil {
		var docErr *document.DocumentError
		if errors.As(err, &docErr) {
			Errorf("Document operation failed - operation: %s, error: %v", docErr.Operation, docErr.Cause)
			fmt.Printf("Operation: %s, Error: %v\n", docErr.Operation, docErr.Cause)
		}
	}

# Logging

You can configure the log level to control output:

	// Set to debug mode
	document.SetGlobalLevel(document.LogLevelDebug)

	// Show only errors
	document.SetGlobalLevel(document.LogLevelError)

# Text Formatting

The TextFormat struct supports various text formatting options:

	format := &document.TextFormat{
		Bold:      true,           // Bold
		Italic:    true,           // Italic
		FontSize:  14,             // Font size (in points)
		FontColor: "0000FF",       // Font color (hexadecimal)
		FontName:  "Times New Roman", // Font name
	}

# Paragraph Alignment

Four alignment modes are supported:

	para.SetAlignment(document.AlignLeft)     // Left aligned
	para.SetAlignment(document.AlignCenter)   // Center aligned
	para.SetAlignment(document.AlignRight)    // Right aligned
	para.SetAlignment(document.AlignJustify)  // Justified

# Spacing Configuration

Paragraph spacing can be precisely controlled:

	config := &document.SpacingConfig{
		LineSpacing:     1.5, // Line spacing (multiplier)
		BeforePara:      12,  // Before paragraph spacing (in points)
		AfterPara:       6,   // After paragraph spacing (in points)
		FirstLineIndent: 24,  // First line indent (in points)
	}
	para.SetSpacing(config)

# Notes

- Font sizes are in points; they are automatically converted to Word's half-point units internally
- Color values use hexadecimal format without the # prefix
- Spacing values are in points; they are converted to TWIPs internally (1 point = 20 TWIPs)
- All text content uses UTF-8 encoding
- Header/footer types include: Default, First, and Even
- Footnotes and endnotes are automatically numbered, with support for multiple numbering formats and restart rules
- Lists support multi-level nesting, up to 9 levels of indentation
- The table of contents feature requires adding headings with bookmarks first, then calling the TOC generation method

For more details and examples, refer to the documentation of each type and function.
*/
package document
