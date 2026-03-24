// Package document provides table of contents generation for Word documents.
package document

import (
	"encoding/xml"
	"fmt"
	"strings"
)

const (
	// styleHeading1 is the style name for heading level 1.
	styleHeading1 = "Heading1"
)

// TOCConfig holds table of contents configuration.
type TOCConfig struct {
	Title        string // TOC title, defaults to "Table of Contents"
	MaxLevel     int    // Maximum heading level, defaults to 3 (displays levels 1-3)
	ShowPageNum  bool   // Whether to show page numbers, defaults to true
	RightAlign   bool   // Whether to right-align page numbers, defaults to true
	UseHyperlink bool   // Whether to use hyperlinks, defaults to true
	DotLeader    bool   // Whether to use dot leaders, defaults to true
}

// TOCEntry represents a single table of contents entry.
type TOCEntry struct {
	Text       string // Entry text
	Level      int    // Heading level (1-9)
	PageNum    int    // Page number
	BookmarkID string // Bookmark ID (used for hyperlinks)
}

// TOCField represents a table of contents field.
type TOCField struct {
	XMLName xml.Name `xml:"w:fldSimple"`
	Instr   string   `xml:"w:instr,attr"`
	Runs    []Run    `xml:"w:r"`
}

// Hyperlink represents a hyperlink structure.
type Hyperlink struct {
	XMLName xml.Name `xml:"w:hyperlink"`
	Anchor  string   `xml:"w:anchor,attr,omitempty"`
	Runs    []Run    `xml:"w:r"`
}

// BookmarkEnd represents a bookmark end element.
type BookmarkEnd struct {
	XMLName xml.Name `xml:"w:bookmarkEnd"`
	ID      string   `xml:"w:id,attr"`
}

// ElementType returns the element type for a bookmark end.
func (b *BookmarkEnd) ElementType() string {
	return "bookmarkEnd"
}

// BookmarkStart represents a bookmark start element.
type BookmarkStart struct {
	XMLName xml.Name `xml:"w:bookmarkStart"`
	ID      string   `xml:"w:id,attr"`
	Name    string   `xml:"w:name,attr"`
}

// ElementType returns the element type for a bookmark start.
func (b *BookmarkStart) ElementType() string {
	return "bookmarkStart"
}

// DefaultTOCConfig returns the default table of contents configuration.
func DefaultTOCConfig() *TOCConfig {
	return &TOCConfig{
		Title:        "Table of Contents",
		MaxLevel:     3,
		ShowPageNum:  true,
		RightAlign:   true,
		UseHyperlink: true,
		DotLeader:    true,
	}
}

// GenerateTOC generates a table of contents for the document.
func (d *Document) GenerateTOC(config *TOCConfig) error {
	if config == nil {
		config = DefaultTOCConfig()
	}

	// Collect heading information
	entries := d.collectHeadings(config.MaxLevel)

	// Create the TOC SDT
	tocSDT := d.CreateTOCSDT(config.Title, config.MaxLevel)

	// Add each heading entry to the TOC
	for i, entry := range entries {
		entryID := fmt.Sprintf("14746%d", 3000+i)
		tocSDT.AddTOCEntry(entry.Text, entry.Level, entry.PageNum, entryID)
	}

	// Finalize the TOC SDT structure
	tocSDT.FinalizeTOCSDT()

	// Append to the document
	d.Body.Elements = append(d.Body.Elements, tocSDT)

	return nil
}

// UpdateTOC updates the existing table of contents.
func (d *Document) UpdateTOC() error {
	// Find the existing TOC SDT
	tocSDT, tocIndex := d.findTOCSDT()
	if tocSDT == nil {
		// If no SDT-type TOC found, try finding a paragraph-type TOC
		tocStart := d.findTOCStart()
		if tocStart == -1 {
			return fmt.Errorf("table of contents not found")
		}

		// Remove existing TOC entries
		d.removeTOCEntries(tocStart)

		// Regenerate TOC entries
		config := DefaultTOCConfig()
		entries := d.collectHeadings(config.MaxLevel)
		for _, entry := range entries {
			if err := d.addTOCEntry(entry, config); err != nil {
				return fmt.Errorf("failed to update TOC entry: %w", err)
			}
		}
		return nil
	}

	// Handle SDT-type TOC
	// Use default TOC configuration
	config := DefaultTOCConfig()

	// Re-collect heading information
	entries := d.collectHeadings(config.MaxLevel)

	// Clear and rebuild SDT content
	tocSDT.Content.Elements = []interface{}{}

	// Add TOC title paragraph
	titlePara := &Paragraph{
		Properties: &ParagraphProperties{
			Spacing: &Spacing{
				Before: "0",
				After:  "0",
				Line:   "240",
			},
			Indentation: &Indentation{
				Left:      "0",
				Right:     "0",
				FirstLine: "0",
			},
			Justification: &Justification{Val: "center"},
		},
		Runs: []Run{
			{
				Text: Text{Content: config.Title},
				Properties: &RunProperties{
					FontFamily: &FontFamily{ASCII: "宋体"},
					FontSize:   &FontSize{Val: "21"},
				},
			},
		},
	}

	// Add bookmark start
	bookmarkStart := &BookmarkStart{
		ID:   "0",
		Name: "_Toc11693_WPSOffice_Type3",
	}

	tocSDT.Content.Elements = append(tocSDT.Content.Elements, bookmarkStart, titlePara)

	// Add each heading entry to the TOC
	for i, entry := range entries {
		entryID := fmt.Sprintf("14746%d", 3000+i)
		tocSDT.AddTOCEntry(entry.Text, entry.Level, entry.PageNum, entryID)
	}

	// Finalize the TOC SDT structure
	tocSDT.FinalizeTOCSDT()

	// Update the SDT in the document
	d.Body.Elements[tocIndex] = tocSDT

	return nil
}

// AddHeadingWithBookmark adds a heading with a bookmark to the document.
func (d *Document) AddHeadingWithBookmark(text string, level int, bookmarkName string) *Paragraph {
	if bookmarkName == "" {
		bookmarkName = fmt.Sprintf("_Toc_%s", strings.ReplaceAll(text, " ", "_"))
	}

	// Add bookmark start
	bookmarkID := fmt.Sprintf("%d", len(d.Body.Elements))
	bookmark := &BookmarkStart{
		ID:   bookmarkID,
		Name: bookmarkName,
	}

	// Create heading paragraph
	paragraph := d.AddHeadingParagraph(text, level)

	// Insert bookmark into the paragraph's Runs
	if len(paragraph.Runs) > 0 {
		// Insert bookmark start before the first Run
		bookmarkRun := Run{
			Properties: &RunProperties{},
		}
		// This requires special XML serialization handling to insert bookmark elements
		paragraph.Runs = append([]Run{bookmarkRun}, paragraph.Runs...)
	}

	// Add bookmark end
	bookmarkEnd := &BookmarkEnd{
		ID: bookmarkID,
	}

	// Add bookmark to the document (simplified handling)
	_ = bookmark // mark as used
	d.Body.Elements = append(d.Body.Elements, bookmarkEnd)

	return paragraph
}

// collectHeadings collects heading information from the document.
func (d *Document) collectHeadings(maxLevel int) []TOCEntry {
	var entries []TOCEntry
	pageNum := 1 // Simplified; actual implementation would calculate real page numbers

	for _, element := range d.Body.Elements {
		if paragraph, ok := element.(*Paragraph); ok {
			level := d.getHeadingLevel(paragraph)
			if level > 0 && level <= maxLevel {
				text := d.extractParagraphText(paragraph)
				if text != "" {
					entry := TOCEntry{
						Text:       text,
						Level:      level,
						PageNum:    pageNum,
						BookmarkID: fmt.Sprintf("_Toc_%s", strings.ReplaceAll(text, " ", "_")),
					}
					entries = append(entries, entry)
				}
			}
		}
	}

	return entries
}

// getHeadingLevel returns the heading level of a paragraph.
func (d *Document) getHeadingLevel(paragraph *Paragraph) int {
	if paragraph.Properties != nil && paragraph.Properties.ParagraphStyle != nil {
		styleVal := paragraph.Properties.ParagraphStyle.Val

		// Map heading level by style ID - supports numeric IDs
		switch styleVal {
		case "1": // heading 1 (some documents use 1 for heading 1)
			return 1
		case "2": // heading 1 (Word defaults to 2 for heading 1)
			return 1
		case "3": // heading 2
			return 2
		case "4": // heading 3
			return 3
		case "5": // heading 4
			return 4
		case "6": // heading 5
			return 5
		case "7": // heading 6
			return 6
		case "8": // heading 7
			return 7
		case "9": // heading 8
			return 8
		case "10": // heading 9
			return 9
		}

		// Support standard style name matching
		switch styleVal {
		case styleHeading1, "heading1", "Title1":
			return 1
		case "Heading2", "heading2", "Title2":
			return 2
		case "Heading3", "heading3", "Title3":
			return 3
		case "Heading4", "heading4", "Title4":
			return 4
		case "Heading5", "heading5", "Title5":
			return 5
		case "Heading6", "heading6", "Title6":
			return 6
		case "Heading7", "heading7", "Title7":
			return 7
		case "Heading8", "heading8", "Title8":
			return 8
		case "Heading9", "heading9", "Title9":
			return 9
		}

		// Support generic pattern matching (handle "Heading" followed by a number)
		if strings.HasPrefix(strings.ToLower(styleVal), "heading") {
			// Extract the numeric portion
			numStr := strings.TrimPrefix(strings.ToLower(styleVal), "heading")
			if numStr != "" {
				if level := parseInt(numStr); level >= 1 && level <= 9 {
					return level
				}
			}
		}
	}
	return 0
}

// parseInt is a simple string-to-integer conversion function.
func parseInt(s string) int {
	switch s {
	case "1":
		return 1
	case "2":
		return 2
	case "3":
		return 3
	case "4":
		return 4
	case "5":
		return 5
	case "6":
		return 6
	case "7":
		return 7
	case "8":
		return 8
	case "9":
		return 9
	default:
		return 0
	}
}

// extractParagraphText extracts the text content from a paragraph.
func (d *Document) extractParagraphText(paragraph *Paragraph) string {
	var text strings.Builder
	for _, run := range paragraph.Runs {
		text.WriteString(run.Text.Content)
	}
	return text.String()
}

// addTOCEntry adds a table of contents entry to the document.
func (d *Document) addTOCEntry(entry TOCEntry, config *TOCConfig) error {
	// Create the TOC entry paragraph
	entryPara := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: fmt.Sprintf("TOC%d", entry.Level)},
		},
	}

	if config.UseHyperlink {
		// Create hyperlink
		hyperlink := &Hyperlink{
			Anchor: entry.BookmarkID,
		}

		// Title text
		titleRun := Run{
			Properties: &RunProperties{},
			Text:       Text{Content: entry.Text},
		}
		hyperlink.Runs = append(hyperlink.Runs, titleRun)

		// If showing page numbers, add leader and page number
		if config.ShowPageNum {
			if config.DotLeader {
				// Add dot leader
				leaderRun := Run{
					Properties: &RunProperties{},
					Text:       Text{Content: strings.Repeat(".", 20)}, // Simplified handling
				}
				hyperlink.Runs = append(hyperlink.Runs, leaderRun)
			}

			// Add page number
			pageRun := Run{
				Properties: &RunProperties{},
				Text:       Text{Content: fmt.Sprintf("%d", entry.PageNum)},
			}
			hyperlink.Runs = append(hyperlink.Runs, pageRun)
		}

		// Add hyperlink to paragraph
		// This needs special handling since Hyperlink is not a standard Run
		// Simplified: add directly as text
		hyperlinkRun := Run{
			Properties: &RunProperties{},
			Text:       Text{Content: entry.Text},
		}
		entryPara.Runs = append(entryPara.Runs, hyperlinkRun)

		if config.ShowPageNum {
			pageRun := Run{
				Properties: &RunProperties{},
				Text:       Text{Content: fmt.Sprintf("\t%d", entry.PageNum)},
			}
			entryPara.Runs = append(entryPara.Runs, pageRun)
		}
	} else {
		// Plain text without hyperlinks
		titleRun := Run{
			Properties: &RunProperties{},
			Text:       Text{Content: entry.Text},
		}
		entryPara.Runs = append(entryPara.Runs, titleRun)

		if config.ShowPageNum {
			pageRun := Run{
				Properties: &RunProperties{},
				Text:       Text{Content: fmt.Sprintf("\t%d", entry.PageNum)},
			}
			entryPara.Runs = append(entryPara.Runs, pageRun)
		}
	}

	d.Body.Elements = append(d.Body.Elements, entryPara)
	return nil
}

// findTOCStart finds the starting position of the table of contents.
func (d *Document) findTOCStart() int {
	for i, element := range d.Body.Elements {
		if paragraph, ok := element.(*Paragraph); ok {
			if paragraph.Properties != nil && paragraph.Properties.ParagraphStyle != nil {
				if strings.HasPrefix(paragraph.Properties.ParagraphStyle.Val, "TOC") {
					return i
				}
			}
		}
	}
	return -1
}

// findTOCSDT finds the TOC SDT structure in the document.
func (d *Document) findTOCSDT() (*SDT, int) {
	for i, element := range d.Body.Elements {
		sdt, ok := element.(*SDT)
		if !ok {
			continue
		}

		// Check if this is a TOC SDT
		if sdt.Properties == nil || sdt.Properties.DocPartObj == nil {
			continue
		}

		if sdt.Properties.DocPartObj.DocPartGallery == nil {
			continue
		}

		if sdt.Properties.DocPartObj.DocPartGallery.Val == "Table of Contents" {
			return sdt, i
		}
	}
	return nil, -1
}

// removeTOCEntries removes existing table of contents entries.
func (d *Document) removeTOCEntries(startIndex int) {
	// Simplified: find and remove all TOC-styled paragraphs starting from startIndex
	var newElements []interface{}

	// Retain elements before startIndex
	newElements = append(newElements, d.Body.Elements[:startIndex]...)

	// Skip TOC-related elements
	for i := startIndex; i < len(d.Body.Elements); i++ {
		element := d.Body.Elements[i]
		if paragraph, ok := element.(*Paragraph); ok {
			if paragraph.Properties != nil && paragraph.Properties.ParagraphStyle != nil {
				if !strings.HasPrefix(paragraph.Properties.ParagraphStyle.Val, "TOC") {
					// Not a TOC style, retain all subsequent elements
					newElements = append(newElements, d.Body.Elements[i:]...)
					break
				}
			}
		}
	}

	d.Body.Elements = newElements
}

// SetTOCStyle sets the style for a specific TOC level.
func (d *Document) SetTOCStyle(level int, style *TextFormat) error {
	if level < 1 || level > 9 {
		return fmt.Errorf("TOC level must be between 1 and 9")
	}

	styleName := fmt.Sprintf("TOC%d", level)

	// Set TOC style via the style manager
	styleManager := d.GetStyleManager()

	// Create paragraph style (needs integration with the style system)
	// Simplified; actual implementation requires a complete style definition
	_ = styleManager
	_ = styleName
	_ = style

	return nil
}

// AutoGenerateTOC automatically generates a TOC by detecting headings in the document.
func (d *Document) AutoGenerateTOC(config *TOCConfig) error {
	if config == nil {
		config = DefaultTOCConfig()
	}

	// Find existing TOC position
	tocStart := d.findTOCStart()
	var insertIndex int

	if tocStart != -1 {
		// If a TOC already exists, remove existing entries
		d.removeTOCEntries(tocStart)
		insertIndex = tocStart
	} else {
		// If no TOC exists, insert at the beginning of the document
		insertIndex = 0
	}

	// Collect all headings and add bookmarks
	entries := d.collectHeadingsAndAddBookmarks(config.MaxLevel)

	if len(entries) == 0 {
		return fmt.Errorf("no headings found in the document (paragraphs with style IDs 2-10)")
	}

	// Generate the TOC using real Word field codes instead of a simplified SDT
	tocElements := d.createWordFieldTOC(config, entries)

	// Insert the TOC at the specified position
	if insertIndex == 0 {
		// Insert at the beginning
		d.Body.Elements = append(tocElements, d.Body.Elements...)
	} else {
		// Replace at the specified position
		newElements := make([]interface{}, 0, len(d.Body.Elements)+len(tocElements))
		newElements = append(newElements, d.Body.Elements[:insertIndex]...)
		newElements = append(newElements, tocElements...)
		newElements = append(newElements, d.Body.Elements[insertIndex:]...)
		d.Body.Elements = newElements
	}

	return nil
}

// GetHeadingCount returns the number of headings in the document by level, useful for debugging.
func (d *Document) GetHeadingCount() map[int]int {
	counts := make(map[int]int)

	for _, element := range d.Body.Elements {
		if paragraph, ok := element.(*Paragraph); ok {
			level := d.getHeadingLevel(paragraph)
			if level > 0 {
				counts[level]++
			}
		}
	}

	return counts
}

// ListHeadings lists all headings in the document, useful for debugging.
func (d *Document) ListHeadings() []TOCEntry {
	return d.collectHeadings(9) // Retrieve headings at all levels
}

// createWordFieldTOC creates a TOC using real Word field codes.
func (d *Document) createWordFieldTOC(config *TOCConfig, entries []TOCEntry) []interface{} {
	var elements []interface{}

	// Create the TOC SDT container
	tocSDT := &SDT{
		Properties: &SDTProperties{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: "宋体", HAnsi: "宋体", EastAsia: "宋体", CS: "Times New Roman"},
				FontSize:   &FontSize{Val: "21"},
			},
			ID:    &SDTID{Val: "147458718"},
			Color: &SDTColor{Val: "DBDBDB"},
			DocPartObj: &DocPartObj{
				DocPartGallery: &DocPartGallery{Val: "Table of Contents"},
				DocPartUnique:  &DocPartUnique{},
			},
		},
		EndPr: &SDTEndPr{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: "Calibri", HAnsi: "Calibri", EastAsia: "宋体", CS: "Times New Roman"},
				Bold:       &Bold{},
				Color:      &Color{Val: "2F5496"},
				FontSize:   &FontSize{Val: "32"},
			},
		},
		Content: &SDTContent{
			Elements: []interface{}{},
		},
	}

	// Add TOC title paragraph
	titlePara := &Paragraph{
		Properties: &ParagraphProperties{
			Spacing: &Spacing{
				Before: "0",
				After:  "0",
				Line:   "240",
			},
			Justification: &Justification{Val: "center"},
			Indentation: &Indentation{
				Left:      "0",
				Right:     "0",
				FirstLine: "0",
			},
		},
		Runs: []Run{
			{
				Text: Text{Content: config.Title},
				Properties: &RunProperties{
					FontFamily: &FontFamily{ASCII: "宋体"},
					FontSize:   &FontSize{Val: "21"},
				},
			},
		},
	}

	tocSDT.Content.Elements = append(tocSDT.Content.Elements, titlePara)

	// Create the main TOC field paragraph
	tocFieldPara := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: "12"}, // TOC style
			Tabs: &Tabs{
				Tabs: []TabDef{
					{
						Val:    "right",
						Leader: "dot",
						Pos:    "8640",
					},
				},
			},
		},
		Runs: []Run{},
	}

	// Add TOC field begin
	tocFieldPara.Runs = append(tocFieldPara.Runs, Run{
		Properties: &RunProperties{
			Bold:     &Bold{},
			Color:    &Color{Val: "2F5496"},
			FontSize: &FontSize{Val: "32"},
		},
		FieldChar: &FieldChar{
			FieldCharType: "begin",
		},
	})

	// Add TOC instruction
	instrContent := fmt.Sprintf("TOC \\o \"1-%d\" \\h \\u", config.MaxLevel)
	tocFieldPara.Runs = append(tocFieldPara.Runs, Run{
		Properties: &RunProperties{
			Bold:     &Bold{},
			Color:    &Color{Val: "2F5496"},
			FontSize: &FontSize{Val: "32"},
		},
		InstrText: &InstrText{
			Space:   "preserve",
			Content: instrContent,
		},
	})

	// Add TOC field separator
	tocFieldPara.Runs = append(tocFieldPara.Runs, Run{
		Properties: &RunProperties{
			Bold:     &Bold{},
			Color:    &Color{Val: "2F5496"},
			FontSize: &FontSize{Val: "32"},
		},
		FieldChar: &FieldChar{
			FieldCharType: "separate",
		},
	})

	tocSDT.Content.Elements = append(tocSDT.Content.Elements, tocFieldPara)

	// Create a hyperlink paragraph for each entry
	for _, entry := range entries {
		entryPara := d.createTOCEntryWithFields(entry, config)
		tocSDT.Content.Elements = append(tocSDT.Content.Elements, entryPara)
	}

	// Add TOC field end paragraph
	endPara := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: "2"},
			Spacing: &Spacing{
				Before: "240",
				After:  "0",
			},
		},
		Runs: []Run{
			{
				Properties: &RunProperties{
					Color: &Color{Val: "2F5496"},
				},
				FieldChar: &FieldChar{
					FieldCharType: "end",
				},
			},
		},
	}

	tocSDT.Content.Elements = append(tocSDT.Content.Elements, endPara)
	elements = append(elements, tocSDT)

	return elements
}

// createTOCEntryWithFields creates a TOC entry with field codes.
func (d *Document) createTOCEntryWithFields(entry TOCEntry, config *TOCConfig) *Paragraph {
	// Determine the TOC style ID
	var styleVal string
	switch entry.Level {
	case 1:
		styleVal = "13" // TOC 1
	case 2:
		styleVal = "14" // TOC 2
	case 3:
		styleVal = "15" // TOC 3
	default:
		styleVal = fmt.Sprintf("%d", 12+entry.Level)
	}

	para := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: styleVal},
			Tabs: &Tabs{
				Tabs: []TabDef{
					{
						Val:    "right",
						Leader: "dot",
						Pos:    "8640",
					},
				},
			},
		},
		Runs: []Run{},
	}

	// Generate a unique bookmark ID for each entry
	anchor := fmt.Sprintf("_Toc%d", generateUniqueID(entry.Text))

	// Create hyperlink field begin
	para.Runs = append(para.Runs, Run{
		Properties: &RunProperties{
			Color: &Color{Val: "2F5496"},
		},
		FieldChar: &FieldChar{
			FieldCharType: "begin",
		},
	})

	// Add hyperlink instruction
	para.Runs = append(para.Runs, Run{
		InstrText: &InstrText{
			Space:   "preserve",
			Content: fmt.Sprintf(" HYPERLINK \\l %s ", anchor),
		},
	})

	// Hyperlink field separator
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "separate",
		},
	})

	// Add heading text
	para.Runs = append(para.Runs, Run{
		Text: Text{Content: entry.Text},
	})

	// Add tab character
	para.Runs = append(para.Runs, Run{
		Text: Text{Content: "\t"},
	})

	// Add page reference field
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "begin",
		},
	})

	para.Runs = append(para.Runs, Run{
		InstrText: &InstrText{
			Space:   "preserve",
			Content: fmt.Sprintf(" PAGEREF %s \\h ", anchor),
		},
	})

	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "separate",
		},
	})

	// Page number text
	para.Runs = append(para.Runs, Run{
		Text: Text{Content: fmt.Sprintf("%d", entry.PageNum)},
	})

	// Page number field end
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "end",
		},
	})

	// Hyperlink field end
	para.Runs = append(para.Runs, Run{
		Properties: &RunProperties{
			Color: &Color{Val: "2F5496"},
		},
		FieldChar: &FieldChar{
			FieldCharType: "end",
		},
	})

	return para
}

// generateUniqueID generates a unique ID based on text content.
func generateUniqueID(text string) int {
	// Use a simple hash algorithm to generate a unique ID
	hash := 0
	for _, char := range text {
		hash = hash*31 + int(char)
	}
	// Ensure the result is positive and within a reasonable range
	if hash < 0 {
		hash = -hash
	}
	return (hash % 90000) + 10000 // Generate a number between 10000-99999
}

// collectHeadingsAndAddBookmarks collects heading information and adds bookmarks.
func (d *Document) collectHeadingsAndAddBookmarks(maxLevel int) []TOCEntry {
	var entries []TOCEntry
	pageNum := 1 // Simplified; actual implementation would calculate real page numbers

	// Build a new Elements slice to insert bookmarks
	newElements := make([]interface{}, 0, len(d.Body.Elements)*2)
	entryIndex := 0

	for _, element := range d.Body.Elements {
		if paragraph, ok := element.(*Paragraph); ok {
			level := d.getHeadingLevel(paragraph)
			if level > 0 && level <= maxLevel {
				text := d.extractParagraphText(paragraph)
				if text != "" {
					// Generate a unique bookmark ID for each entry (consistent with the TOC entries)
					anchor := fmt.Sprintf("_Toc%d", generateUniqueID(text))

					entry := TOCEntry{
						Text:       text,
						Level:      level,
						PageNum:    pageNum,
						BookmarkID: anchor,
					}
					entries = append(entries, entry)

					// Add bookmark start before the heading paragraph
					bookmarkStart := &BookmarkStart{
						ID:   fmt.Sprintf("%d", entryIndex),
						Name: anchor,
					}
					newElements = append(newElements, bookmarkStart)

					// Add the original paragraph
					newElements = append(newElements, element)

					// Add bookmark end after the heading paragraph
					bookmarkEnd := &BookmarkEnd{
						ID: fmt.Sprintf("%d", entryIndex),
					}
					newElements = append(newElements, bookmarkEnd)

					entryIndex++
					continue
				}
			}
		}
		// Non-heading paragraphs are added directly
		newElements = append(newElements, element)
	}

	// Update the document elements
	d.Body.Elements = newElements

	return entries
}
