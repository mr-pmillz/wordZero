// Package document provides header and footer operations for Word documents
package document

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// HeaderFooterType represents the header/footer type
type HeaderFooterType string

const (
	// HeaderFooterTypeDefault represents the default header/footer
	HeaderFooterTypeDefault HeaderFooterType = "default"
	// HeaderFooterTypeFirst represents the first page header/footer
	HeaderFooterTypeFirst HeaderFooterType = "first"
	// HeaderFooterTypeEven represents the even page header/footer
	HeaderFooterTypeEven HeaderFooterType = "even"
)

// Header represents a header structure
type Header struct {
	XMLName     xml.Name     `xml:"w:hdr"`
	XmlnsWPC    string       `xml:"xmlns:wpc,attr"`
	XmlnsMC     string       `xml:"xmlns:mc,attr"`
	XmlnsO      string       `xml:"xmlns:o,attr"`
	XmlnsR      string       `xml:"xmlns:r,attr"`
	XmlnsM      string       `xml:"xmlns:m,attr"`
	XmlnsV      string       `xml:"xmlns:v,attr"`
	XmlnsWP14   string       `xml:"xmlns:wp14,attr"`
	XmlnsWP     string       `xml:"xmlns:wp,attr"`
	XmlnsW10    string       `xml:"xmlns:w10,attr"`
	XmlnsW      string       `xml:"xmlns:w,attr"`
	XmlnsW14    string       `xml:"xmlns:w14,attr"`
	XmlnsW15    string       `xml:"xmlns:w15,attr"`
	XmlnsWPG    string       `xml:"xmlns:wpg,attr"`
	XmlnsWPI    string       `xml:"xmlns:wpi,attr"`
	XmlnsWNE    string       `xml:"xmlns:wne,attr"`
	XmlnsWPS    string       `xml:"xmlns:wps,attr"`
	XmlnsWPSCD  string       `xml:"xmlns:wpsCustomData,attr"`
	MCIgnorable string       `xml:"mc:Ignorable,attr"`
	Paragraphs  []*Paragraph `xml:"w:p"`
}

// Footer represents a footer structure
type Footer struct {
	XMLName     xml.Name     `xml:"w:ftr"`
	XmlnsWPC    string       `xml:"xmlns:wpc,attr"`
	XmlnsMC     string       `xml:"xmlns:mc,attr"`
	XmlnsO      string       `xml:"xmlns:o,attr"`
	XmlnsR      string       `xml:"xmlns:r,attr"`
	XmlnsM      string       `xml:"xmlns:m,attr"`
	XmlnsV      string       `xml:"xmlns:v,attr"`
	XmlnsWP14   string       `xml:"xmlns:wp14,attr"`
	XmlnsWP     string       `xml:"xmlns:wp,attr"`
	XmlnsW10    string       `xml:"xmlns:w10,attr"`
	XmlnsW      string       `xml:"xmlns:w,attr"`
	XmlnsW14    string       `xml:"xmlns:w14,attr"`
	XmlnsW15    string       `xml:"xmlns:w15,attr"`
	XmlnsWPG    string       `xml:"xmlns:wpg,attr"`
	XmlnsWPI    string       `xml:"xmlns:wpi,attr"`
	XmlnsWNE    string       `xml:"xmlns:wne,attr"`
	XmlnsWPS    string       `xml:"xmlns:wps,attr"`
	XmlnsWPSCD  string       `xml:"xmlns:wpsCustomData,attr"`
	MCIgnorable string       `xml:"mc:Ignorable,attr"`
	Paragraphs  []*Paragraph `xml:"w:p"`
}

// HeaderFooterReference represents a header reference
type HeaderFooterReference struct {
	XMLName xml.Name `xml:"w:headerReference"`
	Type    string   `xml:"w:type,attr"`
	ID      string   `xml:"r:id,attr"`
}

// FooterReference represents a footer reference
type FooterReference struct {
	XMLName xml.Name `xml:"w:footerReference"`
	Type    string   `xml:"w:type,attr"`
	ID      string   `xml:"r:id,attr"`
}

// TitlePage represents the different first page setting
type TitlePage struct {
	XMLName xml.Name `xml:"w:titlePg"`
}

// PageNumber represents a page number field
type PageNumber struct {
	XMLName xml.Name `xml:"w:fldSimple"`
	Instr   string   `xml:"w:instr,attr"`
	Text    *Text    `xml:"w:t,omitempty"`
}

// createStandardHeader creates a standard header structure
func createStandardHeader() *Header {
	return &Header{
		XmlnsWPC:    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
		XmlnsMC:     "http://schemas.openxmlformats.org/markup-compatibility/2006",
		XmlnsO:      "urn:schemas-microsoft-com:office:office",
		XmlnsR:      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsM:      "http://schemas.openxmlformats.org/officeDocument/2006/math",
		XmlnsV:      "urn:schemas-microsoft-com:vml",
		XmlnsWP14:   "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
		XmlnsWP:     "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
		XmlnsW10:    "urn:schemas-microsoft-com:office:word",
		XmlnsW:      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsW14:    "http://schemas.microsoft.com/office/word/2010/wordml",
		XmlnsW15:    "http://schemas.microsoft.com/office/word/2012/wordml",
		XmlnsWPG:    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
		XmlnsWPI:    "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
		XmlnsWNE:    "http://schemas.microsoft.com/office/word/2006/wordml",
		XmlnsWPS:    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
		XmlnsWPSCD:  "http://www.wps.cn/officeDocument/2013/wpsCustomData",
		MCIgnorable: "w14 w15 wp14",
		Paragraphs:  make([]*Paragraph, 0),
	}
}

// createStandardFooter creates a standard footer structure
func createStandardFooter() *Footer {
	return &Footer{
		XmlnsWPC:    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
		XmlnsMC:     "http://schemas.openxmlformats.org/markup-compatibility/2006",
		XmlnsO:      "urn:schemas-microsoft-com:office:office",
		XmlnsR:      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsM:      "http://schemas.openxmlformats.org/officeDocument/2006/math",
		XmlnsV:      "urn:schemas-microsoft-com:vml",
		XmlnsWP14:   "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
		XmlnsWP:     "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
		XmlnsW10:    "urn:schemas-microsoft-com:office:word",
		XmlnsW:      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsW14:    "http://schemas.microsoft.com/office/word/2010/wordml",
		XmlnsW15:    "http://schemas.microsoft.com/office/word/2012/wordml",
		XmlnsWPG:    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
		XmlnsWPI:    "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
		XmlnsWNE:    "http://schemas.microsoft.com/office/word/2006/wordml",
		XmlnsWPS:    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
		XmlnsWPSCD:  "http://www.wps.cn/officeDocument/2013/wpsCustomData",
		MCIgnorable: "w14 w15 wp14",
		Paragraphs:  make([]*Paragraph, 0),
	}
}

// createPageNumberRuns creates a set of Runs for page number field codes
func createPageNumberRuns() []Run {
	return []Run{
		{
			FieldChar: &FieldChar{
				FieldCharType: "begin",
			},
		},
		{
			InstrText: &InstrText{
				Space:   "preserve",
				Content: " PAGE  \\* MERGEFORMAT ",
			},
		},
		{
			FieldChar: &FieldChar{
				FieldCharType: "separate",
			},
		},
		{
			Text: Text{
				Content: "1",
			},
		},
		{
			FieldChar: &FieldChar{
				FieldCharType: "end",
			},
		},
	}
}

// getFileNameForType returns the header/footer file name for the given type
func getFileNameForType(typePrefix string, headerType HeaderFooterType) string {
	switch headerType {
	case HeaderFooterTypeDefault:
		return fmt.Sprintf("%s1.xml", typePrefix)
	case HeaderFooterTypeFirst:
		return fmt.Sprintf("%sfirst.xml", typePrefix)
	case HeaderFooterTypeEven:
		return fmt.Sprintf("%seven.xml", typePrefix)
	default:
		return fmt.Sprintf("%s1.xml", typePrefix)
	}
}

// AddHeader adds a header to the document
func (d *Document) AddHeader(headerType HeaderFooterType, text string) error {
	header := createStandardHeader()

	// Create header paragraph
	paragraph := &Paragraph{}
	if text != "" {
		run := Run{
			Text: Text{
				Content: text,
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}
	header.Paragraphs = append(header.Paragraphs, paragraph)

	// Generate relationship ID
	headerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2 because rId1 is reserved for styles

	// Serialize header
	headerXML, err := xml.MarshalIndent(header, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize header: %v", err)
	}

	// Add XML declaration
	fullXML := append([]byte(xml.Header), headerXML...)

	// Get file name
	fileName := getFileNameForType("header", headerType)
	headerPartName := fmt.Sprintf("word/%s", fileName)

	// Store header content
	d.parts[headerPartName] = fullXML

	// Add relationship to document relationships
	relationship := Relationship{
		ID:     headerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// Add content type
	d.addContentType(headerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")

	// Update section properties
	d.addHeaderReference(headerType, headerID)

	return nil
}

// AddFooter adds a footer to the document
func (d *Document) AddFooter(footerType HeaderFooterType, text string) error {
	footer := createStandardFooter()

	// Create footer paragraph
	paragraph := &Paragraph{}
	if text != "" {
		run := Run{
			Text: Text{
				Content: text,
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}
	footer.Paragraphs = append(footer.Paragraphs, paragraph)

	// Generate relationship ID
	footerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2 because rId1 is reserved for styles

	// Serialize footer
	footerXML, err := xml.MarshalIndent(footer, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize footer: %v", err)
	}

	// Add XML declaration
	fullXML := append([]byte(xml.Header), footerXML...)

	// Get file name
	fileName := getFileNameForType("footer", footerType)
	footerPartName := fmt.Sprintf("word/%s", fileName)

	// Store footer content
	d.parts[footerPartName] = fullXML

	// Add relationship to document relationships
	relationship := Relationship{
		ID:     footerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// Add content type
	d.addContentType(footerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")

	// Update section properties
	d.addFooterReference(footerType, footerID)

	return nil
}

// AddHeaderWithPageNumber adds a header with a page number
func (d *Document) AddHeaderWithPageNumber(headerType HeaderFooterType, text string, showPageNum bool) error {
	header := createStandardHeader()

	// Create header paragraph
	paragraph := &Paragraph{}

	if text != "" {
		run := Run{
			Text: Text{
				Content: text,
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}

	if showPageNum {
		// Add "Page" prefix
		pageNumRun := Run{
			Text: Text{
				Content: " Page ",
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, pageNumRun)

		// Add page number field code
		pageNumberRuns := createPageNumberRuns()
		paragraph.Runs = append(paragraph.Runs, pageNumberRuns...)
	}

	header.Paragraphs = append(header.Paragraphs, paragraph)

	// Generate relationship ID
	headerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2 because rId1 is reserved for styles

	// Serialize header
	headerXML, err := xml.MarshalIndent(header, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize header: %v", err)
	}

	// Add XML declaration
	fullXML := append([]byte(xml.Header), headerXML...)

	// Get file name
	fileName := getFileNameForType("header", headerType)
	headerPartName := fmt.Sprintf("word/%s", fileName)

	// Store header content
	d.parts[headerPartName] = fullXML

	// Add relationship to document relationships
	relationship := Relationship{
		ID:     headerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// Add content type
	d.addContentType(headerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")

	// Update section properties
	d.addHeaderReference(headerType, headerID)

	return nil
}

// AddFooterWithPageNumber adds a footer with a page number
func (d *Document) AddFooterWithPageNumber(footerType HeaderFooterType, text string, showPageNum bool) error {
	footer := createStandardFooter()

	// Create footer paragraph
	paragraph := &Paragraph{}

	if text != "" {
		run := Run{
			Text: Text{
				Content: text,
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}

	if showPageNum {
		// Add "Page" prefix
		pageNumRun := Run{
			Text: Text{
				Content: " Page ",
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, pageNumRun)

		// Add page number field code
		pageNumberRuns := createPageNumberRuns()
		paragraph.Runs = append(paragraph.Runs, pageNumberRuns...)
	}

	footer.Paragraphs = append(footer.Paragraphs, paragraph)

	// Generate relationship ID
	footerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2 because rId1 is reserved for styles

	// Serialize footer
	footerXML, err := xml.MarshalIndent(footer, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize footer: %v", err)
	}

	// Add XML declaration
	fullXML := append([]byte(xml.Header), footerXML...)

	// Get file name
	fileName := getFileNameForType("footer", footerType)
	footerPartName := fmt.Sprintf("word/%s", fileName)

	// Store footer content
	d.parts[footerPartName] = fullXML

	// Add relationship to document relationships
	relationship := Relationship{
		ID:     footerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// Add content type
	d.addContentType(footerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")

	// Update section properties
	d.addFooterReference(footerType, footerID)

	return nil
}

// HeaderFooterConfig represents header/footer configuration
type HeaderFooterConfig struct {
	Text      string        // Text content
	Format    *TextFormat   // Text format configuration
	Alignment AlignmentType // Alignment
}

// createFormattedParagraph creates a formatted paragraph
func createFormattedParagraph(text string, format *TextFormat, alignment AlignmentType) *Paragraph {
	paragraph := &Paragraph{}

	// Set paragraph alignment
	if alignment != "" {
		paragraph.Properties = &ParagraphProperties{
			Justification: &Justification{Val: string(alignment)},
		}
	}

	// If there is text content, create a formatted Run
	if text != "" {
		run := Run{
			Text: Text{
				Content: text,
				Space:   "preserve",
			},
		}

		// Apply text formatting
		if format != nil {
			runProps := &RunProperties{}

			// Set font
			fontName := ""
			if format.FontFamily != "" {
				fontName = format.FontFamily
			} else if format.FontName != "" {
				fontName = format.FontName
			}
			if fontName != "" {
				runProps.FontFamily = &FontFamily{
					ASCII:    fontName,
					HAnsi:    fontName,
					EastAsia: fontName,
					CS:       fontName,
				}
			}

			// Set bold
			if format.Bold {
				runProps.Bold = &Bold{}
			}

			// Set italic
			if format.Italic {
				runProps.Italic = &Italic{}
			}

			// Set font color
			if format.FontColor != "" {
				// Ensure color format is correct (remove # prefix)
				color := strings.TrimPrefix(format.FontColor, "#")
				runProps.Color = &Color{Val: color}
			}

			// Set font size
			if format.FontSize > 0 {
				// Font size in Word is in half-points, so multiply by 2
				runProps.FontSize = &FontSize{Val: strconv.Itoa(format.FontSize * 2)}
			}

			// Set underline
			if format.Underline {
				runProps.Underline = &Underline{Val: "single"}
			}

			// Set strikethrough
			if format.Strike {
				runProps.Strike = &Strike{}
			}

			// Set highlight
			if format.Highlight != "" {
				runProps.Highlight = &Highlight{Val: format.Highlight}
			}

			run.Properties = runProps
		}

		paragraph.Runs = append(paragraph.Runs, run)
	}

	return paragraph
}

// AddFormattedHeader adds a formatted header.
//
// This method allows adding a header with custom text formatting and alignment.
//
// Parameters:
//   - headerType: Header type (HeaderFooterTypeDefault, HeaderFooterTypeFirst, HeaderFooterTypeEven)
//   - config: Header configuration containing text content, format, and alignment
//
// Example:
//
//	doc.AddFormattedHeader(document.HeaderFooterTypeDefault, &document.HeaderFooterConfig{
//		Text: "Company Report",
//		Format: &document.TextFormat{
//			FontSize:   10,
//			FontColor:  "8e8e8e",
//			FontFamily: "Arial",
//		},
//		Alignment: document.AlignCenter,
//	})
func (d *Document) AddFormattedHeader(headerType HeaderFooterType, config *HeaderFooterConfig) error {
	header := createStandardHeader()

	// Create formatted header paragraph
	if config == nil {
		config = &HeaderFooterConfig{}
	}
	paragraph := createFormattedParagraph(config.Text, config.Format, config.Alignment)
	header.Paragraphs = append(header.Paragraphs, paragraph)

	// Generate relationship ID
	headerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2 because rId1 is reserved for styles

	// Serialize header
	headerXML, err := xml.MarshalIndent(header, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize header: %v", err)
	}

	// Add XML declaration
	fullXML := append([]byte(xml.Header), headerXML...)

	// Get file name
	fileName := getFileNameForType("header", headerType)
	headerPartName := fmt.Sprintf("word/%s", fileName)

	// Store header content
	d.parts[headerPartName] = fullXML

	// Add relationship to document relationships
	relationship := Relationship{
		ID:     headerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// Add content type
	d.addContentType(headerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")

	// Update section properties
	d.addHeaderReference(headerType, headerID)

	return nil
}

// AddFormattedFooter adds a formatted footer.
//
// This method allows adding a footer with custom text formatting and alignment.
//
// Parameters:
//   - footerType: Footer type (HeaderFooterTypeDefault, HeaderFooterTypeFirst, HeaderFooterTypeEven)
//   - config: Footer configuration containing text content, format, and alignment
//
// Example:
//
//	doc.AddFormattedFooter(document.HeaderFooterTypeDefault, &document.HeaderFooterConfig{
//		Text: "Page 1",
//		Format: &document.TextFormat{
//			FontSize:   9,
//			FontColor:  "666666",
//			FontFamily: "Arial",
//		},
//		Alignment: document.AlignCenter,
//	})
func (d *Document) AddFormattedFooter(footerType HeaderFooterType, config *HeaderFooterConfig) error {
	footer := createStandardFooter()

	// Create formatted footer paragraph
	if config == nil {
		config = &HeaderFooterConfig{}
	}
	paragraph := createFormattedParagraph(config.Text, config.Format, config.Alignment)
	footer.Paragraphs = append(footer.Paragraphs, paragraph)

	// Generate relationship ID
	footerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2 because rId1 is reserved for styles

	// Serialize footer
	footerXML, err := xml.MarshalIndent(footer, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize footer: %v", err)
	}

	// Add XML declaration
	fullXML := append([]byte(xml.Header), footerXML...)

	// Get file name
	fileName := getFileNameForType("footer", footerType)
	footerPartName := fmt.Sprintf("word/%s", fileName)

	// Store footer content
	d.parts[footerPartName] = fullXML

	// Add relationship to document relationships
	relationship := Relationship{
		ID:     footerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// Add content type
	d.addContentType(footerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")

	// Update section properties
	d.addFooterReference(footerType, footerID)

	return nil
}

// SetDifferentFirstPage sets whether the first page has a different header/footer
func (d *Document) SetDifferentFirstPage(different bool) {
	sectPr := d.getSectionPropertiesForHeaderFooter()
	if different {
		sectPr.TitlePage = &TitlePage{}
	} else {
		sectPr.TitlePage = nil
	}
}

// addHeaderReference adds a header reference to the section properties
func (d *Document) addHeaderReference(headerType HeaderFooterType, headerID string) {
	sectPr := d.getSectionPropertiesForHeaderFooter()

	// Ensure the relationship namespace is set
	if sectPr.XmlnsR == "" {
		sectPr.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	}

	headerRef := &HeaderFooterReference{
		Type: string(headerType),
		ID:   headerID,
	}

	sectPr.HeaderReferences = append(sectPr.HeaderReferences, headerRef)
}

// addFooterReference adds a footer reference to the section properties
func (d *Document) addFooterReference(footerType HeaderFooterType, footerID string) {
	sectPr := d.getSectionPropertiesForHeaderFooter()

	// Ensure the relationship namespace is set
	if sectPr.XmlnsR == "" {
		sectPr.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	}

	footerRef := &FooterReference{
		Type: string(footerType),
		ID:   footerID,
	}

	sectPr.FooterReferences = append(sectPr.FooterReferences, footerRef)
}

// getSectionPropertiesForHeaderFooter returns or creates section properties with header/footer support
func (d *Document) getSectionPropertiesForHeaderFooter() *SectionProperties {
	// Check if section properties already exist in the document
	for _, element := range d.Body.Elements {
		if sectPr, ok := element.(*SectionProperties); ok {
			// Ensure the relationship namespace is set
			if sectPr.XmlnsR == "" {
				sectPr.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
			}
			return sectPr
		}
	}

	// If none exists, create new section properties
	sectPr := &SectionProperties{
		XMLName: xml.Name{Local: "w:sectPr"},
		XmlnsR:  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		PageNumType: &PageNumType{
			Fmt: "decimal",
		},
		Columns: &Columns{
			Space: "720",
			Num:   "1",
		},
	}
	d.Body.Elements = append(d.Body.Elements, sectPr)
	return sectPr
}

// addContentType adds a content type
func (d *Document) addContentType(partName, contentType string) {
	// Check if it already exists
	for _, override := range d.contentTypes.Overrides {
		if override.PartName == "/"+partName {
			return
		}
	}

	// Add new content type override
	override := Override{
		PartName:    "/" + partName,
		ContentType: contentType,
	}
	d.contentTypes.Overrides = append(d.contentTypes.Overrides, override)
}
