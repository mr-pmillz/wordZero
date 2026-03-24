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

	// ooXMLRelationshipsBase is the base namespace for OOXML relationships
	ooXMLRelationshipsBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
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
		XmlnsR:      ooXMLRelationshipsBase,
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
		XmlnsR:      ooXMLRelationshipsBase,
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

// headerFooterKind distinguishes between header and footer for shared helpers.
type headerFooterKind int

const (
	kindHeader headerFooterKind = iota
	kindFooter
)

// registerHeaderFooterPart serializes XML content for a header or footer, stores the part,
// adds the relationship and content type, and updates the section properties reference.
func (d *Document) registerHeaderFooterPart(kind headerFooterKind, hfType HeaderFooterType, xmlContent interface{}) error {
	relID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2)

	xmlBytes, err := xml.MarshalIndent(xmlContent, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize header/footer: %w", err)
	}
	fullXML := append([]byte(xml.Header), xmlBytes...)

	var typePrefix, relType, contentType string
	if kind == kindHeader {
		typePrefix = "header"
		relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
		contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	} else {
		typePrefix = "footer"
		relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
		contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
	}

	fileName := getFileNameForType(typePrefix, hfType)
	partName := fmt.Sprintf("word/%s", fileName)

	d.parts[partName] = fullXML

	relationship := Relationship{
		ID:     relID,
		Type:   relType,
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	d.addContentType(partName, contentType)

	if kind == kindHeader {
		d.addHeaderReference(hfType, relID)
	} else {
		d.addFooterReference(hfType, relID)
	}

	return nil
}

// AddHeader adds a header to the document
func (d *Document) AddHeader(headerType HeaderFooterType, text string) error {
	header := createStandardHeader()
	paragraph := &Paragraph{}
	if text != "" {
		run := Run{
			Text: Text{Content: text, Space: "preserve"},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}
	header.Paragraphs = append(header.Paragraphs, paragraph)
	return d.registerHeaderFooterPart(kindHeader, headerType, header)
}

// AddFooter adds a footer to the document
func (d *Document) AddFooter(footerType HeaderFooterType, text string) error {
	footer := createStandardFooter()
	paragraph := &Paragraph{}
	if text != "" {
		run := Run{
			Text: Text{Content: text, Space: "preserve"},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}
	footer.Paragraphs = append(footer.Paragraphs, paragraph)
	return d.registerHeaderFooterPart(kindFooter, footerType, footer)
}

// buildPageNumberParagraph creates a paragraph with optional text and page number runs.
func buildPageNumberParagraph(text string, showPageNum bool) *Paragraph {
	paragraph := &Paragraph{}
	if text != "" {
		run := Run{
			Text: Text{Content: text, Space: "preserve"},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}
	if showPageNum {
		pageNumRun := Run{
			Text: Text{Content: " Page ", Space: "preserve"},
		}
		paragraph.Runs = append(paragraph.Runs, pageNumRun)
		paragraph.Runs = append(paragraph.Runs, createPageNumberRuns()...)
	}
	return paragraph
}

// AddHeaderWithPageNumber adds a header with a page number
func (d *Document) AddHeaderWithPageNumber(headerType HeaderFooterType, text string, showPageNum bool) error {
	header := createStandardHeader()
	header.Paragraphs = append(header.Paragraphs, buildPageNumberParagraph(text, showPageNum))
	return d.registerHeaderFooterPart(kindHeader, headerType, header)
}

// AddFooterWithPageNumber adds a footer with a page number
func (d *Document) AddFooterWithPageNumber(footerType HeaderFooterType, text string, showPageNum bool) error {
	footer := createStandardFooter()
	footer.Paragraphs = append(footer.Paragraphs, buildPageNumberParagraph(text, showPageNum))
	return d.registerHeaderFooterPart(kindFooter, footerType, footer)
}

// HeaderFooterConfig represents header/footer configuration
type HeaderFooterConfig struct {
	Text      string        // Text content
	Format    *TextFormat   // Text format configuration
	Alignment AlignmentType // Alignment
}

// createFormattedParagraph creates a formatted paragraph
func createFormattedParagraph(text string, format *TextFormat, alignment AlignmentType) *Paragraph { //nolint:gocognit
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
	if config == nil {
		config = &HeaderFooterConfig{}
	}
	paragraph := createFormattedParagraph(config.Text, config.Format, config.Alignment)
	header.Paragraphs = append(header.Paragraphs, paragraph)
	return d.registerHeaderFooterPart(kindHeader, headerType, header)
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
	if config == nil {
		config = &HeaderFooterConfig{}
	}
	paragraph := createFormattedParagraph(config.Text, config.Format, config.Alignment)
	footer.Paragraphs = append(footer.Paragraphs, paragraph)
	return d.registerHeaderFooterPart(kindFooter, footerType, footer)
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
		sectPr.XmlnsR = ooXMLRelationshipsBase
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
		sectPr.XmlnsR = ooXMLRelationshipsBase
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
				sectPr.XmlnsR = ooXMLRelationshipsBase
			}
			return sectPr
		}
	}

	// If none exists, create new section properties
	sectPr := &SectionProperties{
		XMLName: xml.Name{Local: "w:sectPr"},
		XmlnsR:  ooXMLRelationshipsBase,
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

// textFormatToRunProperties converts a TextFormat to RunProperties.
// Returns nil if format is nil.
func textFormatToRunProperties(format *TextFormat) *RunProperties {
	if format == nil {
		return nil
	}
	props := &RunProperties{}
	if format.Bold {
		props.Bold = &Bold{}
	}
	if format.Italic {
		props.Italic = &Italic{}
	}
	if format.FontSize > 0 {
		// Font size in OOXML is in half-points
		props.FontSize = &FontSize{Val: strconv.Itoa(format.FontSize * 2)}
	}
	if format.FontColor != "" {
		props.Color = &Color{Val: strings.TrimPrefix(format.FontColor, "#")}
	}
	if format.FontFamily != "" {
		props.FontFamily = &FontFamily{ASCII: format.FontFamily, HAnsi: format.FontFamily}
	} else if format.FontName != "" {
		props.FontFamily = &FontFamily{ASCII: format.FontName, HAnsi: format.FontName}
	}
	if format.Underline {
		props.Underline = &Underline{Val: "single"}
	}
	if format.Strike {
		props.Strike = &Strike{}
	}
	if format.Highlight != "" {
		props.Highlight = &Highlight{Val: format.Highlight}
	}
	return props
}

// AddStyleHeader adds a styled header with formatted text, optional red secondary text,
// and a horizontal rule at the bottom.
//
// Parameters:
//   - headerType: Header type (HeaderFooterTypeDefault, HeaderFooterTypeFirst, HeaderFooterTypeEven)
//   - text: Primary header text (may contain newlines which are converted to line breaks)
//   - redText: Optional secondary text displayed in red below the primary text
//   - format: Optional text formatting for the primary text (can be nil)
func (d *Document) AddStyleHeader(headerType HeaderFooterType, text, redText string, format *TextFormat) error {
	header := createStandardHeader()

	// Create header paragraph with formatting
	paragraph := &Paragraph{}

	// Split text by newlines and add runs with line breaks between
	lines := strings.Split(text, "\n")
	for i, line := range lines {
		if i > 0 {
			// Add line break between lines
			paragraph.Runs = append(paragraph.Runs, Run{Break: &Break{}})
		}
		run := Run{
			Text:       Text{Content: line, Space: "preserve"},
			Properties: textFormatToRunProperties(format), // fresh copy per run
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}

	// Add red text if provided
	if redText != "" {
		paragraph.Runs = append(paragraph.Runs, Run{Break: &Break{}})
		redProps := textFormatToRunProperties(format)
		if redProps == nil {
			redProps = &RunProperties{}
		}
		redProps.Color = &Color{Val: "FF0000"}
		paragraph.Runs = append(paragraph.Runs, Run{
			Text:       Text{Content: redText, Space: "preserve"},
			Properties: redProps,
		})
	}

	// Add horizontal rule (bottom border)
	paragraph.SetHorizontalRule(BorderStyleSingle, 12, "000000")

	header.Paragraphs = append(header.Paragraphs, paragraph)

	return d.registerHeaderFooterPart(kindHeader, headerType, header)
}

// ClearHeaderFooterReferences removes all header and footer references from section properties.
func (d *Document) ClearHeaderFooterReferences() {
	for _, elem := range d.Body.Elements {
		if sectPr, ok := elem.(*SectionProperties); ok {
			sectPr.HeaderReferences = nil
			sectPr.FooterReferences = nil
		}
	}
}

// AddCurrentHeaderReference adds a header reference to the current (last) section properties.
func (d *Document) AddCurrentHeaderReference(headerType HeaderFooterType, headerID string) {
	sectPr := d.getSectionPropertiesForHeaderFooter()
	headerRef := &HeaderFooterReference{
		Type: string(headerType),
		ID:   headerID,
	}
	sectPr.HeaderReferences = append(sectPr.HeaderReferences, headerRef)
}

// AddCurrentFooterReference adds a footer reference to the current (last) section properties.
func (d *Document) AddCurrentFooterReference(footerType HeaderFooterType, footerID string) {
	sectPr := d.getSectionPropertiesForHeaderFooter()
	footerRef := &FooterReference{
		Type: string(footerType),
		ID:   footerID,
	}
	sectPr.FooterReferences = append(sectPr.FooterReferences, footerRef)
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
