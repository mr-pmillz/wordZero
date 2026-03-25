// Package document provides core operations for Word documents
package document

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/mr-pmillz/wordZero/pkg/style"
)

// XML element and attribute name constants used in parsing.
const (
	xmlElemSectPr   = "sectPr"
	xmlElemGraphic  = "graphic"
	xmlAttrDistT    = "distT"
	xmlAttrDistB    = "distB"
	xmlAttrDistL    = "distL"
	xmlAttrDistR    = "distR"
	xmlAttrName     = "name"
	xmlAttrDescr    = "descr"
	xmlAttrTitle    = "title"
)

// Document represents a Word document
type Document struct {
	// Main document content
	Body *Body
	// Document relationships
	relationships *Relationships
	// Document-level relationships (for headers, footers, etc.)
	documentRelationships *Relationships
	// Content types
	contentTypes *ContentTypes
	// Style manager
	styleManager *style.StyleManager
	// Temporary storage for document parts
	parts map[string][]byte
	// Image ID counter, ensures each image has a unique ID
	nextImageID int
	// Footnote/endnote manager (per-document, avoids global state leaks)
	footnoteManager *FootnoteManager
	// Numbering manager (per-document, avoids global state leaks)
	numberingManager *NumberingManager
}

// Body represents the document body
type Body struct {
	XMLName  xml.Name      `xml:"w:body"`
	Elements []interface{} `xml:"-"` // Not serialized directly; uses custom method
}

// MarshalXML performs custom XML serialization, outputting elements in order
func (b *Body) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Start element
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Separate SectionProperties from other elements
	var sectPr *SectionProperties
	var otherElements []interface{}

	for _, element := range b.Elements {
		if sp, ok := element.(*SectionProperties); ok {
			sectPr = sp // Keep the last SectionProperties
		} else {
			otherElements = append(otherElements, element)
		}
	}

	// Serialize other elements first (paragraphs, tables, etc.)
	for _, element := range otherElements {
		if err := e.Encode(element); err != nil {
			return err
		}
	}

	// Serialize SectionProperties last (if present)
	if sectPr != nil {
		if err := e.Encode(sectPr); err != nil {
			return err
		}
	}

	// End element
	return e.EncodeToken(start.End())
}

// BodyElement is the interface for document body elements
type BodyElement interface {
	ElementType() string
}

// ElementType returns the paragraph element type
func (p *Paragraph) ElementType() string {
	return "paragraph"
}

// ElementType returns the table element type
func (t *Table) ElementType() string {
	return "table"
}

// RawXMLElement preserves an unknown XML element encountered during document parsing.
// It stores the original element name, attributes, and inner XML content so it can
// be re-emitted during serialization without loss.
type RawXMLElement struct {
	XMLName  xml.Name
	Attrs    []xml.Attr `xml:"-"`
	InnerXML string     `xml:",innerxml"`
}

// ElementType returns "raw_xml" for preserved unknown elements.
func (r *RawXMLElement) ElementType() string {
	return "raw_xml"
}

// MarshalXML custom serializes the raw XML element, preserving original attributes.
func (r *RawXMLElement) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Use the preserved element name and attributes
	start.Name = r.XMLName
	start.Attr = r.Attrs

	// For elements with inner XML, we need to write the raw content
	// Use a temporary struct with innerxml tag
	type rawContent struct {
		InnerXML string `xml:",innerxml"`
	}

	return e.EncodeElement(rawContent{InnerXML: r.InnerXML}, start)
}

// Paragraph represents a paragraph
type Paragraph struct {
	XMLName        xml.Name             `xml:"w:p"`
	Properties     *ParagraphProperties `xml:"w:pPr,omitempty"`
	Runs           []Run                `xml:"w:r"`
	RawXMLElements []*RawXMLElement     `xml:"-"` // preserved elements (bookmarks, etc.) for round-trip
}

// MarshalXML custom serializes a Paragraph, emitting properties, runs, then raw elements.
func (p *Paragraph) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "w:p"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Emit paragraph properties
	if p.Properties != nil {
		if err := e.EncodeElement(p.Properties, xml.StartElement{Name: xml.Name{Local: "w:pPr"}}); err != nil {
			return err
		}
	}

	// Emit runs
	for i := range p.Runs {
		if err := e.EncodeElement(&p.Runs[i], xml.StartElement{Name: xml.Name{Local: "w:r"}}); err != nil {
			return err
		}
	}

	// Emit preserved raw XML elements (bookmarks, etc.)
	for _, raw := range p.RawXMLElements {
		if err := e.EncodeElement(raw, xml.StartElement{Name: raw.XMLName}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// ParagraphProperties represents paragraph properties
type ParagraphProperties struct {
	XMLName             xml.Name             `xml:"w:pPr"`
	ParagraphStyle      *ParagraphStyle      `xml:"w:pStyle,omitempty"`
	NumberingProperties *NumberingProperties `xml:"w:numPr,omitempty"`
	ParagraphBorder     *ParagraphBorder     `xml:"w:pBdr,omitempty"`
	Tabs                *Tabs                `xml:"w:tabs,omitempty"`
	SnapToGrid          *SnapToGrid          `xml:"w:snapToGrid,omitempty"` // Snap to grid setting
	Spacing             *Spacing             `xml:"w:spacing,omitempty"`
	Indentation         *Indentation         `xml:"w:ind,omitempty"`
	Justification       *Justification       `xml:"w:jc,omitempty"`
	KeepNext            *KeepNext            `xml:"w:keepNext,omitempty"`        // Keep with next paragraph
	KeepLines           *KeepLines           `xml:"w:keepLines,omitempty"`       // Keep all lines together
	PageBreakBefore     *PageBreakBefore     `xml:"w:pageBreakBefore,omitempty"` // Page break before paragraph
	WidowControl        *WidowControl        `xml:"w:widowControl,omitempty"`    // Widow/orphan control
	OutlineLevel        *OutlineLevel        `xml:"w:outlineLvl,omitempty"`      // Outline level
	SectionProperties   *SectionProperties   `xml:"w:sectPr,omitempty"`          // Section properties (for section breaks)
}

// SnapToGrid controls snap-to-grid alignment.
// Set to "0" or "false" to disable grid alignment, allowing custom line spacing to take effect.
// Note: This type has an identical definition in the style package; this is intentional since both packages can be used independently.
type SnapToGrid struct {
	XMLName xml.Name `xml:"w:snapToGrid"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// ParagraphBorder represents paragraph borders
type ParagraphBorder struct {
	XMLName xml.Name             `xml:"w:pBdr"`
	Top     *ParagraphBorderLine `xml:"w:top,omitempty"`
	Left    *ParagraphBorderLine `xml:"w:left,omitempty"`
	Bottom  *ParagraphBorderLine `xml:"w:bottom,omitempty"`
	Right   *ParagraphBorderLine `xml:"w:right,omitempty"`
}

// ParagraphBorderLine represents a paragraph border line
type ParagraphBorderLine struct {
	Val   string `xml:"w:val,attr"`
	Color string `xml:"w:color,attr"`
	Sz    string `xml:"w:sz,attr"`
	Space string `xml:"w:space,attr"`
}

// Spacing represents spacing settings
type Spacing struct {
	XMLName  xml.Name `xml:"w:spacing"`
	Before   string   `xml:"w:before,attr,omitempty"`
	After    string   `xml:"w:after,attr,omitempty"`
	Line     string   `xml:"w:line,attr,omitempty"`
	LineRule string   `xml:"w:lineRule,attr,omitempty"`
}

// Justification represents text alignment
type Justification struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// KeepNext keeps the paragraph with the next paragraph
type KeepNext struct {
	XMLName xml.Name `xml:"w:keepNext"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// KeepLines keeps all lines in the paragraph together
type KeepLines struct {
	XMLName xml.Name `xml:"w:keepLines"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// PageBreakBefore inserts a page break before the paragraph
type PageBreakBefore struct {
	XMLName xml.Name `xml:"w:pageBreakBefore"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// WidowControl controls widow/orphan lines
type WidowControl struct {
	XMLName xml.Name `xml:"w:widowControl"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// OutlineLevel represents the outline level
type OutlineLevel struct {
	XMLName xml.Name `xml:"w:outlineLvl"`
	Val     string   `xml:"w:val,attr"`
}

// Run represents a text run
type Run struct {
	XMLName               xml.Name               `xml:"w:r"`
	Properties            *RunProperties         `xml:"w:rPr,omitempty"`
	FootnoteReference     *FootnoteReference     `xml:"w:footnoteReference,omitempty"`
	EndnoteReference      *EndnoteReference      `xml:"w:endnoteReference,omitempty"`
	FootnoteRef           *FootnoteRef           `xml:"w:footnoteRef,omitempty"`
	EndnoteRef            *EndnoteRef            `xml:"w:endnoteRef,omitempty"`
	Separator             *Separator             `xml:"w:separator,omitempty"`
	ContinuationSeparator *ContinuationSeparator `xml:"w:continuationSeparator,omitempty"`
	Text                  Text                   `xml:"w:t,omitempty"`
	Break                 *Break                 `xml:"w:br,omitempty"` // Page break
	Drawing               *DrawingElement        `xml:"w:drawing,omitempty"`
	FieldChar             *FieldChar             `xml:"w:fldChar,omitempty"`
	InstrText             *InstrText             `xml:"w:instrText,omitempty"`
	RawXMLContent         []*RawXMLElement       `xml:"-"` // preserved unknown elements for round-trip
}

// MarshalXML performs custom XML serialization for Run.
// This method ensures only non-empty elements are serialized, especially for Drawing elements.
//
//nolint:gocognit
func (r *Run) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Start Run element
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Serialize RunProperties (if present)
	if r.Properties != nil {
		if err := e.EncodeElement(r.Properties, xml.StartElement{Name: xml.Name{Local: "w:rPr"}}); err != nil {
			return err
		}
	}

	// Serialize footnote/endnote reference elements (before Text)
	if r.FootnoteReference != nil {
		if err := e.EncodeElement(r.FootnoteReference, xml.StartElement{Name: xml.Name{Local: "w:footnoteReference"}}); err != nil {
			return err
		}
	}
	if r.EndnoteReference != nil {
		if err := e.EncodeElement(r.EndnoteReference, xml.StartElement{Name: xml.Name{Local: "w:endnoteReference"}}); err != nil {
			return err
		}
	}
	if r.FootnoteRef != nil {
		if err := e.EncodeElement(r.FootnoteRef, xml.StartElement{Name: xml.Name{Local: "w:footnoteRef"}}); err != nil {
			return err
		}
	}
	if r.EndnoteRef != nil {
		if err := e.EncodeElement(r.EndnoteRef, xml.StartElement{Name: xml.Name{Local: "w:endnoteRef"}}); err != nil {
			return err
		}
	}
	if r.Separator != nil {
		if err := e.EncodeElement(r.Separator, xml.StartElement{Name: xml.Name{Local: "w:separator"}}); err != nil {
			return err
		}
	}
	if r.ContinuationSeparator != nil {
		if err := e.EncodeElement(r.ContinuationSeparator, xml.StartElement{Name: xml.Name{Local: "w:continuationSeparator"}}); err != nil {
			return err
		}
	}

	// Serialize Text (only when it has content)
	// This is a key fix: avoid serializing empty Text elements
	if r.Text.Content != "" {
		if err := e.EncodeElement(r.Text, xml.StartElement{Name: xml.Name{Local: "w:t"}}); err != nil {
			return err
		}
	}

	// Serialize Break (if present)
	if r.Break != nil {
		if err := e.EncodeElement(r.Break, xml.StartElement{Name: xml.Name{Local: "w:br"}}); err != nil {
			return err
		}
	}

	// Serialize Drawing (if present)
	if r.Drawing != nil {
		if err := e.EncodeElement(r.Drawing, xml.StartElement{Name: xml.Name{Local: "w:drawing"}}); err != nil {
			return err
		}
	}

	// Serialize FieldChar (if present)
	if r.FieldChar != nil {
		if err := e.EncodeElement(r.FieldChar, xml.StartElement{Name: xml.Name{Local: "w:fldChar"}}); err != nil {
			return err
		}
	}

	// Serialize InstrText (if present)
	if r.InstrText != nil {
		if err := e.EncodeElement(r.InstrText, xml.StartElement{Name: xml.Name{Local: "w:instrText"}}); err != nil {
			return err
		}
	}

	// Serialize preserved raw XML content (tab, lastRenderedPageBreak, commentReference, etc.)
	for _, raw := range r.RawXMLContent {
		if err := e.EncodeElement(raw, xml.StartElement{Name: raw.XMLName}); err != nil {
			return err
		}
	}

	// End Run element
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// RunProperties represents text run properties.
// Note: Field order must conform to the OpenXML standard; w:rStyle comes first, and w:rFonts must precede w:color.
type RunProperties struct {
	XMLName       xml.Name           `xml:"w:rPr"`
	RunStyle      *RunStyle          `xml:"w:rStyle,omitempty"`
	FontFamily    *FontFamily        `xml:"w:rFonts,omitempty"`
	Bold          *Bold              `xml:"w:b,omitempty"`
	BoldCs        *BoldCs            `xml:"w:bCs,omitempty"`
	Italic        *Italic            `xml:"w:i,omitempty"`
	ItalicCs      *ItalicCs          `xml:"w:iCs,omitempty"`
	Underline     *Underline         `xml:"w:u,omitempty"`
	Strike        *Strike            `xml:"w:strike,omitempty"`
	Color         *Color             `xml:"w:color,omitempty"`
	FontSize      *FontSize          `xml:"w:sz,omitempty"`
	FontSizeCs    *FontSizeCs        `xml:"w:szCs,omitempty"`
	Highlight     *Highlight         `xml:"w:highlight,omitempty"`
	VerticalAlign *VerticalAlignment `xml:"w:vertAlign,omitempty"`
}

// RunStyle represents a character style reference
type RunStyle struct {
	XMLName xml.Name `xml:"w:rStyle"`
	Val     string   `xml:"w:val,attr"`
}

// VerticalAlignment represents vertical alignment (superscript/subscript)
type VerticalAlignment struct {
	XMLName xml.Name `xml:"w:vertAlign"`
	Val     string   `xml:"w:val,attr"`
}

// FootnoteRef is the footnote self-reference element (marks the number position within footnote content)
type FootnoteRef struct {
	XMLName xml.Name `xml:"w:footnoteRef"`
}

// EndnoteRef is the endnote self-reference element (marks the number position within endnote content)
type EndnoteRef struct {
	XMLName xml.Name `xml:"w:endnoteRef"`
}

// Separator is the separator line element used in footnote/endnote separators
type Separator struct {
	XMLName xml.Name `xml:"w:separator"`
}

// ContinuationSeparator is the continuation separator element for multi-page footnotes/endnotes
type ContinuationSeparator struct {
	XMLName xml.Name `xml:"w:continuationSeparator"`
}

// Bold represents bold formatting
type Bold struct {
	XMLName xml.Name `xml:"w:b"`
}

// BoldCs represents complex script bold formatting
type BoldCs struct {
	XMLName xml.Name `xml:"w:bCs"`
}

// Italic represents italic formatting
type Italic struct {
	XMLName xml.Name `xml:"w:i"`
}

// ItalicCs represents complex script italic formatting
type ItalicCs struct {
	XMLName xml.Name `xml:"w:iCs"`
}

// Underline represents underline formatting
type Underline struct {
	XMLName xml.Name `xml:"w:u"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// Strike represents strikethrough formatting
type Strike struct {
	XMLName xml.Name `xml:"w:strike"`
}

// FontSize represents font size
type FontSize struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     string   `xml:"w:val,attr"`
}

// FontSizeCs represents complex script font size
type FontSizeCs struct {
	XMLName xml.Name `xml:"w:szCs"`
	Val     string   `xml:"w:val,attr"`
}

// Color represents text color
type Color struct {
	XMLName xml.Name `xml:"w:color"`
	Val     string   `xml:"w:val,attr"`
}

// Highlight represents background highlight color
type Highlight struct {
	XMLName xml.Name `xml:"w:highlight"`
	Val     string   `xml:"w:val,attr"`
}

// Text represents text content
type Text struct {
	XMLName xml.Name `xml:"w:t"`
	Space   string   `xml:"xml:space,attr,omitempty"`
	Content string   `xml:",chardata"`
}

// Break represents page breaks in Word documents
type Break struct {
	XMLName xml.Name `xml:"w:br"`
	Type    string   `xml:"w:type,attr,omitempty"` // "page" indicates a page break
}

// Relationships represents document relationships
type Relationships struct {
	XMLName       xml.Name       `xml:"Relationships"`
	Xmlns         string         `xml:"xmlns,attr"`
	Relationships []Relationship `xml:"Relationship"`
}

// Relationship represents a single relationship
type Relationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

// ContentTypes represents content types
type ContentTypes struct {
	XMLName   xml.Name   `xml:"Types"`
	Xmlns     string     `xml:"xmlns,attr"`
	Defaults  []Default  `xml:"Default"`
	Overrides []Override `xml:"Override"`
}

// Default represents a default content type
type Default struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// Override represents an override content type
type Override struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// FontFamily represents a font family
type FontFamily struct {
	XMLName  xml.Name `xml:"w:rFonts"`
	ASCII    string   `xml:"w:ascii,attr,omitempty"`
	HAnsi    string   `xml:"w:hAnsi,attr,omitempty"`
	EastAsia string   `xml:"w:eastAsia,attr,omitempty"`
	CS       string   `xml:"w:cs,attr,omitempty"`
	Hint     string   `xml:"w:hint,attr,omitempty"`
}

// TextFormat represents text formatting configuration
type TextFormat struct {
	Bold       bool   // Whether to apply bold
	Italic     bool   // Whether to apply italic
	FontSize   int    // Font size in points
	FontColor  string // Font color (hex, e.g. "FF0000" for red)
	FontFamily string // Font name (preferred field)
	FontName   string // Font name alias (for backward compatibility with earlier examples/README that used FontName)
	Underline  bool   // Whether to apply underline
	Strike     bool   // Strikethrough
	Highlight  string // Highlight color
}

// AlignmentType represents text alignment type
type AlignmentType string

const (
	// AlignLeft is left alignment
	AlignLeft AlignmentType = "left"
	// AlignCenter is center alignment
	AlignCenter AlignmentType = "center"
	// AlignRight is right alignment
	AlignRight AlignmentType = "right"
	// AlignJustify is justified alignment
	AlignJustify AlignmentType = "both"
)

// SpacingConfig represents spacing configuration
type SpacingConfig struct {
	LineSpacing     float64 // Line spacing multiplier (e.g. 1.5 for 1.5x line spacing)
	BeforePara      int     // Spacing before paragraph (in points)
	AfterPara       int     // Spacing after paragraph (in points)
	FirstLineIndent int     // First line indent (in points)
}

// Indentation represents indentation settings
type Indentation struct {
	XMLName   xml.Name `xml:"w:ind"`
	FirstLine string   `xml:"w:firstLine,attr,omitempty"`
	Left      string   `xml:"w:left,attr,omitempty"`
	Right     string   `xml:"w:right,attr,omitempty"`
}

// Tabs represents tab stop settings
type Tabs struct {
	XMLName xml.Name `xml:"w:tabs"`
	Tabs    []TabDef `xml:"w:tab"`
}

// TabDef represents a tab stop definition
type TabDef struct {
	XMLName xml.Name `xml:"w:tab"`
	Val     string   `xml:"w:val,attr"`
	Leader  string   `xml:"w:leader,attr,omitempty"`
	Pos     string   `xml:"w:pos,attr"`
}

// ParagraphStyle represents a paragraph style reference
type ParagraphStyle struct {
	XMLName xml.Name `xml:"w:pStyle"`
	Val     string   `xml:"w:val,attr"`
}

// NumberingProperties represents paragraph numbering properties
type NumberingProperties struct {
	XMLName xml.Name `xml:"w:numPr"`
	ILevel  *ILevel  `xml:"w:ilvl,omitempty"`
	NumID   *NumID   `xml:"w:numId,omitempty"`
}

// ILevel represents the numbering level
type ILevel struct {
	XMLName xml.Name `xml:"w:ilvl"`
	Val     string   `xml:"w:val,attr"`
}

// NumID represents the numbering ID
type NumID struct {
	XMLName xml.Name `xml:"w:numId"`
	Val     string   `xml:"w:val,attr"`
}

// New creates a new Word document
func New() *Document {
	DebugMsg(MsgCreatingNewDocument)

	doc := &Document{
		Body: &Body{
			Elements: make([]interface{}, 0),
		},
		styleManager: style.NewStyleManager(),
		parts:        make(map[string][]byte),
		nextImageID:  0, // Initialize image ID counter, starting from 0
		documentRelationships: &Relationships{
			Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
			Relationships: []Relationship{},
		},
		footnoteManager: &FootnoteManager{
			nextFootnoteID: 1,
			nextEndnoteID:  1,
			footnotes:      make(map[string]*Footnote),
			endnotes:       make(map[string]*Endnote),
		},
	}

	// Initialize document structure
	doc.initializeStructure()

	return doc
}

// Open opens an existing Word document.
//
// The filename parameter is the path to the .docx file to open.
// This function parses the entire document structure, including text content, formatting, and properties.
//
// Returns an error if the file does not exist, has an invalid format, or fails to parse.
//
// Example:
//
//	doc, err := document.Open("existing.docx")
//	if err != nil {
//		log.Fatal(err)
//	}
//
//	// Print all paragraph content
//	for i, para := range doc.Body.Paragraphs {
//		fmt.Printf("Paragraph %d: ", i+1)
//		for _, run := range para.Runs {
//			fmt.Print(run.Text.Content)
//		}
//		fmt.Println()
//	}
func Open(filename string) (*Document, error) {
	InfoMsgf(MsgOpeningDocumentPath, filename)

	reader, err := zip.OpenReader(filename)
	if err != nil {
		ErrorMsgf(MsgFailedToOpenFile, filename)
		return nil, WrapErrorWithContext("open_file", err, filename)
	}
	defer reader.Close()

	doc, err := openFromZipReader(&reader.Reader, filename)
	if err != nil {
		return nil, err
	}

	InfoMsgf(MsgDocumentOpenedPath, filename)
	return doc, nil
}

func OpenFromMemory(readCloser io.ReadCloser) (*Document, error) {
	defer readCloser.Close()
	InfoMsg(MsgOpeningDocument)

	fileData, err := io.ReadAll(readCloser)
	if err != nil {
		return nil, fmt.Errorf("failed to read file content: %w", err)
	}

	readerAt := bytes.NewReader(fileData)
	size := int64(len(fileData))
	reader, err := zip.NewReader(readerAt, size)
	if err != nil {
		ErrorMsg(MsgFailedToOpenFileSimple)
		return nil, WrapErrorWithContext("open_file", err, "")
	}

	doc, err := openFromZipReader(reader, "memory")
	if err != nil {
		return nil, err
	}

	InfoMsg(MsgDocumentOpened)
	return doc, nil
}

func openFromZipReader(zipReader *zip.Reader, filename string) (*Document, error) {
	doc := &Document{
		parts: make(map[string][]byte),
		documentRelationships: &Relationships{
			Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
			Relationships: []Relationship{},
		},
		nextImageID: 0, // Initialize image ID counter, starting from 0
		footnoteManager: &FootnoteManager{
			nextFootnoteID: 1,
			nextEndnoteID:  1,
			footnotes:      make(map[string]*Footnote),
			endnotes:       make(map[string]*Endnote),
		},
	}

	// Read all file parts
	for _, file := range zipReader.File {
		rc, err := file.Open()
		if err != nil {
			ErrorMsgf(MsgFailedToOpenFilePart, file.Name)
			return nil, WrapErrorWithContext("open_part", err, file.Name)
		}

		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			ErrorMsgf(MsgFailedToReadFilePart, file.Name)
			return nil, WrapErrorWithContext("read_part", err, file.Name)
		}

		doc.parts[file.Name] = data
		DebugMsgf(MsgReadFilePart, file.Name, len(data))
	}

	// Initialize style manager
	doc.styleManager = style.NewStyleManager()

	// Parse content types
	if err := doc.parseContentTypes(); err != nil {
		DebugMsgf(MsgFailedToParseContentTypesDefault, err)
		// If parsing fails, use defaults
		doc.contentTypes = &ContentTypes{
			Xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
			Defaults: []Default{
				{Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"},
				{Extension: "xml", ContentType: "application/xml"},
			},
			Overrides: []Override{
				{PartName: "/word/document.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"},
				{PartName: "/word/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"},
			},
		}
	}

	// Parse relationships
	if err := doc.parseRelationships(); err != nil {
		DebugMsgf(MsgFailedToParseRelationshipsDefault, err)
		// If parsing fails, use defaults
		doc.relationships = &Relationships{
			Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
			Relationships: []Relationship{
				{
					ID:     "rId1",
					Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
					Target: "word/document.xml",
				},
			},
		}
	}

	// Parse main document
	if err := doc.parseDocument(); err != nil {
		ErrorMsgf(MsgFailedToParseDocument, filename)
		return nil, WrapErrorWithContext("parse_document", err, filename)
	}

	// Parse styles file
	if err := doc.parseStyles(); err != nil {
		DebugMsgf(MsgFailedToParseStylesDefault, err)
		// If style parsing fails, reinitialize with default styles
		doc.styleManager = style.NewStyleManager()
	}

	// Parse document relationships (including relationships for images and other resources)
	if err := doc.parseDocumentRelationships(); err != nil {
		DebugMsgf(MsgFailedToParseDocRelDefault, err)
		// If parsing fails, keep the initialized empty relationship list
	}

	// Update nextImageID counter based on existing image relationships
	doc.updateNextImageID()

	// Sync footnote manager with existing footnotes/endnotes in the template
	// to avoid ID collisions and preserve system notes (separator, continuationNotice, etc.)
	doc.syncFootnoteManagerWithExisting()

	return doc, nil
}

// Save saves the document to the specified file path.
//
// The filename parameter is the path to save the file, including file name and extension.
// If the directory does not exist, the required directory structure is created automatically.
//
// The save process includes serializing all document content, compressing it into ZIP format,
// and writing it to the file system.
//
// Example:
//
//	doc := document.New()
//	doc.AddParagraph("Example content")
//
//	// Save to current directory
//	err := doc.Save("example.docx")
//
//	// Save to subdirectory (directory created automatically)
//	err = doc.Save("output/documents/example.docx")
//
//	if err != nil {
//		log.Fatal(err)
//	}
func (d *Document) Save(filename string) error {
	InfoMsgf(MsgSavingDocument, filename)

	// Ensure directory exists
	dir := filepath.Dir(filename)
	if err := os.MkdirAll(dir, 0755); err != nil {
		ErrorMsgf(MsgFailedToCreateDirectory, dir)
		return WrapErrorWithContext("create_dir", err, dir)
	}

	// Create file
	file, err := os.Create(filename)
	if err != nil {
		ErrorMsgf(MsgFailedToCreateFile, filename)
		return WrapErrorWithContext("create_file", err, filename)
	}
	defer file.Close()

	// Create ZIP writer
	zipWriter := zip.NewWriter(file)
	defer zipWriter.Close()

	// Serialize main document
	if err := d.serializeDocument(); err != nil {
		ErrorMsg(MsgFailedToSerializeDocument)
		return WrapError("serialize_document", err)
	}

	// Serialize styles
	if err := d.serializeStyles(); err != nil {
		ErrorMsg(MsgFailedToSerializeStyles)
		return WrapError("serialize_styles", err)
	}

	// Serialize content types
	d.serializeContentTypes()

	// Serialize relationships
	d.serializeRelationships()

	// Serialize document relationships
	d.serializeDocumentRelationships()

	// Write all parts
	for name, data := range d.parts {
		writer, err := zipWriter.Create(name)
		if err != nil {
			ErrorMsgf(MsgFailedToCreateZIPEntry, name)
			return WrapErrorWithContext("create_zip_entry", err, name)
		}

		if _, err := writer.Write(data); err != nil {
			ErrorMsgf(MsgFailedToWriteZIPEntry, name)
			return WrapErrorWithContext("write_zip_entry", err, name)
		}

		DebugMsgf(MsgWrittenZIPEntry, name, len(data))
	}

	InfoMsgf(MsgDocumentSaved, filename)
	return nil
}

// AddParagraph adds a plain paragraph to the document.
//
// The text parameter is the paragraph's text content. The paragraph uses default formatting,
// which can be modified later through the returned Paragraph pointer.
//
// Returns a pointer to the newly created paragraph for further formatting.
//
// Example:
//
//	doc := document.New()
//
//	// Add a plain paragraph
//	para := doc.AddParagraph("This is a paragraph")
//
//	// Set paragraph properties
//	para.SetAlignment(document.AlignCenter)
//	para.SetSpacing(&document.SpacingConfig{
//		LineSpacing: 1.5,
//		BeforePara:  12,
//	})
func (d *Document) AddParagraph(text string) *Paragraph {
	DebugMsgf(MsgAddingParagraph, text)
	p := &Paragraph{
		Runs: []Run{
			{
				Text: Text{
					Content: text,
					Space:   "preserve",
				},
			},
		},
	}

	d.Body.Elements = append(d.Body.Elements, p)
	return p
}

// applyTextFormat applies a TextFormat to RunProperties, setting font, bold, italic,
// color, size, underline, strikethrough, and highlight as specified.
func applyTextFormat(runProps *RunProperties, format *TextFormat) {
	if format == nil {
		return
	}

	// Support both FontFamily and FontName fields
	fontName := ""
	if format.FontFamily != "" {
		fontName = format.FontFamily
	} else if format.FontName != "" { // Backward compatibility
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

	if format.Bold {
		runProps.Bold = &Bold{}
	}

	if format.Italic {
		runProps.Italic = &Italic{}
	}

	if format.FontColor != "" {
		color := strings.TrimPrefix(format.FontColor, "#")
		runProps.Color = &Color{Val: color}
	}

	if format.FontSize > 0 {
		// Word uses half-points for font size, so multiply by 2
		runProps.FontSize = &FontSize{Val: strconv.Itoa(format.FontSize * 2)}
	}

	if format.Underline {
		runProps.Underline = &Underline{Val: "single"}
	}

	if format.Strike {
		runProps.Strike = &Strike{}
	}

	if format.Highlight != "" {
		runProps.Highlight = &Highlight{Val: format.Highlight}
	}
}

// AddFormattedParagraph adds a formatted paragraph to the document.
//
// The text parameter is the paragraph's text content.
// The format parameter specifies text formatting; if nil, default formatting is used.
//
// Returns a pointer to the newly created paragraph for further property setting.
//
// Example:
//
//	doc := document.New()
//
//	// Create format configuration
//	titleFormat := &document.TextFormat{
//		Bold:      true,
//		FontSize:  18,
//		FontColor: "FF0000", // Red
//		FontName:  "Arial",
//	}
//
//	// Add formatted title
//	title := doc.AddFormattedParagraph("Document Title", titleFormat)
//	title.SetAlignment(document.AlignCenter)
func (d *Document) AddFormattedParagraph(text string, format *TextFormat) *Paragraph {
	DebugMsgf(MsgAddingFormattedParagraph, text)

	runProps := &RunProperties{}
	applyTextFormat(runProps, format)

	p := &Paragraph{
		Runs: []Run{
			{
				Properties: runProps,
				Text: Text{
					Content: text,
					Space:   "preserve",
				},
			},
		},
	}

	d.Body.Elements = append(d.Body.Elements, p)
	return p
}

// SetAlignment sets the paragraph alignment.
//
// The alignment parameter specifies the alignment type. Supported values:
//   - AlignLeft: left alignment (default)
//   - AlignCenter: center alignment
//   - AlignRight: right alignment
//   - AlignJustify: justified alignment
//
// Example:
//
//	para := doc.AddParagraph("Centered title")
//	para.SetAlignment(document.AlignCenter)
//
//	para2 := doc.AddParagraph("Right-aligned text")
//	para2.SetAlignment(document.AlignRight)
func (p *Paragraph) SetAlignment(alignment AlignmentType) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	p.Properties.Justification = &Justification{Val: string(alignment)}
	DebugMsgf(MsgSettingParagraphAlignment, alignment)
}

// SetSpacing sets the paragraph spacing configuration.
//
// The config parameter contains various spacing settings; if nil, no changes are made.
// Configuration options include:
//   - LineSpacing: line spacing multiplier (e.g. 1.5 for 1.5x line spacing)
//   - BeforePara: spacing before paragraph (in points)
//   - AfterPara: spacing after paragraph (in points)
//   - FirstLineIndent: first line indent (in points)
//
// Note: Spacing values are automatically converted to Word's internal TWIPs unit (1 point = 20 TWIPs).
//
// Example:
//
//	para := doc.AddParagraph("Paragraph with spacing")
//
//	// Set complex spacing
//	para.SetSpacing(&document.SpacingConfig{
//		LineSpacing:     1.5, // 1.5x line spacing
//		BeforePara:      12,  // 12pt before paragraph
//		AfterPara:       6,   // 6pt after paragraph
//		FirstLineIndent: 24,  // 24pt first line indent
//	})
//
//	// Set only line spacing
//	para2 := doc.AddParagraph("Double spaced")
//	para2.SetSpacing(&document.SpacingConfig{
//		LineSpacing: 2.0,
//	})
func (p *Paragraph) SetSpacing(config *SpacingConfig) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if config != nil {
		spacing := &Spacing{}

		if config.BeforePara > 0 {
			// Convert to TWIPs (1/20 of a point)
			spacing.Before = strconv.Itoa(config.BeforePara * 20)
		}

		if config.AfterPara > 0 {
			// Convert to TWIPs (1/20 of a point)
			spacing.After = strconv.Itoa(config.AfterPara * 20)
		}

		if config.LineSpacing > 0 {
			// Line spacing; 240 represents single spacing
			spacing.Line = strconv.Itoa(int(config.LineSpacing * 240))
		}

		p.Properties.Spacing = spacing

		if config.FirstLineIndent > 0 {
			if p.Properties.Indentation == nil {
				p.Properties.Indentation = &Indentation{}
			}
			// Convert to TWIPs (1/20 of a point)
			p.Properties.Indentation.FirstLine = strconv.Itoa(config.FirstLineIndent * 20)
		}

		DebugMsgf(MsgSettingParagraphSpacing,
			config.BeforePara, config.AfterPara, config.LineSpacing, config.FirstLineIndent)
	}
}

// AddFormattedText adds formatted text content to the paragraph.
//
// This method allows mixing different text formats within a single paragraph.
// The new text is added as a new Run in the paragraph.
//
// The text parameter is the text content to add.
// The format parameter specifies text formatting; if nil, default formatting is used.
//
// Example:
//
//	para := doc.AddParagraph("This paragraph contains ")
//
//	// Add bold red text
//	para.AddFormattedText("bold red", &document.TextFormat{
//		Bold: true,
//		FontColor: "FF0000",
//	})
//
//	// Add plain text
//	para.AddFormattedText(" and plain text", nil)
//
//	// Add italic blue text
//	para.AddFormattedText(" and italic blue", &document.TextFormat{
//		Italic: true,
//		FontColor: "0000FF",
//		FontSize: 14,
//	})
func (p *Paragraph) AddFormattedText(text string, format *TextFormat) {
	runProps := &RunProperties{}
	applyTextFormat(runProps, format)

	run := Run{
		Properties: runProps,
		Text: Text{
			Content: text,
			Space:   "preserve",
		},
	}

	p.Runs = append(p.Runs, run)
	DebugMsgf(MsgAddingFormattedText, text)
}

// AddPageBreak adds a page break to the paragraph.
//
// This method adds a page break within the current paragraph; content after the break appears on a new page.
// Unlike Document.AddPageBreak(), this method does not create a new paragraph but adds the break within the current paragraph's runs.
//
// Example:
//
//	para := doc.AddParagraph("First page content")
//	para.AddPageBreak()
//	para.AddFormattedText("Second page content", nil)
func (p *Paragraph) AddPageBreak() {
	run := Run{
		Break: &Break{
			Type: "page",
		},
	}
	p.Runs = append(p.Runs, run)
	DebugMsg(MsgAddingPageBreakToParagraph)
}

// AddLineBreak adds a line break to the paragraph, optionally followed by text.
// If text is non-empty, it is added as a separate run after the break.
func (p *Paragraph) AddLineBreak(text string) {
	// Add a run with a Break element (line break, not page break)
	breakRun := Run{
		Break: &Break{}, // Empty Type = line break (no Type attr = line break)
	}
	p.Runs = append(p.Runs, breakRun)

	// If text provided, add it as a separate run after the break
	if text != "" {
		textRun := Run{
			Text: Text{Content: text},
		}
		p.Runs = append(p.Runs, textRun)
	}
}

// AddRun adds a formatted text run to the paragraph.
// format and runProps are both optional (can be nil).
// If both are provided, runProps takes precedence for overlapping properties.
func (p *Paragraph) AddRun(text string, format *TextFormat, runProps *RunProperties) {
	run := Run{
		Text: Text{Content: text},
	}

	if runProps != nil {
		run.Properties = runProps
	} else if format != nil {
		props := &RunProperties{}
		applyTextFormat(props, format)
		run.Properties = props
	}

	p.Runs = append(p.Runs, run)
}

// AddHeadingParagraph adds a heading paragraph to the document.
//
// The text parameter is the heading's text content.
// The level parameter is the heading level (1-9), corresponding to Heading1 through Heading9.
//
// Returns a pointer to the newly created paragraph for further property setting.
// This method automatically sets the correct style reference so the heading is recognized by Word's navigation pane.
//
// Example:
//
//	doc := document.New()
//
//	// Add level 1 heading
//	h1 := doc.AddHeadingParagraph("Chapter 1: Overview", 1)
//
//	// Add level 2 heading
//	h2 := doc.AddHeadingParagraph("1.1 Background", 2)
//
//	// Add level 3 heading
//	h3 := doc.AddHeadingParagraph("1.1.1 Research Goals", 3)
func (d *Document) AddHeadingParagraph(text string, level int) *Paragraph {
	return d.AddHeadingParagraphWithBookmark(text, level, "")
}

// AddHeadingParagraphWithBookmark adds a heading paragraph with a bookmark to the document.
//
// The text parameter is the heading's text content.
// The level parameter is the heading level (1-9), corresponding to Heading1 through Heading9.
// The bookmarkName parameter is the bookmark name; if empty, no bookmark is added.
//
// Returns a pointer to the newly created paragraph for further property setting.
// This method automatically sets the correct style reference so the heading is recognized by Word's navigation pane,
// and adds a bookmark when needed to support table of contents navigation and hyperlinks.
//
// Example:
//
//	doc := document.New()
//
//	// Add level 1 heading with bookmark
//	h1 := doc.AddHeadingParagraphWithBookmark("Chapter 1: Overview", 1, "chapter1")
//
//	// Add level 2 heading without bookmark
//	h2 := doc.AddHeadingParagraphWithBookmark("1.1 Background", 2, "")
//
//	// Add level 3 heading with auto-generated bookmark name
//	h3 := doc.AddHeadingParagraphWithBookmark("1.1.1 Research Goals", 3, "auto_bookmark")
func (d *Document) AddHeadingParagraphWithBookmark(text string, level int, bookmarkName string) *Paragraph {
	return d.AddHeadingParagraphWithBookmarkFormatted(text, level, bookmarkName, nil)
}

// AddHeadingParagraphWithBookmarkFormatted adds a heading paragraph with a bookmark and optional text formatting.
// If format is nil, default heading style formatting is used.
//
// The text parameter is the heading's text content.
// The level parameter is the heading level (1-9), corresponding to Heading1 through Heading9.
// The bookmarkName parameter is the bookmark name; if empty, no bookmark is added.
// The format parameter allows overriding the default heading style formatting.
//
// Returns a pointer to the newly created paragraph for further property setting.
//
// Example:
//
//	doc := document.New()
//
//	// Add heading with custom formatting
//	h1 := doc.AddHeadingParagraphWithBookmarkFormatted("Chapter 1", 1, "ch1", &document.TextFormat{
//		Bold:      true,
//		FontSize:  24,
//		FontColor: "FF0000",
//	})
func (d *Document) AddHeadingParagraphWithBookmarkFormatted(text string, level int, bookmarkName string, format *TextFormat) *Paragraph {
	if level < 1 || level > 9 {
		DebugMsgf(MsgHeadingLevelOutOfRange, level)
		level = 1
	}

	styleID := fmt.Sprintf("Heading%d", level)
	DebugMsgf(MsgAddingHeadingParagraph, text, level, styleID, bookmarkName)

	// Get the style from the style manager
	headingStyle := d.styleManager.GetStyle(styleID)
	if headingStyle == nil {
		DebugMsgf(MsgStyleNotFoundUsingDefault, styleID)
		return d.AddParagraph(text)
	}

	// Build run properties from explicit format, falling back to style defaults
	var runProps *RunProperties
	if format != nil {
		runProps = &RunProperties{}
		applyTextFormat(runProps, format)
	} else {
		// Apply character formatting from the style
		runProps = &RunProperties{}
		if headingStyle.RunPr != nil {
			if headingStyle.RunPr.Bold != nil {
				runProps.Bold = &Bold{}
			}
			if headingStyle.RunPr.Italic != nil {
				runProps.Italic = &Italic{}
			}
			if headingStyle.RunPr.FontSize != nil {
				runProps.FontSize = &FontSize{Val: headingStyle.RunPr.FontSize.Val}
			}
			if headingStyle.RunPr.Color != nil {
				runProps.Color = &Color{Val: headingStyle.RunPr.Color.Val}
			}
			if headingStyle.RunPr.FontFamily != nil {
				runProps.FontFamily = &FontFamily{ASCII: headingStyle.RunPr.FontFamily.ASCII}
			}
		}
	}

	// Create paragraph properties, applying paragraph formatting from the style
	paraProps := &ParagraphProperties{
		ParagraphStyle: &ParagraphStyle{Val: styleID},
	}

	// Apply paragraph formatting from the style
	if headingStyle.ParagraphPr != nil {
		if headingStyle.ParagraphPr.Spacing != nil {
			paraProps.Spacing = &Spacing{
				Before: headingStyle.ParagraphPr.Spacing.Before,
				After:  headingStyle.ParagraphPr.Spacing.After,
				Line:   headingStyle.ParagraphPr.Spacing.Line,
			}
		}
		if headingStyle.ParagraphPr.Justification != nil {
			paraProps.Justification = &Justification{
				Val: headingStyle.ParagraphPr.Justification.Val,
			}
		}
		if headingStyle.ParagraphPr.Indentation != nil {
			paraProps.Indentation = &Indentation{
				FirstLine: headingStyle.ParagraphPr.Indentation.FirstLine,
				Left:      headingStyle.ParagraphPr.Indentation.Left,
				Right:     headingStyle.ParagraphPr.Indentation.Right,
			}
		}
	}

	// Create the paragraph's Run list
	runs := make([]Run, 0)

	// If a bookmark is needed, add a bookmark start marker at the beginning
	if bookmarkName != "" {
		// Generate a unique bookmark ID
		bookmarkID := fmt.Sprintf("bookmark_%d_%s", len(d.Body.Elements), bookmarkName)

		// Add bookmark start marker as a separate element in the document body
		d.Body.Elements = append(d.Body.Elements, &BookmarkStart{
			ID:   bookmarkID,
			Name: bookmarkName,
		})

		DebugMsgf(MsgAddingBookmarkStart, bookmarkID, bookmarkName)
	}

	// Add text content
	runs = append(runs, Run{
		Properties: runProps,
		Text: Text{
			Content: text,
			Space:   "preserve",
		},
	})

	// Create paragraph
	p := &Paragraph{
		Properties: paraProps,
		Runs:       runs,
	}

	d.Body.Elements = append(d.Body.Elements, p)

	// If a bookmark is needed, add a bookmark end marker after the paragraph
	if bookmarkName != "" {
		bookmarkID := fmt.Sprintf("bookmark_%d_%s", len(d.Body.Elements)-2, bookmarkName) // -2 because the paragraph has already been added

		// Add bookmark end marker
		d.Body.Elements = append(d.Body.Elements, &BookmarkEnd{
			ID: bookmarkID,
		})

		DebugMsgf(MsgAddingBookmarkEnd, bookmarkID)
	}

	return p
}

// AddPageBreak adds a page break to the document.
//
// The page break forces a new page to begin at the current position.
// This method creates a paragraph containing the page break.
//
// Example:
//
//	doc := document.New()
//	doc.AddParagraph("First page content")
//	doc.AddPageBreak()
//	doc.AddParagraph("Second page content")
func (d *Document) AddPageBreak() {
	DebugMsg(MsgAddingPageBreak)

	// Create a paragraph containing a page break
	p := &Paragraph{
		Runs: []Run{
			{
				Break: &Break{
					Type: "page",
				},
			},
		},
	}

	d.Body.Elements = append(d.Body.Elements, p)
}

// SetStyle sets the paragraph style.
//
// The styleID parameter is the style ID to apply, such as "Heading1", "Normal", etc.
// This method sets the paragraph's style reference, ensuring it uses the specified style.
//
// Example:
//
//	para := doc.AddParagraph("This is a paragraph")
//	para.SetStyle("Heading2")  // Set to Heading 2 style
func (p *Paragraph) SetStyle(styleID string) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	p.Properties.ParagraphStyle = &ParagraphStyle{Val: styleID}
	DebugMsgf(MsgSettingParagraphStyle, styleID)
}

// SetIndentation sets the paragraph indentation properties.
//
// Parameters:
//   - firstLineCm: first line indent in centimeters (negative values create a hanging indent)
//   - leftCm: left indent in centimeters
//   - rightCm: right indent in centimeters
//
// Example:
//
//	para := doc.AddParagraph("This is an indented paragraph")
//	para.SetIndentation(0.5, 0, 0)    // First line indent 0.5cm
//	para.SetIndentation(-0.5, 1, 0)  // Hanging indent 0.5cm, left indent 1cm
func (p *Paragraph) SetIndentation(firstLineCm, leftCm, rightCm float64) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if p.Properties.Indentation == nil {
		p.Properties.Indentation = &Indentation{}
	}

	// Convert centimeters to TWIPs (1 cm = 567 TWIPs)
	if firstLineCm != 0 {
		p.Properties.Indentation.FirstLine = strconv.Itoa(int(firstLineCm * 567))
	}

	if leftCm != 0 {
		p.Properties.Indentation.Left = strconv.Itoa(int(leftCm * 567))
	}

	if rightCm != 0 {
		p.Properties.Indentation.Right = strconv.Itoa(int(rightCm * 567))
	}

	DebugMsgf(MsgSettingParagraphIndent, firstLineCm, leftCm, rightCm)
}

// SetKeepWithNext sets the paragraph to stay on the same page as the next paragraph.
//
// This method ensures the current paragraph and the next paragraph are not separated by a page break.
// Commonly used for heading-body combinations or content that must remain contiguous.
//
// Parameters:
//   - keep: true to enable, false to disable
//
// Example:
//
//	// Keep heading with next paragraph
//	title := doc.AddParagraph("Chapter 1: Overview")
//	title.SetKeepWithNext(true)
//	doc.AddParagraph("This chapter introduces...")  // Stays on the same page as the heading
func (p *Paragraph) SetKeepWithNext(keep bool) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if keep {
		p.Properties.KeepNext = &KeepNext{Val: "1"}
		DebugMsg(MsgSettingKeepWithNext)
	} else {
		p.Properties.KeepNext = nil
		DebugMsg(MsgUnsettingKeepWithNext)
	}
}

// SetKeepLines sets all lines in the paragraph to stay on the same page.
//
// This method prevents the paragraph from being split across multiple pages,
// ensuring all lines are displayed on the same page.
//
// Parameters:
//   - keep: true to enable, false to disable
//
// Example:
//
//	// Ensure the entire paragraph is not split across pages
//	para := doc.AddParagraph("This is an important paragraph that must be displayed in full.")
//	para.SetKeepLines(true)
func (p *Paragraph) SetKeepLines(keep bool) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if keep {
		p.Properties.KeepLines = &KeepLines{Val: "1"}
		DebugMsg(MsgSettingKeepLinesTogether)
	} else {
		p.Properties.KeepLines = nil
		DebugMsg(MsgUnsettingKeepLinesTogether)
	}
}

// SetPageBreakBefore sets a page break before the paragraph.
//
// This method forces a page break before the paragraph, making it start on a new page.
// Commonly used for chapter headings or content that needs its own page.
//
// Parameters:
//   - pageBreak: true to enable page break before, false to disable
//
// Example:
//
//	// Chapter heading starts on a new page
//	chapter := doc.AddParagraph("Chapter 2: Detailed Description")
//	chapter.SetPageBreakBefore(true)
func (p *Paragraph) SetPageBreakBefore(pageBreak bool) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if pageBreak {
		p.Properties.PageBreakBefore = &PageBreakBefore{Val: "1"}
		DebugMsg(MsgSettingPageBreakBefore)
	} else {
		p.Properties.PageBreakBefore = nil
		DebugMsg(MsgUnsettingPageBreakBefore)
	}
}

// SetWidowControl sets the paragraph's widow/orphan control.
//
// Widow/orphan control prevents the first or last line of a paragraph from appearing
// alone at the bottom or top of a page, improving document layout quality.
//
// Parameters:
//   - control: true to enable widow/orphan control (default), false to disable
//
// Example:
//
//	para := doc.AddParagraph("This is a long paragraph...")
//	para.SetWidowControl(true)  // Enable widow/orphan control
func (p *Paragraph) SetWidowControl(control bool) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if control {
		p.Properties.WidowControl = &WidowControl{Val: "1"}
		DebugMsg(MsgEnablingWidowOrphanControl)
	} else {
		p.Properties.WidowControl = &WidowControl{Val: "0"}
		DebugMsg(MsgDisablingWidowOrphanControl)
	}
}

// SetOutlineLevel sets the paragraph's outline level.
//
// The outline level is used to display document structure in the navigation pane; the range is 0-8.
// Typically used for heading paragraphs in conjunction with the table of contents feature.
//
// Parameters:
//   - level: outline level, an integer between 0 and 8 (0 for body text, 1-8 for Heading1 through Heading8)
//
// Example:
//
//	// Set to level 1 heading outline level
//	title := doc.AddParagraph("Chapter 1")
//	title.SetOutlineLevel(0)  // Corresponds to Heading1
//
//	// Set to level 2 heading outline level
//	subtitle := doc.AddParagraph("1.1 Overview")
//	subtitle.SetOutlineLevel(1)  // Corresponds to Heading2
func (p *Paragraph) SetOutlineLevel(level int) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if level < 0 || level > 8 {
		WarnMsg(MsgOutlineLevelAdjusted)
		if level < 0 {
			level = 0
		} else {
			level = 8
		}
	}

	p.Properties.OutlineLevel = &OutlineLevel{Val: strconv.Itoa(level)}
	DebugMsgf(MsgSettingParagraphOutlineLevel, level)
}

// SetSnapToGrid sets the paragraph's snap-to-grid property.
//
// Snap-to-grid controls whether paragraph lines align to the document grid. When the document
// has grid settings enabled (common in CJK documents with the "snap to document grid" option),
// custom line spacing may not take effect precisely because lines automatically align to grid lines.
//
// By setting snapToGrid to false, grid alignment is disabled for this paragraph,
// allowing custom line spacing to take effect precisely.
//
// Parameters:
//   - snapToGrid: true to enable grid alignment (default), false to disable grid alignment
//
// Example:
//
//	// Disable grid alignment so custom line spacing takes effect precisely
//	para := doc.AddParagraph("This text uses precise line spacing")
//	para.SetSpacing(&document.SpacingConfig{LineSpacing: 1.5})
//	para.SetSnapToGrid(false)  // Disable grid alignment
func (p *Paragraph) SetSnapToGrid(snapToGrid bool) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	if !snapToGrid {
		p.Properties.SnapToGrid = &SnapToGrid{Val: "0"}
		DebugMsg(MsgDisablingParagraphGrid)
	} else {
		// When enabling grid alignment, remove the setting (use default behavior)
		p.Properties.SnapToGrid = nil
		DebugMsg(MsgEnablingParagraphGrid)
	}
}

// ParagraphFormatConfig represents paragraph format configuration.
//
// This struct provides a unified configuration interface for all paragraph format properties,
// allowing multiple paragraph properties to be set at once for improved readability and ease of use.
type ParagraphFormatConfig struct {
	// Basic formatting
	Alignment AlignmentType // Alignment (AlignLeft, AlignCenter, AlignRight, AlignJustify)
	Style     string        // Paragraph style ID (e.g. "Heading1", "Normal")

	// Spacing settings
	LineSpacing     float64 // Line spacing multiplier (e.g. 1.5 for 1.5x line spacing)
	BeforePara      int     // Spacing before paragraph (in points)
	AfterPara       int     // Spacing after paragraph (in points)
	FirstLineIndent int     // First line indent (in points)

	// Indentation settings
	FirstLineCm float64 // First line indent in centimeters (negative values create a hanging indent)
	LeftCm      float64 // Left indent (in centimeters)
	RightCm     float64 // Right indent (in centimeters)

	// Pagination and control
	KeepWithNext    bool  // Keep with next paragraph on the same page
	KeepLines       bool  // Keep all lines in the paragraph on the same page
	PageBreakBefore bool  // Page break before paragraph
	WidowControl    bool  // Widow/orphan control
	SnapToGrid      *bool // Snap to grid (set to false to disable grid alignment, allowing custom line spacing to take effect)

	// Outline level
	OutlineLevel int // Outline level (0-8; 0 for body text, 1-8 for Heading1 through Heading8)
}

// SetParagraphFormat sets all paragraph format properties at once using a configuration.
//
// This method provides a convenient way to set all paragraph format properties
// without calling multiple individual setter methods. Only non-zero properties are applied.
//
// Parameters:
//   - config: paragraph format configuration containing all format properties
//
// Example:
//
//	// Create a paragraph with full formatting
//	para := doc.AddParagraph("Important Chapter Title")
//	para.SetParagraphFormat(&document.ParagraphFormatConfig{
//		Alignment:       document.AlignCenter,
//		Style:           "Heading1",
//		LineSpacing:     1.5,
//		BeforePara:      24,
//		AfterPara:       12,
//		KeepWithNext:    true,
//		PageBreakBefore: true,
//		OutlineLevel:    0,
//	})
//
//	// Set an indented body paragraph
//	para2 := doc.AddParagraph("Body content...")
//	para2.SetParagraphFormat(&document.ParagraphFormatConfig{
//		Alignment:       document.AlignJustify,
//		FirstLineCm:     0.5,
//		LineSpacing:     1.5,
//		BeforePara:      6,
//		AfterPara:       6,
//		WidowControl:    true,
//	})
func (p *Paragraph) SetParagraphFormat(config *ParagraphFormatConfig) {
	if config == nil {
		return
	}

	// Set alignment
	if config.Alignment != "" {
		p.SetAlignment(config.Alignment)
	}

	// Set style
	if config.Style != "" {
		p.SetStyle(config.Style)
	}

	// Set spacing (if any spacing settings are provided)
	if config.LineSpacing > 0 || config.BeforePara > 0 || config.AfterPara > 0 || config.FirstLineIndent > 0 {
		p.SetSpacing(&SpacingConfig{
			LineSpacing:     config.LineSpacing,
			BeforePara:      config.BeforePara,
			AfterPara:       config.AfterPara,
			FirstLineIndent: config.FirstLineIndent,
		})
	}

	// Set indentation (if any indentation settings are provided)
	if config.FirstLineCm != 0 || config.LeftCm != 0 || config.RightCm != 0 {
		p.SetIndentation(config.FirstLineCm, config.LeftCm, config.RightCm)
	}

	// Set pagination and control properties
	p.SetKeepWithNext(config.KeepWithNext)
	p.SetKeepLines(config.KeepLines)
	p.SetPageBreakBefore(config.PageBreakBefore)
	p.SetWidowControl(config.WidowControl)

	// Set grid alignment
	if config.SnapToGrid != nil {
		p.SetSnapToGrid(*config.SnapToGrid)
	}

	// Set outline level
	if config.OutlineLevel >= 0 && config.OutlineLevel <= 8 {
		p.SetOutlineLevel(config.OutlineLevel)
	}

	DebugMsgf(MsgApplyingParagraphFormat,
		config.Alignment, config.Style, config.LineSpacing, config.BeforePara, config.AfterPara)
}

// ParagraphBorderConfig represents paragraph border configuration (distinct from table border configuration)
type ParagraphBorderConfig struct {
	Style BorderStyle // Border style
	Size  int         // Border width (in 1/8 points; recommended default is 12, i.e. 1.5pt)
	Color string      // Border color (hex, e.g. "000000" for black)
	Space int         // Spacing between border and text (in points; recommended default is 1)
}

// SetBorder sets the paragraph borders.
//
// This method adds border decorations to a paragraph, particularly useful for rendering Markdown horizontal rules (---).
//
// Parameters:
//   - top: top border configuration; pass nil to skip the top border
//   - left: left border configuration; pass nil to skip the left border
//   - bottom: bottom border configuration; pass nil to skip the bottom border
//   - right: right border configuration; pass nil to skip the right border
//
// Border configuration includes style, width, color, and spacing properties.
//
// Example:
//
//	// Set horizontal rule effect (bottom border only)
//	para := doc.AddParagraph("")
//	para.SetBorder(nil, nil, &document.ParagraphBorderConfig{
//		Style: document.BorderStyleSingle,
//		Size:  12,       // 1.5pt width
//		Color: "000000", // Black
//		Space: 1,        // 1pt spacing
//	}, nil)
//
//	// Set full border
//	para := doc.AddParagraph("Bordered paragraph")
//	borderConfig := &document.ParagraphBorderConfig{
//		Style: document.BorderStyleDouble,
//		Size:  8,
//		Color: "0000FF", // Blue
//		Space: 2,
//	}
//	para.SetBorder(borderConfig, borderConfig, borderConfig, borderConfig)
func (p *Paragraph) SetBorder(top, left, bottom, right *ParagraphBorderConfig) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	// If no border configuration is provided, clear borders
	if top == nil && left == nil && bottom == nil && right == nil {
		p.Properties.ParagraphBorder = nil
		return
	}

	// Create paragraph border
	if p.Properties.ParagraphBorder == nil {
		p.Properties.ParagraphBorder = &ParagraphBorder{}
	}

	// Set top border
	if top != nil {
		p.Properties.ParagraphBorder.Top = &ParagraphBorderLine{
			Val:   string(top.Style),
			Sz:    strconv.Itoa(top.Size),
			Color: top.Color,
			Space: strconv.Itoa(top.Space),
		}
	} else {
		p.Properties.ParagraphBorder.Top = nil
	}

	// Set left border
	if left != nil {
		p.Properties.ParagraphBorder.Left = &ParagraphBorderLine{
			Val:   string(left.Style),
			Sz:    strconv.Itoa(left.Size),
			Color: left.Color,
			Space: strconv.Itoa(left.Space),
		}
	} else {
		p.Properties.ParagraphBorder.Left = nil
	}

	// Set bottom border
	if bottom != nil {
		p.Properties.ParagraphBorder.Bottom = &ParagraphBorderLine{
			Val:   string(bottom.Style),
			Sz:    strconv.Itoa(bottom.Size),
			Color: bottom.Color,
			Space: strconv.Itoa(bottom.Space),
		}
	} else {
		p.Properties.ParagraphBorder.Bottom = nil
	}

	// Set right border
	if right != nil {
		p.Properties.ParagraphBorder.Right = &ParagraphBorderLine{
			Val:   string(right.Style),
			Sz:    strconv.Itoa(right.Size),
			Color: right.Color,
			Space: strconv.Itoa(right.Space),
		}
	} else {
		p.Properties.ParagraphBorder.Right = nil
	}

	DebugMsgf(MsgSettingParagraphBorder, top != nil, left != nil, bottom != nil, right != nil)
}

// SetHorizontalRule sets a horizontal rule.
//
// This method is a simplified version of SetBorder, designed for quickly creating Markdown-style horizontal rules.
// It adds a horizontal line only at the bottom of the paragraph, suitable for Markdown --- or *** syntax.
//
// Parameters:
//   - style: border style, e.g. BorderStyleSingle, BorderStyleDouble
//   - size: border width (in 1/8 points; recommended range 12-18)
//   - color: border color (hex, e.g. "000000")
//
// Example:
//
//	// Create a simple horizontal rule
//	para := doc.AddParagraph("")
//	para.SetHorizontalRule(document.BorderStyleSingle, 12, "000000")
//
//	// Create a thick double-line rule
//	para := doc.AddParagraph("")
//	para.SetHorizontalRule(document.BorderStyleDouble, 18, "808080")
func (p *Paragraph) SetHorizontalRule(style BorderStyle, size int, color string) {
	borderConfig := &ParagraphBorderConfig{
		Style: style,
		Size:  size,
		Color: color,
		Space: 1, // Default 1pt spacing
	}

	p.SetBorder(nil, nil, borderConfig, nil)

	DebugMsgf(MsgSettingHorizontalRule, style, size, color)
}

// SetUnderline sets the underline effect for all text in the paragraph.
//
// The underline parameter indicates whether to enable underline.
// When set to true, a single underline is applied to all runs in the paragraph.
// When set to false, underline is removed from all runs.
//
// Example:
//
//	para := doc.AddParagraph("This is underlined text")
//	para.SetUnderline(true)
func (p *Paragraph) SetUnderline(underline bool) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if underline {
			p.Runs[i].Properties.Underline = &Underline{Val: "single"}
		} else {
			p.Runs[i].Properties.Underline = nil
		}
	}
	DebugMsgf(MsgSettingParagraphUnderline, underline)
}

// SetBold sets the bold effect for all text in the paragraph.
//
// The bold parameter indicates whether to enable bold.
// When set to true, bold is applied to all runs in the paragraph.
// When set to false, bold is removed from all runs.
//
// Example:
//
//	para := doc.AddParagraph("This is bold text")
//	para.SetBold(true)
func (p *Paragraph) SetBold(bold bool) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if bold {
			p.Runs[i].Properties.Bold = &Bold{}
			p.Runs[i].Properties.BoldCs = &BoldCs{}
		} else {
			p.Runs[i].Properties.Bold = nil
			p.Runs[i].Properties.BoldCs = nil
		}
	}
	DebugMsgf(MsgSettingParagraphBold, bold)
}

// SetItalic sets the italic effect for all text in the paragraph.
//
// The italic parameter indicates whether to enable italic.
// When set to true, italic is applied to all runs in the paragraph.
// When set to false, italic is removed from all runs.
//
// Example:
//
//	para := doc.AddParagraph("This is italic text")
//	para.SetItalic(true)
func (p *Paragraph) SetItalic(italic bool) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if italic {
			p.Runs[i].Properties.Italic = &Italic{}
			p.Runs[i].Properties.ItalicCs = &ItalicCs{}
		} else {
			p.Runs[i].Properties.Italic = nil
			p.Runs[i].Properties.ItalicCs = nil
		}
	}
	DebugMsgf(MsgSettingParagraphItalic, italic)
}

// SetStrike sets the strikethrough effect for all text in the paragraph.
//
// The strike parameter indicates whether to enable strikethrough.
// When set to true, strikethrough is applied to all runs in the paragraph.
// When set to false, strikethrough is removed from all runs.
//
// Example:
//
//	para := doc.AddParagraph("This is strikethrough text")
//	para.SetStrike(true)
func (p *Paragraph) SetStrike(strike bool) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if strike {
			p.Runs[i].Properties.Strike = &Strike{}
		} else {
			p.Runs[i].Properties.Strike = nil
		}
	}
	DebugMsgf(MsgSettingParagraphStrikethrough, strike)
}

// SetHighlight sets the highlight color for all text in the paragraph.
//
// The color parameter is the highlight color name. Supported colors include:
// "yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue",
// "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow",
// "darkGray", "lightGray", "black", etc.
// Passing an empty string removes the highlight effect.
//
// Example:
//
//	para := doc.AddParagraph("This is highlighted text")
//	para.SetHighlight("yellow")
func (p *Paragraph) SetHighlight(color string) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if color != "" {
			p.Runs[i].Properties.Highlight = &Highlight{Val: color}
		} else {
			p.Runs[i].Properties.Highlight = nil
		}
	}
	DebugMsgf(MsgSettingParagraphHighlight, color)
}

// SetFontFamily sets the font for all text in the paragraph.
//
// The name parameter is the font name, e.g. "Arial", "Times New Roman", etc.
//
// Example:
//
//	para := doc.AddParagraph("This is custom font text")
//	para.SetFontFamily("Arial")
func (p *Paragraph) SetFontFamily(name string) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if name != "" {
			p.Runs[i].Properties.FontFamily = &FontFamily{
				ASCII:    name,
				HAnsi:    name,
				EastAsia: name,
				CS:       name,
			}
		} else {
			p.Runs[i].Properties.FontFamily = nil
		}
	}
	DebugMsgf(MsgSettingParagraphFont, name)
}

// SetFontSize sets the font size for all text in the paragraph.
//
// The size parameter is the font size in points, e.g. 12, 14, 16.
// Passing 0 or a negative number removes the font size setting.
//
// Example:
//
//	para := doc.AddParagraph("This is large text")
//	para.SetFontSize(16)
func (p *Paragraph) SetFontSize(size int) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if size > 0 {
			// Word uses half-points, so multiply by 2
			sizeStr := strconv.Itoa(size * 2)
			p.Runs[i].Properties.FontSize = &FontSize{Val: sizeStr}
			p.Runs[i].Properties.FontSizeCs = &FontSizeCs{Val: sizeStr}
		} else {
			p.Runs[i].Properties.FontSize = nil
			p.Runs[i].Properties.FontSizeCs = nil
		}
	}
	DebugMsgf(MsgSettingParagraphFontSize, size)
}

// SetColor sets the color for all text in the paragraph.
//
// The color parameter is a hex color value, e.g. "FF0000" (red), "0000FF" (blue).
// The "#" prefix is not required; if present, it is automatically removed.
// Passing an empty string removes the color setting.
//
// Example:
//
//	para := doc.AddParagraph("This is red text")
//	para.SetColor("FF0000")
func (p *Paragraph) SetColor(color string) {
	for i := range p.Runs {
		if p.Runs[i].Properties == nil {
			p.Runs[i].Properties = &RunProperties{}
		}
		if color != "" {
			// Remove possible # prefix
			colorVal := strings.TrimPrefix(color, "#")
			p.Runs[i].Properties.Color = &Color{Val: colorVal}
		} else {
			p.Runs[i].Properties.Color = nil
		}
	}
	DebugMsgf(MsgSettingParagraphColor, color)
}

// GetStyleManager returns the document's style manager.
//
// Returns the style manager, which can be used to access and manage styles.
//
// Example:
//
//	doc := document.New()
//	styleManager := doc.GetStyleManager()
//	headingStyle := styleManager.GetStyle("Heading1")
func (d *Document) GetStyleManager() *style.StyleManager {
	return d.styleManager
}

// GetParts returns the document parts map.
//
// Returns a map containing all document parts, primarily used for testing and debugging.
// Keys are part names and values are byte arrays of part contents.
//
// Example:
//
//	parts := doc.GetParts()
//	settingsXML := parts["word/settings.xml"]
func (d *Document) GetParts() map[string][]byte {
	return d.parts
}

// initializeStructure initializes the basic document structure
func (d *Document) initializeStructure() {
	// Initialize content types
	d.contentTypes = &ContentTypes{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
		Defaults: []Default{
			{Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"},
			{Extension: "xml", ContentType: "application/xml"},
		},
		Overrides: []Override{
			{PartName: "/word/document.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"},
			{PartName: "/word/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"},
		},
	}

	// Initialize main relationships
	d.relationships = &Relationships{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: []Relationship{
			{
				ID:     "rId1",
				Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
				Target: "word/document.xml",
			},
		},
	}

	// Add base parts
	d.serializeContentTypes()
	d.serializeRelationships()
	d.serializeDocumentRelationships()
}

// parseDocument parses the document content
func (d *Document) parseDocument() error {
	DebugMsg(MsgParsingDocumentContent)

	// Parse main document
	docData, ok := d.parts["word/document.xml"]
	if !ok {
		return WrapError("parse_document", ErrDocumentNotFound)
	}

	// First parse the basic structure
	decoder := xml.NewDecoder(bytes.NewReader(docData))
	for {
		token, err := decoder.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return WrapError("parse_document", err)
		}

		if t, ok := token.(xml.StartElement); ok {
			if t.Name.Local == "document" && t.Name.Space == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" {
				// Start parsing document
				if err := d.parseDocumentElement(decoder); err != nil {
					return err
				}
				goto done
			}
		}
	}

done:
	InfoMsgf(MsgParsingComplete, len(d.Body.Elements))
	return nil
}

// parseDocumentElement parses document elements
func (d *Document) parseDocumentElement(decoder *xml.Decoder) error {
	// Initialize Body
	d.Body = &Body{
		Elements: make([]interface{}, 0),
	}

	for {
		token, err := decoder.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return WrapError("parse_document_element", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			if t.Name.Local == "body" {
				// Parse document body
				if err := d.parseBodyElement(decoder); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "document" {
				return nil
			}
		}
	}

	return nil
}

// parseBodyElement parses document body elements
func (d *Document) parseBodyElement(decoder *xml.Decoder) error {
	for {
		token, err := decoder.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return WrapError("parse_body_element", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			element, err := d.parseBodySubElement(decoder, t)
			if err != nil {
				return err
			}
			d.Body.Elements = append(d.Body.Elements, element)
		case xml.EndElement:
			if t.Name.Local == "body" {
				return nil
			}
		}
	}

	return nil
}

// parseBodySubElement parses sub-elements of the document body
func (d *Document) parseBodySubElement(decoder *xml.Decoder, startElement xml.StartElement) (interface{}, error) {
	switch startElement.Name.Local {
	case "p":
		// Parse paragraph
		return d.parseParagraph(decoder, startElement)
	case "tbl":
		// Parse table
		return d.parseTable(decoder, startElement)
	case xmlElemSectPr:
		// Parse section properties
		return d.parseSectionProperties(decoder, startElement)
	default:
		// Preserve unknown element as raw XML for round-trip fidelity
		DebugMsgf(MsgPreservingUnknownElement, startElement.Name.Local)
		return d.captureElement(decoder, startElement)
	}
}

// parseParagraph parses a paragraph
//
//nolint:gocognit
func (d *Document) parseParagraph(decoder *xml.Decoder, startElement xml.StartElement) (*Paragraph, error) {
	paragraph := &Paragraph{
		Runs: make([]Run, 0),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_paragraph", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pPr":
				// Parse paragraph properties
				if err := d.parseParagraphProperties(decoder, paragraph); err != nil {
					return nil, err
				}
			case "r":
				// Parse run
				run, err := d.parseRun(decoder, t)
				if err != nil {
					return nil, err
				}
				if run != nil {
					paragraph.Runs = append(paragraph.Runs, *run)
				}
			default:
				// Capture ALL other paragraph-level elements as raw XML for round-trip
				// preservation. This includes: hyperlinks, bookmarks, SDTs, comment
				// ranges, and any other inline elements.
				raw, err := d.captureElement(decoder, t)
				if err != nil {
					return nil, err
				}
				paragraph.RawXMLElements = append(paragraph.RawXMLElements, raw)
			}
		case xml.EndElement:
			if t.Name.Local == "p" {
				return paragraph, nil
			}
		}
	}
}

// parseParagraphProperties parses paragraph properties
//
//nolint:gocognit
func (d *Document) parseParagraphProperties(decoder *xml.Decoder, paragraph *Paragraph) error {
	paragraph.Properties = &ParagraphProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return WrapError("parse_paragraph_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pStyle":
				// Paragraph style
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					paragraph.Properties.ParagraphStyle = &ParagraphStyle{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "spacing":
				// Spacing
				spacing := &Spacing{}
				spacing.Before = getAttributeValue(t.Attr, "before")
				spacing.After = getAttributeValue(t.Attr, "after")
				spacing.Line = getAttributeValue(t.Attr, "line")
				spacing.LineRule = getAttributeValue(t.Attr, "lineRule")
				paragraph.Properties.Spacing = spacing
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "jc":
				// Alignment
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					paragraph.Properties.Justification = &Justification{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "ind":
				// Indentation
				indentation := &Indentation{}
				indentation.FirstLine = getAttributeValue(t.Attr, "firstLine")
				indentation.Left = getAttributeValue(t.Attr, "left")
				indentation.Right = getAttributeValue(t.Attr, "right")
				paragraph.Properties.Indentation = indentation
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "numPr":
				// Numbering properties
				numPr, err := d.parseNumberingProperties(decoder)
				if err != nil {
					return err
				}
				paragraph.Properties.NumberingProperties = numPr
			case xmlElemSectPr:
				// Section properties within paragraph properties define section breaks.
				// Must be preserved on the paragraph, not moved to body level,
				// to maintain correct page breaks (e.g., cover page → TOC).
				sectPr, err := d.parseSectionProperties(decoder, t)
				if err != nil {
					return err
				}
				paragraph.Properties.SectionProperties = sectPr
			case "keepNext":
				val := getAttributeValue(t.Attr, "val")
				paragraph.Properties.KeepNext = &KeepNext{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "keepLines":
				val := getAttributeValue(t.Attr, "val")
				paragraph.Properties.KeepLines = &KeepLines{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "widowControl":
				val := getAttributeValue(t.Attr, "val")
				paragraph.Properties.WidowControl = &WidowControl{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "pageBreakBefore":
				val := getAttributeValue(t.Attr, "val")
				paragraph.Properties.PageBreakBefore = &PageBreakBefore{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "outlineLvl":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					paragraph.Properties.OutlineLevel = &OutlineLevel{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "pPr" {
				return nil
			}
		}
	}
}

// parseNumberingProperties parses numbering properties
//
//nolint:gocognit
func (d *Document) parseNumberingProperties(decoder *xml.Decoder) (*NumberingProperties, error) {
	numPr := &NumberingProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_numbering_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "ilvl":
				// Numbering level
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					numPr.ILevel = &ILevel{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "numId":
				// Numbering ID
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					numPr.NumID = &NumID{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "numPr" {
				return numPr, nil
			}
		}
	}
}

// parseRun parses a text run
//
//nolint:gocognit
func (d *Document) parseRun(decoder *xml.Decoder, startElement xml.StartElement) (*Run, error) {
	run := &Run{
		Text: Text{},
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_run", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "rPr":
				// Parse run properties
				if err := d.parseRunProperties(decoder, run); err != nil {
					return nil, err
				}
			case "t":
				// Parse text
				space := getAttributeValue(t.Attr, "space")
				run.Text.Space = space

				// Read text content
				content, err := d.readElementText(decoder, "t")
				if err != nil {
					return nil, err
				}
				run.Text.Content = content
			case "drawing":
				// Parse drawing element (images, etc.)
				drawing, err := d.parseDrawingElement(decoder, t)
				if err != nil {
					return nil, err
				}
				run.Drawing = drawing
			case "fldChar":
				fldCharType := getAttributeValue(t.Attr, "fldCharType")
				run.FieldChar = &FieldChar{FieldCharType: fldCharType}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "instrText":
				space := getAttributeValue(t.Attr, "space")
				content, err := d.readElementText(decoder, "instrText")
				if err != nil {
					return nil, err
				}
				run.InstrText = &InstrText{Space: space, Content: content}
			case "footnoteReference":
				id := getAttributeValue(t.Attr, "id")
				run.FootnoteReference = &FootnoteReference{ID: id}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "endnoteReference":
				id := getAttributeValue(t.Attr, "id")
				run.EndnoteReference = &EndnoteReference{ID: id}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "footnoteRef":
				run.FootnoteRef = &FootnoteRef{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "endnoteRef":
				run.EndnoteRef = &EndnoteRef{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "separator":
				run.Separator = &Separator{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "continuationSeparator":
				run.ContinuationSeparator = &ContinuationSeparator{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "br":
				brType := getAttributeValue(t.Attr, "type")
				run.Break = &Break{Type: brType}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				// Capture unknown run elements as raw XML (tab, lastRenderedPageBreak,
				// commentReference, noBreakHyphen, etc.)
				raw, err := d.captureElement(decoder, t)
				if err != nil {
					return nil, err
				}
				run.RawXMLContent = append(run.RawXMLContent, raw)
			}
		case xml.EndElement:
			if t.Name.Local == "r" {
				return run, nil
			}
		}
	}
}

// parseRunProperties parses run properties
//
//nolint:gocognit
func (d *Document) parseRunProperties(decoder *xml.Decoder, run *Run) error {
	run.Properties = &RunProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return WrapError("parse_run_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "b":
				run.Properties.Bold = &Bold{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "bCs":
				run.Properties.BoldCs = &BoldCs{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "i":
				run.Properties.Italic = &Italic{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "iCs":
				run.Properties.ItalicCs = &ItalicCs{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "u":
				val := getAttributeValue(t.Attr, "val")
				run.Properties.Underline = &Underline{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "strike":
				run.Properties.Strike = &Strike{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "sz":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					run.Properties.FontSize = &FontSize{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "szCs":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					run.Properties.FontSizeCs = &FontSizeCs{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "color":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					run.Properties.Color = &Color{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "highlight":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					run.Properties.Highlight = &Highlight{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "rFonts":
				ascii := getAttributeValue(t.Attr, "ascii")
				hAnsi := getAttributeValue(t.Attr, "hAnsi")
				eastAsia := getAttributeValue(t.Attr, "eastAsia")
				cs := getAttributeValue(t.Attr, "cs")
				hint := getAttributeValue(t.Attr, "hint")

				run.Properties.FontFamily = &FontFamily{
					ASCII:    ascii,
					HAnsi:    hAnsi,
					EastAsia: eastAsia,
					CS:       cs,
					Hint:     hint,
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "rStyle":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					run.Properties.RunStyle = &RunStyle{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "vertAlign":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					run.Properties.VerticalAlign = &VerticalAlignment{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "rPr" {
				return nil
			}
		}
	}
}

// parseTable parses a table
//
//nolint:gocognit
func (d *Document) parseTable(decoder *xml.Decoder, startElement xml.StartElement) (*Table, error) {
	table := &Table{
		Rows: make([]TableRow, 0),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tblPr":
				// Parse table properties
				if err := d.parseTableProperties(decoder, table); err != nil {
					return nil, err
				}
			case "tblGrid":
				// Parse table grid
				if err := d.parseTableGrid(decoder, table); err != nil {
					return nil, err
				}
			case "tr":
				// Parse table row
				row, err := d.parseTableRow(decoder, t)
				if err != nil {
					return nil, err
				}
				if row != nil {
					table.Rows = append(table.Rows, *row)
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "tbl" {
				return table, nil
			}
		}
	}
}

// parseTableProperties parses table properties
//
//nolint:gocognit
func (d *Document) parseTableProperties(decoder *xml.Decoder, table *Table) error {
	table.Properties = &TableProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return WrapError("parse_table_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tblW":
				w := getAttributeValue(t.Attr, "w")
				wType := getAttributeValue(t.Attr, "type")
				if w != "" || wType != "" {
					table.Properties.TableW = &TableWidth{W: w, Type: wType}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "jc":
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					table.Properties.TableJc = &TableJc{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "tblLook":
				// Parse table look
				tableLook := &TableLook{
					Val:      getAttributeValue(t.Attr, "val"),
					FirstRow: getAttributeValue(t.Attr, "firstRow"),
					LastRow:  getAttributeValue(t.Attr, "lastRow"),
					FirstCol: getAttributeValue(t.Attr, "firstColumn"),
					LastCol:  getAttributeValue(t.Attr, "lastColumn"),
					NoHBand:  getAttributeValue(t.Attr, "noHBand"),
					NoVBand:  getAttributeValue(t.Attr, "noVBand"),
				}
				table.Properties.TableLook = tableLook
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "tblStyle":
				// Parse table style
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					table.Properties.TableStyle = &TableStyle{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "tblBorders":
				// Parse table borders
				borders, err := d.parseTableBorders(decoder)
				if err != nil {
					return err
				}
				table.Properties.TableBorders = borders
			case "shd":
				// Parse table shading
				shd := &TableShading{
					Val:       getAttributeValue(t.Attr, "val"),
					Color:     getAttributeValue(t.Attr, "color"),
					Fill:      getAttributeValue(t.Attr, "fill"),
					ThemeFill: getAttributeValue(t.Attr, "themeFill"),
				}
				table.Properties.Shd = shd
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "tblCellMar":
				// Parse table cell margins
				margins, err := d.parseTableCellMargins(decoder)
				if err != nil {
					return err
				}
				table.Properties.TableCellMar = margins
			case "tblLayout":
				// Parse table layout
				layoutType := getAttributeValue(t.Attr, "type")
				if layoutType != "" {
					table.Properties.TableLayout = &TableLayoutType{Type: layoutType}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			case "tblInd":
				// Parse table indentation
				w := getAttributeValue(t.Attr, "w")
				indType := getAttributeValue(t.Attr, "type")
				if w != "" || indType != "" {
					table.Properties.TableInd = &TableIndentation{W: w, Type: indType}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			default:
				// Skip other table properties; extend as needed
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "tblPr" {
				return nil
			}
		}
	}
}

// parseTableGrid parses the table grid
func (d *Document) parseTableGrid(decoder *xml.Decoder, table *Table) error {
	table.Grid = &TableGrid{
		Cols: make([]TableGridCol, 0),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return WrapError("parse_table_grid", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "gridCol":
				w := getAttributeValue(t.Attr, "w")
				col := TableGridCol{W: w}
				table.Grid.Cols = append(table.Grid.Cols, col)
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "tblGrid" {
				return nil
			}
		}
	}
}

// parseXMLChildren is a generic helper that iterates over XML child elements, dispatching
// each start element to the provided handler. It returns when it encounters an end element
// matching endTag. The errContext string is used for wrapping errors.
func (d *Document) parseXMLChildren(decoder *xml.Decoder, endTag string, errContext string, handler func(xml.StartElement) error) error {
	for {
		token, err := decoder.Token()
		if err != nil {
			return WrapError(errContext, err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			if err := handler(t); err != nil {
				return err
			}
		case xml.EndElement:
			if t.Name.Local == endTag {
				return nil
			}
		}
	}
}

// parseTableRow parses a table row
//
//nolint:dupl
func (d *Document) parseTableRow(decoder *xml.Decoder, startElement xml.StartElement) (*TableRow, error) {
	row := &TableRow{
		Cells: make([]TableCell, 0),
	}

	err := d.parseXMLChildren(decoder, "tr", "parse_table_row", func(t xml.StartElement) error {
		switch t.Name.Local {
		case "trPr":
			props, err := d.parseTableRowProperties(decoder)
			if err != nil {
				return err
			}
			row.Properties = props
		case "tc":
			cell, err := d.parseTableCell(decoder, t)
			if err != nil {
				return err
			}
			if cell != nil {
				row.Cells = append(row.Cells, *cell)
			}
		default:
			if err := d.skipElement(decoder, t.Name.Local); err != nil {
				return err
			}
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	return row, nil
}

// parseTableCell parses a table cell
//
//nolint:dupl
func (d *Document) parseTableCell(decoder *xml.Decoder, startElement xml.StartElement) (*TableCell, error) {
	cell := &TableCell{
		Paragraphs: make([]Paragraph, 0),
	}

	err := d.parseXMLChildren(decoder, "tc", "parse_table_cell", func(t xml.StartElement) error {
		switch t.Name.Local {
		case "tcPr":
			props, err := d.parseTableCellProperties(decoder)
			if err != nil {
				return err
			}
			cell.Properties = props
		case "p":
			para, err := d.parseParagraph(decoder, t)
			if err != nil {
				return err
			}
			if para != nil {
				cell.Paragraphs = append(cell.Paragraphs, *para)
			}
		default:
			if err := d.skipElement(decoder, t.Name.Local); err != nil {
				return err
			}
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	return cell, nil
}

// parseSectionProperties parses section properties
//
//nolint:gocognit
func (d *Document) parseSectionProperties(decoder *xml.Decoder, startElement xml.StartElement) (*SectionProperties, error) {
	sectPr := &SectionProperties{
		XmlnsR: getAttributeValue(startElement.Attr, "xmlns:r"),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_section_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pgSz":
				// Parse page size
				w := getAttributeValue(t.Attr, "w")
				h := getAttributeValue(t.Attr, "h")
				orient := getAttributeValue(t.Attr, "orient")
				if w != "" || h != "" {
					sectPr.PageSize = &PageSizeXML{W: w, H: h, Orient: orient}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "pgMar":
				// Parse page margins
				margin := &PageMargin{}
				margin.Top = getAttributeValue(t.Attr, "top")
				margin.Right = getAttributeValue(t.Attr, "right")
				margin.Bottom = getAttributeValue(t.Attr, "bottom")
				margin.Left = getAttributeValue(t.Attr, "left")
				margin.Header = getAttributeValue(t.Attr, "header")
				margin.Footer = getAttributeValue(t.Attr, "footer")
				margin.Gutter = getAttributeValue(t.Attr, "gutter")
				sectPr.PageMargins = margin
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "cols":
				// Parse columns
				space := getAttributeValue(t.Attr, "space")
				num := getAttributeValue(t.Attr, "num")
				if space != "" || num != "" {
					sectPr.Columns = &Columns{Space: space, Num: num}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "docGrid":
				// Parse document grid
				docGridType := getAttributeValue(t.Attr, "type")
				linePitch := getAttributeValue(t.Attr, "linePitch")
				charSpace := getAttributeValue(t.Attr, "charSpace")
				if docGridType != "" || linePitch != "" || charSpace != "" {
					sectPr.DocGrid = &DocGrid{
						Type:      docGridType,
						LinePitch: linePitch,
						CharSpace: charSpace,
					}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "headerReference":
				ref := &HeaderFooterReference{
					Type: getAttributeValue(t.Attr, "type"),
					ID:   getAttributeValue(t.Attr, "id"),
				}
				if ref.Type == "" {
					ref.Type = getAttributeValue(t.Attr, "w:type")
				}
				if ref.ID == "" {
					ref.ID = getAttributeValue(t.Attr, "r:id")
				}
				if ref.ID != "" || ref.Type != "" {
					sectPr.HeaderReferences = append(sectPr.HeaderReferences, ref)
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "footerReference":
				ref := &FooterReference{
					Type: getAttributeValue(t.Attr, "type"),
					ID:   getAttributeValue(t.Attr, "id"),
				}
				if ref.Type == "" {
					ref.Type = getAttributeValue(t.Attr, "w:type")
				}
				if ref.ID == "" {
					ref.ID = getAttributeValue(t.Attr, "r:id")
				}
				if ref.ID != "" || ref.Type != "" {
					sectPr.FooterReferences = append(sectPr.FooterReferences, ref)
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "type":
				// Section type (continuous, nextPage, etc.)
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					sectPr.SectionType = &SectionType{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "titlePg":
				// Title page setting
				sectPr.TitlePage = &TitlePage{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "pgNumType":
				// Page number type
				pgNum := &PageNumType{}
				pgNum.Fmt = getAttributeValue(t.Attr, "fmt")
				pgNum.Start = getAttributeValue(t.Attr, "start")
				sectPr.PageNumType = pgNum
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				// Skip other section properties
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == xmlElemSectPr {
				return sectPr, nil
			}
		}
	}
}

// skipElement skips an element and its child elements
func (d *Document) skipElement(decoder *xml.Decoder, elementName string) error {
	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			return WrapError("skip_element", err)
		}

		switch token.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return nil
}

// captureElement captures an unknown XML element and all its children as a RawXMLElement.
// The decoder should have already consumed the start element token.
func (d *Document) captureElement(decoder *xml.Decoder, startElement xml.StartElement) (*RawXMLElement, error) {
	var buf bytes.Buffer
	enc := xml.NewEncoder(&buf)

	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("capture_element", err)
		}

		switch token.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}

		// Encode all tokens except the final end element
		if depth > 0 {
			if err := enc.EncodeToken(xml.CopyToken(token)); err != nil {
				return nil, WrapError("capture_element", err)
			}
		}
	}

	if err := enc.Flush(); err != nil {
		return nil, WrapError("capture_element", err)
	}

	return &RawXMLElement{
		XMLName:  startElement.Name,
		Attrs:    startElement.Attr,
		InnerXML: buf.String(),
	}, nil
}

// readElementText reads the text content of an element
func (d *Document) readElementText(decoder *xml.Decoder, elementName string) (string, error) {
	var content string
	for {
		token, err := decoder.Token()
		if err != nil {
			return "", WrapError("read_element_text", err)
		}

		switch t := token.(type) {
		case xml.CharData:
			content += string(t)
		case xml.EndElement:
			if t.Name.Local == elementName {
				return content, nil
			}
		}
	}
}

// getAttributeValue gets an attribute value
func getAttributeValue(attrs []xml.Attr, name string) string {
	for _, attr := range attrs {
		if attr.Name.Local == name {
			return attr.Value
		}
	}
	return ""
}

// serializeDocument serializes the document content
func (d *Document) serializeDocument() error {
	DebugMsg(MsgSerializingDocument)

	// Create document structure
	type documentXML struct {
		XMLName  xml.Name `xml:"w:document"`
		Xmlns    string   `xml:"xmlns:w,attr"`
		XmlnsW15 string   `xml:"xmlns:w15,attr"`
		XmlnsWP  string   `xml:"xmlns:wp,attr"`
		XmlnsA   string   `xml:"xmlns:a,attr"`
		XmlnsPic string   `xml:"xmlns:pic,attr"`
		XmlnsR   string   `xml:"xmlns:r,attr"`
		Body     *Body    `xml:"w:body"`
	}

	doc := documentXML{
		Xmlns:    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsW15: "http://schemas.microsoft.com/office/word/2012/wordml",
		XmlnsWP:  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
		XmlnsA:   "http://schemas.openxmlformats.org/drawingml/2006/main",
		XmlnsPic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
		XmlnsR:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		Body:     d.Body,
	}

	// Serialize to XML
	data, err := xml.MarshalIndent(doc, "", "  ")
	if err != nil {
		ErrorMsgf(MsgXMLSerializationFailed, err)
		return WrapError("marshal_xml", err)
	}

	// Add XML declaration
	d.parts["word/document.xml"] = append([]byte(xml.Header), data...)

	DebugMsg(MsgDocumentSerializationComplete)
	return nil
}

// serializeContentTypes serializes content types
func (d *Document) serializeContentTypes() {
	data, _ := xml.MarshalIndent(d.contentTypes, "", "  ")
	d.parts["[Content_Types].xml"] = append([]byte(xml.Header), data...)
}

// serializeRelationships serializes relationships
func (d *Document) serializeRelationships() {
	data, _ := xml.MarshalIndent(d.relationships, "", "  ")
	d.parts["_rels/.rels"] = append([]byte(xml.Header), data...)
}

// serializeDocumentRelationships serializes document relationships
func (d *Document) serializeDocumentRelationships() {
	// Get existing relationships, starting from index 1 (reserved for styles.xml)
	relationships := []Relationship{
		{
			ID:     "rId1",
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
			Target: "styles.xml",
		},
	}

	// Add dynamically created document-level relationships (headers, footers, etc.)
	relationships = append(relationships, d.documentRelationships.Relationships...)

	// Create document relationships
	docRels := &Relationships{
		Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: relationships,
	}

	data, _ := xml.MarshalIndent(docRels, "", "  ")
	d.parts["word/_rels/document.xml.rels"] = append([]byte(xml.Header), data...)
}

// serializeStyles serializes styles
func (d *Document) serializeStyles() error {
	DebugMsg(MsgSerializingStyles)

	// If a complete styles.xml was preserved during document cloning (including docDefaults, etc.),
	// skip regeneration to avoid losing the template's original default paragraph/character settings.
	if existing, ok := d.parts["word/styles.xml"]; ok && len(existing) > 0 {
		DebugMsg(MsgExistingStylesDetected)
		return nil
	}

	// Create styles structure with full namespaces
	type stylesXML struct {
		XMLName     xml.Name       `xml:"w:styles"`
		XmlnsW      string         `xml:"xmlns:w,attr"`
		XmlnsMC     string         `xml:"xmlns:mc,attr"`
		XmlnsO      string         `xml:"xmlns:o,attr"`
		XmlnsR      string         `xml:"xmlns:r,attr"`
		XmlnsM      string         `xml:"xmlns:m,attr"`
		XmlnsV      string         `xml:"xmlns:v,attr"`
		XmlnsW14    string         `xml:"xmlns:w14,attr"`
		XmlnsW10    string         `xml:"xmlns:w10,attr"`
		XmlnsSL     string         `xml:"xmlns:sl,attr"`
		XmlnsWPS    string         `xml:"xmlns:wpsCustomData,attr"`
		MCIgnorable string         `xml:"mc:Ignorable,attr"`
		Styles      []*style.Style `xml:"w:style"`
	}

	doc := stylesXML{
		XmlnsW:      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsMC:     "http://schemas.openxmlformats.org/markup-compatibility/2006",
		XmlnsO:      "urn:schemas-microsoft-com:office:office",
		XmlnsR:      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsM:      "http://schemas.openxmlformats.org/officeDocument/2006/math",
		XmlnsV:      "urn:schemas-microsoft-com:vml",
		XmlnsW14:    "http://schemas.microsoft.com/office/word/2010/wordml",
		XmlnsW10:    "urn:schemas-microsoft-com:office:word",
		XmlnsSL:     "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
		XmlnsWPS:    "http://www.wps.cn/officeDocument/2013/wpsCustomData",
		MCIgnorable: "w14",
		Styles:      d.styleManager.GetAllStyles(),
	}

	// Serialize to XML
	data, err := xml.MarshalIndent(doc, "", "  ")
	if err != nil {
		ErrorMsgf(MsgXMLSerializationFailed, err)
		return WrapError("marshal_xml", err)
	}

	// Add XML declaration
	d.parts["word/styles.xml"] = append([]byte(xml.Header), data...)

	DebugMsg(MsgStyleSerializationComplete)
	return nil
}

// parseContentTypes parses the content types file
func (d *Document) parseContentTypes() error {
	DebugMsg(MsgParsingContentTypes)

	// Find content types file
	contentTypesData, ok := d.parts["[Content_Types].xml"]
	if !ok {
		return WrapError("parse_content_types", fmt.Errorf("content types file not found"))
	}

	// Parse XML
	var contentTypes ContentTypes
	if err := xml.Unmarshal(contentTypesData, &contentTypes); err != nil {
		return WrapError("parse_content_types", err)
	}

	d.contentTypes = &contentTypes
	DebugMsg(MsgContentTypesParsed)
	return nil
}

// parseRelationships parses the relationships file
func (d *Document) parseRelationships() error {
	DebugMsg(MsgParsingRelationships)

	// Find relationships file
	relsData, ok := d.parts["_rels/.rels"]
	if !ok {
		return WrapError("parse_relationships", fmt.Errorf("relationships file not found"))
	}

	// Parse XML
	var relationships Relationships
	if err := xml.Unmarshal(relsData, &relationships); err != nil {
		return WrapError("parse_relationships", err)
	}

	d.relationships = &relationships
	DebugMsg(MsgRelationshipsParsed)
	return nil
}

// parseStyles parses the styles file
func (d *Document) parseStyles() error {
	DebugMsg(MsgParsingStyles)

	// Find styles file
	stylesData, ok := d.parts["word/styles.xml"]
	if !ok {
		return WrapError("parse_styles", fmt.Errorf("styles file not found"))
	}

	// Parse styles using the style manager
	if err := d.styleManager.LoadStylesFromDocument(stylesData); err != nil {
		return WrapError("parse_styles", err)
	}

	DebugMsg(MsgStylesParsed)
	return nil
}

// parseDocumentRelationships parses the document relationships file (word/_rels/document.xml.rels).
// This file contains relationships for images, headers, footers, and other resources in the document.
func (d *Document) parseDocumentRelationships() error {
	DebugMsg(MsgParsingDocumentRelationships)

	// Find document relationships file
	docRelsData, ok := d.parts["word/_rels/document.xml.rels"]
	if !ok {
		// The document may not have a relationships file (no images or other resources); this is not an error
		DebugMsg(MsgDocRelFileNotFound)
		return nil
	}

	// Parse XML
	var relationships Relationships
	if err := xml.Unmarshal(docRelsData, &relationships); err != nil {
		return WrapError("parse_document_relationships", err)
	}

	// Save parsed relationships (excluding styles.xml, which is auto-added in serializeDocumentRelationships)
	// Filter out styles.xml relationship since it is always rId1 and auto-added on save
	filteredRels := make([]Relationship, 0)
	for _, rel := range relationships.Relationships {
		if rel.Type != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" {
			filteredRels = append(filteredRels, rel)
		}
	}

	d.documentRelationships.Relationships = filteredRels
	DebugMsgf(MsgDocumentRelationshipsParsed, len(filteredRels))
	return nil
}

// updateNextImageID updates the nextImageID counter based on existing image relationships.
// This ensures newly added image IDs do not conflict with existing images.
func (d *Document) updateNextImageID() {
	maxImageID := -1

	// Iterate through all parts to find the highest existing image file ID
	for partName := range d.parts {
		// Check if this is an image file (word/media/imageN.xxx)
		if len(partName) > 11 && partName[:11] == "word/media/" {
			// Extract image ID from filename (image0.png -> 0, image1.png -> 1, etc.)
			filename := partName[11:] // Remove "word/media/" prefix
			var id int
			if _, err := fmt.Sscanf(filename, "image%d.", &id); err == nil {
				if id > maxImageID {
					maxImageID = id
				}
			}
		}
	}

	// Set nextImageID to max image ID + 1
	// If there are no existing images, maxImageID is -1, so nextImageID should be 0
	d.nextImageID = maxImageID + 1

	DebugMsgf(MsgUpdatingImageIDCounter, d.nextImageID)
}

// ToBytes converts the document to a byte array
func (d *Document) ToBytes() ([]byte, error) {
	buf := new(bytes.Buffer)
	zipWriter := zip.NewWriter(buf)

	// Serialize document
	if err := d.serializeDocument(); err != nil {
		return nil, err
	}

	// Serialize styles
	if err := d.serializeStyles(); err != nil {
		return nil, err
	}

	// Serialize content types
	d.serializeContentTypes()

	// Serialize relationships
	d.serializeRelationships()

	// Serialize document relationships
	d.serializeDocumentRelationships()

	// Write all parts
	for name, data := range d.parts {
		writer, err := zipWriter.Create(name)
		if err != nil {
			return nil, err
		}
		if _, err := writer.Write(data); err != nil {
			return nil, err
		}
	}

	if err := zipWriter.Close(); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

// GetParagraphs returns all paragraphs
func (b *Body) GetParagraphs() []*Paragraph {
	paragraphs := make([]*Paragraph, 0)
	for _, element := range b.Elements {
		if p, ok := element.(*Paragraph); ok {
			paragraphs = append(paragraphs, p)
		}
	}
	return paragraphs
}

// GetTables returns all tables
func (b *Body) GetTables() []*Table {
	tables := make([]*Table, 0)
	for _, element := range b.Elements {
		if t, ok := element.(*Table); ok {
			tables = append(tables, t)
		}
	}
	return tables
}

// AddElement adds an element to the document body
func (b *Body) AddElement(element interface{}) {
	b.Elements = append(b.Elements, element)
}

// RemoveParagraph removes the specified paragraph from the document.
//
// The paragraph parameter is the paragraph object to remove.
// If the paragraph does not exist in the document, this method has no effect.
//
// Returns whether the paragraph was successfully removed.
//
// Example:
//
//	doc := document.New()
//	para := doc.AddParagraph("Paragraph to remove")
//	doc.RemoveParagraph(para)
func (d *Document) RemoveParagraph(paragraph *Paragraph) bool {
	for i, element := range d.Body.Elements {
		if p, ok := element.(*Paragraph); ok && p == paragraph {
			// Remove element
			d.Body.Elements = append(d.Body.Elements[:i], d.Body.Elements[i+1:]...)
			DebugMsgf(MsgDeletingElement, i)
			return true
		}
	}
	DebugMsg(MsgParagraphToDeleteNotFound)
	return false
}

// RemoveParagraphAt removes a paragraph by index.
//
// The index parameter is the paragraph's index among all paragraphs (0-based).
// If the index is out of range, this method returns false.
//
// Returns whether the paragraph was successfully removed.
//
// Example:
//
//	doc := document.New()
//	doc.AddParagraph("First paragraph")
//	doc.AddParagraph("Second paragraph")
//	doc.RemoveParagraphAt(0)  // Remove the first paragraph
func (d *Document) RemoveParagraphAt(index int) bool {
	// Validate negative index early
	if index < 0 {
		DebugMsgf(MsgParagraphIndexNegative, index)
		return false
	}

	// Optimization: single pass to find the target paragraph and its element index
	paragraphCount := 0
	for i, element := range d.Body.Elements {
		if _, ok := element.(*Paragraph); ok {
			if paragraphCount == index {
				// Found the target paragraph; remove it
				d.Body.Elements = append(d.Body.Elements[:i], d.Body.Elements[i+1:]...)
				DebugMsgf(MsgDeletingElement, i)
				return true
			}
			paragraphCount++
		}
	}

	DebugMsgf(MsgParagraphIndexOutOfRange, index, paragraphCount)
	return false
}

// RemoveElementAt removes an element by index (including paragraphs, tables, etc.).
//
// The index parameter is the element's index in the document body (0-based).
// If the index is out of range, this method returns false.
//
// Returns whether the element was successfully removed.
//
// Example:
//
//	doc := document.New()
//	doc.AddParagraph("A paragraph")
//	doc.AddTable(&document.TableConfig{Rows: 2, Cols: 2})
//	doc.RemoveElementAt(0)  // Remove the first element (paragraph)
func (d *Document) RemoveElementAt(index int) bool {
	if index < 0 || index >= len(d.Body.Elements) {
		DebugMsgf(MsgElementIndexOutOfRange, index, len(d.Body.Elements))
		return false
	}

	// Remove element
	d.Body.Elements = append(d.Body.Elements[:index], d.Body.Elements[index+1:]...)
	DebugMsgf(MsgDeletingElement, index)
	return true
}

// parseTableBorders parses table borders
func (d *Document) parseTableBorders(decoder *xml.Decoder) (*TableBorders, error) {
	borders := &TableBorders{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_borders", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			border := &TableBorder{
				Val:        getAttributeValue(t.Attr, "val"),
				Sz:         getAttributeValue(t.Attr, "sz"),
				Space:      getAttributeValue(t.Attr, "space"),
				Color:      getAttributeValue(t.Attr, "color"),
				ThemeColor: getAttributeValue(t.Attr, "themeColor"),
			}

			switch t.Name.Local {
			case string(CellVAlignTop):
				borders.Top = border
			case string(CellAlignLeft):
				borders.Left = border
			case string(CellVAlignBottom):
				borders.Bottom = border
			case string(CellAlignRight):
				borders.Right = border
			case "insideH":
				borders.InsideH = border
			case "insideV":
				borders.InsideV = border
			}

			if err := d.skipElement(decoder, t.Name.Local); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == "tblBorders" {
				return borders, nil
			}
		}
	}
}

// parseTableCellMargins parses table cell margins
//
//nolint:dupl
func (d *Document) parseTableCellMargins(decoder *xml.Decoder) (*TableCellMargins, error) {
	margins := &TableCellMargins{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_cell_margins", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			space := &TableCellSpace{
				W:    getAttributeValue(t.Attr, "w"),
				Type: getAttributeValue(t.Attr, "type"),
			}

			switch t.Name.Local {
			case string(CellVAlignTop):
				margins.Top = space
			case string(CellAlignLeft):
				margins.Left = space
			case string(CellVAlignBottom):
				margins.Bottom = space
			case string(CellAlignRight):
				margins.Right = space
			}

			if err := d.skipElement(decoder, t.Name.Local); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == "tblCellMar" {
				return margins, nil
			}
		}
	}
}

// parseTableCellProperties parses table cell properties
//
//nolint:gocognit
func (d *Document) parseTableCellProperties(decoder *xml.Decoder) (*TableCellProperties, error) {
	props := &TableCellProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_cell_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tcW":
				// Parse cell width
				w := getAttributeValue(t.Attr, "w")
				wType := getAttributeValue(t.Attr, "type")
				if w != "" || wType != "" {
					props.TableCellW = &TableCellW{W: w, Type: wType}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "vAlign":
				// Parse vertical alignment
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					props.VAlign = &VAlign{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "gridSpan":
				// Parse grid span
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					props.GridSpan = &GridSpan{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "vMerge":
				// Parse vertical merge
				val := getAttributeValue(t.Attr, "val")
				props.VMerge = &VMerge{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "textDirection":
				// Parse text direction
				val := getAttributeValue(t.Attr, "val")
				if val != "" {
					props.TextDirection = &TextDirection{Val: val}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "shd":
				// Parse cell shading
				shd := &TableCellShading{
					Val:       getAttributeValue(t.Attr, "val"),
					Color:     getAttributeValue(t.Attr, "color"),
					Fill:      getAttributeValue(t.Attr, "fill"),
					ThemeFill: getAttributeValue(t.Attr, "themeFill"),
				}
				props.Shd = shd
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "tcBorders":
				// Parse cell borders
				borders, err := d.parseTableCellBorders(decoder)
				if err != nil {
					return nil, err
				}
				props.TcBorders = borders
			case "tcMar":
				// Parse cell margins
				margins, err := d.parseTableCellMarginsCell(decoder)
				if err != nil {
					return nil, err
				}
				props.TcMar = margins
			case "noWrap":
				// Parse no-wrap
				val := getAttributeValue(t.Attr, "val")
				props.NoWrap = &NoWrap{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "hideMark":
				// Parse hide mark
				val := getAttributeValue(t.Attr, "val")
				props.HideMark = &HideMark{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				// Skip other unhandled cell properties
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "tcPr" {
				return props, nil
			}
		}
	}
}

// parseTableCellBorders parses table cell borders
func (d *Document) parseTableCellBorders(decoder *xml.Decoder) (*TableCellBorders, error) {
	borders := &TableCellBorders{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_cell_borders", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			border := &TableCellBorder{
				Val:        getAttributeValue(t.Attr, "val"),
				Sz:         getAttributeValue(t.Attr, "sz"),
				Space:      getAttributeValue(t.Attr, "space"),
				Color:      getAttributeValue(t.Attr, "color"),
				ThemeColor: getAttributeValue(t.Attr, "themeColor"),
			}

			switch t.Name.Local {
			case string(CellVAlignTop):
				borders.Top = border
			case string(CellAlignLeft):
				borders.Left = border
			case string(CellVAlignBottom):
				borders.Bottom = border
			case string(CellAlignRight):
				borders.Right = border
			case "insideH":
				borders.InsideH = border
			case "insideV":
				borders.InsideV = border
			case "tl2br":
				borders.TL2BR = border
			case "tr2bl":
				borders.TR2BL = border
			}

			if err := d.skipElement(decoder, t.Name.Local); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == "tcBorders" {
				return borders, nil
			}
		}
	}
}

// parseTableCellMarginsCell parses table cell margins (cell level)
//
//nolint:dupl
func (d *Document) parseTableCellMarginsCell(decoder *xml.Decoder) (*TableCellMarginsCell, error) {
	margins := &TableCellMarginsCell{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_cell_margins_cell", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			space := &TableCellSpaceCell{
				W:    getAttributeValue(t.Attr, "w"),
				Type: getAttributeValue(t.Attr, "type"),
			}

			switch t.Name.Local {
			case string(CellVAlignTop):
				margins.Top = space
			case string(CellAlignLeft):
				margins.Left = space
			case string(CellVAlignBottom):
				margins.Bottom = space
			case string(CellAlignRight):
				margins.Right = space
			}

			if err := d.skipElement(decoder, t.Name.Local); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if t.Name.Local == "tcMar" {
				return margins, nil
			}
		}
	}
}

// parseTableRowProperties parses table row properties
//
//nolint:gocognit
func (d *Document) parseTableRowProperties(decoder *xml.Decoder) (*TableRowProperties, error) {
	props := &TableRowProperties{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_table_row_properties", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "trHeight":
				// Parse row height
				val := getAttributeValue(t.Attr, "val")
				hRule := getAttributeValue(t.Attr, "hRule")
				if val != "" || hRule != "" {
					props.TableRowH = &TableRowH{Val: val, HRule: hRule}
				}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "cantSplit":
				// Parse prevent row from splitting across pages
				val := getAttributeValue(t.Attr, "val")
				props.CantSplit = &CantSplit{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "tblHeader":
				// Parse repeat header row
				val := getAttributeValue(t.Attr, "val")
				props.TblHeader = &TblHeader{Val: val}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				// Skip other row properties
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "trPr" {
				return props, nil
			}
		}
	}
}

// parseDrawingElement parses drawing elements (images, etc.).
// This method parses the complete drawing element structure from XML.
func (d *Document) parseDrawingElement(decoder *xml.Decoder, startElement xml.StartElement) (*DrawingElement, error) {
	drawing := &DrawingElement{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_drawing_element", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "inline":
				// Parse inline drawing
				inline, err := d.parseInlineDrawing(decoder, t)
				if err != nil {
					return nil, err
				}
				drawing.Inline = inline
			case "anchor":
				// Parse anchor (floating) drawing
				anchor, err := d.parseAnchorDrawing(decoder, t)
				if err != nil {
					return nil, err
				}
				drawing.Anchor = anchor
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "drawing" {
				return drawing, nil
			}
		}
	}
}

// parseInlineDrawing parses an inline drawing
//
//nolint:gocognit
func (d *Document) parseInlineDrawing(decoder *xml.Decoder, startElement xml.StartElement) (*InlineDrawing, error) {
	inline := &InlineDrawing{}

	// Parse attributes
	for _, attr := range startElement.Attr {
		switch attr.Name.Local {
		case xmlAttrDistT:
			inline.DistT = attr.Value
		case xmlAttrDistB:
			inline.DistB = attr.Value
		case xmlAttrDistL:
			inline.DistL = attr.Value
		case xmlAttrDistR:
			inline.DistR = attr.Value
		}
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_inline_drawing", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "extent":
				extent := &DrawingExtent{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "cx":
						extent.Cx = attr.Value
					case "cy":
						extent.Cy = attr.Value
					}
				}
				inline.Extent = extent
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "docPr":
				docPr := &DrawingDocPr{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "id":
						docPr.ID = attr.Value
					case xmlAttrName:
						docPr.Name = attr.Value
					case xmlAttrDescr:
						docPr.Descr = attr.Value
					case xmlAttrTitle:
						docPr.Title = attr.Value
					}
				}
				inline.DocPr = docPr
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case xmlElemGraphic:
				graphic, err := d.parseDrawingGraphic(decoder, t)
				if err != nil {
					return nil, err
				}
				inline.Graphic = graphic
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "inline" {
				return inline, nil
			}
		}
	}
}

// parseAnchorDrawing parses an anchor (floating) drawing
//
//nolint:gocognit
func (d *Document) parseAnchorDrawing(decoder *xml.Decoder, startElement xml.StartElement) (*AnchorDrawing, error) {
	anchor := &AnchorDrawing{}

	// Parse attributes
	for _, attr := range startElement.Attr {
		switch attr.Name.Local {
		case xmlAttrDistT:
			anchor.DistT = attr.Value
		case xmlAttrDistB:
			anchor.DistB = attr.Value
		case xmlAttrDistL:
			anchor.DistL = attr.Value
		case xmlAttrDistR:
			anchor.DistR = attr.Value
		case "simplePos":
			anchor.SimplePos = attr.Value
		case "relativeHeight":
			anchor.RelativeHeight = attr.Value
		case "behindDoc":
			anchor.BehindDoc = attr.Value
		case "locked":
			anchor.Locked = attr.Value
		case "layoutInCell":
			anchor.LayoutInCell = attr.Value
		case "allowOverlap":
			anchor.AllowOverlap = attr.Value
		}
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_anchor_drawing", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "extent":
				extent := &DrawingExtent{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "cx":
						extent.Cx = attr.Value
					case "cy":
						extent.Cy = attr.Value
					}
				}
				anchor.Extent = extent
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "docPr":
				docPr := &DrawingDocPr{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "id":
						docPr.ID = attr.Value
					case xmlAttrName:
						docPr.Name = attr.Value
					case xmlAttrDescr:
						docPr.Descr = attr.Value
					case xmlAttrTitle:
						docPr.Title = attr.Value
					}
				}
				anchor.DocPr = docPr
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case xmlElemGraphic:
				graphic, err := d.parseDrawingGraphic(decoder, t)
				if err != nil {
					return nil, err
				}
				anchor.Graphic = graphic
			case "wrapNone":
				anchor.WrapNone = &WrapNone{}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "wrapSquare":
				wrapSquare := &WrapSquare{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "wrapText":
						wrapSquare.WrapText = attr.Value
					case xmlAttrDistT:
						wrapSquare.DistT = attr.Value
					case xmlAttrDistB:
						wrapSquare.DistB = attr.Value
					case xmlAttrDistL:
						wrapSquare.DistL = attr.Value
					case xmlAttrDistR:
						wrapSquare.DistR = attr.Value
					}
				}
				anchor.WrapSquare = wrapSquare
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "anchor" {
				return anchor, nil
			}
		}
	}
}

// parseDrawingGraphic parses drawing graphic elements
func (d *Document) parseDrawingGraphic(decoder *xml.Decoder, startElement xml.StartElement) (*DrawingGraphic, error) {
	graphic := &DrawingGraphic{}

	// Parse xmlns attributes
	for _, attr := range startElement.Attr {
		// Check xmlns attributes (namespace declarations)
		if attr.Name.Space == "xmlns" || (attr.Name.Space == "" && strings.HasPrefix(attr.Name.Local, "xmlns")) {
			if attr.Value == "http://schemas.openxmlformats.org/drawingml/2006/main" {
				graphic.Xmlns = attr.Value
			}
		}
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_drawing_graphic", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "graphicData":
				graphicData, err := d.parseGraphicData(decoder, t)
				if err != nil {
					return nil, err
				}
				graphic.GraphicData = graphicData
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == xmlElemGraphic {
				return graphic, nil
			}
		}
	}
}

// parseGraphicData parses graphic data elements
func (d *Document) parseGraphicData(decoder *xml.Decoder, startElement xml.StartElement) (*GraphicData, error) {
	graphicData := &GraphicData{}

	// Parse attributes
	for _, attr := range startElement.Attr {
		if attr.Name.Local == "uri" {
			graphicData.URI = attr.Value
		}
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_graphic_data", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pic":
				pic, err := d.parsePicElement(decoder, t)
				if err != nil {
					return nil, err
				}
				graphicData.Pic = pic
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "graphicData" {
				return graphicData, nil
			}
		}
	}
}

// parsePicElement parses a picture element
//
//nolint:gocognit
func (d *Document) parsePicElement(decoder *xml.Decoder, startElement xml.StartElement) (*PicElement, error) {
	pic := &PicElement{}

	// Parse xmlns attributes
	for _, attr := range startElement.Attr {
		// Check xmlns attributes (namespace declarations)
		if attr.Name.Space == "xmlns" || (attr.Name.Space == "" && strings.HasPrefix(attr.Name.Local, "xmlns")) {
			if attr.Value == "http://schemas.openxmlformats.org/drawingml/2006/picture" {
				pic.Xmlns = attr.Value
			}
		}
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_pic_element", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "nvPicPr":
				nvPicPr, err := d.parseNvPicPr(decoder, t)
				if err != nil {
					return nil, err
				}
				pic.NvPicPr = nvPicPr
			case "blipFill":
				blipFill, err := d.parseBlipFill(decoder, t)
				if err != nil {
					return nil, err
				}
				pic.BlipFill = blipFill
			case "spPr":
				spPr, err := d.parseSpPr(decoder, t)
				if err != nil {
					return nil, err
				}
				pic.SpPr = spPr
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "pic" {
				return pic, nil
			}
		}
	}
}

// parseNvPicPr parses non-visual picture properties
//
//nolint:gocognit
func (d *Document) parseNvPicPr(decoder *xml.Decoder, startElement xml.StartElement) (*NvPicPr, error) {
	nvPicPr := &NvPicPr{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_nv_pic_pr", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "cNvPr":
				cNvPr := &CNvPr{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "id":
						cNvPr.ID = attr.Value
					case xmlAttrName:
						cNvPr.Name = attr.Value
					case xmlAttrDescr:
						cNvPr.Descr = attr.Value
					case "title":
						cNvPr.Title = attr.Value
					}
				}
				nvPicPr.CNvPr = cNvPr
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "cNvPicPr":
				cNvPicPr := &CNvPicPr{}
				// Parse picLocks if present
				nvPicPr.CNvPicPr = cNvPicPr
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "nvPicPr" {
				return nvPicPr, nil
			}
		}
	}
}

// parseBlipFill parses picture fill
//
//nolint:gocognit
func (d *Document) parseBlipFill(decoder *xml.Decoder, startElement xml.StartElement) (*BlipFill, error) {
	blipFill := &BlipFill{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_blip_fill", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "blip":
				blip := &Blip{}
				for _, attr := range t.Attr {
					if attr.Name.Local == "embed" {
						blip.Embed = attr.Value
					}
				}
				// Capture inner XML to preserve child elements like alphaModFix (transparency)
				raw, err := d.captureElement(decoder, t)
				if err != nil {
					return nil, err
				}
				blip.InnerXML = raw.InnerXML
				blipFill.Blip = blip
			case "stretch":
				blipFill.Stretch = &Stretch{FillRect: &FillRect{}}
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "blipFill" {
				return blipFill, nil
			}
		}
	}
}

// parseSpPr parses shape properties
//
//nolint:gocognit
func (d *Document) parseSpPr(decoder *xml.Decoder, startElement xml.StartElement) (*SpPr, error) {
	spPr := &SpPr{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_sp_pr", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "xfrm":
				xfrm, err := d.parseXfrm(decoder, t)
				if err != nil {
					return nil, err
				}
				spPr.Xfrm = xfrm
			case "prstGeom":
				prstGeom := &PrstGeom{AvLst: &AvLst{}}
				for _, attr := range t.Attr {
					if attr.Name.Local == "prst" {
						prstGeom.Prst = attr.Value
					}
				}
				spPr.PrstGeom = prstGeom
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "spPr" {
				return spPr, nil
			}
		}
	}
}

// parseXfrm parses transform elements
//
//nolint:gocognit
func (d *Document) parseXfrm(decoder *xml.Decoder, startElement xml.StartElement) (*Xfrm, error) {
	xfrm := &Xfrm{}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, WrapError("parse_xfrm", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "off":
				off := &Off{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "x":
						off.X = attr.Value
					case "y":
						off.Y = attr.Value
					}
				}
				xfrm.Off = off
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "ext":
				ext := &Ext{}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "cx":
						ext.Cx = attr.Value
					case "cy":
						ext.Cy = attr.Value
					}
				}
				xfrm.Ext = ext
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := d.skipElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "xfrm" {
				return xfrm, nil
			}
		}
	}
}
