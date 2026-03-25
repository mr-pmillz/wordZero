// Package document provides page settings functionality for Word documents.
package document

import (
	"encoding/xml"
	"errors"
	"fmt"
	"strconv"
)

// PageOrientation represents the page orientation type.
type PageOrientation string

const (
	// OrientationPortrait represents portrait orientation.
	OrientationPortrait PageOrientation = "portrait"
	// OrientationLandscape represents landscape orientation.
	OrientationLandscape PageOrientation = "landscape"
)

// DocGridType represents the document grid type.
type DocGridType string

const (
	// DocGridDefault represents the default grid type.
	DocGridDefault DocGridType = "default"
	// DocGridLines affects line spacing only.
	DocGridLines DocGridType = "lines"
	// DocGridSnapToChars snaps text to character grid.
	DocGridSnapToChars DocGridType = "snapToChars"
	// DocGridSnapToLines snaps text lines to the grid.
	DocGridSnapToLines DocGridType = "snapToLines"
)

// PageSize represents a page size type.
type PageSize string

const (
	// PageSizeA4 represents A4 paper (210mm x 297mm).
	PageSizeA4 PageSize = "A4"
	// PageSizeLetter represents US Letter paper (8.5" x 11").
	PageSizeLetter PageSize = "Letter"
	// PageSizeLegal represents US Legal paper (8.5" x 14").
	PageSizeLegal PageSize = "Legal"
	// PageSizeA3 represents A3 paper (297mm x 420mm).
	PageSizeA3 PageSize = "A3"
	// PageSizeA5 represents A5 paper (148mm x 210mm).
	PageSizeA5 PageSize = "A5"
	// PageSizeCustom represents a custom page size.
	PageSizeCustom PageSize = "Custom"
)

// Page settings related errors
var (
	// ErrInvalidPageSettings indicates invalid page settings.
	ErrInvalidPageSettings = errors.New("invalid page settings")
)

// SectionType represents the section type (e.g., continuous, nextPage).
type SectionType struct {
	XMLName xml.Name `xml:"w:type"`
	Val     string   `xml:"w:val,attr"`
}

// SectionProperties holds section properties including page settings.
// Field order must match OOXML CT_SectPr schema (ECMA-376 §17.6.17).
type SectionProperties struct {
	XMLName          xml.Name                 `xml:"w:sectPr"`
	XmlnsR           string                   `xml:"xmlns:r,attr,omitempty"`
	HeaderReferences []*HeaderFooterReference `xml:"w:headerReference,omitempty"`
	FooterReferences []*FooterReference       `xml:"w:footerReference,omitempty"`
	SectionType      *SectionType             `xml:"w:type,omitempty"`
	PageSize         *PageSizeXML             `xml:"w:pgSz,omitempty"`
	PageMargins      *PageMargin              `xml:"w:pgMar,omitempty"`
	Columns          *Columns                 `xml:"w:cols,omitempty"`
	PageNumType      *PageNumType             `xml:"w:pgNumType,omitempty"`
	TitlePage        *TitlePage               `xml:"w:titlePg,omitempty"`
	DocGrid          *DocGrid                 `xml:"w:docGrid,omitempty"`
}

// PageSizeXML represents the page size XML structure.
type PageSizeXML struct {
	XMLName xml.Name `xml:"w:pgSz"`
	W       string   `xml:"w:w,attr"`                // Page width (twips)
	H       string   `xml:"w:h,attr"`                // Page height (twips)
	Orient  string   `xml:"w:orient,attr,omitempty"` // Page orientation (omit when empty to match Word behavior)
}

// PageMargin represents page margin settings.
type PageMargin struct {
	XMLName xml.Name `xml:"w:pgMar"`
	Top     string   `xml:"w:top,attr"`    // Top margin (twips)
	Right   string   `xml:"w:right,attr"`  // Right margin (twips)
	Bottom  string   `xml:"w:bottom,attr"` // Bottom margin (twips)
	Left    string   `xml:"w:left,attr"`   // Left margin (twips)
	Header  string   `xml:"w:header,attr"` // Header distance (twips)
	Footer  string   `xml:"w:footer,attr"` // Footer distance (twips)
	Gutter  string   `xml:"w:gutter,attr"` // Gutter width (twips)
}

// Columns represents column layout settings.
type Columns struct {
	XMLName xml.Name `xml:"w:cols"`
	Space   string   `xml:"w:space,attr,omitempty"` // Column spacing
	Num     string   `xml:"w:num,attr,omitempty"`   // Number of columns
}

// PageNumType represents page number type settings.
type PageNumType struct {
	XMLName xml.Name `xml:"w:pgNumType"`
	Fmt     string   `xml:"w:fmt,attr,omitempty"`
	Start   string   `xml:"w:start,attr,omitempty"` // Starting page number
}

// PageSettings holds page settings configuration.
type PageSettings struct {
	// Page size
	Size PageSize
	// Custom dimensions (used when Size is Custom)
	CustomWidth  float64 // Custom width (millimeters)
	CustomHeight float64 // Custom height (millimeters)
	// Page orientation
	Orientation PageOrientation
	// Page margins (millimeters)
	MarginTop    float64
	MarginRight  float64
	MarginBottom float64
	MarginLeft   float64
	// Header and footer distance (millimeters)
	HeaderDistance float64
	FooterDistance float64
	// Gutter width (millimeters)
	GutterWidth float64
	// Document grid settings
	DocGridType      DocGridType // Document grid type
	DocGridLinePitch int         // Line grid pitch (1/20 of a point)
	DocGridCharSpace int         // Character spacing
}

// Predefined page sizes (millimeters)
var predefinedSizes = map[PageSize]struct {
	width  float64
	height float64
}{
	PageSizeA4:     {210, 297},
	PageSizeLetter: {215.9, 279.4}, // 8.5" x 11"
	PageSizeLegal:  {215.9, 355.6}, // 8.5" x 14"
	PageSizeA3:     {297, 420},
	PageSizeA5:     {148, 210},
}

// DefaultPageSettings returns the default page settings (A4, portrait).
func DefaultPageSettings() *PageSettings {
	return &PageSettings{
		Size:             PageSizeA4,
		Orientation:      OrientationPortrait,
		MarginTop:        25.4, // 1 inch
		MarginRight:      25.4, // 1 inch
		MarginBottom:     25.4, // 1 inch
		MarginLeft:       25.4, // 1 inch
		HeaderDistance:   12.7, // 0.5 inch
		FooterDistance:   12.7, // 0.5 inch
		GutterWidth:      0,    // No gutter
		DocGridType:      DocGridLines,
		DocGridLinePitch: 312, // Default line grid pitch
		DocGridCharSpace: 0,
	}
}

// SetPageSettings sets the page properties for the document.
func (d *Document) SetPageSettings(settings *PageSettings) error {
	if settings == nil {
		return WrapError("SetPageSettings", errors.New("page settings cannot be nil"))
	}

	// Validate page settings
	if err := validatePageSettings(settings); err != nil {
		return WrapError("SetPageSettings", err)
	}

	// Get or create section properties
	sectPr := d.getSectionProperties()

	// Set page size
	width, height := getPageDimensions(settings)
	sectPr.PageSize = &PageSizeXML{
		W:      fmt.Sprintf("%.0f", mmToTwips(width)),
		H:      fmt.Sprintf("%.0f", mmToTwips(height)),
		Orient: string(settings.Orientation),
	}

	// Set page margins
	sectPr.PageMargins = &PageMargin{
		Top:    fmt.Sprintf("%.0f", mmToTwips(settings.MarginTop)),
		Right:  fmt.Sprintf("%.0f", mmToTwips(settings.MarginRight)),
		Bottom: fmt.Sprintf("%.0f", mmToTwips(settings.MarginBottom)),
		Left:   fmt.Sprintf("%.0f", mmToTwips(settings.MarginLeft)),
		Header: fmt.Sprintf("%.0f", mmToTwips(settings.HeaderDistance)),
		Footer: fmt.Sprintf("%.0f", mmToTwips(settings.FooterDistance)),
		Gutter: fmt.Sprintf("%.0f", mmToTwips(settings.GutterWidth)),
	}

	// Set document grid
	if settings.DocGridType != "" {
		sectPr.DocGrid = &DocGrid{
			Type:      string(settings.DocGridType),
			LinePitch: strconv.Itoa(settings.DocGridLinePitch),
		}

		if settings.DocGridCharSpace > 0 {
			sectPr.DocGrid.CharSpace = strconv.Itoa(settings.DocGridCharSpace)
		}
	}

	InfoMsgf(MsgPageSettingsUpdated, settings.Size, settings.Orientation)
	return nil
}

// GetPageSettings returns the current page settings for the document.
func (d *Document) GetPageSettings() *PageSettings {
	sectPr := d.getSectionProperties()
	settings := DefaultPageSettings()

	if sectPr.PageSize != nil {
		// Parse page dimensions
		width := twipsToMM(parseFloat(sectPr.PageSize.W))
		height := twipsToMM(parseFloat(sectPr.PageSize.H))

		// Determine if this is a predefined size
		settings.Size = identifyPageSize(width, height)
		if settings.Size == PageSizeCustom {
			settings.CustomWidth = width
			settings.CustomHeight = height
		}

		// Set orientation
		if sectPr.PageSize.Orient == string(OrientationLandscape) {
			settings.Orientation = OrientationLandscape
		} else {
			settings.Orientation = OrientationPortrait
		}
	}

	if sectPr.PageMargins != nil {
		// Parse page margins
		settings.MarginTop = twipsToMM(parseFloat(sectPr.PageMargins.Top))
		settings.MarginRight = twipsToMM(parseFloat(sectPr.PageMargins.Right))
		settings.MarginBottom = twipsToMM(parseFloat(sectPr.PageMargins.Bottom))
		settings.MarginLeft = twipsToMM(parseFloat(sectPr.PageMargins.Left))
		settings.HeaderDistance = twipsToMM(parseFloat(sectPr.PageMargins.Header))
		settings.FooterDistance = twipsToMM(parseFloat(sectPr.PageMargins.Footer))
		settings.GutterWidth = twipsToMM(parseFloat(sectPr.PageMargins.Gutter))
	}

	// Parse document grid settings
	if sectPr.DocGrid != nil {
		if sectPr.DocGrid.Type != "" {
			settings.DocGridType = DocGridType(sectPr.DocGrid.Type)
		}

		if sectPr.DocGrid.LinePitch != "" {
			settings.DocGridLinePitch = int(parseFloat(sectPr.DocGrid.LinePitch))
		}

		if sectPr.DocGrid.CharSpace != "" {
			settings.DocGridCharSpace = int(parseFloat(sectPr.DocGrid.CharSpace))
		}
	}

	return settings
}

// SetPageSize sets the page size for the document.
func (d *Document) SetPageSize(size PageSize) error {
	settings := d.GetPageSettings()
	settings.Size = size
	return d.SetPageSettings(settings)
}

// SetCustomPageSize sets a custom page size in millimeters.
func (d *Document) SetCustomPageSize(width, height float64) error {
	if width <= 0 || height <= 0 {
		return WrapError("SetCustomPageSize", errors.New("page dimensions must be greater than 0"))
	}

	settings := d.GetPageSettings()
	settings.Size = PageSizeCustom
	settings.CustomWidth = width
	settings.CustomHeight = height
	return d.SetPageSettings(settings)
}

// SetPageOrientation sets the page orientation for the document.
func (d *Document) SetPageOrientation(orientation PageOrientation) error {
	settings := d.GetPageSettings()
	settings.Orientation = orientation
	return d.SetPageSettings(settings)
}

// SetPageMargins sets the page margins in millimeters.
func (d *Document) SetPageMargins(top, right, bottom, left float64) error {
	if top < 0 || right < 0 || bottom < 0 || left < 0 {
		return WrapError("SetPageMargins", errors.New("page margins cannot be negative"))
	}

	settings := d.GetPageSettings()
	settings.MarginTop = top
	settings.MarginRight = right
	settings.MarginBottom = bottom
	settings.MarginLeft = left
	return d.SetPageSettings(settings)
}

// SetHeaderFooterDistance sets the header and footer distance in millimeters.
func (d *Document) SetHeaderFooterDistance(header, footer float64) error {
	if header < 0 || footer < 0 {
		return WrapError("SetHeaderFooterDistance", errors.New("header and footer distance cannot be negative"))
	}

	settings := d.GetPageSettings()
	settings.HeaderDistance = header
	settings.FooterDistance = footer
	return d.SetPageSettings(settings)
}

// SetGutterWidth sets the gutter width in millimeters.
func (d *Document) SetGutterWidth(width float64) error {
	if width < 0 {
		return WrapError("SetGutterWidth", errors.New("gutter width cannot be negative"))
	}

	settings := d.GetPageSettings()
	settings.GutterWidth = width
	return d.SetPageSettings(settings)
}

// getSectionProperties returns or creates the section properties.
func (d *Document) getSectionProperties() *SectionProperties {
	if d.Body == nil {
		return &SectionProperties{}
	}

	// Search for existing SectionProperties in Elements (may be at any position)
	for _, element := range d.Body.Elements {
		if sectPr, ok := element.(*SectionProperties); ok {
			return sectPr
		}
	}

	// If not found, create new section properties and append to the end
	sectPr := &SectionProperties{}
	d.Body.Elements = append(d.Body.Elements, sectPr)

	return sectPr
}

// SetSectionProperties replaces or sets the document-level section properties.
func (d *Document) SetSectionProperties(sectPr *SectionProperties) {
	if sectPr == nil {
		return
	}

	if d.Body == nil {
		d.Body = &Body{Elements: []interface{}{sectPr}}
		return
	}

	for i, element := range d.Body.Elements {
		if _, ok := element.(*SectionProperties); ok {
			d.Body.Elements[i] = sectPr
			return
		}
	}

	d.Body.Elements = append(d.Body.Elements, sectPr)
}

// ElementType returns the element type for section properties.
func (s *SectionProperties) ElementType() string {
	return "sectionProperties"
}

// validatePageSettings validates the page settings.
func validatePageSettings(settings *PageSettings) error {
	// Validate page dimensions
	if settings.Size == PageSizeCustom {
		if settings.CustomWidth <= 0 || settings.CustomHeight <= 0 {
			return errors.New("custom page dimensions must be greater than 0")
		}

		// Check dimension range (minimum and maximum sizes supported by Word)
		const minSize = 12.7  // 0.5 inch
		const maxSize = 558.8 // 22 inches

		if settings.CustomWidth < minSize || settings.CustomWidth > maxSize ||
			settings.CustomHeight < minSize || settings.CustomHeight > maxSize {
			return fmt.Errorf("page dimensions must be within %.1f-%.1fmm range", minSize, maxSize)
		}
	}

	// Validate orientation
	if settings.Orientation != OrientationPortrait && settings.Orientation != OrientationLandscape {
		return errors.New("invalid page orientation")
	}

	return nil
}

// getPageDimensions returns the page dimensions in millimeters.
func getPageDimensions(settings *PageSettings) (width, height float64) {
	if settings.Size == PageSizeCustom {
		width = settings.CustomWidth
		height = settings.CustomHeight
	} else {
		size, exists := predefinedSizes[settings.Size]
		if !exists {
			// Default to A4
			size = predefinedSizes[PageSizeA4]
		}
		width = size.width
		height = size.height
	}

	// If landscape, swap width and height
	if settings.Orientation == OrientationLandscape {
		width, height = height, width
	}

	return width, height
}

// identifyPageSize identifies the page size type from dimensions.
func identifyPageSize(width, height float64) PageSize {
	// Allow 1mm tolerance
	const tolerance = 1.0

	for size, dims := range predefinedSizes {
		if (abs(width-dims.width) < tolerance && abs(height-dims.height) < tolerance) ||
			(abs(width-dims.height) < tolerance && abs(height-dims.width) < tolerance) {
			return size
		}
	}

	return PageSizeCustom
}

// mmToTwips converts millimeters to twips (1mm = 56.69 twips).
func mmToTwips(mm float64) float64 {
	return mm * 56.692913385827
}

// twipsToMM converts twips to millimeters.
func twipsToMM(twips float64) float64 {
	return twips / 56.692913385827
}

// parseFloat safely parses a float string.
func parseFloat(s string) float64 {
	if s == "" {
		return 0
	}

	// Use strconv.ParseFloat to parse the float
	if val, err := strconv.ParseFloat(s, 64); err == nil {
		return val
	}

	return 0
}

// abs returns the absolute value of a float.
func abs(x float64) float64 {
	if x < 0 {
		return -x
	}
	return x
}

// DocGrid represents document grid settings.
type DocGrid struct {
	XMLName   xml.Name `xml:"w:docGrid"`
	Type      string   `xml:"w:type,attr,omitempty"`      // Grid type
	LinePitch string   `xml:"w:linePitch,attr,omitempty"` // Line grid pitch
	CharSpace string   `xml:"w:charSpace,attr,omitempty"` // Character spacing
}

// SetDocGrid sets the document grid settings.
func (d *Document) SetDocGrid(gridType DocGridType, linePitch int, charSpace int) error {
	if gridType == "" {
		return WrapError("SetDocGrid", errors.New("grid type cannot be empty"))
	}

	settings := d.GetPageSettings()
	settings.DocGridType = gridType
	settings.DocGridLinePitch = linePitch
	settings.DocGridCharSpace = charSpace
	return d.SetPageSettings(settings)
}

// ClearDocGrid clears the document grid settings.
func (d *Document) ClearDocGrid() error {
	sectPr := d.getSectionProperties()
	sectPr.DocGrid = nil
	return nil
}

// SetStartPageNumber sets the starting page number for the document.
// Valid range is 0-32767 per OOXML specification.
func (d *Document) SetStartPageNumber(startPage int) error {
	if startPage < 0 {
		return fmt.Errorf("page number cannot be negative: %d", startPage)
	}
	sectPr := d.getSectionPropertiesForHeaderFooter()
	if sectPr.PageNumType == nil {
		sectPr.PageNumType = &PageNumType{}
	}
	sectPr.PageNumType.Start = strconv.Itoa(startPage)
	return nil
}

// ResetPageNumber resets the page number at the most recent section break.
// It searches Body.Elements in reverse for a paragraph with SectionProperties
// and sets the PageNumType.Start on it. If no section break is found, it falls
// back to the document end section properties.
func (d *Document) ResetPageNumber(startNumber int) {
	if d.Body == nil {
		return
	}

	// Search Body.Elements in reverse for a paragraph with SectionProperties
	for i := len(d.Body.Elements) - 1; i >= 0; i-- {
		if para, ok := d.Body.Elements[i].(*Paragraph); ok {
			if para.Properties != nil && para.Properties.SectionProperties != nil {
				sectPr := para.Properties.SectionProperties
				if sectPr.PageNumType == nil {
					sectPr.PageNumType = &PageNumType{}
				}
				sectPr.PageNumType.Start = strconv.Itoa(startNumber)
				return
			}
		}
	}

	// Fall back to document end section properties
	sectPr := d.getSectionPropertiesForHeaderFooter()
	if sectPr.PageNumType == nil {
		sectPr.PageNumType = &PageNumType{}
	}
	sectPr.PageNumType.Start = strconv.Itoa(startNumber)
}

// RestartPageNumber restarts page numbering from 1 at the most recent section break.
func (d *Document) RestartPageNumber() {
	d.ResetPageNumber(1)
}
