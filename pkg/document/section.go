// Package document provides section break functionality for Word documents.
package document

import (
	"fmt"
	"strconv"
)

// AddSectionBreak adds a section break to a paragraph with the given orientation.
// It continues page numbering from the previous section and inherits header/footer
// references from the current section.
func (p *Paragraph) AddSectionBreak(orient PageOrientation, doc *Document) {
	p.AddSectionBreakWithStartPage(orient, doc, 0, true)
}

// AddSectionBreakWithStartPage adds a section break with page number control.
// startPage=0 means continue numbering from the previous section.
// startPage>0 restarts numbering at that page number in the next section.
// inheritHeaderFooter copies header/footer references from the current section
// into the section being closed.
func (p *Paragraph) AddSectionBreakWithStartPage(orient PageOrientation, doc *Document, startPage int, inheritHeaderFooter bool) {
	// Ensure paragraph properties exist
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	// Get the current section properties from the document
	currentSectPr := doc.getCurrentSectionProperties()

	// Build page size based on orientation
	// Default A4 dimensions in TWIPs: width=11906, height=16838
	var widthTwips, heightTwips string
	currentSettings := doc.GetPageSettings()
	pageWidth, pageHeight := getPageDimensions(currentSettings)

	if orient == OrientationLandscape {
		// For landscape, swap width and height if the current settings are portrait
		if currentSettings.Orientation == OrientationPortrait {
			widthTwips = fmt.Sprintf("%.0f", mmToTwips(pageHeight))
			heightTwips = fmt.Sprintf("%.0f", mmToTwips(pageWidth))
		} else {
			widthTwips = fmt.Sprintf("%.0f", mmToTwips(pageWidth))
			heightTwips = fmt.Sprintf("%.0f", mmToTwips(pageHeight))
		}
	} else {
		// Portrait
		if currentSettings.Orientation == OrientationLandscape {
			widthTwips = fmt.Sprintf("%.0f", mmToTwips(pageHeight))
			heightTwips = fmt.Sprintf("%.0f", mmToTwips(pageWidth))
		} else {
			widthTwips = fmt.Sprintf("%.0f", mmToTwips(pageWidth))
			heightTwips = fmt.Sprintf("%.0f", mmToTwips(pageHeight))
		}
	}

	// Build margins from current settings or use defaults (1 inch = 1440 TWIPs)
	var margins *PageMargin
	if currentSectPr != nil && currentSectPr.PageMargins != nil {
		margins = &PageMargin{
			Top:    currentSectPr.PageMargins.Top,
			Right:  currentSectPr.PageMargins.Right,
			Bottom: currentSectPr.PageMargins.Bottom,
			Left:   currentSectPr.PageMargins.Left,
			Header: currentSectPr.PageMargins.Header,
			Footer: currentSectPr.PageMargins.Footer,
			Gutter: currentSectPr.PageMargins.Gutter,
		}
	} else {
		margins = &PageMargin{
			Top:    "1440",
			Right:  "1440",
			Bottom: "1440",
			Left:   "1440",
			Header: "720",
			Footer: "720",
			Gutter: "0",
		}
	}

	// Create the section properties for the section being closed
	sectPr := &SectionProperties{
		PageSize: &PageSizeXML{
			W:      widthTwips,
			H:      heightTwips,
			Orient: string(orient),
		},
		PageMargins: margins,
	}

	// If inheritHeaderFooter, copy header/footer references from the current section
	if inheritHeaderFooter && currentSectPr != nil {
		if len(currentSectPr.HeaderReferences) > 0 {
			sectPr.XmlnsR = ooXMLRelationshipsBase
			sectPr.HeaderReferences = make([]*HeaderFooterReference, len(currentSectPr.HeaderReferences))
			for i, ref := range currentSectPr.HeaderReferences {
				sectPr.HeaderReferences[i] = &HeaderFooterReference{
					Type: ref.Type,
					ID:   ref.ID,
				}
			}
		}
		if len(currentSectPr.FooterReferences) > 0 {
			if sectPr.XmlnsR == "" {
				sectPr.XmlnsR = ooXMLRelationshipsBase
			}
			sectPr.FooterReferences = make([]*FooterReference, len(currentSectPr.FooterReferences))
			for i, ref := range currentSectPr.FooterReferences {
				sectPr.FooterReferences[i] = &FooterReference{
					Type: ref.Type,
					ID:   ref.ID,
				}
			}
		}
	}

	// If startPage > 0, set PageNumType.Start on this section break's properties.
	// This tells Word to restart page numbering at the specified value for the
	// section that follows this break.
	if startPage > 0 {
		sectPr.PageNumType = &PageNumType{Start: strconv.Itoa(startPage)}
	}

	// Assign section properties to the paragraph
	p.Properties.SectionProperties = sectPr
}

// AddSectionBreakWithPageNumber is a convenience wrapper that always inherits
// headers/footers from the current section.
func (p *Paragraph) AddSectionBreakWithPageNumber(orient PageOrientation, doc *Document, startPage int) {
	p.AddSectionBreakWithStartPage(orient, doc, startPage, true)
}

// AddSectionBreakContinuous adds a continuous section break (no page break).
// The content continues on the same page but starts a new section, allowing
// different section-level formatting (e.g., column layout, page numbering).
func (p *Paragraph) AddSectionBreakContinuous(orient PageOrientation, doc *Document) {
	// Ensure paragraph properties exist
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	// Get the current section properties from the document
	currentSectPr := doc.getCurrentSectionProperties()

	// Build page size based on orientation
	var widthTwips, heightTwips string
	currentSettings := doc.GetPageSettings()
	pageWidth, pageHeight := getPageDimensions(currentSettings)

	if orient == OrientationLandscape {
		if currentSettings.Orientation == OrientationPortrait {
			widthTwips = fmt.Sprintf("%.0f", mmToTwips(pageHeight))
			heightTwips = fmt.Sprintf("%.0f", mmToTwips(pageWidth))
		} else {
			widthTwips = fmt.Sprintf("%.0f", mmToTwips(pageWidth))
			heightTwips = fmt.Sprintf("%.0f", mmToTwips(pageHeight))
		}
	} else {
		if currentSettings.Orientation == OrientationLandscape {
			widthTwips = fmt.Sprintf("%.0f", mmToTwips(pageHeight))
			heightTwips = fmt.Sprintf("%.0f", mmToTwips(pageWidth))
		} else {
			widthTwips = fmt.Sprintf("%.0f", mmToTwips(pageWidth))
			heightTwips = fmt.Sprintf("%.0f", mmToTwips(pageHeight))
		}
	}

	// Build margins from current settings or use defaults
	var margins *PageMargin
	if currentSectPr != nil && currentSectPr.PageMargins != nil {
		margins = &PageMargin{
			Top:    currentSectPr.PageMargins.Top,
			Right:  currentSectPr.PageMargins.Right,
			Bottom: currentSectPr.PageMargins.Bottom,
			Left:   currentSectPr.PageMargins.Left,
			Header: currentSectPr.PageMargins.Header,
			Footer: currentSectPr.PageMargins.Footer,
			Gutter: currentSectPr.PageMargins.Gutter,
		}
	} else {
		margins = &PageMargin{
			Top:    "1440",
			Right:  "1440",
			Bottom: "1440",
			Left:   "1440",
			Header: "720",
			Footer: "720",
			Gutter: "0",
		}
	}

	// Create section properties with continuous type
	sectPr := &SectionProperties{
		SectionType: &SectionType{Val: "continuous"},
		PageSize: &PageSizeXML{
			W:      widthTwips,
			H:      heightTwips,
			Orient: string(orient),
		},
		PageMargins: margins,
	}

	// Copy header/footer references from the current section
	if currentSectPr != nil {
		if len(currentSectPr.HeaderReferences) > 0 {
			sectPr.XmlnsR = ooXMLRelationshipsBase
			sectPr.HeaderReferences = make([]*HeaderFooterReference, len(currentSectPr.HeaderReferences))
			for i, ref := range currentSectPr.HeaderReferences {
				sectPr.HeaderReferences[i] = &HeaderFooterReference{
					Type: ref.Type,
					ID:   ref.ID,
				}
			}
		}
		if len(currentSectPr.FooterReferences) > 0 {
			if sectPr.XmlnsR == "" {
				sectPr.XmlnsR = ooXMLRelationshipsBase
			}
			sectPr.FooterReferences = make([]*FooterReference, len(currentSectPr.FooterReferences))
			for i, ref := range currentSectPr.FooterReferences {
				sectPr.FooterReferences[i] = &FooterReference{
					Type: ref.Type,
					ID:   ref.ID,
				}
			}
		}
	}

	// Assign section properties to the paragraph
	p.Properties.SectionProperties = sectPr
}

// getCurrentSectionProperties returns the most recent section properties.
// Priority:
//  1. Last standalone SectionProperties in Body.Elements
//  2. Last paragraph with SectionProperties in its properties
//  3. Creates a new SectionProperties via getSectionPropertiesForHeaderFooter
func (d *Document) getCurrentSectionProperties() *SectionProperties {
	if d.Body == nil {
		return d.getSectionPropertiesForHeaderFooter()
	}

	// First, look for standalone SectionProperties in Body.Elements (end section)
	var lastStandalone *SectionProperties
	for _, element := range d.Body.Elements {
		if sectPr, ok := element.(*SectionProperties); ok {
			lastStandalone = sectPr
		}
	}

	// Also search for the last paragraph with SectionProperties
	var lastParaSectPr *SectionProperties
	for _, element := range d.Body.Elements {
		if para, ok := element.(*Paragraph); ok {
			if para.Properties != nil && para.Properties.SectionProperties != nil {
				lastParaSectPr = para.Properties.SectionProperties
			}
		}
	}

	// Prefer standalone over paragraph-level (standalone is the document end section)
	if lastStandalone != nil {
		return lastStandalone
	}
	if lastParaSectPr != nil {
		return lastParaSectPr
	}

	// Nothing found; create new section properties
	return d.getSectionPropertiesForHeaderFooter()
}
