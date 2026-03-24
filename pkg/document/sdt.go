// Package document provides SDT (Structured Document Tag) structures for Word documents.
package document

import (
	"encoding/xml"
	"fmt"
)

// SDT represents a structured document tag, used for features such as table of contents.
type SDT struct {
	XMLName    xml.Name       `xml:"w:sdt"`
	Properties *SDTProperties `xml:"w:sdtPr"`
	EndPr      *SDTEndPr      `xml:"w:sdtEndPr,omitempty"`
	Content    *SDTContent    `xml:"w:sdtContent"`
}

// ElementType returns the SDT element type.
func (sdt *SDT) ElementType() string {
	return "sdt"
}

// SDTProperties represents SDT properties.
type SDTProperties struct {
	XMLName     xml.Name        `xml:"w:sdtPr"`
	RunPr       *RunProperties  `xml:"w:rPr,omitempty"`
	ID          *SDTID          `xml:"w:id,omitempty"`
	Color       *SDTColor       `xml:"w15:color,omitempty"`
	DocPartObj  *DocPartObj     `xml:"w:docPartObj,omitempty"`
	Placeholder *SDTPlaceholder `xml:"w:placeholder,omitempty"`
}

// SDTEndPr represents SDT end properties.
type SDTEndPr struct {
	XMLName xml.Name       `xml:"w:sdtEndPr"`
	RunPr   *RunProperties `xml:"w:rPr,omitempty"`
}

// SDTContent represents SDT content.
type SDTContent struct {
	XMLName  xml.Name      `xml:"w:sdtContent"`
	Elements []interface{} `xml:"-"` // Uses custom serialization
}

// MarshalXML performs custom XML serialization.
func (s *SDTContent) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Start element
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Serialize each element
	for _, element := range s.Elements {
		if err := e.Encode(element); err != nil {
			return err
		}
	}

	// End element
	return e.EncodeToken(start.End())
}

// SDTID represents an SDT identifier.
type SDTID struct {
	XMLName xml.Name `xml:"w:id"`
	Val     string   `xml:"w:val,attr"`
}

// SDTColor represents an SDT color.
type SDTColor struct {
	XMLName xml.Name `xml:"w15:color"`
	Val     string   `xml:"w:val,attr"`
}

// DocPartObj represents a document part object.
type DocPartObj struct {
	XMLName        xml.Name        `xml:"w:docPartObj"`
	DocPartGallery *DocPartGallery `xml:"w:docPartGallery,omitempty"`
	DocPartUnique  *DocPartUnique  `xml:"w:docPartUnique,omitempty"`
}

// DocPartGallery represents a document part gallery.
type DocPartGallery struct {
	XMLName xml.Name `xml:"w:docPartGallery"`
	Val     string   `xml:"w:val,attr"`
}

// DocPartUnique represents a document part unique identifier.
type DocPartUnique struct {
	XMLName xml.Name `xml:"w:docPartUnique"`
}

// SDTPlaceholder represents an SDT placeholder.
type SDTPlaceholder struct {
	XMLName xml.Name `xml:"w:placeholder"`
	DocPart *DocPart `xml:"w:docPart,omitempty"`
}

// DocPart represents a document part.
type DocPart struct {
	XMLName xml.Name `xml:"w:docPart"`
	Val     string   `xml:"w:val,attr"`
}

// Tab represents a tab character.
type Tab struct {
	XMLName xml.Name `xml:"w:tab"`
}

// CreateTOCSDT creates a table of contents SDT structure.
func (d *Document) CreateTOCSDT(title string, maxLevel int) *SDT {
	sdt := &SDT{
		Properties: &SDTProperties{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: "宋体"},
				FontSize:   &FontSize{Val: "21"},
			},
			ID:    &SDTID{Val: "147476628"},
			Color: &SDTColor{Val: "DBDBDB"},
			DocPartObj: &DocPartObj{
				DocPartGallery: &DocPartGallery{Val: "Table of Contents"},
				DocPartUnique:  &DocPartUnique{},
			},
		},
		EndPr: &SDTEndPr{
			RunPr: &RunProperties{
				FontSize: &FontSize{Val: "20"},
			},
		},
		Content: &SDTContent{
			Elements: []interface{}{},
		},
	}

	// Add table of contents title paragraph
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
				Text: Text{Content: title},
				Properties: &RunProperties{
					FontFamily: &FontFamily{ASCII: "宋体"},
					FontSize:   &FontSize{Val: "21"},
				},
			},
		},
	}

	// Add bookmark start - uses the existing BookmarkStart type
	bookmarkStart := &BookmarkStart{
		ID:   "0",
		Name: "_Toc11693_WPSOffice_Type3",
	}

	sdt.Content.Elements = append(sdt.Content.Elements, bookmarkStart, titlePara)

	return sdt
}

// AddTOCEntry adds an entry to the table of contents SDT.
func (sdt *SDT) AddTOCEntry(text string, level int, pageNum int, entryID string) {
	// Determine TOC style ID (13=toc 1, 14=toc 2, 15=toc 3, etc.)
	styleVal := fmt.Sprintf("%d", 12+level)

	// Create table of contents entry paragraph
	entryPara := &Paragraph{
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

	// Create nested SDT for placeholder text
	placeholderSDT := &SDT{
		Properties: &SDTProperties{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: "Calibri"},
				FontSize:   &FontSize{Val: "22"},
			},
			ID: &SDTID{Val: entryID},
			Placeholder: &SDTPlaceholder{
				DocPart: &DocPart{Val: generatePlaceholderGUID(level)},
			},
			Color: &SDTColor{Val: "509DF3"},
		},
		EndPr: &SDTEndPr{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: "Calibri"},
				FontSize:   &FontSize{Val: "22"},
			},
		},
		Content: &SDTContent{
			Elements: []interface{}{
				Run{
					Text: Text{Content: text},
				},
			},
		},
	}

	// Add placeholder SDT to the paragraph
	sdt.Content.Elements = append(sdt.Content.Elements, placeholderSDT)

	// Create text runs containing the tab and page number
	tabRun := Run{
		Text: Text{Content: "\t"},
	}

	pageRun := Run{
		Text: Text{Content: fmt.Sprintf("%d", pageNum)},
	}

	entryPara.Runs = append(entryPara.Runs, tabRun, pageRun)

	// Add paragraph to SDT content
	sdt.Content.Elements = append(sdt.Content.Elements, entryPara)
}

// generatePlaceholderGUID generates a placeholder GUID.
func generatePlaceholderGUID(level int) string {
	guids := map[int]string{
		1: "{b5fdec38-8301-4b26-9716-d8b31c00c718}",
		2: "{a500490c-aaae-4252-8340-aa59729b9870}",
		3: "{d7310822-77d9-4e43-95e1-4649f1e215b3}",
	}

	if guid, exists := guids[level]; exists {
		return guid
	}
	return "{b5fdec38-8301-4b26-9716-d8b31c00c718}" // Default to level 1
}

// FinalizeTOCSDT completes the table of contents SDT construction.
func (sdt *SDT) FinalizeTOCSDT() {
	// Add bookmark end - uses the existing BookmarkEnd type
	bookmarkEnd := &BookmarkEnd{
		ID: "0",
	}
	sdt.Content.Elements = append(sdt.Content.Elements, bookmarkEnd)
}
