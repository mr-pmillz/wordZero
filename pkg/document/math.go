// Package document provides core Word document operations.
package document

import (
	"encoding/xml"
)

// OfficeMath represents an Office math formula element,
// corresponding to the m:oMath element in OMML.
type OfficeMath struct {
	XMLName xml.Name `xml:"m:oMath"`
	Xmlns   string   `xml:"xmlns:m,attr,omitempty"`
	RawXML  string   `xml:",innerxml"` // Stores the inner XML content
}

// OfficeMathPara represents an Office math formula paragraph (for block-level formulas),
// corresponding to the m:oMathPara element in OMML.
type OfficeMathPara struct {
	XMLName xml.Name    `xml:"m:oMathPara"`
	Xmlns   string      `xml:"xmlns:m,attr,omitempty"`
	Math    *OfficeMath `xml:"m:oMath"`
}

// MathParagraph represents a paragraph containing a math formula,
// used to embed math formulas in a document.
type MathParagraph struct {
	XMLName    xml.Name             `xml:"w:p"`
	Properties *ParagraphProperties `xml:"w:pPr,omitempty"`
	Math       *OfficeMath          `xml:"m:oMath,omitempty"`
	MathPara   *OfficeMathPara      `xml:"m:oMathPara,omitempty"`
	Runs       []Run                `xml:"w:r"`
}

// MarshalXML provides custom XML serialization for MathParagraph.
func (mp *MathParagraph) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Start the paragraph element
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Serialize paragraph properties
	if mp.Properties != nil {
		if err := e.Encode(mp.Properties); err != nil {
			return err
		}
	}

	// Serialize Runs (text before the formula)
	for _, run := range mp.Runs {
		if err := e.Encode(run); err != nil {
			return err
		}
	}

	// Serialize math formula (block-level)
	if mp.MathPara != nil {
		if err := e.Encode(mp.MathPara); err != nil {
			return err
		}
	}

	// Serialize math formula (inline)
	if mp.Math != nil {
		if err := e.Encode(mp.Math); err != nil {
			return err
		}
	}

	// End the paragraph element
	return e.EncodeToken(start.End())
}

// ElementType returns the element type for the math paragraph.
func (mp *MathParagraph) ElementType() string {
	return "math_paragraph"
}

// AddMathFormula adds a math formula to the document.
// latex: the math formula in LaTeX format.
// isBlock: whether it is a block-level formula (true for block, false for inline).
func (d *Document) AddMathFormula(latex string, isBlock bool) *MathParagraph {
	DebugMsgf(MsgAddingMathFormula, latex, isBlock)

	mp := &MathParagraph{
		Runs: []Run{},
	}

	// Create the formula content.
	// Note: RawXML is used to store the formula content because the OMML structure is complex.
	// The actual LaTeX-to-OMML conversion is done by the LaTeXToOMML function in the markdown package.
	if isBlock {
		mp.MathPara = &OfficeMathPara{
			Xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/math",
			Math: &OfficeMath{
				Xmlns:  "http://schemas.openxmlformats.org/officeDocument/2006/math",
				RawXML: latex, // This stores pre-processed OMML content
			},
		}
	} else {
		mp.Math = &OfficeMath{
			Xmlns:  "http://schemas.openxmlformats.org/officeDocument/2006/math",
			RawXML: latex,
		}
	}

	d.Body.Elements = append(d.Body.Elements, mp)
	return mp
}

// AddInlineMath adds an inline math formula to the paragraph.
// This appends a math formula at the end of the current paragraph.
func (p *Paragraph) AddInlineMath(ommlContent string) {
	DebugMsg(MsgAddingInlineMathFormula)

	// Create a special Run to contain formula reference
	// Note: In Word, inline formulas are implemented through special oMath elements
	// Here we use a placeholder method, actual implementation needs to modify paragraph serialization logic
	run := Run{
		Text: Text{
			Content: "[Formula]", // Placeholder, actual formula content is processed during serialization
			Space:   "preserve",
		},
	}
	p.Runs = append(p.Runs, run)
}
