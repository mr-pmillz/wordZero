// Package document provides footnote and endnote operations for Word documents
package document

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

// FootnoteType represents the type of a note
type FootnoteType string

const (
	// FootnoteTypeFootnote represents a footnote
	FootnoteTypeFootnote FootnoteType = "footnote"
	// FootnoteTypeEndnote represents an endnote
	FootnoteTypeEndnote FootnoteType = "endnote"
)

// Footnotes represents a collection of footnotes
type Footnotes struct {
	XMLName   xml.Name    `xml:"w:footnotes"`
	Xmlns     string      `xml:"xmlns:w,attr"`
	Footnotes []*Footnote `xml:"w:footnote"`
}

// Endnotes represents a collection of endnotes
type Endnotes struct {
	XMLName  xml.Name   `xml:"w:endnotes"`
	Xmlns    string     `xml:"xmlns:w,attr"`
	Endnotes []*Endnote `xml:"w:endnote"`
}

// Footnote represents a footnote structure
type Footnote struct {
	XMLName    xml.Name     `xml:"w:footnote"`
	Type       string       `xml:"w:type,attr,omitempty"`
	ID         string       `xml:"w:id,attr"`
	Paragraphs []*Paragraph `xml:"w:p"`
}

// Endnote represents an endnote structure
type Endnote struct {
	XMLName    xml.Name     `xml:"w:endnote"`
	Type       string       `xml:"w:type,attr,omitempty"`
	ID         string       `xml:"w:id,attr"`
	Paragraphs []*Paragraph `xml:"w:p"`
}

// FootnoteReference represents a footnote reference
type FootnoteReference struct {
	XMLName xml.Name `xml:"w:footnoteReference"`
	ID      string   `xml:"w:id,attr"`
}

// EndnoteReference represents an endnote reference
type EndnoteReference struct {
	XMLName xml.Name `xml:"w:endnoteReference"`
	ID      string   `xml:"w:id,attr"`
}

// FootnoteConfig represents footnote configuration
type FootnoteConfig struct {
	NumberFormat FootnoteNumberFormat // Number format
	StartNumber  int                  // Starting number
	RestartEach  FootnoteRestart      // Restart rule
	Position     FootnotePosition     // Position
}

// FootnoteNumberFormat represents the numbering format for footnotes
type FootnoteNumberFormat string

const (
	// FootnoteFormatDecimal represents decimal numbers
	FootnoteFormatDecimal FootnoteNumberFormat = "decimal"
	// FootnoteFormatLowerRoman represents lowercase Roman numerals
	FootnoteFormatLowerRoman FootnoteNumberFormat = "lowerRoman"
	// FootnoteFormatUpperRoman represents uppercase Roman numerals
	FootnoteFormatUpperRoman FootnoteNumberFormat = "upperRoman"
	// FootnoteFormatLowerLetter represents lowercase letters
	FootnoteFormatLowerLetter FootnoteNumberFormat = "lowerLetter"
	// FootnoteFormatUpperLetter represents uppercase letters
	FootnoteFormatUpperLetter FootnoteNumberFormat = "upperLetter"
	// FootnoteFormatSymbol represents symbols
	FootnoteFormatSymbol FootnoteNumberFormat = "symbol"
)

// FootnoteRestart represents the restart rule for footnote numbering
type FootnoteRestart string

const (
	// FootnoteRestartContinuous represents continuous numbering
	FootnoteRestartContinuous FootnoteRestart = "continuous"
	// FootnoteRestartEachSection restarts numbering at each section
	FootnoteRestartEachSection FootnoteRestart = "eachSect"
	// FootnoteRestartEachPage restarts numbering at each page
	FootnoteRestartEachPage FootnoteRestart = "eachPage"
)

// FootnotePosition represents the position of footnotes
type FootnotePosition string

const (
	// FootnotePositionPageBottom places footnotes at the bottom of the page
	FootnotePositionPageBottom FootnotePosition = "pageBottom"
	// FootnotePositionBeneathText places footnotes beneath the text
	FootnotePositionBeneathText FootnotePosition = "beneathText"
	// FootnotePositionSectionEnd places footnotes at the end of the section
	FootnotePositionSectionEnd FootnotePosition = "sectEnd"
	// FootnotePositionDocumentEnd places footnotes at the end of the document
	FootnotePositionDocumentEnd FootnotePosition = "docEnd"
)

// FootnoteProperties represents footnote properties
type FootnoteProperties struct {
	NumberFormat string `xml:"w:numFmt,attr,omitempty"`
	StartNumber  int    `xml:"w:numStart,attr,omitempty"`
	RestartRule  string `xml:"w:numRestart,attr,omitempty"`
	Position     string `xml:"w:pos,attr,omitempty"`
}

// EndnoteProperties represents endnote properties
type EndnoteProperties struct {
	NumberFormat string `xml:"w:numFmt,attr,omitempty"`
	StartNumber  int    `xml:"w:numStart,attr,omitempty"`
	RestartRule  string `xml:"w:numRestart,attr,omitempty"`
	Position     string `xml:"w:pos,attr,omitempty"`
}

// Settings represents the document settings XML structure
type Settings struct {
	XMLName                 xml.Name                 `xml:"w:settings"`
	Xmlns                   string                   `xml:"xmlns:w,attr"`
	DefaultTabStop          *DefaultTabStop          `xml:"w:defaultTabStop,omitempty"`
	CharacterSpacingControl *CharacterSpacingControl `xml:"w:characterSpacingControl,omitempty"`
	FootnotePr              *FootnotePr              `xml:"w:footnotePr,omitempty"`
	EndnotePr               *EndnotePr               `xml:"w:endnotePr,omitempty"`
}

// DefaultTabStop represents the default tab stop setting
type DefaultTabStop struct {
	XMLName xml.Name `xml:"w:defaultTabStop"`
	Val     string   `xml:"w:val,attr"`
}

// CharacterSpacingControl represents the character spacing control setting
type CharacterSpacingControl struct {
	XMLName xml.Name `xml:"w:characterSpacingControl"`
	Val     string   `xml:"w:val,attr"`
}

// FootnotePr represents footnote property settings
type FootnotePr struct {
	XMLName    xml.Name            `xml:"w:footnotePr"`
	NumFmt     *FootnoteNumFmt     `xml:"w:numFmt,omitempty"`
	NumStart   *FootnoteNumStart   `xml:"w:numStart,omitempty"`
	NumRestart *FootnoteNumRestart `xml:"w:numRestart,omitempty"`
	Pos        *FootnotePos        `xml:"w:pos,omitempty"`
}

// EndnotePr represents endnote property settings
type EndnotePr struct {
	XMLName    xml.Name           `xml:"w:endnotePr"`
	NumFmt     *EndnoteNumFmt     `xml:"w:numFmt,omitempty"`
	NumStart   *EndnoteNumStart   `xml:"w:numStart,omitempty"`
	NumRestart *EndnoteNumRestart `xml:"w:numRestart,omitempty"`
	Pos        *EndnotePos        `xml:"w:pos,omitempty"`
}

// FootnoteNumFmt represents the footnote number format
type FootnoteNumFmt struct {
	XMLName xml.Name `xml:"w:numFmt"`
	Val     string   `xml:"w:val,attr"`
}

// FootnoteNumStart represents the footnote starting number
type FootnoteNumStart struct {
	XMLName xml.Name `xml:"w:numStart"`
	Val     string   `xml:"w:val,attr"`
}

// FootnoteNumRestart represents the footnote numbering restart rule
type FootnoteNumRestart struct {
	XMLName xml.Name `xml:"w:numRestart"`
	Val     string   `xml:"w:val,attr"`
}

// FootnotePos represents the footnote position
type FootnotePos struct {
	XMLName xml.Name `xml:"w:pos"`
	Val     string   `xml:"w:val,attr"`
}

// EndnoteNumFmt represents the endnote number format
type EndnoteNumFmt struct {
	XMLName xml.Name `xml:"w:numFmt"`
	Val     string   `xml:"w:val,attr"`
}

// EndnoteNumStart represents the endnote starting number
type EndnoteNumStart struct {
	XMLName xml.Name `xml:"w:numStart"`
	Val     string   `xml:"w:val,attr"`
}

// EndnoteNumRestart represents the endnote numbering restart rule
type EndnoteNumRestart struct {
	XMLName xml.Name `xml:"w:numRestart"`
	Val     string   `xml:"w:val,attr"`
}

// EndnotePos represents the endnote position
type EndnotePos struct {
	XMLName xml.Name `xml:"w:pos"`
	Val     string   `xml:"w:val,attr"`
}

// FootnoteManager manages footnotes and endnotes
type FootnoteManager struct {
	nextFootnoteID int
	nextEndnoteID  int
	footnotes      map[string]*Footnote
	endnotes       map[string]*Endnote
}

// getFootnoteManager returns the document's footnote manager (lazy initialization)
func (d *Document) getFootnoteManager() *FootnoteManager {
	if d.footnoteManager == nil {
		d.footnoteManager = &FootnoteManager{
			nextFootnoteID: 1,
			nextEndnoteID:  1,
			footnotes:      make(map[string]*Footnote),
			endnotes:       make(map[string]*Endnote),
		}
	}
	return d.footnoteManager
}

// DefaultFootnoteConfig returns the default footnote configuration
func DefaultFootnoteConfig() *FootnoteConfig {
	return &FootnoteConfig{
		NumberFormat: FootnoteFormatDecimal,
		StartNumber:  1,
		RestartEach:  FootnoteRestartContinuous,
		Position:     FootnotePositionPageBottom,
	}
}

// AddFootnote adds a footnote to the document
func (d *Document) AddFootnote(text string, footnoteText string) error {
	return d.addFootnoteOrEndnote(text, footnoteText, FootnoteTypeFootnote)
}

// AddEndnote adds an endnote to the document
func (d *Document) AddEndnote(text string, endnoteText string) error {
	return d.addFootnoteOrEndnote(text, endnoteText, FootnoteTypeEndnote)
}

// addFootnoteOrEndnote is a shared method for adding footnotes or endnotes
func (d *Document) addFootnoteOrEndnote(text string, noteText string, noteType FootnoteType) error {
	manager := d.getFootnoteManager()

	// Ensure the footnote/endnote system is initialized
	d.ensureFootnoteInitialized(noteType)

	var noteID string
	if noteType == FootnoteTypeFootnote {
		noteID = strconv.Itoa(manager.nextFootnoteID)
		manager.nextFootnoteID++
	} else {
		noteID = strconv.Itoa(manager.nextEndnoteID)
		manager.nextEndnoteID++
	}

	// Create a paragraph containing the note reference
	paragraph := &Paragraph{}

	// Add body text
	if text != "" {
		textRun := Run{
			Text: Text{Content: text},
		}
		paragraph.Runs = append(paragraph.Runs, textRun)
	}

	// Add footnote/endnote reference (using standard OOXML elements)
	// Note: only rStyle is needed in body references — the style itself provides superscript
	var refStyleVal string
	if noteType == FootnoteTypeFootnote {
		refStyleVal = "FootnoteReference"
	} else {
		refStyleVal = "EndnoteReference"
	}

	refRun := Run{
		Properties: &RunProperties{
			RunStyle: &RunStyle{Val: refStyleVal},
		},
	}

	if noteType == FootnoteTypeFootnote {
		refRun.FootnoteReference = &FootnoteReference{ID: noteID}
	} else {
		refRun.EndnoteReference = &EndnoteReference{ID: noteID}
	}

	paragraph.Runs = append(paragraph.Runs, refRun)
	d.Body.Elements = append(d.Body.Elements, paragraph)

	// Create footnote/endnote content
	if err := d.createNoteContent(noteID, noteText, noteType); err != nil {
		return fmt.Errorf("failed to create %s content: %w", noteType, err)
	}

	return nil
}

// AddFootnoteToRun adds a footnote reference to an existing Run (deprecated, use AddFootnoteToParagraph instead).
// Note: this method modifies the passed Run, appending a footnote reference marker after its text.
func (d *Document) AddFootnoteToRun(run *Run, footnoteText string) error {
	manager := d.getFootnoteManager()
	d.ensureFootnoteInitialized(FootnoteTypeFootnote)

	noteID := strconv.Itoa(manager.nextFootnoteID)
	manager.nextFootnoteID++

	// Set the Run to footnote reference style
	run.Properties = &RunProperties{
		RunStyle: &RunStyle{Val: "FootnoteReference"},
	}
	run.FootnoteReference = &FootnoteReference{ID: noteID}
	run.Text = Text{} // Footnote reference Run does not need text content

	// Create footnote content
	return d.createNoteContent(noteID, footnoteText, FootnoteTypeFootnote)
}

// AddFootnoteToParagraph appends a footnote reference Run at the end of a paragraph
func (d *Document) AddFootnoteToParagraph(para *Paragraph, footnoteText string) error {
	manager := d.getFootnoteManager()
	d.ensureFootnoteInitialized(FootnoteTypeFootnote)

	noteID := strconv.Itoa(manager.nextFootnoteID)
	manager.nextFootnoteID++

	// Create footnote reference Run (rStyle only — the style handles superscript)
	refRun := Run{
		Properties: &RunProperties{
			RunStyle: &RunStyle{Val: "FootnoteReference"},
		},
		FootnoteReference: &FootnoteReference{ID: noteID},
	}
	para.Runs = append(para.Runs, refRun)

	// Create footnote content
	return d.createNoteContent(noteID, footnoteText, FootnoteTypeFootnote)
}

// AddEndnoteToParagraph appends an endnote reference Run at the end of a paragraph
func (d *Document) AddEndnoteToParagraph(para *Paragraph, footnoteText string) error {
	manager := d.getFootnoteManager()
	d.ensureFootnoteInitialized(FootnoteTypeEndnote)

	noteID := strconv.Itoa(manager.nextEndnoteID)
	manager.nextEndnoteID++

	// Create endnote reference Run (rStyle only — the style handles superscript)
	refRun := Run{
		Properties: &RunProperties{
			RunStyle: &RunStyle{Val: "EndnoteReference"},
		},
		EndnoteReference: &EndnoteReference{ID: noteID},
	}
	para.Runs = append(para.Runs, refRun)

	// Create endnote content
	return d.createNoteContent(noteID, footnoteText, FootnoteTypeEndnote)
}

// SetFootnoteConfig sets the footnote configuration
func (d *Document) SetFootnoteConfig(config *FootnoteConfig) error {
	if config == nil {
		config = DefaultFootnoteConfig()
	}

	// Ensure document settings are initialized
	d.ensureSettingsInitialized()

	// Create footnote properties XML structure
	footnoteProps := &FootnoteProperties{
		NumberFormat: string(config.NumberFormat),
		StartNumber:  config.StartNumber,
		RestartRule:  string(config.RestartEach),
		Position:     string(config.Position),
	}

	// Create endnote properties XML structure
	endnoteProps := &EndnoteProperties{
		NumberFormat: string(config.NumberFormat),
		StartNumber:  config.StartNumber,
		RestartRule:  string(config.RestartEach),
		Position:     string(config.Position),
	}

	// Update document settings
	if err := d.updateDocumentSettings(footnoteProps, endnoteProps); err != nil {
		return fmt.Errorf("failed to update footnote configuration: %w", err)
	}

	return nil
}

// ensureFootnoteInitialized ensures the footnote/endnote system is initialized
func (d *Document) ensureFootnoteInitialized(noteType FootnoteType) {
	if noteType == FootnoteTypeFootnote {
		if _, exists := d.parts["word/footnotes.xml"]; !exists {
			d.initializeFootnotes()
		}
	} else {
		if _, exists := d.parts["word/endnotes.xml"]; !exists {
			d.initializeEndnotes()
		}
	}
}

// initializeNotes is a shared helper that initializes either the footnote or endnote system.
func (d *Document) initializeNotes(noteType FootnoteType) {
	separatorParagraphs := []*Paragraph{
		{
			Properties: &ParagraphProperties{
				Spacing: &Spacing{After: "0", Line: "240", LineRule: "auto"},
			},
			Runs: []Run{
				{Separator: &Separator{}},
			},
		},
	}
	continuationParagraphs := []*Paragraph{
		{
			Properties: &ParagraphProperties{
				Spacing: &Spacing{After: "0", Line: "240", LineRule: "auto"},
			},
			Runs: []Run{
				{ContinuationSeparator: &ContinuationSeparator{}},
			},
		},
	}

	var xmlBytes []byte
	var err error
	var partName, contentTypeName, relType, target string

	if noteType == FootnoteTypeFootnote {
		footnotes := &Footnotes{
			Xmlns: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
			Footnotes: []*Footnote{
				{Type: "separator", ID: "-1", Paragraphs: separatorParagraphs},
				{Type: "continuationSeparator", ID: "0", Paragraphs: continuationParagraphs},
			},
		}
		xmlBytes, err = xml.MarshalIndent(footnotes, "", "  ")
		partName = "word/footnotes.xml"
		contentTypeName = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
		relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
		target = "footnotes.xml"
	} else {
		endnotes := &Endnotes{
			Xmlns: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
			Endnotes: []*Endnote{
				{Type: "separator", ID: "-1", Paragraphs: separatorParagraphs},
				{Type: "continuationSeparator", ID: "0", Paragraphs: continuationParagraphs},
			},
		}
		xmlBytes, err = xml.MarshalIndent(endnotes, "", "  ")
		partName = "word/endnotes.xml"
		contentTypeName = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
		relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
		target = "endnotes.xml"
	}

	if err != nil {
		return
	}

	// Add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts[partName] = append(xmlDeclaration, xmlBytes...)

	// Add content type
	d.addContentType(partName, contentTypeName)

	// Add relationship
	relationshipID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2)
	relationship := Relationship{
		ID:     relationshipID,
		Type:   relType,
		Target: target,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)
}

// initializeFootnotes initializes the footnote system
func (d *Document) initializeFootnotes() {
	d.initializeNotes(FootnoteTypeFootnote)
}

// initializeEndnotes initializes the endnote system
func (d *Document) initializeEndnotes() {
	d.initializeNotes(FootnoteTypeEndnote)
}

// createNoteContent creates footnote/endnote content
func (d *Document) createNoteContent(noteID string, noteText string, noteType FootnoteType) error {
	manager := d.getFootnoteManager()

	// Determine paragraph style and self-reference element
	var pStyleVal string
	var refRunStyle string
	if noteType == FootnoteTypeFootnote {
		pStyleVal = "FootnoteText"
		refRunStyle = "FootnoteReference"
	} else {
		pStyleVal = "EndnoteText"
		refRunStyle = "EndnoteReference"
	}

	// Create self-reference Run (displays footnote/endnote number)
	selfRefRun := Run{
		Properties: &RunProperties{
			RunStyle:      &RunStyle{Val: refRunStyle},
			VerticalAlign: &VerticalAlignment{Val: "superscript"},
		},
	}
	if noteType == FootnoteTypeFootnote {
		selfRefRun.FootnoteRef = &FootnoteRef{}
	} else {
		selfRefRun.EndnoteRef = &EndnoteRef{}
	}

	// Create text content Run (with leading space)
	textRun := Run{
		Text: Text{Content: " " + noteText, Space: "preserve"},
	}

	// Create footnote/endnote paragraph (with paragraph style and self-reference + text Runs)
	noteParagraph := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: pStyleVal},
		},
		Runs: []Run{selfRefRun, textRun},
	}

	if noteType == FootnoteTypeFootnote {
		// Create footnote
		footnote := &Footnote{
			ID:         noteID,
			Paragraphs: []*Paragraph{noteParagraph},
		}
		manager.footnotes[noteID] = footnote

		// Update footnotes file
		d.updateFootnotesFile()
	} else {
		// Create endnote
		endnote := &Endnote{
			ID:         noteID,
			Paragraphs: []*Paragraph{noteParagraph},
		}
		manager.endnotes[noteID] = endnote

		// Update endnotes file
		d.updateEndnotesFile()
	}

	return nil
}

// updateNotesFile is a shared helper that updates either the footnotes or endnotes file.
func (d *Document) updateNotesFile(noteType FootnoteType) {
	manager := d.getFootnoteManager()

	separatorParagraphs := []*Paragraph{
		{
			Properties: &ParagraphProperties{
				Spacing: &Spacing{After: "0", Line: "240", LineRule: "auto"},
			},
			Runs: []Run{
				{Separator: &Separator{}},
			},
		},
	}
	continuationParagraphs := []*Paragraph{
		{
			Properties: &ParagraphProperties{
				Spacing: &Spacing{After: "0", Line: "240", LineRule: "auto"},
			},
			Runs: []Run{
				{ContinuationSeparator: &ContinuationSeparator{}},
			},
		},
	}

	var xmlBytes []byte
	var err error
	var partName string

	if noteType == FootnoteTypeFootnote {
		footnotes := &Footnotes{
			Xmlns: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
			Footnotes: []*Footnote{
				{Type: "separator", ID: "-1", Paragraphs: separatorParagraphs},
				{Type: "continuationSeparator", ID: "0", Paragraphs: continuationParagraphs},
			},
		}
		for _, footnote := range manager.footnotes {
			footnotes.Footnotes = append(footnotes.Footnotes, footnote)
		}
		xmlBytes, err = xml.MarshalIndent(footnotes, "", "  ")
		partName = "word/footnotes.xml"
	} else {
		endnotes := &Endnotes{
			Xmlns: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
			Endnotes: []*Endnote{
				{Type: "separator", ID: "-1", Paragraphs: separatorParagraphs},
				{Type: "continuationSeparator", ID: "0", Paragraphs: continuationParagraphs},
			},
		}
		for _, endnote := range manager.endnotes {
			endnotes.Endnotes = append(endnotes.Endnotes, endnote)
		}
		xmlBytes, err = xml.MarshalIndent(endnotes, "", "  ")
		partName = "word/endnotes.xml"
	}

	if err != nil {
		return
	}

	// Add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts[partName] = append(xmlDeclaration, xmlBytes...)
}

// updateFootnotesFile updates the footnotes file
func (d *Document) updateFootnotesFile() {
	d.updateNotesFile(FootnoteTypeFootnote)
}

// updateEndnotesFile updates the endnotes file
func (d *Document) updateEndnotesFile() {
	d.updateNotesFile(FootnoteTypeEndnote)
}

// GetFootnoteCount returns the number of footnotes
func (d *Document) GetFootnoteCount() int {
	manager := d.getFootnoteManager()
	return len(manager.footnotes)
}

// GetEndnoteCount returns the number of endnotes
func (d *Document) GetEndnoteCount() int {
	manager := d.getFootnoteManager()
	return len(manager.endnotes)
}

// RemoveFootnote removes the specified footnote
func (d *Document) RemoveFootnote(footnoteID string) error {
	manager := d.getFootnoteManager()

	if _, exists := manager.footnotes[footnoteID]; !exists {
		return fmt.Errorf("footnote %s does not exist", footnoteID)
	}

	delete(manager.footnotes, footnoteID)
	d.updateFootnotesFile()

	return nil
}

// RemoveEndnote removes the specified endnote
func (d *Document) RemoveEndnote(endnoteID string) error {
	manager := d.getFootnoteManager()

	if _, exists := manager.endnotes[endnoteID]; !exists {
		return fmt.Errorf("endnote %s does not exist", endnoteID)
	}

	delete(manager.endnotes, endnoteID)
	d.updateEndnotesFile()

	return nil
}

// ensureSettingsInitialized ensures document settings are initialized
func (d *Document) ensureSettingsInitialized() {
	// Check if settings.xml exists; if not, create default settings
	if _, exists := d.parts["word/settings.xml"]; !exists {
		d.initializeSettings()
	}
}

// initializeSettings initializes the document settings
func (d *Document) initializeSettings() {
	// Create default settings
	settings := d.createDefaultSettings()

	// Save settings
	if err := d.saveSettings(settings); err != nil {
		// If saving fails, use the hardcoded fallback
		settingsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="708"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>`
		d.parts["word/settings.xml"] = []byte(settingsXML)
	}

	// Add content type
	d.addContentType("word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")

	// Add relationship
	d.addSettingsRelationship()
}

// updateDocumentSettings updates the footnote and endnote configuration in document settings
func (d *Document) updateDocumentSettings(footnoteProps *FootnoteProperties, endnoteProps *EndnoteProperties) error {
	// Parse existing settings.xml
	settings, err := d.parseSettings()
	if err != nil {
		return fmt.Errorf("failed to parse settings file: %w", err)
	}

	// Update footnote settings
	if footnoteProps != nil {
		footnotePr := &FootnotePr{}

		if footnoteProps.NumberFormat != "" {
			footnotePr.NumFmt = &FootnoteNumFmt{Val: footnoteProps.NumberFormat}
		}

		if footnoteProps.StartNumber > 0 {
			footnotePr.NumStart = &FootnoteNumStart{Val: strconv.Itoa(footnoteProps.StartNumber)}
		}

		if footnoteProps.RestartRule != "" {
			footnotePr.NumRestart = &FootnoteNumRestart{Val: footnoteProps.RestartRule}
		}

		if footnoteProps.Position != "" {
			footnotePr.Pos = &FootnotePos{Val: footnoteProps.Position}
		}

		settings.FootnotePr = footnotePr
	}

	// Update endnote settings
	if endnoteProps != nil {
		endnotePr := &EndnotePr{}

		if endnoteProps.NumberFormat != "" {
			endnotePr.NumFmt = &EndnoteNumFmt{Val: endnoteProps.NumberFormat}
		}

		if endnoteProps.StartNumber > 0 {
			endnotePr.NumStart = &EndnoteNumStart{Val: strconv.Itoa(endnoteProps.StartNumber)}
		}

		if endnoteProps.RestartRule != "" {
			endnotePr.NumRestart = &EndnoteNumRestart{Val: endnoteProps.RestartRule}
		}

		if endnoteProps.Position != "" {
			endnotePr.Pos = &EndnotePos{Val: endnoteProps.Position}
		}

		settings.EndnotePr = endnotePr
	}

	// Save the updated settings.xml
	return d.saveSettings(settings)
}

// parseSettings parses the settings.xml file
func (d *Document) parseSettings() (*Settings, error) {
	settingsData, exists := d.parts["word/settings.xml"]
	if !exists {
		// If settings.xml does not exist, return default settings
		return d.createDefaultSettings(), nil
	}

	var settings Settings

	// Using xml.Unmarshal directly may have namespace issues, so we use a
	// string replacement approach instead. We replace w:settings with settings, etc.,
	// then parse with a simplified structure.
	settingsStr := string(settingsData)

	// If the XML contains w: prefix, it is serialized XML; create default settings and update.
	// This is a simplified approach to avoid namespace parsing issues.
	if len(settingsStr) > 0 {
		// If the file exists and is not empty, use default settings as a base
		settings = *d.createDefaultSettings()

		// More complex XML parsing logic can be added here later.
		// For now, use simplified handling and return default settings.
		return &settings, nil
	}

	return d.createDefaultSettings(), nil
}

// createDefaultSettings creates default settings
func (d *Document) createDefaultSettings() *Settings {
	return &Settings{
		Xmlns: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		DefaultTabStop: &DefaultTabStop{
			Val: "708",
		},
		CharacterSpacingControl: &CharacterSpacingControl{
			Val: "doNotCompress",
		},
	}
}

// saveSettings saves the settings.xml file
func (d *Document) saveSettings(settings *Settings) error {
	// Serialize to XML
	settingsXML, err := xml.MarshalIndent(settings, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize settings.xml: %w", err)
	}

	// Add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/settings.xml"] = append(xmlDeclaration, settingsXML...)

	return nil
}

// addSettingsRelationship adds the settings file relationship
func (d *Document) addSettingsRelationship() {
	relationshipID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2)

	relationship := Relationship{
		ID:     relationshipID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
		Target: "settings.xml",
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)
}
