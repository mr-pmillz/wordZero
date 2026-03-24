// Package document provides Word document property operations.
package document

import (
	"encoding/xml"
	"fmt"
	"strings"
	"time"
)

// DocumentProperties represents the document properties structure.
type DocumentProperties struct {
	// Core properties
	Title       string // Document title
	Subject     string // Document subject
	Creator     string // Creator
	Keywords    string // Keywords
	Description string // Description
	Language    string // Language
	Category    string // Category
	Version     string // Version
	Revision    string // Revision

	// Time properties
	Created      time.Time // Creation time
	LastModified time.Time // Last modified time
	LastPrinted  time.Time // Last printed time

	// Statistical properties
	Pages      int // Page count
	Words      int // Word count
	Characters int // Character count
	Paragraphs int // Paragraph count
	Lines      int // Line count
}

// CoreProperties is the core properties XML structure.
type CoreProperties struct {
	XMLName       xml.Name `xml:"cp:coreProperties"`
	XmlnsCP       string   `xml:"xmlns:cp,attr"`
	XmlnsDC       string   `xml:"xmlns:dc,attr"`
	XmlnsDCTerms  string   `xml:"xmlns:dcterms,attr"`
	XmlnsDCMIType string   `xml:"xmlns:dcmitype,attr"`
	XmlnsXSI      string   `xml:"xmlns:xsi,attr"`
	Title         *DCText  `xml:"dc:title,omitempty"`
	Subject       *DCText  `xml:"dc:subject,omitempty"`
	Creator       *DCText  `xml:"dc:creator,omitempty"`
	Keywords      *CPText  `xml:"cp:keywords,omitempty"`
	Description   *DCText  `xml:"dc:description,omitempty"`
	Language      *DCText  `xml:"dc:language,omitempty"`
	Category      *CPText  `xml:"cp:category,omitempty"`
	Version       *CPText  `xml:"cp:version,omitempty"`
	Revision      *CPText  `xml:"cp:revision,omitempty"`
	Created       *DCDate  `xml:"dcterms:created,omitempty"`
	Modified      *DCDate  `xml:"dcterms:modified,omitempty"`
	LastPrinted   *DCDate  `xml:"cp:lastPrinted,omitempty"`
}

// AppProperties is the application properties XML structure.
type AppProperties struct {
	XMLName       xml.Name `xml:"Properties"`
	Xmlns         string   `xml:"xmlns,attr"`
	XmlnsVT       string   `xml:"xmlns:vt,attr"`
	Application   string   `xml:"Application,omitempty"`
	DocSecurity   int      `xml:"DocSecurity,omitempty"`
	ScaleCrop     bool     `xml:"ScaleCrop,omitempty"`
	LinksUpToDate bool     `xml:"LinksUpToDate,omitempty"`
	Pages         int      `xml:"Pages,omitempty"`
	Words         int      `xml:"Words,omitempty"`
	Characters    int      `xml:"Characters,omitempty"`
	Paragraphs    int      `xml:"Paragraphs,omitempty"`
	Lines         int      `xml:"Lines,omitempty"`
}

// DCText is a DC namespace text element.
type DCText struct {
	Text string `xml:",chardata"`
}

// CPText is a CP namespace text element.
type CPText struct {
	Text string `xml:",chardata"`
}

// DCDate is a DC namespace date element.
type DCDate struct {
	XSIType string    `xml:"xsi:type,attr"`
	Date    time.Time `xml:",chardata"`
}

// SetDocumentProperties sets the document properties.
func (d *Document) SetDocumentProperties(properties *DocumentProperties) error {
	if properties == nil {
		return fmt.Errorf("document properties cannot be nil")
	}

	// Generate core properties XML
	if err := d.generateCoreProperties(properties); err != nil {
		return fmt.Errorf("failed to generate core properties: %w", err)
	}

	// Generate application properties XML
	if err := d.generateAppProperties(properties); err != nil {
		return fmt.Errorf("failed to generate application properties: %w", err)
	}

	// Add content types and relationships
	d.addPropertiesContentTypes()
	d.addPropertiesRelationships()

	return nil
}

// GetDocumentProperties retrieves the document properties.
func (d *Document) GetDocumentProperties() (*DocumentProperties, error) {
	properties := &DocumentProperties{
		Created:      time.Now(),
		LastModified: time.Now(),
		Language:     "zh-CN",
	}

	// Read from saved properties if they exist
	if coreData, exists := d.parts["docProps/core.xml"]; exists {
		if err := d.parseCoreProperties(coreData, properties); err != nil {
			return nil, fmt.Errorf("failed to parse core properties: %w", err)
		}
	}

	if appData, exists := d.parts["docProps/app.xml"]; exists {
		if err := d.parseAppProperties(appData, properties); err != nil {
			return nil, fmt.Errorf("failed to parse application properties: %w", err)
		}
	}

	return properties, nil
}

// SetTitle sets the document title.
func (d *Document) SetTitle(title string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Title = title
	return d.SetDocumentProperties(properties)
}

// SetAuthor sets the document author.
func (d *Document) SetAuthor(author string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Creator = author
	return d.SetDocumentProperties(properties)
}

// SetSubject sets the document subject.
func (d *Document) SetSubject(subject string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Subject = subject
	return d.SetDocumentProperties(properties)
}

// SetKeywords sets the document keywords.
func (d *Document) SetKeywords(keywords string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Keywords = keywords
	return d.SetDocumentProperties(properties)
}

// SetDescription sets the document description.
func (d *Document) SetDescription(description string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Description = description
	return d.SetDocumentProperties(properties)
}

// SetCategory sets the document category.
func (d *Document) SetCategory(category string) error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}
	properties.Category = category
	return d.SetDocumentProperties(properties)
}

// UpdateStatistics updates the document statistics.
func (d *Document) UpdateStatistics() error {
	properties, err := d.GetDocumentProperties()
	if err != nil {
		properties = &DocumentProperties{}
	}

	// Calculate statistics
	properties.Paragraphs = len(d.Body.GetParagraphs())
	properties.Words = d.countWords()
	properties.Characters = d.countCharacters()
	properties.Lines = d.countLines()
	properties.Pages = 1 // Simplified; actual implementation requires complex calculation

	// Update last modified time
	properties.LastModified = time.Now()

	return d.SetDocumentProperties(properties)
}

// generateCoreProperties generates the core properties XML.
func (d *Document) generateCoreProperties(properties *DocumentProperties) error {
	coreProps := &CoreProperties{
		XmlnsCP:       "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
		XmlnsDC:       "http://purl.org/dc/elements/1.1/",
		XmlnsDCTerms:  "http://purl.org/dc/terms/",
		XmlnsDCMIType: "http://purl.org/dc/dcmitype/",
		XmlnsXSI:      "http://www.w3.org/2001/XMLSchema-instance",
	}

	// Set property values
	if properties.Title != "" {
		coreProps.Title = &DCText{Text: properties.Title}
	}
	if properties.Subject != "" {
		coreProps.Subject = &DCText{Text: properties.Subject}
	}
	if properties.Creator != "" {
		coreProps.Creator = &DCText{Text: properties.Creator}
	}
	if properties.Keywords != "" {
		coreProps.Keywords = &CPText{Text: properties.Keywords}
	}
	if properties.Description != "" {
		coreProps.Description = &DCText{Text: properties.Description}
	}
	if properties.Language != "" {
		coreProps.Language = &DCText{Text: properties.Language}
	}
	if properties.Category != "" {
		coreProps.Category = &CPText{Text: properties.Category}
	}
	if properties.Version != "" {
		coreProps.Version = &CPText{Text: properties.Version}
	}
	if properties.Revision != "" {
		coreProps.Revision = &CPText{Text: properties.Revision}
	}

	// Set time properties
	if !properties.Created.IsZero() {
		coreProps.Created = &DCDate{
			XSIType: "dcterms:W3CDTF",
			Date:    properties.Created,
		}
	}
	if !properties.LastModified.IsZero() {
		coreProps.Modified = &DCDate{
			XSIType: "dcterms:W3CDTF",
			Date:    properties.LastModified,
		}
	}
	if !properties.LastPrinted.IsZero() {
		coreProps.LastPrinted = &DCDate{
			XSIType: "dcterms:W3CDTF",
			Date:    properties.LastPrinted,
		}
	}

	// Serialize to XML
	coreXML, err := xml.MarshalIndent(coreProps, "", "  ")
	if err != nil {
		return err
	}

	// Add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["docProps/core.xml"] = append(xmlDeclaration, coreXML...)

	return nil
}

// generateAppProperties generates the application properties XML.
func (d *Document) generateAppProperties(properties *DocumentProperties) error {
	appProps := &AppProperties{
		Xmlns:         "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
		XmlnsVT:       "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
		Application:   "WordZero/1.0",
		DocSecurity:   0,
		ScaleCrop:     false,
		LinksUpToDate: false,
		Pages:         properties.Pages,
		Words:         properties.Words,
		Characters:    properties.Characters,
		Paragraphs:    properties.Paragraphs,
		Lines:         properties.Lines,
	}

	// Serialize to XML
	appXML, err := xml.MarshalIndent(appProps, "", "  ")
	if err != nil {
		return err
	}

	// Add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["docProps/app.xml"] = append(xmlDeclaration, appXML...)

	return nil
}

// parseCoreProperties parses the core properties.
func (d *Document) parseCoreProperties(data []byte, properties *DocumentProperties) error {
	var coreProps CoreProperties
	if err := xml.Unmarshal(data, &coreProps); err != nil {
		return err
	}

	if coreProps.Title != nil {
		properties.Title = coreProps.Title.Text
	}
	if coreProps.Subject != nil {
		properties.Subject = coreProps.Subject.Text
	}
	if coreProps.Creator != nil {
		properties.Creator = coreProps.Creator.Text
	}
	if coreProps.Keywords != nil {
		properties.Keywords = coreProps.Keywords.Text
	}
	if coreProps.Description != nil {
		properties.Description = coreProps.Description.Text
	}
	if coreProps.Language != nil {
		properties.Language = coreProps.Language.Text
	}
	if coreProps.Category != nil {
		properties.Category = coreProps.Category.Text
	}
	if coreProps.Version != nil {
		properties.Version = coreProps.Version.Text
	}
	if coreProps.Revision != nil {
		properties.Revision = coreProps.Revision.Text
	}

	if coreProps.Created != nil {
		properties.Created = coreProps.Created.Date
	}
	if coreProps.Modified != nil {
		properties.LastModified = coreProps.Modified.Date
	}
	if coreProps.LastPrinted != nil {
		properties.LastPrinted = coreProps.LastPrinted.Date
	}

	return nil
}

// parseAppProperties parses the application properties.
func (d *Document) parseAppProperties(data []byte, properties *DocumentProperties) error {
	var appProps AppProperties
	if err := xml.Unmarshal(data, &appProps); err != nil {
		return err
	}

	properties.Pages = appProps.Pages
	properties.Words = appProps.Words
	properties.Characters = appProps.Characters
	properties.Paragraphs = appProps.Paragraphs
	properties.Lines = appProps.Lines

	return nil
}

// addPropertiesContentTypes adds property-related content types.
func (d *Document) addPropertiesContentTypes() {
	d.addContentType("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml")
	d.addContentType("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")
}

// addPropertiesRelationships adds property-related relationships.
func (d *Document) addPropertiesRelationships() {
	// These relationships are typically defined in the package-level _rels/.rels.
	// Simplified here; actual implementation needs to manage package-level relationships.
}

// countWords counts the number of words.
func (d *Document) countWords() int {
	count := 0
	for _, paragraph := range d.Body.GetParagraphs() {
		for _, run := range paragraph.Runs {
			// Simplified counting by splitting on whitespace
			words := len(strings.Fields(run.Text.Content))
			count += words
		}
	}
	return count
}

// countCharacters counts the number of characters.
func (d *Document) countCharacters() int {
	count := 0
	for _, paragraph := range d.Body.GetParagraphs() {
		for _, run := range paragraph.Runs {
			count += len(run.Text.Content)
		}
	}
	return count
}

// countLines counts the number of lines.
func (d *Document) countLines() int {
	count := 0
	for _, paragraph := range d.Body.GetParagraphs() {
		for _, run := range paragraph.Runs {
			// Simplified counting by newline characters
			lines := strings.Count(run.Text.Content, "\n") + 1
			count += lines
		}
	}
	return count
}
