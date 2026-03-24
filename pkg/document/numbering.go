// Package document provides list and numbering operations for Word documents.
package document

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

// ListType represents a list type.
type ListType string

const (
	// ListTypeBullet represents an unordered (bullet) list.
	ListTypeBullet ListType = "bullet"
	// ListTypeNumber represents an ordered (numbered) list.
	ListTypeNumber ListType = "number"
	// ListTypeDecimal represents decimal numbering.
	ListTypeDecimal ListType = "decimal"
	// ListTypeLowerLetter represents lowercase letter numbering.
	ListTypeLowerLetter ListType = "lowerLetter"
	// ListTypeUpperLetter represents uppercase letter numbering.
	ListTypeUpperLetter ListType = "upperLetter"
	// ListTypeLowerRoman represents lowercase Roman numeral numbering.
	ListTypeLowerRoman ListType = "lowerRoman"
	// ListTypeUpperRoman represents uppercase Roman numeral numbering.
	ListTypeUpperRoman ListType = "upperRoman"
)

// BulletType represents a bullet symbol type.
type BulletType string

const (
	// BulletTypeDot represents a filled circle bullet.
	BulletTypeDot BulletType = "•"
	// BulletTypeCircle represents a hollow circle bullet.
	BulletTypeCircle BulletType = "○"
	// BulletTypeSquare represents a square bullet.
	BulletTypeSquare BulletType = "■"
	// BulletTypeDash represents a dash bullet.
	BulletTypeDash BulletType = "–"
	// BulletTypeArrow represents an arrow bullet.
	BulletTypeArrow BulletType = "→"
)

// Numbering represents the numbering definitions for a document.
type Numbering struct {
	XMLName            xml.Name       `xml:"w:numbering"`
	Xmlns              string         `xml:"xmlns:w,attr"`
	AbstractNums       []*AbstractNum `xml:"w:abstractNum"`
	NumberingInstances []*NumInstance `xml:"w:num"`
}

// AbstractNum represents an abstract numbering definition.
type AbstractNum struct {
	XMLName       xml.Name `xml:"w:abstractNum"`
	AbstractNumID string   `xml:"w:abstractNumId,attr"`
	Levels        []*Level `xml:"w:lvl"`
}

// NumInstance represents a numbering instance.
type NumInstance struct {
	XMLName       xml.Name              `xml:"w:num"`
	NumID         string                `xml:"w:numId,attr"`
	AbstractNumID *AbstractNumReference `xml:"w:abstractNumId"`
}

// AbstractNumReference represents a reference to an abstract numbering definition.
type AbstractNumReference struct {
	XMLName xml.Name `xml:"w:abstractNumId"`
	Val     string   `xml:"w:val,attr"`
}

// Level represents a numbering level definition.
type Level struct {
	XMLName   xml.Name   `xml:"w:lvl"`
	ILevel    string     `xml:"w:ilvl,attr"`
	Start     *Start     `xml:"w:start,omitempty"`
	NumFmt    *NumFmt    `xml:"w:numFmt,omitempty"`
	LevelText *LevelText `xml:"w:lvlText,omitempty"`
	LevelJc   *LevelJc   `xml:"w:lvlJc,omitempty"`
	PPr       *LevelPPr  `xml:"w:pPr,omitempty"`
	RPr       *LevelRPr  `xml:"w:rPr,omitempty"`
}

// Start represents the starting number for a numbering level.
type Start struct {
	XMLName xml.Name `xml:"w:start"`
	Val     string   `xml:"w:val,attr"`
}

// NumFmt represents a number format definition.
type NumFmt struct {
	XMLName xml.Name `xml:"w:numFmt"`
	Val     string   `xml:"w:val,attr"`
}

// LevelText represents the text format for a numbering level.
type LevelText struct {
	XMLName xml.Name `xml:"w:lvlText"`
	Val     string   `xml:"w:val,attr"`
}

// LevelJc represents the justification for a numbering level.
type LevelJc struct {
	XMLName xml.Name `xml:"w:lvlJc"`
	Val     string   `xml:"w:val,attr"`
}

// LevelPPr represents paragraph properties for a numbering level.
type LevelPPr struct {
	XMLName xml.Name     `xml:"w:pPr"`
	Ind     *LevelIndent `xml:"w:ind,omitempty"`
}

// LevelIndent represents indentation settings for a numbering level.
type LevelIndent struct {
	XMLName xml.Name `xml:"w:ind"`
	Left    string   `xml:"w:left,attr,omitempty"`
	Hanging string   `xml:"w:hanging,attr,omitempty"`
}

// LevelRPr represents run properties for a numbering level.
type LevelRPr struct {
	XMLName    xml.Name    `xml:"w:rPr"`
	FontFamily *FontFamily `xml:"w:rFonts,omitempty"`
}

// ListConfig holds list configuration options.
type ListConfig struct {
	Type         ListType   // List type
	BulletSymbol BulletType // Bullet symbol (only for unordered lists)
	StartNumber  int        // Starting number (only for ordered lists)
	IndentLevel  int        // Indentation level (0-8)
}

// NumberingManager manages numbering definitions for lists.
type NumberingManager struct {
	nextAbstractNumID int
	nextNumID         int
	abstractNums      map[string]*AbstractNum
	numInstances      map[string]*NumInstance
}

// getNumberingManager returns the document's numbering manager (lazy init).
func (d *Document) getNumberingManager() *NumberingManager {
	if d.numberingManager == nil {
		d.numberingManager = &NumberingManager{
			nextAbstractNumID: 0,
			nextNumID:         1,
			abstractNums:      make(map[string]*AbstractNum),
			numInstances:      make(map[string]*NumInstance),
		}
	}
	return d.numberingManager
}

// AddListItem adds a list item to the document.
func (d *Document) AddListItem(text string, config *ListConfig) *Paragraph {
	if config == nil {
		config = &ListConfig{
			Type:         ListTypeBullet,
			BulletSymbol: BulletTypeDot,
			IndentLevel:  0,
		}
	}

	// Ensure the numbering manager is initialized
	d.ensureNumberingInitialized()

	// Get or create a numbering definition
	numID := d.getOrCreateNumbering(config)

	// Create the paragraph
	paragraph := &Paragraph{
		Properties: &ParagraphProperties{
			NumberingProperties: &NumberingProperties{
				ILevel: &ILevel{Val: strconv.Itoa(config.IndentLevel)},
				NumID:  &NumID{Val: numID},
			},
		},
	}

	// Add text content
	if text != "" {
		run := Run{
			Text: Text{
				Content: text,
			},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}

	// Append to the document
	d.Body.Elements = append(d.Body.Elements, paragraph)
	return paragraph
}

// AddBulletList adds an unordered (bullet) list item to the document.
func (d *Document) AddBulletList(text string, level int, bulletType BulletType) *Paragraph {
	config := &ListConfig{
		Type:         ListTypeBullet,
		BulletSymbol: bulletType,
		IndentLevel:  level,
	}
	return d.AddListItem(text, config)
}

// AddNumberedList adds an ordered (numbered) list item to the document.
func (d *Document) AddNumberedList(text string, level int, numType ListType) *Paragraph {
	config := &ListConfig{
		Type:        numType,
		IndentLevel: level,
		StartNumber: 1,
	}
	return d.AddListItem(text, config)
}

// CreateMultiLevelList creates a multi-level list from the given items.
func (d *Document) CreateMultiLevelList(items []ListItem) error {
	for _, item := range items {
		config := &ListConfig{
			Type:         item.Type,
			BulletSymbol: item.BulletSymbol,
			IndentLevel:  item.Level,
			StartNumber:  item.StartNumber,
		}
		d.AddListItem(item.Text, config)
	}
	return nil
}

// ListItem represents a list item structure.
type ListItem struct {
	Text         string     // Text content
	Level        int        // Indentation level
	Type         ListType   // List type
	BulletSymbol BulletType // Bullet symbol
	StartNumber  int        // Starting number
}

// ensureNumberingInitialized ensures the numbering system is initialized.
func (d *Document) ensureNumberingInitialized() {
	// Check if numbering definitions already exist
	if _, exists := d.parts["word/numbering.xml"]; !exists {
		d.initializeNumbering()
	}
}

// initializeNumbering initializes the numbering system.
func (d *Document) initializeNumbering() {
	numbering := &Numbering{
		Xmlns:              "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		AbstractNums:       []*AbstractNum{},
		NumberingInstances: []*NumInstance{},
	}

	// Serialize the numbering definitions
	numberingXML, err := xml.MarshalIndent(numbering, "", "  ")
	if err != nil {
		return
	}

	// Add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/numbering.xml"] = append(xmlDeclaration, numberingXML...)

	// Add content type
	d.addContentType("word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")

	// Add relationship
	d.addNumberingRelationship()
}

// getOrCreateNumbering gets or creates a numbering definition.
func (d *Document) getOrCreateNumbering(config *ListConfig) string {
	manager := d.getNumberingManager()

	// Generate abstract numbering key
	abstractKey := fmt.Sprintf("%s_%s_%d", config.Type, config.BulletSymbol, config.IndentLevel)

	// Check if an abstract numbering already exists
	var abstractNum *AbstractNum
	if existing, exists := manager.abstractNums[abstractKey]; exists {
		abstractNum = existing
	} else {
		// Create a new abstract numbering
		abstractNumID := strconv.Itoa(manager.nextAbstractNumID)
		manager.nextAbstractNumID++

		abstractNum = d.createAbstractNum(abstractNumID, config)
		manager.abstractNums[abstractKey] = abstractNum
	}

	// Create a numbering instance
	numID := strconv.Itoa(manager.nextNumID)
	manager.nextNumID++

	numInstance := &NumInstance{
		NumID: numID,
		AbstractNumID: &AbstractNumReference{
			Val: abstractNum.AbstractNumID,
		},
	}
	manager.numInstances[numID] = numInstance

	// Update the numbering definition file
	d.updateNumberingFile()

	return numID
}

// createAbstractNum creates an abstract numbering definition.
func (d *Document) createAbstractNum(abstractNumID string, config *ListConfig) *AbstractNum {
	abstractNum := &AbstractNum{
		AbstractNumID: abstractNumID,
		Levels:        []*Level{},
	}

	// Create multiple levels (supports 9-level lists)
	for i := 0; i <= 8; i++ {
		level := d.createLevel(i, config)
		abstractNum.Levels = append(abstractNum.Levels, level)
	}

	return abstractNum
}

// createLevel creates a numbering level definition.
func (d *Document) createLevel(levelIndex int, config *ListConfig) *Level {
	level := &Level{
		ILevel:  strconv.Itoa(levelIndex),
		Start:   &Start{Val: strconv.Itoa(config.StartNumber)},
		LevelJc: &LevelJc{Val: "left"},
		PPr: &LevelPPr{
			Ind: &LevelIndent{
				Left:    strconv.Itoa((levelIndex + 1) * 720), // 720 twips = 0.5 inch
				Hanging: "360",                                // 360 twips = 0.25 inch
			},
		},
	}

	// Set number format and text
	switch config.Type {
	case ListTypeBullet:
		level.NumFmt = &NumFmt{Val: "bullet"}
		level.LevelText = &LevelText{Val: string(config.BulletSymbol)}
		level.RPr = &LevelRPr{
			FontFamily: &FontFamily{ASCII: "Symbol"},
		}
	case ListTypeNumber, ListTypeDecimal:
		level.NumFmt = &NumFmt{Val: "decimal"}
		level.LevelText = &LevelText{Val: fmt.Sprintf("%%%d.", levelIndex+1)}
	case ListTypeLowerLetter:
		level.NumFmt = &NumFmt{Val: "lowerLetter"}
		level.LevelText = &LevelText{Val: fmt.Sprintf("%%%d.", levelIndex+1)}
	case ListTypeUpperLetter:
		level.NumFmt = &NumFmt{Val: "upperLetter"}
		level.LevelText = &LevelText{Val: fmt.Sprintf("%%%d.", levelIndex+1)}
	case ListTypeLowerRoman:
		level.NumFmt = &NumFmt{Val: "lowerRoman"}
		level.LevelText = &LevelText{Val: fmt.Sprintf("%%%d.", levelIndex+1)}
	case ListTypeUpperRoman:
		level.NumFmt = &NumFmt{Val: "upperRoman"}
		level.LevelText = &LevelText{Val: fmt.Sprintf("%%%d.", levelIndex+1)}
	}

	return level
}

// updateNumberingFile updates the numbering definition file.
func (d *Document) updateNumberingFile() {
	manager := d.getNumberingManager()

	numbering := &Numbering{
		Xmlns:              "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		AbstractNums:       []*AbstractNum{},
		NumberingInstances: []*NumInstance{},
	}

	// Add all abstract numbering definitions
	for _, abstractNum := range manager.abstractNums {
		numbering.AbstractNums = append(numbering.AbstractNums, abstractNum)
	}

	// Add all numbering instances
	for _, numInstance := range manager.numInstances {
		numbering.NumberingInstances = append(numbering.NumberingInstances, numInstance)
	}

	// Serialize
	numberingXML, err := xml.MarshalIndent(numbering, "", "  ")
	if err != nil {
		return
	}

	// Add XML declaration
	xmlDeclaration := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	d.parts["word/numbering.xml"] = append(xmlDeclaration, numberingXML...)
}

// addNumberingRelationship adds the numbering relationship to the document.
func (d *Document) addNumberingRelationship() {
	// Generate relationship ID
	relationshipID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2 because styles.xml is already defined

	// Add relationship
	relationship := Relationship{
		ID:     relationshipID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
		Target: "numbering.xml",
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)
}

// RestartNumbering restarts the numbering sequence.
func (d *Document) RestartNumbering(numID string) {
	// Reset the numbering counter
	// In a full implementation, a new numbering instance would be created to reset the count
	manager := d.getNumberingManager()

	// Create a new numbering instance
	newNumID := strconv.Itoa(manager.nextNumID)
	manager.nextNumID++

	// If the original instance exists, copy its abstract numbering reference
	if existing, exists := manager.numInstances[numID]; exists {
		newInstance := &NumInstance{
			NumID: newNumID,
			AbstractNumID: &AbstractNumReference{
				Val: existing.AbstractNumID.Val,
			},
		}
		manager.numInstances[newNumID] = newInstance
		d.updateNumberingFile()
	}
}
