package document

import (
	"fmt"
	"regexp"
	"strings"
)

// TemplateRenderer is a dedicated engine for template rendering.
type TemplateRenderer struct {
	engine *TemplateEngine
	logger *TemplateLogger
}

// TemplateLogger is a logger for template operations.
type TemplateLogger struct {
	enabled bool
}

// NewTemplateRenderer creates a new template renderer.
func NewTemplateRenderer() *TemplateRenderer {
	return &TemplateRenderer{
		engine: NewTemplateEngine(),
		logger: &TemplateLogger{enabled: true},
	}
}

// SetLogging enables or disables logging.
func (tr *TemplateRenderer) SetLogging(enabled bool) {
	tr.logger.enabled = enabled
}

// logInfof logs an informational message.
func (tr *TemplateRenderer) logInfof(format string, args ...interface{}) {
	if tr.logger.enabled {
		Infof("[TemplateEngine] "+format, args...)
	}
}

// logErrorf logs an error message.
func (tr *TemplateRenderer) logErrorf(format string, args ...interface{}) {
	if tr.logger.enabled {
		Errorf("[TemplateEngine] "+format, args...)
	}
}

// LoadTemplateFromFile loads a template from a file.
func (tr *TemplateRenderer) LoadTemplateFromFile(name, filePath string) (*Template, error) {
	doc, err := Open(filePath)
	if err != nil {
		tr.logErrorf("failed to open template file %s: %v", filePath, err)
		return nil, err
	}

	template, err := tr.engine.LoadTemplateFromDocument(name, doc)
	if err != nil {
		tr.logErrorf("failed to load template from document %s: %v", name, err)
		return nil, err
	}

	tr.logInfof("successfully loaded template: %s (source: %s)", name, filePath)
	return template, nil
}

// RenderTemplate renders a template to a new document.
func (tr *TemplateRenderer) RenderTemplate(templateName string, data *TemplateData) (*Document, error) {
	tr.logInfof("starting template rendering: %s", templateName)

	// Check data integrity first
	if err := tr.validateTemplateData(data); err != nil {
		tr.logErrorf("template data validation failed: %v", err)
		return nil, err
	}

	// Render template
	doc, err := tr.engine.RenderTemplateToDocument(templateName, data)
	if err != nil {
		tr.logErrorf("template rendering failed: %v", err)
		return nil, err
	}

	tr.logInfof("template rendering completed: %s", templateName)
	return doc, nil
}

// validateTemplateData validates the template data.
func (tr *TemplateRenderer) validateTemplateData(data *TemplateData) error {
	if data == nil {
		return fmt.Errorf("template data must not be nil")
	}

	// Log data statistics
	tr.logInfof("template data stats: variables=%d, lists=%d, conditions=%d",
		len(data.Variables), len(data.Lists), len(data.Conditions))

	// Check list data format
	for listName, listData := range data.Lists {
		if len(listData) == 0 {
			tr.logInfof("warning: list '%s' is empty", listName)
			continue
		}

		// Check list item format consistency
		for i, item := range listData {
			if itemMap, ok := item.(map[string]interface{}); ok {
				if i == 0 {
					tr.logInfof("list '%s' contains %d items, first item fields: %v",
						listName, len(listData), tr.getMapKeys(itemMap))
				}
			} else {
				tr.logInfof("list '%s' item %d is not a map type: %T", listName, i, item)
			}
		}
	}

	return nil
}

// getMapKeys returns the keys of a map.
func (tr *TemplateRenderer) getMapKeys(m map[string]interface{}) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, k)
	}
	return keys
}

// AnalyzeTemplate analyzes the structure of a template.
func (tr *TemplateRenderer) AnalyzeTemplate(templateName string) (*TemplateAnalysis, error) {
	template, err := tr.engine.GetTemplate(templateName)
	if err != nil {
		return nil, err
	}

	analysis := &TemplateAnalysis{
		TemplateName: templateName,
		Variables:    make(map[string]bool),
		Lists:        make(map[string]bool),
		Conditions:   make(map[string]bool),
		Tables:       make([]*TableAnalysis, 0),
	}

	// Analyze base document
	if template.BaseDoc != nil {
		tr.analyzeDocument(template.BaseDoc, analysis)
	}

	tr.logInfof("template analysis completed: %s", templateName)
	tr.logInfof("- variables: %d", len(analysis.Variables))
	tr.logInfof("- lists: %d", len(analysis.Lists))
	tr.logInfof("- conditions: %d", len(analysis.Conditions))
	tr.logInfof("- tables: %d", len(analysis.Tables))

	return analysis, nil
}

// analyzeDocument analyzes the document structure.
func (tr *TemplateRenderer) analyzeDocument(doc *Document, analysis *TemplateAnalysis) {
	for i, element := range doc.Body.Elements {
		switch elem := element.(type) {
		case *Paragraph:
			tr.analyzeParagraph(elem, analysis)
		case *Table:
			tableAnalysis := tr.analyzeTable(elem, i)
			analysis.Tables = append(analysis.Tables, tableAnalysis)
		}
	}

	// Analyze template variables in headers and footers
	tr.analyzeHeadersFooters(doc, analysis)
}

// analyzeHeadersFooters analyzes template variables in headers and footers.
func (tr *TemplateRenderer) analyzeHeadersFooters(doc *Document, analysis *TemplateAnalysis) {
	if doc == nil || doc.parts == nil {
		return
	}

	// Iterate over all parts to find header and footer files
	for partName, partData := range doc.parts {
		if strings.HasPrefix(partName, "word/header") || strings.HasPrefix(partName, "word/footer") {
			// Parse header/footer XML and extract template variables
			tr.analyzeHeaderFooterXML(partData, analysis)
		}
	}
}

// analyzeHeaderFooterXML analyzes template variables in header/footer XML.
func (tr *TemplateRenderer) analyzeHeaderFooterXML(xmlData []byte, analysis *TemplateAnalysis) {
	// Use regex to extract text content from <w:t> tags
	textPattern := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
	matches := textPattern.FindAllSubmatch(xmlData, -1)

	var fullText strings.Builder
	for _, match := range matches {
		if len(match) >= 2 {
			fullText.Write(match[1])
		}
	}

	// Extract template variables
	tr.extractTemplateVariables(fullText.String(), analysis)
}

// analyzeParagraph analyzes a paragraph.
func (tr *TemplateRenderer) analyzeParagraph(para *Paragraph, analysis *TemplateAnalysis) {
	fullText := ""
	for _, run := range para.Runs {
		fullText += run.Text.Content
	}

	tr.extractTemplateVariables(fullText, analysis)
}

// analyzeTable analyzes a table.
func (tr *TemplateRenderer) analyzeTable(table *Table, index int) *TableAnalysis {
	tableAnalysis := &TableAnalysis{
		Index:         index,
		RowCount:      len(table.Rows),
		HasTemplate:   false,
		TemplateVars:  make(map[string]bool),
		LoopVariables: make([]string, 0),
	}

	if len(table.Rows) > 0 {
		tableAnalysis.ColCount = len(table.Rows[0].Cells)
	}

	// Check if this is a template table
	for rowIndex, row := range table.Rows {
		rowHasLoop := false
		for _, cell := range row.Cells {
			for _, para := range cell.Paragraphs {
				fullText := ""
				for _, run := range para.Runs {
					fullText += run.Text.Content
				}

				// Check loop syntax
				eachPattern := regexp.MustCompile(`\{\{#each\s+(\w+)\}\}`)
				if matches := eachPattern.FindStringSubmatch(fullText); len(matches) > 1 {
					tableAnalysis.HasTemplate = true
					tableAnalysis.TemplateRowIndex = rowIndex
					tableAnalysis.LoopVariables = append(tableAnalysis.LoopVariables, matches[1])
					rowHasLoop = true
				}

				// Extract variables
				varPattern := regexp.MustCompile(`\{\{(\w+)\}\}`)
				varMatches := varPattern.FindAllStringSubmatch(fullText, -1)
				for _, match := range varMatches {
					if len(match) >= 2 {
						tableAnalysis.TemplateVars[match[1]] = true
					}
				}
			}
		}

		if rowHasLoop {
			break
		}
	}

	return tableAnalysis
}

// extractTemplateVariables extracts template variables from text.
func (tr *TemplateRenderer) extractTemplateVariables(text string, analysis *TemplateAnalysis) {
	// Variables: {{variableName}}
	varPattern := regexp.MustCompile(`\{\{(\w+)\}\}`)
	varMatches := varPattern.FindAllStringSubmatch(text, -1)
	for _, match := range varMatches {
		if len(match) >= 2 {
			analysis.Variables[match[1]] = true
		}
	}

	// Conditions: {{#if condition}}
	ifPattern := regexp.MustCompile(`\{\{#if\s+(\w+)\}\}`)
	ifMatches := ifPattern.FindAllStringSubmatch(text, -1)
	for _, match := range ifMatches {
		if len(match) >= 2 {
			analysis.Conditions[match[1]] = true
		}
	}

	// Loops: {{#each list}}
	eachPattern := regexp.MustCompile(`\{\{#each\s+(\w+)\}\}`)
	eachMatches := eachPattern.FindAllStringSubmatch(text, -1)
	for _, match := range eachMatches {
		if len(match) >= 2 {
			analysis.Lists[match[1]] = true
		}
	}
}

// TemplateAnalysis holds the results of analyzing a template.
type TemplateAnalysis struct {
	TemplateName string           // template name
	Variables    map[string]bool  // variable list
	Lists        map[string]bool  // list variables
	Conditions   map[string]bool  // condition variables
	Tables       []*TableAnalysis // table analysis results
}

// TableAnalysis holds the results of analyzing a table.
type TableAnalysis struct {
	Index            int             // table index
	RowCount         int             // row count
	ColCount         int             // column count
	HasTemplate      bool            // whether it contains template syntax
	TemplateRowIndex int             // template row index
	TemplateVars     map[string]bool // template variables
	LoopVariables    []string        // loop variables
}

// GetRequiredData returns a TemplateData structure with sample values for all required fields.
func (analysis *TemplateAnalysis) GetRequiredData() *TemplateData {
	data := NewTemplateData()

	// Set sample variable values
	for varName := range analysis.Variables {
		data.SetVariable(varName, fmt.Sprintf("sample_%s", varName))
	}

	// Set sample condition values
	for condName := range analysis.Conditions {
		data.SetCondition(condName, true)
	}

	// Set sample list values
	for listName := range analysis.Lists {
		sampleList := []interface{}{
			map[string]interface{}{
				"sample_field1": "sample_value1",
				"sample_field2": "sample_value2",
			},
			map[string]interface{}{
				"sample_field1": "sample_value3",
				"sample_field2": "sample_value4",
			},
		}
		data.SetList(listName, sampleList)
	}

	return data
}
