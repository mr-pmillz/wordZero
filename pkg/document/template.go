// Package document implements template functionality.
package document

import (
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"strconv"
	"strings"
	"sync"
)

// Template-related errors
var (
	// ErrTemplateNotFound indicates the template was not found.
	ErrTemplateNotFound = NewDocumentError("template_not_found", fmt.Errorf("template not found"), "")

	// ErrTemplateSyntaxError indicates a template syntax error.
	ErrTemplateSyntaxError = NewDocumentError("template_syntax_error", fmt.Errorf("template syntax error"), "")

	// ErrTemplateRenderError indicates a template rendering error.
	ErrTemplateRenderError = NewDocumentError("template_render_error", fmt.Errorf("template render error"), "")

	// ErrInvalidTemplateData indicates invalid template data.
	ErrInvalidTemplateData = NewDocumentError("invalid_template_data", fmt.Errorf("invalid template data"), "")

	// ErrBlockNotFound indicates the block was not found.
	ErrBlockNotFound = NewDocumentError("block_not_found", fmt.Errorf("block not found"), "")

	// ErrInvalidBlockDefinition indicates an invalid block definition.
	ErrInvalidBlockDefinition = NewDocumentError("invalid_block_definition", fmt.Errorf("invalid block definition"), "")
)

// Pre-compiled regex for header/footer variable substitution
var headerFooterVarPattern = regexp.MustCompile(`\{\{(\w+)\}\}`)

// TemplateEngine manages template loading, caching, and rendering.
type TemplateEngine struct {
	cache    map[string]*Template // template cache
	mutex    sync.RWMutex         // read-write lock
	basePath string               // base path
}

// Template represents a parsed template.
type Template struct {
	Name          string                    // template name
	Content       string                    // template content
	BaseDoc       *Document                 // base document
	Variables     map[string]string         // template variables
	Blocks        []*TemplateBlock          // template block list
	Parent        *Template                 // parent template (for inheritance)
	DefinedBlocks map[string]*TemplateBlock // defined block map
}

// TemplateBlock represents a block within a template.
type TemplateBlock struct {
	Type           string                 // block type: variable, if, each, inherit, block, image
	Name           string                 // block name (used by block type)
	Content        string                 // block content
	Condition      string                 // condition (used by if block)
	Variable       string                 // variable name (used by each block)
	Children       []*TemplateBlock       // child blocks
	Data           map[string]interface{} // block data
	DefaultContent string                 // default content (for optional override)
	IsOverridden   bool                   // whether it has been overridden
}

// TemplateData holds the data used to render a template.
type TemplateData struct {
	Variables  map[string]interface{}        // variable data
	Lists      map[string][]interface{}      // list data
	Conditions map[string]bool               // condition data
	Images     map[string]*TemplateImageData // image data
}

// TemplateImageData holds image data for template rendering.
type TemplateImageData struct {
	FilePath string       // image file path
	Data     []byte       // image binary data (takes precedence)
	Config   *ImageConfig // image config (size, position, style, etc.)
	AltText  string       // image alt text
	Title    string       // image title
}

// NewTemplateEngine creates a new template engine.
func NewTemplateEngine() *TemplateEngine {
	return &TemplateEngine{
		cache: make(map[string]*Template),
		mutex: sync.RWMutex{},
	}
}

// SetBasePath sets the base path for templates.
func (te *TemplateEngine) SetBasePath(path string) {
	te.mutex.Lock()
	defer te.mutex.Unlock()
	te.basePath = path
}

// LoadTemplate loads a template from a string.
func (te *TemplateEngine) LoadTemplate(name, content string) (*Template, error) {
	te.mutex.Lock()
	defer te.mutex.Unlock()

	template := &Template{
		Name:          name,
		Content:       content,
		Variables:     make(map[string]string),
		Blocks:        make([]*TemplateBlock, 0),
		DefinedBlocks: make(map[string]*TemplateBlock),
	}

	// Parse template content
	if err := te.parseTemplate(template); err != nil {
		return nil, WrapErrorWithContext("load_template", err, name)
	}

	// Cache template
	te.cache[name] = template

	return template, nil
}

// LoadTemplateFromDocument creates a template from an existing document.
func (te *TemplateEngine) LoadTemplateFromDocument(name string, doc *Document) (*Template, error) {
	te.mutex.Lock()
	defer te.mutex.Unlock()

	// Extract template content from document
	content, err := te.extractTemplateContentFromDocument(doc)
	if err != nil {
		return nil, WrapErrorWithContext("load_template_from_document", err, name)
	}

	template := &Template{
		Name:          name,
		Content:       content,
		BaseDoc:       doc,
		Variables:     make(map[string]string),
		Blocks:        make([]*TemplateBlock, 0),
		DefinedBlocks: make(map[string]*TemplateBlock),
	}

	// Parse template content
	if err := te.parseTemplate(template); err != nil {
		return nil, WrapErrorWithContext("load_template_from_document", err, name)
	}

	// Cache template
	te.cache[name] = template

	return template, nil
}

// GetTemplate retrieves a cached template by name.
func (te *TemplateEngine) GetTemplate(name string) (*Template, error) {
	te.mutex.RLock()
	defer te.mutex.RUnlock()

	if template, exists := te.cache[name]; exists {
		return template, nil
	}

	return nil, WrapErrorWithContext("get_template", ErrTemplateNotFound.Cause, name)
}

// getTemplateInternal retrieves a cached template (internal method, no locking).
func (te *TemplateEngine) getTemplateInternal(name string) (*Template, error) {
	if template, exists := te.cache[name]; exists {
		return template, nil
	}

	return nil, WrapErrorWithContext("get_template", ErrTemplateNotFound.Cause, name)
}

// ClearCache clears the template cache.
func (te *TemplateEngine) ClearCache() {
	te.mutex.Lock()
	defer te.mutex.Unlock()
	te.cache = make(map[string]*Template)
}

// RemoveTemplate removes a template by name.
func (te *TemplateEngine) RemoveTemplate(name string) {
	te.mutex.Lock()
	defer te.mutex.Unlock()
	delete(te.cache, name)
}

// parseTemplate parses the template content.
func (te *TemplateEngine) parseTemplate(template *Template) error {
	content := template.Content

	// Parse variables: {{variableName}}
	varPattern := regexp.MustCompile(`\{\{(\w+)\}\}`)
	varMatches := varPattern.FindAllStringSubmatch(content, -1)
	for _, match := range varMatches {
		if len(match) >= 2 {
			varName := match[1]
			template.Variables[varName] = ""
		}
	}

	// Parse block definitions: {{#block "blockName"}}...{{/block}}
	blockPattern := regexp.MustCompile(`(?s)\{\{#block\s+"([^"]+)"\}\}(.*?)\{\{/block\}\}`)
	blockMatches := blockPattern.FindAllStringSubmatch(content, -1)
	for _, match := range blockMatches {
		if len(match) >= 3 {
			blockName := match[1]
			blockContent := match[2]

			block := &TemplateBlock{
				Type:           "block",
				Name:           blockName,
				Content:        blockContent,
				DefaultContent: blockContent,
				Children:       make([]*TemplateBlock, 0),
			}

			template.Blocks = append(template.Blocks, block)
			template.DefinedBlocks[blockName] = block
		}
	}

	// Parse conditional statements: {{#if condition}}...{{/if}} (fix: add (?s) flag to match newlines)
	ifPattern := regexp.MustCompile(`(?s)\{\{#if\s+(\w+)\}\}(.*?)\{\{/if\}\}`)
	ifMatches := ifPattern.FindAllStringSubmatch(content, -1)
	for _, match := range ifMatches {
		if len(match) >= 3 {
			condition := match[1]
			blockContent := match[2]

			block := &TemplateBlock{
				Type:      "if",
				Condition: condition,
				Content:   blockContent,
				Children:  make([]*TemplateBlock, 0),
			}

			template.Blocks = append(template.Blocks, block)
		}
	}

	// Parse loop statements: {{#each list}}...{{/each}} (fix: add (?s) flag to match newlines)
	eachPattern := regexp.MustCompile(`(?s)\{\{#each\s+(\w+)\}\}(.*?)\{\{/each\}\}`)
	eachMatches := eachPattern.FindAllStringSubmatch(content, -1)
	for _, match := range eachMatches {
		if len(match) >= 3 {
			listVar := match[1]
			blockContent := match[2]

			block := &TemplateBlock{
				Type:     "each",
				Variable: listVar,
				Content:  blockContent,
				Children: make([]*TemplateBlock, 0),
			}

			template.Blocks = append(template.Blocks, block)
		}
	}

	// Parse image placeholders: {{#image imageName}}
	imagePattern := regexp.MustCompile(`\{\{#image\s+(\w+)\}\}`)
	imageMatches := imagePattern.FindAllStringSubmatch(content, -1)
	for _, match := range imageMatches {
		if len(match) >= 2 {
			imageName := match[1]

			block := &TemplateBlock{
				Type:     "image",
				Name:     imageName,
				Content:  match[0], // save the full placeholder text
				Children: make([]*TemplateBlock, 0),
			}

			template.Blocks = append(template.Blocks, block)
		}
	}

	// Parse inheritance: {{extends "base_template"}}
	extendsPattern := regexp.MustCompile(`\{\{extends\s+"([^"]+)"\}\}`)
	extendsMatches := extendsPattern.FindStringSubmatch(content)
	if len(extendsMatches) >= 2 {
		baseName := extendsMatches[1]
		baseTemplate, err := te.getTemplateInternal(baseName)
		if err == nil {
			template.Parent = baseTemplate
			// Process block overrides
			te.processBlockOverrides(template, baseTemplate)
		}
	}

	return nil
}

// processBlockOverrides processes block overrides between child and parent templates.
func (te *TemplateEngine) processBlockOverrides(childTemplate, parentTemplate *Template) {
	// Iterate over child template block definitions and check for parent block overrides
	for blockName, childBlock := range childTemplate.DefinedBlocks {
		if parentBlock, exists := parentTemplate.DefinedBlocks[blockName]; exists {
			// Mark parent block as overridden
			parentBlock.IsOverridden = true
			parentBlock.Content = childBlock.Content
		}
	}

	// Recursively process grandparent templates
	if parentTemplate.Parent != nil {
		te.processBlockOverrides(childTemplate, parentTemplate.Parent)
	}
}

// RenderToDocument renders a template to a new document.
func (te *TemplateEngine) RenderToDocument(templateName string, data *TemplateData) (*Document, error) {
	template, err := te.GetTemplate(templateName)
	if err != nil {
		return nil, WrapErrorWithContext("render_to_document", err, templateName)
	}

	// Create new document
	var doc *Document
	if template.BaseDoc != nil {
		// Create based on base document
		doc = te.cloneDocument(template.BaseDoc)
	} else {
		// Create new document
		doc = New()
	}

	// Render template content
	renderedContent, err := te.renderTemplate(template, data)
	if err != nil {
		return nil, WrapErrorWithContext("render_to_document", err, templateName)
	}

	// Apply rendered content to document
	if err := te.applyRenderedContentToDocument(doc, renderedContent); err != nil {
		return nil, WrapErrorWithContext("render_to_document", err, templateName)
	}

	// Process image placeholders
	if err := te.processImagePlaceholders(doc, data); err != nil {
		return nil, WrapErrorWithContext("render_to_document", err, templateName)
	}

	return doc, nil
}

// renderTemplate renders a template with the given data.
func (te *TemplateEngine) renderTemplate(template *Template, data *TemplateData) (string, error) {
	var content string

	// Handle inheritance: if there is a parent template, use it as the base
	if template.Parent != nil {
		// Render parent template as the base content
		parentContent, err := te.renderTemplate(template.Parent, data)
		if err != nil {
			return "", err
		}
		content = parentContent

		// Apply child template block overrides to parent template content
		content = te.applyBlockOverrides(content, template)
	} else {
		// No parent template, use current template content directly
		content = template.Content
	}

	// Render block definitions
	content = te.renderBlocks(content, template, data)

	// Render variables
	content = te.renderVariables(content, data.Variables)

	// Render loop statements (process loops first; conditions inside loops are handled internally)
	content = te.renderLoops(content, data.Lists)

	// Render conditional statements (handle conditions outside loops)
	content = te.renderConditionals(content, data.Conditions)

	// Render image placeholders
	content = te.renderImages(content, data.Images)

	return content, nil
}

// applyBlockOverrides applies child template block overrides to parent template content.
func (te *TemplateEngine) applyBlockOverrides(content string, template *Template) string {
	// Replace parent template block placeholders with child template block content
	blockPattern := regexp.MustCompile(`(?s)\{\{#block\s+"([^"]+)"\}\}.*?\{\{/block\}\}`)

	return blockPattern.ReplaceAllStringFunc(content, func(match string) string {
		matches := blockPattern.FindStringSubmatch(match)
		if len(matches) >= 2 {
			blockName := matches[1]
			// If this block is defined in the child template, use the child's content
			if childBlock, exists := template.DefinedBlocks[blockName]; exists {
				return childBlock.Content
			}
		}
		return match // keep as is
	})
}

// renderBlocks renders block definitions.
func (te *TemplateEngine) renderBlocks(content string, template *Template, data *TemplateData) string {
	blockPattern := regexp.MustCompile(`(?s)\{\{#block\s+"([^"]+)"\}\}(.*?)\{\{/block\}\}`)

	return blockPattern.ReplaceAllStringFunc(content, func(match string) string {
		matches := blockPattern.FindStringSubmatch(match)
		if len(matches) >= 3 {
			blockName := matches[1]
			blockContent := matches[2]

			// Check if a block is defined
			if block, exists := template.DefinedBlocks[blockName]; exists {
				// If the block is overridden, use the overridden content; otherwise use default
				if block.IsOverridden {
					return block.Content
				}
				return block.DefaultContent
			}

			// If no block is defined, use the original content
			return blockContent
		}
		return match
	})
}

// renderVariables renders variable placeholders.
func (te *TemplateEngine) renderVariables(content string, variables map[string]interface{}) string {
	varPattern := regexp.MustCompile(`\{\{(\w+)\}\}`)

	return varPattern.ReplaceAllStringFunc(content, func(match string) string {
		varName := varPattern.FindStringSubmatch(match)[1]
		if value, exists := variables[varName]; exists {
			return te.interfaceToString(value)
		}
		return match // keep as is
	})
}

// renderConditionals renders conditional statements (supports if-else syntax).
func (te *TemplateEngine) renderConditionals(content string, conditions map[string]bool) string {
	ifElsePattern := regexp.MustCompile(`(?s)\{\{#if\s+(\w+)\}\}(.*?)\{\{/if\}\}`)

	return ifElsePattern.ReplaceAllStringFunc(content, func(match string) string {
		matches := ifElsePattern.FindStringSubmatch(match)
		if len(matches) >= 3 {
			condition := matches[1]
			blockContent := matches[2]

			// Check if there is an else clause
			elsePattern := regexp.MustCompile(`(?s)(.*?)\{\{else\}\}(.*?)`)
			elseMatches := elsePattern.FindStringSubmatch(blockContent)

			if len(elseMatches) >= 3 {
				// Has else clause
				ifContent := elseMatches[1]
				elseContent := elseMatches[2]

				if condValue, exists := conditions[condition]; exists && condValue {
					return ifContent
				} else {
					return elseContent
				}
			} else {
				// No else clause, handle with original logic
				if condValue, exists := conditions[condition]; exists && condValue {
					return blockContent
				}
			}
		}
		return "" // condition not met, return empty string
	})
}

// renderLoops renders loop statements.
func (te *TemplateEngine) renderLoops(content string, lists map[string][]interface{}) string {
	// Use a stack-based approach to correctly handle nested loops
	return te.renderLoopsNested(content, lists, 0)
}

// renderLoopsNested handles nested loops using recursion.
func (te *TemplateEngine) renderLoopsNested(content string, lists map[string][]interface{}, depth int) string {
	// Find the first {{#each}} tag
	eachStartPattern := regexp.MustCompile(`\{\{#each\s+(\w+)\}\}`)
	startMatch := eachStartPattern.FindStringIndex(content)

	if startMatch == nil {
		// No loop found, return as is
		return content
	}

	// Found loop start tag, now find the matching end tag
	startPos := startMatch[0]
	listVarMatch := eachStartPattern.FindStringSubmatch(content[startPos:])
	if len(listVarMatch) < 2 {
		return content
	}

	listVar := listVarMatch[1]
	blockStart := startMatch[1] // position after {{#each xxx}}

	// Use a stack to find the matching {{/each}}
	depth_counter := 1
	pos := blockStart
	blockEnd := -1

	for pos < len(content) {
		// Find the next {{#each}} or {{/each}}
		nextEach := eachStartPattern.FindStringIndex(content[pos:])
		endPattern := regexp.MustCompile(`\{\{/each\}\}`)
		nextEnd := endPattern.FindStringIndex(content[pos:])

		if nextEnd == nil {
			// No end tag found, syntax error
			break
		}

		// Determine whether the next match is a start or end tag
		if nextEach != nil && nextEach[0] < nextEnd[0] {
			// Next is a nested start tag
			depth_counter++
			pos = pos + nextEach[1]
		} else {
			// Next is an end tag
			depth_counter--
			if depth_counter == 0 {
				// Found the matching end tag
				blockEnd = pos + nextEnd[0]
				break
			}
			pos = pos + nextEnd[1]
		}
	}

	if blockEnd == -1 {
		// No matching end tag found
		return content
	}

	// Extract loop block content
	blockContent := content[blockStart:blockEnd]

	// Process loop
	var result strings.Builder

	// Add content before the loop
	result.WriteString(content[:startPos])

	// Render loop
	if listData, exists := lists[listVar]; exists {
		for i, item := range listData {
			// Create loop context variables
			loopContent := strings.ReplaceAll(blockContent, "{{this}}", te.interfaceToString(item))
			loopContent = strings.ReplaceAll(loopContent, "{{@index}}", strconv.Itoa(i))
			loopContent = strings.ReplaceAll(loopContent, "{{@first}}", strconv.FormatBool(i == 0))
			loopContent = strings.ReplaceAll(loopContent, "{{@last}}", strconv.FormatBool(i == len(listData)-1))

			// If item is a map, handle property access
			if itemMap, ok := item.(map[string]interface{}); ok {
				// Process nested loops first (before replacing variables)
				// Create a new lists map for nested loops containing the current item's list data
				nestedLists := make(map[string][]interface{})
				for key, value := range itemMap {
					// Check if the value is a list type
					if listValue, ok := value.([]interface{}); ok {
						nestedLists[key] = listValue
					}
				}

				// If there are nested lists, recursively process nested loops
				if len(nestedLists) > 0 {
					loopContent = te.renderLoopsNested(loopContent, nestedLists, depth+1)
				}

				// Then replace regular variables
				for key, value := range itemMap {
					placeholder := fmt.Sprintf("{{%s}}", key)
					// Only replace non-list type values
					if _, isList := value.([]interface{}); !isList {
						loopContent = strings.ReplaceAll(loopContent, placeholder, te.interfaceToString(value))
					}
				}

				// Process conditional statements inside the loop
				loopContent = te.renderLoopConditionals(loopContent, itemMap)
			}

			result.WriteString(loopContent)
		}
	}

	// Add content after the loop, and recursively process remaining loops
	remainingContent := content[blockEnd+len("{{/each}}"):]
	remainingContent = te.renderLoopsNested(remainingContent, lists, depth)
	result.WriteString(remainingContent)

	return result.String()
}

// renderLoopConditionals renders conditional statements inside loops (supports if-else syntax).
func (te *TemplateEngine) renderLoopConditionals(content string, itemData map[string]interface{}) string {
	ifElsePattern := regexp.MustCompile(`(?s)\{\{#if\s+(\w+)\}\}(.*?)\{\{/if\}\}`)

	return ifElsePattern.ReplaceAllStringFunc(content, func(match string) string {
		matches := ifElsePattern.FindStringSubmatch(match)
		if len(matches) >= 3 {
			condition := matches[1]
			blockContent := matches[2]

			// Check if there is an else clause
			elsePattern := regexp.MustCompile(`(?s)(.*?)\{\{else\}\}(.*?)`)
			elseMatches := elsePattern.FindStringSubmatch(blockContent)

			var ifContent, elseContent string
			hasElse := false

			if len(elseMatches) >= 3 {
				// Has else clause
				ifContent = elseMatches[1]
				elseContent = elseMatches[2]
				hasElse = true
			} else {
				// No else clause
				ifContent = blockContent
			}

			// Check if the condition exists in the current loop item's data
			if condValue, exists := itemData[condition]; exists {
				// Convert to boolean
				conditionMet := false
				switch v := condValue.(type) {
				case bool:
					conditionMet = v
				case string:
					conditionMet = v == "true" || v == "1" || v == "yes" || v != ""
				case int:
					conditionMet = v != 0
				case int64:
					conditionMet = v != 0
				case float64:
					conditionMet = v != 0.0
				default:
					// For other types, consider non-nil as true
					conditionMet = v != nil
				}

				if conditionMet {
					return ifContent
				} else if hasElse {
					return elseContent
				}
			} else if hasElse {
				// Condition not found, return else content
				return elseContent
			}
		}
		return "" // condition not met and no else clause, return empty string
	})
}

// interfaceToString converts an interface{} value to a string.
func (te *TemplateEngine) interfaceToString(value interface{}) string {
	if value == nil {
		return ""
	}

	switch v := value.(type) {
	case string:
		return v
	case int:
		return strconv.Itoa(v)
	case int64:
		return strconv.FormatInt(v, 10)
	case float64:
		return strconv.FormatFloat(v, 'f', -1, 64)
	case bool:
		return strconv.FormatBool(v)
	default:
		return fmt.Sprintf("%v", v)
	}
}

// ValidateTemplate validates template syntax.
func (te *TemplateEngine) ValidateTemplate(template *Template) error {
	content := template.Content

	// Check bracket pairing
	if err := te.validateBrackets(content); err != nil {
		return WrapErrorWithContext("validate_template", err, template.Name)
	}

	// Check block statement pairing
	if err := te.validateBlockStatements(content); err != nil {
		return WrapErrorWithContext("validate_template", err, template.Name)
	}

	// Check if statement pairing
	if err := te.validateIfStatements(content); err != nil {
		return WrapErrorWithContext("validate_template", err, template.Name)
	}

	// Check each statement pairing
	if err := te.validateEachStatements(content); err != nil {
		return WrapErrorWithContext("validate_template", err, template.Name)
	}

	return nil
}

// validateBrackets validates bracket pairing.
func (te *TemplateEngine) validateBrackets(content string) error {
	openCount := strings.Count(content, "{{")
	closeCount := strings.Count(content, "}}")

	if openCount != closeCount {
		return NewValidationError("brackets", content, "mismatched template brackets")
	}

	return nil
}

// validateBlockStatements validates block statement pairing.
func (te *TemplateEngine) validateBlockStatements(content string) error {
	blockCount := len(regexp.MustCompile(`\{\{#block\s+"[^"]+"\}\}`).FindAllString(content, -1))
	endblockCount := len(regexp.MustCompile(`\{\{/block\}\}`).FindAllString(content, -1))

	if blockCount != endblockCount {
		return NewValidationError("block_statements", content, "mismatched block/endblock statements")
	}

	return nil
}

// validateIfStatements validates if statement pairing.
func (te *TemplateEngine) validateIfStatements(content string) error {
	ifCount := len(regexp.MustCompile(`\{\{#if\s+\w+\}\}`).FindAllString(content, -1))
	endifCount := len(regexp.MustCompile(`\{\{/if\}\}`).FindAllString(content, -1))

	if ifCount != endifCount {
		return NewValidationError("if_statements", content, "mismatched if/endif statements")
	}

	return nil
}

// validateEachStatements validates each statement pairing.
func (te *TemplateEngine) validateEachStatements(content string) error {
	eachCount := len(regexp.MustCompile(`\{\{#each\s+\w+\}\}`).FindAllString(content, -1))
	endeachCount := len(regexp.MustCompile(`\{\{/each\}\}`).FindAllString(content, -1))

	if eachCount != endeachCount {
		return NewValidationError("each_statements", content, "mismatched each/endeach statements")
	}

	return nil
}

// documentToTemplateString converts a document to a template string.
func (te *TemplateEngine) documentToTemplateString(doc *Document) (string, error) {
	// No longer convert to a plain string; instead preserve the original document structure.
	// Variable substitution should be performed directly on the original document.
	return "", nil // handled in a separate method
}

// extractTemplateContentFromDocument extracts template content from a document.
func (te *TemplateEngine) extractTemplateContentFromDocument(doc *Document) (string, error) {
	var contentBuilder strings.Builder

	// Iterate over document elements and extract text content
	for _, element := range doc.Body.Elements {
		switch elem := element.(type) {
		case *Paragraph:
			// Extract text from paragraph
			for _, run := range elem.Runs {
				contentBuilder.WriteString(run.Text.Content)
			}
			contentBuilder.WriteString("\n")

		case *Table:
			// Skip tables for now, focus on template syntax in paragraphs.
			// Template syntax in tables is handled in RenderTemplateToDocument.
			continue
		}
	}

	// Extract template content from headers and footers
	te.extractHeaderFooterContent(doc, &contentBuilder)

	return contentBuilder.String(), nil
}

// extractHeaderFooterContent extracts template content from headers and footers.
func (te *TemplateEngine) extractHeaderFooterContent(doc *Document, contentBuilder *strings.Builder) {
	if doc.parts == nil {
		return
	}

	// Iterate over all parts to find header and footer files
	for partName, partData := range doc.parts {
		if strings.HasPrefix(partName, "word/header") || strings.HasPrefix(partName, "word/footer") {
			// Parse header/footer XML and extract text
			text := te.extractTextFromHeaderFooterXML(partData)
			if text != "" {
				contentBuilder.WriteString(text)
				contentBuilder.WriteString("\n")
			}
		}
	}
}

// extractTextFromHeaderFooterXML extracts text content from header/footer XML.
func (te *TemplateEngine) extractTextFromHeaderFooterXML(xmlData []byte) string {
	var contentBuilder strings.Builder

	// Use regex to extract text content from <w:t> tags.
	// This is a simplified parsing approach suitable for extracting template variables.
	textPattern := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
	matches := textPattern.FindAllSubmatch(xmlData, -1)

	for _, match := range matches {
		if len(match) >= 2 {
			contentBuilder.Write(match[1])
		}
	}

	return contentBuilder.String()
}

// cloneDocument deep copies all document elements and properties.
func (te *TemplateEngine) cloneDocument(source *Document) *Document {
	// Create new document
	doc := New()

	// Deep copy document elements
	for _, element := range source.Body.Elements {
		switch elem := element.(type) {
		case *Paragraph:
			clonedPara := te.cloneParagraph(elem)
			doc.Body.Elements = append(doc.Body.Elements, clonedPara)

		case *Table:
			clonedTable := te.cloneTable(elem)
			doc.Body.Elements = append(doc.Body.Elements, clonedTable)

		case *SectionProperties:
			clonedSectPr := te.cloneSectionProperties(elem)
			doc.Body.Elements = append(doc.Body.Elements, clonedSectPr)

		default:
			// Other types: copy reference for now
			doc.Body.Elements = append(doc.Body.Elements, element)
		}
	}

	// Deep copy style manager to ensure rendered styles match the original template
	if source.styleManager != nil {
		doc.styleManager = source.styleManager.Clone()
		// No longer force-modify the Normal style paragraph line spacing to avoid overriding
		// the template's own default line spacing/after-paragraph settings.
		// To unify line spacing, set it explicitly in the template rather than hardcoding in code.
	}

	// Copy all document parts to fully preserve the original document structure
	if doc.parts == nil {
		doc.parts = make(map[string][]byte)
	}
	te.cloneAllDocumentParts(source, doc)

	// Copy document relationships (including header/footer references)
	if source.documentRelationships != nil {
		doc.documentRelationships = &Relationships{
			Xmlns:         source.documentRelationships.Xmlns,
			Relationships: make([]Relationship, len(source.documentRelationships.Relationships)),
		}
		copy(doc.documentRelationships.Relationships, source.documentRelationships.Relationships)
	}

	// Copy content types
	if source.contentTypes != nil {
		doc.contentTypes = &ContentTypes{
			Xmlns:     source.contentTypes.Xmlns,
			Defaults:  make([]Default, len(source.contentTypes.Defaults)),
			Overrides: make([]Override, len(source.contentTypes.Overrides)),
		}
		copy(doc.contentTypes.Defaults, source.contentTypes.Defaults)
		copy(doc.contentTypes.Overrides, source.contentTypes.Overrides)
	}

	// Copy image ID counter
	doc.nextImageID = source.nextImageID

	return doc
}

// cloneAllDocumentParts copies all document parts to fully preserve the original document structure.
func (te *TemplateEngine) cloneAllDocumentParts(source, dest *Document) {
	if source.parts == nil {
		return
	}

	for partName, partData := range source.parts {
		// Skip word/document.xml because it will be regenerated on save
		if partName == "word/document.xml" {
			continue
		}

		// Copy part data
		dest.parts[partName] = make([]byte, len(partData))
		copy(dest.parts[partName], partData)
	}
}

// cloneSectionProperties deep copies section properties.
func (te *TemplateEngine) cloneSectionProperties(source *SectionProperties) *SectionProperties {
	if source == nil {
		return nil
	}

	sectPr := &SectionProperties{
		XmlnsR: source.XmlnsR,
	}

	// Copy page size
	if source.PageSize != nil {
		sectPr.PageSize = &PageSizeXML{
			W:      source.PageSize.W,
			H:      source.PageSize.H,
			Orient: source.PageSize.Orient,
		}
	}

	// Copy page margins
	if source.PageMargins != nil {
		sectPr.PageMargins = &PageMargin{
			Top:    source.PageMargins.Top,
			Right:  source.PageMargins.Right,
			Bottom: source.PageMargins.Bottom,
			Left:   source.PageMargins.Left,
			Header: source.PageMargins.Header,
			Footer: source.PageMargins.Footer,
			Gutter: source.PageMargins.Gutter,
		}
	}

	// Copy column settings
	if source.Columns != nil {
		sectPr.Columns = &Columns{
			Space: source.Columns.Space,
			Num:   source.Columns.Num,
		}
	}

	// Copy header references
	if source.HeaderReferences != nil {
		sectPr.HeaderReferences = make([]*HeaderFooterReference, len(source.HeaderReferences))
		for i, ref := range source.HeaderReferences {
			sectPr.HeaderReferences[i] = &HeaderFooterReference{
				Type: ref.Type,
				ID:   ref.ID,
			}
		}
	}

	// Copy footer references
	if source.FooterReferences != nil {
		sectPr.FooterReferences = make([]*FooterReference, len(source.FooterReferences))
		for i, ref := range source.FooterReferences {
			sectPr.FooterReferences[i] = &FooterReference{
				Type: ref.Type,
				ID:   ref.ID,
			}
		}
	}

	// Copy title page (different first page) setting
	if source.TitlePage != nil {
		sectPr.TitlePage = &TitlePage{}
	}

	// Copy page number type
	if source.PageNumType != nil {
		sectPr.PageNumType = &PageNumType{
			Fmt: source.PageNumType.Fmt,
		}
	}

	// Copy document grid
	if source.DocGrid != nil {
		sectPr.DocGrid = &DocGrid{
			Type:      source.DocGrid.Type,
			LinePitch: source.DocGrid.LinePitch,
			CharSpace: source.DocGrid.CharSpace,
		}
	}

	return sectPr
}

// cloneHeaderFooterParts copies header/footer parts (kept for backward compatibility; now handled by cloneAllDocumentParts).
func (te *TemplateEngine) cloneHeaderFooterParts(source, dest *Document) {
	if source.parts == nil {
		return
	}

	for partName, partData := range source.parts {
		// Copy header files
		if strings.HasPrefix(partName, "word/header") {
			dest.parts[partName] = make([]byte, len(partData))
			copy(dest.parts[partName], partData)
		}
		// Copy footer files
		if strings.HasPrefix(partName, "word/footer") {
			dest.parts[partName] = make([]byte, len(partData))
			copy(dest.parts[partName], partData)
		}
	}
}

// cloneParagraph deep copies a paragraph.
func (te *TemplateEngine) cloneParagraph(source *Paragraph) *Paragraph {
	newPara := &Paragraph{
		Properties: te.cloneParagraphProperties(source.Properties),
		Runs:       make([]Run, len(source.Runs)),
	}

	for i, run := range source.Runs {
		newPara.Runs[i] = te.cloneRun(&run)
	}

	return newPara
}

// cloneParagraphProperties deep copies paragraph properties.
func (te *TemplateEngine) cloneParagraphProperties(source *ParagraphProperties) *ParagraphProperties {
	if source == nil {
		return nil
	}

	props := &ParagraphProperties{}

	// Copy paragraph style
	if source.ParagraphStyle != nil {
		props.ParagraphStyle = &ParagraphStyle{
			Val: source.ParagraphStyle.Val,
		}
	}

	// Copy numbering properties
	if source.NumberingProperties != nil {
		props.NumberingProperties = &NumberingProperties{}
		if source.NumberingProperties.ILevel != nil {
			props.NumberingProperties.ILevel = &ILevel{Val: source.NumberingProperties.ILevel.Val}
		}
		if source.NumberingProperties.NumID != nil {
			props.NumberingProperties.NumID = &NumID{Val: source.NumberingProperties.NumID.Val}
		}
	}

	// Copy spacing
	if source.Spacing != nil {
		props.Spacing = &Spacing{
			Before:   source.Spacing.Before,
			After:    source.Spacing.After,
			Line:     source.Spacing.Line,
			LineRule: source.Spacing.LineRule,
		}
	}

	// Copy justification
	if source.Justification != nil {
		props.Justification = &Justification{
			Val: source.Justification.Val,
		}
	}

	// Copy indentation
	if source.Indentation != nil {
		props.Indentation = &Indentation{
			FirstLine: source.Indentation.FirstLine,
			Left:      source.Indentation.Left,
			Right:     source.Indentation.Right,
		}
	}

	// Copy tabs
	if source.Tabs != nil {
		props.Tabs = &Tabs{
			Tabs: make([]TabDef, len(source.Tabs.Tabs)),
		}
		for i, tab := range source.Tabs.Tabs {
			props.Tabs.Tabs[i] = TabDef{
				Val:    tab.Val,
				Leader: tab.Leader,
				Pos:    tab.Pos,
			}
		}
	}

	return props
}

// cloneRun deep copies a text run.
func (te *TemplateEngine) cloneRun(source *Run) Run {
	newRun := Run{
		Properties: te.cloneRunProperties(source.Properties),
		Text:       Text{Content: source.Text.Content, Space: source.Text.Space},
	}

	// Copy drawing (if present)
	if source.Drawing != nil {
		// Keep a simple copy for now; deep copying drawings is complex
		newRun.Drawing = source.Drawing
	}

	// Copy field char (if present)
	if source.FieldChar != nil {
		newRun.FieldChar = source.FieldChar
	}

	// Copy instruction text (if present)
	if source.InstrText != nil {
		newRun.InstrText = source.InstrText
	}

	return newRun
}

// cloneRunProperties deep copies run properties.
func (te *TemplateEngine) cloneRunProperties(source *RunProperties) *RunProperties {
	if source == nil {
		return nil
	}

	props := &RunProperties{}

	// Copy bold
	if source.Bold != nil {
		props.Bold = &Bold{}
	}

	// Copy complex script bold
	if source.BoldCs != nil {
		props.BoldCs = &BoldCs{}
	}

	// Copy italic
	if source.Italic != nil {
		props.Italic = &Italic{}
	}

	// Copy complex script italic
	if source.ItalicCs != nil {
		props.ItalicCs = &ItalicCs{}
	}

	// Copy underline
	if source.Underline != nil {
		props.Underline = &Underline{
			Val: source.Underline.Val,
		}
	}

	// Copy strikethrough
	if source.Strike != nil {
		props.Strike = &Strike{}
	}

	// Copy font size
	if source.FontSize != nil {
		props.FontSize = &FontSize{
			Val: source.FontSize.Val,
		}
	}

	// Copy complex script font size
	if source.FontSizeCs != nil {
		props.FontSizeCs = &FontSizeCs{
			Val: source.FontSizeCs.Val,
		}
	}

	// Copy color
	if source.Color != nil {
		props.Color = &Color{
			Val: source.Color.Val,
		}
	}

	// Copy highlight
	if source.Highlight != nil {
		props.Highlight = &Highlight{
			Val: source.Highlight.Val,
		}
	}

	// Copy font family properties including all font settings
	if source.FontFamily != nil {
		props.FontFamily = &FontFamily{
			ASCII:    source.FontFamily.ASCII,
			HAnsi:    source.FontFamily.HAnsi,
			EastAsia: source.FontFamily.EastAsia,
			CS:       source.FontFamily.CS,
			Hint:     source.FontFamily.Hint,
		}
	}

	return props
}

// cloneTable deep copies a table.
func (te *TemplateEngine) cloneTable(source *Table) *Table {
	newTable := &Table{
		Properties: te.cloneTableProperties(source.Properties),
		Grid:       te.cloneTableGrid(source.Grid),
		Rows:       make([]TableRow, len(source.Rows)),
	}

	for i, row := range source.Rows {
		newTable.Rows[i] = *te.cloneTableRow(&row)
	}

	return newTable
}

// cloneTableProperties deep copies table properties.
func (te *TemplateEngine) cloneTableProperties(source *TableProperties) *TableProperties {
	if source == nil {
		Debug("cloneTableProperties: source properties are nil")
		return nil
	}

	props := &TableProperties{}

	// Copy table width
	if source.TableW != nil {
		props.TableW = &TableWidth{
			W:    source.TableW.W,
			Type: source.TableW.Type,
		}
	}

	// Copy table alignment
	if source.TableJc != nil {
		props.TableJc = &TableJc{
			Val: source.TableJc.Val,
		}
	}

	// Copy table look
	if source.TableLook != nil {
		props.TableLook = &TableLook{
			Val:      source.TableLook.Val,
			FirstRow: source.TableLook.FirstRow,
			LastRow:  source.TableLook.LastRow,
			FirstCol: source.TableLook.FirstCol,
			LastCol:  source.TableLook.LastCol,
			NoHBand:  source.TableLook.NoHBand,
			NoVBand:  source.TableLook.NoVBand,
		}
	}

	// Copy table style
	if source.TableStyle != nil {
		props.TableStyle = &TableStyle{
			Val: source.TableStyle.Val,
		}
	}

	// Copy table borders
	if source.TableBorders != nil {
		props.TableBorders = te.cloneTableBorders(source.TableBorders)
	}

	// Copy table shading
	if source.Shd != nil {
		props.Shd = &TableShading{
			Val:       source.Shd.Val,
			Color:     source.Shd.Color,
			Fill:      source.Shd.Fill,
			ThemeFill: source.Shd.ThemeFill,
		}
	}

	// Copy table cell margins
	if source.TableCellMar != nil {
		props.TableCellMar = te.cloneTableCellMargins(source.TableCellMar)
	}

	// Copy table layout
	if source.TableLayout != nil {
		props.TableLayout = &TableLayoutType{
			Type: source.TableLayout.Type,
		}
	}

	// Copy table indentation
	if source.TableInd != nil {
		props.TableInd = &TableIndentation{
			W:    source.TableInd.W,
			Type: source.TableInd.Type,
		}
	}

	return props
}

// cloneTableBorders deep copies table borders.
func (te *TemplateEngine) cloneTableBorders(source *TableBorders) *TableBorders {
	if source == nil {
		return nil
	}

	borders := &TableBorders{}

	if source.Top != nil {
		borders.Top = &TableBorder{
			Val:        source.Top.Val,
			Sz:         source.Top.Sz,
			Space:      source.Top.Space,
			Color:      source.Top.Color,
			ThemeColor: source.Top.ThemeColor,
		}
	}

	if source.Left != nil {
		borders.Left = &TableBorder{
			Val:        source.Left.Val,
			Sz:         source.Left.Sz,
			Space:      source.Left.Space,
			Color:      source.Left.Color,
			ThemeColor: source.Left.ThemeColor,
		}
	}

	if source.Bottom != nil {
		borders.Bottom = &TableBorder{
			Val:        source.Bottom.Val,
			Sz:         source.Bottom.Sz,
			Space:      source.Bottom.Space,
			Color:      source.Bottom.Color,
			ThemeColor: source.Bottom.ThemeColor,
		}
	}

	if source.Right != nil {
		borders.Right = &TableBorder{
			Val:        source.Right.Val,
			Sz:         source.Right.Sz,
			Space:      source.Right.Space,
			Color:      source.Right.Color,
			ThemeColor: source.Right.ThemeColor,
		}
	}

	if source.InsideH != nil {
		borders.InsideH = &TableBorder{
			Val:        source.InsideH.Val,
			Sz:         source.InsideH.Sz,
			Space:      source.InsideH.Space,
			Color:      source.InsideH.Color,
			ThemeColor: source.InsideH.ThemeColor,
		}
	}

	if source.InsideV != nil {
		borders.InsideV = &TableBorder{
			Val:        source.InsideV.Val,
			Sz:         source.InsideV.Sz,
			Space:      source.InsideV.Space,
			Color:      source.InsideV.Color,
			ThemeColor: source.InsideV.ThemeColor,
		}
	}

	return borders
}

// cloneTableCellMargins deep copies table cell margins.
func (te *TemplateEngine) cloneTableCellMargins(source *TableCellMargins) *TableCellMargins {
	if source == nil {
		return nil
	}

	margins := &TableCellMargins{}

	if source.Top != nil {
		margins.Top = &TableCellSpace{
			W:    source.Top.W,
			Type: source.Top.Type,
		}
	}

	if source.Left != nil {
		margins.Left = &TableCellSpace{
			W:    source.Left.W,
			Type: source.Left.Type,
		}
	}

	if source.Bottom != nil {
		margins.Bottom = &TableCellSpace{
			W:    source.Bottom.W,
			Type: source.Bottom.Type,
		}
	}

	if source.Right != nil {
		margins.Right = &TableCellSpace{
			W:    source.Right.W,
			Type: source.Right.Type,
		}
	}

	return margins
}

// cloneTableGrid deep copies a table grid.
func (te *TemplateEngine) cloneTableGrid(source *TableGrid) *TableGrid {
	if source == nil {
		return nil
	}

	grid := &TableGrid{
		Cols: make([]TableGridCol, len(source.Cols)),
	}

	for i, col := range source.Cols {
		grid.Cols[i] = TableGridCol{
			W: col.W,
		}
	}

	return grid
}

// cloneTableCellMarginsCell deep copies table cell margins (cell level).
func (te *TemplateEngine) cloneTableCellMarginsCell(source *TableCellMarginsCell) *TableCellMarginsCell {
	if source == nil {
		return nil
	}

	margins := &TableCellMarginsCell{}

	if source.Top != nil {
		margins.Top = &TableCellSpaceCell{
			W:    source.Top.W,
			Type: source.Top.Type,
		}
	}

	if source.Left != nil {
		margins.Left = &TableCellSpaceCell{
			W:    source.Left.W,
			Type: source.Left.Type,
		}
	}

	if source.Bottom != nil {
		margins.Bottom = &TableCellSpaceCell{
			W:    source.Bottom.W,
			Type: source.Bottom.Type,
		}
	}

	if source.Right != nil {
		margins.Right = &TableCellSpaceCell{
			W:    source.Right.W,
			Type: source.Right.Type,
		}
	}

	return margins
}

// cloneTableCellBorders deep copies table cell borders.
func (te *TemplateEngine) cloneTableCellBorders(source *TableCellBorders) *TableCellBorders {
	if source == nil {
		return nil
	}

	borders := &TableCellBorders{}

	if source.Top != nil {
		borders.Top = &TableCellBorder{
			Val:        source.Top.Val,
			Sz:         source.Top.Sz,
			Space:      source.Top.Space,
			Color:      source.Top.Color,
			ThemeColor: source.Top.ThemeColor,
		}
	}

	if source.Left != nil {
		borders.Left = &TableCellBorder{
			Val:        source.Left.Val,
			Sz:         source.Left.Sz,
			Space:      source.Left.Space,
			Color:      source.Left.Color,
			ThemeColor: source.Left.ThemeColor,
		}
	}

	if source.Bottom != nil {
		borders.Bottom = &TableCellBorder{
			Val:        source.Bottom.Val,
			Sz:         source.Bottom.Sz,
			Space:      source.Bottom.Space,
			Color:      source.Bottom.Color,
			ThemeColor: source.Bottom.ThemeColor,
		}
	}

	if source.Right != nil {
		borders.Right = &TableCellBorder{
			Val:        source.Right.Val,
			Sz:         source.Right.Sz,
			Space:      source.Right.Space,
			Color:      source.Right.Color,
			ThemeColor: source.Right.ThemeColor,
		}
	}

	if source.InsideH != nil {
		borders.InsideH = &TableCellBorder{
			Val:        source.InsideH.Val,
			Sz:         source.InsideH.Sz,
			Space:      source.InsideH.Space,
			Color:      source.InsideH.Color,
			ThemeColor: source.InsideH.ThemeColor,
		}
	}

	if source.InsideV != nil {
		borders.InsideV = &TableCellBorder{
			Val:        source.InsideV.Val,
			Sz:         source.InsideV.Sz,
			Space:      source.InsideV.Space,
			Color:      source.InsideV.Color,
			ThemeColor: source.InsideV.ThemeColor,
		}
	}

	if source.TL2BR != nil {
		borders.TL2BR = &TableCellBorder{
			Val:        source.TL2BR.Val,
			Sz:         source.TL2BR.Sz,
			Space:      source.TL2BR.Space,
			Color:      source.TL2BR.Color,
			ThemeColor: source.TL2BR.ThemeColor,
		}
	}

	if source.TR2BL != nil {
		borders.TR2BL = &TableCellBorder{
			Val:        source.TR2BL.Val,
			Sz:         source.TR2BL.Sz,
			Space:      source.TR2BL.Space,
			Color:      source.TR2BL.Color,
			ThemeColor: source.TR2BL.ThemeColor,
		}
	}

	return borders
}

// cloneTableRow deep copies a table row.
func (te *TemplateEngine) cloneTableRow(source *TableRow) *TableRow {
	newRow := &TableRow{
		Properties: te.cloneTableRowProperties(source.Properties),
		Cells:      make([]TableCell, len(source.Cells)),
	}

	for i, cell := range source.Cells {
		newRow.Cells[i] = te.cloneTableCell(&cell)
	}

	return newRow
}

// cloneTableRowProperties deep copies table row properties.
func (te *TemplateEngine) cloneTableRowProperties(source *TableRowProperties) *TableRowProperties {
	if source == nil {
		return nil
	}

	props := &TableRowProperties{}

	// Copy row height
	if source.TableRowH != nil {
		props.TableRowH = &TableRowH{
			Val:   source.TableRowH.Val,
			HRule: source.TableRowH.HRule,
		}
	}

	// Copy can't split across pages
	if source.CantSplit != nil {
		props.CantSplit = &CantSplit{
			Val: source.CantSplit.Val,
		}
	}

	// Copy header row repeat
	if source.TblHeader != nil {
		props.TblHeader = &TblHeader{
			Val: source.TblHeader.Val,
		}
	}

	return props
}

// cloneTableCell deep copies a table cell.
func (te *TemplateEngine) cloneTableCell(source *TableCell) TableCell {
	newCell := TableCell{
		Properties: te.cloneTableCellProperties(source.Properties),
		Paragraphs: make([]Paragraph, len(source.Paragraphs)),
		Tables:     make([]Table, len(source.Tables)), // copy nested tables
	}

	for i, para := range source.Paragraphs {
		newCell.Paragraphs[i] = *te.cloneParagraph(&para)
	}

	// Deep copy nested tables
	for i, table := range source.Tables {
		newCell.Tables[i] = *te.cloneTable(&table)
	}

	return newCell
}

// cloneTableCellProperties deep copies table cell properties.
func (te *TemplateEngine) cloneTableCellProperties(source *TableCellProperties) *TableCellProperties {
	if source == nil {
		return nil
	}

	props := &TableCellProperties{}

	// Copy cell width
	if source.TableCellW != nil {
		props.TableCellW = &TableCellW{
			W:    source.TableCellW.W,
			Type: source.TableCellW.Type,
		}
	}

	// Copy cell margins
	if source.TcMar != nil {
		props.TcMar = te.cloneTableCellMarginsCell(source.TcMar)
	}

	// Copy cell borders
	if source.TcBorders != nil {
		props.TcBorders = te.cloneTableCellBorders(source.TcBorders)
	}

	// Copy cell shading
	if source.Shd != nil {
		props.Shd = &TableCellShading{
			Val:       source.Shd.Val,
			Color:     source.Shd.Color,
			Fill:      source.Shd.Fill,
			ThemeFill: source.Shd.ThemeFill,
		}
	}

	// Copy cell vertical alignment
	if source.VAlign != nil {
		props.VAlign = &VAlign{
			Val: source.VAlign.Val,
		}
	}

	// Copy grid span
	if source.GridSpan != nil {
		props.GridSpan = &GridSpan{
			Val: source.GridSpan.Val,
		}
	}

	// Copy vertical merge
	if source.VMerge != nil {
		props.VMerge = &VMerge{
			Val: source.VMerge.Val,
		}
	}

	// Copy text direction
	if source.TextDirection != nil {
		props.TextDirection = &TextDirection{
			Val: source.TextDirection.Val,
		}
	}

	// Copy no-wrap
	if source.NoWrap != nil {
		props.NoWrap = &NoWrap{
			Val: source.NoWrap.Val,
		}
	}

	// Copy hide mark
	if source.HideMark != nil {
		props.HideMark = &HideMark{
			Val: source.HideMark.Val,
		}
	}

	return props
}

// applyRenderedContentToDocument applies rendered content to a document.
func (te *TemplateEngine) applyRenderedContentToDocument(doc *Document, content string) error {
	// If content is empty, return immediately
	if strings.TrimSpace(content) == "" {
		return nil
	}

	// Split content by lines and create paragraphs
	lines := strings.Split(content, "\n")

	for _, line := range lines {
		// Create a new paragraph even for empty lines (to preserve formatting)
		para := &Paragraph{
			Properties: &ParagraphProperties{
				ParagraphStyle: &ParagraphStyle{Val: "Normal"},
			},
			Runs: []Run{},
		}

		// If the line is not empty, add text content
		if strings.TrimSpace(line) != "" {
			run := Run{
				Text: Text{Content: line},
				Properties: &RunProperties{
					FontFamily: &FontFamily{
						ASCII:    "仿宋",
						HAnsi:    "仿宋",
						EastAsia: "仿宋",
					},
					FontSize: &FontSize{Val: "24"}, // 12pt = 24 half-points
				},
			}
			para.Runs = append(para.Runs, run)
		} else {
			// Empty lines still need an empty Run to maintain paragraph structure
			run := Run{
				Text: Text{Content: ""},
				Properties: &RunProperties{
					FontFamily: &FontFamily{
						ASCII:    "仿宋",
						HAnsi:    "仿宋",
						EastAsia: "仿宋",
					},
					FontSize: &FontSize{Val: "24"},
				},
			}
			para.Runs = append(para.Runs, run)
		}

		// Add paragraph to document
		doc.Body.Elements = append(doc.Body.Elements, para)
	}

	return nil
}

// RenderTemplateToDocument renders a template to a new document (primary method).
func (te *TemplateEngine) RenderTemplateToDocument(templateName string, data *TemplateData) (*Document, error) {
	template, err := te.GetTemplate(templateName)
	if err != nil {
		return nil, WrapErrorWithContext("render_template_to_document", err, templateName)
	}

	// If there is a base document, clone it and perform variable substitution on it
	if template.BaseDoc != nil {
		doc := te.cloneDocument(template.BaseDoc)

		// Perform variable substitution directly in the document structure
		err := te.replaceVariablesInDocument(doc, data)
		if err != nil {
			return nil, WrapErrorWithContext("render_template_to_document", err, templateName)
		}

		return doc, nil
	}

	// If there is no base document, use the original approach
	return te.RenderToDocument(templateName, data)
}

// replaceVariablesInDocument replaces variables directly in the document structure.
func (te *TemplateEngine) replaceVariablesInDocument(doc *Document, data *TemplateData) error {
	// Process document-level loops (across paragraphs) first
	err := te.processDocumentLevelLoops(doc, data)
	if err != nil {
		return err
	}

	for _, element := range doc.Body.Elements {
		switch elem := element.(type) {
		case *Paragraph:
			// Replace variables in paragraph
			err := te.replaceVariablesInParagraph(elem, data)
			if err != nil {
				return err
			}

		case *Table:
			// Replace variables and process template syntax in table
			err := te.replaceVariablesInTable(elem, data)
			if err != nil {
				return err
			}
		}
	}

	// Replace variables in headers and footers
	err = te.replaceVariablesInHeadersFooters(doc, data)
	if err != nil {
		return err
	}

	// Process image placeholders
	err = te.processImagePlaceholders(doc, data)
	if err != nil {
		return err
	}

	return nil
}

// replaceVariablesInHeadersFooters replaces variables in headers and footers.
func (te *TemplateEngine) replaceVariablesInHeadersFooters(doc *Document, data *TemplateData) error {
	if doc.parts == nil {
		return nil
	}

	for partName, partData := range doc.parts {
		// Process header files
		if strings.HasPrefix(partName, "word/header") && strings.HasSuffix(partName, ".xml") {
			newData, err := te.replaceVariablesInXMLPart(partData, data)
			if err != nil {
				return fmt.Errorf("failed to replace variables in header %s: %v", partName, err)
			}
			doc.parts[partName] = newData
		}
		// Process footer files
		if strings.HasPrefix(partName, "word/footer") && strings.HasSuffix(partName, ".xml") {
			newData, err := te.replaceVariablesInXMLPart(partData, data)
			if err != nil {
				return fmt.Errorf("failed to replace variables in footer %s: %v", partName, err)
			}
			doc.parts[partName] = newData
		}
	}

	return nil
}

// replaceVariablesInXMLPart replaces variables in an XML part.
func (te *TemplateEngine) replaceVariablesInXMLPart(xmlData []byte, data *TemplateData) ([]byte, error) {
	content := string(xmlData)

	// Replace variables: {{variableName}} - using pre-compiled regex
	content = headerFooterVarPattern.ReplaceAllStringFunc(content, func(match string) string {
		matches := headerFooterVarPattern.FindStringSubmatch(match)
		if len(matches) < 2 {
			return match
		}
		varName := matches[1]
		if value, exists := data.Variables[varName]; exists {
			// Escape XML content
			return te.escapeXMLContent(te.interfaceToString(value))
		}
		return match // keep as is
	})

	// Replace conditional statements
	content = te.renderConditionals(content, data.Conditions)

	return []byte(content), nil
}

// escapeXMLContent escapes XML special characters.
func (te *TemplateEngine) escapeXMLContent(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	s = strings.ReplaceAll(s, "'", "&apos;")
	return s
}

// processDocumentLevelLoops processes document-level loops (across paragraphs).
func (te *TemplateEngine) processDocumentLevelLoops(doc *Document, data *TemplateData) error {
	elements := doc.Body.Elements
	newElements := make([]interface{}, 0)

	i := 0
	for i < len(elements) {
		element := elements[i]

		// Check if the current element contains a loop start tag
		if para, ok := element.(*Paragraph); ok {
			// Get the paragraph's full text
			fullText := ""
			for _, run := range para.Runs {
				fullText += run.Text.Content
			}

			// Check if it contains a loop start tag
			eachPattern := regexp.MustCompile(`\{\{#each\s+(\w+)\}\}`)
			matches := eachPattern.FindStringSubmatch(fullText)

			if len(matches) > 1 {
				listVarName := matches[1]

				// Find loop end position
				loopEndIndex := -1
				templateElements := make([]interface{}, 0)

				// Collect loop template elements (from current position to end tag)
				for j := i; j < len(elements); j++ {
					templateElements = append(templateElements, elements[j])

					if nextPara, ok := elements[j].(*Paragraph); ok {
						nextText := ""
						for _, run := range nextPara.Runs {
							nextText += run.Text.Content
						}

						if strings.Contains(nextText, "{{/each}}") {
							loopEndIndex = j
							break
						}
					}
				}

				if loopEndIndex >= 0 {
					// Process loop
					if listData, exists := data.Lists[listVarName]; exists {
						// Generate elements for each data item
						for _, item := range listData {
							if itemMap, ok := item.(map[string]interface{}); ok {
								// Clone template elements and replace variables
								for _, templateElement := range templateElements {
									if templatePara, ok := templateElement.(*Paragraph); ok {
										newPara := te.cloneParagraph(templatePara)

										// Process paragraph text
										fullText := ""
										for _, run := range newPara.Runs {
											fullText += run.Text.Content
										}

										// Remove loop tags
										content := fullText
										content = regexp.MustCompile(`\{\{#each\s+\w+\}\}`).ReplaceAllString(content, "")
										content = regexp.MustCompile(`\{\{/each\}\}`).ReplaceAllString(content, "")

										// Replace variables
										for key, value := range itemMap {
											placeholder := fmt.Sprintf("{{%s}}", key)
											content = strings.ReplaceAll(content, placeholder, te.interfaceToString(value))
										}

										// If content is not empty, create a new paragraph
										if strings.TrimSpace(content) != "" {
											// Preserve the original paragraph style, do not force bold (Fix for Issue #88)
											if len(newPara.Runs) > 0 {
												// Preserve the original Run's properties
												newPara.Runs[0].Text.Content = content
												newPara.Runs = newPara.Runs[:1]
											} else {
												// If there is no original Run, create a new unstyled Run
												newPara.Runs = []Run{{
													Text: Text{Content: content},
												}}
											}
											newElements = append(newElements, newPara)
										}
									}
								}
							}
						}
					}

					// Skip loop template elements
					i = loopEndIndex + 1
					continue
				}
			}
		}

		// Not a loop element, add directly
		newElements = append(newElements, element)
		i++
	}

	// Update document elements
	doc.Body.Elements = newElements
	return nil
}

// replaceVariablesInParagraph replaces variables in a paragraph (improved version with better style preservation).
func (te *TemplateEngine) replaceVariablesInParagraph(para *Paragraph, data *TemplateData) error {
	// First identify all variable placeholder positions
	fullText := ""
	runInfos := make([]struct {
		startIndex int
		endIndex   int
		run        *Run
	}, 0)

	currentIndex := 0
	for i := range para.Runs {
		runText := para.Runs[i].Text.Content
		if runText != "" {
			runInfos = append(runInfos, struct {
				startIndex int
				endIndex   int
				run        *Run
			}{
				startIndex: currentIndex,
				endIndex:   currentIndex + len(runText),
				run:        &para.Runs[i],
			})
			fullText += runText
			currentIndex += len(runText)
		}
	}

	// If there is no text content, return immediately
	if fullText == "" {
		return nil
	}

	// Process loop statements first (including non-table loops)
	processedText, hasLoopChanges := te.processNonTableLoops(fullText, data)
	if hasLoopChanges {
		// Rebuild paragraph
		para.Runs = []Run{{
			Text: Text{Content: processedText},
			Properties: &RunProperties{
				FontFamily: &FontFamily{
					ASCII:    "仿宋",
					HAnsi:    "仿宋",
					EastAsia: "仿宋",
				},
				Bold: &Bold{},
			},
		}}
		fullText = processedText
	}

	// Use the new sequential variable replacement method
	newRuns, hasVarChanges := te.replaceVariablesSequentially(runInfos, fullText, data)

	// If there are changes, update the paragraph's Runs
	if hasVarChanges || hasLoopChanges {
		para.Runs = newRuns
	}

	return nil
}

// processNonTableLoops processes non-table loops.
func (te *TemplateEngine) processNonTableLoops(content string, data *TemplateData) (string, bool) {
	eachPattern := regexp.MustCompile(`(?s)\{\{#each\s+(\w+)\}\}(.*?)\{\{/each\}\}`)
	matches := eachPattern.FindAllStringSubmatchIndex(content, -1)

	if len(matches) == 0 {
		return content, false
	}

	var result strings.Builder
	lastEnd := 0
	hasChanges := false

	for _, match := range matches {
		// Extract variable name and block content
		fullMatch := content[match[0]:match[1]]
		submatch := eachPattern.FindStringSubmatch(fullMatch)
		if len(submatch) >= 3 {
			listVar := submatch[1]
			blockContent := submatch[2]

			// Add content before the loop
			result.WriteString(content[lastEnd:match[0]])

			// Process loop
			if listData, exists := data.Lists[listVar]; exists {
				for _, item := range listData {
					if itemMap, ok := item.(map[string]interface{}); ok {
						loopContent := blockContent
						for key, value := range itemMap {
							placeholder := fmt.Sprintf("{{%s}}", key)
							loopContent = strings.ReplaceAll(loopContent, placeholder, te.interfaceToString(value))
						}
						result.WriteString(loopContent)
					}
				}
			}

			lastEnd = match[1]
			hasChanges = true
		}
	}

	// Add remaining content
	if lastEnd < len(content) {
		result.WriteString(content[lastEnd:])
	}

	return result.String(), hasChanges
}

// replaceVariablesSequentially replaces variables one by one while preserving styles.
func (te *TemplateEngine) replaceVariablesSequentially(originalRunInfos []struct {
	startIndex int
	endIndex   int
	run        *Run
}, originalText string, data *TemplateData) ([]Run, bool) {

	// Find all variable positions
	varPattern := regexp.MustCompile(`\{\{(\w+)\}\}`)
	varMatches := varPattern.FindAllStringSubmatchIndex(originalText, -1)

	if len(varMatches) == 0 {
		// No variables, check conditional statements
		return te.processConditionals(originalRunInfos, originalText, data)
	}

	newRuns := make([]Run, 0)
	currentPos := 0
	hasChanges := false

	for _, varMatch := range varMatches {
		varStart := varMatch[0]
		varEnd := varMatch[1]
		varNameStart := varMatch[2]
		varNameEnd := varMatch[3]

		// Add text before the variable (preserving original style)
		if varStart > currentPos {
			beforeText := originalText[currentPos:varStart]
			beforeRuns := te.extractRunsForSegment(originalRunInfos, currentPos, varStart, beforeText)
			newRuns = append(newRuns, beforeRuns...)
		}

		// Process variable replacement
		varName := originalText[varNameStart:varNameEnd]
		if value, exists := data.Variables[varName]; exists {
			replacementText := te.interfaceToString(value)

			// Select appropriate style for the variable (use the Run style covering the variable position)
			varRun := te.findRunForPosition(originalRunInfos, varStart)
			if varRun != nil {
				newRun := te.cloneRun(varRun)
				newRun.Text.Content = replacementText
				newRuns = append(newRuns, newRun)
				hasChanges = true
			}
		} else {
			// Variable not found, keep original placeholder
			varText := originalText[varStart:varEnd]
			varRun := te.findRunForPosition(originalRunInfos, varStart)
			if varRun != nil {
				newRun := te.cloneRun(varRun)
				newRun.Text.Content = varText
				newRuns = append(newRuns, newRun)
			}
		}

		currentPos = varEnd
	}

	// Add the remaining text at the end
	if currentPos < len(originalText) {
		afterText := originalText[currentPos:]
		afterRuns := te.extractRunsForSegment(originalRunInfos, currentPos, len(originalText), afterText)
		newRuns = append(newRuns, afterRuns...)
	}

	// If no variables were found but text changed, process conditional statements
	if !hasChanges {
		return te.processConditionals(originalRunInfos, originalText, data)
	}

	// Process conditional statements on the result (while keeping each Run independent)
	if hasChanges {
		finalRuns := te.processConditionalsPreservingRuns(newRuns, data)
		return finalRuns, true
	}

	return newRuns, hasChanges
}

// processConditionalsPreservingRuns processes conditional statements while preserving Run independence.
func (te *TemplateEngine) processConditionalsPreservingRuns(runs []Run, data *TemplateData) []Run {
	finalRuns := make([]Run, 0)

	for _, run := range runs {
		originalContent := run.Text.Content
		processedContent := te.renderConditionals(originalContent, data.Conditions)

		// If content changed, update this Run
		if processedContent != originalContent {
			newRun := run // copy Run struct
			newRun.Text.Content = processedContent
			finalRuns = append(finalRuns, newRun)
		} else {
			// Content unchanged, keep as is
			finalRuns = append(finalRuns, run)
		}
	}

	return finalRuns
}

// processConditionals processes conditional statements.
func (te *TemplateEngine) processConditionals(originalRunInfos []struct {
	startIndex int
	endIndex   int
	run        *Run
}, originalText string, data *TemplateData) ([]Run, bool) {

	processedText := te.renderConditionals(originalText, data.Conditions)

	if processedText == originalText {
		// No changes, return original Runs
		newRuns := make([]Run, len(originalRunInfos))
		for i, runInfo := range originalRunInfos {
			newRuns[i] = te.cloneRun(runInfo.run)
		}
		return newRuns, false
	}

	// Conditionals were processed, simplify handling
	if len(originalRunInfos) == 1 {
		newRun := te.cloneRun(originalRunInfos[0].run)
		newRun.Text.Content = processedText
		return []Run{newRun}, true
	}

	// Multiple Runs case: use the first Run's style
	newRun := te.cloneRun(originalRunInfos[0].run)
	newRun.Text.Content = processedText
	return []Run{newRun}, true
}

// extractRunsForSegment extracts the corresponding Runs for a text segment (improved version).
func (te *TemplateEngine) extractRunsForSegment(originalRunInfos []struct {
	startIndex int
	endIndex   int
	run        *Run
}, segmentStart, segmentEnd int, segmentText string) []Run {
	runs := make([]Run, 0)

	for _, runInfo := range originalRunInfos {
		// Check if the Run overlaps with the text segment
		if runInfo.endIndex > segmentStart && runInfo.startIndex < segmentEnd {
			overlapStart := max(runInfo.startIndex, segmentStart)
			overlapEnd := min(runInfo.endIndex, segmentEnd)

			if overlapEnd > overlapStart {
				newRun := te.cloneRun(runInfo.run)
				// Calculate relative position within the segment text
				relativeStart := overlapStart - segmentStart
				relativeEnd := overlapEnd - segmentStart

				// Ensure indices are within valid range
				if relativeStart >= 0 && relativeEnd <= len(segmentText) && relativeStart < relativeEnd {
					newRun.Text.Content = segmentText[relativeStart:relativeEnd]
					if newRun.Text.Content != "" {
						runs = append(runs, newRun)
					}
				}
			}
		}
	}

	return runs
}

// findRunForPosition finds the Run covering the specified position.
func (te *TemplateEngine) findRunForPosition(originalRunInfos []struct {
	startIndex int
	endIndex   int
	run        *Run
}, position int) *Run {
	for _, runInfo := range originalRunInfos {
		if position >= runInfo.startIndex && position < runInfo.endIndex {
			return runInfo.run
		}
	}
	// If not found, return the first Run
	if len(originalRunInfos) > 0 {
		return originalRunInfos[0].run
	}
	return nil
}

// max returns the larger of two integers.
func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}

// min returns the smaller of two integers.
func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}

// replaceVariablesInTable replaces variables and processes table templates.
func (te *TemplateEngine) replaceVariablesInTable(table *Table, data *TemplateData) error {
	// Check if there is a table loop template
	if len(table.Rows) > 0 && te.isTableTemplate(table) {
		return te.renderTableTemplate(table, data)
	}

	// Regular table variable replacement
	for i := range table.Rows {
		for j := range table.Rows[i].Cells {
			for k := range table.Rows[i].Cells[j].Paragraphs {
				err := te.replaceVariablesInParagraph(&table.Rows[i].Cells[j].Paragraphs[k], data)
				if err != nil {
					return err
				}
			}
			// Recursively process nested tables
			for k := range table.Rows[i].Cells[j].Tables {
				err := te.replaceVariablesInTable(&table.Rows[i].Cells[j].Tables[k], data)
				if err != nil {
					return err
				}
			}
		}
	}

	return nil
}

// isTableTemplate checks if a table contains template syntax (supports cross-Run detection).
func (te *TemplateEngine) isTableTemplate(table *Table) bool {
	if len(table.Rows) == 0 {
		return false
	}

	// Check all rows for loop syntax, with cross-Run detection support
	for _, row := range table.Rows {
		for _, cell := range row.Cells {
			for _, para := range cell.Paragraphs {
				// Use the new cross-Run detection method
				if te.containsTemplateLoopInRuns(para.Runs) {
					return true
				}
			}
		}
	}

	return false
}

// containsTemplateLoop checks if text contains loop template syntax.
func (te *TemplateEngine) containsTemplateLoop(text string) bool {
	eachPattern := regexp.MustCompile(`\{\{#each\s+\w+\}\}`)
	return eachPattern.MatchString(text)
}

// containsTemplateLoopInRuns checks if a list of Runs contains loop template syntax (cross-Run detection).
func (te *TemplateEngine) containsTemplateLoopInRuns(runs []Run) bool {
	// Merge text from all Runs
	fullText := ""
	for _, run := range runs {
		fullText += run.Text.Content
	}

	return te.containsTemplateLoop(fullText)
}

// renderTableTemplate renders a table template.
func (te *TemplateEngine) renderTableTemplate(table *Table, data *TemplateData) error {
	if len(table.Rows) == 0 {
		return nil
	}

	// Find the template row (row containing loop syntax)
	templateRowIndex := -1
	var listVarName string

	for i, row := range table.Rows {
		found := false
		// Check all cells in the row, merge text to resolve cross-Run variable issues
		for _, cell := range row.Cells {
			for _, para := range cell.Paragraphs {
				// Merge text from all Runs
				fullText := ""
				for _, run := range para.Runs {
					fullText += run.Text.Content
				}

				// Check if the merged text contains loop syntax
				eachPattern := regexp.MustCompile(`\{\{#each\s+(\w+)\}\}`)
				matches := eachPattern.FindStringSubmatch(fullText)
				if len(matches) > 1 {
					templateRowIndex = i
					listVarName = matches[1]
					found = true
					break
				}
			}
			if found {
				break
			}
		}
		if found {
			break
		}
	}

	if templateRowIndex < 0 || listVarName == "" {
		return nil
	}

	// Get list data
	listData, exists := data.Lists[listVarName]
	if !exists || len(listData) == 0 {
		// Remove template row
		table.Rows = append(table.Rows[:templateRowIndex], table.Rows[templateRowIndex+1:]...)
		return nil
	}

	// Save template row
	templateRow := table.Rows[templateRowIndex]
	newRows := make([]TableRow, 0)

	// Preserve rows before the template row (deep clone to maintain styles)
	for _, row := range table.Rows[:templateRowIndex] {
		clonedRow := te.cloneTableRow(&row)
		newRows = append(newRows, *clonedRow)
	}

	// Generate new rows for each data item
	for _, item := range listData {
		newRow := te.cloneTableRow(&templateRow)

		// Replace variables in the new row
		if itemMap, ok := item.(map[string]interface{}); ok {
			for i := range newRow.Cells {
				for j := range newRow.Cells[i].Paragraphs {
					// Merge text from all Runs
					fullText := ""
					originalRuns := newRow.Cells[i].Paragraphs[j].Runs
					for _, run := range originalRuns {
						fullText += run.Text.Content
					}

					// Remove template syntax tags
					content := fullText
					content = regexp.MustCompile(`\{\{#each\s+\w+\}\}`).ReplaceAllString(content, "")
					content = regexp.MustCompile(`\{\{/each\}\}`).ReplaceAllString(content, "")

					// Replace variables
					for key, value := range itemMap {
						placeholder := fmt.Sprintf("{{%s}}", key)
						content = strings.ReplaceAll(content, placeholder, te.interfaceToString(value))
					}

					// Process conditional statements
					content = te.renderLoopConditionals(content, itemMap)

					// Rebuild Run structure for better style inheritance
					if len(originalRuns) > 0 {
						// Find the first Run with actual content or style as a style template
						var templateRun *Run
						for k := range originalRuns {
							if originalRuns[k].Properties != nil || originalRuns[k].Text.Content != "" {
								templateRun = &originalRuns[k]
								break
							}
						}

						if templateRun != nil {
							newRun := te.cloneRun(templateRun)
							newRun.Text.Content = content
							newRow.Cells[i].Paragraphs[j].Runs = []Run{newRun}
						} else {
							// Use the first Run but ensure basic styling
							newRun := te.cloneRun(&originalRuns[0])
							newRun.Text.Content = content
							// Ensure basic font settings
							if newRun.Properties == nil {
								newRun.Properties = &RunProperties{}
							}
							if newRun.Properties.FontFamily == nil {
								newRun.Properties.FontFamily = &FontFamily{
									ASCII:    "仿宋",
									HAnsi:    "仿宋",
									EastAsia: "仿宋",
								}
							}
							newRow.Cells[i].Paragraphs[j].Runs = []Run{newRun}
						}
					} else {
						// If there is no original Run, create a new one but try to inherit paragraph style
						newRun := Run{
							Text: Text{Content: content},
							Properties: &RunProperties{
								FontFamily: &FontFamily{
									ASCII:    "仿宋",
									HAnsi:    "仿宋",
									EastAsia: "仿宋",
								},
								Bold: &Bold{},
							},
						}

						// If the paragraph has default Run properties, try to inherit them
						if len(templateRow.Cells) > i && len(templateRow.Cells[i].Paragraphs) > j {
							templatePara := &templateRow.Cells[i].Paragraphs[j]
							if len(templatePara.Runs) > 0 && templatePara.Runs[0].Properties != nil {
								newRun.Properties = te.cloneRunProperties(templatePara.Runs[0].Properties)
							}
						}

						newRow.Cells[i].Paragraphs[j].Runs = []Run{newRun}
					}
				}

				// Process variable replacement in nested tables
				for k := range newRow.Cells[i].Tables {
					// Create template data for nested table variable replacement
					nestedData := NewTemplateData()
					nestedData.Variables = make(map[string]interface{})
					for key, value := range itemMap {
						nestedData.Variables[key] = value
					}
					// Recursively process nested table
					err := te.replaceVariablesInTable(&newRow.Cells[i].Tables[k], nestedData)
					if err != nil {
						Debugf("error replacing variables in nested table: %v", err)
					}
				}
			}
		}

		newRows = append(newRows, *newRow)
	}

	// Preserve rows after the template row (deep clone to maintain styles)
	for _, row := range table.Rows[templateRowIndex+1:] {
		clonedRow := te.cloneTableRow(&row)
		newRows = append(newRows, *clonedRow)
	}

	// Update table rows
	table.Rows = newRows

	return nil
}

// NewTemplateData creates a new TemplateData instance.
func NewTemplateData() *TemplateData {
	return &TemplateData{
		Variables:  make(map[string]interface{}),
		Lists:      make(map[string][]interface{}),
		Conditions: make(map[string]bool),
		Images:     make(map[string]*TemplateImageData),
	}
}

// SetVariable sets a variable value.
func (td *TemplateData) SetVariable(name string, value interface{}) {
	td.Variables[name] = value
}

// SetList sets a list value.
func (td *TemplateData) SetList(name string, list []interface{}) {
	td.Lists[name] = list
}

// SetCondition sets a condition value.
func (td *TemplateData) SetCondition(name string, value bool) {
	td.Conditions[name] = value
}

// SetVariables sets multiple variables at once.
func (td *TemplateData) SetVariables(variables map[string]interface{}) {
	for name, value := range variables {
		td.Variables[name] = value
	}
}

// GetVariable retrieves a variable value.
func (td *TemplateData) GetVariable(name string) (interface{}, bool) {
	value, exists := td.Variables[name]
	return value, exists
}

// GetList retrieves a list value.
func (td *TemplateData) GetList(name string) ([]interface{}, bool) {
	list, exists := td.Lists[name]
	return list, exists
}

// GetCondition retrieves a condition value.
func (td *TemplateData) GetCondition(name string) (bool, bool) {
	value, exists := td.Conditions[name]
	return value, exists
}

// GetImage retrieves image data.
func (td *TemplateData) GetImage(name string) (*TemplateImageData, bool) {
	value, exists := td.Images[name]
	return value, exists
}

// Merge merges another TemplateData into this one.
func (td *TemplateData) Merge(other *TemplateData) {
	// Merge variables
	for key, value := range other.Variables {
		td.Variables[key] = value
	}

	// Merge lists
	for key, value := range other.Lists {
		td.Lists[key] = value
	}

	// Merge conditions
	for key, value := range other.Conditions {
		td.Conditions[key] = value
	}

	// Merge images
	for key, value := range other.Images {
		td.Images[key] = value
	}
}

// Clear clears all template data.
func (td *TemplateData) Clear() {
	td.Variables = make(map[string]interface{})
	td.Lists = make(map[string][]interface{})
	td.Conditions = make(map[string]bool)
	td.Images = make(map[string]*TemplateImageData)
}

// FromStruct populates template data from a struct.
func (td *TemplateData) FromStruct(data interface{}) error {
	value := reflect.ValueOf(data)
	if value.Kind() == reflect.Ptr {
		value = value.Elem()
	}

	if value.Kind() != reflect.Struct {
		return NewValidationError("data_type", "struct", "expected struct type")
	}

	typ := value.Type()
	for i := 0; i < value.NumField(); i++ {
		field := typ.Field(i)
		fieldValue := value.Field(i)

		// Skip unexported fields
		if !fieldValue.CanInterface() {
			continue
		}

		fieldName := strings.ToLower(field.Name)
		td.Variables[fieldName] = fieldValue.Interface()
	}

	return nil
}

// SetImage sets image data by file path.
func (td *TemplateData) SetImage(name, filePath string, config *ImageConfig) {
	imageData := &TemplateImageData{
		FilePath: filePath,
		Config:   config,
	}
	td.Images[name] = imageData
}

// SetImageFromData sets image data from binary data.
func (td *TemplateData) SetImageFromData(name string, data []byte, config *ImageConfig) {
	imageData := &TemplateImageData{
		Data:   data,
		Config: config,
	}
	td.Images[name] = imageData
}

// SetImageWithDetails sets image data with full configuration.
func (td *TemplateData) SetImageWithDetails(name, filePath string, data []byte, config *ImageConfig, altText, title string) {
	imageData := &TemplateImageData{
		FilePath: filePath,
		Data:     data,
		Config:   config,
		AltText:  altText,
		Title:    title,
	}
	td.Images[name] = imageData
}

// renderImages renders image placeholders.
func (te *TemplateEngine) renderImages(content string, images map[string]*TemplateImageData) string {
	// Image placeholder pattern: {{#image imageName}}
	imagePattern := regexp.MustCompile(`\{\{#image\s+(\w+)\}\}`)

	return imagePattern.ReplaceAllStringFunc(content, func(match string) string {
		matches := imagePattern.FindStringSubmatch(match)
		if len(matches) >= 2 {
			imageName := matches[1]

			// Look up image data
			if _, exists := images[imageName]; exists {
				// In traditional string templates, return an image placeholder marker.
				// Actual image processing happens in RenderTemplateToDocument.
				return fmt.Sprintf("[IMAGE:%s]", imageName)
			}
		}
		// If image data is not found, keep as is or return error message
		return fmt.Sprintf("[IMAGE_NOT_FOUND:%s]", matches[1])
	})
}

// processImagePlaceholders processes image placeholders in a document.
func (te *TemplateEngine) processImagePlaceholders(doc *Document, data *TemplateData) error {
	// Iterate over document elements to find and replace image placeholders
	for i, element := range doc.Body.Elements {
		switch elem := element.(type) {
		case *Paragraph:
			// Check if paragraph contains image placeholders
			newElements, err := te.processImagePlaceholdersInParagraph(elem, data, doc)
			if err != nil {
				return err
			}

			// If images were replaced, update document elements
			if len(newElements) > 1 || (len(newElements) == 1 && newElements[0] != elem) {
				// Remove original paragraph and insert new elements (may include image paragraphs)
				doc.Body.Elements = append(doc.Body.Elements[:i], append(newElements, doc.Body.Elements[i+1:]...)...)
			}
		case *Table:
			// Process image placeholders in tables (Fix for Issue #91)
			if err := te.processImagePlaceholdersInTable(elem, data, doc); err != nil {
				return err
			}
		}
	}
	return nil
}

// processImagePlaceholdersInTable processes image placeholders in a table (Fix for Issue #91).
func (te *TemplateEngine) processImagePlaceholdersInTable(table *Table, data *TemplateData, doc *Document) error {
	for rowIdx := range table.Rows {
		for cellIdx := range table.Rows[rowIdx].Cells {
			cell := &table.Rows[rowIdx].Cells[cellIdx]
			// Process each paragraph in the cell
			for paraIdx := range cell.Paragraphs {
				para := &cell.Paragraphs[paraIdx]
				newElements, err := te.processImagePlaceholdersInParagraph(para, data, doc)
				if err != nil {
					return err
				}

				// If images were replaced
				if len(newElements) > 0 {
					// Check if the returned elements differ from the original paragraph
					if len(newElements) == 1 {
						if newPara, ok := newElements[0].(*Paragraph); ok {
							cell.Paragraphs[paraIdx] = *newPara
						}
					} else {
						// Multiple elements: replace current paragraph with the first, append the rest
						newParagraphs := make([]Paragraph, 0, len(cell.Paragraphs)-1+len(newElements))
						newParagraphs = append(newParagraphs, cell.Paragraphs[:paraIdx]...)
						for _, elem := range newElements {
							if p, ok := elem.(*Paragraph); ok {
								newParagraphs = append(newParagraphs, *p)
							}
						}
						newParagraphs = append(newParagraphs, cell.Paragraphs[paraIdx+1:]...)
						cell.Paragraphs = newParagraphs
					}
				}
			}
		}
	}
	return nil
}

// processImagePlaceholdersInParagraph processes image placeholders in a paragraph.
func (te *TemplateEngine) processImagePlaceholdersInParagraph(para *Paragraph, data *TemplateData, doc *Document) ([]interface{}, error) {
	// Get the paragraph's full text
	fullText := ""
	for _, run := range para.Runs {
		fullText += run.Text.Content
	}

	// Check for image placeholders (supports two formats)
	// 1. Original template format: {{#image imageName}}
	// 2. Rendered format: [IMAGE:imageName]
	originalImagePattern := regexp.MustCompile(`\{\{#image\s+(\w+)\}\}`)
	renderedImagePattern := regexp.MustCompile(`\[IMAGE:(\w+)\]`)

	originalMatches := originalImagePattern.FindAllStringSubmatch(fullText, -1)
	renderedMatches := renderedImagePattern.FindAllStringSubmatch(fullText, -1)

	// Merge match results from both formats
	allMatches := make([][2]string, 0)
	for _, match := range originalMatches {
		allMatches = append(allMatches, [2]string{match[0], match[1]})
	}
	for _, match := range renderedMatches {
		allMatches = append(allMatches, [2]string{match[0], match[1]})
	}

	if len(allMatches) == 0 {
		// No image placeholders, return original paragraph
		return []interface{}{para}, nil
	}

	result := make([]interface{}, 0)
	lastEnd := 0

	// Process each image placeholder
	for _, match := range allMatches {
		imageName := match[1]
		matchStart := strings.Index(fullText[lastEnd:], match[0]) + lastEnd
		matchEnd := matchStart + len(match[0])

		// Add text before the image placeholder (if any)
		if matchStart > lastEnd {
			beforeText := fullText[lastEnd:matchStart]
			if strings.TrimSpace(beforeText) != "" {
				beforePara := te.createTextParagraph(beforeText, para)
				result = append(result, beforePara)
			}
		}

		// Create image paragraph
		if imageData, exists := data.Images[imageName]; exists {
			imagePara, err := te.createImageParagraph(imageData, doc)
			if err != nil {
				return nil, fmt.Errorf("failed to create image paragraph: %v", err)
			}
			result = append(result, imagePara)
		} else {
			// Image data not found, create error text paragraph
			errorPara := te.createTextParagraph(fmt.Sprintf("[IMAGE_NOT_FOUND: %s]", imageName), para)
			result = append(result, errorPara)
		}

		lastEnd = matchEnd
	}

	// Add the remaining text at the end (if any)
	if lastEnd < len(fullText) {
		afterText := fullText[lastEnd:]
		if strings.TrimSpace(afterText) != "" {
			afterPara := te.createTextParagraph(afterText, para)
			result = append(result, afterPara)
		}
	}

	// If there is no content, return an empty paragraph
	if len(result) == 0 {
		emptyPara := te.createTextParagraph("", para)
		result = append(result, emptyPara)
	}

	return result, nil
}

// createTextParagraph creates a text paragraph (preserving the original paragraph's style).
func (te *TemplateEngine) createTextParagraph(text string, originalPara *Paragraph) *Paragraph {
	newPara := te.cloneParagraph(originalPara)

	// Set text content while preserving original style
	if len(newPara.Runs) > 0 {
		newPara.Runs[0].Text.Content = text
		newPara.Runs = newPara.Runs[:1] // keep only the first run
	} else {
		// If the original paragraph has no runs, create a default one
		newPara.Runs = []Run{{
			Text: Text{Content: text},
		}}
	}

	return newPara
}

// createImageParagraph creates a paragraph containing an image.
func (te *TemplateEngine) createImageParagraph(imageData *TemplateImageData, doc *Document) (*Paragraph, error) {
	// Create image config
	config := imageData.Config
	if config == nil {
		config = &ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
		}
	}

	// Add image to document
	var imageInfo *ImageInfo
	var err error

	if len(imageData.Data) > 0 {
		// Use binary data
		var format ImageFormat
		format, err = detectImageFormat(imageData.Data)
		if err != nil {
			return nil, fmt.Errorf("failed to detect image format: %v", err)
		}

		var width, height int
		width, height, err = getImageDimensions(imageData.Data, format)
		if err != nil {
			return nil, fmt.Errorf("failed to get image dimensions: %v", err)
		}

		// Use a unique filename that includes the image ID counter
		fileName := fmt.Sprintf("image_%d.%s", doc.nextImageID, string(format))
		// Use the method that does not create a paragraph element; the template engine manages paragraphs
		imageInfo, err = doc.AddImageFromDataWithoutElement(imageData.Data, fileName, format, width, height, config)
	} else if imageData.FilePath != "" {
		// Use file path, but first read the data, then use AddImageFromDataWithoutElement
		data, readErr := os.ReadFile(imageData.FilePath)
		if readErr != nil {
			return nil, fmt.Errorf("failed to read image file: %v", readErr)
		}

		var format ImageFormat
		format, err = detectImageFormat(data)
		if err != nil {
			return nil, fmt.Errorf("failed to detect image format: %v", err)
		}

		var width, height int
		width, height, err = getImageDimensions(data, format)
		if err != nil {
			return nil, fmt.Errorf("failed to get image dimensions: %v", err)
		}

		fileName := filepath.Base(imageData.FilePath)
		imageInfo, err = doc.AddImageFromDataWithoutElement(data, fileName, format, width, height, config)
	} else {
		return nil, fmt.Errorf("both image data and file path are empty")
	}

	if err != nil {
		return nil, fmt.Errorf("failed to add image: %v", err)
	}

	// Set image description and title
	if imageData.AltText != "" {
		doc.SetImageAltText(imageInfo, imageData.AltText)
	}
	if imageData.Title != "" {
		doc.SetImageTitle(imageInfo, imageData.Title)
	}

	// Create paragraph containing the image
	imagePara := doc.createImageParagraph(imageInfo)
	return imagePara, nil
}
