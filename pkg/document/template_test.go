// Package document template functionality tests
package document

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

const testTitleKey = "title"

// TestNewTemplateEngine tests creating a template engine
func TestNewTemplateEngine(t *testing.T) {
	engine := NewTemplateEngine()
	if engine == nil {
		t.Fatal("Expected template engine to be created")
	}

	if engine.cache == nil {
		t.Fatal("Expected cache to be initialized")
	}
}

// TestTemplateVariableReplacement tests variable replacement functionality
func TestTemplateVariableReplacement(t *testing.T) {
	engine := NewTemplateEngine()

	// Create template with variables
	templateContent := "Hello {{name}}, welcome to {{company}}!"
	template, err := engine.LoadTemplate("test_template", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	// Verify template variable parsing
	if len(template.Variables) != 2 {
		t.Errorf("Expected 2 variables, got %d", len(template.Variables))
	}

	if _, exists := template.Variables["name"]; !exists {
		t.Error("Expected 'name' variable to be found")
	}

	if _, exists := template.Variables["company"]; !exists {
		t.Error("Expected 'company' variable to be found")
	}

	// Create template data
	data := NewTemplateData()
	data.SetVariable("name", "张三")
	data.SetVariable("company", "WordZero公司")

	// Render template
	doc, err := engine.RenderToDocument("test_template", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}

	// Check document content
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}
}

// TestTemplateConditionalStatements tests conditional statement functionality
func TestTemplateConditionalStatements(t *testing.T) {
	engine := NewTemplateEngine()

	// Create template with conditional statements
	templateContent := `{{#if showWelcome}}欢迎使用WordZero！{{/if}}
{{#if showDescription}}这是一个强大的Word文档操作库。{{/if}}`

	template, err := engine.LoadTemplate("conditional_template", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	// Verify conditional block parsing
	if len(template.Blocks) < 2 {
		t.Errorf("Expected at least 2 blocks, got %d", len(template.Blocks))
	}

	// Test condition being true
	data := NewTemplateData()
	data.SetCondition("showWelcome", true)
	data.SetCondition("showDescription", false)

	doc, err := engine.RenderToDocument("conditional_template", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}
}

// TestTemplateLoopStatements tests loop statement functionality
func TestTemplateLoopStatements(t *testing.T) {
	engine := NewTemplateEngine()

	// Create template with loop statements
	templateContent := `产品列表：
{{#each products}}
- 产品名称：{{name}}，价格：{{price}}元
{{/each}}`

	template, err := engine.LoadTemplate("loop_template", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	// Verify loop block parsing
	foundEachBlock := false
	for _, block := range template.Blocks {
		if block.Type == "each" && block.Variable == "products" {
			foundEachBlock = true
			break
		}
	}

	if !foundEachBlock {
		t.Error("Expected to find 'each products' block")
	}

	// Create list data
	data := NewTemplateData()
	products := []interface{}{
		map[string]interface{}{
			"name":  "iPhone",
			"price": 8999,
		},
		map[string]interface{}{
			"name":  "iPad",
			"price": 5999,
		},
	}
	data.SetList("products", products)

	doc, err := engine.RenderToDocument("loop_template", data)
	if err != nil {
		t.Fatalf("Failed to render template: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}
}

// TestTemplateInheritance tests template inheritance functionality
func TestTemplateInheritance(t *testing.T) {
	engine := NewTemplateEngine()

	// Create base template
	baseTemplateContent := `文档标题：{{title}}
基础内容：这是基础模板的内容。`

	_, err := engine.LoadTemplate("base_template", baseTemplateContent)
	if err != nil {
		t.Fatalf("Failed to load base template: %v", err)
	}

	// Create child template
	childTemplateContent := `{{extends "base_template"}}
扩展内容：这是子模板的内容。`

	childTemplate, err := engine.LoadTemplate("child_template", childTemplateContent)
	if err != nil {
		t.Fatalf("Failed to load child template: %v", err)
	}

	// Verify inheritance relationship
	if childTemplate.Parent == nil {
		t.Error("Expected child template to have parent")
	}

	if childTemplate.Parent.Name != "base_template" {
		t.Errorf("Expected parent template name to be 'base_template', got '%s'", childTemplate.Parent.Name)
	}
}

// TestTemplateValidation tests template validation functionality
func TestTemplateValidation(t *testing.T) {
	engine := NewTemplateEngine()

	// Test valid template
	validTemplate := `Hello {{name}}!
{{#if showMessage}}This is a message.{{/if}}
{{#each items}}Item: {{this}}{{/each}}`

	template, err := engine.LoadTemplate("valid_template", validTemplate)
	if err != nil {
		t.Fatalf("Failed to load valid template: %v", err)
	}

	err = engine.ValidateTemplate(template)
	if err != nil {
		t.Errorf("Expected valid template to pass validation, got error: %v", err)
	}

	// Test invalid template - mismatched brackets
	invalidTemplate1 := `Hello {{name}!`
	template1, err := engine.LoadTemplate("invalid_template1", invalidTemplate1)
	if err != nil {
		t.Fatalf("Failed to load invalid template: %v", err)
	}

	err = engine.ValidateTemplate(template1)
	if err == nil {
		t.Error("Expected invalid template (mismatched brackets) to fail validation")
	}

	// Test invalid template - mismatched if statements
	invalidTemplate2 := `{{#if condition}}Hello`
	template2, err := engine.LoadTemplate("invalid_template2", invalidTemplate2)
	if err != nil {
		t.Fatalf("Failed to load invalid template: %v", err)
	}

	err = engine.ValidateTemplate(template2)
	if err == nil {
		t.Error("Expected invalid template (mismatched if statements) to fail validation")
	}
}

// TestTemplateData tests template data functionality
func TestTemplateData(t *testing.T) {
	data := NewTemplateData()

	// Test setting and getting variables
	data.SetVariable("name", "测试")
	value, exists := data.GetVariable("name")
	if !exists {
		t.Error("Expected variable 'name' to exist")
	}
	if value != "测试" {
		t.Errorf("Expected variable value to be '测试', got '%v'", value)
	}

	// Test setting and getting lists
	items := []interface{}{"item1", "item2", "item3"}
	data.SetList("items", items)
	list, exists := data.GetList("items")
	if !exists {
		t.Error("Expected list 'items' to exist")
	}
	if len(list) != 3 {
		t.Errorf("Expected list length to be 3, got %d", len(list))
	}

	// Test setting and getting conditions
	data.SetCondition("enabled", true)
	condition, exists := data.GetCondition("enabled")
	if !exists {
		t.Error("Expected condition 'enabled' to exist")
	}
	if !condition {
		t.Error("Expected condition value to be true")
	}

	// Test batch setting variables
	variables := map[string]interface{}{
		testTitleKey: "测试标题",
		"content":    "测试内容",
	}
	data.SetVariables(variables)

	title, exists := data.GetVariable(testTitleKey)
	if !exists || title != "测试标题" {
		t.Error("Expected batch set variables to work")
	}
}

// TestTemplateDataFromStruct tests creating template data from a struct
func TestTemplateDataFromStruct(t *testing.T) {
	type TestStruct struct {
		Name    string
		Age     int
		Enabled bool
	}

	testData := TestStruct{
		Name:    "张三",
		Age:     30,
		Enabled: true,
	}

	templateData := NewTemplateData()
	err := templateData.FromStruct(testData)
	if err != nil {
		t.Fatalf("Failed to create template data from struct: %v", err)
	}

	// Verify variables were correctly set
	name, exists := templateData.GetVariable("name")
	if !exists || name != "张三" {
		t.Error("Expected 'name' variable to be set correctly")
	}

	age, exists := templateData.GetVariable("age")
	if !exists || age != 30 {
		t.Error("Expected 'age' variable to be set correctly")
	}

	enabled, exists := templateData.GetVariable("enabled")
	if !exists || enabled != true {
		t.Error("Expected 'enabled' variable to be set correctly")
	}
}

// TestTemplateMerge tests template data merging
func TestTemplateMerge(t *testing.T) {
	data1 := NewTemplateData()
	data1.SetVariable("name", "张三")
	data1.SetCondition("enabled", true)

	data2 := NewTemplateData()
	data2.SetVariable("age", 30)
	data2.SetList("items", []interface{}{"item1", "item2"})

	// Merge data
	data1.Merge(data2)

	// Verify merge result
	name, exists := data1.GetVariable("name")
	if !exists || name != "张三" {
		t.Error("Expected original variable to remain")
	}

	age, exists := data1.GetVariable("age")
	if !exists || age != 30 {
		t.Error("Expected merged variable to be present")
	}

	enabled, exists := data1.GetCondition("enabled")
	if !exists || !enabled {
		t.Error("Expected original condition to remain")
	}

	items, exists := data1.GetList("items")
	if !exists || len(items) != 2 {
		t.Error("Expected merged list to be present")
	}
}

// TestTemplateCache tests template caching functionality
func TestTemplateCache(t *testing.T) {
	engine := NewTemplateEngine()

	// Load template
	templateContent := "Hello {{name}}!"
	template1, err := engine.LoadTemplate("cached_template", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template: %v", err)
	}

	// Get template from cache
	template2, err := engine.GetTemplate("cached_template")
	if err != nil {
		t.Fatalf("Failed to get template from cache: %v", err)
	}

	// Verify same template instance
	if template1 != template2 {
		t.Error("Expected to get same template instance from cache")
	}

	// Clear cache
	engine.ClearCache()

	// Try to get cleared template
	_, err = engine.GetTemplate("cached_template")
	if err == nil {
		t.Error("Expected error when getting template after cache clear")
	}
}

// TestComplexTemplateRendering tests complex template rendering
func TestComplexTemplateRendering(t *testing.T) {
	engine := NewTemplateEngine()

	// Create complex template
	complexTemplate := `报告标题：{{title}}
作者：{{author}}

{{#if showSummary}}
概要：{{summary}}
{{/if}}

详细内容：
{{#each sections}}
章节 {{@index}}: {{title}}
内容：{{content}}

{{/each}}

{{#if showFooter}}
报告完毕。
{{/if}}`

	_, err := engine.LoadTemplate("complex_template", complexTemplate)
	if err != nil {
		t.Fatalf("Failed to load complex template: %v", err)
	}

	// Create complex data
	data := NewTemplateData()
	data.SetVariable("title", "WordZero功能测试报告")
	data.SetVariable("author", "开发团队")
	data.SetVariable("summary", "本报告测试了WordZero的模板功能。")

	data.SetCondition("showSummary", true)
	data.SetCondition("showFooter", true)

	sections := []interface{}{
		map[string]interface{}{
			"title":   "基础功能",
			"content": "测试了基础的文档操作功能。",
		},
		map[string]interface{}{
			"title":   "模板功能",
			"content": "测试了模板引擎的各种功能。",
		},
	}
	data.SetList("sections", sections)

	// Render complex template
	doc, err := engine.RenderToDocument("complex_template", data)
	if err != nil {
		t.Fatalf("Failed to render complex template: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}

	// Verify document has content
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}
}

// TestImagePlaceholder tests image placeholder functionality
//nolint:gocognit
func TestImagePlaceholder(t *testing.T) {
	engine := NewTemplateEngine()

	// Test basic image placeholder parsing
	t.Run("Parse image placeholder", func(t *testing.T) {
		templateContent := `文档标题

这里有一个图片：
{{#image testImage}}

更多内容...`

		template, err := engine.LoadTemplate("image_test", templateContent)
		if err != nil {
			t.Fatalf("failed to load template: %v", err)
		}

		// Check if image placeholder was parsed correctly
		hasImageBlock := false
		for _, block := range template.Blocks {
			if block.Type == "image" && block.Name == "testImage" {
				hasImageBlock = true
				break
			}
		}

		if !hasImageBlock {
			t.Error("template should contain an image block")
		}
	})

	// Test image placeholder rendering (string template)
	t.Run("Render image placeholder to string", func(t *testing.T) {
		templateContent := `产品介绍：{{productName}}

产品图片：
{{#image productImage}}

描述：{{description}}`

		_, err := engine.LoadTemplate("product", templateContent)
		if err != nil {
			t.Fatalf("failed to load template: %v", err)
		}

		data := NewTemplateData()
		data.SetVariable("productName", "测试产品")
		data.SetVariable("description", "这是一个测试产品")

		// Create image config
		imageConfig := &ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
			Size: &ImageSize{
				Width:           100,
				KeepAspectRatio: true,
			},
			AltText: "测试图片",
			Title:   "测试产品图片",
		}

		// Set image data (using example binary data)
		imageData := createTestImageData()
		data.SetImageFromData("productImage", imageData, imageConfig)

		// Render template
		doc, err := engine.RenderToDocument("product", data)
		if err != nil {
			t.Fatalf("failed to render template: %v", err)
		}

		if doc == nil {
			t.Error("render result should not be nil")
		}
	})

	// Test rendering image placeholder from document template
	t.Run("Render image placeholder from document template", func(t *testing.T) {
		// Create base document
		baseDoc := New()
		baseDoc.AddParagraph("报告标题：{{title}}")
		baseDoc.AddParagraph("{{#image reportChart}}")
		baseDoc.AddParagraph("总结：{{summary}}")

		// Create template from document
		template, err := engine.LoadTemplateFromDocument("report_template", baseDoc)
		if err != nil {
			t.Fatalf("failed to create template from document: %v", err)
		}

		if len(template.Variables) == 0 {
			t.Error("template should contain variables")
		}

		// Prepare data
		data := NewTemplateData()
		data.SetVariable("title", "月度报告")
		data.SetVariable("summary", "数据显示增长趋势良好")

		chartConfig := &ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
			Size: &ImageSize{
				Width: 120,
			},
		}

		imageData := createTestImageData()
		data.SetImageFromData("reportChart", imageData, chartConfig)

		// Use RenderTemplateToDocument method (recommended for document templates)
		doc, err := engine.RenderTemplateToDocument("report_template", data)
		if err != nil {
			t.Fatalf("failed to render document template: %v", err)
		}

		if doc == nil {
			t.Fatal("render result should not be nil")
		}

		// Check if document has elements
		if len(doc.Body.Elements) == 0 {
			t.Error("document should contain elements")
		}
	})

	// Test image data management methods
	t.Run("Test image data management", func(t *testing.T) {
		data := NewTemplateData()

		// Test SetImage method
		config := &ImageConfig{
			Position: ImagePositionInline,
			Size:     &ImageSize{Width: 50},
		}
		data.SetImage("test1", "path/to/image.jpg", config)

		// Test SetImageFromData method
		imageData := createTestImageData()
		data.SetImageFromData("test2", imageData, config)

		// Test SetImageWithDetails method
		data.SetImageWithDetails("test3", "path/to/image2.jpg", imageData, config, "alt text", testTitleKey)

		// Test GetImage method
		img1, exists1 := data.GetImage("test1")
		if !exists1 || img1.FilePath != "path/to/image.jpg" {
			t.Error("image 1 data is incorrect")
		}

		img2, exists2 := data.GetImage("test2")
		if !exists2 || len(img2.Data) == 0 {
			t.Error("image 2 data is incorrect")
		}

		img3, exists3 := data.GetImage("test3")
		if !exists3 || img3.AltText != "alt text" || img3.Title != testTitleKey {
			t.Error("image 3 data is incorrect")
		}

		// Test non-existent image
		_, exists4 := data.GetImage("nonexistent")
		if exists4 {
			t.Error("non-existent image should not return true")
		}
	})

	// Test image placeholder compatibility with other template syntax
	t.Run("Image placeholder compatibility with other syntax", func(t *testing.T) {
		templateContent := `{{#if showImage}}
图片标题：{{imageTitle}}
{{#image dynamicImage}}
{{/if}}

{{#each items}}
项目：{{name}}
{{#image itemImage}}
描述：{{description}}
{{/each}}`

		_, err := engine.LoadTemplate("complex", templateContent)
		if err != nil {
			t.Fatalf("failed to load complex template: %v", err)
		}

		data := NewTemplateData()
		data.SetCondition("showImage", true)
		data.SetVariable("imageTitle", "主要图片")

		items := []interface{}{
			map[string]interface{}{
				"name":        "项目1",
				"description": "项目1描述",
			},
		}
		data.SetList("items", items)

		config := &ImageConfig{Position: ImagePositionInline}
		imageData := createTestImageData()
		data.SetImageFromData("dynamicImage", imageData, config)
		data.SetImageFromData("itemImage", imageData, config)

		// Rendering should not produce an error
		doc, err := engine.RenderToDocument("complex", data)
		if err != nil {
			t.Fatalf("failed to render complex template: %v", err)
		}

		if doc == nil {
			t.Error("render result should not be nil")
		}
	})
}

// createTestImageData creates test image data
func createTestImageData() []byte {
	// Create a minimal PNG image data for testing
	return []byte{
		0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
		0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
		0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
		0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
		0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
		0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
		0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
		0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
		0x42, 0x60, 0x82,
	}
}

// TestNestedLoops tests nested loop functionality
func TestNestedLoops(t *testing.T) {
	engine := NewTemplateEngine()

	// Create template with nested loops
	templateContent := `会议纪要

日期：{{date}}

参会人员：
{{#each attendees}}
- {{name}} ({{role}})
  任务清单：
  {{#each tasks}}
  * {{taskName}} - 状态: {{status}}
  {{/each}}
{{/each}}

会议总结：{{summary}}`

	template, err := engine.LoadTemplate("meeting_minutes", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template with nested loops: %v", err)
	}

	if len(template.Blocks) < 1 {
		t.Error("Expected at least 1 block in template")
	}

	// Create nested data structure
	data := NewTemplateData()
	data.SetVariable("date", "2024-12-01")
	data.SetVariable("summary", "会议圆满结束")

	attendees := []interface{}{
		map[string]interface{}{
			"name": "张三",
			"role": "项目经理",
			"tasks": []interface{}{
				map[string]interface{}{
					"taskName": "制定项目计划",
					"status":   "进行中",
				},
				map[string]interface{}{
					"taskName": "分配资源",
					"status":   "已完成",
				},
			},
		},
		map[string]interface{}{
			"name": "李四",
			"role": "开发工程师",
			"tasks": []interface{}{
				map[string]interface{}{
					"taskName": "实现核心功能",
					"status":   "进行中",
				},
				map[string]interface{}{
					"taskName": "编写单元测试",
					"status":   "待开始",
				},
			},
		},
	}
	data.SetList("attendees", attendees)

	// Render template
	doc, err := engine.RenderToDocument("meeting_minutes", data)
	if err != nil {
		t.Fatalf("Failed to render template with nested loops: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}

	// Verify document content
	if len(doc.Body.Elements) == 0 {
		t.Error("Expected document to have content")
	}

	// Check if generated content contains expected nested data
	foundNestedContent := false
	for _, element := range doc.Body.Elements {
		if para, ok := element.(*Paragraph); ok {
			fullText := ""
			for _, run := range para.Runs {
				fullText += run.Text.Content
			}

			// Check for content generated by nested loops (task names)
			if fullText == "  * 制定项目计划 - 状态: 进行中" ||
				fullText == "  * 实现核心功能 - 状态: 进行中" {
				foundNestedContent = true
			}

			// Ensure no unprocessed template syntax remains
			if fullText == "{{#each tasks}}" || fullText == "  * {{taskName}} - 状态: {{status}}" {
				t.Errorf("Found unprocessed template syntax in output: %s", fullText)
			}
		}
	}

	if !foundNestedContent {
		t.Error("Expected to find nested loop content in rendered document")
	}
}

// TestDeepNestedLoops tests deep nested loops (three levels)
func TestDeepNestedLoops(t *testing.T) {
	engine := NewTemplateEngine()

	// Create template with three-level nested loops
	templateContent := `组织架构：
{{#each departments}}
部门：{{name}}
{{#each teams}}
  团队：{{teamName}}
  {{#each members}}
    成员：{{memberName}} - {{position}}
  {{/each}}
{{/each}}
{{/each}}`

	_, err := engine.LoadTemplate("org_structure", templateContent)
	if err != nil {
		t.Fatalf("Failed to load template with deep nested loops: %v", err)
	}

	// Create three-level nested data
	data := NewTemplateData()

	departments := []interface{}{
		map[string]interface{}{
			"name": "技术部",
			"teams": []interface{}{
				map[string]interface{}{
					"teamName": "前端团队",
					"members": []interface{}{
						map[string]interface{}{
							"memberName": "王五",
							"position":   "前端工程师",
						},
						map[string]interface{}{
							"memberName": "赵六",
							"position":   "UI设计师",
						},
					},
				},
				map[string]interface{}{
					"teamName": "后端团队",
					"members": []interface{}{
						map[string]interface{}{
							"memberName": "孙七",
							"position":   "后端工程师",
						},
					},
				},
			},
		},
	}
	data.SetList("departments", departments)

	// Render template
	doc, err := engine.RenderToDocument("org_structure", data)
	if err != nil {
		t.Fatalf("Failed to render template with deep nested loops: %v", err)
	}

	if doc == nil {
		t.Fatal("Expected document to be created")
	}

	// Verify third-level nested content is correctly rendered
	foundDeepContent := false
	for _, element := range doc.Body.Elements {
		if para, ok := element.(*Paragraph); ok {
			fullText := ""
			for _, run := range para.Runs {
				fullText += run.Text.Content
			}

			// Check third-level nested content
			if fullText == "    成员：王五 - 前端工程师" ||
				fullText == "    成员：孙七 - 后端工程师" {
				foundDeepContent = true
			}

			// Ensure no unprocessed template syntax remains
			if fullText == "{{#each members}}" || fullText == "    成员：{{memberName}} - {{position}}" {
				t.Errorf("Found unprocessed template syntax in deep nested output: %s", fullText)
			}
		}
	}

	if !foundDeepContent {
		t.Error("Expected to find deep nested loop content in rendered document")
	}
}

// TestHeaderFooterTemplateVariables tests template variable identification and replacement in headers/footers
func TestHeaderFooterTemplateVariables(t *testing.T) {
	// Create document with headers/footers
	doc := New()

	// Add body content
	doc.AddParagraph("{{title}}")
	doc.AddParagraph("文档内容")

	// Add header with template variables
	err := doc.AddHeader(HeaderFooterTypeDefault, "{{headerTitle}} - {{headerID}}")
	if err != nil {
		t.Fatalf("failed to add header: %v", err)
	}

	// Add footer with template variables
	err = doc.AddFooter(HeaderFooterTypeDefault, "{{footerText}} - 第 {{pageNum}} 页")
	if err != nil {
		t.Fatalf("failed to add footer: %v", err)
	}

	// Create template engine and load document as template
	engine := NewTemplateEngine()
	template, err := engine.LoadTemplateFromDocument("header_footer_test", doc)
	if err != nil {
		t.Fatalf("failed to load template from document: %v", err)
	}

	// Verify template variables were correctly identified
	expectedVars := []string{"title", "headerTitle", "headerID", "footerText", "pageNum"}
	for _, varName := range expectedVars {
		if _, exists := template.Variables[varName]; !exists {
			t.Errorf("template variable '%s' should be identified but was not found", varName)
		}
	}

	// Test using TemplateRenderer to analyze template with headers/footers
	// Create a new document with headers/footers for testing analysis functionality
	doc2 := New()
	doc2.AddParagraph("{{mainContent}}")
	err = doc2.AddHeader(HeaderFooterTypeDefault, "{{documentTitle}}")
	if err != nil {
		t.Fatalf("failed to add header: %v", err)
	}

	// Load via engine
	engine2 := NewTemplateEngine()
	_, err = engine2.LoadTemplateFromDocument("analyze_test", doc2)
	if err != nil {
		t.Fatalf("failed to load template from document: %v", err)
	}

	// Create renderer and use the loaded template
	renderer := &TemplateRenderer{
		engine: engine2,
		logger: &TemplateLogger{enabled: false},
	}

	// Analyze template
	analysis, err := renderer.AnalyzeTemplate("analyze_test")
	if err != nil {
		t.Fatalf("failed to analyze template: %v", err)
	}

	// Verify analysis result contains header variable
	if _, exists := analysis.Variables["documentTitle"]; !exists {
		t.Error("analysis result should contain header variable 'documentTitle'")
	}
	if _, exists := analysis.Variables["mainContent"]; !exists {
		t.Error("analysis result should contain body variable 'mainContent'")
	}

	t.Logf("analyzed variables: %v", analysis.Variables)
}

// TestHeaderFooterVariableReplacement tests variable replacement in headers/footers
func TestHeaderFooterVariableReplacement(t *testing.T) {
	// Create document with headers/footers
	doc := New()

	// Add body content
	doc.AddParagraph("{{title}}")
	doc.AddParagraph("正文内容")

	// Add header with template variables
	err := doc.AddHeader(HeaderFooterTypeDefault, "报告编号: {{reportID}}")
	if err != nil {
		t.Fatalf("failed to add header: %v", err)
	}

	// Add footer with template variables
	err = doc.AddFooter(HeaderFooterTypeDefault, "作者: {{author}}")
	if err != nil {
		t.Fatalf("failed to add footer: %v", err)
	}

	// Create template engine and load document as template
	engine := NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("replacement_test", doc)
	if err != nil {
		t.Fatalf("failed to load template from document: %v", err)
	}

	// Prepare template data
	data := NewTemplateData()
	data.SetVariable("title", "测试报告标题")
	data.SetVariable("reportID", "RPT-2024-001")
	data.SetVariable("author", "测试作者")

	// Render template
	resultDoc, err := engine.RenderTemplateToDocument("replacement_test", data)
	if err != nil {
		t.Fatalf("failed to render template: %v", err)
	}

	// Verify variables in header were replaced
	headerReplaced := false
	footerReplaced := false

	for partName, partData := range resultDoc.parts {
		content := string(partData)

		if partName == "word/header1.xml" {
			if !strings.Contains(content, "{{reportID}}") && strings.Contains(content, "RPT-2024-001") {
				headerReplaced = true
			}
			t.Logf("header content: %s", content)
		}

		if partName == "word/footer1.xml" {
			if !strings.Contains(content, "{{author}}") && strings.Contains(content, "测试作者") {
				footerReplaced = true
			}
			t.Logf("footer content: %s", content)
		}
	}

	if !headerReplaced {
		t.Error("variables in header should be replaced")
	}

	if !footerReplaced {
		t.Error("variables in footer should be replaced")
	}
}

// TestTemplateFromFileWithParagraphSectionProperties ensures section properties within paragraphs still preserve header/footer
func TestTemplateFromFileWithParagraphSectionProperties(t *testing.T) {
	doc := New()
	doc.AddParagraph("{{title}}")

	if err := doc.AddHeader(HeaderFooterTypeDefault, "报告编号: {{reportID}}"); err != nil {
		t.Fatalf("failed to add header: %v", err)
	}
	if err := doc.AddFooter(HeaderFooterTypeDefault, "撰写人: {{author}}"); err != nil {
		t.Fatalf("failed to add footer: %v", err)
	}

	sectionMarker := "__SECTION_BREAK__"
	doc.AddParagraph(sectionMarker)

	tmpDir := t.TempDir()
	basePath := filepath.Join(tmpDir, "base_paragraph_section.docx")
	if err := doc.Save(basePath); err != nil {
		t.Fatalf("failed to save base document: %v", err)
	}

	modifiedPath := filepath.Join(tmpDir, "paragraph_section_template.docx")
	if err := moveSectPrIntoParagraph(basePath, modifiedPath, sectionMarker); err != nil {
		t.Fatalf("failed to adjust section properties position: %v", err)
	}

	loadedDoc, err := Open(modifiedPath)
	if err != nil {
		t.Fatalf("failed to open modified document: %v", err)
	}

	engine := NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("paragraph_section_template", loadedDoc)
	if err != nil {
		t.Fatalf("failed to load template: %v", err)
	}

	data := NewTemplateData()
	data.SetVariable("title", "段落节属性测试")
	data.SetVariable("reportID", "RPT-2024-009")
	data.SetVariable("author", "测试作者")

	renderedDoc, err := engine.RenderTemplateToDocument("paragraph_section_template", data)
	if err != nil {
		t.Fatalf("failed to render template: %v", err)
	}

	headerContent := string(renderedDoc.parts["word/header1.xml"])
	if strings.Contains(headerContent, "{{reportID}}") {
		t.Error("variables in header should be replaced, even when section properties are inside paragraph")
	}
	if !strings.Contains(headerContent, "RPT-2024-009") {
		t.Error("header is missing replaced value")
	}

	footerContent := string(renderedDoc.parts["word/footer1.xml"])
	if strings.Contains(footerContent, "{{author}}") {
		t.Error("variables in footer should be replaced")
	}
	if !strings.Contains(footerContent, "测试作者") {
		t.Error("footer is missing replaced value")
	}
}

func moveSectPrIntoParagraph(srcPath, dstPath, marker string) error {
	reader, err := zip.OpenReader(srcPath)
	if err != nil {
		return err
	}
	defer reader.Close()

	output, err := os.Create(dstPath)
	if err != nil {
		return err
	}
	defer output.Close()

	zipWriter := zip.NewWriter(output)
	defer zipWriter.Close()

	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			return err
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return err
		}

		if file.Name == "word/document.xml" {
			data, err = rewriteSectPrIntoParagraph(data, marker)
			if err != nil {
				return err
			}
		}

		writer, err := zipWriter.Create(file.Name)
		if err != nil {
			return err
		}
		if _, err := writer.Write(data); err != nil {
			return err
		}
	}

	return nil
}

func rewriteSectPrIntoParagraph(xmlData []byte, marker string) ([]byte, error) {
	content := string(xmlData)
	sectStart := strings.Index(content, "<w:sectPr")
	if sectStart == -1 {
		return nil, fmt.Errorf("sectPr not found")
	}

	sectEndRel := strings.Index(content[sectStart:], "</w:sectPr>")
	if sectEndRel == -1 {
		return nil, fmt.Errorf("sectPr missing closing tag")
	}
	sectEnd := sectStart + sectEndRel + len("</w:sectPr>")
	sectBlock := content[sectStart:sectEnd]

	sanitized := removeHeaderFooterReferences(sectBlock)
	content = content[:sectStart] + sanitized + content[sectEnd:]

	markerIndex := strings.Index(content, marker)
	if markerIndex == -1 {
		return nil, fmt.Errorf("marker paragraph not found")
	}

	pStart := strings.LastIndex(content[:markerIndex], "<w:p")
	if pStart == -1 {
		return nil, fmt.Errorf("paragraph start tag not found")
	}
	openEnd := strings.Index(content[pStart:], ">")
	if openEnd == -1 {
		return nil, fmt.Errorf("paragraph tag not closed")
	}
	insertPos := pStart + openEnd + 1

	insert := "<w:pPr>" + sectBlock + "</w:pPr>"
	modified := content[:insertPos] + insert + content[insertPos:]

	return []byte(modified), nil
}

func removeHeaderFooterReferences(block string) string {
	block = stripReferenceTag(block, "<w:headerReference")
	block = stripReferenceTag(block, "<w:footerReference")
	return block
}

func stripReferenceTag(block, tag string) string {
	for {
		start := strings.Index(block, tag)
		if start == -1 {
			break
		}
		end := strings.Index(block[start:], "/>")
		if end == -1 {
			break
		}
		block = block[:start] + block[start+end+2:]
	}
	return block
}

// TestTemplateDocumentPartsPreservation tests complete preservation of document parts during template rendering
func TestTemplateDocumentPartsPreservation(t *testing.T) {
	// Create source document with multiple document parts
	doc := New()

	// Add header and footer
	err := doc.AddHeader(HeaderFooterTypeDefault, "Template Header - {{headerVar}}")
	if err != nil {
		t.Fatalf("failed to add header: %v", err)
	}

	err = doc.AddFooter(HeaderFooterTypeDefault, "Template Footer - {{footerVar}}")
	if err != nil {
		t.Fatalf("failed to add footer: %v", err)
	}

	// Set page settings
	settings := DefaultPageSettings()
	settings.Size = PageSizeA4
	settings.Orientation = OrientationPortrait
	err = doc.SetPageSettings(settings)
	if err != nil {
		t.Fatalf("failed to set page settings: %v", err)
	}

	// Add heading and content
	doc.AddHeadingParagraph("{{docTitle}}", 1)
	doc.AddParagraph("Content with {{variable1}} and more text.")

	// Save original document
	originalPath := "test_parts_preservation_original.docx"
	err = doc.Save(originalPath)
	if err != nil {
		t.Fatalf("failed to save original document: %v", err)
	}
	defer func() {
		if err := os.Remove(originalPath); err != nil {
			t.Logf("failed to clean up original document: %v", err)
		}
	}()

	// Open original document as template
	templateDoc, err := Open(originalPath)
	if err != nil {
		t.Fatalf("failed to open template document: %v", err)
	}

	// Record original document parts
	originalParts := make(map[string]bool)
	for partName := range templateDoc.parts {
		originalParts[partName] = true
	}

	// Create template engine and load template
	engine := NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("parts_test", templateDoc)
	if err != nil {
		t.Fatalf("failed to load template: %v", err)
	}

	// Render template
	data := NewTemplateData()
	data.SetVariable("headerVar", "Header Value")
	data.SetVariable("footerVar", "Footer Value")
	data.SetVariable("docTitle", "Document Title")
	data.SetVariable("variable1", "Variable 1 Value")

	renderedDoc, err := engine.RenderTemplateToDocument("parts_test", data)
	if err != nil {
		t.Fatalf("failed to render template: %v", err)
	}

	// Save rendered document
	renderedPath := "test_parts_preservation_rendered.docx"
	err = renderedDoc.Save(renderedPath)
	if err != nil {
		t.Fatalf("failed to save rendered document: %v", err)
	}
	defer func() {
		if err := os.Remove(renderedPath); err != nil {
			t.Logf("failed to clean up rendered document: %v", err)
		}
	}()

	// Check rendered document parts
	renderedParts := make(map[string]bool)
	for partName := range renderedDoc.parts {
		renderedParts[partName] = true
	}

	// Verify critical parts are preserved
	criticalParts := []string{
		"word/styles.xml",
		"word/header1.xml",
		"word/footer1.xml",
	}

	for _, part := range criticalParts {
		if originalParts[part] && !renderedParts[part] {
			t.Errorf("critical part %s exists in original document but is missing in rendered document", part)
		}
	}

	// Verify header/footer variables were replaced
	headerContent := string(renderedDoc.parts["word/header1.xml"])
	if strings.Contains(headerContent, "{{headerVar}}") {
		t.Error("variables in header should be replaced")
	}
	if !strings.Contains(headerContent, "Header Value") {
		t.Error("header should contain the replaced value")
	}

	footerContent := string(renderedDoc.parts["word/footer1.xml"])
	if strings.Contains(footerContent, "{{footerVar}}") {
		t.Error("variables in footer should be replaced")
	}
	if !strings.Contains(footerContent, "Footer Value") {
		t.Error("footer should contain the replaced value")
	}

	t.Log("document parts preservation test passed")
}

// TestTemplateSectionPropertiesPreservation tests section properties preservation during template rendering
func TestTemplateSectionPropertiesPreservation(t *testing.T) {
	// Create source document with section properties
	doc := New()

	// Set page settings (this creates SectionProperties)
	settings := DefaultPageSettings()
	settings.Size = PageSizeA4
	settings.MarginTop = 30.0
	settings.MarginBottom = 25.0
	settings.MarginLeft = 20.0
	settings.MarginRight = 20.0
	err := doc.SetPageSettings(settings)
	if err != nil {
		t.Fatalf("failed to set page settings: %v", err)
	}

	// Add content
	doc.AddParagraph("Content with {{variable}}")

	// Save original document
	originalPath := "test_section_props_original.docx"
	err = doc.Save(originalPath)
	if err != nil {
		t.Fatalf("failed to save original document: %v", err)
	}
	defer func() {
		if err := os.Remove(originalPath); err != nil {
			t.Logf("failed to clean up original document: %v", err)
		}
	}()

	// Open original document as template
	templateDoc, err := Open(originalPath)
	if err != nil {
		t.Fatalf("failed to open template document: %v", err)
	}

	// Get original document page settings
	originalSettings := templateDoc.GetPageSettings()

	// Create template engine and load template
	engine := NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("section_test", templateDoc)
	if err != nil {
		t.Fatalf("failed to load template: %v", err)
	}

	// Render template
	data := NewTemplateData()
	data.SetVariable("variable", "Value")

	renderedDoc, err := engine.RenderTemplateToDocument("section_test", data)
	if err != nil {
		t.Fatalf("failed to render template: %v", err)
	}

	// Get rendered document page settings
	renderedSettings := renderedDoc.GetPageSettings()

	// Verify page settings are preserved
	if renderedSettings.Size != originalSettings.Size {
		t.Errorf("page size mismatch: expected %v, got %v", originalSettings.Size, renderedSettings.Size)
	}

	// Allow 1mm tolerance
	tolerance := 1.0
	if abs(renderedSettings.MarginTop-originalSettings.MarginTop) > tolerance {
		t.Errorf("top margin mismatch: expected %.1f, got %.1f", originalSettings.MarginTop, renderedSettings.MarginTop)
	}
	if abs(renderedSettings.MarginBottom-originalSettings.MarginBottom) > tolerance {
		t.Errorf("bottom margin mismatch: expected %.1f, got %.1f", originalSettings.MarginBottom, renderedSettings.MarginBottom)
	}
	if abs(renderedSettings.MarginLeft-originalSettings.MarginLeft) > tolerance {
		t.Errorf("left margin mismatch: expected %.1f, got %.1f", originalSettings.MarginLeft, renderedSettings.MarginLeft)
	}
	if abs(renderedSettings.MarginRight-originalSettings.MarginRight) > tolerance {
		t.Errorf("right margin mismatch: expected %.1f, got %.1f", originalSettings.MarginRight, renderedSettings.MarginRight)
	}

	t.Log("section properties preservation test passed")
}

// TestTemplateNumberingPropertiesPreservation tests numbering properties preservation during template rendering
func TestTemplateNumberingPropertiesPreservation(t *testing.T) {
	// Create document with numbered paragraphs
	doc := New()

	// Add numbered list items
	config := &ListConfig{
		Type:        ListTypeNumber,
		IndentLevel: 0,
		StartNumber: 1,
	}
	doc.AddListItem("第一条 {{itemTitle}}", config)
	doc.AddListItem("第二条 {{itemContent}}", config)

	// Save original document
	originalPath := "test_numbering_preservation_original.docx"
	err := doc.Save(originalPath)
	if err != nil {
		t.Fatalf("failed to save original document: %v", err)
	}
	defer func() {
		if err := os.Remove(originalPath); err != nil {
			t.Logf("failed to clean up original document: %v", err)
		}
	}()

	// Open original document as template
	templateDoc, err := Open(originalPath)
	if err != nil {
		t.Fatalf("failed to open template document: %v", err)
	}

	// Verify original document numbering properties were parsed correctly
	paragraphs := templateDoc.Body.GetParagraphs()
	if len(paragraphs) < 2 {
		t.Fatalf("expected at least 2 paragraphs, got %d", len(paragraphs))
	}

	// Check first paragraph numbering properties
	if paragraphs[0].Properties == nil || paragraphs[0].Properties.NumberingProperties == nil {
		t.Error("first paragraph numbering properties should be parsed")
	}

	// Create template engine and load template
	engine := NewTemplateEngine()
	_, err = engine.LoadTemplateFromDocument("numbering_test", templateDoc)
	if err != nil {
		t.Fatalf("failed to load template: %v", err)
	}

	// Render template
	data := NewTemplateData()
	data.SetVariable("itemTitle", "合作项目情况")
	data.SetVariable("itemContent", "合作项目背景")

	renderedDoc, err := engine.RenderTemplateToDocument("numbering_test", data)
	if err != nil {
		t.Fatalf("failed to render template: %v", err)
	}

	// Save rendered document
	renderedPath := "test_numbering_preservation_rendered.docx"
	err = renderedDoc.Save(renderedPath)
	if err != nil {
		t.Fatalf("failed to save rendered document: %v", err)
	}
	defer func() {
		if err := os.Remove(renderedPath); err != nil {
			t.Logf("failed to clean up rendered document: %v", err)
		}
	}()

	// Verify numbering properties are preserved in rendered document
	renderedParagraphs := renderedDoc.Body.GetParagraphs()
	if len(renderedParagraphs) < 2 {
		t.Fatalf("expected at least 2 paragraphs after rendering, got %d", len(renderedParagraphs))
	}

	// Check rendered paragraph numbering properties are preserved
	for i, para := range renderedParagraphs[:2] {
		if para.Properties == nil {
			t.Errorf("paragraph %d properties should not be nil", i+1)
			continue
		}
		if para.Properties.NumberingProperties == nil {
			t.Errorf("paragraph %d numbering properties should not be nil", i+1)
			continue
		}
		if para.Properties.NumberingProperties.NumID == nil {
			t.Errorf("paragraph %d numbering ID should not be nil", i+1)
		}
		if para.Properties.NumberingProperties.ILevel == nil {
			t.Errorf("paragraph %d numbering level should not be nil", i+1)
		}
	}

	// Verify variables were replaced
	firstParaText := ""
	for _, run := range renderedParagraphs[0].Runs {
		firstParaText += run.Text.Content
	}
	if !strings.Contains(firstParaText, "合作项目情况") {
		t.Errorf("first paragraph should contain replaced variable value, actual content: %s", firstParaText)
	}

	t.Log("numbering properties preservation test passed")
}
