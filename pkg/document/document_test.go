package document

import (
	"os"
	"testing"

	"github.com/mr-pmillz/wordZero/pkg/style"
)

const (
	testColorBlue      = "0000FF"
	testHeaderPartName = "word/header1.xml"
)

// assertParagraphContent verifies the text content of a paragraph with bounds checking
func assertParagraphContent(t *testing.T, paragraphs []*Paragraph, index int, expectedContent string) {
	t.Helper()
	if index >= len(paragraphs) {
		t.Errorf("paragraph index %d out of range, only %d paragraphs total", index, len(paragraphs))
		return
	}

	para := paragraphs[index]
	if len(para.Runs) == 0 {
		t.Errorf("paragraph at index %d has no runs", index)
		return
	}

	actualContent := para.Runs[0].Text.Content
	if actualContent != expectedContent {
		t.Errorf("paragraph at index %d content should be '%s', got '%s'", index, expectedContent, actualContent)
	}
}

// TestNewDocument tests new document creation
func TestNewDocument(t *testing.T) {
	doc := New()

	// Verify basic structure
	if doc == nil {
		t.Fatal("Failed to create new document")
	}

	if doc.Body == nil {
		t.Fatal("Document body is nil")
	}

	if doc.styleManager == nil {
		t.Fatal("Style manager is nil")
	}

	// Verify initial state
	if len(doc.Body.GetParagraphs()) != 0 {
		t.Errorf("Expected 0 paragraphs, got %d", len(doc.Body.GetParagraphs()))
	}

	// Verify style manager initialization
	styles := doc.styleManager.GetAllStyles()
	if len(styles) == 0 {
		t.Error("Style manager should have predefined styles")
	}
}

// TestAddParagraph tests adding a plain paragraph
func TestAddParagraph(t *testing.T) {
	doc := New()
	text := "测试段落内容"

	para := doc.AddParagraph(text)

	// Verify paragraph was added
	if len(doc.Body.GetParagraphs()) != 1 {
		t.Errorf("Expected 1 paragraph, got %d", len(doc.Body.GetParagraphs()))
	}

	// Verify paragraph content
	if len(para.Runs) != 1 {
		t.Errorf("Expected 1 run, got %d", len(para.Runs))
	}

	if para.Runs[0].Text.Content != text {
		t.Errorf("Expected %s, got %s", text, para.Runs[0].Text.Content)
	}

	// Verify the returned pointer is correct
	paragraphs := doc.Body.GetParagraphs()
	if paragraphs[0] != para {
		t.Error("Returned paragraph pointer is incorrect")
	}
}

// TestAddHeadingParagraph tests adding heading paragraphs
func TestAddHeadingParagraph(t *testing.T) {
	doc := New()

	testCases := []struct {
		text    string
		level   int
		styleID string
	}{
		{"第一级标题", 1, styleHeading1},
		{"第二级标题", 2, "Heading2"},
		{"第三级标题", 3, "Heading3"},
		{"第九级标题", 9, "Heading9"},
	}

	for _, tc := range testCases {
		para := doc.AddHeadingParagraph(tc.text, tc.level)

		// Verify paragraph style is set
		if para.Properties == nil {
			t.Errorf("Heading paragraph should have properties")
			continue
		}

		if para.Properties.ParagraphStyle == nil {
			t.Errorf("Heading paragraph should have style reference")
			continue
		}

		if para.Properties.ParagraphStyle.Val != tc.styleID {
			t.Errorf("Expected style %s, got %s", tc.styleID, para.Properties.ParagraphStyle.Val)
		}

		// Verify content
		if len(para.Runs) != 1 {
			t.Errorf("Expected 1 run, got %d", len(para.Runs))
			continue
		}

		if para.Runs[0].Text.Content != tc.text {
			t.Errorf("Expected %s, got %s", tc.text, para.Runs[0].Text.Content)
		}
	}

	// Test out-of-range level
	para := doc.AddHeadingParagraph("超出范围", 10)
	if para.Properties.ParagraphStyle.Val != styleHeading1 {
		t.Error("Out of range level should default to Heading1")
	}

	para = doc.AddHeadingParagraph("负数级别", -1)
	if para.Properties.ParagraphStyle.Val != styleHeading1 {
		t.Error("Negative level should default to Heading1")
	}
}

// TestAddFormattedParagraph tests adding a formatted paragraph
func TestAddFormattedParagraph(t *testing.T) {
	doc := New()
	text := "格式化文本"

	format := &TextFormat{
		Bold:       true,
		Italic:     true,
		FontSize:   14,
		FontColor:  "FF0000",
		FontFamily: "宋体",
	}

	para := doc.AddFormattedParagraph(text, format)

	// Verify paragraph was added
	if len(doc.Body.GetParagraphs()) != 1 {
		t.Error("Failed to add formatted paragraph")
	}

	// Verify formatting
	run := para.Runs[0]
	if run.Properties == nil {
		t.Fatal("Run properties should not be nil")
	}

	if run.Properties.Bold == nil {
		t.Error("Bold property should be set")
	}

	if run.Properties.Italic == nil {
		t.Error("Italic property should be set")
	}

	if run.Properties.FontSize == nil || run.Properties.FontSize.Val != "28" {
		t.Errorf("Expected font size 28, got %v", run.Properties.FontSize)
	}

	if run.Properties.Color == nil || run.Properties.Color.Val != "FF0000" {
		t.Errorf("Expected color FF0000, got %v", run.Properties.Color)
	}

	if run.Properties.FontFamily == nil || run.Properties.FontFamily.ASCII != "宋体" {
		t.Errorf("Expected font family 宋体, got %v", run.Properties.FontFamily)
	}
}

// TestParagraphSetAlignment tests paragraph alignment setting
func TestParagraphSetAlignment(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试对齐")

	testCases := []AlignmentType{
		AlignLeft,
		AlignCenter,
		AlignRight,
		AlignJustify,
	}

	for _, alignment := range testCases {
		para.SetAlignment(alignment)

		if para.Properties == nil {
			t.Fatal("Properties should not be nil after setting alignment")
		}

		if para.Properties.Justification == nil {
			t.Fatal("Justification should not be nil")
		}

		if para.Properties.Justification.Val != string(alignment) {
			t.Errorf("Expected alignment %s, got %s", alignment, para.Properties.Justification.Val)
		}
	}
}

// TestParagraphSetSpacing tests paragraph spacing setting
func TestParagraphSetSpacing(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试间距")

	config := &SpacingConfig{
		LineSpacing:     1.5,
		BeforePara:      12,
		AfterPara:       6,
		FirstLineIndent: 24,
	}

	para.SetSpacing(config)

	// Verify property settings
	if para.Properties == nil {
		t.Fatal("Properties should not be nil")
	}

	if para.Properties.Spacing == nil {
		t.Fatal("Spacing should not be nil")
	}

	// Verify spacing values (converted to TWIPs)
	spacing := para.Properties.Spacing
	if spacing.Before != "240" { // 12 * 20
		t.Errorf("Expected before spacing 240, got %s", spacing.Before)
	}

	if spacing.After != "120" { // 6 * 20
		t.Errorf("Expected after spacing 120, got %s", spacing.After)
	}

	if spacing.Line != "360" { // 1.5 * 240
		t.Errorf("Expected line spacing 360, got %s", spacing.Line)
	}

	// Verify first line indent
	if para.Properties.Indentation == nil {
		t.Fatal("Indentation should not be nil")
	}

	if para.Properties.Indentation.FirstLine != "480" { // 24 * 20
		t.Errorf("Expected first line indent 480, got %s", para.Properties.Indentation.FirstLine)
	}
}

// TestParagraphAddFormattedText tests adding formatted text to a paragraph
func TestParagraphAddFormattedText(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("初始文本")

	// Add formatted text
	format := &TextFormat{
		Bold:      true,
		FontColor: testColorBlue,
	}

	para.AddFormattedText("格式化文本", format)

	// Verify run count
	if len(para.Runs) != 2 {
		t.Errorf("Expected 2 runs, got %d", len(para.Runs))
	}

	// Verify the second run's formatting
	run := para.Runs[1]
	if run.Properties == nil {
		t.Fatal("Second run should have properties")
	}

	if run.Properties.Bold == nil {
		t.Error("Second run should be bold")
	}

	if run.Properties.Color == nil || run.Properties.Color.Val != testColorBlue {
		t.Error("Second run should be blue")
	}

	if run.Text.Content != "格式化文本" {
		t.Errorf("Expected '格式化文本', got '%s'", run.Text.Content)
	}
}

// TestParagraphSetStyle tests paragraph style setting
func TestParagraphSetStyle(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试样式")

	para.SetStyle(styleHeading1)

	if para.Properties == nil {
		t.Fatal("Properties should not be nil")
	}

	if para.Properties.ParagraphStyle == nil {
		t.Fatal("ParagraphStyle should not be nil")
	}

	if para.Properties.ParagraphStyle.Val != styleHeading1 {
		t.Errorf("Expected style Heading1, got %s", para.Properties.ParagraphStyle.Val)
	}
}

// TestParagraphSetIndentation tests paragraph indentation setting
func TestParagraphSetIndentation(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试缩进")

	// Test first line indent
	para.SetIndentation(0.5, 0, 0)

	if para.Properties == nil {
		t.Fatal("Properties should not be nil")
	}

	if para.Properties.Indentation == nil {
		t.Fatal("Indentation should not be nil")
	}

	// 0.5 cm = 283.5 TWIPs, rounded to 284
	expectedFirstLine := "283"
	if para.Properties.Indentation.FirstLine != expectedFirstLine {
		t.Errorf("Expected FirstLine %s, got %s", expectedFirstLine, para.Properties.Indentation.FirstLine)
	}

	// Test left and right indent
	para.SetIndentation(-0.5, 1.0, 0.5)

	expectedFirstLine = "-283" // hanging indent
	expectedLeft := "567"      // 1 cm
	expectedRight := "283"     // 0.5 cm

	if para.Properties.Indentation.FirstLine != expectedFirstLine {
		t.Errorf("Expected FirstLine %s, got %s", expectedFirstLine, para.Properties.Indentation.FirstLine)
	}
	if para.Properties.Indentation.Left != expectedLeft {
		t.Errorf("Expected Left %s, got %s", expectedLeft, para.Properties.Indentation.Left)
	}
	if para.Properties.Indentation.Right != expectedRight {
		t.Errorf("Expected Right %s, got %s", expectedRight, para.Properties.Indentation.Right)
	}
}

// TestParagraphSetKeepWithNext tests keeping a paragraph with the next one
func TestParagraphSetKeepWithNext(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试保持与下一段")

	// Test enabling
	para.SetKeepWithNext(true)

	if para.Properties == nil {
		t.Fatal("Properties should not be nil")
	}

	if para.Properties.KeepNext == nil {
		t.Fatal("KeepNext should not be nil")
	}

	if para.Properties.KeepNext.Val != "1" {
		t.Errorf("Expected KeepNext Val to be '1', got '%s'", para.Properties.KeepNext.Val)
	}

	// Test disabling
	para.SetKeepWithNext(false)

	if para.Properties.KeepNext != nil {
		t.Error("KeepNext should be nil when disabled")
	}
}

// TestParagraphSetKeepLines tests keeping paragraph lines together
func TestParagraphSetKeepLines(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试行保持")

	// Test enabling
	para.SetKeepLines(true)

	if para.Properties == nil {
		t.Fatal("Properties should not be nil")
	}

	if para.Properties.KeepLines == nil {
		t.Fatal("KeepLines should not be nil")
	}

	if para.Properties.KeepLines.Val != "1" {
		t.Errorf("Expected KeepLines Val to be '1', got '%s'", para.Properties.KeepLines.Val)
	}

	// Test disabling
	para.SetKeepLines(false)

	if para.Properties.KeepLines != nil {
		t.Error("KeepLines should be nil when disabled")
	}
}

// TestParagraphSetPageBreakBefore tests page break before paragraph
func TestParagraphSetPageBreakBefore(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试段前分页")

	// Test enabling
	para.SetPageBreakBefore(true)

	if para.Properties == nil {
		t.Fatal("Properties should not be nil")
	}

	if para.Properties.PageBreakBefore == nil {
		t.Fatal("PageBreakBefore should not be nil")
	}

	if para.Properties.PageBreakBefore.Val != "1" {
		t.Errorf("Expected PageBreakBefore Val to be '1', got '%s'", para.Properties.PageBreakBefore.Val)
	}

	// Test disabling
	para.SetPageBreakBefore(false)

	if para.Properties.PageBreakBefore != nil {
		t.Error("PageBreakBefore should be nil when disabled")
	}
}

// TestParagraphSetWidowControl tests widow/orphan control
func TestParagraphSetWidowControl(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试孤行控制")

	// Test enabling
	para.SetWidowControl(true)

	if para.Properties == nil {
		t.Fatal("Properties should not be nil")
	}

	if para.Properties.WidowControl == nil {
		t.Fatal("WidowControl should not be nil")
	}

	if para.Properties.WidowControl.Val != "1" {
		t.Errorf("Expected WidowControl Val to be '1', got '%s'", para.Properties.WidowControl.Val)
	}

	// Test disabling
	para.SetWidowControl(false)

	if para.Properties.WidowControl == nil {
		t.Fatal("WidowControl should not be nil when set to false")
	}

	if para.Properties.WidowControl.Val != "0" {
		t.Errorf("Expected WidowControl Val to be '0' when disabled, got '%s'", para.Properties.WidowControl.Val)
	}
}

// TestParagraphSetOutlineLevel tests outline level setting
func TestParagraphSetOutlineLevel(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试大纲级别")

	// Test valid level
	para.SetOutlineLevel(0)

	if para.Properties == nil {
		t.Fatal("Properties should not be nil")
	}

	if para.Properties.OutlineLevel == nil {
		t.Fatal("OutlineLevel should not be nil")
	}

	if para.Properties.OutlineLevel.Val != "0" {
		t.Errorf("Expected OutlineLevel Val to be '0', got '%s'", para.Properties.OutlineLevel.Val)
	}

	// Test other levels
	para.SetOutlineLevel(3)
	if para.Properties.OutlineLevel.Val != "3" {
		t.Errorf("Expected OutlineLevel Val to be '3', got '%s'", para.Properties.OutlineLevel.Val)
	}

	// Test boundary values
	para.SetOutlineLevel(8)
	if para.Properties.OutlineLevel.Val != "8" {
		t.Errorf("Expected OutlineLevel Val to be '8', got '%s'", para.Properties.OutlineLevel.Val)
	}

	// Test out-of-range values (should be clamped)
	para.SetOutlineLevel(10)
	if para.Properties.OutlineLevel.Val != "8" {
		t.Errorf("Expected OutlineLevel to be capped at '8', got '%s'", para.Properties.OutlineLevel.Val)
	}

	para.SetOutlineLevel(-1)
	if para.Properties.OutlineLevel.Val != "0" {
		t.Errorf("Expected OutlineLevel to be floored at '0', got '%s'", para.Properties.OutlineLevel.Val)
	}
}

// TestParagraphSetParagraphFormat tests comprehensive paragraph format setting
func TestParagraphSetParagraphFormat(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试综合格式设置")

	// Test full configuration
	config := &ParagraphFormatConfig{
		Alignment:       AlignCenter,
		Style:           styleHeading1,
		LineSpacing:     1.5,
		BeforePara:      24,
		AfterPara:       12,
		FirstLineCm:     0.5,
		LeftCm:          1.0,
		RightCm:         0.5,
		KeepWithNext:    true,
		KeepLines:       true,
		PageBreakBefore: true,
		WidowControl:    true,
		OutlineLevel:    0,
	}

	para.SetParagraphFormat(config)

	// Verify all properties
	if para.Properties == nil {
		t.Fatal("Properties should not be nil")
	}

	// Verify alignment
	if para.Properties.Justification == nil || para.Properties.Justification.Val != string(AlignCenter) {
		t.Error("Alignment not set correctly")
	}

	// Verify style
	if para.Properties.ParagraphStyle == nil || para.Properties.ParagraphStyle.Val != styleHeading1 {
		t.Error("Style not set correctly")
	}

	// Verify spacing
	if para.Properties.Spacing == nil {
		t.Fatal("Spacing should not be nil")
	}

	// Verify indentation
	if para.Properties.Indentation == nil {
		t.Fatal("Indentation should not be nil")
	}

	// Verify pagination control
	if para.Properties.KeepNext == nil || para.Properties.KeepNext.Val != "1" {
		t.Error("KeepNext not set correctly")
	}

	if para.Properties.KeepLines == nil || para.Properties.KeepLines.Val != "1" {
		t.Error("KeepLines not set correctly")
	}

	if para.Properties.PageBreakBefore == nil || para.Properties.PageBreakBefore.Val != "1" {
		t.Error("PageBreakBefore not set correctly")
	}

	if para.Properties.WidowControl == nil || para.Properties.WidowControl.Val != "1" {
		t.Error("WidowControl not set correctly")
	}

	// Verify outline level
	if para.Properties.OutlineLevel == nil || para.Properties.OutlineLevel.Val != "0" {
		t.Error("OutlineLevel not set correctly")
	}
}

// TestParagraphSetParagraphFormatNil tests nil configuration
func TestParagraphSetParagraphFormatNil(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试nil配置")

	// nil config should not cause a panic
	para.SetParagraphFormat(nil)

	// Paragraph should remain in default state
	if para.Properties != nil && para.Properties.Justification != nil {
		t.Error("Properties should remain unchanged with nil config")
	}
}

// TestParagraphSetParagraphFormatPartial tests partial configuration
func TestParagraphSetParagraphFormatPartial(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试部分配置")

	// Set only some properties
	config := &ParagraphFormatConfig{
		Alignment:    AlignRight,
		KeepWithNext: true,
		LineSpacing:  2.0,
	}

	para.SetParagraphFormat(config)

	// Verify the set properties
	if para.Properties == nil {
		t.Fatal("Properties should not be nil")
	}

	if para.Properties.Justification == nil || para.Properties.Justification.Val != string(AlignRight) {
		t.Error("Alignment not set correctly")
	}

	if para.Properties.KeepNext == nil || para.Properties.KeepNext.Val != "1" {
		t.Error("KeepNext not set correctly")
	}

	if para.Properties.Spacing == nil {
		t.Error("Spacing should be set")
	}

	// Verify unset properties remain at defaults
	if para.Properties.PageBreakBefore != nil {
		t.Error("PageBreakBefore should remain nil")
	}
}

// TestDocumentSave tests document saving
func TestDocumentSave(t *testing.T) {
	doc := New()
	doc.AddParagraph("测试保存功能")

	filename := "test_save.docx"
	defer os.Remove(filename) // Clean up test file

	err := doc.Save(filename)
	if err != nil {
		t.Fatalf("Failed to save document: %v", err)
	}

	// Verify file exists
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		t.Error("Saved file does not exist")
	}

	// Verify file size
	stat, err := os.Stat(filename)
	if err != nil {
		t.Fatalf("Failed to get file stats: %v", err)
	}

	if stat.Size() == 0 {
		t.Error("Saved file is empty")
	}
}

// TestDocumentGetStyleManager tests getting the style manager
func TestDocumentGetStyleManager(t *testing.T) {
	doc := New()

	styleManager := doc.GetStyleManager()
	if styleManager == nil {
		t.Fatal("Style manager should not be nil")
	}

	// Verify style manager functionality
	if !styleManager.StyleExists("Normal") {
		t.Error("Normal style should exist")
	}

	if !styleManager.StyleExists(styleHeading1) {
		t.Error("Heading1 style should exist")
	}
}

// TestComplexDocument tests complex document creation
func TestComplexDocument(t *testing.T) {
	doc := New()

	// Add title
	title := doc.AddFormattedParagraph("文档标题", &TextFormat{
		Bold:     true,
		FontSize: 18,
	})
	title.SetAlignment(AlignCenter)

	// Add various heading levels
	doc.AddHeadingParagraph("第一章", 1)
	doc.AddHeadingParagraph("1.1 概述", 2)
	doc.AddHeadingParagraph("1.1.1 背景", 3)

	// Add paragraph with spacing
	para := doc.AddParagraph("这是一个带有特殊间距的段落")
	para.SetSpacing(&SpacingConfig{
		LineSpacing: 1.5,
		BeforePara:  12,
		AfterPara:   6,
	})

	// Add mixed-format paragraph
	mixed := doc.AddParagraph("这段文字包含")
	mixed.AddFormattedText("粗体", &TextFormat{Bold: true})
	mixed.AddFormattedText("和", nil)
	mixed.AddFormattedText("斜体", &TextFormat{Italic: true})
	mixed.AddFormattedText("文本。", nil)

	// Verify document structure
	if len(doc.Body.GetParagraphs()) != 6 {
		t.Errorf("Expected 6 paragraphs, got %d", len(doc.Body.GetParagraphs()))
	}

	// Save and verify
	filename := "test_complex.docx"
	defer os.Remove(filename)

	err := doc.Save(filename)
	if err != nil {
		t.Fatalf("Failed to save complex document: %v", err)
	}
}

// TestDocumentOpen tests opening a document (requires creating a test document first)
func TestDocumentOpen(t *testing.T) {
	// First create a test document
	originalDoc := New()
	originalDoc.AddParagraph("第一段")
	originalDoc.AddParagraph("第二段")
	originalDoc.AddHeadingParagraph("标题", 1)

	filename := "test_open.docx"
	defer os.Remove(filename)

	err := originalDoc.Save(filename)
	if err != nil {
		t.Fatalf("Failed to save test document: %v", err)
	}

	// Open document
	loadedDoc, err := Open(filename)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}

	// Verify document content
	if len(loadedDoc.Body.GetParagraphs()) != 3 {
		t.Errorf("Expected 3 paragraphs, got %d", len(loadedDoc.Body.GetParagraphs()))
	}

	// Verify first paragraph content
	if len(loadedDoc.Body.GetParagraphs()[0].Runs) > 0 {
		content := loadedDoc.Body.GetParagraphs()[0].Runs[0].Text.Content
		if content != "第一段" {
			t.Errorf("Expected '第一段', got '%s'", content)
		}
	}
}

// TestErrorHandling tests error handling
func TestErrorHandling(t *testing.T) {
	// Test opening a non-existent file
	_, err := Open("nonexistent.docx")
	if err == nil {
		t.Error("Should return error when opening non-existent file")
	}

	// Test saving to a read-only directory (skip if creation fails)
	doc := New()
	doc.AddParagraph("测试")

	// Try saving to an invalid filename containing a null character
	invalidPath := "test\x00invalid.docx"
	err = doc.Save(invalidPath)
	if err == nil {
		// If the first test didn't fail, try another strategy
		// Try saving to an excessively long path
		longPath := string(make([]byte, 300)) + ".docx"
		err = doc.Save(longPath)
		if err == nil {
			t.Log("Warning: Unable to trigger save error - filesystem may be permissive")
		}
	}
}

// TestStyleIntegration tests style integration
func TestStyleIntegration(t *testing.T) {
	doc := New()
	styleManager := doc.GetStyleManager()
	quickAPI := style.NewQuickStyleAPI(styleManager)

	// Create custom style
	config := style.QuickStyleConfig{
		ID:      "TestStyle",
		Name:    "测试样式",
		Type:    style.StyleTypeParagraph,
		BasedOn: "Normal",
		RunConfig: &style.QuickRunConfig{
			Bold:      true,
			FontColor: "FF0000",
		},
	}

	_, err := quickAPI.CreateQuickStyle(config)
	if err != nil {
		t.Fatalf("Failed to create custom style: %v", err)
	}

	// Use custom style
	para := doc.AddParagraph("使用自定义样式")
	para.SetStyle("TestStyle")

	// Verify style is applied
	if para.Properties == nil || para.Properties.ParagraphStyle == nil {
		t.Fatal("Style should be applied to paragraph")
	}

	if para.Properties.ParagraphStyle.Val != "TestStyle" {
		t.Errorf("Expected TestStyle, got %s", para.Properties.ParagraphStyle.Val)
	}

	// Verify style exists
	if !styleManager.StyleExists("TestStyle") {
		t.Error("Custom style should exist in style manager")
	}
}

// BenchmarkAddParagraph benchmarks paragraph addition performance
func BenchmarkAddParagraph(b *testing.B) {
	doc := New()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		doc.AddParagraph("基准测试段落")
	}
}

// BenchmarkDocumentSave benchmarks document save performance
func BenchmarkDocumentSave(b *testing.B) {
	doc := New()

	// Create a medium-sized document
	for i := 0; i < 100; i++ {
		doc.AddParagraph("基准测试段落内容")
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		filename := "benchmark_save.docx"
		err := doc.Save(filename)
		if err != nil {
			b.Fatalf("Failed to save: %v", err)
		}
		os.Remove(filename)
	}
}

// TestTextFormatValidation tests text format validation
func TestTextFormatValidation(t *testing.T) {
	doc := New()

	// Test color format
	testCases := []struct {
		color    string
		expected string
	}{
		{"#FF0000", "FF0000"}, // with # prefix
		{"FF0000", "FF0000"},  // without # prefix
		{"#123456", "123456"},
		{"ABCDEF", "ABCDEF"},
	}

	for _, tc := range testCases {
		format := &TextFormat{
			FontColor: tc.color,
		}

		para := doc.AddFormattedParagraph("测试颜色", format)
		if para.Runs[0].Properties.Color.Val != tc.expected {
			t.Errorf("Color %s should be formatted as %s, got %s",
				tc.color, tc.expected, para.Runs[0].Properties.Color.Val)
		}
	}
}

// TestMemoryUsage tests memory usage
func TestMemoryUsage(t *testing.T) {
	doc := New()

	// Add a large number of paragraphs to test memory usage
	const numParagraphs = 1000
	for i := 0; i < numParagraphs; i++ {
		doc.AddParagraph("内存测试段落")
	}

	if len(doc.Body.GetParagraphs()) != numParagraphs {
		t.Errorf("Expected %d paragraphs, got %d", numParagraphs, len(doc.Body.GetParagraphs()))
	}

	// Test saving a large document
	filename := "test_memory.docx"
	defer os.Remove(filename)

	err := doc.Save(filename)
	if err != nil {
		t.Fatalf("Failed to save large document: %v", err)
	}
}

func TestDocumentOpenFromMemory(t *testing.T) {
	// First create a test document
	originalDoc := New()
	originalDoc.AddParagraph("第一段")
	originalDoc.AddParagraph("第二段")
	originalDoc.AddHeadingParagraph("标题", 1)

	filename := "test_open.docx"
	defer os.Remove(filename)

	err := originalDoc.Save(filename)
	if err != nil {
		t.Fatalf("Failed to save test document: %v", err)
	}

	// Open document
	files, err := os.Open(filename)
	if err != nil {
		t.Fatalf("Failed to open test document: %v", err)
	}
	defer files.Close()

	loadedDoc, err := OpenFromMemory(files)
	if err != nil {
		t.Fatalf("Failed to open document: %v", err)
	}

	for _, paragraphs := range loadedDoc.Body.GetParagraphs() {
		for _, run := range paragraphs.Runs {
			t.Log(run.Text.Content)
		}
	}
}

// TestAddPageBreak tests the add page break functionality
func TestAddPageBreak(t *testing.T) {
	doc := New()

	// Add first page content
	doc.AddParagraph("第一页内容")

	// Add page break
	doc.AddPageBreak()

	// Add second page content
	doc.AddParagraph("第二页内容")

	// Verify document contains 3 elements (paragraph, page break paragraph, paragraph)
	if len(doc.Body.Elements) != 3 {
		t.Errorf("expected document to contain 3 elements, got %d", len(doc.Body.Elements))
	}

	// Verify the second element is a paragraph containing a page break
	if p, ok := doc.Body.Elements[1].(*Paragraph); ok {
		if len(p.Runs) == 0 || p.Runs[0].Break == nil {
			t.Error("second element should be a paragraph containing a page break")
		} else if p.Runs[0].Break.Type != "page" {
			t.Errorf("page break type should be 'page', got '%s'", p.Runs[0].Break.Type)
		}
	} else {
		t.Error("second element should be a paragraph type")
	}

	// Save and verify document can be generated correctly
	filename := "test_page_break.docx"
	defer os.Remove(filename)

	err := doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document with page break: %v", err)
	}

	// Verify file exists
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		t.Error("saved document file does not exist")
	}
}

// TestParagraphAddPageBreak tests adding a page break within a paragraph
func TestParagraphAddPageBreak(t *testing.T) {
	doc := New()

	// Create a paragraph and add a page break
	para := doc.AddParagraph("分页符前的内容")
	para.AddPageBreak()
	para.AddFormattedText("分页符后的内容", nil)

	// Verify paragraph contains 3 runs (text, page break, text)
	if len(para.Runs) != 3 {
		t.Errorf("expected paragraph to contain 3 runs, got %d", len(para.Runs))
	}

	// Verify the second run is a page break
	if para.Runs[1].Break == nil {
		t.Error("second run should be a page break")
	} else if para.Runs[1].Break.Type != "page" {
		t.Errorf("page break type should be 'page', got '%s'", para.Runs[1].Break.Type)
	}

	// Save and verify document can be generated correctly
	filename := "test_paragraph_page_break.docx"
	defer os.Remove(filename)

	err := doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document with paragraph page break: %v", err)
	}

	// Verify file exists
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		t.Error("saved document file does not exist")
	}
}

// TestRemoveParagraph tests the remove paragraph functionality
func TestRemoveParagraph(t *testing.T) {
	doc := New()

	// Add three paragraphs
	para1 := doc.AddParagraph("第一段")
	para2 := doc.AddParagraph("第二段")
	para3 := doc.AddParagraph("第三段")

	// Verify initial state
	if len(doc.Body.Elements) != 3 {
		t.Fatalf("expected document to contain 3 paragraphs, got %d", len(doc.Body.Elements))
	}

	// Remove the second paragraph
	if !doc.RemoveParagraph(para2) {
		t.Error("removing paragraph should succeed")
	}

	// Verify state after removal
	if len(doc.Body.Elements) != 2 {
		t.Errorf("after removal, expected document to contain 2 paragraphs, got %d", len(doc.Body.Elements))
	}

	// Verify the remaining paragraphs are correct
	paragraphs := doc.Body.GetParagraphs()
	if len(paragraphs) != 2 {
		t.Fatalf("expected to get 2 paragraphs, got %d", len(paragraphs))
	}

	if paragraphs[0] != para1 {
		t.Error("first paragraph should be para1")
	}
	if paragraphs[1] != para3 {
		t.Error("second paragraph should be para3")
	}

	// Try to remove an already-removed paragraph (should return false)
	if doc.RemoveParagraph(para2) {
		t.Error("removing a non-existent paragraph should return false")
	}
}

// TestRemoveParagraphAt tests removing a paragraph by index
func TestRemoveParagraphAt(t *testing.T) {
	doc := New()

	// Add three paragraphs
	doc.AddParagraph("第一段")
	doc.AddParagraph("第二段")
	doc.AddParagraph("第三段")

	// Verify initial state
	paragraphs := doc.Body.GetParagraphs()
	if len(paragraphs) != 3 {
		t.Fatalf("expected document to contain 3 paragraphs, got %d", len(paragraphs))
	}

	// Remove the paragraph at index 1 (second paragraph)
	if !doc.RemoveParagraphAt(1) {
		t.Error("removing paragraph at index 1 should succeed")
	}

	// Verify state after removal
	paragraphs = doc.Body.GetParagraphs()
	if len(paragraphs) != 2 {
		t.Errorf("after removal, expected document to contain 2 paragraphs, got %d", len(paragraphs))
	}

	// Verify remaining paragraph content
	assertParagraphContent(t, paragraphs, 0, "第一段")
	assertParagraphContent(t, paragraphs, 1, "第三段")

	// Try to remove an out-of-range index
	if doc.RemoveParagraphAt(10) {
		t.Error("removing an out-of-range index should return false")
	}

	if doc.RemoveParagraphAt(-1) {
		t.Error("removing a negative index should return false")
	}
}

// TestRemoveElementAt tests removing an element by index
func TestRemoveElementAt(t *testing.T) {
	doc := New()

	// Add paragraphs and a table
	doc.AddParagraph("段落1")
	_, err := doc.AddTable(&TableConfig{Rows: 2, Cols: 2})
	if err != nil {
		t.Fatalf("failed to add table: %v", err)
	}
	doc.AddParagraph("段落2")

	// Verify initial state
	if len(doc.Body.Elements) != 3 {
		t.Fatalf("expected document to contain 3 elements, got %d", len(doc.Body.Elements))
	}

	// Remove the element at index 1 (table)
	if !doc.RemoveElementAt(1) {
		t.Error("removing element at index 1 should succeed")
	}

	// Verify state after removal
	if len(doc.Body.Elements) != 2 {
		t.Errorf("after removal, expected document to contain 2 elements, got %d", len(doc.Body.Elements))
	}

	// Verify the remaining elements are all paragraphs
	paragraphs := doc.Body.GetParagraphs()
	if len(paragraphs) != 2 {
		t.Errorf("expected to get 2 paragraphs, got %d", len(paragraphs))
	}

	// Try to remove an out-of-range index
	if doc.RemoveElementAt(10) {
		t.Error("removing an out-of-range index should return false")
	}
}

// TestPageBreakAndDeletion is an integration test for page breaks and deletion
func TestPageBreakAndDeletion(t *testing.T) {
	doc := New()

	// Create a document with page breaks
	doc.AddParagraph("第一页 - 段落1")
	doc.AddParagraph("第一页 - 段落2")
	doc.AddPageBreak()
	doc.AddParagraph("第二页 - 段落1")
	doc.AddPageBreak()
	doc.AddParagraph("第三页 - 段落1")

	// Verify initial state (2 paragraphs + 1 page break + 1 paragraph + 1 page break + 1 paragraph = 6 elements)
	if len(doc.Body.Elements) != 6 {
		t.Fatalf("expected document to contain 6 elements, got %d", len(doc.Body.Elements))
	}

	// Remove the first page break (index 2)
	if !doc.RemoveElementAt(2) {
		t.Error("removing page break should succeed")
	}

	// Verify state after removal
	if len(doc.Body.Elements) != 5 {
		t.Errorf("after removal, expected document to contain 5 elements, got %d", len(doc.Body.Elements))
	}

	// Save and verify document
	filename := "test_pagebreak_deletion.docx"
	defer os.Remove(filename)

	err := doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}

	// Verify file exists
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		t.Error("saved document file does not exist")
	}
}

// TestAddFormattedHeader tests adding a formatted header
func TestAddFormattedHeader(t *testing.T) {
	doc := New()

	// Test adding a formatted header
	config := &HeaderFooterConfig{
		Text: "公司报告",
		Format: &TextFormat{
			FontSize:   10,
			FontColor:  "8e8e8e",
			FontFamily: "Arial",
		},
		Alignment: AlignCenter,
	}

	err := doc.AddFormattedHeader(HeaderFooterTypeDefault, config)
	if err != nil {
		t.Fatalf("failed to add formatted header: %v", err)
	}

	// Verify header file was created
	headerPartName := testHeaderPartName
	if _, ok := doc.parts[headerPartName]; !ok {
		t.Errorf("header file %s was not created", headerPartName)
	}

	// Save and verify document
	filename := "test_formatted_header.docx"
	defer os.Remove(filename)

	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}

	// Verify file exists
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		t.Error("saved document file does not exist")
	}
}

// TestAddFormattedFooter tests adding a formatted footer
func TestAddFormattedFooter(t *testing.T) {
	doc := New()

	// Test adding a formatted footer
	config := &HeaderFooterConfig{
		Text: "第 1 页",
		Format: &TextFormat{
			FontSize:   9,
			FontColor:  "666666",
			FontFamily: "宋体",
			Bold:       true,
		},
		Alignment: AlignCenter,
	}

	err := doc.AddFormattedFooter(HeaderFooterTypeDefault, config)
	if err != nil {
		t.Fatalf("failed to add formatted footer: %v", err)
	}

	// Verify footer file was created
	footerPartName := "word/footer1.xml"
	if _, ok := doc.parts[footerPartName]; !ok {
		t.Errorf("footer file %s was not created", footerPartName)
	}

	// Save and verify document
	filename := "test_formatted_footer.docx"
	defer os.Remove(filename)

	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}

	// Verify file exists
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		t.Error("saved document file does not exist")
	}
}

// TestAddFormattedHeaderWithNilConfig tests adding a header with nil config
func TestAddFormattedHeaderWithNilConfig(t *testing.T) {
	doc := New()

	// Test adding a header with nil config
	err := doc.AddFormattedHeader(HeaderFooterTypeDefault, nil)
	if err != nil {
		t.Fatalf("failed to add header with nil config: %v", err)
	}

	// Verify header file was created
	headerPartName := testHeaderPartName
	if _, ok := doc.parts[headerPartName]; !ok {
		t.Errorf("header file %s was not created", headerPartName)
	}
}

// TestAddFormattedHeaderWithAllFormats tests all formatting options
func TestAddFormattedHeaderWithAllFormats(t *testing.T) {
	doc := New()

	// Test all formatting options
	config := &HeaderFooterConfig{
		Text: "格式化测试",
		Format: &TextFormat{
			Bold:       true,
			Italic:     true,
			FontSize:   12,
			FontColor:  "FF0000",
			FontFamily: "Times New Roman",
			Underline:  true,
			Strike:     true,
			Highlight:  "yellow",
		},
		Alignment: AlignRight,
	}

	err := doc.AddFormattedHeader(HeaderFooterTypeDefault, config)
	if err != nil {
		t.Fatalf("failed to add formatted header: %v", err)
	}

	// Save and verify document
	filename := "test_all_formats_header.docx"
	defer os.Remove(filename)

	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
}

// TestCreateFormattedParagraph tests the createFormattedParagraph function
func TestCreateFormattedParagraph(t *testing.T) {
	// Test basic paragraph creation
	para := createFormattedParagraph("测试文本", nil, "")
	if len(para.Runs) != 1 {
		t.Errorf("expected 1 Run, got %d", len(para.Runs))
	}
	if para.Runs[0].Text.Content != "测试文本" {
		t.Errorf("expected text '测试文本', got '%s'", para.Runs[0].Text.Content)
	}

	// Test paragraph with alignment
	para2 := createFormattedParagraph("居中文本", nil, AlignCenter)
	if para2.Properties == nil {
		t.Fatal("paragraph properties should not be nil")
	}
	if para2.Properties.Justification == nil {
		t.Fatal("alignment property should not be nil")
	}
	if para2.Properties.Justification.Val != string(AlignCenter) {
		t.Errorf("expected alignment 'center', got '%s'", para2.Properties.Justification.Val)
	}

	// Test paragraph with formatting
	format := &TextFormat{
		Bold:       true,
		FontSize:   14,
		FontColor:  testColorBlue,
		FontFamily: "Arial",
	}
	para3 := createFormattedParagraph("格式化文本", format, AlignLeft)
	if para3.Runs[0].Properties == nil {
		t.Fatal("Run properties should not be nil")
	}
	if para3.Runs[0].Properties.Bold == nil {
		t.Error("bold property should not be nil")
	}
	if para3.Runs[0].Properties.FontSize == nil {
		t.Error("font size property should not be nil")
	}
	if para3.Runs[0].Properties.FontSize.Val != "28" { // 14 * 2 = 28
		t.Errorf("expected font size '28', got '%s'", para3.Runs[0].Properties.FontSize.Val)
	}
	if para3.Runs[0].Properties.Color == nil {
		t.Error("color property should not be nil")
	}
	if para3.Runs[0].Properties.Color.Val != testColorBlue {
		t.Errorf("expected color '0000FF', got '%s'", para3.Runs[0].Properties.Color.Val)
	}
	if para3.Runs[0].Properties.FontFamily == nil {
		t.Error("font family property should not be nil")
	}
	if para3.Runs[0].Properties.FontFamily.ASCII != "Arial" {
		t.Errorf("expected font 'Arial', got '%s'", para3.Runs[0].Properties.FontFamily.ASCII)
	}

	// Test empty text
	para4 := createFormattedParagraph("", nil, AlignCenter)
	if len(para4.Runs) != 0 {
		t.Errorf("empty text should not add a Run, got %d", len(para4.Runs))
	}
}

// TestParagraphSetUnderline tests paragraph underline setting
func TestParagraphSetUnderline(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试下划线文本")

	// Test enabling underline
	para.SetUnderline(true)
	if len(para.Runs) == 0 {
		t.Fatal("paragraph should contain at least one Run")
	}
	if para.Runs[0].Properties == nil {
		t.Fatal("Run properties should not be nil")
	}
	if para.Runs[0].Properties.Underline == nil {
		t.Error("underline property should not be nil")
	}
	if para.Runs[0].Properties.Underline.Val != "single" {
		t.Errorf("expected underline type 'single', got '%s'", para.Runs[0].Properties.Underline.Val)
	}

	// Test disabling underline
	para.SetUnderline(false)
	if para.Runs[0].Properties.Underline != nil {
		t.Error("underline property should be nil after disabling")
	}
}

// TestParagraphSetBold tests paragraph bold setting
func TestParagraphSetBold(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试粗体文本")

	// Test enabling bold
	para.SetBold(true)
	if para.Runs[0].Properties == nil {
		t.Fatal("Run properties should not be nil")
	}
	if para.Runs[0].Properties.Bold == nil {
		t.Error("bold property should not be nil")
	}
	if para.Runs[0].Properties.BoldCs == nil {
		t.Error("complex script bold property should not be nil")
	}

	// Test disabling bold
	para.SetBold(false)
	if para.Runs[0].Properties.Bold != nil {
		t.Error("bold property should be nil after disabling")
	}
	if para.Runs[0].Properties.BoldCs != nil {
		t.Error("complex script bold property should be nil after disabling")
	}
}

// TestParagraphSetItalic tests paragraph italic setting
func TestParagraphSetItalic(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试斜体文本")

	// Test enabling italic
	para.SetItalic(true)
	if para.Runs[0].Properties == nil {
		t.Fatal("Run properties should not be nil")
	}
	if para.Runs[0].Properties.Italic == nil {
		t.Error("italic property should not be nil")
	}
	if para.Runs[0].Properties.ItalicCs == nil {
		t.Error("complex script italic property should not be nil")
	}

	// Test disabling italic
	para.SetItalic(false)
	if para.Runs[0].Properties.Italic != nil {
		t.Error("italic property should be nil after disabling")
	}
	if para.Runs[0].Properties.ItalicCs != nil {
		t.Error("complex script italic property should be nil after disabling")
	}
}

// TestParagraphSetStrike tests paragraph strikethrough setting
func TestParagraphSetStrike(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试删除线文本")

	// Test enabling strikethrough
	para.SetStrike(true)
	if para.Runs[0].Properties == nil {
		t.Fatal("Run properties should not be nil")
	}
	if para.Runs[0].Properties.Strike == nil {
		t.Error("strikethrough property should not be nil")
	}

	// Test disabling strikethrough
	para.SetStrike(false)
	if para.Runs[0].Properties.Strike != nil {
		t.Error("strikethrough property should be nil after disabling")
	}
}

// TestParagraphSetHighlight tests paragraph highlight setting
func TestParagraphSetHighlight(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试高亮文本")

	// Test setting yellow highlight
	para.SetHighlight("yellow")
	if para.Runs[0].Properties == nil {
		t.Fatal("Run properties should not be nil")
	}
	if para.Runs[0].Properties.Highlight == nil {
		t.Error("highlight property should not be nil")
	}
	if para.Runs[0].Properties.Highlight.Val != "yellow" {
		t.Errorf("expected highlight color 'yellow', got '%s'", para.Runs[0].Properties.Highlight.Val)
	}

	// Test removing highlight
	para.SetHighlight("")
	if para.Runs[0].Properties.Highlight != nil {
		t.Error("highlight property should be nil after removal")
	}
}

// TestParagraphSetFontFamily tests paragraph font family setting
func TestParagraphSetFontFamily(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试字体文本")

	// Test setting font family
	para.SetFontFamily("微软雅黑")
	if para.Runs[0].Properties == nil {
		t.Fatal("Run properties should not be nil")
	}
	if para.Runs[0].Properties.FontFamily == nil {
		t.Error("font family property should not be nil")
	}
	if para.Runs[0].Properties.FontFamily.ASCII != "微软雅黑" {
		t.Errorf("expected font '微软雅黑', got '%s'", para.Runs[0].Properties.FontFamily.ASCII)
	}
	if para.Runs[0].Properties.FontFamily.EastAsia != "微软雅黑" {
		t.Errorf("expected East Asian font '微软雅黑', got '%s'", para.Runs[0].Properties.FontFamily.EastAsia)
	}

	// Test removing font family
	para.SetFontFamily("")
	if para.Runs[0].Properties.FontFamily != nil {
		t.Error("font family property should be nil after removal")
	}
}

// TestParagraphSetFontSize tests paragraph font size setting
func TestParagraphSetFontSize(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试字体大小文本")

	// Test setting font size
	para.SetFontSize(14)
	if para.Runs[0].Properties == nil {
		t.Fatal("Run properties should not be nil")
	}
	if para.Runs[0].Properties.FontSize == nil {
		t.Error("font size property should not be nil")
	}
	// Word uses half-point units, 14pt = 28 half-points
	if para.Runs[0].Properties.FontSize.Val != "28" {
		t.Errorf("expected font size '28' (14pt), got '%s'", para.Runs[0].Properties.FontSize.Val)
	}
	if para.Runs[0].Properties.FontSizeCs == nil {
		t.Error("complex script font size property should not be nil")
	}

	// Test removing font size
	para.SetFontSize(0)
	if para.Runs[0].Properties.FontSize != nil {
		t.Error("font size property should be nil after removal")
	}
}

// TestParagraphSetColor tests paragraph color setting
func TestParagraphSetColor(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("测试颜色文本")

	// Test setting color (without # prefix)
	para.SetColor("FF0000")
	if para.Runs[0].Properties == nil {
		t.Fatal("Run properties should not be nil")
	}
	if para.Runs[0].Properties.Color == nil {
		t.Error("color property should not be nil")
	}
	if para.Runs[0].Properties.Color.Val != "FF0000" {
		t.Errorf("expected color 'FF0000', got '%s'", para.Runs[0].Properties.Color.Val)
	}

	// Test setting color (with # prefix, should be removed)
	para.SetColor("#0000FF")
	if para.Runs[0].Properties.Color.Val != testColorBlue {
		t.Errorf("expected color '0000FF' (# prefix should be removed), got '%s'", para.Runs[0].Properties.Color.Val)
	}

	// Test removing color
	para.SetColor("")
	if para.Runs[0].Properties.Color != nil {
		t.Error("color property should be nil after removal")
	}
}

// TestParagraphMultipleRunsFormatting tests formatting with multiple Runs
func TestParagraphMultipleRunsFormatting(t *testing.T) {
	doc := New()
	// Create paragraph with first text, not an empty string
	para := doc.AddParagraph("第一段文本")

	// Add more Runs
	para.AddFormattedText("第二段文本", nil)
	para.AddFormattedText("第三段文本", nil)

	if len(para.Runs) != 3 {
		t.Fatalf("expected 3 Runs, got %d", len(para.Runs))
	}

	// Test that underline is applied to all Runs
	para.SetUnderline(true)
	for i, run := range para.Runs {
		if run.Properties == nil || run.Properties.Underline == nil {
			t.Errorf("Run %d should have underline property", i)
		}
	}

	// Test that bold is applied to all Runs
	para.SetBold(true)
	for i, run := range para.Runs {
		if run.Properties == nil || run.Properties.Bold == nil {
			t.Errorf("Run %d should have bold property", i)
		}
	}

	// Test that font family is applied to all Runs
	para.SetFontFamily("Arial")
	for i, run := range para.Runs {
		if run.Properties == nil || run.Properties.FontFamily == nil {
			t.Errorf("Run %d should have font family property", i)
		}
		if run.Properties.FontFamily.ASCII != "Arial" {
			t.Errorf("Run %d font should be 'Arial', got '%s'", i, run.Properties.FontFamily.ASCII)
		}
	}
}

// TestParagraphFormattingIntegration tests text formatting integration
func TestParagraphFormattingIntegration(t *testing.T) {
	doc := New()
	para := doc.AddParagraph("完整格式化测试文本")

	// Apply all formatting
	para.SetBold(true)
	para.SetItalic(true)
	para.SetUnderline(true)
	para.SetStrike(true)
	para.SetHighlight("yellow")
	para.SetFontFamily("Times New Roman")
	para.SetFontSize(16)
	para.SetColor(testColorBlue)

	// Verify all formatting has been applied
	props := para.Runs[0].Properties
	if props.Bold == nil {
		t.Error("bold property not set")
	}
	if props.Italic == nil {
		t.Error("italic property not set")
	}
	if props.Underline == nil {
		t.Error("underline property not set")
	}
	if props.Strike == nil {
		t.Error("strikethrough property not set")
	}
	if props.Highlight == nil || props.Highlight.Val != "yellow" {
		t.Error("highlight property not set correctly")
	}
	if props.FontFamily == nil || props.FontFamily.ASCII != "Times New Roman" {
		t.Error("font family property not set correctly")
	}
	if props.FontSize == nil || props.FontSize.Val != "32" {
		t.Errorf("font size property not set correctly, expected '32', got '%s'", props.FontSize.Val)
	}
	if props.Color == nil || props.Color.Val != testColorBlue {
		t.Error("color property not set correctly")
	}

	// Save document to verify
	filename := "test_formatting_integration.docx"
	defer os.Remove(filename)

	err := doc.Save(filename)
	if err != nil {
		t.Errorf("failed to save document: %v", err)
	}
}
