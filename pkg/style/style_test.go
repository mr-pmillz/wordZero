package style

import (
	"testing"
)

// TestNewStyleManager tests style manager creation
func TestNewStyleManager(t *testing.T) {
	sm := NewStyleManager()

	if sm == nil {
		t.Fatal("StyleManager should not be nil")
	}

	// Verify predefined styles are loaded
	styles := sm.GetAllStyles()
	if len(styles) == 0 {
		t.Error("Should have predefined styles loaded")
	}

	// Verify basic styles exist
	expectedStyles := []string{"Normal", "Heading1", "Heading2", "Title", "Subtitle"}
	for _, styleID := range expectedStyles {
		if !sm.StyleExists(styleID) {
			t.Errorf("Style %s should exist", styleID)
		}
	}
}

// TestStyleExists tests style existence check
func TestStyleExists(t *testing.T) {
	sm := NewStyleManager()

	// Test existing style
	if !sm.StyleExists("Normal") {
		t.Error("Normal style should exist")
	}

	if !sm.StyleExists("Heading1") {
		t.Error("Heading1 style should exist")
	}

	// Test non-existing style
	if sm.StyleExists("NonExistentStyle") {
		t.Error("NonExistentStyle should not exist")
	}
}

// TestGetStyle tests getting a style
func TestGetStyle(t *testing.T) {
	sm := NewStyleManager()

	// Test getting existing style
	normalStyle := sm.GetStyle("Normal")
	if normalStyle == nil {
		t.Fatal("Normal style should not be nil")
	}

	if normalStyle.StyleID != "Normal" {
		t.Errorf("Expected StyleID Normal, got %s", normalStyle.StyleID)
	}

	// Test getting non-existing style
	nonExistent := sm.GetStyle("NonExistentStyle")
	if nonExistent != nil {
		t.Error("NonExistentStyle should return nil")
	}
}

// TestGetHeadingStyles tests getting heading styles
func TestGetHeadingStyles(t *testing.T) {
	sm := NewStyleManager()

	headingStyles := sm.GetHeadingStyles()

	// Should have 9 heading styles
	if len(headingStyles) != 9 {
		t.Errorf("Expected 9 heading styles, got %d", len(headingStyles))
	}

	// Verify heading style IDs
	expectedHeadings := []string{"Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6", "Heading7", "Heading8", "Heading9"}
	styleMap := make(map[string]bool)
	for _, style := range headingStyles {
		styleMap[style.StyleID] = true
	}

	for _, expected := range expectedHeadings {
		if !styleMap[expected] {
			t.Errorf("Heading style %s should be included", expected)
		}
	}
}

// TestAddStyle tests adding a custom style
func TestAddStyle(t *testing.T) {
	sm := NewStyleManager()

	customStyle := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "CustomTest",
		Name:    &StyleName{Val: "测试样式"},
		RunPr: &RunProperties{
			Bold:  &Bold{},
			Color: &Color{Val: "FF0000"},
		},
	}

	sm.AddStyle(customStyle)

	// Verify style was added
	if !sm.StyleExists("CustomTest") {
		t.Error("Custom style should exist after adding")
	}

	// Verify style content
	retrieved := sm.GetStyle("CustomTest")
	if retrieved == nil {
		t.Fatal("Retrieved custom style should not be nil")
	}

	if retrieved.StyleID != "CustomTest" {
		t.Errorf("Expected StyleID CustomTest, got %s", retrieved.StyleID)
	}

	if retrieved.Name.Val != "测试样式" {
		t.Errorf("Expected name 测试样式, got %s", retrieved.Name.Val)
	}
}

// TestRemoveStyle tests removing a style
func TestRemoveStyle(t *testing.T) {
	sm := NewStyleManager()

	// First add a test style
	testStyle := &Style{
		Type:    string(StyleTypeParagraph),
		StyleID: "TestRemove",
		Name:    &StyleName{Val: "待删除样式"},
	}

	sm.AddStyle(testStyle)

	// Verify style exists
	if !sm.StyleExists("TestRemove") {
		t.Fatal("Test style should exist before removal")
	}

	// Remove style
	sm.RemoveStyle("TestRemove")

	// Verify style was removed
	if sm.StyleExists("TestRemove") {
		t.Error("Test style should not exist after removal")
	}

	// Try removing non-existing style (should not error)
	sm.RemoveStyle("NonExistentStyle")
}

// TestGetStyleWithInheritance tests style inheritance
func TestGetStyleWithInheritance(t *testing.T) {
	sm := NewStyleManager()

	// Get Heading1 style with inheritance
	heading1 := sm.GetStyleWithInheritance("Heading1")
	if heading1 == nil {
		t.Fatal("Heading1 with inheritance should not be nil")
	}

	// Heading1 is based on Normal, should inherit Normal properties
	if heading1.BasedOn == nil {
		t.Error("Heading1 should have BasedOn reference")
	}

	// Verify inherited properties
	if heading1.RunPr == nil {
		t.Error("Heading1 should have run properties")
	}

	// Test non-existing style
	nonExistent := sm.GetStyleWithInheritance("NonExistentStyle")
	if nonExistent != nil {
		t.Error("Non-existent style with inheritance should return nil")
	}
}

// TestQuickStyleAPI tests quick API functionality
func TestQuickStyleAPI(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	if api == nil {
		t.Fatal("QuickStyleAPI should not be nil")
	}

	if api.styleManager != sm {
		t.Error("QuickStyleAPI should reference the provided StyleManager")
	}

	// Test getting all style info
	stylesInfo := api.GetAllStylesInfo()
	if len(stylesInfo) == 0 {
		t.Error("Should have style information")
	}

	// Verify returned info structure
	for _, info := range stylesInfo {
		if info.ID == "" {
			t.Error("Style info should have ID")
		}
		if info.Name == "" {
			t.Error("Style info should have Name")
		}
		if info.Type == "" {
			t.Error("Style info should have Type")
		}
	}
}

// TestQuickStyleAPI_GetStyleInfo tests getting single style info
func TestQuickStyleAPI_GetStyleInfo(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	// Test getting existing style info
	info, err := api.GetStyleInfo("Normal")
	if err != nil {
		t.Fatalf("Error getting Normal style info: %v", err)
	}

	if info.ID != "Normal" {
		t.Errorf("Expected ID Normal, got %s", info.ID)
	}

	if info.Name != "Normal" {
		t.Errorf("Expected name Normal, got %s", info.Name)
	}

	// Test getting non-existing style info
	_, err = api.GetStyleInfo("NonExistentStyle")
	if err == nil {
		t.Error("Should return error for non-existent style")
	}
}

// TestQuickStyleAPI_CreateStyle tests quick style creation
func TestQuickStyleAPI_CreateStyle(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	config := QuickStyleConfig{
		ID:      "QuickTest",
		Name:    "快速测试样式",
		Type:    StyleTypeParagraph,
		BasedOn: "Normal",
		ParagraphConfig: &QuickParagraphConfig{
			Alignment:   "center",
			LineSpacing: 1.5,
			SpaceBefore: 12,
			SpaceAfter:  6,
		},
		RunConfig: &QuickRunConfig{
			FontName:  "宋体",
			FontSize:  14,
			FontColor: "FF0000",
			Bold:      true,
			Italic:    false,
		},
	}

	style, err := api.CreateQuickStyle(config)
	if err != nil {
		t.Fatalf("Failed to create quick style: %v", err)
	}

	// Verify style creation
	if style.StyleID != "QuickTest" {
		t.Errorf("Expected StyleID QuickTest, got %s", style.StyleID)
	}

	if style.Name.Val != "快速测试样式" {
		t.Errorf("Expected name 快速测试样式, got %s", style.Name.Val)
	}

	// Verify style was added to manager
	if !sm.StyleExists("QuickTest") {
		t.Error("Quick style should exist in style manager")
	}

	// Verify paragraph properties
	if style.ParagraphPr == nil {
		t.Fatal("Paragraph properties should not be nil")
	}

	if style.ParagraphPr.Justification == nil || style.ParagraphPr.Justification.Val != "center" {
		t.Error("Alignment should be center")
	}

	// Verify run properties
	if style.RunPr == nil {
		t.Fatal("Run properties should not be nil")
	}

	if style.RunPr.Bold == nil {
		t.Error("Should be bold")
	}

	if style.RunPr.Color == nil || style.RunPr.Color.Val != "FF0000" {
		t.Error("Color should be FF0000")
	}

	if style.RunPr.FontSize == nil || style.RunPr.FontSize.Val != "28" {
		t.Error("Font size should be 28 (14*2)")
	}
}

// TestQuickStyleAPI_StylesByType tests getting styles by type
func TestQuickStyleAPI_StylesByType(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	// Test getting paragraph styles
	paragraphStyles := api.GetParagraphStylesInfo()
	if len(paragraphStyles) == 0 {
		t.Error("Should have paragraph styles")
	}

	// Verify all returned are paragraph styles
	for _, info := range paragraphStyles {
		if info.Type != "paragraph" {
			t.Errorf("Expected paragraph type, got %s", info.Type)
		}
	}

	// Test getting character styles
	characterStyles := api.GetCharacterStylesInfo()
	for _, info := range characterStyles {
		if info.Type != "character" {
			t.Errorf("Expected character type, got %s", info.Type)
		}
	}

	// Test getting heading styles
	headingStyles := api.GetHeadingStylesInfo()
	if len(headingStyles) != 9 {
		t.Errorf("Expected 9 heading styles, got %d", len(headingStyles))
	}
}

// BenchmarkStyleLookup benchmarks style lookup performance
func BenchmarkStyleLookup(b *testing.B) {
	sm := NewStyleManager()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		sm.GetStyle("Heading1")
	}
}

// BenchmarkStyleWithInheritance benchmarks inherited style performance
func BenchmarkStyleWithInheritance(b *testing.B) {
	sm := NewStyleManager()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		sm.GetStyleWithInheritance("Heading1")
	}
}
