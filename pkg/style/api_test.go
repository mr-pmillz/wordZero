package style

import (
	"testing"
)

const testAlignCenter = "center"

func TestNewQuickStyleAPI(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	if api == nil {
		t.Fatal("NewQuickStyleAPI returned nil")
	}

	if api.styleManager != sm {
		t.Error("QuickStyleAPI styleManager not set correctly")
	}
}

func TestGetStyleInfo(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	// Test getting existing style info
	info, err := api.GetStyleInfo("Heading1")
	if err != nil {
		t.Fatalf("failed to get style info: %v", err)
	}

	if info.ID != "Heading1" {
		t.Errorf("expected style ID 'Heading1', got '%s'", info.ID)
	}

	if info.Name != "heading 1" {
		t.Errorf("expected style name 'heading 1', got '%s'", info.Name)
	}

	if info.Type != StyleTypeParagraph {
		t.Errorf("expected style type '%s', got '%s'", StyleTypeParagraph, info.Type)
	}

	if !info.IsBuiltIn {
		t.Error("Heading1 should be a built-in style")
	}

	// Test getting non-existent style info
	_, err = api.GetStyleInfo("NonExistentStyle")
	if err == nil {
		t.Error("expected error when getting non-existent style")
	}
}

func TestGetAllStylesInfo(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	allStyles := api.GetAllStylesInfo()

	if len(allStyles) == 0 {
		t.Error("expected style info list to be non-empty")
	}

	// Check if it contains expected styles
	styleFound := false
	for _, info := range allStyles {
		if info.ID == StyleNormal {
			styleFound = true
			break
		}
	}

	if !styleFound {
		t.Error("expected to find 'Normal' style in style list")
	}
}

func TestGetHeadingStylesInfo(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	headingStyles := api.GetHeadingStylesInfo()

	expectedCount := 9 // Heading1 through Heading9
	if len(headingStyles) != expectedCount {
		t.Errorf("expected %d heading styles, got %d", expectedCount, len(headingStyles))
	}

	// Check heading style order and IDs
	for i, info := range headingStyles {
		expectedID := "Heading" + string(rune('1'+i))
		if info.ID != expectedID {
			t.Errorf("expected heading style %d ID to be '%s', got '%s'", i+1, expectedID, info.ID)
		}

		if info.Type != StyleTypeParagraph {
			t.Errorf("heading style '%s' should be paragraph type", info.ID)
		}
	}
}

func TestGetParagraphStylesInfo(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	paragraphStyles := api.GetParagraphStylesInfo()

	if len(paragraphStyles) == 0 {
		t.Error("expected paragraph style list to be non-empty")
	}

	// Check all returned styles are paragraph type
	for _, info := range paragraphStyles {
		if info.Type != StyleTypeParagraph {
			t.Errorf("style '%s' should be paragraph type, got '%s'", info.ID, info.Type)
		}
	}
}

func TestGetCharacterStylesInfo(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	characterStyles := api.GetCharacterStylesInfo()

	if len(characterStyles) == 0 {
		t.Error("expected character style list to be non-empty")
	}

	// Check all returned styles are character type
	for _, info := range characterStyles {
		if info.Type != StyleTypeCharacter {
			t.Errorf("style '%s' should be character type, got '%s'", info.ID, info.Type)
		}
	}
}

func TestCreateQuickStyle(t *testing.T) {
	sm := NewStyleManager()
	api := NewQuickStyleAPI(sm)

	// Test creating custom paragraph style
	config := QuickStyleConfig{
		ID:      "TestCustomStyle",
		Name:    "测试自定义样式",
		Type:    StyleTypeParagraph,
		BasedOn: StyleNormal,
		ParagraphConfig: &QuickParagraphConfig{
			Alignment:   testAlignCenter,
			LineSpacing: 1.5,
			SpaceBefore: 12,
			SpaceAfter:  6,
		},
		RunConfig: &QuickRunConfig{
			FontName:  "微软雅黑",
			FontSize:  14,
			FontColor: "FF0000",
			Bold:      true,
		},
	}

	style, err := api.CreateQuickStyle(config)
	if err != nil {
		t.Fatalf("failed to create custom style: %v", err)
	}

	if style.StyleID != "TestCustomStyle" {
		t.Errorf("expected style ID 'TestCustomStyle', got '%s'", style.StyleID)
	}

	if !style.CustomStyle {
		t.Error("created style should be marked as custom")
	}

	// Verify paragraph properties
	if style.ParagraphPr == nil {
		t.Error("custom style should have paragraph properties")
	} else if style.ParagraphPr.Justification == nil || style.ParagraphPr.Justification.Val != testAlignCenter {
		t.Error("paragraph alignment not set correctly")
	}

	// Verify run properties
	switch {
	case style.RunPr == nil:
		t.Error("custom style should have run properties")
	case style.RunPr.Bold == nil:
		t.Error("bold property not set correctly")
	case style.RunPr.FontSize == nil || style.RunPr.FontSize.Val != "28":
		t.Error("font size not set correctly")
	}

	// Test creating style with duplicate ID
	_, err = api.CreateQuickStyle(config)
	if err == nil {
		t.Error("expected error when creating style with duplicate ID")
	}
}

func TestCreateParagraphProperties(t *testing.T) {
	config := &QuickParagraphConfig{
		Alignment:       testAlignCenter,
		LineSpacing:     1.5,
		SpaceBefore:     12,
		SpaceAfter:      6,
		FirstLineIndent: 24,
		LeftIndent:      36,
		RightIndent:     36,
	}

	props := createParagraphProperties(config)

	if props == nil {
		t.Fatal("createParagraphProperties returned nil")
	}

	// Check alignment
	if props.Justification == nil || props.Justification.Val != testAlignCenter {
		t.Error("alignment not set correctly")
	}

	// Check spacing
	if props.Spacing == nil {
		t.Error("spacing properties not set")
	} else {
		if props.Spacing.Before != "240" { // 12 * 20
			t.Errorf("space before not set correctly, expected '240', got '%s'", props.Spacing.Before)
		}
		if props.Spacing.After != "120" { // 6 * 20
			t.Errorf("space after not set correctly, expected '120', got '%s'", props.Spacing.After)
		}
	}

	// Check indentation
	if props.Indentation == nil {
		t.Error("indentation properties not set")
	} else if props.Indentation.FirstLine != "480" { // 24 * 20
		t.Errorf("first line indent not set correctly, expected '480', got '%s'", props.Indentation.FirstLine)
	}
}

func TestCreateRunProperties(t *testing.T) {
	config := &QuickRunConfig{
		FontName:  "微软雅黑",
		FontSize:  14,
		FontColor: "FF0000",
		Bold:      true,
		Italic:    true,
		Underline: true,
		Strike:    true,
		Highlight: "yellow",
	}

	props := createRunProperties(config)

	if props == nil {
		t.Fatal("createRunProperties returned nil")
	}

	// Check font settings
	if props.FontFamily == nil {
		t.Error("font family not set")
	} else if props.FontFamily.ASCII != "微软雅黑" {
		t.Errorf("ASCII font not set correctly, expected '微软雅黑', got '%s'", props.FontFamily.ASCII)
	}

	if props.FontSize == nil || props.FontSize.Val != "28" { // 14 * 2
		t.Error("font size not set correctly")
	}

	if props.Color == nil || props.Color.Val != "FF0000" {
		t.Error("font color not set correctly")
	}

	// Check formatting
	if props.Bold == nil {
		t.Error("bold not set correctly")
	}

	if props.Italic == nil {
		t.Error("italic not set correctly")
	}

	if props.Underline == nil || props.Underline.Val != "single" {
		t.Error("underline not set correctly")
	}

	if props.Strike == nil {
		t.Error("strikethrough not set correctly")
	}

	if props.Highlight == nil || props.Highlight.Val != "yellow" {
		t.Error("highlight not set correctly")
	}
}

func TestCreateParagraphPropertiesWithSnapToGrid(t *testing.T) {
	// Test SnapToGrid = false disables grid alignment
	snapToGridFalse := false
	config := &QuickParagraphConfig{
		Alignment:   "left",
		LineSpacing: 1.5,
		SnapToGrid:  &snapToGridFalse,
	}

	props := createParagraphProperties(config)

	if props == nil {
		t.Fatal("createParagraphProperties returned nil")
	}

	// Check SnapToGrid setting
	if props.SnapToGrid == nil {
		t.Error("SnapToGrid should be set")
	} else if props.SnapToGrid.Val != "0" {
		t.Errorf("SnapToGrid.Val not set correctly, expected '0', got '%s'", props.SnapToGrid.Val)
	}

	// Check line spacing
	if props.Spacing == nil {
		t.Error("spacing properties not set")
	} else {
		if props.Spacing.Line != "360" { // 1.5 * 240
			t.Errorf("line spacing not set correctly, expected '360', got '%s'", props.Spacing.Line)
		}
		if props.Spacing.LineRule != "auto" {
			t.Errorf("LineRule not set correctly, expected 'auto', got '%s'", props.Spacing.LineRule)
		}
	}

	// Test SnapToGrid = true does not set (keep default)
	snapToGridTrue := true
	configWithGridEnabled := &QuickParagraphConfig{
		Alignment:   "left",
		LineSpacing: 1.5,
		SnapToGrid:  &snapToGridTrue,
	}

	propsWithGrid := createParagraphProperties(configWithGridEnabled)

	if propsWithGrid.SnapToGrid != nil {
		t.Error("when SnapToGrid = true, SnapToGrid property should not be set (keep default behavior)")
	}

	// Test SnapToGrid = nil does not set
	configWithoutGrid := &QuickParagraphConfig{
		Alignment:   "left",
		LineSpacing: 1.5,
		SnapToGrid:  nil,
	}

	propsWithoutGrid := createParagraphProperties(configWithoutGrid)

	if propsWithoutGrid.SnapToGrid != nil {
		t.Error("when SnapToGrid = nil, SnapToGrid property should not be set")
	}
}
