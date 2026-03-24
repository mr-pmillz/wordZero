package document

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"testing"
)

// createCellTestImage creates PNG image data for testing
func createCellTestImage(width, height int) []byte {
	img := image.NewRGBA(image.Rect(0, 0, width, height))

	// Fill with red background
	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			img.Set(x, y, color.RGBA{255, 100, 100, 255})
		}
	}

	buf := new(bytes.Buffer)
	png.Encode(buf, img)
	return buf.Bytes()
}

// TestAddCellParagraph tests adding paragraphs to a cell
func TestAddCellParagraph(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 4000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test adding a paragraph
	para, err := table.AddCellParagraph(0, 0, "第一段内容")
	if err != nil {
		t.Errorf("failed to add paragraph: %v", err)
	}
	if para == nil {
		t.Error("returned paragraph should not be nil")
	}

	// Add a second paragraph
	para2, err := table.AddCellParagraph(0, 0, "第二段内容")
	if err != nil {
		t.Errorf("failed to add second paragraph: %v", err)
	}
	if para2 == nil {
		t.Error("returned second paragraph should not be nil")
	}

	// Verify paragraph count
	paragraphs, err := table.GetCellParagraphs(0, 0)
	if err != nil {
		t.Errorf("failed to get paragraphs: %v", err)
	}

	// Initial empty paragraph plus two new paragraphs
	if len(paragraphs) < 3 {
		t.Errorf("expected at least 3 paragraphs, got %d", len(paragraphs))
	}

	// Test invalid index
	_, err = table.AddCellParagraph(10, 10, "无效")
	if err == nil {
		t.Error("expected invalid index to fail, but it succeeded")
	}
}

// TestAddCellFormattedParagraph tests adding formatted paragraphs to a cell
func TestAddCellFormattedParagraph(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 4000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test adding a formatted paragraph
	format := &TextFormat{
		Bold:       true,
		Italic:     true,
		FontSize:   14,
		FontColor:  "FF0000",
		FontFamily: "Arial",
		Underline:  true,
	}

	para, err := table.AddCellFormattedParagraph(0, 0, "格式化内容", format)
	if err != nil {
		t.Errorf("failed to add formatted paragraph: %v", err)
	}
	if para == nil {
		t.Error("returned paragraph should not be nil")
	}

	// Verify format
	if len(para.Runs) == 0 {
		t.Error("paragraph should contain at least one Run")
	}

	run := para.Runs[0]
	if run.Properties == nil {
		t.Error("Run should have properties")
	} else {
		if run.Properties.Bold == nil {
			t.Error("expected bold property")
		}
		if run.Properties.Italic == nil {
			t.Error("expected italic property")
		}
		if run.Properties.Underline == nil {
			t.Error("expected underline property")
		}
	}
}

// TestClearCellParagraphs tests clearing cell paragraphs
func TestClearCellParagraphs(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 4000,
		Data: [][]string{
			{"A1", "B1"},
			{"A2", "B2"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Add multiple paragraphs
	table.AddCellParagraph(0, 0, "段落1")
	table.AddCellParagraph(0, 0, "段落2")

	// Clear paragraphs
	err = table.ClearCellParagraphs(0, 0)
	if err != nil {
		t.Errorf("failed to clear paragraphs: %v", err)
	}

	// Verify only one empty paragraph remains after clearing
	paragraphs, err := table.GetCellParagraphs(0, 0)
	if err != nil {
		t.Errorf("failed to get paragraphs: %v", err)
	}

	if len(paragraphs) != 1 {
		t.Errorf("expected 1 paragraph after clearing, got %d", len(paragraphs))
	}

	// Test invalid index
	err = table.ClearCellParagraphs(10, 10)
	if err == nil {
		t.Error("expected invalid index to fail, but it succeeded")
	}
}

// TestAddNestedTable tests adding nested tables to a cell
func TestAddNestedTable(t *testing.T) {
	doc := New()

	// Create main table
	mainConfig := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 8000,
	}

	mainTable, err := doc.CreateTable(mainConfig)
	if err != nil {
		t.Fatalf("failed to create main table: %v", err)
	}

	// Create nested table config
	nestedConfig := &TableConfig{
		Rows:  2,
		Cols:  3,
		Width: 3000,
		Data: [][]string{
			{"嵌套1", "嵌套2", "嵌套3"},
			{"数据1", "数据2", "数据3"},
		},
	}

	// Add nested table
	nestedTable, err := mainTable.AddNestedTable(0, 0, nestedConfig)
	if err != nil {
		t.Errorf("failed to add nested table: %v", err)
	}
	if nestedTable == nil {
		t.Error("returned nested table should not be nil")
	}

	// Verify nested table structure
	if nestedTable.GetRowCount() != 2 {
		t.Errorf("expected nested table to have 2 rows, got %d", nestedTable.GetRowCount())
	}
	if nestedTable.GetColumnCount() != 3 {
		t.Errorf("expected nested table to have 3 columns, got %d", nestedTable.GetColumnCount())
	}

	// Verify nested table content
	cellText, err := nestedTable.GetCellText(0, 0)
	if err != nil {
		t.Errorf("failed to get nested table cell content: %v", err)
	}
	if cellText != "嵌套1" {
		t.Errorf("expected nested table content '嵌套1', got '%s'", cellText)
	}

	// Get nested tables list
	nestedTables, err := mainTable.GetNestedTables(0, 0)
	if err != nil {
		t.Errorf("failed to get nested tables list: %v", err)
	}
	if len(nestedTables) != 1 {
		t.Errorf("expected 1 nested table, got %d", len(nestedTables))
	}
}

// TestAddNestedTableInvalidConfig tests nested table with invalid config
func TestAddNestedTableInvalidConfig(t *testing.T) {
	doc := New()

	mainConfig := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 4000,
	}

	mainTable, err := doc.CreateTable(mainConfig)
	if err != nil {
		t.Fatalf("failed to create main table: %v", err)
	}

	// Test invalid row/column counts
	_, err = mainTable.AddNestedTable(0, 0, &TableConfig{Rows: 0, Cols: 2, Width: 2000})
	if err == nil {
		t.Error("expected 0 rows to fail, but it succeeded")
	}

	_, err = mainTable.AddNestedTable(0, 0, &TableConfig{Rows: 2, Cols: 0, Width: 2000})
	if err == nil {
		t.Error("expected 0 columns to fail, but it succeeded")
	}

	// Test invalid cell index
	_, err = mainTable.AddNestedTable(10, 10, &TableConfig{Rows: 2, Cols: 2, Width: 2000})
	if err == nil {
		t.Error("expected invalid index to fail, but it succeeded")
	}
}

// TestAddCellList tests adding lists to a cell
func TestAddCellList(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  3,
		Cols:  2,
		Width: 6000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test adding an unordered list
	bulletListConfig := &CellListConfig{
		Type:         ListTypeBullet,
		BulletSymbol: BulletTypeDot,
		Items:        []string{"项目一", "项目二", "项目三"},
	}

	err = table.AddCellList(0, 0, bulletListConfig)
	if err != nil {
		t.Errorf("failed to add unordered list: %v", err)
	}

	// Verify list item count
	paragraphs, err := table.GetCellParagraphs(0, 0)
	if err != nil {
		t.Errorf("failed to get paragraphs: %v", err)
	}

	// Initial empty paragraph plus 3 list items
	expectedCount := 1 + 3
	if len(paragraphs) != expectedCount {
		t.Errorf("expected %d paragraphs, got %d", expectedCount, len(paragraphs))
	}

	// Test adding an ordered list
	numberListConfig := &CellListConfig{
		Type:  ListTypeNumber,
		Items: []string{"第一步", "第二步", "第三步"},
	}

	err = table.AddCellList(1, 0, numberListConfig)
	if err != nil {
		t.Errorf("failed to add ordered list: %v", err)
	}

	// Test adding a lowercase letter list
	letterListConfig := &CellListConfig{
		Type:  ListTypeLowerLetter,
		Items: []string{"选项a", "选项b"},
	}

	err = table.AddCellList(2, 0, letterListConfig)
	if err != nil {
		t.Errorf("failed to add letter list: %v", err)
	}
}

// TestAddCellListInvalidConfig tests list with invalid config
func TestAddCellListInvalidConfig(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 4000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test nil config
	err = table.AddCellList(0, 0, nil)
	if err == nil {
		t.Error("expected nil config to fail, but it succeeded")
	}

	// Test empty list items
	err = table.AddCellList(0, 0, &CellListConfig{Type: ListTypeBullet, Items: []string{}})
	if err == nil {
		t.Error("expected empty list items to fail, but it succeeded")
	}

	// Test invalid index
	err = table.AddCellList(10, 10, &CellListConfig{Type: ListTypeBullet, Items: []string{"测试"}})
	if err == nil {
		t.Error("expected invalid index to fail, but it succeeded")
	}
}

// TestAddCellImage tests adding images to a cell
func TestAddCellImage(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 6000,
	}

	table, err := doc.AddTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Create test image data
	imageData := createCellTestImage(100, 100)

	// Test adding image from data
	imageInfo, err := doc.AddCellImageFromData(table, 0, 0, imageData, 30)
	if err != nil {
		t.Errorf("failed to add image: %v", err)
	}
	if imageInfo == nil {
		t.Error("returned image info should not be nil")
	}

	// Verify image ID is not empty
	if imageInfo.ID == "" {
		t.Error("image ID should not be empty")
	}

	// Verify relation ID is not empty
	if imageInfo.RelationID == "" {
		t.Error("relation ID should not be empty")
	}

	// Verify cell paragraphs contain image
	paragraphs, err := table.GetCellParagraphs(0, 0)
	if err != nil {
		t.Errorf("failed to get paragraphs: %v", err)
	}

	hasImage := false
	for _, para := range paragraphs {
		for _, run := range para.Runs {
			if run.Drawing != nil {
				hasImage = true
				break
			}
		}
	}

	if !hasImage {
		t.Error("cell should contain an image")
	}
}

// TestAddCellImageWithConfig tests adding images with config
func TestAddCellImageWithConfig(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 6000,
	}

	table, err := doc.AddTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Create test image data
	imageData := createCellTestImage(200, 150)

	// Add image with full config
	imageConfig := &CellImageConfig{
		Data:            imageData,
		Width:           50,
		Height:          40,
		KeepAspectRatio: false,
		AltText:         "测试图片",
		Title:           "单元格图片",
	}

	imageInfo, err := doc.AddCellImage(table, 0, 0, imageConfig)
	if err != nil {
		t.Errorf("failed to add image: %v", err)
	}

	// Verify image config
	if imageInfo.Config == nil {
		t.Error("image config should not be nil")
	} else {
		if imageInfo.Config.AltText != "测试图片" {
			t.Errorf("expected alt text '测试图片', got '%s'", imageInfo.Config.AltText)
		}
		if imageInfo.Config.Title != "单元格图片" {
			t.Errorf("expected title '单元格图片', got '%s'", imageInfo.Config.Title)
		}
	}
}

// TestAddCellImageInvalidCases tests adding images with invalid cases
func TestAddCellImageInvalidCases(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 4000,
	}

	table, err := doc.AddTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test nil table
	_, err = doc.AddCellImage(nil, 0, 0, &CellImageConfig{Data: createCellTestImage(100, 100)})
	if err == nil {
		t.Error("expected nil table to fail, but it succeeded")
	}

	// Test invalid index
	_, err = doc.AddCellImage(table, 10, 10, &CellImageConfig{Data: createCellTestImage(100, 100)})
	if err == nil {
		t.Error("expected invalid index to fail, but it succeeded")
	}

	// Test config without data
	_, err = doc.AddCellImage(table, 0, 0, &CellImageConfig{})
	if err == nil {
		t.Error("expected config without data to fail, but it succeeded")
	}
}

// TestComplexTableStructure tests complex table structure
func TestComplexTableStructure(t *testing.T) {
	doc := New()

	// Create main table
	mainConfig := &TableConfig{
		Rows:  3,
		Cols:  3,
		Width: 9000,
	}

	table, err := doc.AddTable(mainConfig)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// First cell: add multiple paragraphs
	table.AddCellParagraph(0, 0, "第一段")
	table.AddCellFormattedParagraph(0, 0, "格式化段落", &TextFormat{Bold: true})

	// Second cell: add list
	listConfig := &CellListConfig{
		Type:         ListTypeBullet,
		BulletSymbol: BulletTypeDot,
		Items:        []string{"列表项1", "列表项2"},
	}
	table.AddCellList(0, 1, listConfig)

	// Third cell: add nested table
	nestedConfig := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 2500,
		Data: [][]string{
			{"A", "B"},
			{"C", "D"},
		},
	}
	table.AddNestedTable(0, 2, nestedConfig)

	// Fourth cell: add image
	imageData := createCellTestImage(50, 50)
	doc.AddCellImageFromData(table, 1, 0, imageData, 20)

	// Verify complex structure
	paragraphs00, _ := table.GetCellParagraphs(0, 0)
	if len(paragraphs00) < 3 { // initial 1 + added 2
		t.Errorf("cell(0,0) should have at least 3 paragraphs, got %d", len(paragraphs00))
	}

	paragraphs01, _ := table.GetCellParagraphs(0, 1)
	if len(paragraphs01) < 3 { // initial 1 + 2 list items
		t.Errorf("cell(0,1) should have at least 3 paragraphs, got %d", len(paragraphs01))
	}

	nestedTables, _ := table.GetNestedTables(0, 2)
	if len(nestedTables) != 1 {
		t.Errorf("cell(0,2) should have 1 nested table, got %d", len(nestedTables))
	}

	paragraphs10, _ := table.GetCellParagraphs(1, 0)
	hasImage := false
	for _, para := range paragraphs10 {
		for _, run := range para.Runs {
			if run.Drawing != nil {
				hasImage = true
				break
			}
		}
	}
	if !hasImage {
		t.Error("cell(1,0) should contain an image")
	}
}

// TestSaveComplexTable tests saving a complex table
func TestSaveComplexTable(t *testing.T) {
	doc := New()

	// Create table
	config := &TableConfig{
		Rows:  3,
		Cols:  2,
		Width: 6000,
	}

	table, err := doc.AddTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Add complex content
	table.AddCellParagraph(0, 0, "复杂表格测试")
	table.AddCellFormattedParagraph(0, 0, "粗体文本", &TextFormat{Bold: true})

	listConfig := &CellListConfig{
		Type:  ListTypeNumber,
		Items: []string{"第一项", "第二项"},
	}
	table.AddCellList(0, 1, listConfig)

	nestedConfig := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 2000,
		Data: [][]string{
			{"X", "Y"},
			{"Z", "W"},
		},
	}
	table.AddNestedTable(1, 0, nestedConfig)

	// Add image
	imageData := createCellTestImage(80, 60)
	doc.AddCellImageFromData(table, 1, 1, imageData, 25)

	// Save and verify
	outputDir := "test_output"
	if _, err := os.Stat(outputDir); os.IsNotExist(err) {
		os.MkdirAll(outputDir, 0755)
	}

	outputFile := outputDir + "/complex_table_test.docx"
	err = doc.Save(outputFile)
	if err != nil {
		t.Errorf("failed to save document: %v", err)
	}

	// Verify file exists
	if _, err := os.Stat(outputFile); os.IsNotExist(err) {
		t.Error("output file does not exist")
	}

	// Cleanup
	defer os.RemoveAll(outputDir)
}

// TestRomanNumerals tests Roman numeral conversion
func TestRomanNumerals(t *testing.T) {
	testCases := []struct {
		num      int
		expected string
	}{
		{1, "I"},
		{2, "II"},
		{3, "III"},
		{4, "IV"},
		{5, "V"},
		{9, "IX"},
		{10, "X"},
		{40, "XL"},
		{50, "L"},
		{90, "XC"},
		{100, "C"},
		{400, "CD"},
		{500, "D"},
		{900, "CM"},
		{1000, "M"},
		{1999, "MCMXCIX"},
		{2024, "MMXXIV"},
	}

	for _, tc := range testCases {
		result := toRomanUpper(tc.num)
		if result != tc.expected {
			t.Errorf("toRomanUpper(%d) = %s, expected %s", tc.num, result, tc.expected)
		}
	}

	// Test edge cases
	if toRomanUpper(0) != "0" {
		t.Error("0 should return string '0'")
	}

	if toRomanUpper(4000) != "4000" {
		t.Error("4000 should return string '4000'")
	}
}

// TestAddCellListAllTypes tests all list types
func TestAddCellListAllTypes(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  7,
		Cols:  1,
		Width: 3000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	testTypes := []struct {
		listType ListType
		name     string
	}{
		{ListTypeBullet, "bullet list"},
		{ListTypeNumber, "number list"},
		{ListTypeDecimal, "decimal list"},
		{ListTypeLowerLetter, "lowercase letter list"},
		{ListTypeUpperLetter, "uppercase letter list"},
		{ListTypeLowerRoman, "lowercase Roman list"},
		{ListTypeUpperRoman, "uppercase Roman list"},
	}

	for i, tc := range testTypes {
		listConfig := &CellListConfig{
			Type:  tc.listType,
			Items: []string{"项目1", "项目2", "项目3"},
		}

		err := table.AddCellList(i, 0, listConfig)
		if err != nil {
			t.Errorf("failed to add %s: %v", tc.name, err)
		}
	}
}
