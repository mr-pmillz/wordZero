package document

import (
	"testing"
)

// TestCreateTable tests table creation functionality
func TestCreateTable(t *testing.T) {
	doc := New()

	// Test basic table creation
	config := &TableConfig{
		Rows:  3,
		Cols:  4,
		Width: 8000,
		Data: [][]string{
			{"A1", "B1", "C1", "D1"},
			{"A2", "B2", "C2", "D2"},
			{"A3", "B3", "C3", "D3"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Verify table dimensions
	if table.GetRowCount() != 3 {
		t.Errorf("expected 3 rows, got %d", table.GetRowCount())
	}
	if table.GetColumnCount() != 4 {
		t.Errorf("expected 4 columns, got %d", table.GetColumnCount())
	}

	// Verify table content
	cellText, err := table.GetCellText(0, 0)
	if err != nil {
		t.Errorf("failed to get cell content: %v", err)
	}
	if cellText != "A1" {
		t.Errorf("expected cell content 'A1', got '%s'", cellText)
	}

	cellText, err = table.GetCellText(2, 3)
	if err != nil {
		t.Errorf("failed to get cell content: %v", err)
	}
	if cellText != "D3" {
		t.Errorf("expected cell content 'D3', got '%s'", cellText)
	}
}

// TestCreateTableWithInvalidConfig tests table creation with invalid config
func TestCreateTableWithInvalidConfig(t *testing.T) {
	doc := New()

	// Test with 0 rows
	config := &TableConfig{
		Rows:  0,
		Cols:  3,
		Width: 6000,
	}
	_, err := doc.CreateTable(config)
	if err == nil {
		t.Error("expected creation to fail, but it succeeded")
	}

	// Test with 0 columns
	config = &TableConfig{
		Rows:  3,
		Cols:  0,
		Width: 6000,
	}
	_, err = doc.CreateTable(config)
	if err == nil {
		t.Error("expected creation to fail, but it succeeded")
	}

	// Test with mismatched column width count
	config = &TableConfig{
		Rows:      3,
		Cols:      4,
		Width:     6000,
		ColWidths: []int{1000, 2000}, // only 2 column widths, but 4 columns
	}
	_, err = doc.CreateTable(config)
	if err == nil {
		t.Error("expected creation to fail, but it succeeded")
	}
}

// TestAddTable tests adding a table to the document
func TestAddTable(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  3,
		Width: 6000,
	}

	initialTableCount := len(doc.Body.GetTables())
	table, err := doc.AddTable(config)

	if err != nil {
		t.Fatalf("failed to add table: %v", err)
	}

	if len(doc.Body.GetTables()) != initialTableCount+1 {
		t.Errorf("expected table count %d, got %d", initialTableCount+1, len(doc.Body.GetTables()))
	}

	_ = table // use table variable to avoid compiler warning
}

// TestTableCellOperations tests cell operations
func TestTableCellOperations(t *testing.T) {
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

	// Test setting cell content
	err = table.SetCellText(0, 0, "测试内容")
	if err != nil {
		t.Errorf("failed to set cell content: %v", err)
	}

	// Test getting cell content
	cellText, err := table.GetCellText(0, 0)
	if err != nil {
		t.Errorf("failed to get cell content: %v", err)
	}
	if cellText != "测试内容" {
		t.Errorf("expected cell content '测试内容', got '%s'", cellText)
	}

	// Test invalid cell index
	err = table.SetCellText(5, 5, "无效")
	if err == nil {
		t.Error("expected setting invalid cell to fail, but it succeeded")
	}

	_, err = table.GetCellText(5, 5)
	if err == nil {
		t.Error("expected getting invalid cell to fail, but it succeeded")
	}
}

// TestInsertRow tests row insertion functionality
func TestInsertRow(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  3,
		Width: 6000,
		Data: [][]string{
			{"A1", "B1", "C1"},
			{"A2", "B2", "C2"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	initialRowCount := table.GetRowCount()

	// Insert row in the middle
	err = table.InsertRow(1, []string{"A1.5", "B1.5", "C1.5"})
	if err != nil {
		t.Errorf("failed to insert row: %v", err)
	}

	if table.GetRowCount() != initialRowCount+1 {
		t.Errorf("expected %d rows, got %d", initialRowCount+1, table.GetRowCount())
	}

	// Verify inserted content
	cellText, err := table.GetCellText(1, 0)
	if err != nil {
		t.Errorf("failed to get inserted row content: %v", err)
	}
	if cellText != "A1.5" {
		t.Errorf("expected inserted row content 'A1.5', got '%s'", cellText)
	}

	// Test appending row at the end
	err = table.AppendRow([]string{"A末", "B末", "C末"})
	if err != nil {
		t.Errorf("failed to append row: %v", err)
	}

	if table.GetRowCount() != initialRowCount+2 {
		t.Errorf("expected %d rows, got %d", initialRowCount+2, table.GetRowCount())
	}
}

// TestInsertRowInvalidCases tests invalid row insertion cases
func TestInsertRowInvalidCases(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  3,
		Width: 6000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test invalid position
	err = table.InsertRow(-1, []string{"A", "B", "C"})
	if err == nil {
		t.Error("expected inserting at invalid position to fail, but it succeeded")
	}

	err = table.InsertRow(10, []string{"A", "B", "C"})
	if err == nil {
		t.Error("expected inserting at invalid position to fail, but it succeeded")
	}

	// Test too many data columns
	err = table.InsertRow(1, []string{"A", "B", "C", "D", "E"})
	if err == nil {
		t.Error("expected inserting too many columns to fail, but it succeeded")
	}
}

// TestDeleteRow tests row deletion functionality
func TestDeleteRow(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  4,
		Cols:  3,
		Width: 6000,
		Data: [][]string{
			{"A1", "B1", "C1"},
			{"A2", "B2", "C2"},
			{"A3", "B3", "C3"},
			{"A4", "B4", "C4"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	initialRowCount := table.GetRowCount()

	// Delete 2nd row (index 1)
	err = table.DeleteRow(1)
	if err != nil {
		t.Errorf("failed to delete row: %v", err)
	}

	if table.GetRowCount() != initialRowCount-1 {
		t.Errorf("expected %d rows, got %d", initialRowCount-1, table.GetRowCount())
	}

	// Verify content after deletion (original 3rd row should now be 2nd row)
	cellText, err := table.GetCellText(1, 0)
	if err != nil {
		t.Errorf("failed to get content after deletion: %v", err)
	}
	if cellText != "A3" {
		t.Errorf("expected content after deletion 'A3', got '%s'", cellText)
	}
}

// TestDeleteRowInvalidCases tests invalid row deletion cases
func TestDeleteRowInvalidCases(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  1,
		Cols:  3,
		Width: 6000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test deleting the only row
	err = table.DeleteRow(0)
	if err == nil {
		t.Error("expected deleting the only row to fail, but it succeeded")
	}

	// Add a row to test invalid index
	table.AppendRow([]string{"A", "B", "C"})

	// Test invalid index
	err = table.DeleteRow(-1)
	if err == nil {
		t.Error("expected deleting with invalid index to fail, but it succeeded")
	}

	err = table.DeleteRow(10)
	if err == nil {
		t.Error("expected deleting with invalid index to fail, but it succeeded")
	}
}

// TestDeleteRows tests deleting multiple rows
func TestDeleteRows(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  5,
		Cols:  2,
		Width: 4000,
		Data: [][]string{
			{"A1", "B1"},
			{"A2", "B2"},
			{"A3", "B3"},
			{"A4", "B4"},
			{"A5", "B5"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	initialRowCount := table.GetRowCount()

	// Delete rows 2 to 4 (index 1 to 3)
	err = table.DeleteRows(1, 3)
	if err != nil {
		t.Errorf("failed to delete multiple rows: %v", err)
	}

	expectedRowCount := initialRowCount - 3
	if table.GetRowCount() != expectedRowCount {
		t.Errorf("expected %d rows, got %d", expectedRowCount, table.GetRowCount())
	}

	// Verify remaining content
	cellText, err := table.GetCellText(1, 0)
	if err != nil {
		t.Errorf("failed to get content after deletion: %v", err)
	}
	if cellText != "A5" {
		t.Errorf("expected content after deletion 'A5', got '%s'", cellText)
	}
}

// TestInsertColumn tests column insertion functionality
func TestInsertColumn(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  3,
		Cols:  2,
		Width: 4000,
		Data: [][]string{
			{"A1", "B1"},
			{"A2", "B2"},
			{"A3", "B3"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	initialColCount := table.GetColumnCount()

	// Insert column in the middle
	err = table.InsertColumn(1, []string{"C1", "C2", "C3"}, 1000)
	if err != nil {
		t.Errorf("failed to insert column: %v", err)
	}

	if table.GetColumnCount() != initialColCount+1 {
		t.Errorf("expected %d columns, got %d", initialColCount+1, table.GetColumnCount())
	}

	// Verify inserted content
	cellText, err := table.GetCellText(0, 1)
	if err != nil {
		t.Errorf("failed to get inserted column content: %v", err)
	}
	if cellText != "C1" {
		t.Errorf("expected inserted column content 'C1', got '%s'", cellText)
	}

	// Test appending column at the end
	err = table.AppendColumn([]string{"D1", "D2", "D3"}, 1000)
	if err != nil {
		t.Errorf("failed to append column: %v", err)
	}

	if table.GetColumnCount() != initialColCount+2 {
		t.Errorf("expected %d columns, got %d", initialColCount+2, table.GetColumnCount())
	}
}

// TestDeleteColumn tests column deletion functionality
func TestDeleteColumn(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  3,
		Cols:  4,
		Width: 8000,
		Data: [][]string{
			{"A1", "B1", "C1", "D1"},
			{"A2", "B2", "C2", "D2"},
			{"A3", "B3", "C3", "D3"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	initialColCount := table.GetColumnCount()

	// Delete 2nd column (index 1)
	err = table.DeleteColumn(1)
	if err != nil {
		t.Errorf("failed to delete column: %v", err)
	}

	if table.GetColumnCount() != initialColCount-1 {
		t.Errorf("expected %d columns, got %d", initialColCount-1, table.GetColumnCount())
	}

	// Verify content after deletion (original 3rd column should now be 2nd column)
	cellText, err := table.GetCellText(0, 1)
	if err != nil {
		t.Errorf("failed to get content after deletion: %v", err)
	}
	if cellText != "C1" {
		t.Errorf("expected content after deletion 'C1', got '%s'", cellText)
	}
}

// TestDeleteColumns tests deleting multiple columns
func TestDeleteColumns(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  5,
		Width: 10000,
		Data: [][]string{
			{"A1", "B1", "C1", "D1", "E1"},
			{"A2", "B2", "C2", "D2", "E2"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	initialColCount := table.GetColumnCount()

	// Delete columns 2 to 4 (index 1 to 3)
	err = table.DeleteColumns(1, 3)
	if err != nil {
		t.Errorf("failed to delete multiple columns: %v", err)
	}

	expectedColCount := initialColCount - 3
	if table.GetColumnCount() != expectedColCount {
		t.Errorf("expected %d columns, got %d", expectedColCount, table.GetColumnCount())
	}

	// Verify remaining content
	cellText, err := table.GetCellText(0, 1)
	if err != nil {
		t.Errorf("failed to get content after deletion: %v", err)
	}
	if cellText != "E1" {
		t.Errorf("expected content after deletion 'E1', got '%s'", cellText)
	}
}

// TestClearTable tests clearing table content
func TestClearTable(t *testing.T) {
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

	// Clear the table
	table.ClearTable()

	// Verify all cells are empty
	for i := 0; i < table.GetRowCount(); i++ {
		for j := 0; j < table.GetColumnCount(); j++ {
			cellText, err := table.GetCellText(i, j)
			if err != nil {
				t.Errorf("failed to get cell content after clearing: %v", err)
			}
			if cellText != "" {
				t.Errorf("expected cell to be empty after clearing, got '%s'", cellText)
			}
		}
	}

	// Verify table structure is preserved
	if table.GetRowCount() != 2 {
		t.Errorf("expected 2 rows after clearing, got %d", table.GetRowCount())
	}
	if table.GetColumnCount() != 2 {
		t.Errorf("expected 2 columns after clearing, got %d", table.GetColumnCount())
	}
}

// TestCopyTable tests table copy functionality
func TestCopyTable(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 4000,
		Data: [][]string{
			{"原始1", "原始2"},
			{"原始3", "原始4"},
		},
	}

	originalTable, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create original table: %v", err)
	}

	// Copy the table
	copiedTable := originalTable.CopyTable()
	if copiedTable == nil {
		t.Fatal("failed to copy table")
	}

	// Verify copied table structure
	if copiedTable.GetRowCount() != originalTable.GetRowCount() {
		t.Errorf("copied table row count mismatch: expected %d, got %d",
			originalTable.GetRowCount(), copiedTable.GetRowCount())
	}
	if copiedTable.GetColumnCount() != originalTable.GetColumnCount() {
		t.Errorf("copied table column count mismatch: expected %d, got %d",
			originalTable.GetColumnCount(), copiedTable.GetColumnCount())
	}

	// Verify copied table content
	for i := 0; i < originalTable.GetRowCount(); i++ {
		for j := 0; j < originalTable.GetColumnCount(); j++ {
			originalText, _ := originalTable.GetCellText(i, j)
			copiedText, _ := copiedTable.GetCellText(i, j)
			if originalText != copiedText {
				t.Errorf("copied table content mismatch: position(%d,%d) expected '%s', got '%s'",
					i, j, originalText, copiedText)
			}
		}
	}

	// Modify the copied table to verify independence
	err = copiedTable.SetCellText(0, 0, "修改后")
	if err != nil {
		t.Errorf("failed to modify copied table: %v", err)
	}

	originalText, _ := originalTable.GetCellText(0, 0)
	copiedText, _ := copiedTable.GetCellText(0, 0)

	if originalText == copiedText {
		t.Error("copied table is not independent, modification affected the original table")
	}
}

// TestTableWithCustomColumnWidths tests table with custom column widths
func TestTableWithCustomColumnWidths(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:      2,
		Cols:      3,
		Width:     6000,
		ColWidths: []int{1000, 2000, 3000},
		Data: [][]string{
			{"窄列", "中列", "宽列"},
			{"A", "B", "C"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table with custom column widths: %v", err)
	}

	// Verify table created successfully
	if table.GetRowCount() != 2 {
		t.Errorf("expected 2 rows, got %d", table.GetRowCount())
	}
	if table.GetColumnCount() != 3 {
		t.Errorf("expected 3 columns, got %d", table.GetColumnCount())
	}

	// Verify grid column count
	if len(table.Grid.Cols) != 3 {
		t.Errorf("expected 3 grid columns, got %d", len(table.Grid.Cols))
	}
}

// TestTableElementType tests table element type
func TestTableElementType(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  1,
		Cols:  1,
		Width: 2000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test table element type
	if table.ElementType() != "table" {
		t.Errorf("expected table element type 'table', got '%s'", table.ElementType())
	}

	// Test paragraph element type
	para := doc.AddParagraph("测试段落")
	if para.ElementType() != "paragraph" {
		t.Errorf("expected paragraph element type 'paragraph', got '%s'", para.ElementType())
	}
}

// TestCellFormattedText tests cell rich text functionality
func TestCellFormattedText(t *testing.T) {
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

	// Test setting rich text content
	format := &TextFormat{
		Bold:       true,
		Italic:     true,
		FontSize:   14,
		FontColor:  "FF0000",
		FontFamily: "Arial",
	}

	err = table.SetCellFormattedText(0, 0, "富文本测试", format)
	if err != nil {
		t.Errorf("failed to set rich text content: %v", err)
	}

	// Verify content
	cellText, err := table.GetCellText(0, 0)
	if err != nil {
		t.Errorf("failed to get cell content: %v", err)
	}
	if cellText != "富文本测试" {
		t.Errorf("expected content '富文本测试', got '%s'", cellText)
	}

	// Test appending formatted text
	err = table.AddCellFormattedText(0, 0, " 追加文本", &TextFormat{Bold: false, FontColor: "00FF00"})
	if err != nil {
		t.Errorf("failed to add formatted text: %v", err)
	}
}

// TestCellFormat tests cell format settings
func TestCellFormat(t *testing.T) {
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

	// Set cell content
	err = table.SetCellText(0, 0, "格式测试")
	if err != nil {
		t.Errorf("failed to set cell content: %v", err)
	}

	// Test setting cell format
	format := &CellFormat{
		TextFormat: &TextFormat{
			Bold:     true,
			FontSize: 16,
		},
		HorizontalAlign: CellAlignCenter,
		VerticalAlign:   CellVAlignCenter,
	}

	err = table.SetCellFormat(0, 0, format)
	if err != nil {
		t.Errorf("failed to set cell format: %v", err)
	}

	// Get and verify format
	retrievedFormat, err := table.GetCellFormat(0, 0)
	if err != nil {
		t.Errorf("failed to get cell format: %v", err)
	}

	if retrievedFormat.HorizontalAlign != CellAlignCenter {
		t.Errorf("expected horizontal alignment 'center', got '%s'", retrievedFormat.HorizontalAlign)
	}

	if retrievedFormat.VerticalAlign != CellVAlignCenter {
		t.Errorf("expected vertical alignment 'center', got '%s'", retrievedFormat.VerticalAlign)
	}

	if retrievedFormat.TextFormat == nil || !retrievedFormat.TextFormat.Bold {
		t.Error("expected text format to be bold")
	}
}

// TestCellMergeHorizontal tests horizontal cell merging
func TestCellMergeHorizontal(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  3,
		Cols:  4,
		Width: 8000,
		Data: [][]string{
			{"A1", "B1", "C1", "D1"},
			{"A2", "B2", "C2", "D2"},
			{"A3", "B3", "C3", "D3"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	initialColCount := table.GetColumnCount()

	// Merge columns 2 to 4 (index 1 to 3) in the first row
	err = table.MergeCellsHorizontal(0, 1, 3)
	if err != nil {
		t.Errorf("failed to merge cells horizontally: %v", err)
	}

	// Verify first row column count decreased after merging
	if len(table.Rows[0].Cells) != initialColCount-2 {
		t.Errorf("expected first row column count %d, got %d", initialColCount-2, len(table.Rows[0].Cells))
	}

	// Verify merge status
	isMerged, err := table.IsCellMerged(0, 1)
	if err != nil {
		t.Errorf("failed to check merge status: %v", err)
	}
	if !isMerged {
		t.Error("expected cell to be merged")
	}

	// Get merge info
	mergeInfo, err := table.GetMergedCellInfo(0, 1)
	if err != nil {
		t.Errorf("failed to get merge info: %v", err)
	}

	if !mergeInfo["is_merged"].(bool) {
		t.Error("expected cell to be in merged state")
	}

	if mergeInfo["horizontal_span"].(int) != 3 {
		t.Errorf("expected horizontal span 3, got %d", mergeInfo["horizontal_span"].(int))
	}
}

// TestCellMergeVertical tests vertical cell merging
func TestCellMergeVertical(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  4,
		Cols:  3,
		Width: 6000,
		Data: [][]string{
			{"A1", "B1", "C1"},
			{"A2", "B2", "C2"},
			{"A3", "B3", "C3"},
			{"A4", "B4", "C4"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Merge rows 1 to 3 (index 0 to 2) in column 2
	err = table.MergeCellsVertical(0, 2, 1)
	if err != nil {
		t.Errorf("failed to merge cells vertically: %v", err)
	}

	// Verify merge status
	isMerged, err := table.IsCellMerged(0, 1)
	if err != nil {
		t.Errorf("failed to check merge status: %v", err)
	}
	if !isMerged {
		t.Error("expected cell to be merged")
	}

	// Verify merged cells also have merge markers
	isMerged, err = table.IsCellMerged(1, 1)
	if err != nil {
		t.Errorf("failed to check merge status: %v", err)
	}
	if !isMerged {
		t.Error("expected merged cell to also have merge marker")
	}
}

// TestCellMergeRange tests range merging
func TestCellMergeRange(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  4,
		Cols:  4,
		Width: 8000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Merge 2x2 range (rows 0-1, columns 1-2)
	err = table.MergeCellsRange(0, 1, 1, 2)
	if err != nil {
		t.Errorf("failed to merge range: %v", err)
	}

	// Verify merge status
	isMerged, err := table.IsCellMerged(0, 1)
	if err != nil {
		t.Errorf("failed to check merge status: %v", err)
	}
	if !isMerged {
		t.Error("expected cell to be merged")
	}
}

// TestUnmergeCells tests unmerging cells
func TestUnmergeCells(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  3,
		Cols:  3,
		Width: 6000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// First perform horizontal merge
	err = table.MergeCellsHorizontal(0, 0, 1)
	if err != nil {
		t.Errorf("failed to merge horizontally: %v", err)
	}

	// Verify merge status
	isMerged, err := table.IsCellMerged(0, 0)
	if err != nil {
		t.Errorf("failed to check merge status: %v", err)
	}
	if !isMerged {
		t.Error("expected cell to be merged")
	}

	// Unmerge cells
	err = table.UnmergeCells(0, 0)
	if err != nil {
		t.Errorf("failed to unmerge cells: %v", err)
	}

	// Verify status after unmerging
	isMerged, err = table.IsCellMerged(0, 0)
	if err != nil {
		t.Errorf("failed to check merge status: %v", err)
	}
	if isMerged {
		t.Error("expected cell to be unmerged")
	}
}

// TestCellContentOperations tests cell content operations
func TestCellContentOperations(t *testing.T) {
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

	// Set formatted content
	format := &TextFormat{
		Bold:     true,
		FontSize: 12,
	}
	err = table.SetCellFormattedText(0, 0, "测试内容", format)
	if err != nil {
		t.Errorf("failed to set formatted content: %v", err)
	}

	// Clear content but preserve format
	err = table.ClearCellContent(0, 0)
	if err != nil {
		t.Errorf("failed to clear cell content: %v", err)
	}

	// Verify content is cleared
	content, err := table.GetCellText(0, 0)
	if err != nil {
		t.Errorf("failed to get cell content: %v", err)
	}
	if content != "" {
		t.Errorf("expected content to be empty, got '%s'", content)
	}

	// Set content again
	err = table.SetCellText(0, 0, "新内容")
	if err != nil {
		t.Errorf("failed to set new content: %v", err)
	}

	// Clear format but preserve content
	err = table.ClearCellFormat(0, 0)
	if err != nil {
		t.Errorf("failed to clear cell format: %v", err)
	}

	// Verify content is preserved
	content, err = table.GetCellText(0, 0)
	if err != nil {
		t.Errorf("failed to get cell content: %v", err)
	}
	if content != "新内容" {
		t.Errorf("expected content '新内容', got '%s'", content)
	}
}

// TestCellMergeInvalidCases tests invalid merge cases
func TestCellMergeInvalidCases(t *testing.T) {
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

	// Test invalid horizontal merge
	err = table.MergeCellsHorizontal(0, 0, 0)
	if err == nil {
		t.Error("expected same-column merge to fail, but it succeeded")
	}

	err = table.MergeCellsHorizontal(0, 1, 0)
	if err == nil {
		t.Error("expected reverse merge to fail, but it succeeded")
	}

	// Test invalid vertical merge
	err = table.MergeCellsVertical(0, 0, 0)
	if err == nil {
		t.Error("expected same-row merge to fail, but it succeeded")
	}

	err = table.MergeCellsVertical(1, 0, 0)
	if err == nil {
		t.Error("expected reverse merge to fail, but it succeeded")
	}

	// Test invalid index
	err = table.MergeCellsHorizontal(-1, 0, 1)
	if err == nil {
		t.Error("expected invalid row index to fail, but it succeeded")
	}

	err = table.MergeCellsVertical(0, 1, -1)
	if err == nil {
		t.Error("expected invalid column index to fail, but it succeeded")
	}
}

// TestCellPadding tests cell padding
func TestCellPadding(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  1,
		Cols:  1,
		Width: 2000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test setting padding
	err = table.SetCellPadding(0, 0, 10)
	if err != nil {
		t.Errorf("failed to set cell padding: %v", err)
	}

	// Test invalid index
	err = table.SetCellPadding(5, 5, 10)
	if err == nil {
		t.Error("expected invalid index to fail, but it succeeded")
	}
}

// TestCellTextDirection tests cell text direction settings
func TestCellTextDirection(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  3,
		Cols:  3,
		Width: 6000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test setting different text directions
	testCases := []struct {
		name      string
		direction CellTextDirection
		row       int
		col       int
	}{
		{"从左到右", TextDirectionLR, 0, 0},
		{"从上到下", TextDirectionTB, 0, 1},
		{"从下到上", TextDirectionBT, 0, 2},
		{"从右到左", TextDirectionRL, 1, 0},
		{"从上到下垂直", TextDirectionTBV, 1, 1},
		{"从下到上垂直", TextDirectionBTV, 1, 2},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			// Set cell text
			err := table.SetCellText(tc.row, tc.col, tc.name)
			if err != nil {
				t.Errorf("failed to set cell text: %v", err)
			}

			// Set text direction
			err = table.SetCellTextDirection(tc.row, tc.col, tc.direction)
			if err != nil {
				t.Errorf("failed to set text direction: %v", err)
			}

			// Verify text direction
			actualDirection, err := table.GetCellTextDirection(tc.row, tc.col)
			if err != nil {
				t.Errorf("failed to get text direction: %v", err)
			}

			if actualDirection != tc.direction {
				t.Errorf("text direction mismatch, expected: %s, got: %s", tc.direction, actualDirection)
			}
		})
	}
}

// TestCellFormatWithTextDirection tests setting text direction via CellFormat
func TestCellFormatWithTextDirection(t *testing.T) {
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

	// Set full format via CellFormat, including text direction
	format := &CellFormat{
		TextFormat: &TextFormat{
			Bold:     true,
			FontSize: 14,
		},
		HorizontalAlign: CellAlignCenter,
		VerticalAlign:   CellVAlignCenter,
		TextDirection:   TextDirectionTB, // top to bottom
	}

	err = table.SetCellText(0, 0, "竖排文字")
	if err != nil {
		t.Errorf("failed to set cell text: %v", err)
	}

	err = table.SetCellFormat(0, 0, format)
	if err != nil {
		t.Errorf("failed to set cell format: %v", err)
	}

	// Verify format is set correctly
	retrievedFormat, err := table.GetCellFormat(0, 0)
	if err != nil {
		t.Errorf("failed to get cell format: %v", err)
	}

	if retrievedFormat.TextDirection != TextDirectionTB {
		t.Errorf("text direction mismatch, expected: %s, got: %s", TextDirectionTB, retrievedFormat.TextDirection)
	}

	if retrievedFormat.HorizontalAlign != CellAlignCenter {
		t.Errorf("horizontal alignment mismatch, expected: %s, got: %s", CellAlignCenter, retrievedFormat.HorizontalAlign)
	}

	if retrievedFormat.VerticalAlign != CellVAlignCenter {
		t.Errorf("vertical alignment mismatch, expected: %s, got: %s", CellVAlignCenter, retrievedFormat.VerticalAlign)
	}
}

// TestTextDirectionConstants tests text direction constants
func TestTextDirectionConstants(t *testing.T) {
	directions := []CellTextDirection{
		TextDirectionLR,
		TextDirectionTB,
		TextDirectionBT,
		TextDirectionRL,
		TextDirectionTBV,
		TextDirectionBTV,
	}

	expectedValues := []string{
		"lrTb",
		"tbRl",
		"btLr",
		"rlTb",
		"tbLrV",
		"btLrV",
	}

	for i, direction := range directions {
		if string(direction) != expectedValues[i] {
			t.Errorf("text direction constant value mismatch, expected: %s, got: %s", expectedValues[i], string(direction))
		}
	}
}

// TestRowHeight tests row height settings
func TestRowHeight(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  3,
		Cols:  2,
		Width: 4000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test setting exact row height
	heightConfig := &RowHeightConfig{
		Height: 30,
		Rule:   RowHeightExact,
	}

	err = table.SetRowHeight(0, heightConfig)
	if err != nil {
		t.Errorf("failed to set row height: %v", err)
	}

	// Test getting row height
	retrievedConfig, err := table.GetRowHeight(0)
	if err != nil {
		t.Errorf("failed to get row height: %v", err)
	}

	if retrievedConfig.Height != 30 {
		t.Errorf("expected row height 30, got %d", retrievedConfig.Height)
	}

	if retrievedConfig.Rule != RowHeightExact {
		t.Errorf("expected row height rule %s, got %s", RowHeightExact, retrievedConfig.Rule)
	}

	// Test batch setting row height
	batchConfig := &RowHeightConfig{
		Height: 25,
		Rule:   RowHeightMinimum,
	}

	err = table.SetRowHeightRange(1, 2, batchConfig)
	if err != nil {
		t.Errorf("failed to batch set row height: %v", err)
	}

	// Verify batch setting results
	for i := 1; i <= 2; i++ {
		config, err := table.GetRowHeight(i)
		if err != nil {
			t.Errorf("failed to get row %d height: %v", i, err)
		}
		if config.Height != 25 {
			t.Errorf("row %d expected height 25, got %d", i, config.Height)
		}
		if config.Rule != RowHeightMinimum {
			t.Errorf("row %d expected rule %s, got %s", i, RowHeightMinimum, config.Rule)
		}
	}

	// Test invalid index
	err = table.SetRowHeight(10, heightConfig)
	if err == nil {
		t.Error("expected setting invalid row index to fail, but it succeeded")
	}

	_, err = table.GetRowHeight(10)
	if err == nil {
		t.Error("expected getting invalid row index to fail, but it succeeded")
	}
}

// TestTableLayout tests table layout and positioning
func TestTableLayout(t *testing.T) {
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

	// Test setting table layout
	layoutConfig := &TableLayoutConfig{
		Alignment: TableAlignCenter,
		TextWrap:  TextWrapNone,
		Position:  PositionInline,
	}

	err = table.SetTableLayout(layoutConfig)
	if err != nil {
		t.Errorf("failed to set table layout: %v", err)
	}

	// Test getting table layout
	retrievedLayout := table.GetTableLayout()
	if retrievedLayout.Alignment != TableAlignCenter {
		t.Errorf("expected alignment %s, got %s", TableAlignCenter, retrievedLayout.Alignment)
	}

	// Test shortcut method for setting alignment
	err = table.SetTableAlignment(TableAlignRight)
	if err != nil {
		t.Errorf("failed to set table alignment: %v", err)
	}

	retrievedLayout = table.GetTableLayout()
	if retrievedLayout.Alignment != TableAlignRight {
		t.Errorf("expected alignment %s, got %s", TableAlignRight, retrievedLayout.Alignment)
	}
}

// TestTablePageBreak tests table page break control
func TestTablePageBreak(t *testing.T) {
	doc := New()

	config := &TableConfig{
		Rows:  4,
		Cols:  2,
		Width: 4000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test setting row keep together (no page split)
	err = table.SetRowKeepTogether(0, true)
	if err != nil {
		t.Errorf("failed to set row keep together: %v", err)
	}

	// Test checking if row keeps together
	keepTogether, err := table.IsRowKeepTogether(0)
	if err != nil {
		t.Errorf("failed to check row keep together setting: %v", err)
	}
	if !keepTogether {
		t.Error("expected row keep together to be true, got false")
	}

	// Test setting header row
	err = table.SetRowAsHeader(0, true)
	if err != nil {
		t.Errorf("failed to set header row: %v", err)
	}

	// Test checking if row is header
	isHeader, err := table.IsRowHeader(0)
	if err != nil {
		t.Errorf("failed to check header row setting: %v", err)
	}
	if !isHeader {
		t.Error("expected row 0 to be a header row, but it is not")
	}

	// Test setting header row range
	err = table.SetHeaderRows(0, 1)
	if err != nil {
		t.Errorf("failed to set header row range: %v", err)
	}

	// Verify header row range setting
	for i := 0; i <= 1; i++ {
		isHeader, err := table.IsRowHeader(i)
		if err != nil {
			t.Errorf("failed to check row %d header setting: %v", i, err)
		}
		if !isHeader {
			t.Errorf("expected row %d to be a header row, but it is not", i)
		}
	}

	// Test table break info
	breakInfo := table.GetTableBreakInfo()
	if breakInfo["total_rows"] != 4 {
		t.Errorf("expected total rows 4, got %v", breakInfo["total_rows"])
	}
	if breakInfo["header_rows"] != 2 {
		t.Errorf("expected header rows 2, got %v", breakInfo["header_rows"])
	}

	// Test table page break config
	pageBreakConfig := &TablePageBreakConfig{
		KeepWithNext:    true,
		KeepLines:       true,
		PageBreakBefore: false,
		WidowControl:    true,
	}

	err = table.SetTablePageBreak(pageBreakConfig)
	if err != nil {
		t.Errorf("failed to set table page break config: %v", err)
	}

	// Test row keep with next
	err = table.SetRowKeepWithNext(1, true)
	if err != nil {
		t.Errorf("failed to set row keep with next: %v", err)
	}

	// Test invalid index
	err = table.SetRowKeepTogether(10, true)
	if err == nil {
		t.Error("expected setting invalid row index to fail, but it succeeded")
	}

	err = table.SetRowAsHeader(10, true)
	if err == nil {
		t.Error("expected setting invalid row index to fail, but it succeeded")
	}

	_, err = table.IsRowHeader(10)
	if err == nil {
		t.Error("expected checking invalid row index to fail, but it succeeded")
	}

	_, err = table.IsRowKeepTogether(10)
	if err == nil {
		t.Error("expected checking invalid row index to fail, but it succeeded")
	}
}

// TestRowHeightConstants tests row height rule constants
func TestRowHeightConstants(t *testing.T) {
	// Verify row height rule constants are correctly defined
	if RowHeightAuto != "auto" {
		t.Errorf("expected RowHeightAuto to be 'auto', got '%s'", RowHeightAuto)
	}
	if RowHeightMinimum != "atLeast" {
		t.Errorf("expected RowHeightMinimum to be 'atLeast', got '%s'", RowHeightMinimum)
	}
	if RowHeightExact != "exact" {
		t.Errorf("expected RowHeightExact to be 'exact', got '%s'", RowHeightExact)
	}
}

// TestTableAlignmentConstants tests table alignment constants
func TestTableAlignmentConstants(t *testing.T) {
	// Verify table alignment constants are correctly defined
	if TableAlignLeft != "left" {
		t.Errorf("expected TableAlignLeft to be 'left', got '%s'", TableAlignLeft)
	}
	if TableAlignCenter != "center" {
		t.Errorf("expected TableAlignCenter to be 'center', got '%s'", TableAlignCenter)
	}
	if TableAlignRight != "right" {
		t.Errorf("expected TableAlignRight to be 'right', got '%s'", TableAlignRight)
	}
}

// TestTableRowPropertiesExtensions tests TableRowProperties extension methods
func TestTableRowPropertiesExtensions(t *testing.T) {
	trp := &TableRowProperties{}

	// Test SetCantSplit method
	trp.SetCantSplit(true)
	if trp.CantSplit == nil || trp.CantSplit.Val != "1" {
		t.Error("failed to set CantSplit")
	}

	trp.SetCantSplit(false)
	if trp.CantSplit != nil {
		t.Error("failed to clear CantSplit")
	}

	// Test SetTblHeader method
	trp.SetTblHeader(true)
	if trp.TblHeader == nil || trp.TblHeader.Val != "1" {
		t.Error("failed to set TblHeader")
	}

	trp.SetTblHeader(false)
	if trp.TblHeader != nil {
		t.Error("failed to clear TblHeader")
	}
}
