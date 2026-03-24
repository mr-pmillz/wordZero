package document

import (
	"fmt"
	"testing"
)

// TestCellIterator tests basic cell iterator functionality
func TestCellIterator(t *testing.T) {
	// Create a 3x3 test table
	doc := New()
	config := &TableConfig{
		Rows:  3,
		Cols:  3,
		Width: 5000,
		Data: [][]string{
			{"A1", "B1", "C1"},
			{"A2", "B2", "C2"},
			{"A3", "B3", "C3"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test iterator creation
	iterator := table.NewCellIterator()
	if iterator == nil {
		t.Fatal("failed to create iterator")
	}

	// Test Total method
	expectedTotal := 9
	if iterator.Total() != expectedTotal {
		t.Errorf("Total() expected %d, got %d", expectedTotal, iterator.Total())
	}

	// Test iterator traversal
	cellCount := 0
	expectedCells := []struct {
		row  int
		col  int
		text string
	}{
		{0, 0, "A1"}, {0, 1, "B1"}, {0, 2, "C1"},
		{1, 0, "A2"}, {1, 1, "B2"}, {1, 2, "C2"},
		{2, 0, "A3"}, {2, 1, "B3"}, {2, 2, "C3"},
	}

	for iterator.HasNext() {
		cellInfo, err := iterator.Next()
		if err != nil {
			t.Fatalf("iterator Next() failed: %v", err)
		}

		if cellCount >= len(expectedCells) {
			t.Fatalf("iterator returned too many cells")
		}

		expected := expectedCells[cellCount]
		if cellInfo.Row != expected.row || cellInfo.Col != expected.col {
			t.Errorf("cell position mismatch: expected (%d,%d), got (%d,%d)",
				expected.row, expected.col, cellInfo.Row, cellInfo.Col)
		}

		if cellInfo.Text != expected.text {
			t.Errorf("cell text mismatch: expected '%s', got '%s'",
				expected.text, cellInfo.Text)
		}

		if cellInfo.Cell == nil {
			t.Error("cell reference is nil")
		}

		// Test IsLast flag
		if cellCount == len(expectedCells)-1 && !cellInfo.IsLast {
			t.Error("last cell's IsLast flag should be true")
		}

		cellCount++
	}

	if cellCount != expectedTotal {
		t.Errorf("iterated cell count mismatch: expected %d, got %d", expectedTotal, cellCount)
	}
}

// TestCellIteratorReset tests iterator reset functionality
func TestCellIteratorReset(t *testing.T) {
	doc := New()
	config := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 3000,
		Data: [][]string{
			{"A1", "B1"},
			{"A2", "B2"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}
	iterator := table.NewCellIterator()

	// Iterate some cells first
	iterator.Next()
	iterator.Next()

	// Check current position
	row, col := iterator.Current()
	if row != 1 || col != 0 {
		t.Errorf("iterator position incorrect: expected (1,0), got (%d,%d)", row, col)
	}

	// Reset iterator
	iterator.Reset()

	// Check position after reset
	row, col = iterator.Current()
	if row != 0 || col != 0 {
		t.Errorf("position after reset incorrect: expected (0,0), got (%d,%d)", row, col)
	}

	// Ensure it can iterate again
	if !iterator.HasNext() {
		t.Error("should have next cell after reset")
	}
}

// TestCellIteratorProgress tests progress calculation
func TestCellIteratorProgress(t *testing.T) {
	doc := New()
	config := &TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 3000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}
	iterator := table.NewCellIterator()

	// Initial progress should be 0
	if iterator.Progress() != 0.0 {
		t.Errorf("initial progress should be 0.0, got %f", iterator.Progress())
	}

	// Iterate one cell
	iterator.Next()
	expectedProgress := 0.25 // 1/4
	if iterator.Progress() != expectedProgress {
		t.Errorf("progress after one cell should be %f, got %f", expectedProgress, iterator.Progress())
	}

	// Iterate to the end
	for iterator.HasNext() {
		iterator.Next()
	}

	if iterator.Progress() != 1.0 {
		t.Errorf("progress after completion should be 1.0, got %f", iterator.Progress())
	}
}

// TestTableForEach tests the ForEach method
func TestTableForEach(t *testing.T) {
	doc := New()
	config := &TableConfig{
		Rows:  2,
		Cols:  3,
		Width: 4000,
		Data: [][]string{
			{"A1", "B1", "C1"},
			{"A2", "B2", "C2"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test ForEach traversal
	var visitedCells []string
	err = table.ForEach(func(row, col int, cell *TableCell, text string) error {
		visitedCells = append(visitedCells, fmt.Sprintf("%d-%d:%s", row, col, text))
		return nil
	})

	if err != nil {
		t.Fatalf("ForEach execution failed: %v", err)
	}

	expectedCells := []string{
		"0-0:A1", "0-1:B1", "0-2:C1",
		"1-0:A2", "1-1:B2", "1-2:C2",
	}

	if len(visitedCells) != len(expectedCells) {
		t.Errorf("visited cell count mismatch: expected %d, got %d", len(expectedCells), len(visitedCells))
	}

	for i, expected := range expectedCells {
		if i < len(visitedCells) && visitedCells[i] != expected {
			t.Errorf("cell visit order incorrect: expected '%s', got '%s'", expected, visitedCells[i])
		}
	}
}

// TestForEachInRow tests row traversal
//
//nolint:dupl
func TestForEachInRow(t *testing.T) {
	doc := New()
	config := &TableConfig{
		Rows:  3,
		Cols:  3,
		Width: 4000,
		Data: [][]string{
			{"A1", "B1", "C1"},
			{"A2", "B2", "C2"},
			{"A3", "B3", "C3"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test traversing row 2 (index 1)
	var visitedCells []string
	err = table.ForEachInRow(1, func(col int, cell *TableCell, text string) error {
		visitedCells = append(visitedCells, fmt.Sprintf("%d:%s", col, text))
		return nil
	})

	if err != nil {
		t.Fatalf("ForEachInRow execution failed: %v", err)
	}

	expectedCells := []string{"0:A2", "1:B2", "2:C2"}
	if len(visitedCells) != len(expectedCells) {
		t.Errorf("visited cell count mismatch: expected %d, got %d", len(expectedCells), len(visitedCells))
	}

	for i, expected := range expectedCells {
		if i < len(visitedCells) && visitedCells[i] != expected {
			t.Errorf("cell visit order incorrect: expected '%s', got '%s'", expected, visitedCells[i])
		}
	}

	// Test invalid row index
	err = table.ForEachInRow(5, func(col int, cell *TableCell, text string) error {
		return nil
	})
	if err == nil {
		t.Error("should return invalid row index error")
	}
}

// TestForEachInColumn tests column traversal
//nolint:dupl
func TestForEachInColumn(t *testing.T) {
	doc := New()
	config := &TableConfig{
		Rows:  3,
		Cols:  3,
		Width: 4000,
		Data: [][]string{
			{"A1", "B1", "C1"},
			{"A2", "B2", "C2"},
			{"A3", "B3", "C3"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test traversing column 2 (index 1)
	var visitedCells []string
	err = table.ForEachInColumn(1, func(row int, cell *TableCell, text string) error {
		visitedCells = append(visitedCells, fmt.Sprintf("%d:%s", row, text))
		return nil
	})

	if err != nil {
		t.Fatalf("ForEachInColumn execution failed: %v", err)
	}

	expectedCells := []string{"0:B1", "1:B2", "2:B3"}
	if len(visitedCells) != len(expectedCells) {
		t.Errorf("visited cell count mismatch: expected %d, got %d", len(expectedCells), len(visitedCells))
	}

	for i, expected := range expectedCells {
		if i < len(visitedCells) && visitedCells[i] != expected {
			t.Errorf("cell visit order incorrect: expected '%s', got '%s'", expected, visitedCells[i])
		}
	}

	// Test invalid column index
	err = table.ForEachInColumn(5, func(row int, cell *TableCell, text string) error {
		return nil
	})
	if err == nil {
		t.Error("should return invalid column index error")
	}
}

// TestGetCellRange tests getting a cell range
func TestGetCellRange(t *testing.T) {
	doc := New()
	config := &TableConfig{
		Rows:  4,
		Cols:  4,
		Width: 5000,
		Data: [][]string{
			{"A1", "B1", "C1", "D1"},
			{"A2", "B2", "C2", "D2"},
			{"A3", "B3", "C3", "D3"},
			{"A4", "B4", "C4", "D4"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Test getting 2x2 range (1,1) to (2,2)
	cells, err := table.GetCellRange(1, 1, 2, 2)
	if err != nil {
		t.Fatalf("GetCellRange execution failed: %v", err)
	}

	expectedCells := []struct {
		row  int
		col  int
		text string
	}{
		{1, 1, "B2"}, {1, 2, "C2"},
		{2, 1, "B3"}, {2, 2, "C3"},
	}

	if len(cells) != len(expectedCells) {
		t.Errorf("returned cell count mismatch: expected %d, got %d", len(expectedCells), len(cells))
	}

	for i, expected := range expectedCells {
		if i < len(cells) {
			cell := cells[i]
			if cell.Row != expected.row || cell.Col != expected.col {
				t.Errorf("cell position mismatch: expected (%d,%d), got (%d,%d)",
					expected.row, expected.col, cell.Row, cell.Col)
			}
			if cell.Text != expected.text {
				t.Errorf("cell text mismatch: expected '%s', got '%s'",
					expected.text, cell.Text)
			}
		}
	}

	// Test invalid range
	_, err = table.GetCellRange(2, 2, 1, 1) // start position greater than end position
	if err == nil {
		t.Error("should return invalid range error")
	}

	_, err = table.GetCellRange(0, 0, 10, 10) // exceeds table bounds
	if err == nil {
		t.Error("should return out of bounds error")
	}
}

// TestFindCells tests cell search functionality
func TestFindCells(t *testing.T) {
	doc := New()
	config := &TableConfig{
		Rows:  3,
		Cols:  3,
		Width: 4000,
		Data: [][]string{
			{"apple", "banana", "cherry"},
			{"dog", "apple", "fish"},
			{"grape", "horse", "apple"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Find cells containing "apple"
	cells, err := table.FindCells(func(row, col int, cell *TableCell, text string) bool {
		return text == "apple"
	})

	if err != nil {
		t.Fatalf("FindCells execution failed: %v", err)
	}

	expectedPositions := [][2]int{{0, 0}, {1, 1}, {2, 2}}
	if len(cells) != len(expectedPositions) {
		t.Errorf("found cell count mismatch: expected %d, got %d", len(expectedPositions), len(cells))
	}

	for i, expected := range expectedPositions {
		if i < len(cells) {
			cell := cells[i]
			if cell.Row != expected[0] || cell.Col != expected[1] {
				t.Errorf("found cell position incorrect: expected (%d,%d), got (%d,%d)",
					expected[0], expected[1], cell.Row, cell.Col)
			}
			if cell.Text != "apple" {
				t.Errorf("found cell text incorrect: expected 'apple', got '%s'", cell.Text)
			}
		}
	}
}

// TestFindCellsByText tests finding cells by text
func TestFindCellsByText(t *testing.T) {
	doc := New()
	config := &TableConfig{
		Rows:  2,
		Cols:  3,
		Width: 4000,
		Data: [][]string{
			{"test", "testing", "other"},
			{"notest", "test123", "test"},
		},
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}

	// Exact match
	cells, err := table.FindCellsByText("test", true)
	if err != nil {
		t.Fatalf("FindCellsByText execution failed: %v", err)
	}

	expectedCount := 2 // (0,0) and (1,2)
	if len(cells) != expectedCount {
		t.Errorf("exact match found cell count mismatch: expected %d, got %d", expectedCount, len(cells))
	}

	// Fuzzy match
	cells, err = table.FindCellsByText("test", false)
	if err != nil {
		t.Fatalf("FindCellsByText execution failed: %v", err)
	}

	expectedCount = 5 // all cells containing "test"
	if len(cells) != expectedCount {
		t.Errorf("fuzzy match found cell count mismatch: expected %d, got %d", expectedCount, len(cells))
	}
}

// TestEmptyTable tests iterator on an empty table
func TestEmptyTable(t *testing.T) {
	doc := New()

	// Create an empty table (minimum is 1x1)
	config := &TableConfig{
		Rows:  1,
		Cols:  1,
		Width: 2000,
	}

	table, err := doc.CreateTable(config)
	if err != nil {
		t.Fatalf("failed to create table: %v", err)
	}
	iterator := table.NewCellIterator()

	if iterator.Total() != 1 {
		t.Errorf("empty table total cell count should be 1, got %d", iterator.Total())
	}

	if !iterator.HasNext() {
		t.Error("1x1 table should have one cell")
	}
}
