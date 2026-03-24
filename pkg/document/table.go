// Package document provides table operations for Word documents.
package document

import (
	"encoding/xml"
	"fmt"
	"strings"
)

// Table represents a table.
type Table struct {
	XMLName    xml.Name         `xml:"w:tbl"`
	Properties *TableProperties `xml:"w:tblPr,omitempty"`
	Grid       *TableGrid       `xml:"w:tblGrid,omitempty"`
	Rows       []TableRow       `xml:"w:tr"`
}

// TableProperties represents table properties.
type TableProperties struct {
	XMLName      xml.Name          `xml:"w:tblPr"`
	TableW       *TableWidth       `xml:"w:tblW,omitempty"`
	TableJc      *TableJc          `xml:"w:jc,omitempty"`
	TableLook    *TableLook        `xml:"w:tblLook,omitempty"`
	TableStyle   *TableStyle       `xml:"w:tblStyle,omitempty"`   // table style
	TableBorders *TableBorders     `xml:"w:tblBorders,omitempty"` // table borders
	Shd          *TableShading     `xml:"w:shd,omitempty"`        // table shading/background
	TableCellMar *TableCellMargins `xml:"w:tblCellMar,omitempty"` // table cell margins
	TableLayout  *TableLayoutType  `xml:"w:tblLayout,omitempty"`  // table layout type
	TableInd     *TableIndentation `xml:"w:tblInd,omitempty"`     // table indentation
}

// TableWidth represents the table width.
type TableWidth struct {
	XMLName xml.Name `xml:"w:tblW"`
	W       string   `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// TableJc represents the table alignment.
type TableJc struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr"`
}

// TableLook represents the table appearance.
type TableLook struct {
	XMLName  xml.Name `xml:"w:tblLook"`
	Val      string   `xml:"w:val,attr"`
	FirstRow string   `xml:"w:firstRow,attr,omitempty"`
	LastRow  string   `xml:"w:lastRow,attr,omitempty"`
	FirstCol string   `xml:"w:firstColumn,attr,omitempty"`
	LastCol  string   `xml:"w:lastColumn,attr,omitempty"`
	NoHBand  string   `xml:"w:noHBand,attr,omitempty"`
	NoVBand  string   `xml:"w:noVBand,attr,omitempty"`
}

// TableGrid represents the table grid.
type TableGrid struct {
	XMLName xml.Name       `xml:"w:tblGrid"`
	Cols    []TableGridCol `xml:"w:gridCol"`
}

// TableGridCol represents a table grid column.
type TableGridCol struct {
	XMLName xml.Name `xml:"w:gridCol"`
	W       string   `xml:"w:w,attr,omitempty"`
}

// TableRow represents a table row.
type TableRow struct {
	XMLName    xml.Name            `xml:"w:tr"`
	Properties *TableRowProperties `xml:"w:trPr,omitempty"`
	Cells      []TableCell         `xml:"w:tc"`
}

// TableRowProperties represents table row properties.
type TableRowProperties struct {
	XMLName   xml.Name   `xml:"w:trPr"`
	TableRowH *TableRowH `xml:"w:trHeight,omitempty"`
	CantSplit *CantSplit `xml:"w:cantSplit,omitempty"` // prevent page break within row
	TblHeader *TblHeader `xml:"w:tblHeader,omitempty"` // repeat as header row
}

// TableRowH represents the table row height.
type TableRowH struct {
	XMLName xml.Name `xml:"w:trHeight"`
	Val     string   `xml:"w:val,attr,omitempty"`
	HRule   string   `xml:"w:hRule,attr,omitempty"`
}

// TableCell represents a table cell.
type TableCell struct {
	XMLName    xml.Name             `xml:"w:tc"`
	Properties *TableCellProperties `xml:"w:tcPr,omitempty"`
	Paragraphs []Paragraph          `xml:"w:p"`
	Tables     []Table              `xml:"w:tbl"` // nested tables
}

// MarshalXML implements custom XML serialization to ensure nested tables are serialized correctly.
// OOXML requires cell content to output paragraphs and tables in original document order.
func (tc *TableCell) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// start element <w:tc>
	start.Name = xml.Name{Local: "w:tc"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// serialize properties <w:tcPr>
	if tc.Properties != nil {
		if err := e.Encode(tc.Properties); err != nil {
			return err
		}
	}

	// serialize paragraphs <w:p>
	for i := range tc.Paragraphs {
		if err := e.Encode(&tc.Paragraphs[i]); err != nil {
			return err
		}
	}

	// serialize nested tables <w:tbl>
	for i := range tc.Tables {
		if err := e.Encode(&tc.Tables[i]); err != nil {
			return err
		}
	}

	// end element </w:tc>
	return e.EncodeToken(start.End())
}

// TableCellProperties represents table cell properties.
type TableCellProperties struct {
	XMLName       xml.Name              `xml:"w:tcPr"`
	TableCellW    *TableCellW           `xml:"w:tcW,omitempty"`
	VAlign        *VAlign               `xml:"w:vAlign,omitempty"`
	GridSpan      *GridSpan             `xml:"w:gridSpan,omitempty"`
	VMerge        *VMerge               `xml:"w:vMerge,omitempty"`
	TextDirection *TextDirection        `xml:"w:textDirection,omitempty"`
	Shd           *TableCellShading     `xml:"w:shd,omitempty"`       // cell background
	TcBorders     *TableCellBorders     `xml:"w:tcBorders,omitempty"` // cell borders
	TcMar         *TableCellMarginsCell `xml:"w:tcMar,omitempty"`     // cell margins
	NoWrap        *NoWrap               `xml:"w:noWrap,omitempty"`    // no wrap
	HideMark      *HideMark             `xml:"w:hideMark,omitempty"`  // hide mark
}

// TableCellMarginsCell represents cell margins (different XML structure from table margins).
type TableCellMarginsCell struct {
	XMLName xml.Name            `xml:"w:tcMar"`
	Top     *TableCellSpaceCell `xml:"w:top,omitempty"`
	Left    *TableCellSpaceCell `xml:"w:left,omitempty"`
	Bottom  *TableCellSpaceCell `xml:"w:bottom,omitempty"`
	Right   *TableCellSpaceCell `xml:"w:right,omitempty"`
}

// TableCellSpaceCell represents cell spacing settings.
type TableCellSpaceCell struct {
	W    string `xml:"w:w,attr"`
	Type string `xml:"w:type,attr"`
}

// TableCellW represents the cell width.
type TableCellW struct {
	XMLName xml.Name `xml:"w:tcW"`
	W       string   `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// VAlign represents vertical alignment.
type VAlign struct {
	XMLName xml.Name `xml:"w:vAlign"`
	Val     string   `xml:"w:val,attr"`
}

// GridSpan represents the grid span (column merge).
type GridSpan struct {
	XMLName xml.Name `xml:"w:gridSpan"`
	Val     string   `xml:"w:val,attr"`
}

// VMerge represents a vertical merge (row merge).
type VMerge struct {
	XMLName xml.Name `xml:"w:vMerge"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// TableConfig represents the table configuration.
type TableConfig struct {
	Rows      int        // number of rows
	Cols      int        // number of columns
	Width     int        // total table width (in twips)
	ColWidths []int      // column widths (in twips); evenly distributed if empty
	Data      [][]string // initial data
	Emphases  [][]int    // cell styles: 1=italic, 2=bold
}

// CreateTable creates a new table.
// Parameters:
//   - config: table configuration
//
// Returns:
//   - *Table: the created table object
//   - error: an error if the configuration is invalid
func (d *Document) CreateTable(config *TableConfig) (*Table, error) {
	if config.Rows <= 0 || config.Cols <= 0 {
		Error("table rows and columns must be greater than 0")
		return nil, NewValidationError("TableConfig", "", "table rows and columns must be greater than 0")
	}

	table := &Table{
		Properties: &TableProperties{
			TableW: &TableWidth{
				W:    fmt.Sprintf("%d", config.Width),
				Type: "dxa", // unit in twips
			},
			TableJc: &TableJc{
				Val: "center", // default center alignment
			},
			TableLook: &TableLook{
				Val:      "04A0",
				FirstRow: "1",
				LastRow:  "0",
				FirstCol: "1",
				LastCol:  "0",
				NoHBand:  "0",
				NoVBand:  "1",
			},
			// add default table borders using single-line border style matching the tmp_test reference table
			TableBorders: &TableBorders{
				Top: &TableBorder{
					Val:   "single", // single line border style
					Sz:    "4",      // border thickness (1/8 point)
					Space: "0",      // border spacing
					Color: "auto",   // automatic color
				},
				Left: &TableBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				Bottom: &TableBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				Right: &TableBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				InsideH: &TableBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
				InsideV: &TableBorder{
					Val:   "single",
					Sz:    "4",
					Space: "0",
					Color: "auto",
				},
			},
			// add table layout and cell margin settings consistent with the reference table
			TableLayout: &TableLayoutType{
				Type: "autofit", // auto-fit layout
			},
			TableCellMar: &TableCellMargins{
				Left: &TableCellSpace{
					W:    "108", // left margin (matching reference table)
					Type: "dxa",
				},
				Right: &TableCellSpace{
					W:    "108", // right margin (matching reference table)
					Type: "dxa",
				},
			},
		},
		Grid: &TableGrid{},
		Rows: make([]TableRow, 0, config.Rows),
	}

	// set column widths
	colWidths := config.ColWidths
	if len(colWidths) == 0 {
		// distribute width evenly
		avgWidth := config.Width / config.Cols
		colWidths = make([]int, config.Cols)
		for i := range colWidths {
			colWidths[i] = avgWidth
		}
	} else if len(colWidths) != config.Cols {
		Error("column width count does not match column count")
		return nil, NewValidationError("TableConfig.ColWidths", "", "column width count does not match column count")
	}

	// create table grid
	for _, width := range colWidths {
		table.Grid.Cols = append(table.Grid.Cols, TableGridCol{
			W: fmt.Sprintf("%d", width),
		})
	}

	// create table rows and cells
	for i := 0; i < config.Rows; i++ {
		row := TableRow{
			Cells: make([]TableCell, 0, config.Cols),
		}

		for j := 0; j < config.Cols; j++ {
			cell := TableCell{
				Properties: &TableCellProperties{
					TableCellW: &TableCellW{
						W:    fmt.Sprintf("%d", colWidths[j]),
						Type: "dxa",
					},
					VAlign: &VAlign{
						Val: "center",
					},
				},
				Paragraphs: []Paragraph{
					{
						Runs: []Run{
							{
								Text: Text{
									Content: "",
								},
							},
						},
					},
				},
			}

			// set cell content if initial data is provided
			if config.Data != nil && i < len(config.Data) && j < len(config.Data[i]) {
				cell.Paragraphs[0].Runs[0].Text.Content = config.Data[i][j]
			}

			if config.Emphases != nil && i < len(config.Emphases) && j < len(config.Emphases[i]) {
				switch config.Emphases[i][j] {
				case 1:
					cell.Paragraphs[0].Runs[0].Properties = &RunProperties{Italic: &Italic{}}
				case 2:
					cell.Paragraphs[0].Runs[0].Properties = &RunProperties{Bold: &Bold{}}
				}
			}

			row.Cells = append(row.Cells, cell)
		}

		table.Rows = append(table.Rows, row)
	}

	InfoMsgf(MsgTableCreated, config.Rows, config.Cols)
	return table, nil
}

// AddTable adds a table to the document.
// Parameters:
//   - config: table configuration
//
// Returns:
//   - *Table: the added table object
//   - error: an error if the configuration is invalid
func (d *Document) AddTable(config *TableConfig) (*Table, error) {
	table, err := d.CreateTable(config)
	if err != nil {
		return nil, err
	}

	// add the table to the document body
	d.Body.Elements = append(d.Body.Elements, table)

	InfoMsgf(MsgTableAddedToDocument, len(d.Body.GetTables()))
	return table, nil
}

// InsertRow inserts a row at the specified position.
func (t *Table) InsertRow(position int, data []string) error {
	if position < 0 || position > len(t.Rows) {
		return fmt.Errorf("invalid insert position: %d, table has %d rows", position, len(t.Rows))
	}

	if len(t.Rows) == 0 {
		return fmt.Errorf("table has no column definitions, cannot insert row")
	}

	colCount := len(t.Rows[0].Cells)
	if len(data) > colCount {
		return fmt.Errorf("data column count (%d) exceeds table column count (%d)", len(data), colCount)
	}

	// create new row
	newRow := TableRow{
		Cells: make([]TableCell, colCount),
	}

	// copy first row cell properties as template
	templateRow := t.Rows[0]
	for i := 0; i < colCount; i++ {
		// deep copy cell properties
		var cellProps *TableCellProperties
		if templateRow.Cells[i].Properties != nil {
			cellProps = &TableCellProperties{}
			// copy width
			if templateRow.Cells[i].Properties.TableCellW != nil {
				cellProps.TableCellW = &TableCellW{
					W:    templateRow.Cells[i].Properties.TableCellW.W,
					Type: templateRow.Cells[i].Properties.TableCellW.Type,
				}
			}
			// copy vertical alignment
			if templateRow.Cells[i].Properties.VAlign != nil {
				cellProps.VAlign = &VAlign{
					Val: templateRow.Cells[i].Properties.VAlign.Val,
				}
			}
			// copy other necessary properties
			// note: do not copy GridSpan and VMerge as these are merge-related properties
		}

		newRow.Cells[i] = TableCell{
			Properties: cellProps,
			Paragraphs: []Paragraph{
				{
					Runs: []Run{
						{
							Text: Text{
								Content: "",
							},
						},
					},
				},
			},
		}

		// set data
		if i < len(data) {
			newRow.Cells[i].Paragraphs[0].Runs[0].Text.Content = data[i]
		}
	}

	// insert row
	if position == len(t.Rows) {
		// append at end
		t.Rows = append(t.Rows, newRow)
	} else {
		// insert in the middle
		t.Rows = append(t.Rows[:position+1], t.Rows[position:]...)
		t.Rows[position] = newRow
	}

	InfoMsgf(MsgRowInserted, position)
	return nil
}

// AppendRow appends a row at the end of the table.
func (t *Table) AppendRow(data []string) error {
	return t.InsertRow(len(t.Rows), data)
}

// DeleteRow deletes the specified row.
func (t *Table) DeleteRow(rowIndex int) error {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return fmt.Errorf("invalid row index: %d, table has %d rows", rowIndex, len(t.Rows))
	}

	if len(t.Rows) <= 1 {
		return fmt.Errorf("table must have at least one row")
	}

	// delete row
	t.Rows = append(t.Rows[:rowIndex], t.Rows[rowIndex+1:]...)

	InfoMsgf(MsgRowDeleted, rowIndex)
	return nil
}

// DeleteRows deletes the specified range of rows.
func (t *Table) DeleteRows(startIndex, endIndex int) error {
	if startIndex < 0 || endIndex >= len(t.Rows) || startIndex > endIndex {
		return fmt.Errorf("invalid row index range: [%d, %d], table has %d rows", startIndex, endIndex, len(t.Rows))
	}

	deleteCount := endIndex - startIndex + 1
	if len(t.Rows)-deleteCount < 1 {
		return fmt.Errorf("table must have at least one row after deletion")
	}

	// delete row range
	t.Rows = append(t.Rows[:startIndex], t.Rows[endIndex+1:]...)

	InfoMsgf(MsgRowsDeleted, startIndex, endIndex)
	return nil
}

// InsertColumn inserts a column at the specified position.
func (t *Table) InsertColumn(position int, data []string, width int) error {
	if len(t.Rows) == 0 {
		return fmt.Errorf("table has no rows, cannot insert column")
	}

	colCount := len(t.Rows[0].Cells)
	if position < 0 || position > colCount {
		return fmt.Errorf("invalid insert position: %d, table has %d columns", position, colCount)
	}

	if len(data) > len(t.Rows) {
		return fmt.Errorf("data row count (%d) exceeds table row count (%d)", len(data), len(t.Rows))
	}

	// update table grid
	newGridCol := TableGridCol{
		W: fmt.Sprintf("%d", width),
	}
	if position == len(t.Grid.Cols) {
		t.Grid.Cols = append(t.Grid.Cols, newGridCol)
	} else {
		t.Grid.Cols = append(t.Grid.Cols[:position+1], t.Grid.Cols[position:]...)
		t.Grid.Cols[position] = newGridCol
	}

	// insert new cell for each row
	for i := range t.Rows {
		newCell := TableCell{
			Properties: &TableCellProperties{
				TableCellW: &TableCellW{
					W:    fmt.Sprintf("%d", width),
					Type: "dxa",
				},
				VAlign: &VAlign{
					Val: "center",
				},
			},
			Paragraphs: []Paragraph{
				{
					Runs: []Run{
						{
							Text: Text{
								Content: "",
							},
						},
					},
				},
			},
		}

		// set data
		if i < len(data) {
			newCell.Paragraphs[0].Runs[0].Text.Content = data[i]
		}

		// insert cell
		if position == len(t.Rows[i].Cells) {
			t.Rows[i].Cells = append(t.Rows[i].Cells, newCell)
		} else {
			t.Rows[i].Cells = append(t.Rows[i].Cells[:position+1], t.Rows[i].Cells[position:]...)
			t.Rows[i].Cells[position] = newCell
		}
	}

	InfoMsgf(MsgColumnInserted, position)
	return nil
}

// AppendColumn appends a column at the end of the table.
func (t *Table) AppendColumn(data []string, width int) error {
	colCount := 0
	if len(t.Rows) > 0 {
		colCount = len(t.Rows[0].Cells)
	}
	return t.InsertColumn(colCount, data, width)
}

// DeleteColumn deletes the specified column.
func (t *Table) DeleteColumn(colIndex int) error {
	if len(t.Rows) == 0 {
		return fmt.Errorf("table has no rows")
	}

	colCount := len(t.Rows[0].Cells)
	if colIndex < 0 || colIndex >= colCount {
		return fmt.Errorf("invalid column index: %d, table has %d columns", colIndex, colCount)
	}

	if colCount <= 1 {
		return fmt.Errorf("table must have at least one column")
	}

	// delete grid column
	t.Grid.Cols = append(t.Grid.Cols[:colIndex], t.Grid.Cols[colIndex+1:]...)

	// delete corresponding cell from each row
	for i := range t.Rows {
		t.Rows[i].Cells = append(t.Rows[i].Cells[:colIndex], t.Rows[i].Cells[colIndex+1:]...)
	}

	InfoMsgf(MsgColumnDeleted, colIndex)
	return nil
}

// DeleteColumns deletes the specified range of columns.
func (t *Table) DeleteColumns(startIndex, endIndex int) error {
	if len(t.Rows) == 0 {
		return fmt.Errorf("table has no rows")
	}

	colCount := len(t.Rows[0].Cells)
	if startIndex < 0 || endIndex >= colCount || startIndex > endIndex {
		return fmt.Errorf("invalid column index range: [%d, %d], table has %d columns", startIndex, endIndex, colCount)
	}

	deleteCount := endIndex - startIndex + 1
	if colCount-deleteCount < 1 {
		return fmt.Errorf("table must have at least one column after deletion")
	}

	// delete grid column range
	t.Grid.Cols = append(t.Grid.Cols[:startIndex], t.Grid.Cols[endIndex+1:]...)

	// delete corresponding cell range from each row
	for i := range t.Rows {
		t.Rows[i].Cells = append(t.Rows[i].Cells[:startIndex], t.Rows[i].Cells[endIndex+1:]...)
	}

	InfoMsgf(MsgColumnsDeleted, startIndex, endIndex)
	return nil
}

// GetCell returns the cell at the specified position.
func (t *Table) GetCell(row, col int) (*TableCell, error) {
	if row < 0 || row >= len(t.Rows) {
		return nil, fmt.Errorf("invalid row index: %d, table has %d rows", row, len(t.Rows))
	}

	if col < 0 || col >= len(t.Rows[row].Cells) {
		return nil, fmt.Errorf("invalid column index: %d, row %d has %d columns", col, row, len(t.Rows[row].Cells))
	}

	return &t.Rows[row].Cells[col], nil
}

// SetCellText sets the cell text content.
func (t *Table) SetCellText(row, col int, text string) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// ensure cell has paragraphs and runs
	if len(cell.Paragraphs) == 0 {
		cell.Paragraphs = []Paragraph{
			{
				Runs: []Run{
					{
						Text: Text{Content: text},
					},
				},
			},
		}
	} else {
		if len(cell.Paragraphs[0].Runs) == 0 {
			cell.Paragraphs[0].Runs = []Run{
				{
					Text: Text{Content: text},
				},
			}
		} else {
			cell.Paragraphs[0].Runs[0].Text.Content = text
		}
	}

	return nil
}

// GetCellText returns the cell text content.
func (t *Table) GetCellText(row, col int) (string, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return "", err
	}

	if len(cell.Paragraphs) == 0 {
		return "", nil
	}

	var result string
	for idx, para := range cell.Paragraphs {
		for _, run := range para.Runs {
			result += run.Text.Content
		}
		// add soft line break between paragraphs (except the last)
		if idx < len(cell.Paragraphs)-1 {
			result += "\n"
		}
	}
	return result, nil
}

// GetRowCount returns the number of rows in the table.
func (t *Table) GetRowCount() int {
	return len(t.Rows)
}

// GetColumnCount returns the number of columns in the table.
func (t *Table) GetColumnCount() int {
	if len(t.Rows) == 0 {
		return 0
	}
	return len(t.Rows[0].Cells)
}

// ClearTable clears the table content while preserving structure.
func (t *Table) ClearTable() {
	for i := range t.Rows {
		for j := range t.Rows[i].Cells {
			t.Rows[i].Cells[j].Paragraphs = []Paragraph{
				{
					Runs: []Run{
						{
							Text: Text{Content: ""},
						},
					},
				},
			}
		}
	}
	InfoMsg(MsgTableContentCleared)
}

// CopyTable creates a copy of the table.
func (t *Table) CopyTable() *Table {
	// deep copy table structure
	newTable := &Table{
		Properties: t.Properties,
		Grid:       t.Grid,
		Rows:       make([]TableRow, len(t.Rows)),
	}

	// copy all rows and cells
	for i, row := range t.Rows {
		newTable.Rows[i] = TableRow{
			Properties: row.Properties,
			Cells:      make([]TableCell, len(row.Cells)),
		}

		for j, cell := range row.Cells {
			newTable.Rows[i].Cells[j] = TableCell{
				Properties: cell.Properties,
				Paragraphs: make([]Paragraph, len(cell.Paragraphs)),
			}

			// copy paragraph content
			for k, para := range cell.Paragraphs {
				newTable.Rows[i].Cells[j].Paragraphs[k] = Paragraph{
					Properties: para.Properties,
					Runs:       make([]Run, len(para.Runs)),
				}

				for l, run := range para.Runs {
					newTable.Rows[i].Cells[j].Paragraphs[k].Runs[l] = Run{
						Properties: run.Properties,
						Text:       Text{Content: run.Text.Content},
					}
				}
			}
		}
	}

	InfoMsg(MsgTableCopied)
	return newTable
}

// CellAlignment represents the cell alignment type.
type CellAlignment string

const (
	// CellAlignLeft represents left alignment.
	CellAlignLeft CellAlignment = "left"
	// CellAlignCenter represents center alignment.
	CellAlignCenter CellAlignment = "center"
	// CellAlignRight represents right alignment.
	CellAlignRight CellAlignment = "right"
	// CellAlignJustify represents justified alignment.
	CellAlignJustify CellAlignment = "both"
)

// CellVerticalAlignment represents the cell vertical alignment type.
type CellVerticalAlignment string

const (
	// CellVAlignTop represents top alignment.
	CellVAlignTop CellVerticalAlignment = "top"
	// CellVAlignCenter represents center alignment.
	CellVAlignCenter CellVerticalAlignment = "center"
	// CellVAlignBottom represents bottom alignment.
	CellVAlignBottom CellVerticalAlignment = "bottom"
)

// CellTextDirection represents the cell text direction.
type CellTextDirection string

const (
	// TextDirectionLR represents left-to-right direction (default).
	TextDirectionLR CellTextDirection = "lrTb"
	// TextDirectionTB represents top-to-bottom direction.
	TextDirectionTB CellTextDirection = "tbRl"
	// TextDirectionBT represents bottom-to-top direction.
	TextDirectionBT CellTextDirection = "btLr"
	// TextDirectionRL represents right-to-left direction.
	TextDirectionRL CellTextDirection = "rlTb"
	// TextDirectionTBV represents top-to-bottom vertical display.
	TextDirectionTBV CellTextDirection = "tbLrV"
	// TextDirectionBTV represents bottom-to-top vertical display.
	TextDirectionBTV CellTextDirection = "btLrV"
)

// CellFormat represents the cell format configuration.
type CellFormat struct {
	TextFormat      *TextFormat           // text format
	HorizontalAlign CellAlignment         // horizontal alignment
	VerticalAlign   CellVerticalAlignment // vertical alignment
	TextDirection   CellTextDirection     // text direction
	BackgroundColor string                // background color
	BorderStyle     string                // border style
	Padding         int                   // padding (in twips)
}

// SetCellFormat sets the cell format.
func (t *Table) SetCellFormat(row, col int, format *CellFormat) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// ensure cell has properties
	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	// set vertical alignment
	if format.VerticalAlign != "" {
		cell.Properties.VAlign = &VAlign{
			Val: string(format.VerticalAlign),
		}
	}

	// set text direction
	if format.TextDirection != "" {
		cell.Properties.TextDirection = &TextDirection{
			Val: string(format.TextDirection),
		}
	}

	// ensure cell has paragraphs
	if len(cell.Paragraphs) == 0 {
		cell.Paragraphs = []Paragraph{{}}
	}

	// set horizontal alignment
	if format.HorizontalAlign != "" {
		if cell.Paragraphs[0].Properties == nil {
			cell.Paragraphs[0].Properties = &ParagraphProperties{}
		}
		cell.Paragraphs[0].Properties.Justification = &Justification{
			Val: string(format.HorizontalAlign),
		}
	}

	// set text format
	if format.TextFormat != nil {
		// ensure runs exist
		if len(cell.Paragraphs[0].Runs) == 0 {
			cell.Paragraphs[0].Runs = []Run{{}}
		}

		run := &cell.Paragraphs[0].Runs[0]
		if run.Properties == nil {
			run.Properties = &RunProperties{}
		}

		// set bold
		if format.TextFormat.Bold {
			run.Properties.Bold = &Bold{}
		}

		// set italic
		if format.TextFormat.Italic {
			run.Properties.Italic = &Italic{}
		}

		// set font size
		if format.TextFormat.FontSize > 0 {
			run.Properties.FontSize = &FontSize{
				Val: fmt.Sprintf("%d", format.TextFormat.FontSize*2), // Word uses half-points
			}
		}

		// set font color
		if format.TextFormat.FontColor != "" {
			run.Properties.Color = &Color{
				Val: format.TextFormat.FontColor,
			}
		}

		// set font family
		if format.TextFormat.FontFamily != "" {
			run.Properties.FontFamily = &FontFamily{
				ASCII: format.TextFormat.FontFamily,
			}
		}
	}

	InfoMsgf(MsgCellFormatSet, row, col)
	return nil
}

// SetCellFormattedText sets the cell rich text content.
func (t *Table) SetCellFormattedText(row, col int, text string, format *TextFormat) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// create formatted run
	run := Run{
		Text: Text{Content: text},
	}

	if format != nil {
		run.Properties = &RunProperties{}

		if format.FontFamily != "" {
			run.Properties.FontFamily = &FontFamily{
				ASCII: format.FontFamily,
			}
		}

		if format.Bold {
			run.Properties.Bold = &Bold{}
		}

		if format.Italic {
			run.Properties.Italic = &Italic{}
		}

		if format.FontColor != "" {
			run.Properties.Color = &Color{
				Val: format.FontColor,
			}
		}

		if format.FontSize > 0 {
			run.Properties.FontSize = &FontSize{
				Val: fmt.Sprintf("%d", format.FontSize*2),
			}
		}
	}

	// set cell content
	cell.Paragraphs = []Paragraph{
		{
			Runs: []Run{run},
		},
	}

	InfoMsgf(MsgCellRichTextSet, row, col)
	return nil
}

// AddCellFormattedText appends formatted text to a cell.
func (t *Table) AddCellFormattedText(row, col int, text string, format *TextFormat) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// ensure cell has paragraphs
	if len(cell.Paragraphs) == 0 {
		cell.Paragraphs = []Paragraph{{}}
	}

	// create formatted run
	run := Run{
		Text: Text{Content: text},
	}

	if format != nil {
		run.Properties = &RunProperties{}

		if format.FontFamily != "" {
			run.Properties.FontFamily = &FontFamily{
				ASCII: format.FontFamily,
			}
		}

		if format.Bold {
			run.Properties.Bold = &Bold{}
		}

		if format.Italic {
			run.Properties.Italic = &Italic{}
		}

		if format.FontColor != "" {
			run.Properties.Color = &Color{
				Val: format.FontColor,
			}
		}

		if format.FontSize > 0 {
			run.Properties.FontSize = &FontSize{
				Val: fmt.Sprintf("%d", format.FontSize*2),
			}
		}
	}

	// append run to first paragraph
	cell.Paragraphs[0].Runs = append(cell.Paragraphs[0].Runs, run)

	InfoMsgf(MsgFormattedTextAddedToCell, row, col)
	return nil
}

// MergeCellsHorizontal merges cells horizontally (column merge).
func (t *Table) MergeCellsHorizontal(row, startCol, endCol int) error {
	if row < 0 || row >= len(t.Rows) {
		return fmt.Errorf("invalid row index: %d", row)
	}

	if startCol < 0 || endCol >= len(t.Rows[row].Cells) || startCol > endCol {
		return fmt.Errorf("invalid column index range: [%d, %d]", startCol, endCol)
	}

	if startCol == endCol {
		return fmt.Errorf("start column and end column cannot be the same")
	}

	// set grid span on the starting cell
	startCell := &t.Rows[row].Cells[startCol]
	if startCell.Properties == nil {
		startCell.Properties = &TableCellProperties{}
	}

	spanCount := endCol - startCol + 1
	startCell.Properties.GridSpan = &GridSpan{
		Val: fmt.Sprintf("%d", spanCount),
	}

	// remove merged cells
	newCells := make([]TableCell, 0, len(t.Rows[row].Cells)-(endCol-startCol))
	newCells = append(newCells, t.Rows[row].Cells[:startCol+1]...)
	if endCol+1 < len(t.Rows[row].Cells) {
		newCells = append(newCells, t.Rows[row].Cells[endCol+1:]...)
	}
	t.Rows[row].Cells = newCells

	InfoMsgf(MsgHorizontalMerge, row, startCol, endCol)
	return nil
}

// MergeCellsVertical merges cells vertically (row merge).
func (t *Table) MergeCellsVertical(startRow, endRow, col int) error {
	if startRow < 0 || endRow >= len(t.Rows) || startRow > endRow {
		return fmt.Errorf("invalid row index range: [%d, %d]", startRow, endRow)
	}

	if col < 0 {
		return fmt.Errorf("invalid column index: %d", col)
	}

	if startRow == endRow {
		return fmt.Errorf("start row and end row cannot be the same")
	}

	// check column count for all rows
	for i := startRow; i <= endRow; i++ {
		if col >= len(t.Rows[i].Cells) {
			return fmt.Errorf("row %d does not have column %d", i, col)
		}
	}

	// set starting cell as merge start
	startCell := &t.Rows[startRow].Cells[col]
	if startCell.Properties == nil {
		startCell.Properties = &TableCellProperties{}
	}
	startCell.Properties.VMerge = &VMerge{
		Val: "restart",
	}

	// set subsequent cells as merge continue
	for i := startRow + 1; i <= endRow; i++ {
		cell := &t.Rows[i].Cells[col]
		if cell.Properties == nil {
			cell.Properties = &TableCellProperties{}
		}
		cell.Properties.VMerge = &VMerge{
			Val: "continue",
		}
		// clear content of merged cells
		cell.Paragraphs = []Paragraph{{}}
	}

	InfoMsgf(MsgVerticalMerge, startRow, endRow, col)
	return nil
}

// MergeCellsRange merges a range of cells (multiple rows and columns).
func (t *Table) MergeCellsRange(startRow, endRow, startCol, endCol int) error {
	// validate range
	if startRow < 0 || endRow >= len(t.Rows) || startRow > endRow {
		return fmt.Errorf("invalid row index range: [%d, %d]", startRow, endRow)
	}

	// first merge each row horizontally
	for i := startRow; i <= endRow; i++ {
		if startCol >= len(t.Rows[i].Cells) || endCol >= len(t.Rows[i].Cells) {
			return fmt.Errorf("invalid column index range for row %d: [%d, %d]", i, startCol, endCol)
		}

		if startCol != endCol {
			err := t.MergeCellsHorizontal(i, startCol, endCol)
			if err != nil {
				return fmt.Errorf("failed to merge row %d horizontally: %v", i, err)
			}
		}
	}

	// then merge the first column vertically
	if startRow != endRow {
		err := t.MergeCellsVertical(startRow, endRow, startCol)
		if err != nil {
			return fmt.Errorf("vertical merge failed: %v", err)
		}
	}

	InfoMsgf(MsgRangeMerge, startRow, endRow, startCol, endCol)
	return nil
}

// UnmergeCells cancels cell merge.
func (t *Table) UnmergeCells(row, col int) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	if cell.Properties == nil {
		return fmt.Errorf("cell is not merged")
	}

	// check for horizontal merge
	if cell.Properties.GridSpan != nil {
		// get the number of merged columns
		spanCount := 1
		if cell.Properties.GridSpan.Val != "" {
			fmt.Sscanf(cell.Properties.GridSpan.Val, "%d", &spanCount)
		}

		// insert the previously merged cells
		for i := 1; i < spanCount; i++ {
			newCell := TableCell{
				Properties: &TableCellProperties{
					TableCellW: cell.Properties.TableCellW,
					VAlign:     cell.Properties.VAlign,
				},
				Paragraphs: []Paragraph{{}},
			}

			// insert new cell at specified position
			insertPos := col + i
			if insertPos <= len(t.Rows[row].Cells) {
				t.Rows[row].Cells = append(t.Rows[row].Cells[:insertPos], append([]TableCell{newCell}, t.Rows[row].Cells[insertPos:]...)...)
			}
		}

		// remove horizontal merge attribute
		cell.Properties.GridSpan = nil
	}

	// check for vertical merge
	if cell.Properties.VMerge != nil {
		// remove vertical merge attribute
		cell.Properties.VMerge = nil

		// find and restore merged cells
		for i := row + 1; i < len(t.Rows); i++ {
			if col < len(t.Rows[i].Cells) {
				otherCell := &t.Rows[i].Cells[col]
				if otherCell.Properties != nil && otherCell.Properties.VMerge != nil {
					if otherCell.Properties.VMerge.Val == "continue" {
						// restore cell content
						otherCell.Properties.VMerge = nil
						if len(otherCell.Paragraphs) == 0 {
							otherCell.Paragraphs = []Paragraph{{}}
						}
					} else {
						break
					}
				} else {
					break
				}
			}
		}
	}

	InfoMsgf(MsgCellMergeCancelled, row, col)
	return nil
}

// IsCellMerged checks whether a cell is merged.
func (t *Table) IsCellMerged(row, col int) (bool, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return false, err
	}

	if cell.Properties == nil {
		return false, nil
	}

	// check for horizontal merge
	if cell.Properties.GridSpan != nil && cell.Properties.GridSpan.Val != "" && cell.Properties.GridSpan.Val != "1" {
		return true, nil
	}

	// check for vertical merge
	if cell.Properties.VMerge != nil {
		return true, nil
	}

	return false, nil
}

// GetMergedCellInfo returns information about merged cells.
func (t *Table) GetMergedCellInfo(row, col int) (map[string]interface{}, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	info := make(map[string]interface{})
	info["is_merged"] = false

	if cell.Properties == nil {
		return info, nil
	}

	// check for horizontal merge
	if cell.Properties.GridSpan != nil && cell.Properties.GridSpan.Val != "" {
		spanCount := 1
		fmt.Sscanf(cell.Properties.GridSpan.Val, "%d", &spanCount)
		if spanCount > 1 {
			info["is_merged"] = true
			info["horizontal_span"] = spanCount
			info["merge_type"] = "horizontal"
		}
	}

	// check for vertical merge
	if cell.Properties.VMerge != nil {
		info["is_merged"] = true
		if cell.Properties.VMerge.Val == "restart" {
			info["vertical_merge_start"] = true
			info["merge_type"] = "vertical"
		} else if cell.Properties.VMerge.Val == "continue" {
			info["vertical_merge_continue"] = true
			info["merge_type"] = "vertical"
		}
	}

	return info, nil
}

// ClearCellContent clears cell content while preserving format.
func (t *Table) ClearCellContent(row, col int) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// preserve format, only clear text content
	for i := range cell.Paragraphs {
		for j := range cell.Paragraphs[i].Runs {
			cell.Paragraphs[i].Runs[j].Text.Content = ""
		}
	}

	InfoMsgf(MsgCellContentCleared, row, col)
	return nil
}

// ClearCellFormat clears cell format while preserving content.
func (t *Table) ClearCellFormat(row, col int) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// clear format from cell properties
	if cell.Properties != nil {
		// preserve merge info and base width, clear other format
		oldGridSpan := cell.Properties.GridSpan
		oldVMerge := cell.Properties.VMerge
		oldWidth := cell.Properties.TableCellW

		cell.Properties = &TableCellProperties{
			TableCellW: oldWidth,
			GridSpan:   oldGridSpan,
			VMerge:     oldVMerge,
		}
	}

	// clear paragraph and run format
	for i := range cell.Paragraphs {
		cell.Paragraphs[i].Properties = nil
		for j := range cell.Paragraphs[i].Runs {
			cell.Paragraphs[i].Runs[j].Properties = nil
		}
	}

	InfoMsgf(MsgCellFormatCleared, row, col)
	return nil
}

// SetCellPadding sets the cell padding.
func (t *Table) SetCellPadding(row, col int, padding int) error {
	_, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// cell padding is set via table properties; this is a placeholder interface
	// actual implementation requires setting default padding at the table level
	InfoMsgf(MsgCellPaddingSet, row, col, padding)
	return nil
}

// SetCellTextDirection sets the cell text direction.
func (t *Table) SetCellTextDirection(row, col int, direction CellTextDirection) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// ensure cell has properties
	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	// set text direction
	cell.Properties.TextDirection = &TextDirection{
		Val: string(direction),
	}

	InfoMsgf(MsgCellTextDirectionSet, row, col, direction)
	return nil
}

// GetCellTextDirection returns the cell text direction.
func (t *Table) GetCellTextDirection(row, col int) (CellTextDirection, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return TextDirectionLR, err
	}

	if cell.Properties != nil && cell.Properties.TextDirection != nil {
		return CellTextDirection(cell.Properties.TextDirection.Val), nil
	}

	// default to left-to-right
	return TextDirectionLR, nil
}

// GetCellFormat returns the cell format information.
func (t *Table) GetCellFormat(row, col int) (*CellFormat, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	format := &CellFormat{}

	// get vertical alignment
	if cell.Properties != nil && cell.Properties.VAlign != nil {
		format.VerticalAlign = CellVerticalAlignment(cell.Properties.VAlign.Val)
	}

	// get text direction
	if cell.Properties != nil && cell.Properties.TextDirection != nil {
		format.TextDirection = CellTextDirection(cell.Properties.TextDirection.Val)
	}

	// get horizontal alignment
	if len(cell.Paragraphs) > 0 && cell.Paragraphs[0].Properties != nil && cell.Paragraphs[0].Properties.Justification != nil {
		format.HorizontalAlign = CellAlignment(cell.Paragraphs[0].Properties.Justification.Val)
	}

	// get text format
	if len(cell.Paragraphs) > 0 && len(cell.Paragraphs[0].Runs) > 0 {
		run := &cell.Paragraphs[0].Runs[0]
		if run.Properties != nil {
			format.TextFormat = &TextFormat{}

			if run.Properties.Bold != nil {
				format.TextFormat.Bold = true
			}

			if run.Properties.Italic != nil {
				format.TextFormat.Italic = true
			}

			if run.Properties.FontSize != nil {
				fmt.Sscanf(run.Properties.FontSize.Val, "%d", &format.TextFormat.FontSize)
				format.TextFormat.FontSize /= 2 // convert to points
			}

			if run.Properties.Color != nil {
				format.TextFormat.FontColor = run.Properties.Color.Val
			}

			if run.Properties.FontFamily != nil {
				format.TextFormat.FontFamily = run.Properties.FontFamily.ASCII
			}
		}
	}

	return format, nil
}

// TextDirection represents the text direction.
type TextDirection struct {
	XMLName xml.Name `xml:"w:textDirection"`
	Val     string   `xml:"w:val,attr"`
}

// RowHeightRule represents the row height rule.
type RowHeightRule string

const (
	// RowHeightAuto represents automatic row height.
	RowHeightAuto RowHeightRule = "auto"
	// RowHeightMinimum represents minimum row height.
	RowHeightMinimum RowHeightRule = "atLeast"
	// RowHeightExact represents exact row height.
	RowHeightExact RowHeightRule = "exact"
)

// RowHeightConfig represents the row height configuration.
type RowHeightConfig struct {
	Height int           // row height value (in points, 1 point = 20 twips)
	Rule   RowHeightRule // row height rule
}

// SetRowHeight sets the row height.
func (t *Table) SetRowHeight(rowIndex int, config *RowHeightConfig) error {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return fmt.Errorf("invalid row index: %d, table has %d rows", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties == nil {
		row.Properties = &TableRowProperties{}
	}

	// set row height properties
	row.Properties.TableRowH = &TableRowH{
		Val:   fmt.Sprintf("%d", config.Height*20), // convert to twips (1 point = 20 twips)
		HRule: string(config.Rule),
	}

	InfoMsgf(MsgRowHeightSet, rowIndex, config.Height, config.Rule)
	return nil
}

// GetRowHeight returns the row height configuration.
func (t *Table) GetRowHeight(rowIndex int) (*RowHeightConfig, error) {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return nil, fmt.Errorf("invalid row index: %d, table has %d rows", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties == nil || row.Properties.TableRowH == nil {
		// return default auto height
		return &RowHeightConfig{
			Height: 0,
			Rule:   RowHeightAuto,
		}, nil
	}

	height := 0
	if row.Properties.TableRowH.Val != "" {
		fmt.Sscanf(row.Properties.TableRowH.Val, "%d", &height)
		height /= 20 // convert to points
	}

	rule := RowHeightAuto
	if row.Properties.TableRowH.HRule != "" {
		rule = RowHeightRule(row.Properties.TableRowH.HRule)
	}

	return &RowHeightConfig{
		Height: height,
		Rule:   rule,
	}, nil
}

// SetRowHeightRange sets the row height for a range of rows.
func (t *Table) SetRowHeightRange(startRow, endRow int, config *RowHeightConfig) error {
	if startRow < 0 || endRow >= len(t.Rows) || startRow > endRow {
		return fmt.Errorf("invalid row index range: [%d, %d], table has %d rows", startRow, endRow, len(t.Rows))
	}

	for i := startRow; i <= endRow; i++ {
		err := t.SetRowHeight(i, config)
		if err != nil {
			return fmt.Errorf("failed to set height for row %d: %v", i, err)
		}
	}

	InfoMsgf(MsgRowsHeightSet, startRow, endRow)
	return nil
}

// TableTextWrap represents the table text wrap type.
type TableTextWrap string

const (
	// TextWrapNone represents no text wrap (default).
	TextWrapNone TableTextWrap = "none"
	// TextWrapAround represents text wrapping around the table.
	TextWrapAround TableTextWrap = "around"
)

// TablePosition represents the table position type.
type TablePosition string

const (
	// PositionInline represents inline positioning (default).
	PositionInline TablePosition = "inline"
	// PositionFloating represents floating positioning.
	PositionFloating TablePosition = "floating"
)

// TableAlignment represents the table alignment type.
type TableAlignment string

const (
	// TableAlignLeft represents left alignment.
	TableAlignLeft TableAlignment = "left"
	// TableAlignCenter represents center alignment.
	TableAlignCenter TableAlignment = "center"
	// TableAlignRight represents right alignment.
	TableAlignRight TableAlignment = "right"
	// TableAlignInside represents inside alignment.
	TableAlignInside TableAlignment = "inside"
	// TableAlignOutside represents outside alignment.
	TableAlignOutside TableAlignment = "outside"
)

// TablePositioning represents the table positioning configuration.
type TablePositioning struct {
	XMLName        xml.Name `xml:"w:tblpPr"`
	LeftFromText   string   `xml:"w:leftFromText,attr,omitempty"`   // distance from left text
	RightFromText  string   `xml:"w:rightFromText,attr,omitempty"`  // distance from right text
	TopFromText    string   `xml:"w:topFromText,attr,omitempty"`    // distance from top text
	BottomFromText string   `xml:"w:bottomFromText,attr,omitempty"` // distance from bottom text
	VertAnchor     string   `xml:"w:vertAnchor,attr,omitempty"`     // vertical anchor
	HorzAnchor     string   `xml:"w:horzAnchor,attr,omitempty"`     // horizontal anchor
	TblpXSpec      string   `xml:"w:tblpXSpec,attr,omitempty"`      // horizontal alignment specification
	TblpYSpec      string   `xml:"w:tblpYSpec,attr,omitempty"`      // vertical alignment specification
	TblpX          string   `xml:"w:tblpX,attr,omitempty"`          // X coordinate
	TblpY          string   `xml:"w:tblpY,attr,omitempty"`          // Y coordinate
}

// TableLayoutConfig represents the table layout configuration.
type TableLayoutConfig struct {
	Alignment   TableAlignment    // table alignment
	TextWrap    TableTextWrap     // text wrap type
	Position    TablePosition     // positioning type
	Positioning *TablePositioning // positioning details (only effective when Position is Floating)
}

// SetTableLayout sets the table layout and positioning.
func (t *Table) SetTableLayout(config *TableLayoutConfig) error {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}

	// set table alignment
	if config.Alignment != "" {
		t.Properties.TableJc = &TableJc{
			Val: string(config.Alignment),
		}
	}

	// set positioning attributes (only effective for floating positioning)
	if config.Position == PositionFloating && config.Positioning != nil {
		// in OOXML, floating table positioning requires special TablePositioning attributes
		// store configuration in table properties
		InfoMsg(MsgTableFloatingMode)
		// note: full floating positioning requires more complex XML structure support
	}

	InfoMsgf(MsgTableLayoutSet, config.Alignment, config.TextWrap, config.Position)
	return nil
}

// GetTableLayout returns the table layout configuration.
func (t *Table) GetTableLayout() *TableLayoutConfig {
	config := &TableLayoutConfig{
		Alignment: TableAlignLeft, // default
		TextWrap:  TextWrapNone,
		Position:  PositionInline,
	}

	if t.Properties != nil && t.Properties.TableJc != nil {
		config.Alignment = TableAlignment(t.Properties.TableJc.Val)
	}

	return config
}

// SetTableAlignment sets the table alignment (shortcut method).
func (t *Table) SetTableAlignment(alignment TableAlignment) error {
	return t.SetTableLayout(&TableLayoutConfig{
		Alignment: alignment,
		TextWrap:  TextWrapNone,
		Position:  PositionInline,
	})
}

// TableBreakRule represents the table page break rule.
type TableBreakRule string

const (
	// BreakAuto represents automatic page break (default).
	BreakAuto TableBreakRule = "auto"
	// BreakPage represents a forced page break.
	BreakPage TableBreakRule = "page"
	// BreakColumn represents a forced column break.
	BreakColumn TableBreakRule = "column"
)

// RowBreakConfig represents the row page break configuration.
type RowBreakConfig struct {
	XMLName   xml.Name   `xml:"w:trPr"`
	CantSplit *CantSplit `xml:"w:cantSplit,omitempty"` // prevent page break within row
	TrHeight  *TableRowH `xml:"w:trHeight,omitempty"`  // row height
	TblHeader *TblHeader `xml:"w:tblHeader,omitempty"` // repeat as header row
}

// CantSplit prevents row splitting across pages.
type CantSplit struct {
	XMLName xml.Name `xml:"w:cantSplit"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// TblHeader represents a table header row.
type TblHeader struct {
	XMLName xml.Name `xml:"w:tblHeader"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// SetRowKeepTogether sets whether a row should not split across pages.
func (t *Table) SetRowKeepTogether(rowIndex int, keepTogether bool) error {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return fmt.Errorf("invalid row index: %d, table has %d rows", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties == nil {
		row.Properties = &TableRowProperties{}
	}

	if keepTogether {
		row.Properties.CantSplit = &CantSplit{
			Val: "1",
		}
	} else {
		row.Properties.CantSplit = nil
	}

	InfoMsgf(MsgRowPageSplitSet, rowIndex, !keepTogether)
	return nil
}

// SetRowAsHeader sets a row as a repeating header row.
func (t *Table) SetRowAsHeader(rowIndex int, isHeader bool) error {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return fmt.Errorf("invalid row index: %d, table has %d rows", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties == nil {
		row.Properties = &TableRowProperties{}
	}

	if isHeader {
		row.Properties.TblHeader = &TblHeader{
			Val: "1",
		}
	} else {
		row.Properties.TblHeader = nil
	}

	InfoMsgf(MsgRowSetAsHeader, rowIndex, isHeader)
	return nil
}

// SetHeaderRows sets the table header row range.
func (t *Table) SetHeaderRows(startRow, endRow int) error {
	if startRow < 0 || endRow >= len(t.Rows) || startRow > endRow {
		return fmt.Errorf("invalid row index range: [%d, %d], table has %d rows", startRow, endRow, len(t.Rows))
	}

	// clear all existing header row settings
	for i := range t.Rows {
		if t.Rows[i].Properties != nil {
			t.Rows[i].Properties.TblHeader = nil
		}
	}

	// set specified range as header rows
	for i := startRow; i <= endRow; i++ {
		err := t.SetRowAsHeader(i, true)
		if err != nil {
			return fmt.Errorf("failed to set row %d as header: %v", i, err)
		}
	}

	InfoMsgf(MsgRowsSetAsHeaders, startRow, endRow)
	return nil
}

// IsRowHeader checks whether a row is a header row.
func (t *Table) IsRowHeader(rowIndex int) (bool, error) {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return false, fmt.Errorf("invalid row index: %d, table has %d rows", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties != nil && row.Properties.TblHeader != nil {
		return row.Properties.TblHeader.Val == "1", nil
	}

	return false, nil
}

// IsRowKeepTogether checks whether a row prevents page break splitting.
func (t *Table) IsRowKeepTogether(rowIndex int) (bool, error) {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return false, fmt.Errorf("invalid row index: %d, table has %d rows", rowIndex, len(t.Rows))
	}

	row := &t.Rows[rowIndex]
	if row.Properties != nil && row.Properties.CantSplit != nil {
		return row.Properties.CantSplit.Val == "1", nil
	}

	return false, nil
}

// TablePageBreakConfig represents the table page break configuration.
type TablePageBreakConfig struct {
	KeepWithNext    bool // keep with next paragraph
	KeepLines       bool // keep lines together
	PageBreakBefore bool // page break before paragraph
	WidowControl    bool // widow/orphan control
}

// SetTablePageBreak sets the table page break control.
func (t *Table) SetTablePageBreak(config *TablePageBreakConfig) error {
	// table-level page break control is typically set in table properties
	// record configuration here; actual XML output requires corresponding implementation
	InfoMsgf(MsgTablePagination, config.KeepWithNext, config.KeepLines, config.PageBreakBefore, config.WidowControl)
	return nil
}

// SetRowKeepWithNext sets whether a row should stay on the same page as the next row.
func (t *Table) SetRowKeepWithNext(rowIndex int, keepWithNext bool) error {
	if rowIndex < 0 || rowIndex >= len(t.Rows) {
		return fmt.Errorf("invalid row index: %d, table has %d rows", rowIndex, len(t.Rows))
	}

	// this feature requires setting specific page break properties in row properties
	// actual implementation requires extending the TableRowProperties struct
	InfoMsgf(MsgRowKeepWithNextSet, rowIndex, keepWithNext)
	return nil
}

// GetTableBreakInfo returns table page break information.
func (t *Table) GetTableBreakInfo() map[string]interface{} {
	info := make(map[string]interface{})

	headerRowCount := 0
	keepTogetherCount := 0

	for i := range t.Rows {
		isHeader, _ := t.IsRowHeader(i)
		if isHeader {
			headerRowCount++
		}

		keepTogether, _ := t.IsRowKeepTogether(i)
		if keepTogether {
			keepTogetherCount++
		}
	}

	info["total_rows"] = len(t.Rows)
	info["header_rows"] = headerRowCount
	info["keep_together_rows"] = keepTogetherCount

	return info
}

// extend TableRowProperties to support page break control
type TableRowPropertiesExtended struct {
	XMLName   xml.Name   `xml:"w:trPr"`
	TableRowH *TableRowH `xml:"w:trHeight,omitempty"`
	CantSplit *CantSplit `xml:"w:cantSplit,omitempty"`
	TblHeader *TblHeader `xml:"w:tblHeader,omitempty"`
	KeepNext  *KeepNext  `xml:"w:keepNext,omitempty"`
	KeepLines *KeepLines `xml:"w:keepLines,omitempty"`
}

// extend the existing TableRowProperties struct
func (trp *TableRowProperties) SetCantSplit(cantSplit bool) {
	if cantSplit {
		trp.CantSplit = &CantSplit{Val: "1"}
	} else {
		trp.CantSplit = nil
	}
}

func (trp *TableRowProperties) SetTblHeader(isHeader bool) {
	if isHeader {
		trp.TblHeader = &TblHeader{Val: "1"}
	} else {
		trp.TblHeader = nil
	}
}

// TableStyle represents a table style reference.
type TableStyle struct {
	XMLName xml.Name `xml:"w:tblStyle"`
	Val     string   `xml:"w:val,attr"`
}

// TableBorders represents table borders.
type TableBorders struct {
	XMLName xml.Name     `xml:"w:tblBorders"`
	Top     *TableBorder `xml:"w:top,omitempty"`     // top border
	Left    *TableBorder `xml:"w:left,omitempty"`    // left border
	Bottom  *TableBorder `xml:"w:bottom,omitempty"`  // bottom border
	Right   *TableBorder `xml:"w:right,omitempty"`   // right border
	InsideH *TableBorder `xml:"w:insideH,omitempty"` // inside horizontal border
	InsideV *TableBorder `xml:"w:insideV,omitempty"` // inside vertical border
}

// TableBorder represents a border definition.
type TableBorder struct {
	Val        string `xml:"w:val,attr"`                  // border style
	Sz         string `xml:"w:sz,attr"`                   // border thickness (1/8 point)
	Space      string `xml:"w:space,attr"`                // border spacing
	Color      string `xml:"w:color,attr"`                // border color
	ThemeColor string `xml:"w:themeColor,attr,omitempty"` // theme color
}

// TableShading represents table shading/background.
type TableShading struct {
	XMLName   xml.Name `xml:"w:shd"`
	Val       string   `xml:"w:val,attr"`                 // shading pattern
	Color     string   `xml:"w:color,attr,omitempty"`     // foreground color
	Fill      string   `xml:"w:fill,attr,omitempty"`      // background color
	ThemeFill string   `xml:"w:themeFill,attr,omitempty"` // theme fill color
}

// TableCellMargins represents table cell margins.
type TableCellMargins struct {
	XMLName xml.Name        `xml:"w:tblCellMar"`
	Top     *TableCellSpace `xml:"w:top,omitempty"`
	Left    *TableCellSpace `xml:"w:left,omitempty"`
	Bottom  *TableCellSpace `xml:"w:bottom,omitempty"`
	Right   *TableCellSpace `xml:"w:right,omitempty"`
}

// TableCellSpace represents table cell spacing.
type TableCellSpace struct {
	W    string `xml:"w:w,attr"`
	Type string `xml:"w:type,attr"`
}

// TableLayoutType represents the table layout type.
type TableLayoutType struct {
	XMLName xml.Name `xml:"w:tblLayout"`
	Type    string   `xml:"w:type,attr"` // fixed, autofit
}

// TableIndentation represents table indentation.
type TableIndentation struct {
	XMLName xml.Name `xml:"w:tblInd"`
	W       string   `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

// TableCellShading represents cell background shading.
type TableCellShading struct {
	XMLName   xml.Name `xml:"w:shd"`
	Val       string   `xml:"w:val,attr"`                 // shading pattern
	Color     string   `xml:"w:color,attr,omitempty"`     // foreground color
	Fill      string   `xml:"w:fill,attr,omitempty"`      // background color
	ThemeFill string   `xml:"w:themeFill,attr,omitempty"` // theme fill color
}

// TableCellBorders represents cell borders.
type TableCellBorders struct {
	XMLName xml.Name         `xml:"w:tcBorders"`
	Top     *TableCellBorder `xml:"w:top,omitempty"`     // top border
	Left    *TableCellBorder `xml:"w:left,omitempty"`    // left border
	Bottom  *TableCellBorder `xml:"w:bottom,omitempty"`  // bottom border
	Right   *TableCellBorder `xml:"w:right,omitempty"`   // right border
	InsideH *TableCellBorder `xml:"w:insideH,omitempty"` // inside horizontal border
	InsideV *TableCellBorder `xml:"w:insideV,omitempty"` // inside vertical border
	TL2BR   *TableCellBorder `xml:"w:tl2br,omitempty"`   // top-left to bottom-right diagonal
	TR2BL   *TableCellBorder `xml:"w:tr2bl,omitempty"`   // top-right to bottom-left diagonal
}

// TableCellBorder represents a cell border definition.
type TableCellBorder struct {
	Val        string `xml:"w:val,attr"`                  // border style
	Sz         string `xml:"w:sz,attr"`                   // border thickness (1/8 point)
	Space      string `xml:"w:space,attr"`                // border spacing
	Color      string `xml:"w:color,attr"`                // border color
	ThemeColor string `xml:"w:themeColor,attr,omitempty"` // theme color
}

// NoWrap prevents text wrapping.
type NoWrap struct {
	XMLName xml.Name `xml:"w:noWrap"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// HideMark hides the end-of-cell marker.
type HideMark struct {
	XMLName xml.Name `xml:"w:hideMark"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// ============== Table Style and Appearance Features ==============

// BorderStyle represents a border style constant.
type BorderStyle string

const (
	BorderStyleNone                   BorderStyle = "none"                   // no border
	BorderStyleSingle                 BorderStyle = "single"                 // single line
	BorderStyleThick                  BorderStyle = "thick"                  // thick line
	BorderStyleDouble                 BorderStyle = "double"                 // double line
	BorderStyleDotted                 BorderStyle = "dotted"                 // dotted line
	BorderStyleDashed                 BorderStyle = "dashed"                 // dashed line
	BorderStyleDotDash                BorderStyle = "dotDash"                // dot-dash line
	BorderStyleDotDotDash             BorderStyle = "dotDotDash"             // dot-dot-dash line
	BorderStyleTriple                 BorderStyle = "triple"                 // triple line
	BorderStyleThinThickSmallGap      BorderStyle = "thinThickSmallGap"      // thin-thick-thin (small gap)
	BorderStyleThickThinSmallGap      BorderStyle = "thickThinSmallGap"      // thick-thin-thick (small gap)
	BorderStyleThinThickThinSmallGap  BorderStyle = "thinThickThinSmallGap"  // thin-thick-thin (small gap)
	BorderStyleThinThickMediumGap     BorderStyle = "thinThickMediumGap"     // thin-thick-thin (medium gap)
	BorderStyleThickThinMediumGap     BorderStyle = "thickThinMediumGap"     // thick-thin-thick (medium gap)
	BorderStyleThinThickThinMediumGap BorderStyle = "thinThickThinMediumGap" // thin-thick-thin (medium gap)
	BorderStyleThinThickLargeGap      BorderStyle = "thinThickLargeGap"      // thin-thick-thin (large gap)
	BorderStyleThickThinLargeGap      BorderStyle = "thickThinLargeGap"      // thick-thin-thick (large gap)
	BorderStyleThinThickThinLargeGap  BorderStyle = "thinThickThinLargeGap"  // thin-thick-thin (large gap)
	BorderStyleWave                   BorderStyle = "wave"                   // wave line
	BorderStyleDoubleWave             BorderStyle = "doubleWave"             // double wave line
	BorderStyleDashSmallGap           BorderStyle = "dashSmallGap"           // dashed (small gap)
	BorderStyleDashDotStroked         BorderStyle = "dashDotStroked"         // dash-dot stroked
	BorderStyleThreeDEmboss           BorderStyle = "threeDEmboss"           // 3D emboss
	BorderStyleThreeDEngrave          BorderStyle = "threeDEngrave"          // 3D engrave
	BorderStyleOutset                 BorderStyle = "outset"                 // outset
	BorderStyleInset                  BorderStyle = "inset"                  // inset
)

// ShadingPattern represents the shading pattern type.
type ShadingPattern string

const (
	ShadingPatternClear             ShadingPattern = "clear"             // clear/solid
	ShadingPatternSolid             ShadingPattern = "clear"             // solid (implemented using clear)
	ShadingPatternPct5              ShadingPattern = "pct5"              // 5%
	ShadingPatternPct10             ShadingPattern = "pct10"             // 10%
	ShadingPatternPct20             ShadingPattern = "pct20"             // 20%
	ShadingPatternPct25             ShadingPattern = "pct25"             // 25%
	ShadingPatternPct30             ShadingPattern = "pct30"             // 30%
	ShadingPatternPct40             ShadingPattern = "pct40"             // 40%
	ShadingPatternPct50             ShadingPattern = "pct50"             // 50%
	ShadingPatternPct60             ShadingPattern = "pct60"             // 60%
	ShadingPatternPct70             ShadingPattern = "pct70"             // 70%
	ShadingPatternPct75             ShadingPattern = "pct75"             // 75%
	ShadingPatternPct80             ShadingPattern = "pct80"             // 80%
	ShadingPatternPct90             ShadingPattern = "pct90"             // 90%
	ShadingPatternHorzStripe        ShadingPattern = "horzStripe"        // horizontal stripe
	ShadingPatternVertStripe        ShadingPattern = "vertStripe"        // vertical stripe
	ShadingPatternReverseDiagStripe ShadingPattern = "reverseDiagStripe" // reverse diagonal stripe
	ShadingPatternDiagStripe        ShadingPattern = "diagStripe"        // diagonal stripe
	ShadingPatternHorzCross         ShadingPattern = "horzCross"         // horizontal cross
	ShadingPatternDiagCross         ShadingPattern = "diagCross"         // diagonal cross
)

// TableStyleTemplate represents a table style template.
type TableStyleTemplate string

const (
	TableStyleTemplateNormal    TableStyleTemplate = "TableNormal"    // normal table
	TableStyleTemplateGrid      TableStyleTemplate = "TableGrid"      // grid table
	TableStyleTemplateList      TableStyleTemplate = "TableList"      // list table
	TableStyleTemplateColorful1 TableStyleTemplate = "TableColorful1" // colorful table 1
	TableStyleTemplateColorful2 TableStyleTemplate = "TableColorful2" // colorful table 2
	TableStyleTemplateColorful3 TableStyleTemplate = "TableColorful3" // colorful table 3
	TableStyleTemplateColumns1  TableStyleTemplate = "TableColumns1"  // column style 1
	TableStyleTemplateColumns2  TableStyleTemplate = "TableColumns2"  // column style 2
	TableStyleTemplateColumns3  TableStyleTemplate = "TableColumns3"  // column style 3
	TableStyleTemplateRows1     TableStyleTemplate = "TableRows1"     // row style 1
	TableStyleTemplateRows2     TableStyleTemplate = "TableRows2"     // row style 2
	TableStyleTemplateRows3     TableStyleTemplate = "TableRows3"     // row style 3
	TableStyleTemplatePlain1    TableStyleTemplate = "TablePlain1"    // plain table 1
	TableStyleTemplatePlain2    TableStyleTemplate = "TablePlain2"    // plain table 2
	TableStyleTemplatePlain3    TableStyleTemplate = "TablePlain3"    // plain table 3
)

// TableStyleConfig represents the table style configuration.
type TableStyleConfig struct {
	Template          TableStyleTemplate // style template
	StyleID           string             // custom style ID
	FirstRowHeader    bool               // first row as header
	LastRowTotal      bool               // last row as total
	FirstColumnHeader bool               // first column as header
	LastColumnTotal   bool               // last column as total
	BandedRows        bool               // alternating row colors
	BandedColumns     bool               // alternating column colors
}

// BorderConfig represents border configuration.
type BorderConfig struct {
	Style BorderStyle // border style
	Width int         // border width (1/8 point)
	Color string      // border color (hex, e.g. "FF0000")
	Space int         // border spacing
}

// ShadingConfig represents shading configuration.
type ShadingConfig struct {
	Pattern         ShadingPattern // shading pattern
	ForegroundColor string         // foreground color (hex)
	BackgroundColor string         // background color (hex)
}

// TableBorderConfig represents table border configuration.
type TableBorderConfig struct {
	Top     *BorderConfig // top border
	Left    *BorderConfig // left border
	Bottom  *BorderConfig // bottom border
	Right   *BorderConfig // right border
	InsideH *BorderConfig // inside horizontal border
	InsideV *BorderConfig // inside vertical border
}

// CellBorderConfig represents cell border configuration.
type CellBorderConfig struct {
	Top      *BorderConfig // top border
	Left     *BorderConfig // left border
	Bottom   *BorderConfig // bottom border
	Right    *BorderConfig // right border
	DiagDown *BorderConfig // top-left to bottom-right diagonal
	DiagUp   *BorderConfig // top-right to bottom-left diagonal
}

// ApplyTableStyle applies a table style.
func (t *Table) ApplyTableStyle(config *TableStyleConfig) error {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}

	// set style template
	if config.Template != "" {
		t.Properties.TableStyle = &TableStyle{
			Val: string(config.Template),
		}
	} else if config.StyleID != "" {
		t.Properties.TableStyle = &TableStyle{
			Val: config.StyleID,
		}
	}

	// set table appearance options
	if t.Properties.TableLook == nil {
		t.Properties.TableLook = &TableLook{}
	}

	// build TableLook value
	lookVal := "0000"
	if config.FirstRowHeader {
		t.Properties.TableLook.FirstRow = "1"
		lookVal = "0400"
	} else {
		t.Properties.TableLook.FirstRow = "0"
	}

	if config.LastRowTotal {
		t.Properties.TableLook.LastRow = "1"
		if lookVal == "0400" {
			lookVal = "0440"
		} else {
			lookVal = "0040"
		}
	} else {
		t.Properties.TableLook.LastRow = "0"
	}

	if config.FirstColumnHeader {
		t.Properties.TableLook.FirstCol = "1"
		switch lookVal {
		case "0400":
			lookVal = "0500"
		case "0040":
			lookVal = "0140"
		case "0440":
			lookVal = "0540"
		default:
			lookVal = "0100"
		}
	} else {
		t.Properties.TableLook.FirstCol = "0"
	}

	if config.LastColumnTotal {
		t.Properties.TableLook.LastCol = "1"
	} else {
		t.Properties.TableLook.LastCol = "0"
	}

	if config.BandedRows {
		t.Properties.TableLook.NoHBand = "0"
	} else {
		t.Properties.TableLook.NoHBand = "1"
	}

	if config.BandedColumns {
		t.Properties.TableLook.NoVBand = "0"
	} else {
		t.Properties.TableLook.NoVBand = "1"
	}

	t.Properties.TableLook.Val = lookVal

	InfoMsgf(MsgTableStyleApplied, config.Template)
	return nil
}

// SetTableBorders sets the table borders.
func (t *Table) SetTableBorders(config *TableBorderConfig) error {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}

	t.Properties.TableBorders = &TableBorders{}

	if config.Top != nil {
		t.Properties.TableBorders.Top = createTableBorder(config.Top)
	}
	if config.Left != nil {
		t.Properties.TableBorders.Left = createTableBorder(config.Left)
	}
	if config.Bottom != nil {
		t.Properties.TableBorders.Bottom = createTableBorder(config.Bottom)
	}
	if config.Right != nil {
		t.Properties.TableBorders.Right = createTableBorder(config.Right)
	}
	if config.InsideH != nil {
		t.Properties.TableBorders.InsideH = createTableBorder(config.InsideH)
	}
	if config.InsideV != nil {
		t.Properties.TableBorders.InsideV = createTableBorder(config.InsideV)
	}

	InfoMsg(MsgTableBorderSet)
	return nil
}

// SetTableShading sets the table background shading.
func (t *Table) SetTableShading(config *ShadingConfig) error {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}

	t.Properties.Shd = &TableShading{
		Val:   string(config.Pattern),
		Color: config.ForegroundColor,
		Fill:  config.BackgroundColor,
	}

	InfoMsg(MsgTableBackgroundSet)
	return nil
}

// SetCellBorders sets the cell borders.
func (t *Table) SetCellBorders(row, col int, config *CellBorderConfig) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	cell.Properties.TcBorders = &TableCellBorders{}

	if config.Top != nil {
		cell.Properties.TcBorders.Top = createTableCellBorder(config.Top)
	}
	if config.Left != nil {
		cell.Properties.TcBorders.Left = createTableCellBorder(config.Left)
	}
	if config.Bottom != nil {
		cell.Properties.TcBorders.Bottom = createTableCellBorder(config.Bottom)
	}
	if config.Right != nil {
		cell.Properties.TcBorders.Right = createTableCellBorder(config.Right)
	}
	if config.DiagDown != nil {
		cell.Properties.TcBorders.TL2BR = createTableCellBorder(config.DiagDown)
	}
	if config.DiagUp != nil {
		cell.Properties.TcBorders.TR2BL = createTableCellBorder(config.DiagUp)
	}

	InfoMsgf(MsgCellBorderSet, row, col)
	return nil
}

// SetCellShading sets the cell background shading.
func (t *Table) SetCellShading(row, col int, config *ShadingConfig) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	if cell.Properties == nil {
		cell.Properties = &TableCellProperties{}
	}

	cell.Properties.Shd = &TableCellShading{
		Val:   string(config.Pattern),
		Color: config.ForegroundColor,
		Fill:  config.BackgroundColor,
	}

	InfoMsgf(MsgCellBackgroundSet, row, col)
	return nil
}

// SetAlternatingRowColors sets alternating row colors.
func (t *Table) SetAlternatingRowColors(evenRowColor, oddRowColor string) error {
	for i := range t.Rows {
		var bgColor string
		if i%2 == 0 {
			bgColor = evenRowColor
		} else {
			bgColor = oddRowColor
		}

		// set background color for all cells in the row
		for j := range t.Rows[i].Cells {
			err := t.SetCellShading(i, j, &ShadingConfig{
				Pattern:         ShadingPatternSolid,
				BackgroundColor: bgColor,
			})
			if err != nil {
				return fmt.Errorf("failed to set background color for row %d, column %d: %v", i, j, err)
			}
		}
	}

	InfoMsg(MsgAlternatingRowColorsSet)
	return nil
}

// RemoveTableBorders removes all table borders.
func (t *Table) RemoveTableBorders() error {
	if t.Properties == nil {
		t.Properties = &TableProperties{}
	}

	// set all borders to none
	noBorderConfig := &BorderConfig{
		Style: BorderStyleNone,
		Width: 0,
		Color: "auto",
		Space: 0,
	}

	borderConfig := &TableBorderConfig{
		Top:     noBorderConfig,
		Left:    noBorderConfig,
		Bottom:  noBorderConfig,
		Right:   noBorderConfig,
		InsideH: noBorderConfig,
		InsideV: noBorderConfig,
	}

	return t.SetTableBorders(borderConfig)
}

// RemoveCellBorders removes all cell borders.
func (t *Table) RemoveCellBorders(row, col int) error {
	noBorderConfig := &BorderConfig{
		Style: BorderStyleNone,
		Width: 0,
		Color: "auto",
		Space: 0,
	}

	cellBorderConfig := &CellBorderConfig{
		Top:    noBorderConfig,
		Left:   noBorderConfig,
		Bottom: noBorderConfig,
		Right:  noBorderConfig,
	}

	return t.SetCellBorders(row, col, cellBorderConfig)
}

// CreateCustomTableStyle creates a custom table style.
func (t *Table) CreateCustomTableStyle(styleID, styleName string,
	borderConfig *TableBorderConfig,
	shadingConfig *ShadingConfig,
	firstRowBold bool) error {

	// apply style to table
	config := &TableStyleConfig{
		StyleID:        styleID,
		FirstRowHeader: firstRowBold,
		BandedRows:     shadingConfig != nil,
	}

	err := t.ApplyTableStyle(config)
	if err != nil {
		return err
	}

	// set borders
	if borderConfig != nil {
		err = t.SetTableBorders(borderConfig)
		if err != nil {
			return err
		}
	}

	// set background
	if shadingConfig != nil {
		err = t.SetTableShading(shadingConfig)
		if err != nil {
			return err
		}
	}

	InfoMsgf(MsgCustomTableStyleCreated, styleID)
	return nil
}

// helper function: create table border
func createTableBorder(config *BorderConfig) *TableBorder {
	return &TableBorder{
		Val:   string(config.Style),
		Sz:    fmt.Sprintf("%d", config.Width),
		Space: fmt.Sprintf("%d", config.Space),
		Color: config.Color,
	}
}

// helper function: create cell border
func createTableCellBorder(config *BorderConfig) *TableCellBorder {
	return &TableCellBorder{
		Val:   string(config.Style),
		Sz:    fmt.Sprintf("%d", config.Width),
		Space: fmt.Sprintf("%d", config.Space),
		Color: config.Color,
	}
}

// CellIterator is a cell iterator.
type CellIterator struct {
	table      *Table
	currentRow int
	currentCol int
	totalRows  int
	totalCols  int
}

// CellInfo represents cell information.
type CellInfo struct {
	Row    int        // row index
	Col    int        // column index
	Cell   *TableCell // cell reference
	Text   string     // cell text
	IsLast bool       // whether this is the last cell
}

// NewCellIterator creates a new cell iterator.
func (t *Table) NewCellIterator() *CellIterator {
	totalRows := t.GetRowCount()
	totalCols := 0
	if totalRows > 0 {
		totalCols = t.GetColumnCount()
	}

	return &CellIterator{
		table:      t,
		currentRow: 0,
		currentCol: 0,
		totalRows:  totalRows,
		totalCols:  totalCols,
	}
}

// HasNext checks whether there is a next cell.
func (iter *CellIterator) HasNext() bool {
	if iter.totalRows == 0 || iter.totalCols == 0 {
		return false
	}

	// check if current position is out of range
	return iter.currentRow < iter.totalRows &&
		(iter.currentRow < iter.totalRows-1 || iter.currentCol < iter.totalCols)
}

// Next returns the next cell information.
func (iter *CellIterator) Next() (*CellInfo, error) {
	if !iter.HasNext() {
		return nil, fmt.Errorf("no more cells")
	}

	// get current cell
	cell, err := iter.table.GetCell(iter.currentRow, iter.currentCol)
	if err != nil {
		return nil, fmt.Errorf("failed to get cell: %v", err)
	}

	// get cell text
	text, _ := iter.table.GetCellText(iter.currentRow, iter.currentCol)

	// create cell info
	cellInfo := &CellInfo{
		Row:  iter.currentRow,
		Col:  iter.currentCol,
		Cell: cell,
		Text: text,
	}

	// update position and check if last
	iter.currentCol++
	if iter.currentCol >= iter.totalCols {
		iter.currentCol = 0
		iter.currentRow++
	}

	// check if this is the last cell
	cellInfo.IsLast = !iter.HasNext()

	return cellInfo, nil
}

// Reset resets the iterator to the starting position.
func (iter *CellIterator) Reset() {
	iter.currentRow = 0
	iter.currentCol = 0
}

// Current returns the current position (without advancing the iterator).
func (iter *CellIterator) Current() (int, int) {
	return iter.currentRow, iter.currentCol
}

// Total returns the total number of cells.
func (iter *CellIterator) Total() int {
	return iter.totalRows * iter.totalCols
}

// Progress returns the iteration progress (0.0-1.0).
func (iter *CellIterator) Progress() float64 {
	if iter.totalRows == 0 || iter.totalCols == 0 {
		return 1.0
	}

	processed := iter.currentRow*iter.totalCols + iter.currentCol
	total := iter.totalRows * iter.totalCols

	return float64(processed) / float64(total)
}

// ForEach iterates over all cells and executes the specified function for each.
func (t *Table) ForEach(fn func(row, col int, cell *TableCell, text string) error) error {
	iterator := t.NewCellIterator()

	for iterator.HasNext() {
		cellInfo, err := iterator.Next()
		if err != nil {
			return fmt.Errorf("iteration failed: %v", err)
		}

		if err := fn(cellInfo.Row, cellInfo.Col, cellInfo.Cell, cellInfo.Text); err != nil {
			return fmt.Errorf("callback execution failed (row: %d, col: %d): %v", cellInfo.Row, cellInfo.Col, err)
		}
	}

	return nil
}

// ForEachInRow iterates over all cells in the specified row.
func (t *Table) ForEachInRow(rowIndex int, fn func(col int, cell *TableCell, text string) error) error {
	if rowIndex < 0 || rowIndex >= t.GetRowCount() {
		return fmt.Errorf("invalid row index: %d", rowIndex)
	}

	colCount := t.GetColumnCount()
	for col := 0; col < colCount; col++ {
		cell, err := t.GetCell(rowIndex, col)
		if err != nil {
			return fmt.Errorf("failed to get cell (row: %d, col: %d): %v", rowIndex, col, err)
		}

		text, _ := t.GetCellText(rowIndex, col)

		if err := fn(col, cell, text); err != nil {
			return fmt.Errorf("callback execution failed (row: %d, col: %d): %v", rowIndex, col, err)
		}
	}

	return nil
}

// ForEachInColumn iterates over all cells in the specified column.
func (t *Table) ForEachInColumn(colIndex int, fn func(row int, cell *TableCell, text string) error) error {
	if colIndex < 0 || colIndex >= t.GetColumnCount() {
		return fmt.Errorf("invalid column index: %d", colIndex)
	}

	rowCount := t.GetRowCount()
	for row := 0; row < rowCount; row++ {
		cell, err := t.GetCell(row, colIndex)
		if err != nil {
			return fmt.Errorf("failed to get cell (row: %d, col: %d): %v", row, colIndex, err)
		}

		text, _ := t.GetCellText(row, colIndex)

		if err := fn(row, cell, text); err != nil {
			return fmt.Errorf("callback execution failed (row: %d, col: %d): %v", row, colIndex, err)
		}
	}

	return nil
}

// GetCellRange returns all cells in the specified range.
func (t *Table) GetCellRange(startRow, startCol, endRow, endCol int) ([]*CellInfo, error) {
	// parameter validation
	if startRow < 0 || startCol < 0 || endRow >= t.GetRowCount() || endCol >= t.GetColumnCount() {
		return nil, fmt.Errorf("invalid range index: (%d,%d) to (%d,%d)", startRow, startCol, endRow, endCol)
	}

	if startRow > endRow || startCol > endCol {
		return nil, fmt.Errorf("start position cannot be greater than end position")
	}

	var cells []*CellInfo

	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			cell, err := t.GetCell(row, col)
			if err != nil {
				return nil, fmt.Errorf("failed to get cell (row: %d, col: %d): %v", row, col, err)
			}

			text, _ := t.GetCellText(row, col)

			cellInfo := &CellInfo{
				Row:    row,
				Col:    col,
				Cell:   cell,
				Text:   text,
				IsLast: row == endRow && col == endCol,
			}

			cells = append(cells, cellInfo)
		}
	}

	return cells, nil
}

// FindCells finds cells matching the given predicate.
func (t *Table) FindCells(predicate func(row, col int, cell *TableCell, text string) bool) ([]*CellInfo, error) {
	var matchedCells []*CellInfo

	err := t.ForEach(func(row, col int, cell *TableCell, text string) error {
		if predicate(row, col, cell, text) {
			cellInfo := &CellInfo{
				Row:  row,
				Col:  col,
				Cell: cell,
				Text: text,
			}
			matchedCells = append(matchedCells, cellInfo)
		}
		return nil
	})

	if err != nil {
		return nil, err
	}

	return matchedCells, nil
}

// FindCellsByText finds cells by text content.
func (t *Table) FindCellsByText(searchText string, exactMatch bool) ([]*CellInfo, error) {
	return t.FindCells(func(row, col int, cell *TableCell, text string) bool {
		if exactMatch {
			return text == searchText
		}
		// use strings.Contains for fuzzy matching
		return strings.Contains(text, searchText)
	})
}

// ============== Cell Complex Content Features ==============
// The following methods support adding paragraphs, images, lists, nested tables, and other complex content to table cells.

// AddCellParagraph adds a paragraph to a cell.
// Parameters:
//   - row: row index (0-based)
//   - col: column index (0-based)
//   - text: paragraph text content
//
// Returns:
//   - *Paragraph: the newly added paragraph object
//   - error: an error if the index is invalid
func (t *Table) AddCellParagraph(row, col int, text string) (*Paragraph, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	// create new paragraph
	para := &Paragraph{
		Runs: []Run{
			{
				Text: Text{
					Content: text,
					Space:   "preserve",
				},
			},
		},
	}

	// add to cell
	cell.Paragraphs = append(cell.Paragraphs, *para)

	InfoMsgf(MsgParagraphAddedToCell, row, col)
	return &cell.Paragraphs[len(cell.Paragraphs)-1], nil
}

// AddCellFormattedParagraph adds a formatted paragraph to a cell.
// Parameters:
//   - row: row index (0-based)
//   - col: column index (0-based)
//   - text: paragraph text content
//   - format: text format configuration
//
// Returns:
//   - *Paragraph: the newly added paragraph object
//   - error: an error if the index is invalid
func (t *Table) AddCellFormattedParagraph(row, col int, text string, format *TextFormat) (*Paragraph, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	// create run properties
	runProps := &RunProperties{}

	if format != nil {
		if format.FontFamily != "" {
			runProps.FontFamily = &FontFamily{
				ASCII:    format.FontFamily,
				HAnsi:    format.FontFamily,
				EastAsia: format.FontFamily,
				CS:       format.FontFamily,
			}
		}

		if format.Bold {
			runProps.Bold = &Bold{}
		}

		if format.Italic {
			runProps.Italic = &Italic{}
		}

		if format.FontColor != "" {
			color := strings.TrimPrefix(format.FontColor, "#")
			runProps.Color = &Color{Val: color}
		}

		if format.FontSize > 0 {
			runProps.FontSize = &FontSize{Val: fmt.Sprintf("%d", format.FontSize*2)}
		}

		if format.Underline {
			runProps.Underline = &Underline{Val: "single"}
		}

		if format.Strike {
			runProps.Strike = &Strike{}
		}

		if format.Highlight != "" {
			runProps.Highlight = &Highlight{Val: format.Highlight}
		}
	}

	// create new paragraph
	para := &Paragraph{
		Runs: []Run{
			{
				Properties: runProps,
				Text: Text{
					Content: text,
					Space:   "preserve",
				},
			},
		},
	}

	// add to cell
	cell.Paragraphs = append(cell.Paragraphs, *para)

	InfoMsgf(MsgFormattedParagraphAddedToCell, row, col)
	return &cell.Paragraphs[len(cell.Paragraphs)-1], nil
}

// ClearCellParagraphs clears all paragraphs in a cell, keeping only one empty paragraph.
// Parameters:
//   - row: row index (0-based)
//   - col: column index (0-based)
//
// Returns:
//   - error: an error if the index is invalid
func (t *Table) ClearCellParagraphs(row, col int) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	// clear paragraphs, keep only one empty paragraph (OOXML spec requires at least one paragraph per cell)
	cell.Paragraphs = []Paragraph{
		{
			Runs: []Run{
				{
					Text: Text{Content: ""},
				},
			},
		},
	}

	InfoMsgf(MsgCellParagraphsCleared, row, col)
	return nil
}

// GetCellParagraphs returns all paragraphs in a cell.
// Parameters:
//   - row: row index (0-based)
//   - col: column index (0-based)
//
// Returns:
//   - []Paragraph: all paragraphs in the cell
//   - error: an error if the index is invalid
func (t *Table) GetCellParagraphs(row, col int) ([]Paragraph, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	return cell.Paragraphs, nil
}

// AddNestedTable adds a nested table to a cell.
// Parameters:
//   - row: row index (0-based)
//   - col: column index (0-based)
//   - config: nested table configuration
//
// Returns:
//   - *Table: the newly created nested table object
//   - error: an error if the index or configuration is invalid
func (t *Table) AddNestedTable(row, col int, config *TableConfig) (*Table, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	if config.Rows <= 0 || config.Cols <= 0 {
		Error("nested table rows and columns must be greater than 0")
		return nil, NewValidationError("TableConfig", "", "nested table rows and columns must be greater than 0")
	}

	// create nested table
	nestedTable := &Table{
		Properties: &TableProperties{
			TableW: &TableWidth{
				W:    fmt.Sprintf("%d", config.Width),
				Type: "dxa",
			},
			TableJc: &TableJc{
				Val: "center",
			},
			TableLook: &TableLook{
				Val:      "04A0",
				FirstRow: "1",
				LastRow:  "0",
				FirstCol: "1",
				LastCol:  "0",
				NoHBand:  "0",
				NoVBand:  "1",
			},
			TableBorders: &TableBorders{
				Top:     &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
				Left:    &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
				Bottom:  &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
				Right:   &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
				InsideH: &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
				InsideV: &TableBorder{Val: "single", Sz: "4", Space: "0", Color: "auto"},
			},
			TableLayout: &TableLayoutType{
				Type: "autofit",
			},
			TableCellMar: &TableCellMargins{
				Left:  &TableCellSpace{W: "108", Type: "dxa"},
				Right: &TableCellSpace{W: "108", Type: "dxa"},
			},
		},
		Grid: &TableGrid{},
		Rows: make([]TableRow, 0, config.Rows),
	}

	// set column widths
	colWidths := config.ColWidths
	if len(colWidths) == 0 {
		avgWidth := config.Width / config.Cols
		colWidths = make([]int, config.Cols)
		for i := range colWidths {
			colWidths[i] = avgWidth
		}
	} else if len(colWidths) != config.Cols {
		Error("nested table column width count does not match column count")
		return nil, NewValidationError("TableConfig.ColWidths", "", "column width count does not match column count")
	}

	// create table grid
	for _, width := range colWidths {
		nestedTable.Grid.Cols = append(nestedTable.Grid.Cols, TableGridCol{
			W: fmt.Sprintf("%d", width),
		})
	}

	// create table rows and cells
	for i := 0; i < config.Rows; i++ {
		tableRow := TableRow{
			Cells: make([]TableCell, 0, config.Cols),
		}

		for j := 0; j < config.Cols; j++ {
			tableCell := TableCell{
				Properties: &TableCellProperties{
					TableCellW: &TableCellW{
						W:    fmt.Sprintf("%d", colWidths[j]),
						Type: "dxa",
					},
					VAlign: &VAlign{
						Val: "center",
					},
				},
				Paragraphs: []Paragraph{
					{
						Runs: []Run{
							{
								Text: Text{Content: ""},
							},
						},
					},
				},
			}

			// set cell content if initial data is provided
			if config.Data != nil && i < len(config.Data) && j < len(config.Data[i]) {
				tableCell.Paragraphs[0].Runs[0].Text.Content = config.Data[i][j]
			}

			tableRow.Cells = append(tableRow.Cells, tableCell)
		}

		nestedTable.Rows = append(nestedTable.Rows, tableRow)
	}

	// add to the cell's nested table list
	cell.Tables = append(cell.Tables, *nestedTable)

	InfoMsgf(MsgNestedTableAddedToCell, row, col, config.Rows, config.Cols)
	return &cell.Tables[len(cell.Tables)-1], nil
}

// GetNestedTables returns all nested tables in a cell.
// Parameters:
//   - row: row index (0-based)
//   - col: column index (0-based)
//
// Returns:
//   - []Table: all nested tables in the cell
//   - error: an error if the index is invalid
func (t *Table) GetNestedTables(row, col int) ([]Table, error) {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	return cell.Tables, nil
}

// CellListConfig represents cell list configuration.
type CellListConfig struct {
	Type         ListType   // list type
	BulletSymbol BulletType // bullet symbol (for unordered lists only)
	Items        []string   // list item content
}

// AddCellList adds a list to a cell.
// Parameters:
//   - row: row index (0-based)
//   - col: column index (0-based)
//   - config: list configuration
//
// Returns:
//   - error: an error if the index is invalid
func (t *Table) AddCellList(row, col int, config *CellListConfig) error {
	cell, err := t.GetCell(row, col)
	if err != nil {
		return err
	}

	if config == nil || len(config.Items) == 0 {
		return NewValidationError("CellListConfig", "", "list configuration cannot be empty and must contain list items")
	}

	// determine prefix based on list type
	for i, item := range config.Items {
		var prefix string
		switch config.Type {
		case ListTypeBullet:
			// use bullet symbol
			bulletSymbol := config.BulletSymbol
			if bulletSymbol == "" {
				bulletSymbol = BulletTypeDot
			}
			prefix = string(bulletSymbol) + " "
		case ListTypeNumber, ListTypeDecimal:
			// use numeric numbering
			prefix = fmt.Sprintf("%d. ", i+1)
		case ListTypeLowerLetter:
			// use lowercase letters
			prefix = fmt.Sprintf("%c. ", 'a'+i)
		case ListTypeUpperLetter:
			// use uppercase letters
			prefix = fmt.Sprintf("%c. ", 'A'+i)
		case ListTypeLowerRoman:
			// use lowercase Roman numerals
			prefix = fmt.Sprintf("%s. ", toRomanLower(i+1))
		case ListTypeUpperRoman:
			// use uppercase Roman numerals
			prefix = fmt.Sprintf("%s. ", toRomanUpper(i+1))
		default:
			// default to bullet symbol
			prefix = string(BulletTypeDot) + " "
		}

		// create list item paragraph
		para := Paragraph{
			Runs: []Run{
				{
					Text: Text{
						Content: prefix + item,
						Space:   "preserve",
					},
				},
			},
		}

		// add to cell
		cell.Paragraphs = append(cell.Paragraphs, para)
	}

	InfoMsgf(MsgListAddedToCell, row, col, len(config.Items))
	return nil
}

// toRomanLower converts a number to lowercase Roman numerals.
func toRomanLower(num int) string {
	return strings.ToLower(toRomanUpper(num))
}

// toRomanUpper converts a number to uppercase Roman numerals.
func toRomanUpper(num int) string {
	values := []int{1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1}
	symbols := []string{"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"}

	if num <= 0 || num > 3999 {
		return fmt.Sprintf("%d", num)
	}

	result := ""
	for i, value := range values {
		for num >= value {
			result += symbols[i]
			num -= value
		}
	}
	return result
}

// CellImageConfig represents cell image configuration.
type CellImageConfig struct {
	// image source - file path
	FilePath string
	// image source - binary data
	Data []byte
	// image format (required when using Data)
	Format ImageFormat
	// image width (mm), 0 for auto
	Width float64
	// image height (mm), 0 for auto
	Height float64
	// whether to maintain aspect ratio
	KeepAspectRatio bool
	// image alt text
	AltText string
	// image title
	Title string
}
