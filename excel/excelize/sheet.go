package excelize

import (
	"fmt"

	"github.com/xuri/excelize/v2"

	"github.com/SPSZerone/sps-go-excel/excel"
)

func newSheet(excel *Excel, name string, index excel.SheetIndex) *Sheet {
	sheet := &Sheet{excel: excel, name: name, index: index}
	sheet.init()
	return sheet
}

type Sheet struct {
	excel   *Excel
	name    string
	index   excel.SheetIndex
	cellsCR map[string]map[excel.RowId]*Cell
	cellsRC map[excel.RowId]map[string]*Cell
}

func (s *Sheet) init() {
	s.cellsCR = make(map[string]map[excel.RowId]*Cell)
	s.cellsRC = make(map[excel.RowId]map[string]*Cell)
}

func (s *Sheet) getExcelFile() *excelize.File {
	return s.excel.excel
}

func (s *Sheet) getCell(colNam string, rowId excel.RowId) (*Cell, error) {
	cellId, err := getCellId(colNam, rowId)
	if err != nil {
		return nil, err
	}
	col := cellId.Col()

	cellRows, ok := s.cellsCR[col]
	if !ok {
		s.cellsCR[col] = make(map[excel.RowId]*Cell)
	}

	_, ok = s.cellsRC[rowId]
	if !ok {
		s.cellsRC[rowId] = make(map[string]*Cell)
	}

	cell, ok := cellRows[rowId]
	if ok {
		return cell, nil
	}

	cellNew, err := newCellCR(col, rowId)
	if err != nil {
		return nil, err
	}

	s.cellsCR[col][rowId] = cellNew
	s.cellsRC[rowId][col] = cellNew
	return cellNew, nil
}

func (s *Sheet) IsWritable() bool {
	return s.excel.IsWritable()
}

func (s *Sheet) Name() string {
	return s.name
}

func (s *Sheet) Index() excel.SheetIndex {
	return s.index
}

func (s *Sheet) SetRows(rows []excel.Row) error {
	if !s.IsWritable() {
		return fmt.Errorf("excel is not writable")
	}
	return fmt.Errorf("Sheet.SetRows not implemented yet")
}

func (s *Sheet) GetRows(opts ...excel.Option) ([]excel.Row, error) {
	excelFile := s.getExcelFile()

	rows, err := excelFile.GetRows(s.name)
	if err != nil {
		return nil, err
	}

	eRows := make([]excel.Row, 0)
	for i, row := range rows {
		rowId := excel.RowId(i + 1)
		eRow, errNew := newRowData(rowId, row)
		if errNew != nil {
			return nil, errNew
		}
		eRows = append(eRows, eRow)
	}

	return eRows, nil
}

func (s *Sheet) SetCols(cols []excel.Col) error {
	if !s.IsWritable() {
		return fmt.Errorf("excel is not writable")
	}
	return fmt.Errorf("Sheet.SetCols not implemented yet")
}

func (s *Sheet) GetCols(opts ...excel.Option) ([]excel.Col, error) {
	excelFile := s.getExcelFile()

	cols, err := excelFile.GetCols(s.name)
	if err != nil {
		return nil, err
	}

	eCols := make([]excel.Col, 0)
	for i, col := range cols {
		colName, errColName := excelize.ColumnNumberToName(i + 1)
		if errColName != nil {
			return nil, errColName
		}
		eCol, errNew := newColData(colName, col)
		if errNew != nil {
			return nil, errNew
		}
		eCols = append(eCols, eCol)
	}

	return eCols, nil
}

func (s *Sheet) SetCell(cell excel.Cell) error {
	return s.SetCellI(cell.Id(), cell.Value())
}

func (s *Sheet) GetCell(cellId excel.CellId, opts ...excel.Option) (excel.Cell, error) {
	return s.GetCellCR(cellId.Col(), cellId.Row(), opts...)
}

func (s *Sheet) SetCellI(cellId excel.CellId, value any) error {
	return s.SetCellCR(cellId.Col(), cellId.Row(), value)
}

func (s *Sheet) SetCellCR(colName string, rowId excel.RowId, value any) error {
	if !s.IsWritable() {
		return fmt.Errorf("excel is not writable")
	}

	cell, err := s.getCell(colName, rowId)
	if err != nil {
		return err
	}

	excelFile := s.getExcelFile()
	cellName, err := joinCellName(colName, rowId)
	if err != nil {
		return err
	}

	err = excelFile.SetCellValue(s.name, cellName, value)
	if err != nil {
		return nil
	}
	return cell.SetValue(value)
}

func (s *Sheet) GetCellCR(colName string, rowId excel.RowId, opts ...excel.Option) (excel.Cell, error) {
	cell, err := s.getCell(colName, rowId)
	if err != nil {
		return nil, err
	}

	excelFile := s.getExcelFile()
	cellName, err := joinCellName(colName, rowId)
	if err != nil {
		return nil, err
	}

	value, err := excelFile.GetCellValue(s.name, cellName)
	if err != nil {
		return nil, err
	}

	err = cell.SetValue(value)
	if err != nil {
		return nil, err
	}

	return cell, nil
}

func newRowData(rowId excel.RowId, data []string) (*Row, error) {
	row := newRow(rowId)
	err := row.SetCellsS(data)
	if err != nil {
		return nil, err
	}
	return row, nil
}

func newRow(rowId excel.RowId) *Row {
	row := &Row{id: rowId}
	row.init()
	return row
}

type Row struct {
	id    excel.RowId
	cols  []string
	cells map[string]excel.Cell
}

func (r *Row) init() {
	r.cols = make([]string, 0)
	r.cells = make(map[string]excel.Cell, 0)
}

func (r *Row) Id() excel.RowId {
	return r.id
}

func (r *Row) SetCells(cells []excel.Cell) error {
	return nil
}

func (r *Row) Cells(opts ...excel.Option) ([]excel.Cell, error) {
	cells := make([]excel.Cell, 0)
	for _, col := range r.cols {
		cell := r.cells[col]
		cells = append(cells, cell)
	}
	return cells, nil
}

func (r *Row) SetCellsS(data []string) error {
	for i, value := range data {
		num := i + 1
		col, err := excelize.ColumnNumberToName(num)
		if err != nil {
			return err
		}

		err = r.SetCellC(col, value)
		if err != nil {
			return err
		}
	}

	return nil
}

func (r *Row) SetCell(cell excel.Cell) error {
	if cell == nil {
		return fmt.Errorf("cell is nil")
	}
	colName := cell.Id().Col()
	r.cols = append(r.cols, colName)
	r.cells[colName] = cell
	return nil
}

func (r *Row) Cell(colName string, opts ...excel.Option) (excel.Cell, error) {
	cell, ok := r.cells[colName]
	if !ok {
		return nil, fmt.Errorf("col '%s' not exist", colName)
	}
	return cell, nil
}

func (r *Row) SetCellC(colName string, value any) error {
	// cell id
	rowId := r.id
	cellId, err := getCellId(colName, rowId)
	if err != nil {
		return err
	}

	// cell
	cell := newCell(cellId)
	err = cell.SetValue(value)
	if err != nil {
		return err
	}

	return r.SetCell(cell)
}

func newColData(colName string, data []string) (*Col, error) {
	col := newCol(colName)
	err := col.SetCellsS(data)
	if err != nil {
		return nil, err
	}
	return col, nil
}

func newCol(colName string) *Col {
	col := &Col{name: colName}
	col.init()
	return col
}

type Col struct {
	name  string
	rows  []excel.RowId
	cells map[excel.RowId]excel.Cell
}

func (c *Col) init() {
	c.rows = make([]excel.RowId, 0)
	c.cells = make(map[excel.RowId]excel.Cell, 0)
}

func (c *Col) Name() string {
	return c.name
}

func (c *Col) SetCells(cells []excel.Cell) error {
	return nil
}

func (c *Col) Cells(opts ...excel.Option) ([]excel.Cell, error) {
	cells := make([]excel.Cell, 0)
	for _, row := range c.rows {
		cell := c.cells[row]
		cells = append(cells, cell)
	}
	return cells, nil
}

func (c *Col) SetCellsS(data []string) error {
	for i, value := range data {
		rowId := excel.RowId(i + 1)

		err := c.SetCellR(rowId, value)
		if err != nil {
			return err
		}
	}

	return nil
}

func (c *Col) SetCell(cell excel.Cell) error {
	if cell == nil {
		return fmt.Errorf("cell is nil")
	}
	rowId := cell.Id().Row()
	c.rows = append(c.rows, rowId)
	c.cells[rowId] = cell
	return nil
}

func (c *Col) Cell(rowId excel.RowId, opts ...excel.Option) (excel.Cell, error) {
	cell, ok := c.cells[rowId]
	if !ok {
		return nil, fmt.Errorf("row '%d' not exist", rowId)
	}
	return cell, nil
}

func (c *Col) SetCellR(rowId excel.RowId, value any) error {
	// cell id
	colName := c.name
	cellId, err := getCellId(colName, rowId)
	if err != nil {
		return err
	}

	// cell
	cell := newCell(cellId)
	err = cell.SetValue(value)
	if err != nil {
		return err
	}

	return c.SetCell(cell)
}

type CellId struct {
	col  string
	row  excel.RowId
	name string
}

func (i *CellId) Col() string {
	return i.col
}

func (i *CellId) Row() excel.RowId {
	return i.row
}

func (i *CellId) Name() string {
	return i.name
}

func newCellCR(col string, row excel.RowId) (*Cell, error) {
	cellId, err := getCellId(col, row)
	if err != nil {
		return nil, err
	}
	cell := newCell(cellId)
	return cell, nil
}

func newCell(cellId excel.CellId) *Cell {
	cell := &Cell{id: cellId}
	return cell
}

type Cell struct {
	id    excel.CellId
	value any
}

func (c *Cell) Id() excel.CellId {
	return c.id
}

func (c *Cell) SetValue(value any) error {
	c.value = value
	return nil
}

func (c *Cell) Value() any {
	return c.value
}

func (c *Cell) String() string {
	return fmt.Sprintf("「%s」=>『%+v』", c.id.Name(), c.value)
}
