package excelize

import (
	"fmt"
	"github.com/SPSZerone/sps-go-excel/excel"
)

func newSheet(excel *Excel, name string, index excel.SheetIndex) excel.Sheet {
	sheet := &Sheet{excel: excel, name: name, index: index}
	sheet.init()
	return sheet
}

type Sheet struct {
	excel   *Excel
	name    string
	index   excel.SheetIndex
	cellsCR map[string]map[excel.RowId]excel.Cell
	cellsRC map[excel.RowId]map[string]excel.Cell
}

func (s *Sheet) init() {
	s.cellsCR = make(map[string]map[excel.RowId]excel.Cell)
	s.cellsRC = make(map[excel.RowId]map[string]excel.Cell)
}

func (s *Sheet) GetExcel() excel.Excel {
	return s.excel
}

func (s *Sheet) SetName(name string) {
	s.name = name
}

func (s *Sheet) GetName() string {
	return s.name
}

func (s *Sheet) SetIndex(index excel.SheetIndex) {
	s.index = index
}

func (s *Sheet) GetIndex() excel.SheetIndex {
	return s.index
}

func (s *Sheet) SetRows(rows []excel.Row) {

}

func (s *Sheet) GetRows() []excel.Row {
	return nil
}

func (s *Sheet) SetCols(cols []excel.Col) {

}

func (s *Sheet) GetCols() []excel.Col {
	return nil
}

func (s *Sheet) SetCell(cell excel.Cell) error {
	if !s.excel.isWritable() {
		return fmt.Errorf("excel is not writable")
	}

	return nil
}

func (s *Sheet) GetCell(cellId excel.CellId) excel.Cell {
	return nil
}

func checkCellId(colName string, rowId excel.RowId) error {
	if colName == "" {
		return fmt.Errorf("col name is empty")
	}
	if rowId == 0 {
		return fmt.Errorf("row id is 0")
	}
	return nil
}

func (s *Sheet) getCell(colName string, rowId excel.RowId) (excel.Cell, error) {
	if err := checkCellId(colName, rowId); err != nil {
		return nil, err
	}

	cellRows, ok := s.cellsCR[colName]
	if !ok {
		s.cellsCR[colName] = make(map[excel.RowId]excel.Cell)
	}

	_, ok = s.cellsRC[rowId]
	if !ok {
		s.cellsRC[rowId] = make(map[string]excel.Cell)
	}

	cell, ok := cellRows[rowId]
	if !ok {
		cell = newCellCR(colName, rowId)
		s.cellsCR[colName][rowId] = cell
		s.cellsRC[rowId][colName] = cell
	}

	return cell, nil
}

func (s *Sheet) SetCellI(cellId excel.CellId, value any) error {
	if !s.excel.isWritable() {
		return fmt.Errorf("excel is not writable")
	}

	return nil
}

func (s *Sheet) SetCellCR(colName string, rowId excel.RowId, value any) error {
	if !s.excel.isWritable() {
		return fmt.Errorf("excel is not writable")
	}

	cell, err := s.getCell(colName, rowId)
	if err != nil {
		return err
	}

	excelFile := s.excel.excel
	cellId := getCellId(colName, rowId)

	err = excelFile.SetCellValue(s.name, cellId, value)
	return cell.SetValue(value)
}

func (s *Sheet) GetCellCR(colName string, rowId excel.RowId, opts ...excel.Option) (excel.Cell, error) {
	cell, err := s.getCell(colName, rowId)
	if err != nil {
		return nil, err
	}

	excelFile := s.excel.excel
	cellId := getCellId(colName, rowId)

	value, err := excelFile.GetCellValue(s.name, cellId)
	if err != nil {
		return nil, err
	}

	err = cell.SetValue(value)
	if err != nil {
		return nil, err
	}

	return cell, nil
}

type Row struct {
	id excel.RowId
}

func (r *Row) SetId(rowId excel.RowId) {
	r.id = rowId
}

func (r *Row) GetId() excel.RowId {
	return r.id
}

func (r *Row) SetCells(cells []excel.Cell) {

}

func (r *Row) GetCells() []excel.Cell {
	return nil
}

func (r *Row) SetCell(cell excel.Cell) {

}

func (r *Row) GetCell(colName string) excel.Cell {
	return nil
}

func (r *Row) SetCellCV(colName string, value any) {

}

type Column struct {
	name string
}

func (c *Column) SetName(name string) {
	c.name = name
}

func (c *Column) GetName() string {
	return c.name
}

func (c *Column) SetCells(cells []excel.Cell) {

}

func (c *Column) GetCells() []excel.Cell {
	return nil
}

func (c *Column) SetCell(cell excel.Cell) {

}

func (c *Column) GetCell(rowId excel.RowId) excel.Cell {
	return nil
}

func (c *Column) SetCellRV(rowId excel.RowId, value any) {

}

func newCellId(col string, row excel.RowId) excel.CellId {
	cellId := &CellId{col: col, row: row}
	return cellId
}

type CellId struct {
	col string
	row excel.RowId
}

func (i *CellId) SetCol(col string) {
	i.col = col
}

func (i *CellId) GetCol() string {
	return i.col
}

func (i *CellId) SetRow(rowId excel.RowId) {
	i.row = rowId
}

func (i *CellId) GetRow() excel.RowId {
	return i.row
}

func newCellCR(col string, row excel.RowId) excel.Cell {
	cellId := newCellId(col, row)
	cell := &Cell{id: cellId}
	return cell
}

func newCell(cellId excel.CellId) excel.Cell {
	cell := &Cell{id: cellId}
	return cell
}

type Cell struct {
	id    excel.CellId
	value any
}

func (c *Cell) SetId(cellId excel.CellId) {
	c.id = cellId
}

func (c *Cell) GetId() excel.CellId {
	return c.id
}

func (c *Cell) SetValue(value any) error {
	c.value = value
	return nil
}

func (c *Cell) GetValue() any {
	return c.value
}
