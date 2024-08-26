package excel

type SheetIndex int
type RowId uint32

type Sheet interface {
	SetName(name string)
	GetName() string

	SetIndex(index SheetIndex)
	GetIndex() SheetIndex

	SetRows(rows []Row)
	GetRows() []Row

	SetCols(cols []Col)
	GetCols() []Col

	SetCell(cell Cell) error
	GetCell(cellId CellId) Cell

	SetCellI(cellId CellId, value any) error
	SetCellCR(colName string, rowId RowId, value any) error
	GetCellCR(colName string, rowId RowId, opts ...Option) (Cell, error)
}

type Row interface {
	SetId(rowId RowId)
	GetId() RowId

	SetCells(cells []Cell)
	GetCells() []Cell

	SetCell(cell Cell)
	GetCell(colName string) Cell

	SetCellCV(colName string, value any)
}

type Col interface {
	SetName(name string)
	GetName() string

	SetCells(cells []Cell)
	GetCells() []Cell

	SetCell(cell Cell)
	GetCell(rowId RowId) Cell

	SetCellRV(rowId RowId, value any)
}

type CellId interface {
	SetCol(col string)
	GetCol() string

	SetRow(rowId RowId)
	GetRow() RowId
}

type Cell interface {
	SetId(cellId CellId)
	GetId() CellId

	SetValue(value any) error
	GetValue() any
}
