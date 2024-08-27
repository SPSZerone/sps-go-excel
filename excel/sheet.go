package excel

type SheetIndex int
type RowId uint32

type Sheet interface {
	SetName(name string)
	GetName() string

	SetIndex(index SheetIndex)
	GetIndex() SheetIndex

	SetRows(rows []Row) error
	GetRows(opts ...Option) ([]Row, error)

	SetCols(cols []Col) error
	GetCols(opts ...Option) ([]Col, error)

	SetCell(cell Cell) error
	GetCell(cellId CellId, opts ...Option) (Cell, error)

	SetCellI(cellId CellId, value any) error
	SetCellCR(colName string, rowId RowId, value any) error
	GetCellCR(colName string, rowId RowId, opts ...Option) (Cell, error)
}

type Row interface {
	SetId(rowId RowId)
	GetId() RowId

	SetCells(cells []Cell) error
	GetCells(opts ...Option) ([]Cell, error)
	SetCellsS(cells []string) error

	SetCell(cell Cell) error
	GetCell(colName string, opts ...Option) (Cell, error)

	SetCellC(colName string, value any) error
}

type Col interface {
	SetName(name string)
	GetName() string

	SetCells(cells []Cell) error
	GetCells(opts ...Option) ([]Cell, error)
	SetCellsS(cells []string) error

	SetCell(cell Cell) error
	GetCell(rowId RowId, opts ...Option) (Cell, error)

	SetCellR(rowId RowId, value any) error
}

type CellId interface {
	Col() string
	Row() RowId
	Name() string
}

type Cell interface {
	SetId(cellId CellId)
	GetId() CellId

	SetValue(value any) error
	GetValue() any

	String() string
}
