package excel

type SheetIndex int
type RowId uint32

type Sheet interface {
	Name() string
	Index() SheetIndex

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
	Id() RowId
	Cells(opts ...Option) ([]Cell, error)
	Cell(colName string, opts ...Option) (Cell, error)
}

type Col interface {
	Name() string
	Cells(opts ...Option) ([]Cell, error)
	Cell(rowId RowId, opts ...Option) (Cell, error)
}

type CellId interface {
	Col() string
	Row() RowId
	Name() string
}

type Cell interface {
	Id() CellId
	Value() any
	String() string
}
