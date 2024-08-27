package excelize

import (
	"github.com/xuri/excelize/v2"

	"github.com/SPSZerone/sps-go-excel/excel"
)

var cellIds = make(map[string]excel.CellId)

func getCellId(colName string, rowId excel.RowId) (excel.CellId, error) {
	cellName, err := joinCellName(colName, rowId)
	if err != nil {
		return nil, err
	}

	cellId, ok := cellIds[cellName]
	if ok {
		return cellId, nil
	}

	col, row, err := splitCellName(cellName)
	if err != nil {
		return nil, err
	}

	cellId = &CellId{col: col, row: row, name: cellName}
	cellIds[cellName] = cellId
	return cellId, nil
}

func joinCellName(colName string, rowId excel.RowId) (string, error) {
	return excelize.JoinCellName(colName, int(rowId))
}

func splitCellName(cellName string) (string, excel.RowId, error) {
	colName, rowId, err := excelize.SplitCellName(cellName)
	if err != nil {
		return "", 0, err
	}
	return colName, excel.RowId(rowId), nil
}
