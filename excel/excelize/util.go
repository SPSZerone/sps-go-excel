package excelize

import (
	"fmt"
	"github.com/SPSZerone/sps-go-excel/excel"
)

func getCellId(colName string, rowId excel.RowId) string {
	cellId := fmt.Sprintf("%s%d", colName, rowId)
	return cellId
}
