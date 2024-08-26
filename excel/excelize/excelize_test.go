package excelize

import (
	"fmt"
	"testing"

	"github.com/SPSZerone/sps-go-excel/excel"
	"github.com/stretchr/testify/require"
)

const (
	testFile = "TestFile.xlsx"
)

func TestNewFile(t *testing.T) {
	file := "NewFile.xlsx"
	e, err := excel.NewFile(file)
	require.NoError(t, err)
	require.NotNil(t, e)

	defer func(e excel.Excel) {
		err := e.Close()
		require.Nil(t, err)
	}(e)

	require.Equal(t, file, e.File())
}

func TestCreateSheet(t *testing.T) {
	e, errNew := excel.NewFile("CreateSheet.xlsx")
	require.NoError(t, errNew)
	require.NotNil(t, e)

	defer func(e excel.Excel) {
		err := e.Close()
		require.Nil(t, err)
	}(e)

	var err error

	sheetName := "sheet-1-0"
	_, err = e.SheetCreate(fmt.Sprintf(sheetName))
	require.NoError(t, err, "SheetCreate fail:%+v", err)
	_, err = e.SheetCreate(fmt.Sprintf(sheetName))
	require.Error(t, err)

	for i := 0; i < 3; i++ {
		_, err = e.SheetCreate(fmt.Sprintf("sheet-2-%d", i))
		require.NoError(t, err)
	}

	sheet := e.Sheets()
	require.Len(t, sheet, 4)
}

func TestWriteFile(t *testing.T) {
	e, errNew := excel.NewFile(testFile, excel.OptFlag(excel.OReadWrite|excel.OCreate))
	require.NoError(t, errNew)
	require.NotNil(t, e)

	defer func(e excel.Excel) {
		err := e.Close()
		require.Nil(t, err)
	}(e)

	var err error

	testSheet := func(sheet excel.Sheet) {
		name := sheet.GetName()
		err = sheet.SetCellCR("A", 1, fmt.Sprintf("%s: Hello, A1.", name))
		require.NoError(t, err, "SetCellCR fail:%+v", err)
		err = sheet.SetCellCR("B", 2, fmt.Sprintf("%s: Hello, B2.", name))
		require.NoError(t, err, "SetCellCR fail:%+v", err)
		err = sheet.SetCellCR("C", 3, fmt.Sprintf("%s: Hello, C3.", name))
		require.NoError(t, err, "SetCellCR fail:%+v", err)
	}

	defaultSheet := e.GetActiveSheet()
	testSheet(defaultSheet)

	// default Sheet1
	for i := 1; i <= 3; i++ {
		sheetName := fmt.Sprintf("My-Sheet-%d", i)
		sheet, err := e.SheetCreate(sheetName)
		require.NoError(t, err, "SheetCreate fail:%+v", err)
		require.NotNil(t, sheet)
		testSheet(sheet)
	}

	_, err = e.Write()
	require.NoError(t, err, "Write fail:%+v", err)

	_, err = e.WriteAs("TestFileAs.xlsx")
	require.NoError(t, err, "Write fail:%+v", err)
}

func TestReadFile(t *testing.T) {
	e, errOpen := excel.OpenFile(testFile, excel.OptFlag(excel.OReadOnly))
	require.NoError(t, errOpen)
	require.NotNil(t, e)

	defer func(e excel.Excel) {
		err := e.Close()
		require.Nil(t, err)
	}(e)

	_, errRead := e.Read()
	require.NoError(t, errRead)

	testSheet := func(sheet excel.Sheet) {
		t.Logf("Sheet:%+v ==================================================", sheet.GetName())
		cell, err := sheet.GetCellCR("A", 1)
		require.NoError(t, err, "SetCellCR fail:%+v", err)
		require.NotNil(t, cell)
		t.Logf("Col:%+v Row:%+v Val:%+v", cell.GetId().GetCol(), cell.GetId().GetRow(), cell.GetValue())
		cell, err = sheet.GetCellCR("B", 2)
		require.NoError(t, err, "SetCellCR fail:%+v", err)
		require.NotNil(t, cell)
		t.Logf("Col:%+v Row:%+v Val:%+v", cell.GetId().GetCol(), cell.GetId().GetRow(), cell.GetValue())
		cell, err = sheet.GetCellCR("C", 3)
		require.NoError(t, err, "SetCellCR fail:%+v", err)
		require.NotNil(t, cell)
		t.Logf("Col:%+v Row:%+v Val:%+v", cell.GetId().GetCol(), cell.GetId().GetRow(), cell.GetValue())
	}

	defaultSheet := e.GetActiveSheet()
	testSheet(defaultSheet)

	err := defaultSheet.SetCellCR("A", 1, fmt.Sprintf("%s: Hello, A1.", defaultSheet.GetName()))
	require.Error(t, err, "SetCellCR fail:%+v", err)

	for i := 1; i <= 3; i++ {
		sheetName := fmt.Sprintf("My-Sheet-%d", i)
		sheet, err := e.Sheet(sheetName)
		require.NoError(t, err, "Sheet fail:%+v", err)
		testSheet(sheet)
	}
}
