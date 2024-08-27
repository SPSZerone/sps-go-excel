package excelize

import (
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"

	spsos "github.com/SPSZerone/sps-go-zerone/os"

	"github.com/SPSZerone/sps-go-excel/excel"
)

func init() {
	newer := func(options *excel.Options) excel.Excel {
		return &Excel{options: options}
	}
	excel.RegisterDefaultNewerExcel(newer)
}

type Excel struct {
	options *excel.Options

	sheets map[string]excel.Sheet
	excel  *excelize.File
}

func (e *Excel) updateOptions(opts ...excel.Option) {
	for _, o := range opts {
		o(e.options)
	}
}

func (e *Excel) isWritable() bool {
	return e.options.Flag.IsWritable()
}

func (e *Excel) newSheet(name string, index excel.SheetIndex) excel.Sheet {
	sheet, ok := e.sheets[name]
	if ok {
		return sheet
	}
	sheetNew := newSheet(e, name, index)
	e.sheets[name] = sheetNew
	return sheetNew
}

func (e *Excel) Init(opts ...excel.Option) error {
	if e.sheets != nil {
		return fmt.Errorf("excel is already initialized")
	}

	e.updateOptions(opts...)

	if err := e.initFile(); err != nil {
		return err
	}

	e.sheets = make(map[string]excel.Sheet)
	return nil
}

func (e *Excel) initFile() error {
	file := e.options.File
	flag := e.options.Flag

	if flag.IsCreate() {
		if flag.IsExist() && spsos.FileExist(file) {
			return fmt.Errorf("cannot create file %s already exists", file)
		}
		e.excel = excelize.NewFile()
	} else {
		excelFile, err := excelize.OpenFile(file)
		if err != nil {
			return err
		}
		e.excel = excelFile
	}

	return nil
}

func (e *Excel) Options() excel.Options {
	return *e.options
}

func (e *Excel) Close() error {
	if e.excel == nil {
		return nil
	}
	if err := e.excel.Close(); err != nil {
		return err
	}
	return nil
}

func (e *Excel) File() string {
	return e.options.File
}

func (e *Excel) Read(opts ...excel.Option) (int64, error) {
	excelFile := e.excel
	sheetNames := excelFile.GetSheetList()
	for _, name := range sheetNames {
		index, err := excelFile.GetSheetIndex(name)
		if err != nil {
			return 0, err
		}
		e.newSheet(name, excel.SheetIndex(index))
	}
	return 0, nil
}

func (e *Excel) ReadFrom(reader io.Reader) (int64, error) {
	return 0, fmt.Errorf("Excel.ReadFrom not implemented yet")
}

func (e *Excel) ReadFromO(reader io.Reader, opts ...excel.Option) (int64, error) {
	return 0, fmt.Errorf("Excel.ReadFromO not implemented yet")
}

func (e *Excel) write(opts ...excel.Option) error {
	if !e.isWritable() {
		return fmt.Errorf("excel is not writable")
	}

	excelFile := e.excel
	for _, sheet := range e.sheets {
		index := int(sheet.GetIndex())
		excelFile.SetActiveSheet(index)
	}

	return nil
}

func (e *Excel) Write(opts ...excel.Option) (int64, error) {
	if err := e.write(opts...); err != nil {
		return 0, err
	}

	file := e.options.File
	err := e.excel.SaveAs(file)
	if err != nil {
		return 0, err
	}

	return 0, nil
}

func (e *Excel) WriteTo(writer io.Writer) (int64, error) {
	return e.WriteToO(writer)
}

func (e *Excel) WriteToO(writer io.Writer, opts ...excel.Option) (int64, error) {
	if err := e.write(opts...); err != nil {
		return 0, err
	}

	return e.excel.WriteTo(writer)
}

func (e *Excel) WriteAs(file string, opts ...excel.Option) (int64, error) {
	if err := e.write(opts...); err != nil {
		return 0, err
	}

	err := e.excel.SaveAs(file)
	if err != nil {
		return 0, err
	}

	return 0, nil
}

func (e *Excel) Sheets() map[string]excel.Sheet {
	return e.sheets
}

func (e *Excel) Sheet(name string) (excel.Sheet, error) {
	sheet, ok := e.sheets[name]
	if !ok {
		return nil, fmt.Errorf("sheet '%s' not exists", name)
	}
	return sheet, nil
}

func (e *Excel) GetActiveSheet() excel.Sheet {
	index := e.excel.GetActiveSheetIndex()
	name := e.excel.GetSheetName(index)

	sheet, ok := e.sheets[name]
	if ok {
		return sheet
	}

	sheetNew := e.newSheet(name, excel.SheetIndex(index))
	return sheetNew
}

func (e *Excel) SheetCreate(name string) (excel.Sheet, error) {
	sheet, ok := e.sheets[name]
	if ok {
		return sheet, fmt.Errorf("sheet '%s' already exists", name)
	}

	index, err := e.excel.NewSheet(name)
	if err != nil {
		return nil, err
	}

	sheetNew := e.newSheet(name, excel.SheetIndex(index))
	return sheetNew, nil
}

func (e *Excel) SheetDelete(name string) error {
	_, ok := e.sheets[name]
	if !ok {
		return fmt.Errorf("sheet '%s' not exist", name)
	}

	err := e.excel.DeleteSheet(name)
	if err != nil {
		return err
	}

	delete(e.sheets, name)
	return nil
}

type Formula struct {
}

type NumberFormat struct {
}
