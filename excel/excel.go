package excel

import (
	"fmt"
	"io"
)

type NewerExcel func(options *Options) Excel

var defaultNewerExcel NewerExcel

func RegisterDefaultNewerExcel(newer NewerExcel) {
	defaultNewerExcel = newer
}

type NewerExcelKey int

var customNewerExcels = make(map[NewerExcelKey]NewerExcel)

func RegisterCustomNewerExcel(key NewerExcelKey, newer NewerExcel) error {
	_, ok := customNewerExcels[key]
	if ok {
		return fmt.Errorf("custom newer excel '%+v' already exists", key)
	}
	customNewerExcels[key] = newer
	return nil
}

func OpenFile(file string, opts ...Option) (Excel, error) {
	options := make([]Option, 0)
	options = append(options, OptFile(file), OptFlag(OReadWrite))
	options = append(options, opts...)
	e, err := NewExcel(options...)
	if err != nil {
		return nil, err
	}
	return e, nil
}

func NewFile(file string, opts ...Option) (Excel, error) {
	options := make([]Option, 0)
	options = append(options, OptFile(file), OptFlag(OReadWrite|OCreate|OExist))
	options = append(options, opts...)
	return NewExcel(options...)
}

func NewExcel(opts ...Option) (Excel, error) {
	if defaultNewerExcel == nil {
		return nil, fmt.Errorf("default newer excel not register")
	}
	return NewExcelCustom(defaultNewerExcel, opts...)
}

func NewExcelCustom(newer NewerExcel, opts ...Option) (Excel, error) {
	options := newOptions()
	e := newer(options)
	err := e.Init(opts...)
	if err != nil {
		return nil, err
	}
	return e, nil
}

type Excel interface {
	Init(opts ...Option) error
	Options() Options
	Close() error

	File() string

	Read(opts ...Option) (int64, error)
	ReadFrom(reader io.Reader) (int64, error)
	ReadFromO(reader io.Reader, opts ...Option) (int64, error)

	Write(opts ...Option) (int64, error)
	WriteTo(writer io.Writer) (int64, error)
	WriteToO(writer io.Writer, opts ...Option) (int64, error)
	WriteAs(file string, opts ...Option) (int64, error)

	Sheets() map[string]Sheet
	Sheet(name string) (Sheet, error)
	GetActiveSheet() Sheet
	SheetCreate(name string) (Sheet, error)
	SheetDelete(name string) error
}

type Formula interface {
}

type NumberFormat interface {
}
