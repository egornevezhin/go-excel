package main

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
)

type ExcelCreator struct {
	file          *excelize.File
	ReportName    string
	columnNames   []string
	currentRow    int // current row number
	currentColumn int // current column number
	Data          interface{}
}

func (e *ExcelCreator) Init() (*ExcelCreator, error) {
	if reflect.TypeOf(e.Data).Kind() != reflect.Slice {
		return nil, errors.New("data must be a slice")
	}

	e.file = excelize.NewFile()
	sheet := e.file.NewSheet(e.ReportName)
	e.file.SetActiveSheet(sheet)
	e.currentRow = 1

	if err := e.initHeader(); err != nil {
		return nil, err
	}

	if err := e.writeData(); err != nil {
		return nil, err
	}

	return e, nil
}

func (e *ExcelCreator) Save(filename string) error {
	if err := e.file.SaveAs(filename); err != nil {
		return err
	}
	return nil
}

func (e *ExcelCreator) initHeader() error {
	row := reflect.TypeOf(e.Data).Elem()

	e.columnNames = createColumnNames(row.NumField())

	currentColumn := 0
	for i := 0; i < row.NumField(); i++ {

		address := getCellAddress(e.columnNames[currentColumn], e.currentRow)
		err := e.file.SetCellValue(e.ReportName, address, row.Field(i).Tag.Get("excel"))
		currentColumn += 1

		if err != nil {
			return err
		}
	}
	e.currentRow += 1
	return nil
}

func (e *ExcelCreator) writeData() error {

	// iterate through rows
	for j := 0; j < reflect.ValueOf(e.Data).Len(); j++ {

		//for _, row := range e.Data.([]interface{}) {
		currentColumn := 0

		fields := reflect.ValueOf(e.Data).Index(j)

		for i := 0; i < fields.NumField(); i++ {
			address := getCellAddress(e.columnNames[currentColumn], e.currentRow)
			err := e.file.SetCellValue(e.ReportName, address, fields.Field(i))
			currentColumn += 1

			if err != nil {
				return err
			}
		}
		e.currentRow += 1
	}
	return nil
}

func intToLetters(number int32) (letters string) {
	if firstLetter := number / 26; firstLetter > 0 {
		letters += intToLetters(firstLetter)
		letters += string('A' + number%26)
	} else {
		letters += string('A' + number)
	}

	return
}

func createColumnNames(numberOfColumns int) []string {
	out := make([]string, numberOfColumns)
	for i := 0; i < numberOfColumns; i++ {
		out[i] = intToLetters(int32(i))
	}
	return out
}

func getCellAddress(row string, column int) string {
	return fmt.Sprintf("%s%d", row, column)
}
