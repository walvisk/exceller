package exceller

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

type sheet struct {
	name   string
	header []string
	body   [][]string
}

type Report struct {
	File   *excelize.File
	sheets []*sheet
}

func NewExcelReport() *Report {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	return &Report{
		File: f,
	}
}

func (rb *Report) AddSheet(sheetname string) *sheet {
	sheet := &sheet{name: sheetname}
	rb.sheets = append(rb.sheets, sheet)
	if len(rb.sheets) == 1 {
		rb.File.SetSheetName(rb.File.GetSheetName(0), sheetname)
	} else {
		rb.File.NewSheet(sheetname)
	}

	return sheet
}

func (sheet *sheet) AddHeader(headers []string) *sheet {
	sheet.header = headers

	return sheet
}

func (sheet *sheet) AddBody(data [][]string) *sheet {
	sheet.body = data

	return sheet
}

func (rb *Report) Build() {
	rb.write()
}

func (rb *Report) BuildAndExport() {
	rb.write()

	if err := rb.File.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func (rb *Report) write() {
	for _, sheet := range rb.sheets {
		x := 'A'
		y := 1

		headers := sheet.header
		for _, header := range headers {
			cell := fmt.Sprintf("%c%d", x, y)
			rb.File.SetCellValue(sheet.name, cell, header)
			x = x + 1
		}

		y = y + 1
		x = 'A'

		for _, outer := range sheet.body {
			for _, inner := range outer {
				cell := fmt.Sprintf("%c%d", x, y)
				rb.File.SetCellValue(sheet.name, cell, inner)
				x = x + 1
			}

			y = y + 1
			x = 'A'
		}
	}
}
