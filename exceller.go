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

	return &Report{
		File: f,
	}
}

func (rb *Report) Close() error {
	if err := rb.File.Close(); err != nil {
		fmt.Println(err)
		return err
	}

	return nil
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

// Build() provides writing to excelize file
func (rb *Report) Build() {
	rb.write()
}

// BuildAndExport provides writing the excelize to local file for debugging purposes
func (rb *Report) BuildAndExport() error {
	rb.write()

	if err := rb.File.SaveAs("Debug.xlsx"); err != nil {
		return err
	}

	return nil
}

// write provides the writing process to exelize file
func (rb *Report) write() {
	for _, sheet := range rb.sheets {
		headers := sheet.header
		y := 1
		for i, header := range headers {
			x := rb.generateColumnLetter(i + 1)
			cell := fmt.Sprintf("%s%d", x, y)

			rb.File.SetCellValue(sheet.name, cell, header)
		}

		if len(headers) > 0 {
			y = y + 1
		}

		for _, outer := range sheet.body {
			for i, inner := range outer {
				x := rb.generateColumnLetter(i + 1)
				cell := fmt.Sprintf("%s%d", x, y)

				rb.File.SetCellValue(sheet.name, cell, inner)
			}
			y = y + 1
		}
	}
}

func (rb *Report) generateColumnLetter(n int) string {
	if n <= 0 {
		return ""
	}

	column := ""
	for n > 0 {
		n--
		column = fmt.Sprintf("%c%s", 'A'+n%26, column)
		n /= 26
	}

	return column
}
