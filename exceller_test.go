package exceller

import (
	"fmt"
	"testing"
)

func TestBuildReport(t *testing.T) {
	var err error
	sheetName := "Bio"
	sheetHeader := []string{
		"No",
		"Name",
		"Age",
	}
	sheetData := [][]string{
		{"1", "John", "20"},
		{"2", "Ken", "20"},
		{"3", "Yuri", "20"},
	}

	report := NewExcelReport()
	report.AddSheet(sheetName).AddHeader(sheetHeader).AddBody(sheetData)
	err = report.BuildAndExport()
	if err != nil {
		t.Error(err)
	}
}

func TestBuildReportEmpty(t *testing.T) {
	var err error

	report := NewExcelReport()
	err = report.BuildAndExport()
	if err != nil {
		t.Error(err)
	}
}

func TestBuildReport_HeaderBeyondZ(t *testing.T) {
	var (
		err         error
		sheetHeader []string
		sheetRow    []string
		sheetBody   [][]string
	)

	for i := 0; i < 40; i++ {
		head := fmt.Sprintf("Header%d", i)
		sheetHeader = append(sheetHeader, head)
	}

	for i := 0; i < 30; i++ {
		for j := 0; j < 40; j++ {
			row := fmt.Sprintf("Row%d", j)
			sheetRow = append(sheetRow, row)
		}
		sheetBody = append(sheetBody, sheetRow)

		sheetRow = []string{}
	}

	report := NewExcelReport()
	report.AddSheet("Headers").AddHeader(sheetHeader).AddBody(sheetBody)
	err = report.BuildAndExport()
	if err != nil {
		t.Error(err)
	}
}
