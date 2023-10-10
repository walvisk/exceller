package exceller

import (
	"fmt"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestBuildReport(t *testing.T) {
	var err error
	sheetName := "Bio"
	sheetHeader := []string{
		"No",
		"Name",
		"Income",
	}
	sheetData := [][]any{
		{"1", "John", 1000},
		{"2", "Ken", float64(2000.35)},
		{"3", "Yuri", "1000"},
	}

	f := excelize.NewFile()
	report := NewExcelReport(f)
	defer report.Close()
	report.AddSheet(sheetName).AddHeader(sheetHeader).AddBody(sheetData)
	err = report.BuildAndExport()
	if err != nil {
		t.Error(err)
	}
}

func TestBuildReportEmpty(t *testing.T) {
	var err error
	f := excelize.NewFile()
	report := NewExcelReport(f)
	defer report.Close()
	err = report.BuildAndExport()
	if err != nil {
		t.Error(err)
	}
}

func TestBuildReport_HeaderBeyondZ(t *testing.T) {
	var (
		err         error
		sheetHeader []string
		sheetRow    []any
		sheetBody   [][]any
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

		sheetRow = []any{}
	}
	f := excelize.NewFile()
	report := NewExcelReport(f)
	defer report.Close()
	report.AddSheet("Headers").AddHeader(sheetHeader).AddBody(sheetBody)
	err = report.BuildAndExport()
	if err != nil {
		t.Error(err)
	}
}

func TestBuildReport_JustBody(t *testing.T) {
	var err error

	sheetData := [][]any{
		{"1", "John", "20"},
		{"2", "Ken", "20"},
		{"3", "Yuri", "20"},
	}
	f := excelize.NewFile()
	report := NewExcelReport(f)
	defer report.Close()
	report.AddSheet("Sheet 1").AddBody(sheetData)
	err = report.BuildAndExport()
	if err != nil {
		t.Error(err)
	}
}
