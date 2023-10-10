# Exceller

## Introduction

Exceller is library written in Go providing abstraction level for writing simple excel report based on [Excelize](https://github.com/qax-os/excelize)

## Usage

### Installation

```bash
go get github.com/walvisk/exceller
```

### Writing Excel
You need to prepare sheet, header for the sheet, and data for the sheet.

```go
package main

func main() {
	f := excelize.NewFile()
	report := NewExcelReport(f)
	defer report.Close()

	sheetName := "Bio"
	sheetHeader := []string{
		"No",
		"Name",
		"Age",
	}
	sheetData := [][]any{
		{"1", "John", 20},
		{"2", "Ken", "20"},
		{"3", "Yuri", "20"},
	}

	report.AddSheet(sheetName).AddHeader(sheetHeader).AddBody(sheetData)
	report.Build()
}
```

It will produce file excel below
![Example](./example.png?raw=true "Example Report")
