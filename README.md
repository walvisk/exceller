# Exceller

## Introduction

Exceller is library written in Go providing abstraction level for writing excel report based on [Excelize](https://github.com/qax-os/excelize)

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
	report := builder.NewExcelReport()

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

	report.AddSheet(sheetName).AddHeader(sheetHeader).AddBody(sheetData)
	report.Build()
}
```
### Todo

- [ ] Give screenshoot of excel on writing report
- [ ] Error Handling
- [ ] Provide test
