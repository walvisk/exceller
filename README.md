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
	report := exceller.NewExcelReport()

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

It will produce file excel below
![Example](./example.png?raw=true "Example Report")
### Todo

- [x] Give screenshoot of excel on writing report
- [ ] Error Handling
- [ ] Provide test
- [ ] Make sure file is close
- [ ] Support header more than 1 row ?
- [x] Support header more than Z column
- [ ] BuilAndExport() can accept path as exported path
