package excel

import "github.com/xuri/excelize/v2"

type Excel struct {
	FileName string
	Sheets   []Sheet
}
type Sheet struct {
	SheetName string
	Headers   []Header
	Values    [][]interface{}
}

type Header struct {
	Width float64 //非必填
	Value string
	Style excelize.Style
}

var headerCell []string

func init() {
	for i := 'A'; i <= 'Z'; i++ {
		headerCell = append(headerCell, string(i))
	}
	for i := 'A'; i <= 'Z'; i++ {
		headerCell = append(headerCell, "A"+string(i))
	}
	for i := 'A'; i <= 'Z'; i++ {
		headerCell = append(headerCell, "B"+string(i))
	}
}

func GetCellFlag(i int) string {
	if i > len(headerCell) || i < 1 {
		return "unsupport"
	}
	return headerCell[i-1]
}
