package main

import (
	"fmt"

	excel "github.com/eric01-tech/go-tools/pkg/excel"
	"github.com/xuri/excelize/v2"
)

type Person struct {
	Name string `json:"name"`
	Age  int    `json:"age"`
}

func main() {
	fileName := "Boox_01.xlsx"

	e := excel.Excel{FileName: fileName}
	firstSheet := excel.Sheet{SheetName: "物料清单"}

	//定义header
	var headers []excel.Header
	headers = append(headers, excel.Header{Value: "名称",
		Width: 40,
		Style: excelize.Style{
			Font: &excelize.Font{Bold: true, Size: 16},
		}})
	headers = append(headers, excel.Header{Value: "年龄"})
	firstSheet.Headers = headers

	//定义数据值
	firstSheet.Values = [][]interface{}{{"张三", 11}, {"王五", 33}}

	e.Sheets = append(e.Sheets, firstSheet)

	//生成excel
	if err := e.Create(); err != nil {
		fmt.Println(err)
	}
}
