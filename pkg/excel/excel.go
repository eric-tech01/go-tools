package excel

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func SetHeader() {

}

func (e Excel) Create() error {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	for _, sheet := range e.Sheets {
		index, err := f.NewSheet(sheet.SheetName)
		if err != nil {
			fmt.Println(err)
			return err
		}
		if err != nil {
			fmt.Println(err)
			return err
		}
		for i, header := range sheet.Headers {
			style, err := f.NewStyle(&header.Style)
			if err != nil {
				fmt.Println(err)
				return err
			}
			f.SetCellValue(sheet.SheetName, GetCellFlag(i+1)+"1", header.Value)
			if header.Width < 5 {
				header.Width = 15
			}
			f.SetColWidth(sheet.SheetName, GetCellFlag(i+1), GetCellFlag(i+1), header.Width)

			f.SetCellStyle(sheet.SheetName, GetCellFlag(i+1)+"1", GetCellFlag(i+1)+"1", style)
		}
		// Set active sheet of the workbook.
		f.SetActiveSheet(index)

		for r, row := range sheet.Values {
			for c, col := range row {
				f.SetCellValue(sheet.SheetName, GetCellFlag(c+1)+strconv.Itoa(r+2), col)
			}
		}
	}
	f.DeleteSheet("sheet1")
	if err := f.SaveAs(e.FileName); err != nil {
		fmt.Println(err)
		return err
	}
	return nil
}
