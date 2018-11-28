package main

import (
	"go-excel/excel"
	"fmt"
	"flag"
)
func main()  {
	var excelFile = flag.String("file","../src/go-xlsx/Golden_China_Fund.xlsm","excelFile-path")
	var sheetName = flag.String("sheet","master","excel-sheet")
	flag.Parse()

	a := new(excel.Workbook)
	a.NewWorkBook(*excelFile)

	workSheet :=a.ChangeSheet(*sheetName)

	fmt.Printf("%s\n", workSheet.FormatJson())
}