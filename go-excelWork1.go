package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"flag"
	"strconv"
)

func check(e error) {
    if e != nil {
        panic(e)
    }
}

//example call should be excelWork -excelFile=file path -outputFolder=folder path - sheetToUse =sheet number 12 18 20
func main() {
	excelFile := flag.String("excelFile", "example.xlsx", "Excel file name with path")
	outputFolder := flag.String("outputFolder", "C:\\output\text\\", "Path to outputfolder with trailing folder file separator")
	sheetToUse := flag.Int("sheetToUse", 1, "the sheet number to use, 1 for first sheet")
	column := flag.Args()
	
	flag.Parse()
	
	xlsxFile, err := xlsx.OpenFile(*excelFile)
	check(err)

	sheet := xlsxFile.Sheets[*sheetToUse]
	fmt.Printf("here2")
	for i, row := range sheet.Rows {
	fmt.Printf("here3")
		if i >= 2 {
			if len(column) > 0 {
				for _, columnNo := range column {
					fileName := fmt.Sprintf("%s%d%s",*outputFolder, i, ".txt")
					fmt.Printf("%s", fileName)
					f, err := os.Create(fileName)
					check(err)
					number, err := strconv.Atoi(columnNo)
					check(err)
					value := row.Cells[number].String()
					f.WriteString(value)
					f.WriteString("\n")
				}
			} else {
				fileName := fmt.Sprintf("%s%d%s",*outputFolder, i, ".txt")
				fmt.Printf("%s", fileName)
				f, err := os.Create(fileName)
				check(err)
				value := row.Cells[1].String()
				f.WriteString(value)
				f.WriteString("\n")
			}
		}
	}
}