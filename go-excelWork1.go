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

//example call should be go-excelWork1 -excelFile=filePath -outputFolder=folderPath - sheetToUse=sheetNO 12 18 20
//12 18 20 are optional and represent the column to concatenate

func main() {
	excelFile := flag.String("excelFile", "testdocs/excel.xlsx", "Excel file name with its path")
	outputFolder := flag.String("outputFolder", "testdocs/results/", "Path to output folder with trailing folder file separator")
	sheetToUse := flag.Int("sheetToUse", 0, "the sheet number to use, 0 for first sheet")
	
	flag.Parse()
	
	column := flag.Args()	
	
	xlsxFile, err := xlsx.OpenFile(*excelFile)
	check(err)

	if *sheetToUse == 0 {
		fmt.Printf("we will use the first sheet\n")
	}
	
	if len(column) == 0 {
		fmt.Printf("we will use the first column\n")
		column = append(column, "0")
	}
	
	sheet := xlsxFile.Sheets[*sheetToUse]
	for i, row := range sheet.Rows {
	//TODO make skipping first row optional
		if i >= 1 {
				//TODO only create a file if the cell[s] contain data
				fileName := fmt.Sprintf("%s%d%s",*outputFolder, (i+1), ".txt")
				fmt.Printf("%s\n", fileName)
				f, err := os.Create(fileName)
				check(err)
				for _, columnNo := range column {
					number, err := strconv.Atoi(columnNo)
					check(err)
					value := row.Cells[number].String()
					f.WriteString(value)
					f.WriteString("\n")
				}
				f.Sync()
		}
	}
}