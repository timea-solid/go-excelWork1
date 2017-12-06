package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
)

func check(e error) {
    if e != nil {
        panic(e)
    }
}

func main() {
	var sheet *xlsx.Sheet
	excelFile := "D:/SemanticWeb/MyClients&Projects/VirtualAssistant/Content/InputDataDE_FinServ2Click.xlsx"
	outputFilePath := "D:/SemanticWeb/MyClients&Projects/VirtualAssistant/Content/InputDataDE_FinServ2Click"
	xlsxFile, err := xlsx.OpenFile(excelFile)
	if err != nil {
		fmt.Printf("there was an error\n")
	}
	sheet = xlsxFile.Sheets[6]
	index := 0
	for _, row := range sheet.Rows {
		if index >= 1 {
			question := row.Cells[12].String()
			answer := row.Cells[18].String()
			hint := row.Cells[20].String()
//			fmt.Printf("%s\n %s\n %s\n",question, answer, hint)
			fileName := fmt.Sprintf("%s%s%d%s",outputFilePath,  "/", (index+1), ".txt")
			f, err := os.Create(fileName)
			check(err)
			f.WriteString(question)
			f.WriteString("\n")
			f.WriteString(answer)
			f.WriteString("\n")
			f.WriteString(hint)
			f.Sync()
		}
		index ++ 
	}
}