package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"strconv"
)

func main() {
	var filePath string
	fmt.Println("冲鸭!!! 请把待处理的Excel拖入这个黑框框")
	fmt.Scan(&filePath)
	var list = make(map[string][]xlsx.Row)
	file, err := xlsx.OpenFile(filePath)
	if err != nil {
		panic(err)
	}
	sheet, ok := file.Sheet["Sheet1"]
	if ok {
	}
	rows := sheet.Rows
	var leng1 int
	for i := range rows {
		var cells []*xlsx.Cell = rows[i].Cells
		leng1 = len(cells)
		name := cells[0].Value
		_, ok := list[name]
		if !ok {
			i2 := []xlsx.Row{*rows[i]}
			list[name] = i2
		} else {
			list[cells[0].Value] = append(list[name], *rows[i])
		}
	}

	for key := range list {
		if key == "姓名" {
			continue
		}
		rows := list[key]
		for i := 0; i < len(rows)-1; i++ {
			for j := i + 1; j < len(rows); j++ {
				rowi := rows[i]
				rowj := rows[j]
				valuei := rowi.Cells[leng1-1].Value
				valuej := rowj.Cells[leng1-1].Value
				iint, _ := strconv.Atoi(valuei)
				jint, _ := strconv.Atoi(valuej)
				if iint > jint {
					rows[i], rows[j] = rows[j], rows[i]
				}
			}
		}
	}

	addSheet, _ := file.AddSheet("汇总")
	for key := range list {
		value := list[key]
		for index := range value {
			row := addSheet.AddRow()
			cells := value[index].Cells
			for ic := range cells {
				cell := cells[ic]
				row.AddCell().Value = cell.Value
			}
		}
	}
	newFile1, err := xlsx.NewFile().AppendSheet(*addSheet, "汇总")
	if err == nil {
		fmt.Println("成功!!!")
	}
	file.Save(filePath)
	fmt.Println(newFile1)
}
