package main

import (
	"bufio"
	"os"

	"github.com/xuri/excelize/v2"
)

const (
	WHITE = "FFFFFF"
	BLACK = "000000"
)

func check(err error) {
	if err != nil {
		panic(err)
	}
}

func main() {
	f, err := os.Open("f.txt")
	check(err)
	defer f.Close()

	var ftext []string
	scanner := bufio.NewScanner(f)
	for scanner.Scan() {
		ftext = append(ftext, scanner.Text())
	}

	excel := excelize.NewFile()
	defer excel.Close()
	sheet := "Sheet1"

	white, err := excel.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{WHITE}, Pattern: 1},
	})
	check(err)

	black, err := excel.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{BLACK}, Pattern: 1},
	})
	check(err)

	colidx := 1
	rowidx := 1

	colWidth := 2.0
	rowHeight := 8.0

	for _, line := range ftext {
		for _, char := range line {
			c := int(char) - '0'
			cell, err := excelize.CoordinatesToCellName(colidx, rowidx)
			check(err)
			var cellStyle int

			if c == 0 {
				cellStyle = white
			} else {
				cellStyle = black
			}

			excel.SetCellStyle(sheet, cell, cell, cellStyle)
			colidx++

			excel.SetColWidth(sheet, string(cell[0]), string(cell[0]), colWidth)
			excel.SetRowHeight(sheet, rowidx, rowHeight)
		}
		rowidx++
		colidx = 1
	}

	excel.SaveAs("out.xlsx")
}
