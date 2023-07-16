package main

import (
	"fmt"
	"image/png"
	"io"
	"os"
	"unicode"

	"github.com/xuri/excelize/v2"
)

type Pixel struct {
	R uint32
	G uint32
	B uint32
	A uint32
}

func rgbaToPixel(r, g, b, a uint32) Pixel {
	return Pixel{r / 257, g / 257, b / 257, a / 257}
}

// stolen from:
// https://stackoverflow.com/a/41185404
func getPixels(f io.Reader) ([][]Pixel, error) {
	im, err := png.Decode(f)
	if err != nil {
		return nil, err
	}

	bounds := im.Bounds()
	w, h := bounds.Max.X, bounds.Max.Y

	var pixels [][]Pixel
	for y := 0; y < h; y++ {
		var row []Pixel
		for x := 0; x < w; x++ {
			row = append(row, rgbaToPixel(im.At(x, y).RGBA()))
		}
		pixels = append(pixels, row)
	}

	return pixels, nil
}

func pixToHex(p Pixel) string {
	return fmt.Sprintf("%02x%02x%02x", p.R, p.G, p.B)
}

func rmCellNameNumber(cellName string) string {
	newCellName := ""

	for _, c := range cellName {
		if !unicode.IsDigit(c) {
			newCellName += string(c)
		}
	}

	return newCellName
}

func check(err error) {
	if err != nil {
		panic(err)
	}
}

func main() {
	filename := os.Args[1]
	f, err := os.Open(filename)
	check(err)
	defer f.Close()

	pixels, err := getPixels(f)
	check(err)

	excel := excelize.NewFile()
	defer excel.Close()
	sheet := "Sheet1"

	colidx := 1
	rowidx := 1

	colWidth := 1.0
	rowHeight := 6.0

	for y, row := range pixels {
		for x, col := range row {
			cell, err := excelize.CoordinatesToCellName(x+1, y+1)
			cellNoNumber := rmCellNameNumber(cell)
			check(err)
			cellStyle, err := excel.NewStyle(&excelize.Style{
				Fill: excelize.Fill{Type: "pattern", Color: []string{pixToHex(col)}, Pattern: 1},
			})
			check(err)

			excel.SetCellStyle(sheet, cell, cell, cellStyle)
			colidx++

			excel.SetColWidth(sheet, cellNoNumber, cellNoNumber, colWidth)
			excel.SetRowHeight(sheet, rowidx, rowHeight)
		}
		rowidx++
		colidx = 1
	}

	excel.SaveAs("out.xlsx")
}
