package xlsx

import (
	"bufio"
	"math"
	"os"
	"strconv"
	"time"
)

type Workbook struct {
	Sheet *Sheet
}

type Sheet struct {
	Name string
	Rows [][]Cell
}

type Cell string

func (workbook *Workbook) SaveToCSV(fileName string) error {
	f, _ := os.Create(fileName)
	defer f.Close()

	buf := bufio.NewWriter(f)
	defer buf.Flush()

	sheet := workbook.Sheet

	for _, row := range sheet.Rows {
		for i, cell := range row {
			if i == 0 {
				buf.WriteString(string(cell))
			} else {
				buf.WriteString(";" + string(cell))
			}
		}
		buf.WriteString("\n")
	}

	return nil
}

func (workbook *Workbook) Get(row, col int) Cell {
	if row >= len(workbook.Sheet.Rows) {
		return ""
	}

	if col >= len(workbook.Sheet.Rows[row]) {
		return ""
	}

	return workbook.Sheet.Rows[row][col]
}

func (cell Cell) String() string {
	return string(cell)
}

func (cell Cell) Integer() int {
	f, err := strconv.ParseFloat(string(cell), 64)
	if err != nil {
		return 0
	}
	return int(math.Round(f))
}

func (cell Cell) Float() float64 {
	f, err := strconv.ParseFloat(string(cell), 64)
	if err != nil {
		return 0
	}
	return f
}

func (cell Cell) ExcelTime() time.Time {
	f, err := strconv.ParseFloat(string(cell), 64)
	if err != nil {
		return time.Time{}
	}
	return time.Date(1900, 1, 1, 0, 0, 0, 0, time.Local).Add(time.Duration(24*float64(time.Hour)*f - 2*24*float64(time.Hour)))
}
