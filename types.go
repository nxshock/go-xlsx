package xlsx

import (
	"math"
	"strconv"
	"time"
)

// Workbook represents Excel workbook
type Workbook struct {
	Sheet *Sheet
}

// Sheet represents Excel sheet
type Sheet struct {
	Name string
	Rows [][]Cell
}

// Cell represents table cell
type Cell string

// String returns cell data as string
func (cell Cell) String() string {
	return string(cell)
}

// Int returns cell data as integer
func (cell Cell) Int() int {
	f, err := strconv.ParseFloat(string(cell), 64)
	if err != nil {
		return 0
	}
	return int(math.Round(f))
}

// Float returns cell data as float64
func (cell Cell) Float() float64 {
	f, err := strconv.ParseFloat(string(cell), 64)
	if err != nil {
		return 0
	}
	return f
}

// ExcelTime returns cell time, if cell contains Excel date/time value
func (cell Cell) ExcelTime() time.Time {
	var excelTimeShift = time.Date(1900, 1, 1, 0, 0, 0, 0, time.Local)

	f, err := strconv.ParseFloat(string(cell), 64)
	if err != nil {
		return time.Time{}
	}
	return excelTimeShift.Add(time.Duration(24*float64(time.Hour)*f - 2*24*float64(time.Hour)))
}

// Time returns cell parsed time with specified layout
func (cell Cell) Time(layout string) time.Time {
	t, err := time.ParseInLocation(layout, string(cell), time.Local)
	if err != nil {
		return time.Time{}
	}

	return t
}
