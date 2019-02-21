package ooxml

import (
	"strconv"
)

func ParseCellName(s string) (col, row int, err error) {
	var (
		colStr string
		rowStr string
	)

	for _, v := range s {
		if v >= 48 && v <= 57 {
			rowStr += string(v)
		} else {
			colStr += string(v)
		}
	}

	for i := 0; i < len(colStr); i++ {
		col *= 26
		col += int(colStr[i]) - int('A') + 1
	}
	col -= 1

	row, err = strconv.Atoi(rowStr)
	if err != nil {
		return 0, 0, err
	}
	row -= 1

	return col, row, nil
}
