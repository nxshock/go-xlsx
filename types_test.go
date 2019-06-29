package xlsx_test

import (
	"testing"
	"time"

	xlsx "github.com/nxshock/go-xlsx"
	"github.com/stretchr/testify/assert"
)

func TestExcelTime(t *testing.T) {
	assert.Equal(t, time.Date(1899, 12, 30, 0, 0, 0, 0, time.Local), xlsx.Cell("0").ExcelTime())
	assert.Equal(t, time.Date(2019, 2, 21, 0, 0, 0, 0, time.Local), xlsx.Cell("43517").ExcelTime())
}

func TestParseTime(t *testing.T) {
	assert.Equal(t, time.Date(2000, 1, 2, 3, 4, 5, 0, time.Local), xlsx.Cell("20000102030405").Time("20060102030405"))
}
