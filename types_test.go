package xlsx_test

import (
	"testing"
	"time"

	xlsx "github.com/nxshock/go-xlsx"
	"github.com/stretchr/testify/assert"
)

func Test(t *testing.T) {
	cell := xlsx.Cell("0")
	assert.Equal(t, time.Date(1899, 12, 30, 0, 0, 0, 0, time.Local), cell.Time())

	cell = xlsx.Cell("43517")
	assert.Equal(t, time.Date(2019, 2, 21, 0, 0, 0, 0, time.Local), cell.Time())
}
