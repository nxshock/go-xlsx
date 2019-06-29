package xlsx_test

import (
	"testing"

	xlsx "github.com/nxshock/go-xlsx"
	"github.com/stretchr/testify/assert"
)

func TestAllFile1(t *testing.T) {
	expectedData := [][]xlsx.Cell{
		[]xlsx.Cell{"column1", "column2"},
		[]xlsx.Cell{"1", "text1"},
		[]xlsx.Cell{"2", "text2"}}

	f, err := xlsx.Open("testdata/test1.xlsx")
	assert.Nil(t, err)

	assert.Equal(t, expectedData, f)
}

func TestByRowFile1(t *testing.T) {
	expectedData := [][]xlsx.Cell{
		[]xlsx.Cell{"column1", "column2"},
		[]xlsx.Cell{"1", "text1"},
		[]xlsx.Cell{"2", "text2"}}

	var gotData [][]xlsx.Cell

	rowsChan, err := xlsx.ReadFileByRow("testdata/test1.xlsx")
	assert.Nil(t, err)

	for row := range rowsChan {
		gotData = append(gotData, row)
	}

	assert.Equal(t, expectedData, gotData)
}
