package ooxml_test

import (
	"testing"

	"github.com/nxshock/go-xlsx/ooxml"
	"github.com/stretchr/testify/assert"
)

func TestParseCellName(t *testing.T) {
	c, r, err := ooxml.ParseCellName("A1")
	assert.Nil(t, err)
	assert.Equal(t, c, 0)
	assert.Equal(t, r, 0)

	c, r, err = ooxml.ParseCellName("XFD1048576")
	assert.Nil(t, err)
	assert.Equal(t, c, 16383)
	assert.Equal(t, r, 1048575)
}
