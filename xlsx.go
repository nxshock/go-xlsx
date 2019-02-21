package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"io/ioutil"
	"strconv"
	"strings"

	"github.com/nxshock/go-xlsx/ooxml"
)

func Open(path string) (*Workbook, error) {
	r, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	var (
		sharedStrings *ooxml.SharedStrings
		//workbook      *ooxml.Workbook
	)
	for _, file := range r.File {
		switch file.Name {
		case "xl/sharedStrings.xml":
			sharedStrings, err = readSharedStrings(file)
			if err != nil {
				return nil, err
			}
			/*case "xl/workbook.xml":
			workbook, err = readWorkbook(file)
			if err != nil {
				return nil, err
			}*/
		}
	}

	w := new(Workbook)
	for _, file := range r.File {
		if !(strings.HasPrefix(file.Name, "xl/worksheets/sheet") && strings.HasSuffix(file.Name, ".xml")) {
			continue
		}

		sheet, err := readSheet(file, sharedStrings)
		if err != nil {
			return nil, err
		}
		if sheet == nil {
			continue
		}
		w.Sheet = sheet
		break
	}

	return w, err
}

func readWorkbook(f *zip.File) (workbook *ooxml.Workbook, err error) {
	r, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer r.Close()

	b, err := ioutil.ReadAll(r)
	if err != nil {
		return nil, err
	}

	err = xml.Unmarshal(b, &workbook)
	return
}

func readSharedStrings(f *zip.File) (sharedStrings *ooxml.SharedStrings, err error) {
	r, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer r.Close()

	b, err := ioutil.ReadAll(r)
	if err != nil {
		return nil, err
	}

	err = xml.Unmarshal(b, &sharedStrings)
	return
}

func readSheet(f *zip.File, sharedStrings *ooxml.SharedStrings) (sheet *Sheet, err error) {
	r, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer r.Close()

	sheet = new(Sheet)

	d := xml.NewDecoder(r)
	for {
		t, err := d.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			panic(err)
		}
		switch t := t.(type) {
		case xml.StartElement:
			if t.Name.Local == "sheetView" {
				var sheetView ooxml.SheetView
				if err := d.DecodeElement(&sheetView, &t); err != nil {
					return nil, err
				}
				if sheetView.TabSelected != 1 {
					return nil, nil
				}
				continue
			}

			if t.Name.Local == "row" {
				var (
					fileRow ooxml.Row
					row     Row
				)
				if err := d.DecodeElement(&fileRow, &t); err != nil {
					return nil, err
				}
				for i := range fileRow.Cells {
					if fileRow.Cells[i].Type == "s" {
						n, err := strconv.Atoi(fileRow.Cells[i].Value)
						if err != nil {
							panic(err)
						}
						fileRow.Cells[i].Value = sharedStrings.Strings[n].Text
					}
					c, _, err := ooxml.ParseCellName(fileRow.Cells[i].R)
					if err != nil {
						return nil, err
					}
					if c > len(row) {
						s := make([]Cell, c-len(row))
						row = append(row, s...)
					}
					row = append(row, []Cell{Cell(fileRow.Cells[i].Value)}...)
				}
				sheet.Rows = append(sheet.Rows, row)
			}
		}
	}
	return
}
