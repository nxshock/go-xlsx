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

func ReadFileByRow(filePath string) (fileData chan []Cell, err error) {
	r, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, err
	}

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

	//TODO: добавить поиск нужного листа

	for _, file := range r.File {
		if !(strings.HasPrefix(file.Name, "xl/worksheets/sheet") && strings.HasSuffix(file.Name, ".xml")) {
			continue
		}

		data, err := readSheetCells(file, sharedStrings)
		if err != nil {
			return nil, err
		}

		fileData = make(chan []Cell)
		go func() {
			defer r.Close()
			defer close(fileData)

			for row := range data {
				fileData <- row
			}
		}()

		break
	}

	return fileData, err
}

func OpenAsCellSlice(filePath string) ([][]Cell, error) {
	rowChan, err := ReadFileByRow(filePath)
	if err != nil {
		return nil, err
	}

	var result [][]Cell

	for row := range rowChan {
		result = append(result, row)
	}

	return result, err
}

func OpenAsStringSlice(filePath string) ([][]string, error) {
	rowChan, err := ReadFileByRow(filePath)
	if err != nil {
		return nil, err
	}

	var result [][]string

	for row := range rowChan {
		var stringRow []string

		for _, cell := range row {
			stringRow = append(stringRow, cell.String())
		}

		result = append(result, stringRow)
	}

	return result, err
}

/*func readWorkbook(f *zip.File) (workbook *ooxml.Workbook, err error) {
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
}*/

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

func readAll(f *zip.File, sharedStrings *ooxml.SharedStrings) (data [][]Cell, err error) {
	cellsChan, err := readSheetCells(f, sharedStrings)
	if err != nil {
		return nil, err
	}

	for cells := range cellsChan {
		data = append(data, cells)
	}

	return data, nil
}

func readSheetCells(f *zip.File, sharedStrings *ooxml.SharedStrings) (cells chan []Cell, err error) {
	cells = make(chan []Cell)

	r, err := f.Open()
	if err != nil {
		return nil, err
	}

	d := xml.NewDecoder(r)
	go func() {
		defer r.Close()
		defer close(cells)

		for {
			t, err := d.Token()
			if err != nil {
				if err == io.EOF {
					break
				}
				panic(err) //TODO: do not panic
			}
			switch t := t.(type) {
			case xml.StartElement:
				if t.Name.Local == "sheetView" {
					var sheetView ooxml.SheetView
					if err := d.DecodeElement(&sheetView, &t); err != nil {
						//return nil, err
						panic(err)
					}
					if sheetView.TabSelected != 1 {
						//return nil, nil
						panic(err) //TODO: переделать на проверку, является ли данный лист активным, раньше в коде
					}
					continue
				}

				if t.Name.Local == "row" {
					var (
						fileRow ooxml.Row
						row     []Cell
					)
					if err := d.DecodeElement(&fileRow, &t); err != nil {
						//return nil, err
						panic(err)
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
							//return nil, err
							panic(err)
						}
						if c > len(row) {
							s := make([]Cell, c-len(row))
							row = append(row, s...)
						}
						row = append(row, []Cell{Cell(fileRow.Cells[i].Value)}...)
					}
					cells <- row
				}
			}
		}
	}()

	return cells, nil
}
