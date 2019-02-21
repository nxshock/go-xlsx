package ooxml

type Workbook struct {
	Sheets []Sheet `xml:"sheets>sheet"`
}

type Sheet struct {
	Name string `xml:"name,attr"`
	ID   int    `xml:"sheetId,attr"`
}
