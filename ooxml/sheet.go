package ooxml

type Cell struct {
	R     string `xml:"r,attr"`
	Type  string `xml:"t,attr"`
	Value string `xml:"v"`
}

type Row struct {
	Num   int    `xml:"r,attr"`
	Cells []Cell `xml:"c"`
}

type SheetData struct {
	Rows []Row `xml:"row"`
}

type SheetView struct {
	TabSelected int `xml:"tabSelected,attr"`
}

/*type Worksheet struct {
	//SheetData SheetData `xml:"sheetData"`
	SheetView SheetView `xml:"sheetViews>sheetView"`
}*/
