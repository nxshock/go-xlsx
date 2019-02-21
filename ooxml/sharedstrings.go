package ooxml

type SharedStrings struct {
	Strings []Si `xml:"si"`
}

type Si struct {
	Text string `xml:"t"`
}
