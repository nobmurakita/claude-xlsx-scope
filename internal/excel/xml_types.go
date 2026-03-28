package excel

import "encoding/xml"

// workbook.xml のパース用構造体

type xmlWorkbook struct {
	Sheets       []xmlSheet       `xml:"sheets>sheet"`
	DefinedNames []xmlDefinedName `xml:"definedNames>definedName"`
}

type xmlSheet struct {
	Name  string `xml:"name,attr"`
	RID   string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	State string `xml:"state,attr"`
}

type xmlDefinedName struct {
	Name         string `xml:"name,attr"`
	LocalSheetID *int   `xml:"localSheetId,attr"`
	Value        string `xml:",chardata"`
}

// workbook.xml.rels 等のリレーション用構造体（複数モジュールで共有）

type xmlRelationships struct {
	XMLName xml.Name          `xml:"Relationships"`
	Rels    []xmlRelationship `xml:"Relationship"`
}

type xmlRelationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}
