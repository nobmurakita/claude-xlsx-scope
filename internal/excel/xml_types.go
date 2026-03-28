package excel

import "encoding/xml"

// リレーション種別のキーワード（Relationship の Type 属性に含まれる文字列で判定する）
const (
	relKeywordDrawing          = "drawing"
	relKeywordComments         = "/comments"
	relKeywordThreadedComments = "threadedcomments" // 大文字小文字混在あり、小文字で比較する
)

// workbook.xml のパース用構造体（複数モジュールで共有）
// ※ 各パーサー固有の XML 構造体は対応する *_parse.go ファイルに定義

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
