package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"path/filepath"
	"strings"
)

// QuickInfo は excelize を使わず ZIP から直接メタデータを取得する。
// info コマンド用の軽量パス。

// SheetInfo はシートの基本情報
type SheetInfo struct {
	Index  int    `json:"index"`
	Name   string `json:"name"`
	Type   string `json:"type"`
	Hidden bool   `json:"hidden,omitempty"`
}

// QuickInfoResult は QuickInfo の結果
type QuickInfoResult struct {
	FileName     string
	Sheets       []SheetInfo
	DefinedNames []DefinedNameInfo
}

// DefinedNameInfo は定義名の情報
type DefinedNameInfo struct {
	Name    string
	Scope   string
	RefersTo string
}

// QuickInfo は ZIP から workbook.xml のみを読み、シート一覧と定義名を返す。
func QuickInfo(path string) (*QuickInfoResult, error) {
	ext := strings.ToLower(filepath.Ext(path))
	if ext != ".xlsx" && ext != ".xlsm" {
		return nil, fmt.Errorf(".xlsx / .xlsm 形式のみ対応しています")
	}

	r, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	// workbook.xml.rels を読んでリレーション種別を取得
	relTypes, err := readRels(r, "xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, err
	}

	// workbook.xml を読んでシートと定義名を取得
	wb, err := readWorkbook(r, "xl/workbook.xml")
	if err != nil {
		return nil, err
	}

	// シート情報を構築
	sheets := make([]SheetInfo, len(wb.Sheets))
	for i, s := range wb.Sheets {
		sheetType := "worksheet"
		if rel, ok := relTypes[s.RID]; ok {
			if strings.Contains(rel, "chartsheet") {
				sheetType = "chartsheet"
			} else if strings.Contains(rel, "dialogsheet") {
				sheetType = "dialogsheet"
			} else if strings.Contains(rel, "macrosheetx") {
				sheetType = "macrosheet"
			}
		}
		sheets[i] = SheetInfo{
			Index:  i,
			Name:   s.Name,
			Type:   sheetType,
			Hidden: s.State != "",
		}
	}

	// 定義名を構築
	var definedNames []DefinedNameInfo
	for _, dn := range wb.DefinedNames {
		scope := ""
		if dn.LocalSheetID != nil {
			idx := *dn.LocalSheetID
			if idx >= 0 && idx < len(wb.Sheets) {
				scope = wb.Sheets[idx].Name
			}
		}
		definedNames = append(definedNames, DefinedNameInfo{
			Name:     dn.Name,
			Scope:    scope,
			RefersTo: dn.Value,
		})
	}

	return &QuickInfoResult{
		FileName:     filepath.Base(path),
		Sheets:       sheets,
		DefinedNames: definedNames,
	}, nil
}

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

// workbook.xml.rels のパース用構造体
type xmlRelationships struct {
	Rels []xmlRelationship `xml:"Relationship"`
}

type xmlRelationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

// relsPathFor は XML パスから対応する .rels ファイルのパスを構築する
func relsPathFor(xmlPath string) string {
	dir := xmlPath[:strings.LastIndex(xmlPath, "/")+1]
	base := xmlPath[strings.LastIndex(xmlPath, "/")+1:]
	return dir + "_rels/" + base + ".rels"
}

// withZipXML は ZIP エントリを開いて xml.Decoder を渡し、終了後にクローズする。
// SAX パーサーの共通ボイラープレートを吸収する。
func withZipXML(entry *zip.File, fn func(decoder *xml.Decoder) error) error {
	rc, err := entry.Open()
	if err != nil {
		return err
	}
	defer rc.Close()
	return fn(xml.NewDecoder(rc))
}

// findZipEntry は ZIP 内の指定パスのエントリを探す
func findZipEntry(zr *zip.ReadCloser, name string) *zip.File {
	for _, f := range zr.File {
		if f.Name == name {
			return f
		}
	}
	return nil
}

// readZipFile は ZIP 内の指定パスのファイルを読み込む
func readZipFile(r *zip.ReadCloser, name string) ([]byte, error) {
	entry := findZipEntry(r, name)
	if entry == nil {
		return nil, fmt.Errorf("ZIP内に %s が見つかりません", name)
	}
	rc, err := entry.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}

func readWorkbook(r *zip.ReadCloser, name string) (*xmlWorkbook, error) {
	data, err := readZipFile(r, name)
	if err != nil {
		return nil, err
	}
	var wb xmlWorkbook
	if err := xml.Unmarshal(data, &wb); err != nil {
		return nil, fmt.Errorf("workbook.xml のパースに失敗: %w", err)
	}
	return &wb, nil
}

func readRels(r *zip.ReadCloser, name string) (map[string]string, error) {
	data, err := readZipFile(r, name)
	if err != nil {
		// rels がない場合は空を返す
		return map[string]string{}, nil
	}
	var rels xmlRelationships
	if err := xml.Unmarshal(data, &rels); err != nil {
		return nil, fmt.Errorf("rels のパースに失敗: %w", err)
	}
	m := make(map[string]string, len(rels.Rels))
	for _, rel := range rels.Rels {
		m[rel.ID] = rel.Type
	}
	return m, nil
}
