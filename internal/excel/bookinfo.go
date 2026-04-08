package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"path/filepath"
	"strings"
)

// BookInfo は ZIP から直接メタデータを取得する。
// info コマンド用の軽量パス。

// SheetInfo はシートの基本情報
type SheetInfo struct {
	Index  int    `json:"index"`
	Name   string `json:"name"`
	Type   string `json:"type"`
	Hidden bool   `json:"hidden,omitempty"`
}

// BookInfoResult は BookInfo の結果
type BookInfoResult struct {
	FileName     string
	Sheets       []SheetInfo
	DefinedNames []DefinedNameInfo
}

// DefinedNameInfo は定義名の情報
type DefinedNameInfo struct {
	Name     string
	Scope    string
	RefersTo string
}

// BookInfo は ZIP から workbook.xml のみを読み、シート一覧と定義名を返す。
func BookInfo(path string) (*BookInfoResult, error) {
	ext := strings.ToLower(filepath.Ext(path))
	if ext != ".xlsx" && ext != ".xlsm" {
		return nil, fmt.Errorf(".xlsx / .xlsm 形式のみ対応しています")
	}

	r, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	zi := newZipIndex(r)

	// workbook.xml.rels を読んでリレーション種別を取得
	relTypes, err := readRels(zi, "xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, err
	}

	// workbook.xml を読んでシートと定義名を取得
	wb, err := readWorkbook(zi, "xl/workbook.xml")
	if err != nil {
		return nil, err
	}

	// シート情報を構築
	sheets := make([]SheetInfo, len(wb.Sheets))
	for i, s := range wb.Sheets {
		sheetType := detectSheetType(relTypes[s.RID])
		sheets[i] = SheetInfo{
			Index:  i,
			Name:   s.Name,
			Type:   sheetType,
			Hidden: s.State != "" && s.State != "visible",
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

	return &BookInfoResult{
		FileName:     filepath.Base(path),
		Sheets:       sheets,
		DefinedNames: definedNames,
	}, nil
}

// detectSheetType はリレーション種別からシートタイプを判定する
func detectSheetType(relType string) string {
	switch {
	case strings.Contains(relType, "chartsheet"):
		return "chartsheet"
	case strings.Contains(relType, "dialogsheet"):
		return "dialogsheet"
	case strings.Contains(relType, "macrosheetx"):
		return "macrosheet"
	default:
		return "worksheet"
	}
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

// zipIndex は ZIP 内のファイルを名前で高速検索するためのインデックス
type zipIndex struct {
	files map[string]*zip.File
}

// newZipIndex は zip.ReadCloser からインデックスを構築する
func newZipIndex(zr *zip.ReadCloser) *zipIndex {
	m := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		m[f.Name] = f
	}
	return &zipIndex{files: m}
}

// lookup は指定パスのファイルを返す。見つからなければ nil
func (zi *zipIndex) lookup(name string) *zip.File {
	return zi.files[name]
}

// readZipFile は ZIP 内の指定パスのファイルを読み込む。
// ファイルが存在しない場合は (nil, nil) を返す。
func readZipFile(zi *zipIndex, name string) ([]byte, error) {
	entry := zi.lookup(name)
	if entry == nil {
		return nil, nil
	}
	rc, err := entry.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}

func readWorkbook(zi *zipIndex, name string) (*xmlWorkbook, error) {
	data, err := readZipFile(zi, name)
	if err != nil {
		return nil, err
	}
	if data == nil {
		return nil, fmt.Errorf("ZIP内に %s が見つかりません", name)
	}
	var wb xmlWorkbook
	if err := xml.Unmarshal(data, &wb); err != nil {
		return nil, fmt.Errorf("workbook.xml のパースに失敗: %w", err)
	}
	return &wb, nil
}

// readRels は .rels ファイルを読み、rId → Type のマップを返す。
// .rels ファイルが存在しない場合はエラーではなく空マップを返す（rels はオプショナルなため）。
func readRels(zi *zipIndex, name string) (map[string]string, error) {
	data, err := readZipFile(zi, name)
	if err != nil {
		return nil, err
	}
	if data == nil {
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
