package excel

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// SheetInfo はシートの基本情報
type SheetInfo struct {
	Index  int    `json:"index"`
	Name   string `json:"name"`
	Type   string `json:"type"`
	Hidden bool   `json:"hidden,omitempty"`
}

// GetSheetList はシート一覧を返す
func (f *File) GetSheetList() ([]SheetInfo, error) {
	list := f.File.GetSheetList()
	sheets := make([]SheetInfo, 0, len(list))
	for i, name := range list {
		info := SheetInfo{
			Index: i,
			Name:  name,
		}
		info.Type = f.getSheetType(name)
		visible, err := f.File.GetSheetVisible(name)
		if err != nil {
			return nil, err
		}
		info.Hidden = !visible
		sheets = append(sheets, info)
	}
	return sheets, nil
}

func (f *File) getSheetType(name string) string {
	// excelize はチャートシートかワークシートを区別できる
	// GetSheetType は excelize v2 にはない。チャートシートかどうかを判定するために
	// チャートシートのXMLを確認する
	props, err := f.File.GetSheetProps(name)
	if err != nil {
		// チャートシートの場合 GetSheetProps はエラーを返す
		return "chartsheet"
	}
	_ = props
	return "worksheet"
}

// ResolveSheet は --sheet オプションの値からシート名を解決する。
// 空文字の場合は最初のシートを返す。
func (f *File) ResolveSheet(sheet string) (string, error) {
	list := f.File.GetSheetList()
	if len(list) == 0 {
		return "", fmt.Errorf("ブックにシートがありません")
	}
	if sheet == "" {
		return list[0], nil
	}

	// インデックス指定を試みる
	if idx, err := strconv.Atoi(sheet); err == nil {
		if idx < 0 || idx >= len(list) {
			return "", fmt.Errorf("シートインデックス %d が範囲外です（利用可能: %s）", idx, formatSheetNames(list))
		}
		return list[idx], nil
	}

	// 名前指定
	for _, name := range list {
		if name == sheet {
			return name, nil
		}
	}
	return "", fmt.Errorf("シート %q が見つかりません（利用可能: %s）", sheet, formatSheetNames(list))
}

// ResolveWorksheet は ResolveSheet と同様だが、ワークシート以外はエラーにする。
// scan / dump / search 用。
func (f *File) ResolveWorksheet(sheet string) (string, error) {
	name, err := f.ResolveSheet(sheet)
	if err != nil {
		return "", err
	}
	if f.getSheetType(name) != "worksheet" {
		return "", fmt.Errorf("シート %q はワークシートではありません（ワークシートのみ対応）", name)
	}
	return name, nil
}

// GetUsedRange はシートの使用範囲を返す。空シートの場合は空文字を返す。
func (f *File) GetUsedRange(sheet string) (string, error) {
	dim, err := f.File.GetSheetDimension(sheet)
	if err != nil {
		return "", err
	}

	// GetSheetDimension が信頼できる値を返した場合
	if dim != "" && dim != "A1:A1" {
		return dim, nil
	}

	// dim が空や "A1:A1" の場合、全行を走査して実際の使用範囲を算出する
	return f.calcUsedRange(sheet)
}

func (f *File) calcUsedRange(sheet string) (string, error) {
	rows, err := f.File.GetRows(sheet)
	if err != nil {
		return "", err
	}
	if len(rows) == 0 {
		return "", nil
	}

	minCol, maxCol := -1, -1
	minRow, maxRow := -1, -1

	for r, row := range rows {
		for c, cell := range row {
			if cell != "" {
				rowIdx := r + 1
				colIdx := c + 1
				if minRow == -1 || rowIdx < minRow {
					minRow = rowIdx
				}
				if maxRow == -1 || rowIdx > maxRow {
					maxRow = rowIdx
				}
				if minCol == -1 || colIdx < minCol {
					minCol = colIdx
				}
				if maxCol == -1 || colIdx > maxCol {
					maxCol = colIdx
				}
			}
		}
	}

	if minRow == -1 {
		return "", nil
	}

	return CellRef(minCol, minRow) + ":" + CellRef(maxCol, maxRow), nil
}

func formatSheetNames(names []string) string {
	quoted := make([]string, len(names))
	for i, n := range names {
		quoted[i] = fmt.Sprintf("%q", n)
	}
	return strings.Join(quoted, ", ")
}

// GetDefinedNames は定義名一覧を返す
func (f *File) GetDefinedNames() []excelize.DefinedName {
	return f.File.GetDefinedName()
}
