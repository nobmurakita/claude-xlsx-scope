package excel

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// SheetInfo はシートの基本情報
type SheetInfo struct {
	Index     int    `json:"index"`
	Name      string `json:"name"`
	Type      string `json:"type"`
	Hidden    bool   `json:"hidden,omitempty"`
	Dimension string `json:"-"` // XMLのdimension属性（info出力用）
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

// GetSheetDimension は excelize の GetSheetDimension をラップする
func (f *File) GetSheetDimension(sheet string) (string, error) {
	return f.File.GetSheetDimension(sheet)
}

// GetUsedRange はシートの使用範囲を返す。空シートの場合は空文字を返す。
// rowCache が非nilの場合はそれを使い、nilの場合は GetSheetDimension にフォールバックする。
func (f *File) GetUsedRange(sheet string, rowCache *RowCache) (string, error) {
	// RowCache がある場合はそこから算出
	if rowCache != nil {
		return rowCache.CalcUsedRange(), nil
	}

	dim, err := f.File.GetSheetDimension(sheet)
	if err != nil {
		return "", err
	}

	if dim != "" && dim != "A1:A1" {
		return dim, nil
	}

	// フォールバック: GetRows で算出
	rc, err := f.LoadRows(sheet)
	if err != nil {
		return "", err
	}
	return rc.CalcUsedRange(), nil
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
