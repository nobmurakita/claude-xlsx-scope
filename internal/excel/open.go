package excel

import (
	"fmt"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

// File はオープンしたExcelファイルを表す
type File struct {
	*excelize.File
	Name string
}

// OpenFile はExcelファイルを開く。.xlsx/.xlsm のみ対応。
func OpenFile(path string) (*File, error) {
	ext := strings.ToLower(filepath.Ext(path))
	if ext != ".xlsx" && ext != ".xlsm" {
		return nil, fmt.Errorf(".xlsx / .xlsm 形式のみ対応しています")
	}
	f, err := excelize.OpenFile(path, excelize.Options{
		// 数値フォーマット処理をスキップし、生の値を返す。
		// セル値のフォーマットは ReadCell 側で行うため excelize 側は不要。
		RawCellValue: true,
	})
	if err != nil {
		return nil, err
	}
	return &File{File: f, Name: filepath.Base(path)}, nil
}
