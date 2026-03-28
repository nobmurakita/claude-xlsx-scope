package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

// File はオープンしたExcelファイルを表す
type File struct {
	*excelize.File
	Name string
	path string

	// StreamSheet 用のキャッシュ（遅延初期化）
	sharedStrings *sharedStrings
	sheetPaths    map[string]string // シート名 → ZIP内のXMLパス
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
	return &File{File: f, Name: filepath.Base(path), path: path}, nil
}

// initStreamData は StreamSheet に必要な共有文字列テーブルとシートパスを遅延初期化する
func (f *File) initStreamData() error {
	if f.sharedStrings != nil {
		return nil
	}

	zr, err := zip.OpenReader(f.path)
	if err != nil {
		return err
	}
	defer zr.Close()

	f.sharedStrings, err = parseSharedStringsFromZip(zr)
	if err != nil {
		return err
	}

	f.sheetPaths, err = buildSheetPaths(zr)
	if err != nil {
		return err
	}

	return nil
}

// buildSheetPaths は workbook.xml と workbook.xml.rels からシート名→XMLパスのマップを構築する
func buildSheetPaths(zr *zip.ReadCloser) (map[string]string, error) {
	wb, err := readWorkbook(zr, "xl/workbook.xml")
	if err != nil {
		return nil, err
	}

	relsData, err := readZipFile(zr, "xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, err
	}
	var rels xmlRelationships
	if err := xml.Unmarshal(relsData, &rels); err != nil {
		return nil, fmt.Errorf("workbook.xml.rels のパースに失敗: %w", err)
	}

	targets := make(map[string]string, len(rels.Rels))
	for _, r := range rels.Rels {
		targets[r.ID] = r.Target
	}

	paths := make(map[string]string, len(wb.Sheets))
	for _, s := range wb.Sheets {
		target, ok := targets[s.RID]
		if !ok {
			continue
		}
		if !strings.HasPrefix(target, "/") {
			target = "xl/" + target
		} else {
			target = target[1:]
		}
		paths[s.Name] = target
	}
	return paths, nil
}
