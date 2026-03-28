package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// File はオープンしたExcelファイルを表す
type File struct {
	*excelize.File       // scan 用（nil の場合は lite モード）
	Name string
	path string

	// 自前パースデータ（lite モードおよび StreamSheet で使用）
	sharedStrings *sharedStrings
	sheetPaths    map[string]string // シート名 → ZIP内のXMLパス
	sheetNames    []string          // シート名（workbook.xml の順序）
	styles        *styleSheet       // styles.xml
	theme         *themeColors      // theme1.xml
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

// OpenFileLite は excelize を使わず、必要なメタデータのみをZIPから直接パースする。
// dump/search コマンド用の軽量オープン。ワークシートXMLの展開を行わない。
func OpenFileLite(path string) (*File, error) {
	ext := strings.ToLower(filepath.Ext(path))
	if ext != ".xlsx" && ext != ".xlsm" {
		return nil, fmt.Errorf(".xlsx / .xlsm 形式のみ対応しています")
	}

	zr, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer zr.Close()

	f := &File{
		Name: filepath.Base(path),
		path: path,
	}

	// 共有文字列テーブル
	f.sharedStrings, err = parseSharedStringsFromZip(zr)
	if err != nil {
		return nil, err
	}

	// シートパスマップ
	f.sheetPaths, f.sheetNames, err = buildSheetPaths(zr)
	if err != nil {
		return nil, err
	}

	// styles.xml
	stylesData, err := readZipFileFromReader(zr, "xl/styles.xml")
	if err != nil {
		return nil, fmt.Errorf("styles.xml の読み込みに失敗: %w", err)
	}
	f.styles, err = parseStyleSheet(stylesData)
	if err != nil {
		return nil, fmt.Errorf("styles.xml のパースに失敗: %w", err)
	}

	// theme1.xml（存在しなくてもエラーにしない）
	themeData, err := readZipFileFromReader(zr, "xl/theme/theme1.xml")
	if err == nil {
		f.theme = parseThemeColors(themeData)
	}

	return f, nil
}

// IsLite は excelize を使わない軽量モードかどうかを返す
func (f *File) IsLite() bool {
	return f.File == nil
}

// Close はファイルを閉じる。lite モードでは何もしない。
func (f *File) CloseLite() {
	if f.File != nil {
		f.File.Close()
	}
}

// LoadSheetMetaLite はワークシートXMLからメタデータを直接パースする（lite モード用）
func (f *File) LoadSheetMetaLite(sheet string) (*SheetMeta, error) {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return nil, fmt.Errorf("シート %q が見つかりません", sheet)
	}
	zr, err := zip.OpenReader(f.path)
	if err != nil {
		return nil, err
	}
	defer zr.Close()
	return LoadSheetMeta(zr, xmlPath)
}

// LoadSheetRelsLite はシートのリレーションを読む（lite モード用）
func (f *File) LoadSheetRelsLite(sheet string) map[string]string {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return nil
	}
	zr, err := zip.OpenReader(f.path)
	if err != nil {
		return nil
	}
	defer zr.Close()
	return LoadSheetRels(zr, xmlPath)
}

// DetectDefaultFontLite は styles.xml からデフォルトフォントを返す（lite モード用）
func (f *File) DetectDefaultFontLite() FontInfo {
	if f.styles == nil {
		return FontInfo{Name: "Calibri", Size: 11}
	}
	name := f.styles.DefaultFontName()
	if name == "" {
		name = "Calibri"
	}
	return FontInfo{Name: name, Size: 11}
}

// DetectDefaultFontFromMeta は SheetMeta の列スタイルから最頻フォントを検出する
func (f *File) DetectDefaultFontFromMeta(meta *SheetMeta) FontInfo {
	bookDefault := f.DetectDefaultFontLite()
	if f.styles == nil || len(meta.Cols) == 0 {
		return bookDefault
	}

	type fontKey struct {
		name string
		size float64
	}
	counts := make(map[fontKey]int)

	for _, ci := range meta.Cols {
		if ci.StyleID == 0 {
			continue
		}
		pf := f.styles.GetFont(ci.StyleID)
		if pf == nil || pf.Name == "" {
			continue
		}
		key := fontKey{name: pf.Name, size: pf.Size}
		colCount := ci.Max - ci.Min + 1
		counts[key] += colCount
	}

	if len(counts) == 0 {
		return bookDefault
	}

	maxCount := 0
	var best fontKey
	for key, count := range counts {
		if count > maxCount {
			maxCount = count
			best = key
		}
	}
	return FontInfo{Name: best.name, Size: best.size}
}

// ResolveSheetLite は lite モードでのシート名解決
func (f *File) ResolveSheetLite(sheet string) (string, error) {
	if len(f.sheetNames) == 0 {
		return "", fmt.Errorf("ブックにシートがありません")
	}
	if sheet == "" {
		return f.sheetNames[0], nil
	}
	// 名前指定
	if _, ok := f.sheetPaths[sheet]; ok {
		return sheet, nil
	}
	// インデックス指定
	if idx, err := strconv.Atoi(sheet); err == nil {
		if idx >= 0 && idx < len(f.sheetNames) {
			return f.sheetNames[idx], nil
		}
		return "", fmt.Errorf("シートインデックス %d が範囲外です", idx)
	}
	return "", fmt.Errorf("シート %q が見つかりません", sheet)
}

// ResolveTabColor は SheetMeta のタブ色をRGB文字列に解決する
func (f *File) ResolveTabColor(meta *SheetMeta) string {
	if meta.TabColorRGB != "" {
		return normalizeHexColor(meta.TabColorRGB)
	}
	if meta.TabColorTheme != nil {
		return resolveColorLite("", meta.TabColorTheme, meta.TabColorTint, f.theme)
	}
	return ""
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

	f.sheetPaths, f.sheetNames, err = buildSheetPaths(zr)
	if err != nil {
		return err
	}

	return nil
}

// buildSheetPaths は workbook.xml と workbook.xml.rels からシート名→XMLパスのマップとシート名リストを構築する
func buildSheetPaths(zr *zip.ReadCloser) (map[string]string, []string, error) {
	wb, err := readWorkbook(zr, "xl/workbook.xml")
	if err != nil {
		return nil, nil, err
	}

	relsData, err := readZipFile(zr, "xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, nil, err
	}
	var rels xmlRelationships
	if err := xml.Unmarshal(relsData, &rels); err != nil {
		return nil, nil, fmt.Errorf("workbook.xml.rels のパースに失敗: %w", err)
	}

	targets := make(map[string]string, len(rels.Rels))
	for _, r := range rels.Rels {
		targets[r.ID] = r.Target
	}

	paths := make(map[string]string, len(wb.Sheets))
	names := make([]string, 0, len(wb.Sheets))
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
		names = append(names, s.Name)
	}
	return paths, names, nil
}
