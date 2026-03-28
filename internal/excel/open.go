package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"path/filepath"
	"strconv"
	"strings"
)

// FontInfo はフォントの基本情報
type FontInfo struct {
	Name string  `json:"name"`
	Size float64 `json:"size"`
}

// File はオープンしたExcelファイルを表す。
// OpenFile() で生成し、使用後は Close() で解放する。
type File struct {
	Name string // ファイル名（パス除去済み）
	path string
	zr   *zip.ReadCloser

	// ZIP 内 XML から自前でパー��したデータ
	sharedStrings *sharedStrings
	sheetPaths    map[string]string // シート名 → ZIP内のXMLパス
	sheetNames    []string          // シート名（workbook.xml の順序）
	styles        *styleSheet
	theme         *themeColors
}

// OpenFile はExcelファイルを開き、メタデータをZIPから直接パースする。
func OpenFile(path string) (result *File, retErr error) {
	ext := strings.ToLower(filepath.Ext(path))
	if ext != ".xlsx" && ext != ".xlsm" {
		return nil, fmt.Errorf(".xlsx / .xlsm 形式のみ対応しています")
	}

	zr, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer func() {
		if retErr != nil {
			zr.Close()
		}
	}()

	f := &File{
		Name: filepath.Base(path),
		path: path,
		zr:   zr,
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
	stylesData, err := readZipFile(zr, "xl/styles.xml")
	if err != nil {
		return nil, fmt.Errorf("styles.xml の読み込みに失敗: %w", err)
	}
	f.styles, err = parseStyleSheet(stylesData)
	if err != nil {
		return nil, fmt.Errorf("styles.xml のパースに失敗: %w", err)
	}

	// theme1.xml（存在しなくてもエラーにしない）
	themeData, err := readZipFile(zr, "xl/theme/theme1.xml")
	if err == nil {
		f.theme = parseThemeColors(themeData)
	}

	return f, nil
}

// Close は File が保持する ZIP リーダーを閉じる
func (f *File) Close() error {
	if f.zr != nil {
		return f.zr.Close()
	}
	return nil
}

// LoadSheetMeta はワークシートXMLからメタデータを直接パースする
func (f *File) LoadSheetMeta(sheet string) (*SheetMeta, error) {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return nil, fmt.Errorf("シート %q が見つかりません", sheet)
	}
	return LoadSheetMeta(f.zr, xmlPath)
}

// LoadSheetRels はシートのリレーションを読む
func (f *File) LoadSheetRels(sheet string) map[string]string {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return nil
	}
	return LoadSheetRelsFromZip(f.zr, xmlPath)
}

// LoadDimension はシートの dimension を高速取得する（XML先頭のみ読む）
func (f *File) LoadDimension(sheet string) string {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return ""
	}
	return LoadDimensionOnly(f.zr, xmlPath)
}

const (
	defaultFontName = "Calibri"
	defaultFontSize = 11
)

// DetectDefaultFont は styles.xml からデフォルトフォントを返す
func (f *File) DetectDefaultFont() FontInfo {
	if f.styles == nil {
		return FontInfo{Name: defaultFontName, Size: defaultFontSize}
	}
	name := f.styles.DefaultFontName()
	if name == "" {
		name = defaultFontName
	}
	return FontInfo{Name: name, Size: defaultFontSize}
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

// ResolveSheet はシート名を解決する
func (f *File) ResolveSheet(sheet string) (string, error) {
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

// StyleByID は自前パーサーから全スタイル情報を返す
func (f *File) StyleByID(styleID int, defaultFont FontInfo) (*FontObj, *FillObj, *BorderObj, *AlignmentObj) {
	if f.styles == nil {
		return nil, nil, nil, nil
	}
	font := buildFontObjFromParsed(f.styles.GetFont(styleID), defaultFont, f.theme)
	fill := buildFillObjFromParsed(f.styles.GetFill(styleID), f.theme)
	border := buildBorderObjFromParsed(f.styles.GetBorder(styleID))
	alignment := buildAlignmentObjFromParsed(f.styles.GetAlignment(styleID))
	return font, fill, border, alignment
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
