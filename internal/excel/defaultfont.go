package excel

import "github.com/xuri/excelize/v2"

// FontInfo はフォントの基本情報
type FontInfo struct {
	Name string  `json:"name"`
	Size float64 `json:"size"`
}

// DetectDefaultFont はシートのデフォルトフォントを検出する。
// used_range 内の列スタイルからフォント頻度をカウントし、最頻フォントを採用する。
// 列スタイル未設定の場合はブックデフォルトにフォールバック。
func (f *File) DetectDefaultFont(sheet string, usedRange CellRange) FontInfo {
	if usedRange.IsEmpty() {
		return f.getBookDefaultFont()
	}

	type fontKey struct {
		name string
		size float64
	}
	counts := make(map[fontKey]int)
	var firstFont fontKey
	firstCol := -1

	for c := usedRange.StartCol; c <= usedRange.EndCol; c++ {
		colStr := colName(c)
		styleID, err := f.File.GetColStyle(sheet, colStr)
		if err != nil || styleID == 0 {
			continue
		}
		style, err := f.File.GetStyle(styleID)
		if err != nil || style == nil || style.Font == nil {
			continue
		}
		key := fontKey{name: style.Font.Family, size: style.Font.Size}
		if key.name == "" {
			continue
		}
		counts[key]++
		if firstCol == -1 || c < firstCol {
			firstFont = key
			firstCol = c
		}
	}

	if len(counts) == 0 {
		return f.getBookDefaultFont()
	}

	// 最頻フォントを選択（同数なら列インデックスが小さい方）
	maxCount := 0
	var bestFont fontKey
	for key, count := range counts {
		if count > maxCount {
			maxCount = count
			bestFont = key
		}
	}
	// 同数チェック: 最頻と同数のフォントがある場合、列インデックスが最小のものを選ぶ
	if counts[firstFont] == maxCount {
		bestFont = firstFont
	}

	return FontInfo{Name: bestFont.name, Size: bestFont.size}
}

func (f *File) getBookDefaultFont() FontInfo {
	fontName, err := f.File.GetDefaultFont()
	if err != nil || fontName == "" {
		fontName = "Calibri"
	}
	return FontInfo{Name: fontName, Size: 11}
}

// getStyleFont はスタイルIDからフォント情報を取得する
func (f *File) getStyleFont(styleID int) *excelize.Font {
	style, err := f.File.GetStyle(styleID)
	if err != nil || style == nil {
		return nil
	}
	return style.Font
}
