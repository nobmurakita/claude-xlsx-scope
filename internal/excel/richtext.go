package excel

import "github.com/xuri/excelize/v2"

// RichTextRun は出力用のリッチテキストラン
type RichTextRun struct {
	Text string   `json:"text"`
	Font *FontObj `json:"font,omitempty"`
}

// GetRichText はセルのリッチテキスト情報を取得する。
// リッチテキストがない場合は nil を返す。
func (f *File) GetRichText(sheet string, col, row int, cellFont *FontObj, defaultFont FontInfo) []RichTextRun {
	axis := CellRef(col, row)
	runs, err := f.File.GetCellRichText(sheet, axis)
	if err != nil || len(runs) <= 1 {
		return nil
	}

	baseFontInfo := richTextBaseFont(cellFont, defaultFont)

	result := make([]RichTextRun, 0, len(runs))
	for _, run := range runs {
		r := RichTextRun{Text: run.Text}
		if run.Font != nil {
			r.Font = richTextFontDiff(run.Font, baseFontInfo, f.File)
		}
		result = append(result, r)
	}
	return result
}

// GetRichTextLite は共有文字列テーブルからリッチテキスト情報を取得する（lite モード用）。
// sharedStrIdx は RawCell.SharedStrIdx。-1 または非共有文字列の場合は nil を返す。
func (f *File) GetRichTextLite(sharedStrIdx int, cellFont *FontObj, defaultFont FontInfo) []RichTextRun {
	if f.sharedStrings == nil || sharedStrIdx < 0 {
		return nil
	}
	rawRuns := f.sharedStrings.GetRichTextRuns(sharedStrIdx)
	if len(rawRuns) <= 1 {
		return nil
	}

	baseFontInfo := richTextBaseFont(cellFont, defaultFont)

	result := make([]RichTextRun, 0, len(rawRuns))
	for _, run := range rawRuns {
		r := RichTextRun{Text: run.Text}
		if run.Font != nil {
			r.Font = richTextFontDiffFromParsed(run.Font, baseFontInfo, f.theme)
		}
		result = append(result, r)
	}
	return result
}

func richTextBaseFont(cellFont *FontObj, defaultFont FontInfo) FontInfo {
	base := defaultFont
	if cellFont != nil {
		if cellFont.Name != "" {
			base.Name = cellFont.Name
		}
		if cellFont.Size != 0 {
			base.Size = cellFont.Size
		}
	}
	return base
}

func richTextFontDiff(font *excelize.Font, base FontInfo, ef *excelize.File) *FontObj {
	obj := &FontObj{}
	if font.Family != "" && font.Family != base.Name {
		obj.Name = font.Family
	}
	if font.Size != 0 && font.Size != base.Size {
		obj.Size = font.Size
	}
	if font.Bold {
		obj.Bold = true
	}
	if font.Italic {
		obj.Italic = true
	}
	if font.Strike {
		obj.Strikethrough = true
	}
	if font.Underline != "" {
		obj.Underline = font.Underline
	}
	color := ResolveColor(font.Color, font.ColorTheme, font.ColorTint, ef)
	if color != "" && color != "#000000" {
		obj.Color = color
	}
	if obj.IsEmpty() {
		return nil
	}
	return obj
}

func richTextFontDiffFromParsed(font *parsedFont, base FontInfo, tc *themeColors) *FontObj {
	obj := &FontObj{}
	if font.Name != "" && font.Name != base.Name {
		obj.Name = font.Name
	}
	if font.Size != 0 && font.Size != base.Size {
		obj.Size = font.Size
	}
	obj.Bold = font.Bold
	obj.Italic = font.Italic
	obj.Strikethrough = font.Strike
	obj.Underline = font.Underline

	color := resolveColorLite(font.Color, font.ColorTheme, font.ColorTint, tc)
	if color != "" && color != "#000000" {
		obj.Color = color
	}
	if obj.IsEmpty() {
		return nil
	}
	return obj
}
