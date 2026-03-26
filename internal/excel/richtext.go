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

	// セルレベルのフォント（差分計算の基準）
	baseFontInfo := defaultFont
	if cellFont != nil {
		if cellFont.Name != "" {
			baseFontInfo.Name = cellFont.Name
		}
		if cellFont.Size != 0 {
			baseFontInfo.Size = cellFont.Size
		}
	}

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
