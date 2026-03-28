package excel

// RichTextRun は出力用のリッチテキストラン
type RichTextRun struct {
	Text string   `json:"text"`
	Font *FontObj `json:"font,omitempty"`
}

// GetRichText は共有文字列テーブルからリッチテキスト情報を取得する。
// sharedStrIdx は RawCell.SharedStrIdx。-1 または非共有文字列の場合は nil を返す。
func (f *File) GetRichText(sharedStrIdx int, cellFont *FontObj, defaultFont FontInfo) []RichTextRun {
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
