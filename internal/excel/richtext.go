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
			r.Font = buildFontObjFromParsed(run.Font, baseFontInfo, f.getTheme())
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

