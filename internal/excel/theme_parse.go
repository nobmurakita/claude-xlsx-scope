package excel

import (
	"bytes"
	"encoding/xml"
)

// themeColors はテーマカラーパレット（12色）と書式スキーム（fmtScheme）
type themeColors struct {
	colors []string // インデックス0-11のRGB色 (#RRGGBB)

	// fmtScheme の塗り/線スタイル一覧。各要素は phClr に適用する色変換（出現順）。
	// fillRef/lnRef の idx（1始まり）→ slice[idx-1]。グラデーションは中央ストップで代表する。
	fillStyles [][]drawingColorOp
	lineStyles [][]drawingColorOp
}

// ApplyFillStyle は fillRef の idx が指すテーマ塗りスタイルの色変換を base に適用する
func (tc *themeColors) ApplyFillStyle(idx int, base string) string {
	if tc == nil || idx < 1 || idx > len(tc.fillStyles) {
		return base
	}
	return applyDrawingColorOps(base, tc.fillStyles[idx-1])
}

// ApplyLineStyle は lnRef の idx が指すテーマ線スタイルの色変換を base に適用する
func (tc *themeColors) ApplyLineStyle(idx int, base string) string {
	if tc == nil || idx < 1 || idx > len(tc.lineStyles) {
		return base
	}
	return applyDrawingColorOps(base, tc.lineStyles[idx-1])
}

// XML 構造体（テーマパース用）

type xmlTheme struct {
	XMLName       xml.Name          `xml:"theme"`
	ThemeElements *xmlThemeElements `xml:"themeElements"`
}

type xmlThemeElements struct {
	ClrScheme *xmlClrScheme `xml:"clrScheme"`
}

type xmlClrScheme struct {
	Dk1      xmlThemeColor `xml:"dk1"`
	Lt1      xmlThemeColor `xml:"lt1"`
	Dk2      xmlThemeColor `xml:"dk2"`
	Lt2      xmlThemeColor `xml:"lt2"`
	Accent1  xmlThemeColor `xml:"accent1"`
	Accent2  xmlThemeColor `xml:"accent2"`
	Accent3  xmlThemeColor `xml:"accent3"`
	Accent4  xmlThemeColor `xml:"accent4"`
	Accent5  xmlThemeColor `xml:"accent5"`
	Accent6  xmlThemeColor `xml:"accent6"`
	Hlink    xmlThemeColor `xml:"hlink"`
	FolHlink xmlThemeColor `xml:"folHlink"`
}

type xmlThemeColor struct {
	SysClr  *xmlSysClr  `xml:"sysClr"`
	SrgbClr *xmlSrgbClr `xml:"srgbClr"`
}

type xmlSysClr struct {
	Val     string `xml:"val,attr"`
	LastClr string `xml:"lastClr,attr"`
}

type xmlSrgbClr struct {
	Val string `xml:"val,attr"`
}

// parseThemeColors は theme1.xml のバイトデータからテーマカラーを解析する
func parseThemeColors(data []byte) *themeColors {
	var theme xmlTheme
	if err := xml.Unmarshal(data, &theme); err != nil {
		return &themeColors{colors: make([]string, 12)}
	}

	tc := &themeColors{colors: make([]string, 12)}

	if theme.ThemeElements == nil || theme.ThemeElements.ClrScheme == nil {
		return tc
	}

	cs := theme.ThemeElements.ClrScheme

	// 格納順序: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
	entries := []xmlThemeColor{
		cs.Dk1, cs.Lt1, cs.Dk2, cs.Lt2,
		cs.Accent1, cs.Accent2, cs.Accent3, cs.Accent4, cs.Accent5, cs.Accent6,
		cs.Hlink, cs.FolHlink,
	}

	for i, e := range entries {
		tc.colors[i] = extractThemeColorValue(e)
	}

	tc.fillStyles, tc.lineStyles = parseThemeStyleLists(data)

	return tc
}

// parseThemeStyleLists は fmtScheme の fillStyleLst / lnStyleLst を解析し、
// 各スタイルエントリの代表色変換（phClr に適用する ops）を返す。
func parseThemeStyleLists(data []byte) (fills, lines [][]drawingColorOp) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "fillStyleLst":
			fills = parseStyleEntries(dec, "fillStyleLst")
		case "lnStyleLst":
			lines = parseStyleEntries(dec, "lnStyleLst")
		}
	}
	return
}

// parseStyleEntries はスタイル一覧（fillStyleLst/lnStyleLst）の各エントリを読み、
// エントリごとの代表色変換を返す。
func parseStyleEntries(dec *xml.Decoder, listName string) [][]drawingColorOp {
	var entries [][]drawingColorOp
	for {
		tok, err := dec.Token()
		if err != nil {
			return entries
		}
		switch t := tok.(type) {
		case xml.StartElement:
			entries = append(entries, extractEntryOps(dec, t))
		case xml.EndElement:
			if t.Name.Local == listName {
				return entries
			}
		}
	}
}

// extractEntryOps は 1 つのスタイルエントリ（solidFill/gradFill/ln 等）を末尾まで読み、
// phClr に適用する代表色変換を返す。gradFill は pos=50000 に最も近いストップを代表とする。
func extractEntryOps(dec *xml.Decoder, start xml.StartElement) []drawingColorOp {
	isGrad := start.Name.Local == "gradFill"
	depth := 1
	var bestOps []drawingColorOp
	bestPos := -1
	haveBest := false
	var curOps []drawingColorOp
	inClr := false
	curPos := 0

	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "gs":
				curPos = safeAtoi(attrVal(t, "pos"))
			case "schemeClr", "srgbClr":
				inClr = true
				curOps = nil
			case "lumMod", "lumOff", "satMod", "satOff", "tint", "shade":
				if inClr {
					curOps = append(curOps, drawingColorOp{
						kind: t.Name.Local,
						val:  float64(safeAtoi(attrVal(t, "val"))) / drawingMLPercentUnit,
					})
				}
			}
		case xml.EndElement:
			depth--
			switch t.Name.Local {
			case "schemeClr", "srgbClr":
				inClr = false
				// solidFill / ln 内の最初の色を代表とする
				if !isGrad && !haveBest {
					bestOps = curOps
					haveBest = true
				}
			case "gs":
				// pos=50000 に最も近いストップを代表とする
				if isGrad && (bestPos < 0 || absInt(curPos-50000) < absInt(bestPos-50000)) {
					bestPos = curPos
					bestOps = curOps
					haveBest = true
				}
			}
		}
	}
	return bestOps
}

func absInt(n int) int {
	if n < 0 {
		return -n
	}
	return n
}

// extractThemeColorValue はテーマカラー要素からRGB色を抽出する
func extractThemeColorValue(c xmlThemeColor) string {
	if c.SrgbClr != nil && c.SrgbClr.Val != "" {
		return normalizeHexColor(c.SrgbClr.Val)
	}
	if c.SysClr != nil && c.SysClr.LastClr != "" {
		return normalizeHexColor(c.SysClr.LastClr)
	}
	return ""
}

// themeIndexMap はテーマカラーインデックスを内部配列インデックスにマッピングする。
// theme=0 → lt1 (index 1)
// theme=1 → dk1 (index 0)
// theme=2 → lt2 (index 3)
// theme=3 → dk2 (index 2)
// theme=4-11 → accent1-6, hlink, folHlink (index 4-11)
var themeIndexMap = map[int]int{
	0: 1, // lt1
	1: 0, // dk1
	2: 3, // lt2
	3: 2, // dk2
}

// Get はテーマカラーインデックスからRGB色を返す。
// styles.xml のセル用テーマインデックス（0=lt1, 1=dk1, 2=lt2, 3=dk2 の入れ替え順）を前提とし、
// themeIndexMap でスワップしてから格納配列を引く。
func (tc *themeColors) Get(index int) string {
	if tc == nil {
		return ""
	}
	actual, ok := themeIndexMap[index]
	if !ok {
		actual = index
	}
	if actual < 0 || actual >= len(tc.colors) {
		return ""
	}
	return tc.colors[actual]
}

// GetScheme は DrawingML スキームカラー用に、スワップなしで色を返す。
// schemeColorIndex が既に自然順（dk1=0, lt1=1, dk2=2, lt2=3, accent1-6=4-9...）を渡すため、
// セル用テーマインデックスのスワップ（themeIndexMap）は適用してはならない。
func (tc *themeColors) GetScheme(index int) string {
	if tc == nil {
		return ""
	}
	if index < 0 || index >= len(tc.colors) {
		return ""
	}
	return tc.colors[index]
}
