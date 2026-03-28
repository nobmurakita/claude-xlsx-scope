package excel

import (
	"encoding/xml"
)

// themeColors はテーマカラーパレット（12色）
type themeColors struct {
	colors []string // インデックス0-11のRGB色 (#RRGGBB)
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

	return tc
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

// Get はテーマカラーインデックスからRGB色を返す
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
