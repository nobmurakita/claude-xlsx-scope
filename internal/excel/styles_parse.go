package excel

import (
	"encoding/xml"
	"math"
	"strconv"
	"strings"
)

// styleSheet は styles.xml から解析したスタイル情報
type styleSheet struct {
	numFmts  map[int]string // numFmtId → formatCode (カスタムフォーマットのみ)
	fonts    []parsedFont
	fills    []parsedFill
	borders  []parsedBorder
	cellXfs  []parsedCellXf // セルスタイル定義（インデックス = styleID）
}

type parsedFont struct {
	Name       string
	Size       float64
	Bold       bool
	Italic     bool
	Strike     bool
	Underline  string
	Color      string // RGB色（ARGBの先頭2バイト除去済み）
	ColorTheme *int
	ColorTint  float64
}

type parsedFill struct {
	PatternType string
	FgColor     string // RGB色
	FgTheme     *int
	FgTint      float64
}

type parsedBorder struct {
	Edges []parsedBorderEdge
}

type parsedBorderEdge struct {
	Type  string // "left", "right", "top", "bottom", "diagonal"
	Style string // "thin", "medium", etc.
	Color string
}

type parsedCellXf struct {
	NumFmtID     int
	FontID       int
	FillID       int
	BorderID     int
	ApplyFont    bool
	ApplyFill    bool
	ApplyBorder  bool
	ApplyAlign   bool
	HasAlignment bool // <alignment> 要素が存在するか
	Alignment    parsedAlignment
}

type parsedAlignment struct {
	Horizontal   string
	Vertical     string
	WrapText     bool
	Indent       int
	TextRotation int
	ShrinkToFit  bool
}

// XML 構造体（パース用）

type xmlStyleSheet struct {
	XMLName xml.Name      `xml:"styleSheet"`
	NumFmts xmlNumFmts    `xml:"numFmts"`
	Fonts   xmlFonts      `xml:"fonts"`
	Fills   xmlFills      `xml:"fills"`
	Borders xmlBorders    `xml:"borders"`
	CellXfs xmlCellXfs    `xml:"cellXfs"`
}

type xmlNumFmts struct {
	NumFmt []xmlNumFmt `xml:"numFmt"`
}

type xmlNumFmt struct {
	NumFmtID   int    `xml:"numFmtId,attr"`
	FormatCode string `xml:"formatCode,attr"`
}

type xmlFonts struct {
	Font []xmlFont `xml:"font"`
}

type xmlFont struct {
	B         *xmlBoolVal   `xml:"b"`
	I         *xmlBoolVal   `xml:"i"`
	Strike    *xmlBoolVal   `xml:"strike"`
	U         *xmlUnderline `xml:"u"`
	Sz        *xmlFloatVal  `xml:"sz"`
	Color     *xmlColor     `xml:"color"`
	Name      *xmlStringVal `xml:"name"`
}

type xmlBoolVal struct {
	Val *string `xml:"val,attr"`
}

type xmlUnderline struct {
	Val string `xml:"val,attr"`
}

type xmlFloatVal struct {
	Val float64 `xml:"val,attr"`
}

type xmlStringVal struct {
	Val string `xml:"val,attr"`
}

type xmlColor struct {
	RGB     string  `xml:"rgb,attr"`
	Theme   *int    `xml:"theme,attr"`
	Indexed *int    `xml:"indexed,attr"`
	Tint    float64 `xml:"tint,attr"`
}

type xmlFills struct {
	Fill []xmlFill `xml:"fill"`
}

type xmlFill struct {
	PatternFill *xmlPatternFill `xml:"patternFill"`
}

type xmlPatternFill struct {
	PatternType string   `xml:"patternType,attr"`
	FgColor     *xmlColor `xml:"fgColor"`
}

type xmlBorders struct {
	Border []xmlBorderDef `xml:"border"`
}

type xmlBorderDef struct {
	Left     *xmlBorderEdge `xml:"left"`
	Right    *xmlBorderEdge `xml:"right"`
	Top      *xmlBorderEdge `xml:"top"`
	Bottom   *xmlBorderEdge `xml:"bottom"`
	Diagonal *xmlBorderEdge `xml:"diagonal"`
}

type xmlBorderEdge struct {
	Style string   `xml:"style,attr"`
	Color *xmlColor `xml:"color"`
}

type xmlCellXfs struct {
	Xf []xmlXf `xml:"xf"`
}

type xmlXf struct {
	NumFmtID    int    `xml:"numFmtId,attr"`
	FontID      int    `xml:"fontId,attr"`
	FillID      int    `xml:"fillId,attr"`
	BorderID    int    `xml:"borderId,attr"`
	ApplyFont   string `xml:"applyFont,attr"`
	ApplyFill   string `xml:"applyFill,attr"`
	ApplyBorder string `xml:"applyBorder,attr"`
	ApplyAlign  string `xml:"applyAlignment,attr"`
	Alignment   *xmlAlignment `xml:"alignment"`
}

type xmlAlignment struct {
	Horizontal   string `xml:"horizontal,attr"`
	Vertical     string `xml:"vertical,attr"`
	WrapText     string `xml:"wrapText,attr"`
	Indent       int    `xml:"indent,attr"`
	TextRotation int    `xml:"textRotation,attr"`
	ShrinkToFit  string `xml:"shrinkToFit,attr"`
}

// indexedColors は Excel 標準の64色パレット
var indexedColors = []string{
	"#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
	"#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
	"#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
	"#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
	"#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
	"#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
	"#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
	"#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
}

// parseStyleSheet は styles.xml のバイトデータを解析する
func parseStyleSheet(data []byte) (*styleSheet, error) {
	var raw xmlStyleSheet
	if err := xml.Unmarshal(data, &raw); err != nil {
		return nil, err
	}

	ss := &styleSheet{
		numFmts: make(map[int]string),
	}

	// numFmts
	for _, nf := range raw.NumFmts.NumFmt {
		ss.numFmts[nf.NumFmtID] = nf.FormatCode
	}

	// fonts
	ss.fonts = make([]parsedFont, len(raw.Fonts.Font))
	for i, f := range raw.Fonts.Font {
		pf := parsedFont{}
		if f.Name != nil {
			pf.Name = f.Name.Val
		}
		if f.Sz != nil {
			pf.Size = f.Sz.Val
		}
		if f.B != nil {
			pf.Bold = xmlBoolTrue(f.B.Val)
		}
		if f.I != nil {
			pf.Italic = xmlBoolTrue(f.I.Val)
		}
		if f.Strike != nil {
			pf.Strike = xmlBoolTrue(f.Strike.Val)
		}
		if f.U != nil {
			if f.U.Val == "" {
				pf.Underline = "single"
			} else {
				pf.Underline = f.U.Val
			}
		}
		if f.Color != nil {
			pf.Color = resolveXMLColor(f.Color)
			pf.ColorTheme = f.Color.Theme
			pf.ColorTint = f.Color.Tint
		}
		ss.fonts[i] = pf
	}

	// fills
	ss.fills = make([]parsedFill, len(raw.Fills.Fill))
	for i, f := range raw.Fills.Fill {
		pf := parsedFill{}
		if f.PatternFill != nil {
			pf.PatternType = f.PatternFill.PatternType
			if f.PatternFill.FgColor != nil {
				pf.FgColor = resolveXMLColor(f.PatternFill.FgColor)
				pf.FgTheme = f.PatternFill.FgColor.Theme
				pf.FgTint = f.PatternFill.FgColor.Tint
			}
		}
		ss.fills[i] = pf
	}

	// borders
	ss.borders = make([]parsedBorder, len(raw.Borders.Border))
	for i, b := range raw.Borders.Border {
		pb := parsedBorder{}
		addEdge := func(edge *xmlBorderEdge, edgeType string) {
			if edge == nil || edge.Style == "" {
				return
			}
			pe := parsedBorderEdge{
				Type:  edgeType,
				Style: edge.Style,
			}
			if edge.Color != nil {
				pe.Color = resolveXMLColor(edge.Color)
			}
			pb.Edges = append(pb.Edges, pe)
		}
		addEdge(b.Left, "left")
		addEdge(b.Right, "right")
		addEdge(b.Top, "top")
		addEdge(b.Bottom, "bottom")
		addEdge(b.Diagonal, "diagonal")
		ss.borders[i] = pb
	}

	// cellXfs
	ss.cellXfs = make([]parsedCellXf, len(raw.CellXfs.Xf))
	for i, xf := range raw.CellXfs.Xf {
		px := parsedCellXf{
			NumFmtID:    xf.NumFmtID,
			FontID:      xf.FontID,
			FillID:      xf.FillID,
			BorderID:    xf.BorderID,
			ApplyFont:   xf.ApplyFont == "1" || strings.EqualFold(xf.ApplyFont, "true"),
			ApplyFill:   xf.ApplyFill == "1" || strings.EqualFold(xf.ApplyFill, "true"),
			ApplyBorder: xf.ApplyBorder == "1" || strings.EqualFold(xf.ApplyBorder, "true"),
			ApplyAlign:  xf.ApplyAlign == "1" || strings.EqualFold(xf.ApplyAlign, "true"),
		}
		if xf.Alignment != nil {
			px.HasAlignment = true
			px.Alignment = parsedAlignment{
				Horizontal:   xf.Alignment.Horizontal,
				Vertical:     xf.Alignment.Vertical,
				WrapText:     xf.Alignment.WrapText == "1" || strings.EqualFold(xf.Alignment.WrapText, "true"),
				Indent:       xf.Alignment.Indent,
				TextRotation: xf.Alignment.TextRotation,
				ShrinkToFit:  xf.Alignment.ShrinkToFit == "1" || strings.EqualFold(xf.Alignment.ShrinkToFit, "true"),
			}
		}
		ss.cellXfs[i] = px
	}

	return ss, nil
}

// xmlBoolTrue は XML のブール属性値を判定する。
// 属性なし（<b/>）の場合は true、"0"/"false" の場合は false。
func xmlBoolTrue(val *string) bool {
	if val == nil {
		return true
	}
	return *val != "0" && !strings.EqualFold(*val, "false")
}

// resolveXMLColor は XML のカラー要素からRGB色文字列を取得する
func resolveXMLColor(c *xmlColor) string {
	if c == nil {
		return ""
	}
	if c.RGB != "" {
		return normalizeHexColor(c.RGB)
	}
	if c.Indexed != nil {
		idx := *c.Indexed
		if idx >= 0 && idx < len(indexedColors) {
			return indexedColors[idx]
		}
	}
	return ""
}

// GetNumFmt は styleID から numFmtId と formatCode を返す
func (ss *styleSheet) GetNumFmt(styleID int) (int, string) {
	if styleID < 0 || styleID >= len(ss.cellXfs) {
		return 0, ""
	}
	xf := ss.cellXfs[styleID]
	code, ok := ss.numFmts[xf.NumFmtID]
	if !ok {
		return xf.NumFmtID, ""
	}
	return xf.NumFmtID, code
}

// GetFont は styleID からフォント情報を返す
func (ss *styleSheet) GetFont(styleID int) *parsedFont {
	if styleID < 0 || styleID >= len(ss.cellXfs) {
		return nil
	}
	xf := ss.cellXfs[styleID]
	if !xf.ApplyFont && xf.FontID == 0 {
		return nil
	}
	if xf.FontID < 0 || xf.FontID >= len(ss.fonts) {
		return nil
	}
	f := ss.fonts[xf.FontID]
	return &f
}

// GetFill は styleID から塗りつぶし情報を返す
func (ss *styleSheet) GetFill(styleID int) *parsedFill {
	if styleID < 0 || styleID >= len(ss.cellXfs) {
		return nil
	}
	xf := ss.cellXfs[styleID]
	if !xf.ApplyFill && xf.FillID == 0 {
		return nil
	}
	if xf.FillID < 0 || xf.FillID >= len(ss.fills) {
		return nil
	}
	f := ss.fills[xf.FillID]
	return &f
}

// GetBorder は styleID から罫線情報を返す
func (ss *styleSheet) GetBorder(styleID int) []parsedBorderEdge {
	if styleID < 0 || styleID >= len(ss.cellXfs) {
		return nil
	}
	xf := ss.cellXfs[styleID]
	if !xf.ApplyBorder && xf.BorderID == 0 {
		return nil
	}
	if xf.BorderID < 0 || xf.BorderID >= len(ss.borders) {
		return nil
	}
	return ss.borders[xf.BorderID].Edges
}

// GetAlignment は styleID から配置情報を返す
func (ss *styleSheet) GetAlignment(styleID int) *parsedAlignment {
	if styleID < 0 || styleID >= len(ss.cellXfs) {
		return nil
	}
	xf := ss.cellXfs[styleID]
	if !xf.ApplyAlign && !xf.HasAlignment {
		return nil
	}
	a := xf.Alignment
	return &a
}

// DefaultFontName はブックのデフォルトフォント名を返す
func (ss *styleSheet) DefaultFontName() string {
	if len(ss.fonts) == 0 {
		return ""
	}
	return ss.fonts[0].Name
}

// applyTint はHEXカラー文字列にtint値（明度調整）を適用する。
// excelize.ThemeColor と同等の処理を行う。
func applyTint(hexColor string, tint float64) string {
	if tint == 0 {
		return hexColor
	}
	hex := strings.TrimPrefix(hexColor, "#")
	if len(hex) != 6 {
		return hexColor
	}
	r, err1 := strconv.ParseUint(hex[0:2], 16, 8)
	g, err2 := strconv.ParseUint(hex[2:4], 16, 8)
	b, err3 := strconv.ParseUint(hex[4:6], 16, 8)
	if err1 != nil || err2 != nil || err3 != nil {
		return hexColor
	}

	h, s, l := rgbToHSL(float64(r)/255.0, float64(g)/255.0, float64(b)/255.0)

	if tint < 0 {
		l = l * (1.0 + tint)
	} else {
		l = l*(1.0-tint) + tint
	}
	l = math.Max(0, math.Min(1, l))

	rr, gg, bb := hslToRGB(h, s, l)
	return "#" + strings.ToUpper(
		toHex(rr)+toHex(gg)+toHex(bb),
	)
}

func toHex(v float64) string {
	n := int(math.Round(v * 255))
	if n < 0 {
		n = 0
	}
	if n > 255 {
		n = 255
	}
	s := strconv.FormatInt(int64(n), 16)
	if len(s) == 1 {
		return "0" + s
	}
	return s
}

// rgbToHSL は RGB (0-1) を HSL (0-1) に変換する
func rgbToHSL(r, g, b float64) (h, s, l float64) {
	max := math.Max(r, math.Max(g, b))
	min := math.Min(r, math.Min(g, b))
	l = (max + min) / 2.0

	if max == min {
		return 0, 0, l
	}

	d := max - min
	if l > 0.5 {
		s = d / (2.0 - max - min)
	} else {
		s = d / (max + min)
	}

	switch max {
	case r:
		h = (g - b) / d
		if g < b {
			h += 6.0
		}
	case g:
		h = (b-r)/d + 2.0
	case b:
		h = (r-g)/d + 4.0
	}
	h /= 6.0
	return h, s, l
}

// hslToRGB は HSL (0-1) を RGB (0-1) に変換する
func hslToRGB(h, s, l float64) (r, g, b float64) {
	if s == 0 {
		return l, l, l
	}

	var q float64
	if l < 0.5 {
		q = l * (1.0 + s)
	} else {
		q = l + s - l*s
	}
	p := 2.0*l - q

	r = hueToRGB(p, q, h+1.0/3.0)
	g = hueToRGB(p, q, h)
	b = hueToRGB(p, q, h-1.0/3.0)
	return r, g, b
}

func hueToRGB(p, q, t float64) float64 {
	if t < 0 {
		t += 1
	}
	if t > 1 {
		t -= 1
	}
	if t < 1.0/6.0 {
		return p + (q-p)*6.0*t
	}
	if t < 1.0/2.0 {
		return q
	}
	if t < 2.0/3.0 {
		return p + (q-p)*(2.0/3.0-t)*6.0
	}
	return p
}
