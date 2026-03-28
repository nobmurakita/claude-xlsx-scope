package excel

import (
	"github.com/xuri/excelize/v2"
)

// FontObj は出力用のフォント情報
type FontObj struct {
	Name          string `json:"name,omitempty"`
	Size          float64 `json:"size,omitempty"`
	Bold          bool   `json:"bold,omitempty"`
	Italic        bool   `json:"italic,omitempty"`
	Strikethrough bool   `json:"strikethrough,omitempty"`
	Underline     string `json:"underline,omitempty"`
	Color         string `json:"color,omitempty"`
}

// IsEmpty はフォント情報が空かどうかを返す
func (fo *FontObj) IsEmpty() bool {
	return fo.Name == "" && fo.Size == 0 && !fo.Bold && !fo.Italic &&
		!fo.Strikethrough && fo.Underline == "" && fo.Color == ""
}

// FillObj は出力用の塗りつぶし情報
type FillObj struct {
	Color string `json:"color,omitempty"`
}

// BorderEdge は罫線の1辺
type BorderEdge struct {
	Style string `json:"style,omitempty"`
	Color string `json:"color,omitempty"`
}

// BorderObj は出力用の罫線情報
type BorderObj struct {
	Top          *BorderEdge `json:"top,omitempty"`
	Bottom       *BorderEdge `json:"bottom,omitempty"`
	Left         *BorderEdge `json:"left,omitempty"`
	Right        *BorderEdge `json:"right,omitempty"`
	DiagonalUp   *BorderEdge `json:"diagonal_up,omitempty"`
	DiagonalDown *BorderEdge `json:"diagonal_down,omitempty"`
}

// IsEmpty は罫線情報が空かどうかを返す
func (bo *BorderObj) IsEmpty() bool {
	return bo.Top == nil && bo.Bottom == nil && bo.Left == nil &&
		bo.Right == nil && bo.DiagonalUp == nil && bo.DiagonalDown == nil
}

// AlignmentObj は出力用の配置情報
type AlignmentObj struct {
	Horizontal   string `json:"horizontal,omitempty"`
	Vertical     string `json:"vertical,omitempty"`
	Wrap         bool   `json:"wrap,omitempty"`
	Indent       int    `json:"indent,omitempty"`
	TextRotation int    `json:"text_rotation,omitempty"`
	ShrinkToFit  bool   `json:"shrink_to_fit,omitempty"`
}

// IsEmpty は配置情報が空かどうかを返す
func (ao *AlignmentObj) IsEmpty() bool {
	return ao.Horizontal == "" && ao.Vertical == "" && !ao.Wrap &&
		ao.Indent == 0 && ao.TextRotation == 0 && !ao.ShrinkToFit
}

// GetCellStyle はセルのスタイルIDを返す
func (f *File) GetCellStyle(sheet, axis string) (int, error) {
	return f.File.GetCellStyle(sheet, axis)
}

// CellStyle はセルの書式情報を取得する
func (f *File) CellStyle(sheet string, col, row int, defaultFont FontInfo) (*FontObj, *FillObj, *BorderObj, *AlignmentObj, error) {
	axis := CellRef(col, row)
	styleID, err := f.File.GetCellStyle(sheet, axis)
	if err != nil {
		return nil, nil, nil, nil, err
	}
	if styleID == 0 {
		return nil, nil, nil, nil, nil
	}

	return f.StyleByID(styleID, defaultFont)
}

// StyleByID はスタイルIDから書式情報を取得する（キャッシュ用）
func (f *File) StyleByID(styleID int, defaultFont FontInfo) (*FontObj, *FillObj, *BorderObj, *AlignmentObj, error) {
	// lite モード: 自前パーサーを使用
	if f.styles != nil {
		return f.styleByIDLite(styleID, defaultFont), nil, nil, nil, nil
	}
	// excelize モード
	style, err := f.File.GetStyle(styleID)
	if err != nil || style == nil {
		return nil, nil, nil, nil, nil
	}

	font := buildFontObj(style.Font, defaultFont, f.File)
	fill := buildFillObj(style.Fill, f.File)
	border := buildBorderObj(style.Border, f.File)
	alignment := buildAlignmentObj(style.Alignment)

	return font, fill, border, alignment, nil
}

// styleByIDLite は自前パーサーのスタイル情報から FontObj を返す（lite モード用）
// 戻り値は FontObj のみだが、getCellStyleByID で全スタイルを取得するため
// StyleByIDLite を別途用意する
func (f *File) styleByIDLite(styleID int, defaultFont FontInfo) *FontObj {
	pf := f.styles.GetFont(styleID)
	if pf == nil {
		return nil
	}
	return buildFontObjFromParsed(pf, defaultFont, f.theme)
}

// StyleByIDLite は自前パーサーから全スタイル情報を返す（lite モード用）
func (f *File) StyleByIDLite(styleID int, defaultFont FontInfo) (*FontObj, *FillObj, *BorderObj, *AlignmentObj) {
	font := f.styleByIDLite(styleID, defaultFont)
	fill := buildFillObjFromParsed(f.styles.GetFill(styleID), f.theme)
	border := buildBorderObjFromParsed(f.styles.GetBorder(styleID))
	alignment := buildAlignmentObjFromParsed(f.styles.GetAlignment(styleID))
	return font, fill, border, alignment
}

func buildFontObjFromParsed(pf *parsedFont, defaultFont FontInfo, tc *themeColors) *FontObj {
	if pf == nil {
		return nil
	}
	obj := &FontObj{}
	if pf.Name != "" && pf.Name != defaultFont.Name {
		obj.Name = pf.Name
	}
	if pf.Size != 0 && pf.Size != defaultFont.Size {
		obj.Size = pf.Size
	}
	obj.Bold = pf.Bold
	obj.Italic = pf.Italic
	obj.Strikethrough = pf.Strike
	obj.Underline = pf.Underline

	color := resolveColorLite(pf.Color, pf.ColorTheme, pf.ColorTint, tc)
	if color != "" && color != "#000000" {
		obj.Color = color
	}
	if obj.IsEmpty() {
		return nil
	}
	return obj
}

func buildFillObjFromParsed(pf *parsedFill, tc *themeColors) *FillObj {
	if pf == nil || pf.PatternType == "" || pf.PatternType == "none" {
		return nil
	}
	color := ""
	if pf.FgTheme != nil {
		color = resolveColorLite("", pf.FgTheme, pf.FgTint, tc)
	} else if pf.FgColor != "" {
		color = pf.FgColor
		if pf.FgTint != 0 {
			color = applyTint(color, pf.FgTint)
		}
	}
	if color == "" {
		return nil
	}
	return &FillObj{Color: color}
}

func buildBorderObjFromParsed(edges []parsedBorderEdge) *BorderObj {
	if len(edges) == 0 {
		return nil
	}
	obj := &BorderObj{}
	for _, e := range edges {
		edge := &BorderEdge{Style: e.Style}
		color := normalizeHexColor(e.Color)
		if color != "" && color != "#000000" {
			edge.Color = color
		}
		switch e.Type {
		case "top":
			obj.Top = edge
		case "bottom":
			obj.Bottom = edge
		case "left":
			obj.Left = edge
		case "right":
			obj.Right = edge
		case "diagonal":
			// diagonal は diagonalUp と diagonalDown の両方を設定
			obj.DiagonalUp = edge
			obj.DiagonalDown = edge
		}
	}
	if obj.IsEmpty() {
		return nil
	}
	return obj
}

func buildAlignmentObjFromParsed(pa *parsedAlignment) *AlignmentObj {
	if pa == nil {
		return nil
	}
	obj := &AlignmentObj{}
	if pa.Horizontal != "" && pa.Horizontal != "general" {
		obj.Horizontal = pa.Horizontal
	}
	if pa.Vertical != "" && pa.Vertical != "bottom" {
		obj.Vertical = pa.Vertical
	}
	obj.Wrap = pa.WrapText
	obj.Indent = pa.Indent
	obj.TextRotation = pa.TextRotation
	obj.ShrinkToFit = pa.ShrinkToFit
	if obj.IsEmpty() {
		return nil
	}
	return obj
}

// resolveColorLite は自前パーサーのテーマカラーを解決する
func resolveColorLite(color string, theme *int, tint float64, tc *themeColors) string {
	if theme != nil && tc != nil {
		base := tc.Get(*theme)
		if base != "" {
			if tint != 0 {
				return applyTint(base, tint)
			}
			return base
		}
	}
	if color != "" {
		if tint != 0 {
			return applyTint(color, tint)
		}
		return normalizeHexColor(color)
	}
	return ""
}

func buildFontObj(font *excelize.Font, defaultFont FontInfo, ef *excelize.File) *FontObj {
	if font == nil {
		return nil
	}
	obj := &FontObj{}
	if font.Family != "" && font.Family != defaultFont.Name {
		obj.Name = font.Family
	}
	if font.Size != 0 && font.Size != defaultFont.Size {
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

func buildFillObj(fill excelize.Fill, ef *excelize.File) *FillObj {
	// ソリッド塗りつぶし（pattern type）の前景色のみ
	if fill.Type != "pattern" || fill.Pattern == 0 {
		return nil
	}
	if len(fill.Color) == 0 || fill.Color[0] == "" {
		return nil
	}
	color := normalizeHexColor(fill.Color[0])
	if color == "" {
		return nil
	}
	return &FillObj{Color: color}
}

var borderStyleNames = map[int]string{
	1: "thin", 2: "medium", 3: "dashed", 4: "dotted",
	5: "thick", 6: "double", 7: "hair",
	8: "mediumDashed", 9: "dashDot", 10: "mediumDashDot",
	11: "dashDotDot", 12: "mediumDashDotDot", 13: "slantDashDot",
}

func buildBorderObj(borders []excelize.Border, ef *excelize.File) *BorderObj {
	if len(borders) == 0 {
		return nil
	}
	obj := &BorderObj{}
	for _, b := range borders {
		styleName := borderStyleNames[b.Style]
		if styleName == "" {
			continue
		}
		edge := &BorderEdge{Style: styleName}
		color := normalizeHexColor(b.Color)
		if color != "" && color != "#000000" {
			edge.Color = color
		}
		switch b.Type {
		case "top":
			obj.Top = edge
		case "bottom":
			obj.Bottom = edge
		case "left":
			obj.Left = edge
		case "right":
			obj.Right = edge
		case "diagonalUp":
			obj.DiagonalUp = edge
		case "diagonalDown":
			obj.DiagonalDown = edge
		}
	}
	if obj.IsEmpty() {
		return nil
	}
	return obj
}

func buildAlignmentObj(align *excelize.Alignment) *AlignmentObj {
	if align == nil {
		return nil
	}
	obj := &AlignmentObj{}
	if align.Horizontal != "" && align.Horizontal != "general" {
		obj.Horizontal = align.Horizontal
	}
	if align.Vertical != "" && align.Vertical != "bottom" {
		obj.Vertical = align.Vertical
	}
	if align.WrapText {
		obj.Wrap = true
	}
	if align.Indent > 0 {
		obj.Indent = align.Indent
	}
	if align.TextRotation != 0 {
		obj.TextRotation = align.TextRotation
	}
	if align.ShrinkToFit {
		obj.ShrinkToFit = true
	}
	if obj.IsEmpty() {
		return nil
	}
	return obj
}
