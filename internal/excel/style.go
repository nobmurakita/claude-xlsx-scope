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
