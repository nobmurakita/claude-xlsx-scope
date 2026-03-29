package excel

// FontObj は出力用のフォント情報
type FontObj struct {
	Name          string  `json:"name,omitempty"`
	Size          float64 `json:"size,omitempty"`
	Bold          bool    `json:"bold,omitempty"`
	Italic        bool    `json:"italic,omitempty"`
	Strikethrough bool    `json:"strikethrough,omitempty"`
	Underline     string  `json:"underline,omitempty"`
	Color         string  `json:"color,omitempty"`
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
	obj.Strikethrough = pf.Strikethrough
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
		case "diagonal_up":
			obj.DiagonalUp = edge
		case "diagonal_down":
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
