package excel

import (
	"encoding/xml"
	"fmt"
	"math"
	"strconv"
	"strings"
)

// schemeColorIndex はスキームカラー名をテーマインデックスにマッピングする
var schemeColorIndex = map[string]int{
	"dk1":      0,
	"lt1":      1,
	"dk2":      2,
	"lt2":      3,
	"accent1":  4,
	"accent2":  5,
	"accent3":  6,
	"accent4":  7,
	"accent5":  8,
	"accent6":  9,
	"hlink":    10,
	"folHlink": 11,
}

// resolveSchemeColor はスキームカラーを解決し、子の色変換要素まで消費する
func (p *drawingParser) resolveSchemeColor(scheme string, decoder *xml.Decoder, startDepth int) string {
	idx, ok := schemeColorIndex[scheme]
	base := ""
	if ok && p.theme != nil {
		base = p.theme.Get(idx)
	}

	// 子要素から lumMod, lumOff, tint, shade を収集
	var lumMod, lumOff float64
	lumMod = 1.0 // デフォルト
	var tint float64
	hasTint := false

	depth := 1
	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "lumMod":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					lumMod = float64(n) / 100000.0
				}
			case "lumOff":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					lumOff = float64(n) / 100000.0
				}
			case "tint":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					tint = float64(n) / 100000.0
					hasTint = true
				}
			case "shade":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					tint = -(1.0 - float64(n)/100000.0)
					hasTint = true
				}
			}
		case xml.EndElement:
			depth--
		}
	}

	if base == "" {
		return ""
	}

	// tint を適用
	if hasTint {
		return applyTint(base, tint)
	}

	// lumMod/lumOff を適用
	if lumMod != 1.0 || lumOff != 0 {
		return applyLuminance(base, lumMod, lumOff)
	}

	return base
}

// applyColorMods は srgbClr の子要素（alpha 等）を消費し、色を返す
func (p *drawingParser) applyColorMods(decoder *xml.Decoder, startDepth int, color string) string {
	clr := normalizeHexColor(color)

	var lumMod, lumOff float64
	lumMod = 1.0
	var tint float64
	hasTint := false

	depth := 1
	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "lumMod":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					lumMod = float64(n) / 100000.0
				}
			case "lumOff":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					lumOff = float64(n) / 100000.0
				}
			case "tint":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					tint = float64(n) / 100000.0
					hasTint = true
				}
			case "shade":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					tint = -(1.0 - float64(n)/100000.0)
					hasTint = true
				}
			}
		case xml.EndElement:
			depth--
		}
	}

	if hasTint {
		return applyTint(clr, tint)
	}
	if lumMod != 1.0 || lumOff != 0 {
		return applyLuminance(clr, lumMod, lumOff)
	}
	return clr
}

// applyLuminance は lumMod/lumOff を適用する
func applyLuminance(hex string, lumMod, lumOff float64) string {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return "#" + strings.ToUpper(hex)
	}
	r, _ := strconv.ParseInt(hex[0:2], 16, 32)
	g, _ := strconv.ParseInt(hex[2:4], 16, 32)
	b, _ := strconv.ParseInt(hex[4:6], 16, 32)

	// HSL に変換して luminance を調整
	h, s, l := rgbToHSL(float64(r)/255, float64(g)/255, float64(b)/255)
	l = l*lumMod + lumOff
	if l < 0 {
		l = 0
	}
	if l > 1 {
		l = 1
	}
	rr, gg, bb := hslToRGB(h, s, l)
	return fmt.Sprintf("#%02X%02X%02X", int(math.Round(rr*255)), int(math.Round(gg*255)), int(math.Round(bb*255)))
}

// assignColor は解決済み色を適切なターゲットに割り当てる
func (p *drawingParser) assignColor(color, ctx string, shapeFill *string, lineStyle *LineStyle, runFont, defFont *parsedFont) {
	if color == "" {
		return
	}
	switch ctx {
	case "sp":
		if shapeFill != nil {
			*shapeFill = color
		}
	case "ln":
		if lineStyle != nil {
			lineStyle.Color = color
		}
	case "rPr":
		if runFont != nil {
			runFont.Color = color
		}
	case "defRPr":
		if defFont != nil {
			defFont.Color = color
		}
	}
}

// attrVal は StartElement から指定属性の値を返す
func attrVal(t xml.StartElement, name string) string {
	for _, attr := range t.Attr {
		if attr.Name.Local == name {
			return attr.Value
		}
	}
	return ""
}

// skipElement は現在の要素を末尾まで読み飛ばす
func skipElement(decoder *xml.Decoder) {
	depth := 1
	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			return
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
}

// parseDrawingFontAttrs は DrawingML の rPr/defRPr 属性からフォント情報を取得する
func parseDrawingFontAttrs(t xml.StartElement, font *parsedFont) {
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "sz":
			// 100分の1ポイント単位
			sz, _ := strconv.Atoi(attr.Value)
			font.Size = float64(sz) / 100
		case "b":
			font.Bold = attr.Value == "1"
		case "i":
			font.Italic = attr.Value == "1"
		case "strike":
			if attr.Value != "" && attr.Value != "noStrike" {
				font.Strike = true
			}
		case "u":
			if attr.Value != "" && attr.Value != "none" {
				font.Underline = attr.Value
			}
		}
	}
}

// buildDrawingFontObj は DrawingML の parsedFont から FontObj を構築する
func buildDrawingFontObj(font *parsedFont, theme *themeColors) *FontObj {
	if font == nil {
		return nil
	}
	obj := &FontObj{
		Name:          font.Name,
		Bold:          font.Bold,
		Italic:        font.Italic,
		Strikethrough: font.Strike,
		Underline:     font.Underline,
	}
	if font.Size != 0 {
		obj.Size = font.Size
	}
	if font.Color != "" {
		obj.Color = font.Color
	} else if font.ColorTheme != nil {
		color := resolveColorLite("", font.ColorTheme, font.ColorTint, theme)
		if color != "" && color != "#000000" {
			obj.Color = color
		}
	}
	if obj.IsEmpty() {
		return nil
	}
	return obj
}

// richTextFontDiffFromDrawing は DrawingML の parsedFont から差分フォントを構築する
func richTextFontDiffFromDrawing(font *parsedFont, theme *themeColors) *FontObj {
	if font == nil {
		return nil
	}
	obj := &FontObj{}
	if font.Name != "" {
		obj.Name = font.Name
	}
	if font.Size != 0 {
		obj.Size = font.Size
	}
	obj.Bold = font.Bold
	obj.Italic = font.Italic
	obj.Strikethrough = font.Strike
	obj.Underline = font.Underline

	if font.Color != "" {
		obj.Color = font.Color
	} else if font.ColorTheme != nil {
		color := resolveColorLite("", font.ColorTheme, font.ColorTint, theme)
		if color != "" && color != "#000000" {
			obj.Color = color
		}
	}
	if obj.IsEmpty() {
		return nil
	}
	return obj
}
