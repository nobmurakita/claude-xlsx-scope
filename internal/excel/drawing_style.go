package excel

import (
	"encoding/xml"
	"fmt"
	"log"
	"math"
	"strconv"
	"strings"
)

// DrawingML の単位変換定数
const (
	emuPerPixel          = 9525    // 1px = 9525 EMU
	emuPerPoint          = 12700   // 1pt = 12700 EMU
	drawingMLPercentUnit = 100000  // DrawingML の色変換パーセント単位
	drawingMLRotUnit     = 60000   // DrawingML の回転角度単位（1度 = 60000）
	drawingMLFontUnit    = 100     // DrawingML のフォントサイズ単位（100分の1ポイント）
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

// colorMods は DrawingML の色変換パラメータ
type colorMods struct {
	lumMod  float64
	lumOff  float64
	tint    float64
	hasTint bool
}

// collectColorMods は decoder から lumMod, lumOff, tint, shade を収集する
func collectColorMods(decoder *xml.Decoder) colorMods {
	cm := colorMods{lumMod: 1.0}
	depth := 1
	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			log.Printf("[WARN] collectColorMods: XMLトークン読み取りに失敗: %v", err)
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "lumMod":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					cm.lumMod = float64(n) / drawingMLPercentUnit
				}
			case "lumOff":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					cm.lumOff = float64(n) / drawingMLPercentUnit
				}
			case "tint":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					cm.tint = float64(n) / drawingMLPercentUnit
					cm.hasTint = true
				}
			case "shade":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					cm.tint = -(1.0 - float64(n)/drawingMLPercentUnit)
					cm.hasTint = true
				}
			}
		case xml.EndElement:
			depth--
		}
	}
	return cm
}

// applyTo は色変換パラメータをベースカラーに適用する
func (cm colorMods) applyTo(base string) string {
	if cm.hasTint {
		return applyTint(base, cm.tint)
	}
	if cm.lumMod != 1.0 || cm.lumOff != 0 {
		return applyLuminance(base, cm.lumMod, cm.lumOff)
	}
	return base
}

// resolveSchemeColor はスキームカラーを解決し、子の色変換要素まで消費する
func (p *drawingParser) resolveSchemeColor(scheme string, decoder *xml.Decoder, startDepth int) string {
	idx, ok := schemeColorIndex[scheme]
	base := ""
	if ok && p.theme != nil {
		base = p.theme.Get(idx)
	}

	cm := collectColorMods(decoder)

	if base == "" {
		return ""
	}
	return cm.applyTo(base)
}

// applyColorMods は srgbClr の子要素（alpha 等）を消費し、色を返す
func (p *drawingParser) applyColorMods(decoder *xml.Decoder, startDepth int, color string) string {
	clr := normalizeHexColor(color)
	cm := collectColorMods(decoder)
	return cm.applyTo(clr)
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

// updateArrow は矢印の方向を更新する
func updateArrow(current *string, headOrTail string, arrowType string) {
	if arrowType == "" || arrowType == "none" {
		return
	}
	switch headOrTail {
	case "head":
		if *current == "end" {
			*current = "both"
		} else {
			*current = "start"
		}
	case "tail":
		if *current == "start" {
			*current = "both"
		} else {
			*current = "end"
		}
	}
}

// parseLineWidth は ln 要素から LineStyle を初期化し線幅を設定する
func parseLineWidth(t xml.StartElement) *LineStyle {
	ls := &LineStyle{}
	for _, attr := range t.Attr {
		if attr.Name.Local == "w" {
			w, _ := strconv.Atoi(attr.Value)
			ls.Width = math.Round(float64(w)/emuPerPoint*100) / 100
		}
	}
	return ls
}

// finalizeLineStyle は LineStyle を最終形に整える（style 未設定なら "solid"）
func finalizeLineStyle(ls *LineStyle) *LineStyle {
	if ls == nil {
		return nil
	}
	if ls.Color == "" && ls.Style == "" && ls.Width == 0 {
		return nil
	}
	if ls.Style == "" && (ls.Color != "" || ls.Width > 0) {
		ls.Style = "solid"
	}
	return ls
}

// newShapeInfo は共通のシェイプ初期化を行う
func (p *drawingParser) newShapeInfo(shapeType string, z int, cell string, groupStack []groupContext) (ShapeInfo, int) {
	id := p.nextID
	p.nextID++
	shape := ShapeInfo{
		ID:   id,
		Type: shapeType,
		Z:    z,
		Cell: cell,
	}
	if len(groupStack) > 0 {
		parentID := groupStack[len(groupStack)-1].seqID
		shape.Parent = &parentID
	}
	return shape, id
}

// registerExcelID は Excel ID から連番 ID へのマッピングを登録する
func (p *drawingParser) registerExcelID(excelID, seqID int) {
	if excelID > 0 {
		p.excelIDMap[excelID] = seqID
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
			log.Printf("[WARN] skipElement: XMLトークン読み取りに失敗: %v", err)
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
			font.Size = float64(sz) / drawingMLFontUnit
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
	return fontObjFromDrawingFont(font, theme)
}

// richTextFontDiffFromDrawing は DrawingML の parsedFont から差分フォントを構築する
func richTextFontDiffFromDrawing(font *parsedFont, theme *themeColors) *FontObj {
	return fontObjFromDrawingFont(font, theme)
}

// fontObjFromDrawingFont は DrawingML の parsedFont から FontObj を構築する共通実装
func fontObjFromDrawingFont(font *parsedFont, theme *themeColors) *FontObj {
	if font == nil {
		return nil
	}
	obj := &FontObj{
		Bold:          font.Bold,
		Italic:        font.Italic,
		Strikethrough: font.Strike,
		Underline:     font.Underline,
	}
	if font.Name != "" {
		obj.Name = font.Name
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
