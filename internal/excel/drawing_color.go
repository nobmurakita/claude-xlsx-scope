package excel

import (
	"encoding/xml"
	"strconv"
	"strings"
)

// schemeColorIndex は schemeClr の val 属性を内部テーマインデックスに対応付ける。
// bg1/bg2/tx1/tx2 は OOXML 仕様で lt1/lt2/dk1/dk2 のエイリアス。
var schemeColorIndex = map[string]int{
	"dk1":      0,
	"tx1":      0, // dk1 のエイリアス（背景色1 の対）
	"lt1":      1,
	"bg1":      1, // lt1 のエイリアス
	"dk2":      2,
	"tx2":      2, // dk2 のエイリアス
	"lt2":      3,
	"bg2":      3, // lt2 のエイリアス
	"accent1":  4,
	"accent2":  5,
	"accent3":  6,
	"accent4":  7,
	"accent5":  8,
	"accent6":  9,
	"hlink":    10,
	"folHlink": 11,
}

// resolveColorElement は色要素（srgbClr/schemeClr/sysClr/prstClr/scrgbClr）と
// 子の色変換要素（lumMod/lumOff/tint/shade/satMod/satOff 等）を末尾まで読み、
// 解決した hex 色（"#RRGGBB"）を返す。解決できない場合は "" を返す。
//
// 呼び出し元は depth に -1 を加算してこの要素分の EndElement 読み込みをスキップすること。
func (p *drawingParser) resolveColorElement(t xml.StartElement, decoder *xml.Decoder) string {
	base := p.colorElementBase(t)
	cm := collectColorMods(decoder)
	if base == "" {
		return ""
	}
	return cm.applyTo(base)
}

// colorElementBase は色要素の属性からベース色（"#RRGGBB"）を返す。
// 子要素は読まない（呼び出し側で collectColorMods が消費する）。
func (p *drawingParser) colorElementBase(t xml.StartElement) string {
	switch t.Name.Local {
	case "srgbClr":
		return normalizeHexColor(attrVal(t, "val"))
	case "schemeClr":
		return lookupSchemeColor(attrVal(t, "val"), p.theme)
	case "sysClr":
		// lastClr は Excel が解決した実色（OS のシステム色設定に依存しないフォールバック）。
		// Excel が書き出すファイルでは常に付与される想定。
		if c := attrVal(t, "lastClr"); c != "" {
			return normalizeHexColor(c)
		}
		return sysColorByName(attrVal(t, "val"))
	case "prstClr":
		return prstColorByName(attrVal(t, "val"))
	case "scrgbClr":
		// scRGB は線形 RGB を 0-100000 のパーセント表現で持つ。
		return scrgbToHex(attrVal(t, "r"), attrVal(t, "g"), attrVal(t, "b"))
	}
	return ""
}

// lookupSchemeColor は schemeClr の val 名（accent1, bg1 等）からテーマ色を引く。
func lookupSchemeColor(name string, tc *themeColors) string {
	idx, ok := schemeColorIndex[name]
	if !ok {
		return ""
	}
	return resolveThemeIndexScheme(idx, tc)
}

// sysColorByName は sysClr の val から既知のシステム色を返す。
// 実用上 Excel は lastClr を付けるためここに到達するのは稀。最低限の対応に留める。
func sysColorByName(name string) string {
	switch name {
	case "window", "background":
		return "#FFFFFF"
	case "windowText", "menuText", "captionText":
		return "#000000"
	}
	return ""
}

// prstColorByName は prstClr の val（HTML/X11 名）から色を返す。
// OOXML 仕様は 147 色を定義するが、実ファイルでの出現はごく少数のため
// 基本 16 色＋ Excel が実際に書き出す black/white を対象とする。
// 未知の名前は "" を返す（呼び出し側で fillRef 等にフォールバックする）。
var prstColorMap = map[string]string{
	"black":   "#000000",
	"white":   "#FFFFFF",
	"silver":  "#C0C0C0",
	"gray":    "#808080",
	"grey":    "#808080",
	"maroon":  "#800000",
	"red":     "#FF0000",
	"purple":  "#800080",
	"fuchsia": "#FF00FF",
	"green":   "#008000",
	"lime":    "#00FF00",
	"olive":   "#808000",
	"yellow":  "#FFFF00",
	"navy":    "#000080",
	"blue":    "#0000FF",
	"teal":    "#008080",
	"aqua":    "#00FFFF",
	"cyan":    "#00FFFF",
	"magenta": "#FF00FF",
	"orange":  "#FFA500",
	"pink":    "#FFC0CB",
}

func prstColorByName(name string) string {
	return prstColorMap[strings.ToLower(name)]
}

// scrgbToHex は scRGB (r/g/b 各 0-100000 のパーセント) を hex 色に変換する。
// 線形 RGB → sRGB の厳密変換ではなく、既存色変換と整合させるため線形値をそのまま正規化する。
func scrgbToHex(rs, gs, bs string) string {
	r, ok1 := scrgbComponent(rs)
	g, ok2 := scrgbComponent(gs)
	b, ok3 := scrgbComponent(bs)
	if !ok1 || !ok2 || !ok3 {
		return ""
	}
	return formatHexRGB(r, g, b)
}

func scrgbComponent(s string) (float64, bool) {
	if s == "" {
		return 0, false
	}
	v, err := strconv.Atoi(s)
	if err != nil {
		return 0, false
	}
	f := float64(v) / drawingMLPercentUnit
	if f < 0 {
		f = 0
	}
	if f > 1 {
		f = 1
	}
	return f, true
}

