package excel

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// ---------- 色変換ユーティリティ ----------

func normalizeHexColor(c string) string {
	c = strings.TrimPrefix(c, "#")
	if c == "" {
		return ""
	}
	// ARGB（FFxxxxxx）の先頭2バイトを除去
	if len(c) == 8 {
		c = c[2:]
	}
	if len(c) != 6 {
		if len(c) > 6 {
			c = c[len(c)-6:]
		} else {
			return "#" + strings.ToUpper(c)
		}
	}
	return "#" + strings.ToUpper(c)
}

// parseHexRGB は "#RRGGBB" or "RRGGBB" 形式の文字列を 0.0〜1.0 の RGB に変換する
func parseHexRGB(hex string) (r, g, b float64, ok bool) {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return 0, 0, 0, false
	}
	ri, err1 := strconv.ParseUint(hex[0:2], 16, 8)
	gi, err2 := strconv.ParseUint(hex[2:4], 16, 8)
	bi, err3 := strconv.ParseUint(hex[4:6], 16, 8)
	if err1 != nil || err2 != nil || err3 != nil {
		return 0, 0, 0, false
	}
	return float64(ri) / 255.0, float64(gi) / 255.0, float64(bi) / 255.0, true
}

// formatHexRGB は 0.0〜1.0 の RGB を "#RRGGBB" 形式に変換する
func formatHexRGB(r, g, b float64) string {
	return fmt.Sprintf("#%02X%02X%02X",
		int(math.Round(r*255)),
		int(math.Round(g*255)),
		int(math.Round(b*255)),
	)
}

// ---------- HSL 色空間変換 ----------

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

// ---------- 色調整関数 ----------

// applyTint はHEXカラー文字列にtint値（明度調整）を適用する。
// excelize.ThemeColor と同等の処理を行う。
func applyTint(hexColor string, tint float64) string {
	if tint == 0 {
		return hexColor
	}
	r, g, b, ok := parseHexRGB(hexColor)
	if !ok {
		return hexColor
	}

	h, s, l := rgbToHSL(r, g, b)

	if tint < 0 {
		l = l * (1.0 + tint)
	} else {
		l = l*(1.0-tint) + tint
	}
	l = math.Max(0, math.Min(1, l))

	rr, gg, bb := hslToRGB(h, s, l)
	return formatHexRGB(rr, gg, bb)
}

// applyLuminance は lumMod/lumOff を適用する
func applyLuminance(hexColor string, lumMod, lumOff float64) string {
	r, g, b, ok := parseHexRGB(hexColor)
	if !ok {
		return normalizeHexColor(hexColor)
	}

	h, s, l := rgbToHSL(r, g, b)
	l = l*lumMod + lumOff
	if l < 0 {
		l = 0
	}
	if l > 1 {
		l = 1
	}
	rr, gg, bb := hslToRGB(h, s, l)
	return formatHexRGB(rr, gg, bb)
}

// ---------- テーマカラー解決 ----------

// resolveThemeIndex はテーマインデックスからベースカラーを取得する
func resolveThemeIndex(idx int, tc *themeColors) string {
	if tc == nil {
		return ""
	}
	return tc.Get(idx)
}

// resolveColorLite は自前パーサーのテーマカラーを解決する
func resolveColorLite(color string, theme *int, tint float64, tc *themeColors) string {
	if theme != nil {
		base := resolveThemeIndex(*theme, tc)
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
