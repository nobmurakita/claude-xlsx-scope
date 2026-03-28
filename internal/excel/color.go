package excel

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// NormalizeColor はカラー文字列を #RRGGBB 形式に正規化する
func NormalizeColor(c string) string {
	return normalizeHexColor(c)
}

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
