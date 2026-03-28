package excel

import "strings"

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
