package excel

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// ResolveColor はexcelizeのスタイル色情報からHEX RGB文字列を解決する。
// テーマカラーの場合はRGBに変換し、tint値がある場合は明度を調整する。
func ResolveColor(color string, theme *int, tint float64, f *excelize.File) string {
	// GetBaseColor でテーマカラー/インデックスカラーをRGBに解決
	baseRGB := f.GetBaseColor(color, 0, theme)
	if baseRGB == "" {
		if color != "" {
			return normalizeHexColor(color)
		}
		if theme != nil {
			return fmt.Sprintf("theme:%d", *theme)
		}
		return ""
	}

	// tint値がある場合は excelize.ThemeColor で明度調整
	if tint != 0 {
		result := excelize.ThemeColor(baseRGB, tint)
		return normalizeHexColor(result)
	}
	return normalizeHexColor(baseRGB)
}

// NormalizeColor はカラー文字列を #RRGGBB 形式に正規化する
func NormalizeColor(c string) string {
	return normalizeHexColor(c)
}

func normalizeHexColor(c string) string {
	c = strings.TrimPrefix(c, "#")
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
