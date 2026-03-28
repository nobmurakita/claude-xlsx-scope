package excel

import (
	"encoding/xml"
	"strconv"
)

// safeAtoi は文字列を int に変換する（エラー時は 0）
func safeAtoi(s string) int {
	n, _ := strconv.Atoi(s)
	return n
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
			// XMLトークン読み取り失敗: スキップ中断
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
