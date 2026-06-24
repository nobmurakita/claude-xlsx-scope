package excel

import (
	"math"
	"strconv"
	"strings"
)

// ── 型定義 ──────────────────────────────────────────

// tokenKind はフォーマットトークンの種別
type tokenKind int

const (
	tkLiteral      tokenKind = iota // リテラル文字列
	tkDateYear                      // yyyy, yy
	tkDateMonth                     // mm, m (日付文脈)
	tkDateDay                       // dd, d, ddd, dddd
	tkTimeHour                      // hh, h
	tkTimeMinute                    // mm, m (時刻文脈)
	tkTimeSecond                    // ss, s
	tkAmPm                          // AM/PM 等
	tkElapsedHour                   // [h]
	tkElapsedMin                    // [m]
	tkElapsedSec                    // [s]
	tkDigitZero                     // 0
	tkDigitHash                     // #
	tkDigitSpace                    // ?
	tkDecimalPoint                  // .
	tkThousandSep                   // , (桁区切り)
	tkPercent                       // %
	tkExponent                      // E+, E-, e+, e-
	tkSlash                         // / (分数区切り)
	tkAt                            // @ (テキスト挿入)
	tkGeneral                       // General
)

// fmtToken はパース済みの1トークン
type fmtToken struct {
	kind  tokenKind
	raw   string // 元のトークン文字列（リテラルの場合はデコード済み文字列）
	width int    // 桁数（yyyy=4, yy=2, hh=2, h=1 等）
}

// sectionKind はセクションの種別
type sectionKind int

const (
	secNumeric sectionKind = iota
	secDate
	secText
	secGeneral
)

// sectionCondition はセクション条件 [>=100] 等
type sectionCondition struct {
	op    string
	value float64
}

// fmtSection はセミコロン区切りの1セクション
type fmtSection struct {
	kind      sectionKind
	tokens    []fmtToken
	condition *sectionCondition
	hasAmPm   bool
	// 数値フォーマット用の事前計算値
	intDigits    int  // 整数部の 0 桁数
	decDigits    int  // 小数部の桁数（0, #, ? の総数）
	useThousand  bool // 桁区切りあり
	scaleDivisor int  // 末尾カンマの個数（÷1000^n）
	percentCount int  // % の個数
	// 分数用
	hasFraction     bool
	fracWholePart   bool // 整数部があるか（"# ?/?" の "#" 部分）
	fracDenomDigits int  // 分母の桁数（?/?=1, ??/??=2）
	fracFixedDenom  int  // 固定分母（0なら可変）
	// 指数用
	hasExponent bool
	expSign     string // "+" or "-"
	expDigits   int    // 指数部の桁数
}

// parsedNumFmt はフォーマット文字列全体のパース結果
type parsedNumFmt struct {
	sections []fmtSection
}

// ── キャッシュ ──────────────────────────────────────

// numFmtCache はパース済みフォーマットのキャッシュ。
// 本ツールは単一 goroutine で動作するため、排他制御は行わない。
var numFmtCache = make(map[string]*parsedNumFmt)

func getOrParseNumFmt(format string) *parsedNumFmt {
	if p, ok := numFmtCache[format]; ok {
		return p
	}
	p := parseNumFmt(format)
	numFmtCache[format] = p
	return p
}

// ── 公開API ─────────────────────────────────────────

// FormatNumericValue はフォーマット文字列を使って数値をフォーマットする。
// numFmtStr が空または "General" の場合は空文字列を返す。
func FormatNumericValue(numFmtStr string, value float64) string {
	if numFmtStr == "" {
		return ""
	}
	lower := strings.ToLower(strings.TrimSpace(numFmtStr))
	if lower == "general" || lower == "@" {
		return ""
	}

	p := getOrParseNumFmt(numFmtStr)
	sec, val := p.selectSection(value)
	if sec == nil {
		return ""
	}

	switch sec.kind {
	case secDate:
		return formatDate(sec, value) // 日付は元のシリアル値を使う
	case secNumeric:
		return formatNumber(sec, val)
	case secText:
		return formatText(sec, strconv.FormatFloat(value, 'f', -1, 64))
	case secGeneral:
		return ""
	}
	return ""
}

// ── セクション選択 ──────────────────────────────────

func (p *parsedNumFmt) selectSection(value float64) (*fmtSection, float64) {
	n := len(p.sections)
	if n == 0 {
		return nil, value
	}

	// 条件付きセクションがある場合
	if n >= 2 && (p.sections[0].condition != nil || p.sections[1].condition != nil) {
		for i := 0; i < n && i < 2; i++ {
			if p.sections[i].condition != nil && evalCondition(p.sections[i].condition, value) {
				return &p.sections[i], value
			}
		}
		// どの条件にもマッチしない → 3番目のセクション（あれば）
		if n >= 3 {
			return &p.sections[2], value
		}
		return &p.sections[n-1], value
	}

	// 条件なしの場合（標準的なセクション分岐）
	switch {
	case n == 1:
		return &p.sections[0], value
	case n == 2:
		if value >= 0 {
			return &p.sections[0], value
		}
		return &p.sections[1], -value // 負のセクションでは符号を反転
	case n >= 3:
		if value > 0 {
			return &p.sections[0], value
		}
		if value < 0 {
			return &p.sections[1], -value
		}
		return &p.sections[2], value // ゼロ
	}
	return &p.sections[0], value
}

func evalCondition(c *sectionCondition, value float64) bool {
	switch c.op {
	case ">":
		return value > c.value
	case ">=":
		return value >= c.value
	case "<":
		return value < c.value
	case "<=":
		return value <= c.value
	case "=":
		return math.Abs(value-c.value) < 1e-9
	case "<>":
		return math.Abs(value-c.value) >= 1e-9
	}
	return false
}

// ── テキストフォーマッター ───────────────────────────

func formatText(sec *fmtSection, text string) string {
	var buf strings.Builder
	for _, t := range sec.tokens {
		switch t.kind {
		case tkAt:
			buf.WriteString(text)
		case tkLiteral:
			buf.WriteString(t.raw)
		}
	}
	return buf.String()
}
