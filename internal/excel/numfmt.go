package excel

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"unicode/utf8"
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

// ── パーサー ────────────────────────────────────────

func parseNumFmt(format string) *parsedNumFmt {
	parts := splitSections(format)
	p := &parsedNumFmt{sections: make([]fmtSection, len(parts))}
	for i, part := range parts {
		p.sections[i] = parseSection(part)
	}
	return p
}

// splitSections はフォーマット文字列を ; で分割する（クォート内は無視）
func splitSections(format string) []string {
	var parts []string
	var buf strings.Builder
	inQuote := false
	for i := 0; i < len(format); i++ {
		ch := format[i]
		if ch == '"' {
			inQuote = !inQuote
			buf.WriteByte(ch)
		} else if ch == '\\' && i+1 < len(format) {
			buf.WriteByte(ch)
			i++
			buf.WriteByte(format[i])
		} else if ch == ';' && !inQuote {
			parts = append(parts, buf.String())
			buf.Reset()
		} else {
			buf.WriteByte(ch)
		}
	}
	parts = append(parts, buf.String())
	if len(parts) > 4 {
		parts = parts[:4]
	}
	return parts
}

func parseSection(s string) fmtSection {
	tokens, cond := tokenize(s)
	tokens = resolveMonthMinute(tokens)
	sec := fmtSection{tokens: tokens, condition: cond}
	classifySection(&sec)
	if sec.kind == secNumeric {
		precomputeNumeric(&sec)
	}
	return sec
}

// ── トークナイザ ────────────────────────────────────

func tokenize(s string) ([]fmtToken, *sectionCondition) {
	var tokens []fmtToken
	var cond *sectionCondition
	i := 0

	for i < len(s) {
		ch := s[i]

		// クォート文字列 "..."
		if ch == '"' {
			j := i + 1
			for j < len(s) && s[j] != '"' {
				j++
			}
			lit := s[i+1 : j]
			if j < len(s) {
				j++ // 閉じクォートをスキップ
			}
			tokens = append(tokens, fmtToken{kind: tkLiteral, raw: lit})
			i = j
			continue
		}

		// エスケープ \X
		if ch == '\\' && i+1 < len(s) {
			_, size := utf8.DecodeRuneInString(s[i+1:])
			tokens = append(tokens, fmtToken{kind: tkLiteral, raw: s[i+1 : i+1+size]})
			i += 1 + size
			continue
		}

		// _X (スキップ幅 → スペース1つ)
		if ch == '_' && i+1 < len(s) {
			_, size := utf8.DecodeRuneInString(s[i+1:])
			tokens = append(tokens, fmtToken{kind: tkLiteral, raw: " "})
			i += 1 + size
			continue
		}

		// *X (繰り返し文字 → 無視)
		if ch == '*' && i+1 < len(s) {
			_, size := utf8.DecodeRuneInString(s[i+1:])
			i += 1 + size
			continue
		}

		// [...] ブラケット
		if ch == '[' {
			j := strings.IndexByte(s[i:], ']')
			if j < 0 {
				tokens = append(tokens, fmtToken{kind: tkLiteral, raw: string(ch)})
				i++
				continue
			}
			inner := s[i+1 : i+j]
			i += j + 1

			// 経過時間 [h], [m], [s]
			lowerInner := strings.ToLower(inner)
			if lowerInner == "h" || lowerInner == "hh" {
				tokens = append(tokens, fmtToken{kind: tkElapsedHour, raw: inner})
				continue
			}
			if lowerInner == "m" || lowerInner == "mm" {
				tokens = append(tokens, fmtToken{kind: tkElapsedMin, raw: inner})
				continue
			}
			if lowerInner == "s" || lowerInner == "ss" {
				tokens = append(tokens, fmtToken{kind: tkElapsedSec, raw: inner})
				continue
			}

			// 条件 [>=100] 等
			if c := parseCondition(inner); c != nil {
				cond = c
				continue
			}

			// ロケール [$¥-ja-JP] → 通貨記号を抽出
			if strings.HasPrefix(inner, "$") {
				currency := extractCurrency(inner)
				if currency != "" {
					tokens = append(tokens, fmtToken{kind: tkLiteral, raw: currency})
				}
				continue
			}

			// 色指定 [Red], [Color1] 等 → 無視
			// DBNum → 無視
			continue
		}

		// AM/PM, am/pm, A/P, a/p
		if (ch == 'A' || ch == 'a') && i+1 < len(s) {
			upper := strings.ToUpper(s[i:])
			if strings.HasPrefix(upper, "AM/PM") {
				tokens = append(tokens, fmtToken{kind: tkAmPm, raw: s[i : i+5]})
				i += 5
				continue
			}
			if strings.HasPrefix(upper, "A/P") {
				tokens = append(tokens, fmtToken{kind: tkAmPm, raw: s[i : i+3]})
				i += 3
				continue
			}
		}

		// General
		if (ch == 'G' || ch == 'g') && i+6 < len(s) {
			if strings.EqualFold(s[i:i+7], "General") {
				tokens = append(tokens, fmtToken{kind: tkGeneral, raw: s[i : i+7]})
				i += 7
				continue
			}
		}

		// 日付トークン（大文字小文字無視）
		upper := strings.ToUpper(s[i:])

		if strings.HasPrefix(upper, "YYYY") {
			tokens = append(tokens, fmtToken{kind: tkDateYear, raw: s[i : i+4], width: 4})
			i += 4
			continue
		}
		if strings.HasPrefix(upper, "YY") {
			tokens = append(tokens, fmtToken{kind: tkDateYear, raw: s[i : i+2], width: 2})
			i += 2
			continue
		}

		// gg, g (和暦) → リテラル空文字として読み飛ばし
		if ch == 'G' || ch == 'g' {
			n := 1
			for i+n < len(s) && (s[i+n] == 'g' || s[i+n] == 'G') {
				n++
			}
			i += n
			continue
		}

		// MMMMM, MMMM, MMM, MM, M (月 or 分 — 後で解決)
		if ch == 'M' || ch == 'm' {
			n := 1
			for i+n < len(s) && (s[i+n] == 'M' || s[i+n] == 'm') {
				n++
			}
			// 暫定的に tkDateMonth として登録（後で resolveMonthMinute で解決）
			tokens = append(tokens, fmtToken{kind: tkDateMonth, raw: s[i : i+n], width: n})
			i += n
			continue
		}

		// DDDD, DDD, DD, D
		if ch == 'D' || ch == 'd' {
			n := 1
			for i+n < len(s) && (s[i+n] == 'D' || s[i+n] == 'd') {
				n++
			}
			tokens = append(tokens, fmtToken{kind: tkDateDay, raw: s[i : i+n], width: n})
			i += n
			continue
		}

		// HH, H
		if ch == 'H' || ch == 'h' {
			n := 1
			for i+n < len(s) && (s[i+n] == 'H' || s[i+n] == 'h') {
				n++
			}
			tokens = append(tokens, fmtToken{kind: tkTimeHour, raw: s[i : i+n], width: n})
			i += n
			continue
		}

		// SS, S（秒の小数部 ss.0 ss.00 にも注意）
		if ch == 'S' || ch == 's' {
			n := 1
			for i+n < len(s) && (s[i+n] == 'S' || s[i+n] == 's') {
				n++
			}
			tokens = append(tokens, fmtToken{kind: tkTimeSecond, raw: s[i : i+n], width: n})
			i += n
			continue
		}

		// 数値トークン
		switch ch {
		case '0':
			tokens = append(tokens, fmtToken{kind: tkDigitZero, raw: "0"})
			i++
			continue
		case '#':
			tokens = append(tokens, fmtToken{kind: tkDigitHash, raw: "#"})
			i++
			continue
		case '?':
			tokens = append(tokens, fmtToken{kind: tkDigitSpace, raw: "?"})
			i++
			continue
		case '.':
			tokens = append(tokens, fmtToken{kind: tkDecimalPoint, raw: "."})
			i++
			continue
		case ',':
			tokens = append(tokens, fmtToken{kind: tkThousandSep, raw: ","})
			i++
			continue
		case '%':
			tokens = append(tokens, fmtToken{kind: tkPercent, raw: "%"})
			i++
			continue
		case '/':
			tokens = append(tokens, fmtToken{kind: tkSlash, raw: "/"})
			i++
			continue
		case '@':
			tokens = append(tokens, fmtToken{kind: tkAt, raw: "@"})
			i++
			continue
		}

		// E+, E-, e+, e-
		if (ch == 'E' || ch == 'e') && i+1 < len(s) && (s[i+1] == '+' || s[i+1] == '-') {
			tokens = append(tokens, fmtToken{kind: tkExponent, raw: s[i : i+2]})
			i += 2
			continue
		}

		// その他 → リテラル1文字（マルチバイト対応）
		_, size := utf8.DecodeRuneInString(s[i:])
		tokens = append(tokens, fmtToken{kind: tkLiteral, raw: s[i : i+size]})
		i += size
	}

	return tokens, cond
}

// resolveMonthMinute は mm/m トークンの文脈を解決する。
// h/hh/[h] の後の mm/m は分、ss/s/[s] の前の mm/m は分、それ以外は月。
func resolveMonthMinute(tokens []fmtToken) []fmtToken {
	for i := range tokens {
		if tokens[i].kind != tkDateMonth {
			continue
		}
		// 前方に h/hh/[h] があるか
		isMinute := false
		for j := i - 1; j >= 0; j-- {
			k := tokens[j].kind
			if k == tkTimeHour || k == tkElapsedHour {
				isMinute = true
				break
			}
			if k == tkLiteral || k == tkTimeSecond || k == tkElapsedSec {
				continue // リテラルは飛ばす
			}
			break // 他のトークンがあれば中断
		}
		// 後方に ss/s/[s] があるか
		if !isMinute {
			for j := i + 1; j < len(tokens); j++ {
				k := tokens[j].kind
				if k == tkTimeSecond || k == tkElapsedSec {
					isMinute = true
					break
				}
				if k == tkLiteral || k == tkDecimalPoint {
					continue
				}
				break
			}
		}
		if isMinute {
			tokens[i].kind = tkTimeMinute
		}
	}
	return tokens
}

// parseCondition は [>=100] 等の条件をパースする
func parseCondition(s string) *sectionCondition {
	s = strings.TrimSpace(s)
	for _, op := range []string{">=", "<=", "<>", ">", "<", "="} {
		if strings.HasPrefix(s, op) {
			val, err := strconv.ParseFloat(strings.TrimSpace(s[len(op):]), 64)
			if err != nil {
				return nil
			}
			return &sectionCondition{op: op, value: val}
		}
	}
	return nil
}

// extractCurrency は [$¥-ja-JP] から通貨記号 "¥" を抽出する
func extractCurrency(inner string) string {
	// inner は "$" で始まっている前提（"$" は除去済みの場合は "$¥-ja-JP"）
	rest := inner[1:] // "$" を除去
	idx := strings.IndexByte(rest, '-')
	if idx < 0 {
		return rest
	}
	return rest[:idx]
}

// ── セクション分類・事前計算 ─────────────────────────

func classifySection(sec *fmtSection) {
	hasDate := false
	hasNumeric := false
	hasAt := false
	hasGeneral := false

	for _, t := range sec.tokens {
		switch t.kind {
		case tkDateYear, tkDateMonth, tkDateDay, tkTimeHour, tkTimeMinute,
			tkTimeSecond, tkAmPm, tkElapsedHour, tkElapsedMin, tkElapsedSec:
			hasDate = true
		case tkDigitZero, tkDigitHash, tkDigitSpace, tkDecimalPoint,
			tkThousandSep, tkPercent, tkExponent, tkSlash:
			hasNumeric = true
		case tkAt:
			hasAt = true
		case tkGeneral:
			hasGeneral = true
		}
	}

	if hasDate {
		sec.kind = secDate
		// AM/PM フラグ
		for _, t := range sec.tokens {
			if t.kind == tkAmPm {
				sec.hasAmPm = true
				break
			}
		}
	} else if hasGeneral {
		sec.kind = secGeneral
	} else if hasAt && !hasNumeric {
		sec.kind = secText
	} else {
		sec.kind = secNumeric
	}
}

func precomputeNumeric(sec *fmtSection) {
	// % カウント
	for _, t := range sec.tokens {
		if t.kind == tkPercent {
			sec.percentCount++
		}
	}

	// 指数判定
	for i, t := range sec.tokens {
		if t.kind == tkExponent {
			sec.hasExponent = true
			if len(t.raw) >= 2 {
				sec.expSign = string(t.raw[1])
			}
			// 指数部の桁数（E+の後の0/#/?の個数）
			for j := i + 1; j < len(sec.tokens); j++ {
				k := sec.tokens[j].kind
				if k == tkDigitZero || k == tkDigitHash || k == tkDigitSpace {
					sec.expDigits++
				} else {
					break
				}
			}
			break
		}
	}

	// 分数判定
	for i, t := range sec.tokens {
		if t.kind == tkSlash {
			sec.hasFraction = true
			// 分母桁数の計算
			for j := i + 1; j < len(sec.tokens); j++ {
				k := sec.tokens[j].kind
				if k == tkDigitZero || k == tkDigitHash || k == tkDigitSpace {
					sec.fracDenomDigits++
				} else if k == tkLiteral && sec.tokens[j].raw == " " {
					continue
				} else {
					break
				}
			}
			// 固定分母の判定（/8, /16 等）
			sec.fracFixedDenom = detectFixedDenom(sec.tokens, i)
			// 整数部の有無
			for j := 0; j < i; j++ {
				if sec.tokens[j].kind == tkDigitHash || sec.tokens[j].kind == tkDigitZero {
					// スラッシュ直前の?/#ブロックは分子。その前にまだ#があれば整数部
					sec.fracWholePart = hasFractionWholePart(sec.tokens[:i])
					break
				}
			}
			break
		}
	}

	if sec.hasFraction || sec.hasExponent {
		return // 分数・指数では桁区切り・スケールは不要
	}

	// 小数点の位置を見つけ、整数部/小数部の桁数を計算
	decPos := -1
	for i, t := range sec.tokens {
		if t.kind == tkDecimalPoint {
			decPos = i
			break
		}
	}

	if decPos < 0 {
		// 小数点なし — 全トークンが整数部
		for _, t := range sec.tokens {
			if t.kind == tkDigitZero {
				sec.intDigits++
			}
		}
	} else {
		// 小数点前が整数部
		for _, t := range sec.tokens[:decPos] {
			if t.kind == tkDigitZero {
				sec.intDigits++
			}
		}
		// 小数点後が小数部
		for _, t := range sec.tokens[decPos+1:] {
			if t.kind == tkDigitZero || t.kind == tkDigitHash || t.kind == tkDigitSpace {
				sec.decDigits++
			}
		}
	}

	// 桁区切り・スケール（末尾カンマ）の判定
	// 数字トークンの範囲内のカンマは桁区切り
	// 数字トークンの後で次の数字トークンがないカンマはスケール
	lastDigitIdx := -1
	for i := len(sec.tokens) - 1; i >= 0; i-- {
		k := sec.tokens[i].kind
		if k == tkDigitZero || k == tkDigitHash || k == tkDigitSpace {
			lastDigitIdx = i
			break
		}
	}

	for i, t := range sec.tokens {
		if t.kind != tkThousandSep {
			continue
		}
		if i < lastDigitIdx {
			// 数字範囲内のカンマ → 桁区切り
			sec.useThousand = true
		} else {
			// 末尾カンマ → スケール
			sec.scaleDivisor++
		}
	}
}

// detectFixedDenom は分数の固定分母を検出する（"# ?/8" → 8）
func detectFixedDenom(tokens []fmtToken, slashIdx int) int {
	// スラッシュの後にリテラル数字のみが続く場合は固定分母
	numStr := ""
	for j := slashIdx + 1; j < len(tokens); j++ {
		t := tokens[j]
		if t.kind == tkDigitZero {
			numStr += "0"
		} else if t.kind == tkLiteral {
			// スペースは無視
			if strings.TrimSpace(t.raw) == "" {
				continue
			}
			// 数字リテラルかチェック
			if _, err := strconv.Atoi(t.raw); err == nil {
				numStr += t.raw
			} else {
				break
			}
		} else if t.kind == tkDigitHash || t.kind == tkDigitSpace {
			return 0 // 可変分母
		} else {
			break
		}
	}
	if numStr != "" && !strings.Contains(numStr, "0") {
		// 全部リテラル数字
	}
	// 実際には分母のトークンが全て0トークンで構成されているかで判定
	// 単純化: fracDenomDigitsが0なら固定分母の可能性
	return 0
}

// hasFractionWholePart は分数フォーマットに整数部があるか判定
func hasFractionWholePart(tokens []fmtToken) bool {
	// 分子部分（最後の連続?/#ブロック）の前に # があるか
	inNumerator := false
	for i := len(tokens) - 1; i >= 0; i-- {
		k := tokens[i].kind
		if k == tkDigitHash || k == tkDigitSpace || k == tkDigitZero {
			inNumerator = true
		} else if inNumerator {
			// 分子ブロックを抜けた → ここから前に#があれば整数部
			for j := i; j >= 0; j-- {
				if tokens[j].kind == tkDigitHash || tokens[j].kind == tkDigitZero {
					return true
				}
			}
			return false
		}
	}
	return false
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

// ── 日付フォーマッター ──────────────────────────────

func formatDate(sec *fmtSection, serial float64) string {
	t, err := excelDateToTime(math.Abs(serial))
	if err != nil {
		return ""
	}

	// 経過時間計算用
	totalHours := int(math.Abs(serial) * 24)
	totalMinutes := int(math.Abs(serial) * 24 * 60)
	totalSeconds := int(math.Abs(serial) * 24 * 60 * 60)

	// 秒の小数部（ss.0 用）
	fracSeconds := math.Abs(serial)*86400 - math.Floor(math.Abs(serial)*86400)

	var buf strings.Builder
	for i := 0; i < len(sec.tokens); i++ {
		tok := sec.tokens[i]
		switch tok.kind {
		case tkLiteral:
			buf.WriteString(tok.raw)
		case tkDateYear:
			if tok.width >= 4 {
				fmt.Fprintf(&buf, "%04d", t.Year())
			} else {
				fmt.Fprintf(&buf, "%02d", t.Year()%100)
			}
		case tkDateMonth:
			m := int(t.Month())
			switch tok.width {
			case 5: // mmmmm → 月名の最初の1文字
				buf.WriteByte(t.Month().String()[0])
			case 4: // mmmm → 月名
				buf.WriteString(t.Month().String())
			case 3: // mmm → 月名略称
				buf.WriteString(t.Month().String()[:3])
			case 2: // mm → ゼロ埋め
				fmt.Fprintf(&buf, "%02d", m)
			default: // m
				fmt.Fprintf(&buf, "%d", m)
			}
		case tkDateDay:
			switch tok.width {
			case 4: // dddd → 曜日
				buf.WriteString(t.Weekday().String())
			case 3: // ddd → 曜日略称
				buf.WriteString(t.Weekday().String()[:3])
			case 2: // dd → ゼロ埋め
				fmt.Fprintf(&buf, "%02d", t.Day())
			default: // d
				fmt.Fprintf(&buf, "%d", t.Day())
			}
		case tkTimeHour:
			h := t.Hour()
			if sec.hasAmPm {
				h = h % 12
				if h == 0 {
					h = 12
				}
			}
			if tok.width >= 2 {
				fmt.Fprintf(&buf, "%02d", h)
			} else {
				fmt.Fprintf(&buf, "%d", h)
			}
		case tkTimeMinute:
			if tok.width >= 2 {
				fmt.Fprintf(&buf, "%02d", t.Minute())
			} else {
				fmt.Fprintf(&buf, "%d", t.Minute())
			}
		case tkTimeSecond:
			if tok.width >= 2 {
				fmt.Fprintf(&buf, "%02d", t.Second())
			} else {
				fmt.Fprintf(&buf, "%d", t.Second())
			}
			// 秒の小数部: ss.0, ss.00 等
			if i+1 < len(sec.tokens) && sec.tokens[i+1].kind == tkDecimalPoint {
				// 小数点の後の0/#の個数をカウント
				decDigits := 0
				for j := i + 2; j < len(sec.tokens); j++ {
					if sec.tokens[j].kind == tkDigitZero || sec.tokens[j].kind == tkDigitHash {
						decDigits++
					} else {
						break
					}
				}
				if decDigits > 0 {
					frac := fracSeconds
					buf.WriteByte('.')
					for d := 0; d < decDigits; d++ {
						frac *= 10
						digit := int(frac) % 10
						buf.WriteByte(byte('0' + digit))
					}
					i += 1 + decDigits // 小数点と桁をスキップ
				}
			}
		case tkAmPm:
			upper := strings.ToUpper(tok.raw)
			if t.Hour() < 12 {
				if strings.HasPrefix(upper, "AM") {
					buf.WriteString("AM")
				} else {
					buf.WriteString("A")
				}
			} else {
				if strings.HasPrefix(upper, "AM") {
					buf.WriteString("PM")
				} else {
					buf.WriteString("P")
				}
			}
		case tkElapsedHour:
			fmt.Fprintf(&buf, "%d", totalHours)
		case tkElapsedMin:
			fmt.Fprintf(&buf, "%d", totalMinutes)
		case tkElapsedSec:
			fmt.Fprintf(&buf, "%d", totalSeconds)
		case tkDecimalPoint:
			buf.WriteByte('.')
		case tkSlash:
			buf.WriteByte('/')
		case tkThousandSep:
			buf.WriteByte(',')
		case tkDigitZero, tkDigitHash, tkDigitSpace:
			// 日付フォーマットでは数字トークンはリテラルとして扱う
		}
	}

	return buf.String()
}

// ── 数値フォーマッター ──────────────────────────────

func formatNumber(sec *fmtSection, value float64) string {
	// パーセント処理
	for i := 0; i < sec.percentCount; i++ {
		value *= 100
	}

	// スケール処理（末尾カンマ）
	for i := 0; i < sec.scaleDivisor; i++ {
		value /= 1000
	}

	// 指数フォーマット
	if sec.hasExponent {
		return formatExponent(sec, value)
	}

	// 分数フォーマット
	if sec.hasFraction {
		return formatFraction(sec, value)
	}

	// 通常数値
	return formatStandardNumber(sec, value)
}

// formatStandardNumber は通常の数値フォーマットを行う。
// 整数部・小数部それぞれの数字トークンに1桁ずつ右詰めで割り当てる。
func formatStandardNumber(sec *fmtSection, value float64) string {
	absVal := math.Abs(value)

	// 小数桁数で丸め
	if sec.decDigits > 0 {
		factor := math.Pow(10, float64(sec.decDigits))
		absVal = math.Round(absVal*factor) / factor
	} else {
		absVal = math.Round(absVal)
	}

	// 整数部と小数部を文字列化
	intPart := int64(absVal)
	intDigits := strconv.FormatInt(intPart, 10) // 純粋な数字列

	fracDigits := ""
	if sec.decDigits > 0 {
		fracPart := absVal - float64(intPart)
		f := strconv.FormatFloat(fracPart, 'f', sec.decDigits, 64)
		if len(f) > 2 {
			fracDigits = f[2:]
		} else {
			fracDigits = strings.Repeat("0", sec.decDigits)
		}
	}

	// 小数点の位置を見つける
	decPos := -1
	for i, t := range sec.tokens {
		if t.kind == tkDecimalPoint {
			decPos = i
			break
		}
	}

	intEnd := len(sec.tokens)
	if decPos >= 0 {
		intEnd = decPos
	}

	// ── 整数部の描画 ──
	// 整数部の数字トークン位置を収集
	var intTokenIndices []int
	for i := 0; i < intEnd; i++ {
		k := sec.tokens[i].kind
		if k == tkDigitZero || k == tkDigitHash || k == tkDigitSpace {
			intTokenIndices = append(intTokenIndices, i)
		}
	}

	nSlots := len(intTokenIndices)
	nDigits := len(intDigits)

	// 各スロットに割り当てる文字を決定（右詰め）
	slotChars := make([]string, nSlots)
	for i := 0; i < nSlots; i++ {
		digitPos := nDigits - nSlots + i // intDigits 内の位置
		if digitPos >= 0 {
			slotChars[i] = string(intDigits[digitPos])
		} else {
			// 値の桁数が足りない → トークン種別でパディング
			switch sec.tokens[intTokenIndices[i]].kind {
			case tkDigitZero:
				slotChars[i] = "0"
			case tkDigitSpace:
				slotChars[i] = " "
			case tkDigitHash:
				slotChars[i] = "" // 先頭の不要なゼロ省略
			}
		}
	}

	// オーバーフロー（値の桁数 > スロット数）→ 先頭スロットに余剰桁を付加
	if nDigits > nSlots && nSlots > 0 {
		overflow := intDigits[:nDigits-nSlots]
		slotChars[0] = overflow + slotChars[0]
	}

	// 桁区切り挿入: slotChars を結合した数字列に適用してから再分配
	if sec.useThousand && nSlots > 0 {
		// 全スロットの数字を結合
		combined := strings.Join(slotChars, "")
		if combined != "" {
			combined = addThousandSep(combined)
		}
		// 結合結果を最初のスロットにまとめ、残りは空に
		slotChars[0] = combined
		for i := 1; i < nSlots; i++ {
			slotChars[i] = ""
		}
	}

	// トークン列に沿って整数部を出力
	var buf strings.Builder
	slotIdx := 0
	for i := 0; i < intEnd; i++ {
		t := sec.tokens[i]
		switch t.kind {
		case tkDigitZero, tkDigitHash, tkDigitSpace:
			buf.WriteString(slotChars[slotIdx])
			slotIdx++
		case tkThousandSep:
			// 桁区切りは slotChars に含まれているのでスキップ
			// 桁区切りでない場合（useThousand=false）もスキップ（スケール用カンマ）
		case tkLiteral:
			buf.WriteString(t.raw)
		case tkPercent:
			buf.WriteString("%")
		}
	}

	// ── 小数部の描画 ──
	if decPos >= 0 {
		buf.WriteByte('.')
		fracIdx := 0
		for i := decPos + 1; i < len(sec.tokens); i++ {
			t := sec.tokens[i]
			switch t.kind {
			case tkDigitZero:
				if fracIdx < len(fracDigits) {
					buf.WriteByte(fracDigits[fracIdx])
				} else {
					buf.WriteByte('0')
				}
				fracIdx++
			case tkDigitHash:
				if fracIdx < len(fracDigits) {
					// # は末尾ゼロを省略: 以降が全てゼロなら出力しない
					remaining := fracDigits[fracIdx:]
					if strings.TrimRight(remaining, "0") != "" {
						buf.WriteByte(fracDigits[fracIdx])
					}
				}
				fracIdx++
			case tkDigitSpace:
				if fracIdx < len(fracDigits) {
					buf.WriteByte(fracDigits[fracIdx])
				} else {
					buf.WriteByte(' ')
				}
				fracIdx++
			case tkLiteral:
				buf.WriteString(t.raw)
			case tkPercent:
				buf.WriteString("%")
			}
		}
	}

	return buf.String()
}

func addThousandSep(s string) string {
	if len(s) <= 3 {
		return s
	}
	var buf strings.Builder
	start := len(s) % 3
	if start > 0 {
		buf.WriteString(s[:start])
	}
	for i := start; i < len(s); i += 3 {
		if buf.Len() > 0 {
			buf.WriteByte(',')
		}
		buf.WriteString(s[i : i+3])
	}
	return buf.String()
}

// ── 指数フォーマッター ──────────────────────────────

func formatExponent(sec *fmtSection, value float64) string {
	if value == 0 {
		return "0" + sec.tokens[0].raw[0:0] // 簡易実装
	}

	absVal := math.Abs(value)
	exp := 0
	if absVal != 0 {
		exp = int(math.Floor(math.Log10(absVal)))
	}

	// 仮数の整数部桁数に合わせて指数を調整
	// 整数部のトークン数を数える
	intDigitCount := 0
	for _, t := range sec.tokens {
		if t.kind == tkExponent {
			break
		}
		if t.kind == tkDigitZero || t.kind == tkDigitHash || t.kind == tkDigitSpace {
			intDigitCount++
		}
		if t.kind == tkDecimalPoint {
			break
		}
	}
	if intDigitCount > 1 {
		exp -= intDigitCount - 1
	}

	mantissa := absVal / math.Pow(10, float64(exp))

	// 仮数部分のフォーマット（指数前のトークンを使う）
	var mantTokens []fmtToken
	expIdx := -1
	for i, t := range sec.tokens {
		if t.kind == tkExponent {
			expIdx = i
			break
		}
		mantTokens = append(mantTokens, t)
	}

	// 仮数用のセクションを作成
	mantSec := fmtSection{kind: secNumeric, tokens: mantTokens}
	precomputeNumeric(&mantSec)
	mantStr := formatStandardNumber(&mantSec, mantissa)

	// 指数部のフォーマット
	var buf strings.Builder
	buf.WriteString(mantStr)
	buf.WriteByte('E')
	if sec.expSign == "+" || exp >= 0 {
		if exp >= 0 {
			buf.WriteByte('+')
		}
	}
	if exp < 0 {
		buf.WriteByte('-')
		exp = -exp
	}
	expStr := strconv.Itoa(exp)
	for len(expStr) < sec.expDigits {
		expStr = "0" + expStr
	}
	buf.WriteString(expStr)

	// 指数後のリテラルトークン
	if expIdx >= 0 {
		for j := expIdx + 1 + sec.expDigits; j < len(sec.tokens); j++ {
			if sec.tokens[j].kind == tkLiteral {
				buf.WriteString(sec.tokens[j].raw)
			}
		}
	}

	return buf.String()
}

// ── 分数フォーマッター ──────────────────────────────

func formatFraction(sec *fmtSection, value float64) string {
	absVal := math.Abs(value)
	wholePart := int(absVal)
	frac := absVal - float64(wholePart)

	var num, den int

	if sec.fracFixedDenom > 0 {
		den = sec.fracFixedDenom
		num = int(math.Round(frac * float64(den)))
	} else {
		maxDen := 1
		for i := 0; i < sec.fracDenomDigits; i++ {
			maxDen *= 10
		}
		maxDen-- // ?/? → max 9, ??/?? → max 99
		if maxDen < 1 {
			maxDen = 9
		}
		num, den = bestFraction(frac, maxDen)
	}

	// 分子が0なら整数のみ
	if num == 0 {
		if sec.fracWholePart {
			return fmt.Sprintf("%d", wholePart)
		}
		return "0"
	}

	// 繰り上がり処理
	if num >= den {
		wholePart += num / den
		num = num % den
	}

	// フォーマット出力
	var buf strings.Builder
	if sec.fracWholePart {
		if wholePart > 0 {
			buf.WriteString(strconv.Itoa(wholePart))
			buf.WriteByte(' ')
		}
	} else {
		num += wholePart * den
	}

	// 分子と分母の桁数調整
	numStr := strconv.Itoa(num)
	denStr := strconv.Itoa(den)

	// 分子のパディング
	for len(numStr) < sec.fracDenomDigits {
		numStr = " " + numStr
	}
	// 分母のパディング
	for len(denStr) < sec.fracDenomDigits {
		denStr = denStr + " "
	}

	buf.WriteString(numStr)
	buf.WriteByte('/')
	buf.WriteString(denStr)

	return buf.String()
}

// bestFraction は値に最も近い分数を求める（Stern-Brocot木）
func bestFraction(value float64, maxDen int) (num, den int) {
	if value <= 0 {
		return 0, 1
	}
	if value >= 1 {
		return 1, 1
	}

	bestNum, bestDen := 0, 1
	bestErr := value

	// メディアント法で近似
	ln, ld := 0, 1 // 左境界: 0/1
	rn, rd := 1, 1 // 右境界: 1/1

	for {
		mn := ln + rn
		md := ld + rd
		if md > maxDen {
			break
		}
		mediant := float64(mn) / float64(md)
		err := math.Abs(value - mediant)
		if err < bestErr {
			bestErr = err
			bestNum = mn
			bestDen = md
		}
		if math.Abs(err) < 1e-12 {
			break
		}
		if mediant < value {
			ln, ld = mn, md
		} else {
			rn, rd = mn, md
		}
	}

	return bestNum, bestDen
}

// ── テキストフォーマッター ───────────────────────────

func formatText(sec *fmtSection, text string) string {
	var buf strings.Builder
	for _, t := range sec.tokens {
		if t.kind == tkAt {
			buf.WriteString(text)
		} else if t.kind == tkLiteral {
			buf.WriteString(t.raw)
		}
	}
	return buf.String()
}
