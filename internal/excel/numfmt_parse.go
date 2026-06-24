package excel

import (
	"strconv"
	"strings"
	"unicode/utf8"
)

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
	before, _, _ := strings.Cut(rest, "-")
	return before
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
			for j := range i {
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
		if v, err := strconv.Atoi(numStr); err == nil && v > 0 {
			return v
		}
	}
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
