package excel

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

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
	for i := range nSlots {
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
