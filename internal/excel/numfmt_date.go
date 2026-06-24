package excel

import (
	"fmt"
	"math"
	"strings"
)

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
