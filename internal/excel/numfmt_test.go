package excel

import (
	"testing"
)

func TestFormatNumericValue(t *testing.T) {
	tests := []struct {
		name   string
		fmt    string
		value  float64
		expect string
	}{
		// ── 空・General ──
		{"empty format", "", 123, ""},
		{"General", "General", 123, ""},
		{"text format @", "@", 123, ""},

		// ── 基本数値 ──
		{"integer zero-padded", "0", 5, "5"},
		{"integer zero-padded 3 digits", "000", 5, "005"},
		{"decimal 0.00", "0.00", 3.5, "3.50"},
		{"decimal 0.00 rounding", "0.00", 3.456, "3.46"},
		{"decimal 0.0", "0.0", 0.1, "0.1"},
		{"hash integer", "#", 0, "0"},
		{"hash integer nonzero", "#", 5, "5"},
		{"hash decimal", "#.##", 3.5, "3.5"},

		// ── 桁区切り ──
		{"thousand sep", "#,##0", 1234567, "1,234,567"},
		{"thousand sep small", "#,##0", 999, "999"},
		{"thousand sep zero", "#,##0", 0, "0"},
		{"thousand sep decimal", "#,##0.00", 1234.5, "1,234.50"},

		// ── パーセント ──
		{"percent integer", "0%", 0.15, "15%"},
		{"percent decimal", "0.00%", 0.1555, "15.55%"},
		{"percent whole", "0%", 1, "100%"},

		// ── スケール（末尾カンマ） ──
		{"scale thousands", "#,##0,", 1234567, "1,235"},
		{"scale millions", "#,##0,,", 1234567890, "1,235"},

		// ── セクション分岐（正;負;ゼロ） ──
		{"section positive", "#,##0;(#,##0)", 1234, "1,234"},
		{"section negative", "#,##0;(#,##0)", -1234, "(1,234)"},
		{"section 3 zero", "#,##0;(#,##0);\"-\"", 0, "-"},
		{"section 2 neg sign removed", "0;0", -42, "42"},

		// ── 条件付きセクション ──
		{"condition >=100", "[>=100]0;0.00", 150, "150"},
		{"condition <100", "[>=100]0;0.00", 50, "50.00"},

		// ── 日付フォーマット ──
		{"date yyyy/m/d", "yyyy/m/d", 45735, "2025/3/19"},
		{"date yyyy-mm-dd", "yyyy-mm-dd", 45735, "2025-03-19"},
		{"date yyyy年m月d日", "yyyy\"年\"m\"月\"d\"日\"", 45735, "2025年3月19日"},
		{"date m/d/yyyy", "m/d/yyyy", 45735, "3/19/2025"},
		{"date dd-mmm-yy", "dd-mmm-yy", 45735, "19-Mar-25"},
		{"date mmm-yy", "mmm-yy", 45735, "Mar-25"},
		{"date d-mmm", "d-mmm", 45735, "19-Mar"},
		{"date yy", "yy/mm/dd", 45735, "25/03/19"},

		// ── 時刻フォーマット ──
		{"time h:mm", "h:mm", 0.4375, "10:30"},
		{"time hh:mm:ss", "hh:mm:ss", 0.4375, "10:30:00"},
		{"time h:mm AM/PM", "h:mm AM/PM", 0.4375, "10:30 AM"},
		{"time h:mm AM/PM pm", "h:mm AM/PM", 0.75, "6:00 PM"},
		{"time h:mm A/P", "h:mm A/P", 0.75, "6:00 P"},

		// ── 日時 ──
		{"datetime", "yyyy/m/d h:mm", 45735.4375, "2025/3/19 10:30"},

		// ── mm の文脈判定（月 vs 分） ──
		{"mm as month", "yyyy/mm/dd", 45735, "2025/03/19"},
		{"mm as minute after h", "h:mm", 0.5, "12:00"},
		{"mm as minute before ss", "mm:ss", 0.000694, "00:59"}, // 0.000694日 ≈ 59.9秒

		// ── 経過時間 ──
		{"elapsed hours", "[h]:mm:ss", 1.5, "36:00:00"},
		{"elapsed hours short", "[h]:mm", 0.5, "12:00"},

		// ── 曜日 ──
		{"weekday ddd", "ddd", 45735, "Wed"},
		{"weekday dddd", "dddd", 45735, "Wednesday"},

		// ── 月名 ──
		{"month mmm", "mmm", 45735, "Mar"},
		{"month mmmm", "mmmm", 45735, "March"},
		{"month mmmmm", "mmmmm", 45735, "M"},

		// ── リテラル・エスケープ ──
		{"literal quoted", "0\" items\"", 42, "42 items"},
		{"literal backslash", "0\\-0", 12, "1-2"},
		{"underscore skip", "0_)", 5, "5 "},

		// ── 色指定（無視） ──
		{"color ignored", "[Red]#,##0", 1234, "1,234"},

		// ── ロケール ──
		{"currency yen", "[$¥-ja-JP]#,##0", 1234, "¥1,234"},

		// ── 指数 ──
		{"exponent E+00", "0.00E+00", 12345, "1.23E+04"},
		{"exponent small", "0.00E+00", 0.00123, "1.23E-03"},
		{"exponent zero", "0.00E+00", 0, "0.00E+00"},

		// ── 組み込みフォーマットID相当 ──
		{"builtin 1: 0", "0", 42, "42"},
		{"builtin 2: 0.00", "0.00", 42, "42.00"},
		{"builtin 3: #,##0", "#,##0", 42000, "42,000"},
		{"builtin 4: #,##0.00", "#,##0.00", 42000, "42,000.00"},
		{"builtin 9: 0%", "0%", 0.42, "42%"},
		{"builtin 10: 0.00%", "0.00%", 0.42, "42.00%"},
		{"builtin 14: m/d/yyyy", "m/d/yyyy", 45735, "3/19/2025"},
		{"builtin 20: h:mm", "h:mm", 0.75, "18:00"},
		{"builtin 21: h:mm:ss", "h:mm:ss", 0.75, "18:00:00"},
		{"builtin 22: m/d/yyyy h:mm", "m/d/yyyy h:mm", 45735.75, "3/19/2025 18:00"},

		// ── エッジケース ──
		{"zero with 0.00", "0.00", 0, "0.00"},
		{"negative with section", "#,##0;-#,##0", -500, "-500"},
		{"date serial 1 (1900-01-01)", "yyyy-mm-dd", 1, "1900-01-01"},
		{"date serial 60 (leap year bug)", "yyyy-mm-dd", 60, "1900-03-01"},
		{"date serial 61", "yyyy-mm-dd", 61, "1900-03-01"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := FormatNumericValue(tt.fmt, tt.value)
			if got != tt.expect {
				t.Errorf("FormatNumericValue(%q, %v)\n  got:    %q\n  expect: %q", tt.fmt, tt.value, got, tt.expect)
			}
		})
	}
}
