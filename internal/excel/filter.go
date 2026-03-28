package excel

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// numericEpsilon は数値比較で等価とみなす閾値
const numericEpsilon = 1e-9

// Filter は search コマンドのセル検索条件。
// Query, Numeric, Type の各条件は AND で結合される。
type Filter struct {
	Text    string       // テキスト部分一致（大文字小文字無視）
	Numeric *NumericExpr // 数値比較条件
	Type    CellType     // セル型フィルタ（空なら全型）
}

// NumericExpr は数値比較式。">100", "<=50", "100:200" 等を表す。
type NumericExpr struct {
	Op    string  // ">", ">=", "<", "<=", "=", ":"（範囲）
	Value float64 // 単一比較の値、または範囲の下限
	Upper float64 // 範囲指定時の上限
}

// ParseNumericExpr は --numeric の値をパースする
func ParseNumericExpr(s string) (*NumericExpr, error) {
	s = strings.TrimSpace(s)

	// 範囲指定: "100:200"
	if parts := strings.SplitN(s, ":", 2); len(parts) == 2 {
		lower, err := strconv.ParseFloat(parts[0], 64)
		if err != nil {
			return nil, fmt.Errorf("数値の解析に失敗しました: %q", parts[0])
		}
		upper, err := strconv.ParseFloat(parts[1], 64)
		if err != nil {
			return nil, fmt.Errorf("数値の解析に失敗しました: %q", parts[1])
		}
		if upper < lower {
			return nil, fmt.Errorf("範囲の上限が下限より小さいです: %s", s)
		}
		return &NumericExpr{Op: ":", Value: lower, Upper: upper}, nil
	}

	// 比較演算子
	for _, op := range []string{">=", "<=", ">", "<", "="} {
		if strings.HasPrefix(s, op) {
			val, err := strconv.ParseFloat(strings.TrimSpace(s[len(op):]), 64)
			if err != nil {
				return nil, fmt.Errorf("数値の解析に失敗しました: %q", s[len(op):])
			}
			return &NumericExpr{Op: op, Value: val}, nil
		}
	}

	return nil, fmt.Errorf("数値比較式の形式が不正です: %q", s)
}

// Match は数値がこの式にマッチするか判定する
func (e *NumericExpr) Match(v float64) bool {
	switch e.Op {
	case ">":
		return v > e.Value
	case ">=":
		return v >= e.Value
	case "<":
		return v < e.Value
	case "<=":
		return v <= e.Value
	case "=":
		return math.Abs(v-e.Value) <= numericEpsilon
	case ":":
		return v >= e.Value && v <= e.Upper
	}
	return false
}

// MatchCell はセルデータがフィルタ条件にマッチするかを判定する（AND結合）
func (f *Filter) MatchCell(data *CellData) bool {
	// --type フィルタ
	// formula セルはキャッシュ値の型も考慮する（例: --type number で数値キャッシュの数式セルもヒット）
	if f.Type != "" {
		if data.Type != f.Type && !matchesCachedType(data, f.Type) {
			return false
		}
	}

	// --text フィルタ
	if f.Text != "" {
		query := strings.ToLower(f.Text)
		matched := false
		if data.Display != "" && strings.Contains(strings.ToLower(data.Display), query) {
			matched = true
		}
		if !matched {
			valStr := valueToSearchString(data.Value)
			if strings.Contains(strings.ToLower(valStr), query) {
				matched = true
			}
		}
		if !matched {
			return false
		}
	}

	// --numeric フィルタ（number 型 + 数式セルの数値キャッシュも対象）
	if f.Numeric != nil {
		num, ok := extractNumericValue(data)
		if !ok {
			return false
		}
		if !f.Numeric.Match(num) {
			return false
		}
	}

	return true
}

// matchesCachedType は formula セルのキャッシュ値の型が targetType に一致するか判定する
func matchesCachedType(data *CellData, targetType CellType) bool {
	if data.Type != CellTypeFormula {
		return false
	}
	switch targetType {
	case CellTypeNumber:
		_, ok := data.Value.(float64)
		return ok
	case CellTypeString:
		_, ok := data.Value.(string)
		return ok && !isErrorValue(data.Value.(string))
	case CellTypeBool:
		_, ok := data.Value.(bool)
		return ok
	}
	return false
}

// extractNumericValue はセルから数値を抽出する（number型 + formulaの数値キャッシュ）
func extractNumericValue(data *CellData) (float64, bool) {
	if data.Type == CellTypeNumber || data.Type == CellTypeFormula {
		num, ok := data.Value.(float64)
		return num, ok
	}
	return 0, false
}

func valueToSearchString(v any) string {
	if v == nil {
		return ""
	}
	switch val := v.(type) {
	case string:
		return val
	case float64:
		return strconv.FormatFloat(val, 'f', -1, 64)
	case bool:
		if val {
			return "TRUE"
		}
		return "FALSE"
	default:
		return fmt.Sprintf("%v", v)
	}
}
