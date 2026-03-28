package excel

import (
	"testing"
)

func TestParseNumericExpr(t *testing.T) {
	tests := []struct {
		name    string
		input   string
		op      string
		value   float64
		upper   float64
		wantErr bool
	}{
		{"greater than", ">100", ">", 100, 0, false},
		{"greater equal", ">=50", ">=", 50, 0, false},
		{"less than", "<10", "<", 10, 0, false},
		{"less equal", "<=0", "<=", 0, 0, false},
		{"equal", "=42", "=", 42, 0, false},
		{"range", "100:200", ":", 100, 200, false},
		{"range with decimals", "1.5:3.5", ":", 1.5, 3.5, false},
		{"negative value", ">-10", ">", -10, 0, false},
		{"with spaces", " >=50 ", ">=", 50, 0, false},

		// エラーケース
		{"invalid format", "abc", "", 0, 0, true},
		{"reversed range", "200:100", "", 0, 0, true},
		{"bad number", ">abc", "", 0, 0, true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			expr, err := ParseNumericExpr(tt.input)
			if tt.wantErr {
				if err == nil {
					t.Errorf("ParseNumericExpr(%q) expected error", tt.input)
				}
				return
			}
			if err != nil {
				t.Errorf("ParseNumericExpr(%q) unexpected error: %v", tt.input, err)
				return
			}
			if expr.Op != tt.op || expr.Value != tt.value || expr.Upper != tt.upper {
				t.Errorf("ParseNumericExpr(%q) = {%s, %v, %v}, want {%s, %v, %v}",
					tt.input, expr.Op, expr.Value, expr.Upper, tt.op, tt.value, tt.upper)
			}
		})
	}
}

func TestNumericExprMatch(t *testing.T) {
	tests := []struct {
		name  string
		op    string
		value float64
		upper float64
		input float64
		want  bool
	}{
		{"gt match", ">", 100, 0, 150, true},
		{"gt no match", ">", 100, 0, 100, false},
		{"gte match", ">=", 100, 0, 100, true},
		{"lt match", "<", 10, 0, 5, true},
		{"lt no match", "<", 10, 0, 10, false},
		{"lte match", "<=", 10, 0, 10, true},
		{"eq match", "=", 42, 0, 42, true},
		{"eq near match", "=", 42, 0, 42.0000000001, true},
		{"eq no match", "=", 42, 0, 43, false},
		{"range match lower", ":", 100, 200, 100, true},
		{"range match upper", ":", 100, 200, 200, true},
		{"range match mid", ":", 100, 200, 150, true},
		{"range no match", ":", 100, 200, 250, false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			expr := &NumericExpr{Op: tt.op, Value: tt.value, Upper: tt.upper}
			got := expr.Match(tt.input)
			if got != tt.want {
				t.Errorf("NumericExpr{%s, %v, %v}.Match(%v) = %v, want %v",
					tt.op, tt.value, tt.upper, tt.input, got, tt.want)
			}
		})
	}
}

func TestMatchCell(t *testing.T) {
	tests := []struct {
		name   string
		filter Filter
		data   CellData
		want   bool
	}{
		// --type フィルタ
		{
			"type string match",
			Filter{Type: CellTypeString},
			CellData{Type: CellTypeString, Value: "hello"},
			true,
		},
		{
			"type string no match",
			Filter{Type: CellTypeString},
			CellData{Type: CellTypeNumber, Value: 42.0},
			false,
		},
		{
			"type number match",
			Filter{Type: CellTypeNumber},
			CellData{Type: CellTypeNumber, Value: 42.0},
			true,
		},
		{
			"type formula with number cache matches number",
			Filter{Type: CellTypeNumber},
			CellData{Type: CellTypeFormula, Value: 42.0},
			true,
		},
		{
			"type formula with string cache matches string",
			Filter{Type: CellTypeString},
			CellData{Type: CellTypeFormula, Value: "hello"},
			true,
		},
		{
			"type formula with bool cache matches bool",
			Filter{Type: CellTypeBool},
			CellData{Type: CellTypeFormula, Value: true},
			true,
		},
		{
			"type formula with error cache does not match string",
			Filter{Type: CellTypeString},
			CellData{Type: CellTypeFormula, Value: "#N/A"},
			false,
		},

		// --query フィルタ
		{
			"query display match",
			Filter{Query: "合計"},
			CellData{Type: CellTypeString, Value: "小合計", Display: "小合計"},
			true,
		},
		{
			"query value match",
			Filter{Query: "100"},
			CellData{Type: CellTypeNumber, Value: 100.0},
			true,
		},
		{
			"query case insensitive",
			Filter{Query: "hello"},
			CellData{Type: CellTypeString, Value: "Hello World", Display: "Hello World"},
			true,
		},
		{
			"query no match",
			Filter{Query: "xyz"},
			CellData{Type: CellTypeString, Value: "hello", Display: "hello"},
			false,
		},
		{
			"query bool TRUE",
			Filter{Query: "TRUE"},
			CellData{Type: CellTypeBool, Value: true},
			true,
		},

		// --numeric フィルタ
		{
			"numeric match",
			Filter{Numeric: &NumericExpr{Op: ">", Value: 50}},
			CellData{Type: CellTypeNumber, Value: 100.0},
			true,
		},
		{
			"numeric no match",
			Filter{Numeric: &NumericExpr{Op: ">", Value: 50}},
			CellData{Type: CellTypeNumber, Value: 30.0},
			false,
		},
		{
			"numeric on string is rejected",
			Filter{Numeric: &NumericExpr{Op: ">", Value: 50}},
			CellData{Type: CellTypeString, Value: "100"},
			false,
		},
		{
			"numeric formula cache match",
			Filter{Numeric: &NumericExpr{Op: "=", Value: 42}},
			CellData{Type: CellTypeFormula, Value: 42.0},
			true,
		},

		// AND 結合
		{
			"AND query + type match",
			Filter{Query: "100", Type: CellTypeNumber},
			CellData{Type: CellTypeNumber, Value: 100.0},
			true,
		},
		{
			"AND query match but type mismatch",
			Filter{Query: "100", Type: CellTypeString},
			CellData{Type: CellTypeNumber, Value: 100.0},
			false,
		},
		{
			"AND query + numeric both match",
			Filter{Query: "50", Numeric: &NumericExpr{Op: ">=", Value: 50}},
			CellData{Type: CellTypeNumber, Value: 150.0, Display: "150"},
			true, // query "50" は "150" に部分一致する
		},
		{
			"AND query match but numeric mismatch",
			Filter{Query: "30", Numeric: &NumericExpr{Op: ">", Value: 100}},
			CellData{Type: CellTypeNumber, Value: 30.0},
			false,
		},
		{
			"AND all three match",
			Filter{Query: "500", Type: CellTypeNumber, Numeric: &NumericExpr{Op: ">=", Value: 100}},
			CellData{Type: CellTypeNumber, Value: 500.0},
			true,
		},

		// 空フィルタ
		{
			"empty filter matches all",
			Filter{},
			CellData{Type: CellTypeString, Value: "anything"},
			true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := tt.filter.MatchCell(&tt.data)
			if got != tt.want {
				t.Errorf("MatchCell() = %v, want %v\n  filter: %+v\n  data: %+v", got, tt.want, tt.filter, tt.data)
			}
		})
	}
}
