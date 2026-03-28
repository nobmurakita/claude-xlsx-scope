package excel

import (
	"testing"
)

func TestParseRange(t *testing.T) {
	tests := []struct {
		name      string
		rangeStr  string
		usedRange string
		expect    CellRange
		wantErr   bool
	}{
		// セル範囲 (A1:H20)
		{"cell range", "A1:H20", "", CellRange{1, 1, 8, 20}, false},
		{"cell range lowercase", "a1:h20", "", CellRange{1, 1, 8, 20}, false},
		{"cell range with spaces", " B2:D10 ", "", CellRange{2, 2, 4, 10}, false},
		{"single cell range", "C5:C5", "", CellRange{3, 5, 3, 5}, false},
		{"wide range", "A1:AZ100", "", CellRange{1, 1, 52, 100}, false},

		// 列範囲 (A:F)
		{"col range", "A:F", "A1:Z50", CellRange{1, 1, 6, 50}, false},
		{"col range single", "C:C", "A1:H20", CellRange{3, 1, 3, 20}, false},
		{"col range no used", "A:F", "", CellRange{}, false},

		// 行範囲 (1:20)
		{"row range", "1:20", "A1:H50", CellRange{1, 1, 8, 20}, false},
		{"row range single", "5:5", "A1:H50", CellRange{1, 5, 8, 5}, false},
		{"row range no used", "1:20", "", CellRange{}, false},

		// 単一セル
		{"single cell", "B3", "", CellRange{2, 3, 2, 3}, false},
		{"single cell A1", "A1", "", CellRange{1, 1, 1, 1}, false},

		// エラーケース
		{"invalid format", "foo", "", CellRange{}, true},
		{"reversed range", "H20:A1", "", CellRange{}, true},
		{"reversed col range", "F:A", "", CellRange{}, true},
		{"reversed row range", "20:1", "", CellRange{}, true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := ParseRange(tt.rangeStr, tt.usedRange)
			if tt.wantErr {
				if err == nil {
					t.Errorf("ParseRange(%q, %q) expected error, got %v", tt.rangeStr, tt.usedRange, got)
				}
				return
			}
			if err != nil {
				t.Errorf("ParseRange(%q, %q) unexpected error: %v", tt.rangeStr, tt.usedRange, err)
				return
			}
			if got != tt.expect {
				t.Errorf("ParseRange(%q, %q)\n  got:    %+v\n  expect: %+v", tt.rangeStr, tt.usedRange, got, tt.expect)
			}
		})
	}
}

func TestStartPosition(t *testing.T) {
	tests := []struct {
		name    string
		cell    string
		col     int
		row     int
		wantErr bool
	}{
		{"A1", "A1", 1, 1, false},
		{"B50", "B50", 2, 50, false},
		{"AA1", "AA1", 27, 1, false},
		{"lowercase", "c3", 3, 3, false},
		{"invalid", "123", 0, 0, true},
		{"empty", "", 0, 0, true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			col, row, err := StartPosition(tt.cell)
			if tt.wantErr {
				if err == nil {
					t.Errorf("StartPosition(%q) expected error", tt.cell)
				}
				return
			}
			if err != nil {
				t.Errorf("StartPosition(%q) unexpected error: %v", tt.cell, err)
				return
			}
			if col != tt.col || row != tt.row {
				t.Errorf("StartPosition(%q) = (%d, %d), want (%d, %d)", tt.cell, col, row, tt.col, tt.row)
			}
		})
	}
}

func TestColNameRoundTrip(t *testing.T) {
	tests := []struct {
		num  int
		name string
	}{
		{1, "A"},
		{26, "Z"},
		{27, "AA"},
		{52, "AZ"},
		{53, "BA"},
		{702, "ZZ"},
		{703, "AAA"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := ColName(tt.num)
			if got != tt.name {
				t.Errorf("ColName(%d) = %q, want %q", tt.num, got, tt.name)
			}
			// ラウンドトリップ確認
			gotNum := colNumber(tt.name)
			if gotNum != tt.num {
				t.Errorf("colNumber(%q) = %d, want %d", tt.name, gotNum, tt.num)
			}
		})
	}
}
