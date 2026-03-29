package excel

import (
	"testing"
	"time"
)

func TestExcelDateToTime(t *testing.T) {
	tests := []struct {
		name   string
		serial float64
		want   time.Time
	}{
		// 基本日付
		{"1900-01-01", 1, time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC)},
		{"1900-01-02", 2, time.Date(1900, 1, 2, 0, 0, 0, 0, time.UTC)},
		// うるう年バグ境界（シリアル値 60 = Excel の 1900-02-29）
		{"leap year bug boundary 59", 59, time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC)},
		{"leap year bug boundary 60", 60, time.Date(1900, 2, 29, 0, 0, 0, 0, time.UTC)},
		{"leap year bug boundary 61", 61, time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC)},
		// 現代の日付
		{"2024-01-01", 45292, time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)},
		// 時刻付き
		{"noon", 1.5, time.Date(1900, 1, 1, 12, 0, 0, 0, time.UTC)},
		{"6:30 AM", 1.270833333333, time.Date(1900, 1, 1, 6, 30, 0, 0, time.UTC)},
		// 時刻のみ（日付部分 0）
		{"time only 0:00", 0, time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := excelDateToTime(tt.serial)
			if err != nil {
				t.Fatalf("excelDateToTime(%v) error: %v", tt.serial, err)
			}
			if !got.Equal(tt.want) {
				t.Errorf("excelDateToTime(%v) = %v, want %v", tt.serial, got, tt.want)
			}
		})
	}
}

func TestExcelDateToTimeError(t *testing.T) {
	_, err := excelDateToTime(-1)
	if err == nil {
		t.Error("excelDateToTime(-1) expected error")
	}
}

func TestParseCellRef(t *testing.T) {
	tests := []struct {
		ref     string
		wantCol int
		wantRow int
	}{
		{"A1", 1, 1},
		{"B2", 2, 2},
		{"Z1", 26, 1},
		{"AA1", 27, 1},
		{"AZ100", 52, 100},
		// 小文字対応
		{"a1", 1, 1},
		{"ab12", 28, 12},
		// 混在（通常ありえないが堅牢性確認）
		{"Ab3", 28, 3},
		// 列部分なし
		{"123", 0, 123},
	}

	for _, tt := range tests {
		t.Run(tt.ref, func(t *testing.T) {
			col, row := parseCellRef(tt.ref)
			if col != tt.wantCol || row != tt.wantRow {
				t.Errorf("parseCellRef(%q) = (%d, %d), want (%d, %d)",
					tt.ref, col, row, tt.wantCol, tt.wantRow)
			}
		})
	}
}

func TestResolveRelTarget(t *testing.T) {
	tests := []struct {
		name     string
		basePath string
		target   string
		want     string
	}{
		{"relative", "xl/worksheets/sheet1.xml", "../drawings/drawing1.xml", "xl/drawings/drawing1.xml"},
		{"absolute", "xl/worksheets/sheet1.xml", "/xl/drawings/drawing1.xml", "xl/drawings/drawing1.xml"},
		{"same dir", "xl/worksheets/sheet1.xml", "sheet2.xml", "xl/worksheets/sheet2.xml"},
		{"double parent", "xl/worksheets/deep/sheet1.xml", "../../drawings/d.xml", "xl/drawings/d.xml"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := resolveRelTarget(tt.basePath, tt.target)
			if got != tt.want {
				t.Errorf("resolveRelTarget(%q, %q) = %q, want %q",
					tt.basePath, tt.target, got, tt.want)
			}
		})
	}
}
