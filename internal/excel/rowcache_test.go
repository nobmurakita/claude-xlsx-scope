package excel

import (
	"testing"
)

func TestRowCache(t *testing.T) {
	rc := NewRowCache()

	// 空の場合
	if got := rc.CalcUsedRange(); got != "" {
		t.Errorf("empty cache: CalcUsedRange() = %q, want empty", got)
	}

	// セルを追加
	rc.Add(2, 3) // B3
	rc.Add(5, 1) // E1
	rc.Add(1, 7) // A7

	got := rc.CalcUsedRange()
	if got != "A1:E7" {
		t.Errorf("CalcUsedRange() = %q, want %q", got, "A1:E7")
	}
}

func TestRowCache_SingleCell(t *testing.T) {
	rc := NewRowCache()
	rc.Add(3, 5) // C5

	got := rc.CalcUsedRange()
	if got != "C5:C5" {
		t.Errorf("CalcUsedRange() = %q, want %q", got, "C5:C5")
	}
}
