package excel

import (
	"testing"
)

func TestRowCache(t *testing.T) {
	rc := newRowCache()

	// 空の場合
	if got := rc.calcUsedRange(); got != "" {
		t.Errorf("empty cache: calcUsedRange() = %q, want empty", got)
	}

	// セルを追加
	rc.add(2, 3) // B3
	rc.add(5, 1) // E1
	rc.add(1, 7) // A7

	got := rc.calcUsedRange()
	if got != "A1:E7" {
		t.Errorf("calcUsedRange() = %q, want %q", got, "A1:E7")
	}
}

func TestRowCache_SingleCell(t *testing.T) {
	rc := newRowCache()
	rc.add(3, 5) // C5

	got := rc.calcUsedRange()
	if got != "C5:C5" {
		t.Errorf("calcUsedRange() = %q, want %q", got, "C5:C5")
	}
}
