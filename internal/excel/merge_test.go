package excel

import (
	"testing"
)

func TestBuildMergeInfo(t *testing.T) {
	sm := &SheetMeta{
		MergeCells: []MergeCellRange{
			{Ref: "A1:C3"},
			{Ref: "E5:F6"},
		},
	}
	mi := sm.BuildMergeInfo()

	// A1 は左上
	if merge, ok := mi.IsTopLeft(1, 1); !ok || merge != "A1:C3" {
		t.Errorf("A1 should be topLeft of A1:C3, got %q, %v", merge, ok)
	}

	// B2 は結合セルの左上以外
	if !mi.IsMergedNonTopLeft(2, 2) {
		t.Error("B2 should be merged non-top-left")
	}

	// D4 はマージなし
	if _, ok := mi.IsTopLeft(4, 4); ok {
		t.Error("D4 should not be topLeft")
	}
	if mi.IsMergedNonTopLeft(4, 4) {
		t.Error("D4 should not be merged")
	}

	// E5 は左上
	if merge, ok := mi.IsTopLeft(5, 5); !ok || merge != "E5:F6" {
		t.Errorf("E5 should be topLeft of E5:F6, got %q, %v", merge, ok)
	}

	// F6 は結合セルの左上以外
	if !mi.IsMergedNonTopLeft(6, 6) {
		t.Error("F6 should be merged non-top-left")
	}
}
