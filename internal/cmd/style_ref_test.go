package cmd

import (
	"testing"

	"github.com/nobmurakita/claude-xlsx-scope/internal/excel"
)

// mockStyler はテスト用のスタイラー
type mockStyler struct {
	styles map[int]*styleResult
}

func (m *mockStyler) StyleByID(styleID int, _ excel.FontInfo) (*excel.FontObj, *excel.FillObj, *excel.BorderObj, *excel.AlignmentObj) {
	if sr, ok := m.styles[styleID]; ok {
		return sr.font, sr.fill, sr.border, sr.alignment
	}
	return nil, nil, nil, nil
}

func (m *mockStyler) GetRichText(_ int, _ *excel.FontObj, _ excel.FontInfo) []excel.RichTextRun {
	return nil
}

func TestResolveStyleRef_NewStyleReturnsDef(t *testing.T) {
	dc := &cellsContext{
		showStyle:      true,
		styleCache:     make(map[int]*styleResult),
		styleRefMap:    make(map[int]int),
		hiddenColCache: make(map[int]bool),
		styler: &mockStyler{styles: map[int]*styleResult{
			1: {font: &excel.FontObj{Bold: true}},
		}},
	}

	idx, def := dc.resolveStyleRef(1)
	if idx != 1 {
		t.Errorf("idx = %d, want 1", idx)
	}
	if def == nil {
		t.Fatal("def should not be nil for first occurrence")
	}
	if def.StyleDef != 1 {
		t.Errorf("def.StyleDef = %d, want 1", def.StyleDef)
	}
	if def.Font == nil || !def.Font.Bold {
		t.Error("def.Font should have Bold=true")
	}
}

func TestResolveStyleRef_SameStyleReturnsNilDef(t *testing.T) {
	dc := &cellsContext{
		showStyle:      true,
		styleCache:     make(map[int]*styleResult),
		styleRefMap:    make(map[int]int),
		hiddenColCache: make(map[int]bool),
		styler: &mockStyler{styles: map[int]*styleResult{
			1: {font: &excel.FontObj{Bold: true}},
		}},
	}

	// 初回
	dc.resolveStyleRef(1)

	// 2回目: 定義は返さない
	idx, def := dc.resolveStyleRef(1)
	if idx != 1 {
		t.Errorf("idx = %d, want 1", idx)
	}
	if def != nil {
		t.Error("def should be nil for second occurrence")
	}
}

func TestResolveStyleRef_EmptyStyleReturnsZero(t *testing.T) {
	dc := &cellsContext{
		showStyle:      true,
		styleCache:     make(map[int]*styleResult),
		styleRefMap:    make(map[int]int),
		hiddenColCache: make(map[int]bool),
		styler: &mockStyler{styles: map[int]*styleResult{
			1: {}, // 全フィールド nil
		}},
	}

	idx, def := dc.resolveStyleRef(1)
	if idx != 0 {
		t.Errorf("idx = %d, want 0 for empty style", idx)
	}
	if def != nil {
		t.Error("def should be nil for empty style")
	}
}

func TestResolveStyleRef_StyleIDZeroReturnsZero(t *testing.T) {
	dc := &cellsContext{
		showStyle:      true,
		styleCache:     make(map[int]*styleResult),
		styleRefMap:    make(map[int]int),
		hiddenColCache: make(map[int]bool),
		styler:         &mockStyler{styles: map[int]*styleResult{}},
	}

	idx, def := dc.resolveStyleRef(0)
	if idx != 0 {
		t.Errorf("idx = %d, want 0", idx)
	}
	if def != nil {
		t.Error("def should be nil for styleID 0")
	}
}

func TestResolveStyleRef_MultipleStyles(t *testing.T) {
	dc := &cellsContext{
		showStyle:      true,
		styleCache:     make(map[int]*styleResult),
		styleRefMap:    make(map[int]int),
		hiddenColCache: make(map[int]bool),
		styler: &mockStyler{styles: map[int]*styleResult{
			1: {font: &excel.FontObj{Bold: true}},
			2: {fill: &excel.FillObj{Color: "#FF0000"}},
			3: {}, // 空
		}},
	}

	idx1, def1 := dc.resolveStyleRef(1)
	idx2, def2 := dc.resolveStyleRef(2)
	idx3, def3 := dc.resolveStyleRef(3)

	if idx1 != 1 || def1 == nil {
		t.Errorf("style 1: idx=%d, def=%v", idx1, def1)
	}
	if idx2 != 2 || def2 == nil {
		t.Errorf("style 2: idx=%d, def=%v", idx2, def2)
	}
	if idx3 != 0 || def3 != nil {
		t.Errorf("style 3 (empty): idx=%d, def=%v", idx3, def3)
	}

	// 再参照: 定義なし
	idx1again, def1again := dc.resolveStyleRef(1)
	if idx1again != 1 || def1again != nil {
		t.Errorf("style 1 again: idx=%d, def should be nil", idx1again)
	}
}

func TestBuildCellOutput_WithStyleRef(t *testing.T) {
	dc := &cellsContext{
		showStyle:      true,
		styleCache:     make(map[int]*styleResult),
		styleRefMap:    make(map[int]int),
		hiddenColCache: make(map[int]bool),
		mergeInfo:      (&excel.SheetMeta{}).BuildMergeInfo(),
		styler: &mockStyler{styles: map[int]*styleResult{
			5: {font: &excel.FontObj{Bold: true}, fill: &excel.FillObj{Color: "#CCFFFF"}},
		}},
	}

	data := &excel.CellData{
		Value:   "テスト",
		Type:    excel.CellTypeString,
		StyleID: 5,
	}

	out, def := dc.buildCellOutput(1, 1, data, nil)

	// スタイル定義が返る
	if def == nil {
		t.Fatal("expected style definition for first occurrence")
	}
	if def.Font == nil || !def.Font.Bold {
		t.Error("style def should have Bold font")
	}
	if def.Fill == nil || def.Fill.Color != "#CCFFFF" {
		t.Error("style def should have Fill color")
	}

	// セル出力はスタイル参照のみ
	if out.StyleRef == nil || *out.StyleRef != 1 {
		t.Errorf("out.StyleRef = %v, want 1", out.StyleRef)
	}

	// 2回目: 定義なし
	data2 := &excel.CellData{Value: "テスト2", Type: excel.CellTypeString, StyleID: 5}
	out2, def2 := dc.buildCellOutput(2, 1, data2, nil)
	if def2 != nil {
		t.Error("second cell should not return style definition")
	}
	if out2.StyleRef == nil || *out2.StyleRef != 1 {
		t.Errorf("out2.StyleRef = %v, want 1", out2.StyleRef)
	}
}

func TestBuildCellOutput_NoStyleWithoutFlag(t *testing.T) {
	dc := &cellsContext{
		showStyle:      false,
		styleCache:     make(map[int]*styleResult),
		styleRefMap:    make(map[int]int),
		hiddenColCache: make(map[int]bool),
		mergeInfo:      (&excel.SheetMeta{}).BuildMergeInfo(),
		styler:         &mockStyler{styles: map[int]*styleResult{}},
	}

	data := &excel.CellData{Value: "hello", Type: excel.CellTypeString, StyleID: 5}
	out, def := dc.buildCellOutput(1, 1, data, nil)

	if def != nil {
		t.Error("should not return style def when showStyle is false")
	}
	if out.StyleRef != nil {
		t.Error("should not have StyleRef when showStyle is false")
	}
}
