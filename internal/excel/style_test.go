package excel

import (
	"testing"
)

func TestBuildFontObjFromParsed(t *testing.T) {
	defaultFont := FontInfo{Name: "Calibri", Size: 11}

	// デフォルトと同じ → nil
	pf := &parsedFont{Name: "Calibri", Size: 11}
	if obj := buildFontObjFromParsed(pf, defaultFont, nil); obj != nil {
		t.Errorf("expected nil for default font, got %+v", obj)
	}

	// nil → nil
	if obj := buildFontObjFromParsed(nil, defaultFont, nil); obj != nil {
		t.Error("expected nil for nil input")
	}

	// 差分あり
	pf = &parsedFont{Name: "Arial", Size: 14, Bold: true, Color: "#FF0000"}
	obj := buildFontObjFromParsed(pf, defaultFont, nil)
	if obj == nil {
		t.Fatal("expected non-nil font")
	}
	if obj.Name != "Arial" {
		t.Errorf("Name = %q, want %q", obj.Name, "Arial")
	}
	if obj.Size != 14 {
		t.Errorf("Size = %v, want 14", obj.Size)
	}
	if !obj.Bold {
		t.Error("Bold = false, want true")
	}
	if obj.Color != "#FF0000" {
		t.Errorf("Color = %q, want %q", obj.Color, "#FF0000")
	}

	// bold のみ差分
	pf = &parsedFont{Name: "Calibri", Size: 11, Bold: true}
	obj = buildFontObjFromParsed(pf, defaultFont, nil)
	if obj == nil {
		t.Fatal("expected non-nil for bold-only diff")
	}
	if obj.Name != "" {
		t.Errorf("Name = %q, want empty (same as default)", obj.Name)
	}
	if !obj.Bold {
		t.Error("Bold = false, want true")
	}
}

func TestBuildFillObjFromParsed(t *testing.T) {
	// nil → nil
	if obj := buildFillObjFromParsed(nil, nil); obj != nil {
		t.Error("expected nil for nil input")
	}

	// none パターン → nil
	pf := &parsedFill{PatternType: "none"}
	if obj := buildFillObjFromParsed(pf, nil); obj != nil {
		t.Error("expected nil for 'none' pattern")
	}

	// solid + FgColor
	pf = &parsedFill{PatternType: "solid", FgColor: "#D9E2F3"}
	obj := buildFillObjFromParsed(pf, nil)
	if obj == nil || obj.Color != "#D9E2F3" {
		t.Errorf("got %+v, want Color=#D9E2F3", obj)
	}

	// solid + no color → nil
	pf = &parsedFill{PatternType: "solid"}
	if obj := buildFillObjFromParsed(pf, nil); obj != nil {
		t.Errorf("expected nil for solid with no color, got %+v", obj)
	}
}

func TestBuildBorderObjFromParsed(t *testing.T) {
	// nil → nil
	if obj := buildBorderObjFromParsed(nil); obj != nil {
		t.Error("expected nil for nil input")
	}

	// 空スライス → nil
	if obj := buildBorderObjFromParsed([]parsedBorderEdge{}); obj != nil {
		t.Error("expected nil for empty edges")
	}

	// 罫線あり
	edges := []parsedBorderEdge{
		{Type: "top", Style: "thin"},
		{Type: "bottom", Style: "medium", Color: "FFFF0000"},
		{Type: "diagonal_up", Style: "thin"},
	}
	obj := buildBorderObjFromParsed(edges)
	if obj == nil {
		t.Fatal("expected non-nil border")
	}
	if obj.Top == nil || obj.Top.Style != "thin" {
		t.Errorf("Top = %+v, want thin", obj.Top)
	}
	if obj.Bottom == nil || obj.Bottom.Style != "medium" || obj.Bottom.Color != "#FF0000" {
		t.Errorf("Bottom = %+v, want medium with #FF0000", obj.Bottom)
	}
	if obj.DiagonalUp == nil || obj.DiagonalUp.Style != "thin" {
		t.Errorf("DiagonalUp = %+v, want thin", obj.DiagonalUp)
	}
	if obj.Left != nil || obj.Right != nil {
		t.Error("Left/Right should be nil")
	}

	// 黒色は省略される
	if obj.Top.Color != "" {
		t.Errorf("Top.Color = %q, want empty (default black omitted)", obj.Top.Color)
	}
}

func TestBuildAlignmentObjFromParsed(t *testing.T) {
	// nil → nil
	if obj := buildAlignmentObjFromParsed(nil); obj != nil {
		t.Error("expected nil for nil input")
	}

	// デフォルト値 → nil
	pa := &parsedAlignment{Horizontal: "general", Vertical: "bottom"}
	if obj := buildAlignmentObjFromParsed(pa); obj != nil {
		t.Errorf("expected nil for default alignment, got %+v", obj)
	}

	// 値あり
	pa = &parsedAlignment{Horizontal: "center", Vertical: "top", WrapText: true, Indent: 2}
	obj := buildAlignmentObjFromParsed(pa)
	if obj == nil {
		t.Fatal("expected non-nil alignment")
	}
	if obj.Horizontal != "center" || obj.Vertical != "top" || !obj.Wrap || obj.Indent != 2 {
		t.Errorf("got %+v", obj)
	}
}
