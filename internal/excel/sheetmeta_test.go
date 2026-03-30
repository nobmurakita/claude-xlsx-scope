package excel

import (
	"testing"
)

func TestEffectiveDefaultWidth(t *testing.T) {
	tests := []struct {
		name  string
		width float64
		want  float64
	}{
		{"explicit width", 12.5, 12.5},
		{"zero uses default", 0, DefaultColWidth},
		{"negative uses default", -1, DefaultColWidth},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			sm := &SheetMeta{DefaultWidth: tt.width}
			if got := sm.EffectiveDefaultWidth(); got != tt.want {
				t.Errorf("EffectiveDefaultWidth() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestBuildHyperlinkMap(t *testing.T) {
	sm := &SheetMeta{
		Hyperlinks: []HyperlinkEntry{
			{Ref: "A1", RID: "rId1"},                   // 外部リンク
			{Ref: "B2", Location: "Sheet2!A1"},          // 内部リンク
			{Ref: "C3", RID: "rId2"},                    // rels に存在しないRID
			{Ref: "D4", RID: "rId3", Location: "local"}, // RID優先
		},
	}

	rels := map[string]string{
		"rId1": "https://example.com",
		"rId3": "mailto:test@example.com",
	}

	m := sm.BuildHyperlinkMap(rels)

	// A1: 外部URL
	if link, ok := m["A1"]; !ok || link.URL != "https://example.com" {
		t.Errorf("A1 link = %+v, want URL=https://example.com", link)
	}

	// B2: 内部リンク
	if link, ok := m["B2"]; !ok || link.Location != "Sheet2!A1" {
		t.Errorf("B2 link = %+v, want Location=Sheet2!A1", link)
	}

	// C3: rels に存在しないRIDは含まれない
	if _, ok := m["C3"]; ok {
		t.Error("C3 should not be in map (rId2 not in rels)")
	}

	// D4: RID が rels にあるので外部リンクとして解決
	if link, ok := m["D4"]; !ok || link.URL != "mailto:test@example.com" {
		t.Errorf("D4 link = %+v, want URL=mailto:test@example.com", link)
	}
}

func TestBuildHyperlinkMap_Empty(t *testing.T) {
	sm := &SheetMeta{}
	m := sm.BuildHyperlinkMap(nil)
	if len(m) != 0 {
		t.Errorf("expected empty map, got %d entries", len(m))
	}
}
