package excel

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestParseSharedStringsSAX(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3">
  <si><t>plain text</t></si>
  <si>
    <r><t>normal</t></r>
    <r><rPr><b/><color rgb="FFFF0000"/></rPr><t>bold red</t></r>
  </si>
  <si><t>third</t></si>
</sst>`

	ss := &sharedStrings{}
	decoder := xml.NewDecoder(strings.NewReader(xmlData))
	if err := parseSharedStringsSAX(decoder, ss); err != nil {
		t.Fatalf("parseSharedStringsSAX error: %v", err)
	}

	if len(ss.items) != 3 {
		t.Fatalf("got %d items, want 3", len(ss.items))
	}

	// プレーンテキスト
	if ss.Get(0) != "plain text" {
		t.Errorf("Get(0) = %q, want %q", ss.Get(0), "plain text")
	}
	if ss.GetRichTextRuns(0) != nil {
		t.Errorf("GetRichTextRuns(0) should be nil for plain text")
	}

	// リッチテキスト
	if ss.Get(1) != "normalbold red" {
		t.Errorf("Get(1) = %q, want %q", ss.Get(1), "normalbold red")
	}
	runs := ss.GetRichTextRuns(1)
	if len(runs) != 2 {
		t.Fatalf("GetRichTextRuns(1) has %d runs, want 2", len(runs))
	}
	if runs[0].Text != "normal" || runs[0].Font != nil {
		t.Errorf("runs[0] = {Text:%q, Font:%v}, want {\"normal\", nil}", runs[0].Text, runs[0].Font)
	}
	if runs[1].Text != "bold red" || runs[1].Font == nil {
		t.Fatalf("runs[1] = {Text:%q, Font:%v}, want bold font", runs[1].Text, runs[1].Font)
	}
	if !runs[1].Font.Bold {
		t.Error("runs[1].Font.Bold = false, want true")
	}
	if runs[1].Font.Color != "#FF0000" {
		t.Errorf("runs[1].Font.Color = %q, want %q", runs[1].Font.Color, "#FF0000")
	}

	// 3番目
	if ss.Get(2) != "third" {
		t.Errorf("Get(2) = %q, want %q", ss.Get(2), "third")
	}
}

func TestSharedStrings_GetOutOfRange(t *testing.T) {
	ss := &sharedStrings{items: []sharedStringItem{{Text: "a"}}}

	if ss.Get(-1) != "" {
		t.Error("Get(-1) should return empty")
	}
	if ss.Get(1) != "" {
		t.Error("Get(1) should return empty for 1-item list")
	}
	if ss.GetRichTextRuns(-1) != nil {
		t.Error("GetRichTextRuns(-1) should return nil")
	}

	// nil sharedStrings
	var nilSS *sharedStrings
	if nilSS.Get(0) != "" {
		t.Error("nil.Get(0) should return empty")
	}
}

func TestParseSharedStringsSAX_WithRuby(t *testing.T) {
	// ルビ（rPh）は無視されるべき
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si>
    <t>漢字</t>
    <rPh sb="0" eb="2"><t>かんじ</t></rPh>
  </si>
</sst>`

	ss := &sharedStrings{}
	decoder := xml.NewDecoder(strings.NewReader(xmlData))
	if err := parseSharedStringsSAX(decoder, ss); err != nil {
		t.Fatalf("error: %v", err)
	}

	if ss.Get(0) != "漢字" {
		t.Errorf("Get(0) = %q, want %q (ruby should be ignored)", ss.Get(0), "漢字")
	}
}
