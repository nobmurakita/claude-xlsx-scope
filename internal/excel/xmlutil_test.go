package excel

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestSafeAtoi(t *testing.T) {
	tests := []struct {
		input string
		want  int
	}{
		{"42", 42},
		{"0", 0},
		{"-1", -1},
		{"", 0},
		{"abc", 0},
		{"3.14", 0},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			if got := safeAtoi(tt.input); got != tt.want {
				t.Errorf("safeAtoi(%q) = %d, want %d", tt.input, got, tt.want)
			}
		})
	}
}

func TestAttrVal(t *testing.T) {
	se := xml.StartElement{
		Name: xml.Name{Local: "test"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "name"}, Value: "hello"},
			{Name: xml.Name{Local: "id"}, Value: "42"},
		},
	}

	if got := attrVal(se, "name"); got != "hello" {
		t.Errorf("attrVal(name) = %q, want %q", got, "hello")
	}
	if got := attrVal(se, "id"); got != "42" {
		t.Errorf("attrVal(id) = %q, want %q", got, "42")
	}
	if got := attrVal(se, "missing"); got != "" {
		t.Errorf("attrVal(missing) = %q, want empty", got)
	}
}

func TestSkipElement(t *testing.T) {
	xmlData := `<outer><inner><deep>text</deep></inner></outer><after/>`
	decoder := xml.NewDecoder(strings.NewReader(xmlData))

	// <outer> を読む
	tok, _ := decoder.Token()
	if se, ok := tok.(xml.StartElement); !ok || se.Name.Local != "outer" {
		t.Fatal("expected <outer>")
	}

	// skipElement で <outer> の中身を全てスキップ
	skipElement(decoder)

	// 次のトークンは <after/> であるべき
	tok, _ = decoder.Token()
	if se, ok := tok.(xml.StartElement); !ok || se.Name.Local != "after" {
		t.Errorf("expected <after/> after skip, got %T %v", tok, tok)
	}
}
