package excel

import (
	"testing"
)

func TestDetectSheetType(t *testing.T) {
	tests := []struct {
		name    string
		relType string
		want    string
	}{
		{"worksheet", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheet"},
		{"chartsheet", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet", "chartsheet"},
		{"dialogsheet", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet", "dialogsheet"},
		{"macrosheet", "http://schemas.microsoft.com/office/2006/relationships/macrosheetx", "macrosheet"},
		{"empty", "", "worksheet"},
		{"unknown", "some-other-type", "worksheet"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := detectSheetType(tt.relType); got != tt.want {
				t.Errorf("detectSheetType(%q) = %q, want %q", tt.relType, got, tt.want)
			}
		})
	}
}

func TestRelsPathFor(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels"},
		{"xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			if got := relsPathFor(tt.input); got != tt.want {
				t.Errorf("relsPathFor(%q) = %q, want %q", tt.input, got, tt.want)
			}
		})
	}
}
