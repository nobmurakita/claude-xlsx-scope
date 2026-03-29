package excel

import (
	"math"
	"testing"
)

func TestNormalizeHexColor(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"FF0000", "#FF0000"},
		{"#FF0000", "#FF0000"},
		{"FFFF0000", "#FF0000"},   // ARGB → RGB
		{"ff0000", "#FF0000"},     // 小文字
		{"", ""},                  // 空文字列
		{"#FFFF0000", "#FF0000"},  // # 付き ARGB
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := normalizeHexColor(tt.input)
			if got != tt.want {
				t.Errorf("normalizeHexColor(%q) = %q, want %q", tt.input, got, tt.want)
			}
		})
	}
}

func TestRGBToHSLRoundTrip(t *testing.T) {
	tests := []struct {
		name    string
		r, g, b float64
	}{
		{"red", 1, 0, 0},
		{"green", 0, 1, 0},
		{"blue", 0, 0, 1},
		{"white", 1, 1, 1},
		{"black", 0, 0, 0},
		{"gray", 0.5, 0.5, 0.5},
		{"arbitrary", 0.2, 0.6, 0.8},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			h, s, l := rgbToHSL(tt.r, tt.g, tt.b)
			rr, gg, bb := hslToRGB(h, s, l)
			if math.Abs(rr-tt.r) > 1e-10 || math.Abs(gg-tt.g) > 1e-10 || math.Abs(bb-tt.b) > 1e-10 {
				t.Errorf("round trip (%v,%v,%v) → HSL(%v,%v,%v) → (%v,%v,%v)",
					tt.r, tt.g, tt.b, h, s, l, rr, gg, bb)
			}
		})
	}
}

func TestApplyTint(t *testing.T) {
	tests := []struct {
		name  string
		color string
		tint  float64
		want  string
	}{
		{"zero tint", "#FF0000", 0, "#FF0000"},
		{"positive tint lightens", "#000000", 0.5, "#808080"},
		{"negative tint darkens", "#FFFFFF", -0.5, "#808080"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := applyTint(tt.color, tt.tint)
			if got != tt.want {
				t.Errorf("applyTint(%q, %v) = %q, want %q", tt.color, tt.tint, got, tt.want)
			}
		})
	}
}
