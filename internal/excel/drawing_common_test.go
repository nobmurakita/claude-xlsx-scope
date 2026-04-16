package excel

import (
	"encoding/xml"
	"testing"
)

func TestConnectorEndpoints(t *testing.T) {
	pos := &Position{X: 10, Y: 20, W: 100, H: 50}

	tests := []struct {
		name                   string
		flip                   string
		wantStartX, wantStartY int
		wantEndX, wantEndY     int
	}{
		{"no flip", "", 10, 20, 110, 70},
		{"flip h", "h", 110, 20, 10, 70},
		{"flip v", "v", 10, 70, 110, 20},
		{"flip hv", "hv", 110, 70, 10, 20},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			start, end := connectorEndpoints(pos, tt.flip)
			if start.X != tt.wantStartX || start.Y != tt.wantStartY {
				t.Errorf("start = (%d,%d), want (%d,%d)", start.X, start.Y, tt.wantStartX, tt.wantStartY)
			}
			if end.X != tt.wantEndX || end.Y != tt.wantEndY {
				t.Errorf("end = (%d,%d), want (%d,%d)", end.X, end.Y, tt.wantEndX, tt.wantEndY)
			}
		})
	}
}

func TestCalcCalloutTarget_Wedge(t *testing.T) {
	pos := &Position{X: 100, Y: 100, W: 200, H: 100}

	// デフォルト adj 値を使用
	pt := calcCalloutTarget(pos, "wedgeRectCallout", nil)
	if pt == nil {
		t.Fatal("expected non-nil point for wedgeRectCallout")
	}
	// adj1=-20833, adj2=62500 → x=100+(-20833*200/100000)=-42+100=58, y=100+(62500*100/100000)=63+100=163
	// 正確な丸め値に依存するので、おおよその範囲をチェック
	if pt.X > 100 || pt.X < 50 {
		t.Errorf("X = %d, expected around 58", pt.X)
	}
	if pt.Y < 150 || pt.Y > 170 {
		t.Errorf("Y = %d, expected around 163", pt.Y)
	}
}

func TestCalcCalloutTarget_CustomAdj(t *testing.T) {
	pos := &Position{X: 0, Y: 0, W: 100, H: 100}

	adjs := map[string]int{"adj1": 50000, "adj2": 50000} // 中央
	pt := calcCalloutTarget(pos, "wedgeEllipseCallout", adjs)
	if pt == nil {
		t.Fatal("expected non-nil point")
	}
	if pt.X != 50 || pt.Y != 50 {
		t.Errorf("point = (%d,%d), want (50,50)", pt.X, pt.Y)
	}
}

func TestCalcCalloutTarget_Unknown(t *testing.T) {
	pos := &Position{X: 0, Y: 0, W: 100, H: 100}
	pt := calcCalloutTarget(pos, "rect", nil)
	if pt != nil {
		t.Error("expected nil for non-callout shape")
	}
}

func TestFinalizeLineStyle(t *testing.T) {
	tests := []struct {
		name string
		ls   *LineStyle
		want *LineStyle
	}{
		{"nil", nil, nil},
		{"empty", &LineStyle{}, nil},
		{"color only", &LineStyle{Color: "#FF0000"}, &LineStyle{Color: "#FF0000", Style: "solid"}},
		{"width only", &LineStyle{Width: 1.5}, &LineStyle{Width: 1.5, Style: "solid"}},
		{"with style", &LineStyle{Color: "#000", Style: "dash"}, &LineStyle{Color: "#000", Style: "dash"}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := finalizeLineStyle(tt.ls)
			if got == nil && tt.want == nil {
				return
			}
			if got == nil || tt.want == nil {
				t.Errorf("got %v, want %v", got, tt.want)
				return
			}
			if got.Color != tt.want.Color || got.Style != tt.want.Style || got.Width != tt.want.Width {
				t.Errorf("got %+v, want %+v", got, tt.want)
			}
		})
	}
}

func TestUpdateArrow(t *testing.T) {
	tests := []struct {
		name       string
		current    string
		headOrTail string
		arrowType  string
		want       string
	}{
		{"tail arrow", "", "tail", "triangle", "end"},
		{"head arrow", "", "head", "triangle", "start"},
		{"both: start+end", "start", "tail", "triangle", "both"},
		{"both: end+start", "end", "head", "triangle", "both"},
		{"none type ignored", "", "tail", "none", ""},
		{"empty type ignored", "", "head", "", ""},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			current := tt.current
			updateArrow(&current, tt.headOrTail, tt.arrowType)
			if current != tt.want {
				t.Errorf("got %q, want %q", current, tt.want)
			}
		})
	}
}

func TestBuildDrawingFontObj(t *testing.T) {
	theme := &themeColors{colors: make([]string, 12)}

	// 空フォント → nil
	pf := &parsedFont{}
	if obj := buildDrawingFontObj(pf, theme); obj != nil {
		t.Errorf("expected nil for empty font, got %+v", obj)
	}

	// nil → nil
	if obj := buildDrawingFontObj(nil, theme); obj != nil {
		t.Error("expected nil for nil input")
	}

	// 属性あり
	pf = &parsedFont{Name: "Arial", Size: 14, Bold: true, Color: "#FF0000"}
	obj := buildDrawingFontObj(pf, theme)
	if obj == nil {
		t.Fatal("expected non-nil font")
	}
	if obj.Name != "Arial" || obj.Size != 14 || !obj.Bold || obj.Color != "#FF0000" {
		t.Errorf("got %+v", obj)
	}
}

func TestParseCNvPr(t *testing.T) {
	tests := []struct {
		name       string
		attrs      []xml.Attr
		wantName   string
		wantID     int
		wantHidden bool
	}{
		{
			"normal",
			[]xml.Attr{
				{Name: xml.Name{Local: "name"}, Value: "Shape 1"},
				{Name: xml.Name{Local: "id"}, Value: "5"},
			},
			"Shape 1", 5, false,
		},
		{
			"no id",
			[]xml.Attr{
				{Name: xml.Name{Local: "name"}, Value: "Shape 2"},
			},
			"Shape 2", 0, false,
		},
		{
			"hidden form control legacy",
			[]xml.Attr{
				{Name: xml.Name{Local: "id"}, Value: "73729"},
				{Name: xml.Name{Local: "name"}, Value: "Check Box 1"},
				{Name: xml.Name{Local: "hidden"}, Value: "1"},
			},
			"Check Box 1", 73729, true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			se := xml.StartElement{
				Name: xml.Name{Local: "cNvPr"},
				Attr: tt.attrs,
			}
			name, id, hidden := parseCNvPr(se)
			if name != tt.wantName || id != tt.wantID || hidden != tt.wantHidden {
				t.Errorf("got (%q, %d, %v), want (%q, %d, %v)", name, id, hidden, tt.wantName, tt.wantID, tt.wantHidden)
			}
		})
	}
}
