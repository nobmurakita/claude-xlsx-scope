package excel

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestParseVMLShapeID(t *testing.T) {
	tests := []struct {
		in   string
		want int
	}{
		{"_x0000_s73729", 73729},
		{"_x0000_s1", 1},
		{"abc", 0},
		{"", 0},
	}
	for _, tt := range tests {
		if got := parseVMLShapeID(tt.in); got != tt.want {
			t.Errorf("parseVMLShapeID(%q) = %d, want %d", tt.in, got, tt.want)
		}
	}
}

func TestParseVMLZIndex(t *testing.T) {
	tests := []struct {
		in      string
		wantN   int
		wantOK  bool
	}{
		{"position:absolute;z-index:1;mso-wrap-style:tight", 1, true},
		{"z-index:42", 42, true},
		{"Z-Index: 7 ;foo", 7, true},
		{"position:absolute", 0, false},
		{"z-index:abc", 0, false},
		{"", 0, false},
	}
	for _, tt := range tests {
		n, ok := parseVMLZIndex(tt.in)
		if n != tt.wantN || ok != tt.wantOK {
			t.Errorf("parseVMLZIndex(%q) = (%d, %v), want (%d, %v)", tt.in, n, ok, tt.wantN, tt.wantOK)
		}
	}
}

func TestParseFormControlPrAttrs(t *testing.T) {
	tests := []struct {
		name string
		xml  string
		want formControlPr
	}{
		{
			"checkbox checked",
			`<formControlPr objectType="CheckBox" checked="Checked" lockText="1" noThreeD="1"/>`,
			formControlPr{ObjectType: "CheckBox", Checked: "Checked"},
		},
		{
			"checkbox unchecked with fmlaLink",
			`<formControlPr objectType="CheckBox" fmlaLink="$A$1" noThreeD="1"/>`,
			formControlPr{ObjectType: "CheckBox", FmlaLink: "$A$1"},
		},
		{
			"drop with selection",
			`<formControlPr objectType="Drop" fmlaRange="$B$1:$B$5" fmlaLink="$C$1" sel="2" dropLines="4"/>`,
			formControlPr{ObjectType: "Drop", FmlaRange: "$B$1:$B$5", FmlaLink: "$C$1", Sel: 2, DropLines: 4},
		},
		{
			"list multi",
			`<formControlPr objectType="List" fmlaRange="$B$1:$B$5" selType="Multi" sel="1"/>`,
			formControlPr{ObjectType: "List", FmlaRange: "$B$1:$B$5", SelType: "Multi", Sel: 1},
		},
		{
			"spin",
			`<formControlPr objectType="Spin" min="0" max="100" val="10" inc="5" fmlaLink="$D$1"/>`,
			formControlPr{
				ObjectType: "Spin", FmlaLink: "$D$1",
				Min: ptrInt(0), Max: ptrInt(100), Val: ptrInt(10), Inc: ptrInt(5),
			},
		},
		{
			"scroll",
			`<formControlPr objectType="Scroll" min="0" max="100" val="50" inc="1" page="10"/>`,
			formControlPr{
				ObjectType: "Scroll",
				Min:        ptrInt(0), Max: ptrInt(100), Val: ptrInt(50), Inc: ptrInt(1), Page: ptrInt(10),
			},
		},
		{
			"button macro",
			`<formControlPr objectType="Button" fmlaMacro="MyMacro"/>`,
			formControlPr{ObjectType: "Button", FmlaMacro: "MyMacro"},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			dec := xml.NewDecoder(strings.NewReader(tt.xml))
			tok, err := dec.Token()
			if err != nil {
				t.Fatal(err)
			}
			se := tok.(xml.StartElement)
			got := parseFormControlPrAttrs(se)
			if !equalFormControlPr(*got, tt.want) {
				t.Errorf("got %+v, want %+v", *got, tt.want)
			}
		})
	}
}

func TestParseVMLShapesSAX_Checkbox(t *testing.T) {
	vml := `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
<v:shape id="_x0000_s73729" type="#_x0000_t201" style='position:absolute;z-index:1'>
  <v:textbox>
    <div><font face="MS P Gothic" size="180">入力</font></div>
  </v:textbox>
  <x:ClientData ObjectType="Checkbox"><x:Checked>1</x:Checked></x:ClientData>
</v:shape>
<v:shape id="_x0000_s73730" style='z-index:2'>
  <v:textbox>
    <div>行1</div>
    <div>行2</div>
  </v:textbox>
</v:shape>
<v:shape id="_x0000_s73731" style='z-index:3'></v:shape>
</xml>`

	vmlMap := make(map[int]vmlShapeInfo)
	dec := xml.NewDecoder(strings.NewReader(vml))
	parseVMLShapesSAX(dec, vmlMap)

	if got := vmlMap[73729]; got.Text != "入力" || got.ZIndex != 1 || !got.HasZ {
		t.Errorf("shape 73729: %+v", got)
	}
	if got := vmlMap[73730]; got.Text != "行1\n行2" || got.ZIndex != 2 {
		t.Errorf("shape 73730 text=%q z=%d, want %q z=2", got.Text, got.ZIndex, "行1\n行2")
	}
	if got := vmlMap[73731]; got.Text != "" || got.ZIndex != 3 || !got.HasShape {
		t.Errorf("shape 73731: %+v", got)
	}
}

func TestParseSheetControlsSAX(t *testing.T) {
	sheet := `<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<controls>
  <mc:AlternateContent>
    <mc:Choice Requires="x14">
      <control shapeId="73729" r:id="rId4" name="Check Box 1">
        <controlPr>
          <anchor>
            <from><xdr:col>4</xdr:col><xdr:colOff>100</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>200</xdr:rowOff></from>
            <to><xdr:col>8</xdr:col><xdr:colOff>300</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>400</xdr:rowOff></to>
          </anchor>
        </controlPr>
      </control>
    </mc:Choice>
    <mc:Fallback>
      <control shapeId="99999" r:id="rIdX" name="Ignore Me"/>
    </mc:Fallback>
  </mc:AlternateContent>
  <control shapeId="73730" r:id="rId5" name="Check Box 2">
    <controlPr>
      <anchor>
        <from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></from>
        <to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></to>
      </anchor>
    </controlPr>
  </control>
</controls>
</worksheet>`

	dec := xml.NewDecoder(strings.NewReader(sheet))
	controls := parseSheetControlsSAX(dec)

	if len(controls) != 2 {
		t.Fatalf("want 2 controls, got %d", len(controls))
	}
	c1 := controls[0]
	if c1.ShapeID != 73729 || c1.RelID != "rId4" || c1.Name != "Check Box 1" {
		t.Errorf("c1 attrs: %+v", c1)
	}
	if c1.From.col != 4 || c1.From.colOff != 100 || c1.From.row != 6 || c1.From.rowOff != 200 {
		t.Errorf("c1 from: %+v", c1.From)
	}
	if c1.To.col != 8 || c1.To.rowOff != 400 {
		t.Errorf("c1 to: %+v", c1.To)
	}
	c2 := controls[1]
	if c2.ShapeID != 73730 || c2.RelID != "rId5" {
		t.Errorf("c2 attrs: %+v", c2)
	}
}

func TestBuildCellRangeFromAnchor(t *testing.T) {
	tests := []struct {
		from, to anchorPos
		want     string
	}{
		{anchorPos{col: 0, row: 0}, anchorPos{col: 0, row: 0}, "A1"},
		{anchorPos{col: 0, row: 0}, anchorPos{col: 2, row: 1}, "A1:C2"},
		{anchorPos{col: 4, row: 6}, anchorPos{col: 8, row: 7}, "E7:I8"},
	}
	for _, tt := range tests {
		if got := buildCellRangeFromAnchor(tt.from, tt.to); got != tt.want {
			t.Errorf("from=%+v to=%+v -> %q, want %q", tt.from, tt.to, got, tt.want)
		}
	}
}

func TestApplyControlProps_Variants(t *testing.T) {
	vml := vmlShapeInfo{Text: "ラベル"}

	t.Run("checkbox checked", func(t *testing.T) {
		s := &ShapeInfo{}
		applyControlProps(s, ShapeTypeCheckbox, &formControlPr{Checked: "Checked", FmlaLink: "$A$1"}, vml)
		if s.Text != "ラベル" || s.Checked == nil || !*s.Checked || s.LinkedCell != "$A$1" {
			t.Errorf("got %+v", s)
		}
	})
	t.Run("checkbox unchecked", func(t *testing.T) {
		s := &ShapeInfo{}
		applyControlProps(s, ShapeTypeCheckbox, &formControlPr{}, vml)
		if s.Checked == nil || *s.Checked {
			t.Errorf("checked should be false, got %+v", s.Checked)
		}
	})
	t.Run("list multi", func(t *testing.T) {
		s := &ShapeInfo{}
		applyControlProps(s, ShapeTypeList, &formControlPr{FmlaRange: "$B$1:$B$5", SelType: "Multi", Sel: 2}, vml)
		if s.ListRange != "$B$1:$B$5" || s.SelType != "multi" || s.SelectedIndex != 2 {
			t.Errorf("got %+v", s)
		}
	})
	t.Run("spin", func(t *testing.T) {
		s := &ShapeInfo{}
		applyControlProps(s, ShapeTypeSpin, &formControlPr{Min: ptrInt(0), Max: ptrInt(10), Val: ptrInt(5), Inc: ptrInt(1)}, vml)
		if *s.Min != 0 || *s.Max != 10 || *s.Val != 5 || *s.Inc != 1 {
			t.Errorf("got %+v", s)
		}
	})
	t.Run("button macro", func(t *testing.T) {
		s := &ShapeInfo{}
		applyControlProps(s, ShapeTypeButton, &formControlPr{FmlaMacro: "RunMe"}, vml)
		if s.Macro != "RunMe" {
			t.Errorf("got %+v", s)
		}
	})
}

func TestNextTopZ(t *testing.T) {
	one := 1
	tests := []struct {
		name   string
		shapes []ShapeInfo
		want   int
	}{
		{"empty", nil, 0},
		{"one top", []ShapeInfo{{ID: 1, Z: 0}}, 1},
		{"multi top", []ShapeInfo{{ID: 1, Z: 0}, {ID: 2, Z: 3}}, 4},
		{"children ignored", []ShapeInfo{{ID: 1, Z: 0}, {ID: 2, Z: 5, Parent: &one}}, 1},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := nextTopZ(tt.shapes); got != tt.want {
				t.Errorf("got %d, want %d", got, tt.want)
			}
		})
	}
}

func TestNextShapeID(t *testing.T) {
	tests := []struct {
		name   string
		shapes []ShapeInfo
		want   int
	}{
		{"empty", nil, 1},
		{"one", []ShapeInfo{{ID: 1}}, 2},
		{"non sequential", []ShapeInfo{{ID: 5}, {ID: 3}}, 6},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := nextShapeID(tt.shapes); got != tt.want {
				t.Errorf("got %d, want %d", got, tt.want)
			}
		})
	}
}

// test helpers

func ptrInt(v int) *int { return &v }

func equalFormControlPr(a, b formControlPr) bool {
	if a.ObjectType != b.ObjectType || a.Checked != b.Checked ||
		a.FmlaLink != b.FmlaLink || a.FmlaRange != b.FmlaRange ||
		a.FmlaMacro != b.FmlaMacro || a.Sel != b.Sel || a.DropLines != b.DropLines ||
		a.SelType != b.SelType {
		return false
	}
	return eqIntPtr(a.Min, b.Min) && eqIntPtr(a.Max, b.Max) &&
		eqIntPtr(a.Val, b.Val) && eqIntPtr(a.Inc, b.Inc) && eqIntPtr(a.Page, b.Page)
}

func eqIntPtr(a, b *int) bool {
	if a == nil && b == nil {
		return true
	}
	if a == nil || b == nil {
		return false
	}
	return *a == *b
}
