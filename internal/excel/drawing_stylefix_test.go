package excel

import (
	"encoding/xml"
	"strings"
	"testing"
)

// testTheme は dk1=黒, lt1=白, accent1=青 のテーマを返す
func testTheme() *themeColors {
	colors := make([]string, 12)
	colors[0] = "#000000" // dk1
	colors[1] = "#FFFFFF" // lt1
	colors[2] = "#44546A" // dk2
	colors[3] = "#E7E6E6" // lt2
	colors[4] = "#5B9BD5" // accent1
	return &themeColors{colors: colors}
}

// TestThemeGetScheme は DrawingML 用 GetScheme がセル用 Get のスワップを受けないことを確認する
func TestThemeGetScheme(t *testing.T) {
	tc := testTheme()

	// セル用 Get は 0↔1, 2↔3 をスワップする
	if got := tc.Get(0); got != "#FFFFFF" {
		t.Errorf("Get(0) = %q, want #FFFFFF (lt1, スワップ後)", got)
	}
	if got := tc.Get(1); got != "#000000" {
		t.Errorf("Get(1) = %q, want #000000 (dk1, スワップ後)", got)
	}

	// DrawingML 用 GetScheme は自然順（スワップなし）
	if got := tc.GetScheme(0); got != "#000000" {
		t.Errorf("GetScheme(0) = %q, want #000000 (dk1)", got)
	}
	if got := tc.GetScheme(1); got != "#FFFFFF" {
		t.Errorf("GetScheme(1) = %q, want #FFFFFF (lt1)", got)
	}
	if got := tc.GetScheme(4); got != "#5B9BD5" {
		t.Errorf("GetScheme(4) = %q, want #5B9BD5 (accent1)", got)
	}
}

// parseShapeXML はテスト用に <xdr:sp> XML 断片を parseShape に通す
func parseShapeXML(t *testing.T, theme *themeColors, spXML string) ShapeInfo {
	t.Helper()
	p := newDrawingParser(drawingParserConfig{theme: theme, includeStyle: true})
	decoder := xml.NewDecoder(strings.NewReader(spXML))
	for {
		tok, err := decoder.Token()
		if err != nil {
			t.Fatalf("sp 開始要素が見つからない: %v", err)
		}
		if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "sp" {
			return p.parseShape(decoder, 0, "", nil, nil)
		}
	}
}

// TestParseShapeFillAndLine は枠線色（dk1→黒）と fillRef 由来の塗り色を取得できることを確認する
func TestParseShapeFillAndLine(t *testing.T) {
	// spPr に明示の ln(dk1)、塗りは fillRef(lt1) 経由
	const spXML = `<xdr:sp xmlns:xdr="x" xmlns:a="a">
		<xdr:spPr>
			<a:prstGeom prst="flowChartProcess"/>
			<a:ln w="9525"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:ln>
		</xdr:spPr>
		<xdr:style>
			<a:lnRef idx="2"><a:schemeClr val="dk1"/></a:lnRef>
			<a:fillRef idx="1"><a:schemeClr val="lt1"/></a:fillRef>
			<a:effectRef idx="0"><a:schemeClr val="dk1"/></a:effectRef>
			<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
		</xdr:style>
	</xdr:sp>`

	shape := parseShapeXML(t, testTheme(), spXML)

	if shape.Fill != "#FFFFFF" {
		t.Errorf("Fill = %q, want #FFFFFF (fillRef lt1 由来)", shape.Fill)
	}
	if shape.Line == nil || shape.Line.Color != "#000000" {
		t.Errorf("Line = %+v, want color #000000 (ln dk1 由来)", shape.Line)
	}
}

// TestParseShapeFillRefAccent は色付き fillRef（accent1）の塗りを取得できることを確認する
func TestParseShapeFillRefAccent(t *testing.T) {
	const spXML = `<xdr:sp xmlns:xdr="x" xmlns:a="a">
		<xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
		<xdr:style>
			<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
			<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
		</xdr:style>
	</xdr:sp>`

	shape := parseShapeXML(t, testTheme(), spXML)

	if shape.Fill != "#5B9BD5" {
		t.Errorf("Fill = %q, want #5B9BD5 (fillRef accent1 由来)", shape.Fill)
	}
	if shape.Line == nil || shape.Line.Color != "#5B9BD5" {
		t.Errorf("Line = %+v, want color #5B9BD5 (lnRef accent1 由来)", shape.Line)
	}
}

// TestParseShapeExplicitFillWins は spPr の明示塗りが fillRef より優先されることを確認する
func TestParseShapeExplicitFillWins(t *testing.T) {
	const spXML = `<xdr:sp xmlns:xdr="x" xmlns:a="a">
		<xdr:spPr>
			<a:prstGeom prst="rect"/>
			<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
		</xdr:spPr>
		<xdr:style>
			<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
		</xdr:style>
	</xdr:sp>`

	shape := parseShapeXML(t, testTheme(), spXML)

	if shape.Fill != "#FF0000" {
		t.Errorf("Fill = %q, want #FF0000 (spPr の明示塗り優先)", shape.Fill)
	}
}

// TestApplyDrawingColorOps は HSL 近似の色変換が DrawingML 定義どおり効くことを確認する
func TestApplyDrawingColorOps(t *testing.T) {
	tests := []struct {
		name string
		base string
		ops  []drawingColorOp
		want string
	}{
		{"ops なしは素通り", "#5B9BD5", nil, "#5B9BD5"},
		{"tint は白へ寄せる", "#000000", []drawingColorOp{{"tint", 0.5}}, "#808080"},
		{"shade は黒へ寄せる", "#FFFFFF", []drawingColorOp{{"shade", 0.5}}, "#808080"},
		{"satMod 0 で無彩色化", "#5B9BD5", []drawingColorOp{{"satMod", 0}}, "#989898"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := applyDrawingColorOps(tt.base, tt.ops); got != tt.want {
				t.Errorf("applyDrawingColorOps(%q, %v) = %q, want %q", tt.base, tt.ops, got, tt.want)
			}
		})
	}
}

// TestThemeFillStyleTransform は fmtScheme の解析と fillRef idx 別の色変換適用を確認する
func TestThemeFillStyleTransform(t *testing.T) {
	const themeXML = `<a:theme xmlns:a="a"><a:themeElements>
		<a:clrScheme><a:dk1><a:srgbClr val="000000"/></a:dk1></a:clrScheme>
		<a:fmtScheme>
			<a:fillStyleLst>
				<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
				<a:gradFill>
					<a:gsLst>
						<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/></a:schemeClr></a:gs>
						<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="60000"/></a:schemeClr></a:gs>
						<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="80000"/></a:schemeClr></a:gs>
					</a:gsLst>
				</a:gradFill>
			</a:fillStyleLst>
		</a:fmtScheme>
	</a:themeElements></a:theme>`

	tc := parseThemeColors([]byte(themeXML))

	base := "#000000"
	// idx=1: solidFill phClr（変換なし）→ 素通り
	if got := tc.ApplyFillStyle(1, base); got != base {
		t.Errorf("idx=1: got %q, want %q", got, base)
	}
	// idx=2: 中央ストップ（pos=50000, tint 60%）を代表 → l=0*0.6+0.4=0.4 → #666666
	if got := tc.ApplyFillStyle(2, base); got != "#666666" {
		t.Errorf("idx=2: got %q, want #666666 (中央ストップ tint 60%%)", got)
	}
	// 範囲外 idx は素通り
	if got := tc.ApplyFillStyle(9, base); got != base {
		t.Errorf("idx=9(範囲外): got %q, want %q", got, base)
	}
}

// TestParseShapeFillRefNoFill は fillRef idx="0"（noFill）で塗りが出力されないことを確認する
func TestParseShapeFillRefNoFill(t *testing.T) {
	const spXML = `<xdr:sp xmlns:xdr="x" xmlns:a="a">
		<xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
		<xdr:style>
			<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
		</xdr:style>
	</xdr:sp>`

	shape := parseShapeXML(t, testTheme(), spXML)

	if shape.Fill != "" {
		t.Errorf("Fill = %q, want \"\" (fillRef idx=0 は noFill)", shape.Fill)
	}
}

// TestParseShapeExplicitNoFill は spPr の明示 noFill が fillRef フォールバックを抑止することを確認する
func TestParseShapeExplicitNoFill(t *testing.T) {
	const spXML = `<xdr:sp xmlns:xdr="x" xmlns:a="a">
		<xdr:spPr>
			<a:prstGeom prst="rect"/>
			<a:noFill/>
		</xdr:spPr>
		<xdr:style>
			<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
		</xdr:style>
	</xdr:sp>`

	shape := parseShapeXML(t, testTheme(), spXML)

	if shape.Fill != "" {
		t.Errorf("Fill = %q, want \"\" (明示 noFill は fillRef より優先)", shape.Fill)
	}
}

// TestParseShapeLineNoFill は ln 内の明示 noFill が線なし（lnRef 抑止）になることを確認する
func TestParseShapeLineNoFill(t *testing.T) {
	const spXML = `<xdr:sp xmlns:xdr="x" xmlns:a="a">
		<xdr:spPr>
			<a:prstGeom prst="rect"/>
			<a:ln w="9525"><a:noFill/></a:ln>
		</xdr:spPr>
		<xdr:style>
			<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
		</xdr:style>
	</xdr:sp>`

	shape := parseShapeXML(t, testTheme(), spXML)

	if shape.Line != nil {
		t.Errorf("Line = %+v, want nil (ln 内 noFill は線なし)", shape.Line)
	}
}
