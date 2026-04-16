package excel

import (
	"encoding/xml"
	"io"
	"log"
	"sort"
	"strconv"
	"strings"
)

// controlTypeMap は objectType（ctrlProp / VML ClientData の値）を内部 type に対応づける。
// 非対応タイプ（EditBox, Dialog, Note）はマップに含めない。
var controlTypeMap = map[string]string{
	"CheckBox": ShapeTypeCheckbox,
	"Checkbox": ShapeTypeCheckbox, // VML 側は Checkbox 表記で出てくるため
	"Radio":    ShapeTypeRadio,
	"Option":   ShapeTypeRadio, // VML 側の別名
	"Drop":     ShapeTypeDrop,
	"List":     ShapeTypeList,
	"Spin":     ShapeTypeSpin,
	"Scroll":   ShapeTypeScroll,
	"Button":   ShapeTypeButton,
	"GBox":     ShapeTypeGroupBox,
	"Label":    ShapeTypeLabel,
}

// formControlPr は xl/ctrlProps/ctrlProp*.xml の <formControlPr> から読み取るプロパティ
type formControlPr struct {
	ObjectType string
	Checked    string // "Checked" 以外（空や "Unchecked"）は未チェック扱い
	FmlaLink   string
	FmlaRange  string
	FmlaMacro  string
	Sel        int
	DropLines  int
	SelType    string
	Min        *int
	Max        *int
	Val        *int
	Inc        *int
	Page       *int
}

// sheetControl は sheet.xml の <controls><control> 1件分
type sheetControl struct {
	ShapeID int
	RelID   string
	Name    string
	From    anchorPos
	To      anchorPos
}

// vmlShapeInfo は VML 側から抽出する表示情報（shapeId をキーにマージ）
type vmlShapeInfo struct {
	ZIndex   int
	HasZ     bool
	Text     string
	HasShape bool
}

// loadFormControls はシートのフォームコントロールを読み込み、ShapeInfo のリストと次の z-order / ID を返す。
// 対応しない種別（EditBox, Dialog 等）は静かにスキップする。
func loadFormControls(zi *zipIndex, sheetXMLPath string, sheetMeta *SheetMeta, startZ, startID int) (shapes []ShapeInfo, nextZ, nextID int) {
	nextZ = startZ
	nextID = startID

	controls := parseSheetControls(zi, sheetXMLPath)
	if len(controls) == 0 {
		return nil, nextZ, nextID
	}

	// sheet.xml.rels から rId → 絶対パスのマップを構築
	rels := loadSheetRelsAll(zi, sheetXMLPath)
	relsMap := make(map[string]string, len(rels))
	for _, r := range rels {
		relsMap[r.ID] = resolveRelTarget(sheetXMLPath, r.Target)
	}

	// VML 側から shapeId → ZIndex / Text を収集（複数の vmlDrawing にまたがる可能性あり）
	vmlMap := make(map[int]vmlShapeInfo)
	for _, r := range rels {
		if !strings.Contains(strings.ToLower(r.Type), "vmldrawing") {
			continue
		}
		vmlPath := resolveRelTarget(sheetXMLPath, r.Target)
		parseVMLShapes(zi, vmlPath, vmlMap)
	}

	// VML の z-index 順でソート（未取得はシート定義順を維持）
	sort.SliceStable(controls, func(i, j int) bool {
		a, b := vmlMap[controls[i].ShapeID], vmlMap[controls[j].ShapeID]
		if a.HasZ && b.HasZ {
			return a.ZIndex < b.ZIndex
		}
		return !a.HasZ && b.HasZ
	})

	var posCalc *posCalculator
	if sheetMeta != nil {
		posCalc = &posCalculator{meta: sheetMeta}
	}

	for _, c := range controls {
		ctrlPath, ok := relsMap[c.RelID]
		if !ok {
			continue
		}
		props := parseCtrlProp(zi, ctrlPath)
		if props == nil {
			continue
		}
		t, ok := controlTypeMap[props.ObjectType]
		if !ok {
			continue
		}
		vml := vmlMap[c.ShapeID]
		shape := ShapeInfo{
			ID:   nextID,
			Type: t,
			Name: c.Name,
			Z:    nextZ,
			Cell: buildCellRangeFromAnchor(c.From, c.To),
		}
		if posCalc != nil {
			shape.Pos = buildPosFromAnchor(posCalc, c.From, c.To)
		}
		applyControlProps(&shape, t, props, vml)
		shapes = append(shapes, shape)
		nextZ++
		nextID++
	}

	return shapes, nextZ, nextID
}

// applyControlProps はコントロール種別に応じてフォームコントロール用フィールドを設定する
func applyControlProps(shape *ShapeInfo, t string, props *formControlPr, vml vmlShapeInfo) {
	shape.Text = vml.Text
	switch t {
	case ShapeTypeCheckbox, ShapeTypeRadio:
		checked := props.Checked == "Checked"
		shape.Checked = &checked
		shape.LinkedCell = props.FmlaLink
	case ShapeTypeDrop:
		shape.LinkedCell = props.FmlaLink
		shape.ListRange = props.FmlaRange
		shape.SelectedIndex = props.Sel
		shape.DropLines = props.DropLines
	case ShapeTypeList:
		shape.LinkedCell = props.FmlaLink
		shape.ListRange = props.FmlaRange
		shape.SelectedIndex = props.Sel
		if props.SelType != "" {
			shape.SelType = strings.ToLower(props.SelType)
		}
	case ShapeTypeSpin:
		shape.LinkedCell = props.FmlaLink
		shape.Min = props.Min
		shape.Max = props.Max
		shape.Val = props.Val
		shape.Inc = props.Inc
	case ShapeTypeScroll:
		shape.LinkedCell = props.FmlaLink
		shape.Min = props.Min
		shape.Max = props.Max
		shape.Val = props.Val
		shape.Inc = props.Inc
		shape.Page = props.Page
	case ShapeTypeButton:
		shape.Macro = props.FmlaMacro
	}
}

// buildCellRangeFromAnchor は from/to のセル位置から "A1:B2" または "A1" の形式の範囲文字列を返す
func buildCellRangeFromAnchor(from, to anchorPos) string {
	fromRef := CellRef(from.col+1, from.row+1)
	toRef := CellRef(to.col+1, to.row+1)
	if fromRef == toRef {
		return fromRef
	}
	return fromRef + ":" + toRef
}

// buildPosFromAnchor は from/to アンカーから Position を算出する
func buildPosFromAnchor(pc *posCalculator, from, to anchorPos) *Position {
	x1 := pc.calcX(from.col, from.colOff)
	y1 := pc.calcY(from.row, from.rowOff)
	x2 := pc.calcX(to.col, to.colOff)
	y2 := pc.calcY(to.row, to.rowOff)
	return &Position{X: x1, Y: y1, W: x2 - x1, H: y2 - y1}
}

// parseSheetControls はシートXMLの <controls> 配下から控除項目の一覧を取り出す。
// <mc:AlternateContent> の <mc:Choice> を採用し、<mc:Fallback> は無視する。
func parseSheetControls(zi *zipIndex, sheetXMLPath string) []sheetControl {
	entry := zi.lookup(sheetXMLPath)
	if entry == nil {
		return nil
	}
	var controls []sheetControl
	_ = withZipXML(entry, func(decoder *xml.Decoder) error {
		controls = parseSheetControlsSAX(decoder)
		return nil
	})
	return controls
}

// parseSheetControlsSAX は SAX パーサーで <controls>/<control> を拾う
func parseSheetControlsSAX(decoder *xml.Decoder) []sheetControl {
	var (
		controls      []sheetControl
		inControls    bool
		fallbackDepth int // >0 の間は mc:Fallback 内
	)

	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Printf("[WARN] parseSheetControlsSAX: %v", err)
			return controls
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if fallbackDepth > 0 {
				fallbackDepth++
				continue
			}
			switch t.Name.Local {
			case "controls":
				inControls = true
			case "Fallback":
				if inControls {
					fallbackDepth = 1
				}
			case "control":
				if inControls {
					c := parseSheetControlAttrs(t)
					if readSheetControlAnchor(decoder, &c) {
						controls = append(controls, c)
					}
				}
			}
		case xml.EndElement:
			if fallbackDepth > 0 {
				fallbackDepth--
				continue
			}
			if t.Name.Local == "controls" {
				inControls = false
			}
		}
	}
	return controls
}

// parseSheetControlAttrs は <control> の属性から shapeId, r:id, name を読む
func parseSheetControlAttrs(t xml.StartElement) sheetControl {
	var c sheetControl
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "shapeId":
			c.ShapeID = safeAtoi(attr.Value)
		case "name":
			c.Name = attr.Value
		case "id":
			if strings.HasSuffix(attr.Name.Space, "relationships") {
				c.RelID = attr.Value
			}
		}
	}
	return c
}

// readSheetControlAnchor は <control> の末尾まで読み進め、<anchor> の from/to を取得する。
// from/to が両方取れた場合のみ true を返す。
func readSheetControlAnchor(decoder *xml.Decoder, c *sheetControl) bool {
	depth := 1
	gotFrom, gotTo := false, false
	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			log.Printf("[WARN] readSheetControlAnchor: %v", err)
			return gotFrom && gotTo
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "from":
				c.From = parseXdrAnchorPos(decoder)
				gotFrom = true
			case "to":
				c.To = parseXdrAnchorPos(decoder)
				gotTo = true
			default:
				depth++
			}
		case xml.EndElement:
			depth--
		}
	}
	return gotFrom && gotTo
}

// parseXdrAnchorPos は <from> または <to> の子として col/colOff/row/rowOff を読み取る。
// 呼び出し時点で <from>/<to> の StartElement は既に消費済み。EndElement まで消費する。
func parseXdrAnchorPos(decoder *xml.Decoder) anchorPos {
	var pos anchorPos
	var field string
	var buf strings.Builder
	depth := 1
	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			log.Printf("[WARN] parseXdrAnchorPos: %v", err)
			return pos
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "col", "colOff", "row", "rowOff":
				field = t.Name.Local
				buf.Reset()
			}
		case xml.EndElement:
			depth--
			switch t.Name.Local {
			case "col":
				pos.col = safeAtoi(buf.String())
				field = ""
			case "colOff":
				pos.colOff = safeAtoi(buf.String())
				field = ""
			case "row":
				pos.row = safeAtoi(buf.String())
				field = ""
			case "rowOff":
				pos.rowOff = safeAtoi(buf.String())
				field = ""
			}
		case xml.CharData:
			if field != "" {
				buf.Write(t)
			}
		}
	}
	return pos
}

// parseCtrlProp は xl/ctrlProps/ctrlProp*.xml をパースする
func parseCtrlProp(zi *zipIndex, path string) *formControlPr {
	entry := zi.lookup(path)
	if entry == nil {
		return nil
	}
	var result *formControlPr
	_ = withZipXML(entry, func(decoder *xml.Decoder) error {
		for {
			tok, err := decoder.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				return err
			}
			if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "formControlPr" {
				result = parseFormControlPrAttrs(se)
				return nil
			}
		}
		return nil
	})
	return result
}

// parseFormControlPrAttrs は <formControlPr> の属性を構造体に取り込む
func parseFormControlPrAttrs(t xml.StartElement) *formControlPr {
	p := &formControlPr{}
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "objectType":
			p.ObjectType = attr.Value
		case "checked":
			p.Checked = attr.Value
		case "fmlaLink":
			p.FmlaLink = attr.Value
		case "fmlaRange":
			p.FmlaRange = attr.Value
		case "fmlaMacro":
			p.FmlaMacro = attr.Value
		case "sel":
			p.Sel = safeAtoi(attr.Value)
		case "dropLines":
			p.DropLines = safeAtoi(attr.Value)
		case "selType":
			p.SelType = attr.Value
		case "min":
			v := safeAtoi(attr.Value)
			p.Min = &v
		case "max":
			v := safeAtoi(attr.Value)
			p.Max = &v
		case "val":
			v := safeAtoi(attr.Value)
			p.Val = &v
		case "inc":
			v := safeAtoi(attr.Value)
			p.Inc = &v
		case "page":
			v := safeAtoi(attr.Value)
			p.Page = &v
		}
	}
	return p
}

// parseVMLShapes は vmlDrawing*.vml をスキャンし、shapeId → {ZIndex, Text} を vmlMap に追記する
func parseVMLShapes(zi *zipIndex, path string, vmlMap map[int]vmlShapeInfo) {
	entry := zi.lookup(path)
	if entry == nil {
		return
	}
	_ = withZipXML(entry, func(decoder *xml.Decoder) error {
		parseVMLShapesSAX(decoder, vmlMap)
		return nil
	})
}

// parseVMLShapesSAX は VML を SAX パースし、v:shape ごとに shapeId をキーに情報を追加する
func parseVMLShapesSAX(decoder *xml.Decoder, vmlMap map[int]vmlShapeInfo) {
	var (
		inShape    bool
		shapeID    int
		zIndex     int
		hasZ       bool
		inTextbox  bool
		divDepth   int // >0 の間は <div> 内（入れ子を数える）
		textDivs   []string
		currentDiv strings.Builder
	)

	flushDiv := func() {
		s := strings.TrimSpace(currentDiv.String())
		if s != "" {
			textDivs = append(textDivs, s)
		}
		currentDiv.Reset()
	}

	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "shape":
				inShape = true
				shapeID = 0
				zIndex = 0
				hasZ = false
				textDivs = nil
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "id":
						shapeID = parseVMLShapeID(attr.Value)
					case "style":
						if z, ok := parseVMLZIndex(attr.Value); ok {
							zIndex = z
							hasZ = true
						}
					}
				}
			case "textbox":
				if inShape {
					inTextbox = true
					divDepth = 0
					currentDiv.Reset()
				}
			case "div":
				if inTextbox {
					// ネストされた div は親 div の一部として扱う（深さだけ数える）
					if divDepth == 0 {
						flushDiv()
					}
					divDepth++
				}
			case "br":
				if inTextbox && divDepth > 0 {
					currentDiv.WriteString("\n")
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "shape":
				if inShape && shapeID > 0 {
					flushDiv()
					text := strings.Join(textDivs, "\n")
					info := vmlMap[shapeID]
					info.HasShape = true
					if hasZ {
						info.ZIndex = zIndex
						info.HasZ = true
					}
					if text != "" {
						info.Text = text
					}
					vmlMap[shapeID] = info
				}
				inShape = false
				inTextbox = false
				divDepth = 0
			case "textbox":
				if inTextbox {
					flushDiv()
					inTextbox = false
					divDepth = 0
				}
			case "div":
				if inTextbox && divDepth > 0 {
					divDepth--
					if divDepth == 0 {
						flushDiv()
					}
				}
			}
		case xml.CharData:
			// div 内のテキストだけ拾う（textbox 直下の空白を除外するため）
			if inTextbox && divDepth > 0 {
				currentDiv.Write(t)
			}
		}
	}
}

// parseVMLShapeID は VML の shape id 属性（例: "_x0000_s73729"）から数値部分を取り出す
func parseVMLShapeID(s string) int {
	// 末尾の数値部分を切り出す
	i := len(s)
	for i > 0 && s[i-1] >= '0' && s[i-1] <= '9' {
		i--
	}
	if i == len(s) {
		return 0
	}
	n, _ := strconv.Atoi(s[i:])
	return n
}

// parseVMLZIndex は VML shape の style 属性から z-index: の値を取り出す
func parseVMLZIndex(style string) (int, bool) {
	lower := strings.ToLower(style)
	key := "z-index:"
	idx := strings.Index(lower, key)
	if idx < 0 {
		return 0, false
	}
	rest := style[idx+len(key):]
	// 次のセミコロンまで
	end := strings.IndexAny(rest, ";")
	if end >= 0 {
		rest = rest[:end]
	}
	rest = strings.TrimSpace(rest)
	n, err := strconv.Atoi(rest)
	if err != nil {
		return 0, false
	}
	return n, true
}

