package excel

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"log"
	"math"
	"strconv"
	"strings"
)

// anchorPos は anchor の from/to 内の位置情報
type anchorPos struct {
	col    int
	colOff int // EMU
	row    int
	rowOff int // EMU
}

// posCalculator はアンカー位置からポイント座標を算出する
type posCalculator struct {
	meta *SheetMeta
}

// calcX は列+オフセットからX座標（ポイント）を算出する（col は 0 始まり、off は EMU）
func (pc *posCalculator) calcX(col, off int) int {
	var x float64
	for c := 1; c <= col; c++ {
		x += pc.meta.ColWidthPt(c)
	}
	x += float64(off) / emuPerPt
	return int(math.Round(x))
}

// calcY は行+オフセットからY座標（ポイント）を算出する（row は 0 始まり、off は EMU）
func (pc *posCalculator) calcY(row, off int) int {
	var y float64
	for r := 1; r <= row; r++ {
		y += pc.meta.RowHeightPt(r)
	}
	y += float64(off) / emuPerPt
	return int(math.Round(y))
}

// drawingParser は drawing XML の SAX パーサー
type drawingParser struct {
	theme        *themeColors
	includeStyle bool

	// 結果
	shapes    []ShapeInfo
	connCount int
	picCount  int

	// Excel図形ID → 連番IDのマッピング
	excelIDMap map[int]int
	nextID     int

	// アンカーの状態
	topZ int // トップレベルの z-order カウンタ

	// コネクタの接続先（後処理用）
	connRefs []connRef

	// 座標計算
	posCalc *posCalculator

	// 画像対応
	drawingPath string
	drawingRels map[string]xmlRelationship
}

type connRef struct {
	shapeIndex int // shapes スライス内のインデックス
	startID    int // Excel の接続元ID
	endID      int // Excel の接続先ID
	startIdx   int // 接続元の接続ポイントインデックス
	endIdx     int // 接続先の接続ポイントインデックス
	hasStart   bool
	hasEnd     bool
}

// グループのコンテキスト
type groupContext struct {
	seqID    int // グループの連番ID
	childZ   int // グループ内の z-order カウンタ
	children []int
}

// drawingParserConfig は drawingParser の初期化パラメータ
type drawingParserConfig struct {
	theme        *themeColors
	includeStyle bool
	drawingPath  string
	drawingRels  map[string]xmlRelationship
	sheetMeta    *SheetMeta
}

func newDrawingParser(cfg drawingParserConfig) *drawingParser {
	p := &drawingParser{
		theme:        cfg.theme,
		includeStyle: cfg.includeStyle,
		excelIDMap:   make(map[int]int),
		nextID:       1,
		drawingPath:  cfg.drawingPath,
		drawingRels:  cfg.drawingRels,
	}
	if cfg.sheetMeta != nil {
		p.posCalc = &posCalculator{meta: cfg.sheetMeta}
	}
	return p
}

func parseDrawingXML(entry *zip.File, cfg drawingParserConfig) (*DrawingResult, error) {
	rc, err := entry.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	p := newDrawingParser(cfg)

	if err := p.parse(rc); err != nil {
		return nil, err
	}

	// コネクタの接続先を解決
	p.resolveConnectors()

	return &DrawingResult{
		Meta: ShapesMeta{
			Meta:           true,
			ShapeCount:     len(p.shapes),
			ConnectorCount: p.connCount,
		},
		Shapes: p.shapes,
	}, nil
}

func (p *drawingParser) parse(r io.Reader) error {
	decoder := xml.NewDecoder(r)

	var (
		// アンカー状態
		inAnchor   bool
		anchorType string // "two", "one", "abs"
		anchorFrom anchorPos
		anchorTo   anchorPos
		hasTo      bool

		// oneCellAnchor / absoluteAnchor 用の ext (EMU)
		anchorExtCX, anchorExtCY int
		// absoluteAnchor 用の pos (EMU)
		anchorAbsX, anchorAbsY int

		// スキップ状態（pic, graphicFrame）
		skipDepth int

		// グループスタック
		groupStack []groupContext
	)

	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if skipDepth > 0 {
				skipDepth++
				continue
			}

			switch t.Name.Local {
			case "twoCellAnchor":
				inAnchor = true
				anchorType = "two"
				anchorFrom = anchorPos{}
				anchorTo = anchorPos{}
				hasTo = false
				anchorExtCX, anchorExtCY = 0, 0

			case "oneCellAnchor":
				inAnchor = true
				anchorType = "one"
				anchorFrom = anchorPos{}
				hasTo = false
				anchorExtCX, anchorExtCY = 0, 0

			case "absoluteAnchor":
				inAnchor = true
				anchorType = "abs"
				hasTo = false
				anchorExtCX, anchorExtCY = 0, 0
				anchorAbsX, anchorAbsY = 0, 0

			case "from":
				if inAnchor && len(groupStack) == 0 {
					anchorFrom = parseAnchorPos(decoder)
				} else {
					skipElement(decoder)
				}

			case "to":
				if inAnchor && len(groupStack) == 0 && anchorType == "two" {
					anchorTo = parseAnchorPos(decoder)
					hasTo = true
				} else {
					skipElement(decoder)
				}

			case "ext":
				// アンカーレベルの ext（oneCellAnchor / absoluteAnchor）
				if inAnchor && len(groupStack) == 0 && anchorType != "two" {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "cx":
							anchorExtCX, _ = strconv.Atoi(attr.Value)
						case "cy":
							anchorExtCY, _ = strconv.Atoi(attr.Value)
						}
					}
				}

			case "pos":
				// absoluteAnchor の位置
				if inAnchor && anchorType == "abs" && len(groupStack) == 0 {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "x":
							anchorAbsX, _ = strconv.Atoi(attr.Value)
						case "y":
							anchorAbsY, _ = strconv.Atoi(attr.Value)
						}
					}
				}

			case "sp":
				if !inAnchor {
					continue
				}
				z, cell, pos := p.prepareShapeContext(anchorType, anchorFrom, anchorTo, hasTo, anchorExtCX, anchorExtCY, anchorAbsX, anchorAbsY, groupStack)
				shape := p.parseShape(decoder, z, cell, pos, groupStack)
				if shape.Hidden {
					continue
				}
				p.incrementZOrder(groupStack)
				p.addShape(shape, groupStack)

			case "cxnSp":
				if !inAnchor {
					continue
				}
				z, cell, pos := p.prepareShapeContext(anchorType, anchorFrom, anchorTo, hasTo, anchorExtCX, anchorExtCY, anchorAbsX, anchorAbsY, groupStack)
				shape := p.parseConnector(decoder, z, cell, pos, groupStack)
				if shape.Hidden {
					continue
				}
				p.incrementZOrder(groupStack)
				p.connCount++
				p.addShape(shape, groupStack)

			case "grpSp":
				if !inAnchor {
					continue
				}
				z, cell, pos := p.prepareShapeContext(anchorType, anchorFrom, anchorTo, hasTo, anchorExtCX, anchorExtCY, anchorAbsX, anchorAbsY, groupStack)
				grpShape := p.startGroup(decoder, z, cell, pos, groupStack)
				if grpShape.Hidden {
					// 非表示グループはスキップ。子要素は本来このグループの後に続くが、
					// grpSp の EndElement まで読み飛ばさないと整合性が崩れる。
					skipDepth = 1
					continue
				}
				p.incrementZOrder(groupStack)
				p.addShape(grpShape, groupStack)
				groupStack = append(groupStack, groupContext{
					seqID: grpShape.ID,
				})

			case "pic":
				if !inAnchor {
					continue
				}
				z, cell, pos := p.prepareShapeContext(anchorType, anchorFrom, anchorTo, hasTo, anchorExtCX, anchorExtCY, anchorAbsX, anchorAbsY, groupStack)
				shape := p.parsePicture(decoder, z, cell, pos, groupStack)
				if shape.Hidden {
					continue
				}
				p.incrementZOrder(groupStack)
				p.picCount++
				p.addShape(shape, groupStack)

			case "graphicFrame":
				// スキップ対象
				skipDepth = 1

			default:
				// アンカー内のその他の要素はスキップ
			}

		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
				continue
			}

			switch t.Name.Local {
			case "twoCellAnchor", "oneCellAnchor", "absoluteAnchor":
				inAnchor = false
				p.topZ++

			case "grpSp":
				p.closeGroup(&groupStack)
			}
		}
	}

	return nil
}

// prepareShapeContext は図形パース前の共通処理（z-order, セル参照, 位置情報の算出）
func (p *drawingParser) prepareShapeContext(
	anchorType string, anchorFrom, anchorTo anchorPos, hasTo bool,
	extCX, extCY, absX, absY int, groupStack []groupContext,
) (z int, cell string, pos *Position) {
	z = p.currentZOrder(groupStack)
	cell = p.buildCell(anchorType, anchorFrom.col, anchorFrom.row, anchorTo.col, anchorTo.row, hasTo)
	if len(groupStack) == 0 {
		pos = p.buildPos(anchorType, anchorFrom, anchorTo, hasTo, extCX, extCY, absX, absY)
	}
	return z, cell, pos
}

// addShape は図形を shapes に追加し、グループの children を更新する
func (p *drawingParser) addShape(shape ShapeInfo, groupStack []groupContext) {
	p.shapes = append(p.shapes, shape)
	if len(groupStack) > 0 {
		groupStack[len(groupStack)-1].children = append(groupStack[len(groupStack)-1].children, shape.ID)
	}
}

// closeGroup はグループスタックの先頭を取り出し、children を設定する
func (p *drawingParser) closeGroup(groupStack *[]groupContext) {
	if len(*groupStack) == 0 {
		return
	}
	top := (*groupStack)[len(*groupStack)-1]
	*groupStack = (*groupStack)[:len(*groupStack)-1]
	for i := range p.shapes {
		if p.shapes[i].ID == top.seqID {
			p.shapes[i].Children = top.children
			break
		}
	}
}

// cellRangeRef は from/to のセル座標（0始まり）から "A1:B2" または "A1" の形式の範囲文字列を返す
func cellRangeRef(fromCol, fromRow, toCol, toRow int) string {
	from := CellRef(fromCol+1, fromRow+1)
	to := CellRef(toCol+1, toRow+1)
	if from == to {
		return from
	}
	return from + ":" + to
}

func (p *drawingParser) currentZOrder(groupStack []groupContext) int {
	if len(groupStack) > 0 {
		return groupStack[len(groupStack)-1].childZ
	}
	return p.topZ
}

func (p *drawingParser) incrementZOrder(groupStack []groupContext) {
	if len(groupStack) > 0 {
		groupStack[len(groupStack)-1].childZ++
	}
}

func (p *drawingParser) buildCell(anchorType string, fromCol, fromRow, toCol, toRow int, hasTo bool) string {
	switch anchorType {
	case "two":
		if hasTo {
			return cellRangeRef(fromCol, fromRow, toCol, toRow)
		}
		return CellRef(fromCol+1, fromRow+1)
	case "one":
		return CellRef(fromCol+1, fromRow+1)
	default:
		return ""
	}
}

// parseAnchorPos は <from> / <to> / <xdr:from> / <xdr:to> 内の col, colOff, row, rowOff を読む。
// 呼び出し時点で <from>/<to> の StartElement は消費済みで、EndElement まで消費する。
func parseAnchorPos(decoder *xml.Decoder) anchorPos {
	depth := 1
	var pos anchorPos
	var field string // "col", "colOff", "row", "rowOff"
	var buf strings.Builder

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			log.Printf("[WARN] parseAnchorPos: XMLトークン読み取りに失敗: %v", err)
			return anchorPos{}
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

// twoAnchorPos は from/to アンカーからポイント座標の Position を算出する
func twoAnchorPos(pc *posCalculator, from, to anchorPos) *Position {
	x1 := pc.calcX(from.col, from.colOff)
	y1 := pc.calcY(from.row, from.rowOff)
	x2 := pc.calcX(to.col, to.colOff)
	y2 := pc.calcY(to.row, to.rowOff)
	return &Position{X: x1, Y: y1, W: x2 - x1, H: y2 - y1}
}

// buildPos はアンカー情報からポイント座標を算出する
func (p *drawingParser) buildPos(anchorType string, from, to anchorPos, hasTo bool, extCX, extCY, absX, absY int) *Position {
	if p.posCalc == nil {
		return nil
	}
	switch anchorType {
	case "two":
		if !hasTo {
			return nil
		}
		return twoAnchorPos(p.posCalc, from, to)
	case "one":
		x := p.posCalc.calcX(from.col, from.colOff)
		y := p.posCalc.calcY(from.row, from.rowOff)
		w := int(math.Round(float64(extCX) / emuPerPt))
		h := int(math.Round(float64(extCY) / emuPerPt))
		return &Position{X: x, Y: y, W: w, H: h}
	case "abs":
		x := int(math.Round(float64(absX) / emuPerPt))
		y := int(math.Round(float64(absY) / emuPerPt))
		w := int(math.Round(float64(extCX) / emuPerPt))
		h := int(math.Round(float64(extCY) / emuPerPt))
		return &Position{X: x, Y: y, W: w, H: h}
	}
	return nil
}

// connectorEndpoints はコネクタの pos と flip から始点・終点を算出する
func connectorEndpoints(pos *Position, flip string) (*Point, *Point) {
	x1, y1 := pos.X, pos.Y
	x2, y2 := pos.X+pos.W, pos.Y+pos.H
	switch flip {
	case "h":
		x1, x2 = x2, x1
	case "v":
		y1, y2 = y2, y1
	case "hv":
		x1, x2 = x2, x1
		y1, y2 = y2, y1
	}
	return &Point{X: x1, Y: y1}, &Point{X: x2, Y: y2}
}

// callout 形状のデフォルト adj 値（adj1=x, adj2=y、1/100000 単位）
var calloutDefaults = map[string][2]int{
	"wedgeRectCallout":      {-20833, 62500},
	"wedgeRoundRectCallout": {-20833, 62500},
	"wedgeEllipseCallout":   {-20833, 62500},
}

// borderCallout 形状のデフォルト adj 値（pointer tip: adj4=x, adj3=y）
var borderCalloutDefaults = map[string][2]int{
	"borderCallout1": {-8333, 112963},
	"accentCallout1": {-8333, 112963},
	"callout1":       {-8333, 112963},
}

// calcCalloutTarget は吹き出し形状のポインタ先ポイント座標を算出する
func calcCalloutTarget(pos *Position, shapeType string, adjs map[string]int) *Point {
	// wedge 系: adj1=x, adj2=y
	if defaults, ok := calloutDefaults[shapeType]; ok {
		adjX, adjY := defaults[0], defaults[1]
		if v, ok := adjs["adj1"]; ok {
			adjX = v
		}
		if v, ok := adjs["adj2"]; ok {
			adjY = v
		}
		x := pos.X + int(math.Round(float64(adjX)*float64(pos.W)/100000))
		y := pos.Y + int(math.Round(float64(adjY)*float64(pos.H)/100000))
		return &Point{X: x, Y: y}
	}
	// border/accent/callout 系: adj4=x, adj3=y
	if defaults, ok := borderCalloutDefaults[shapeType]; ok {
		adjX, adjY := defaults[0], defaults[1]
		if v, ok := adjs["adj4"]; ok {
			adjX = v
		}
		if v, ok := adjs["adj3"]; ok {
			adjY = v
		}
		x := pos.X + int(math.Round(float64(adjX)*float64(pos.W)/100000))
		y := pos.Y + int(math.Round(float64(adjY)*float64(pos.H)/100000))
		return &Point{X: x, Y: y}
	}
	return nil
}

// parseXfrm は xfrm の属性から回転と反転を取得する
func parseXfrm(t xml.StartElement) (rotation float64, flip string) {
	var flipH, flipV bool
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "rot":
			rot, _ := strconv.Atoi(attr.Value)
			rotation = math.Round(float64(rot)/drawingMLRotUnit*100) / 100
		case "flipH":
			flipH = attr.Value == "1"
		case "flipV":
			flipV = attr.Value == "1"
		}
	}
	if flipH && flipV {
		flip = "hv"
	} else if flipH {
		flip = "h"
	} else if flipV {
		flip = "v"
	}
	return rotation, flip
}
