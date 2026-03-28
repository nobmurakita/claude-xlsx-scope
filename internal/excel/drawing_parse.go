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

	// 画像対応
	drawingPath string
	drawingRels map[string]xmlRelationship
	zipEntries  map[string]*zip.File
	extractDir  string // 空なら画像スキップ
}

type connRef struct {
	shapeIndex int // shapes スライス内のインデックス
	startID    int // Excel の接続元ID
	endID      int // Excel の接続先ID
	hasStart   bool
	hasEnd     bool
}

// グループのコンテキスト
type groupContext struct {
	seqID    int // グループの連番ID
	childZ   int // グループ内の z-order カウンタ
	children []int
}

func newDrawingParser(theme *themeColors, includeStyle bool, drawingPath string, drawingRels map[string]xmlRelationship, zipEntries map[string]*zip.File, extractDir string) *drawingParser {
	return &drawingParser{
		theme:        theme,
		includeStyle: includeStyle,
		excelIDMap:   make(map[int]int),
		nextID:       1,
		drawingPath:  drawingPath,
		drawingRels:  drawingRels,
		zipEntries:   zipEntries,
		extractDir:   extractDir,
	}
}

func parseDrawingXML(entry *zip.File, theme *themeColors, includeStyle bool, drawingPath string, drawingRels map[string]xmlRelationship, zipEntries map[string]*zip.File, extractDir string) (*DrawingResult, error) {
	rc, err := entry.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	p := newDrawingParser(theme, includeStyle, drawingPath, drawingRels, zipEntries, extractDir)

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
		inAnchor     bool
		anchorType   string // "two", "one", "abs"
		anchorFromCol, anchorFromRow int
		anchorToCol, anchorToRow     int
		hasTo        bool

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
				anchorFromCol, anchorFromRow = 0, 0
				anchorToCol, anchorToRow = 0, 0
				hasTo = false

			case "oneCellAnchor":
				inAnchor = true
				anchorType = "one"
				anchorFromCol, anchorFromRow = 0, 0
				hasTo = false

			case "absoluteAnchor":
				inAnchor = true
				anchorType = "abs"
				hasTo = false

			case "from":
				if inAnchor && len(groupStack) == 0 {
					anchorFromCol, anchorFromRow = p.parseAnchorPos(decoder)
				} else {
					skipElement(decoder)
				}

			case "to":
				if inAnchor && len(groupStack) == 0 && anchorType == "two" {
					anchorToCol, anchorToRow = p.parseAnchorPos(decoder)
					hasTo = true
				} else {
					skipElement(decoder)
				}

			case "sp":
				if !inAnchor {
					continue
				}
				z := p.currentZOrder(groupStack)
				cell := p.buildCell(anchorType, anchorFromCol, anchorFromRow, anchorToCol, anchorToRow, hasTo)
				shape := p.parseShape(decoder, z, cell, groupStack)
				p.incrementZOrder(groupStack)
				p.addShape(shape, groupStack)

			case "cxnSp":
				if !inAnchor {
					continue
				}
				z := p.currentZOrder(groupStack)
				cell := p.buildCell(anchorType, anchorFromCol, anchorFromRow, anchorToCol, anchorToRow, hasTo)
				shape := p.parseConnector(decoder, z, cell, groupStack)
				p.incrementZOrder(groupStack)
				p.connCount++
				p.addShape(shape, groupStack)

			case "grpSp":
				if !inAnchor {
					continue
				}
				z := p.currentZOrder(groupStack)
				cell := p.buildCell(anchorType, anchorFromCol, anchorFromRow, anchorToCol, anchorToRow, hasTo)
				grpShape := p.startGroup(decoder, z, cell, groupStack)
				p.incrementZOrder(groupStack)
				p.addShape(grpShape, groupStack)
				groupStack = append(groupStack, groupContext{
					seqID: grpShape.ID,
				})

			case "pic":
				if !inAnchor {
					continue
				}
				if p.extractDir == "" {
					skipDepth = 1
					continue
				}
				z := p.currentZOrder(groupStack)
				cell := p.buildCell(anchorType, anchorFromCol, anchorFromRow, anchorToCol, anchorToRow, hasTo)
				shape := p.parsePicture(decoder, z, cell, groupStack)
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
			from := CellRef(fromCol+1, fromRow+1)
			to := CellRef(toCol+1, toRow+1)
			if from == to {
				return from
			}
			return from + ":" + to
		}
		return CellRef(fromCol+1, fromRow+1)
	case "one":
		return CellRef(fromCol+1, fromRow+1)
	default:
		return ""
	}
}

// parseAnchorPos は <from> または <to> 内の col, row を読む
func (p *drawingParser) parseAnchorPos(decoder *xml.Decoder) (col, row int) {
	depth := 1
	var inCol, inRow bool
	var buf strings.Builder

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			log.Printf("[WARN] parseAnchorPos: XMLトークン読み取りに失敗: %v", err)
			return 0, 0
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "col":
				inCol = true
				buf.Reset()
			case "row":
				inRow = true
				buf.Reset()
			}
		case xml.EndElement:
			depth--
			if t.Name.Local == "col" {
				col, _ = strconv.Atoi(buf.String())
				inCol = false
			} else if t.Name.Local == "row" {
				row, _ = strconv.Atoi(buf.String())
				inRow = false
			}
		case xml.CharData:
			if inCol || inRow {
				buf.Write(t)
			}
		}
	}
	return col, row
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
