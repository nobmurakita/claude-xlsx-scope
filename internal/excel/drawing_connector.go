package excel

import (
	"encoding/xml"
	"log"
	"strconv"
	"strings"
)

// connectorParseState は parseConnector の SAX パーサー状態
type connectorParseState struct {
	inNvCxnSpPr bool
	inCxnSpPr   bool // cNvCxnSpPr
	inSpPr      bool
	inPrstGeom  bool
	inAvLst     bool
}

// parseConnector は <cxnSp> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parseConnector(decoder *xml.Decoder, z int, cell string, pos *Position, groupStack []groupContext) ShapeInfo {
	shape := p.newShapeInfo(ShapeTypeConnector, z, cell, groupStack)
	shape.Pos = pos

	var cr connRef
	depth := 1
	var st connectorParseState
	sh := drawingStyleHandler{p: p}
	var adjValues map[string]int
	var (
		textParts   []string
		currentPara strings.Builder
		inTxBody    bool
		inP         bool

		excelID int
	)

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			log.Printf("[WARN] parseConnector: XMLトークン読み取りに失敗: %v", err)
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++

			// スタイル処理（色・線・塗り）
			if handled, adj := sh.handleStartElement(t, decoder, st.inSpPr, false, false, nil, nil); handled {
				depth += adj
				continue
			}

			// 矢印処理
			sh.handleArrow(t, &shape.Arrow)

			switch t.Name.Local {
			case "nvCxnSpPr":
				st.inNvCxnSpPr = true
			case "cNvPr":
				if st.inNvCxnSpPr {
					shape.Name, excelID, shape.Hidden = parseCNvPr(t)
				}
			case "cNvCxnSpPr":
				if st.inNvCxnSpPr {
					st.inCxnSpPr = true
				}
			case "stCxn":
				if st.inCxnSpPr {
					if v := attrVal(t, "id"); v != "" {
						cr.startID, _ = strconv.Atoi(v)
						cr.hasStart = true
					}
					if v := attrVal(t, "idx"); v != "" {
						cr.startIdx, _ = strconv.Atoi(v)
					}
				}
			case "endCxn":
				if st.inCxnSpPr {
					if v := attrVal(t, "id"); v != "" {
						cr.endID, _ = strconv.Atoi(v)
						cr.hasEnd = true
					}
					if v := attrVal(t, "idx"); v != "" {
						cr.endIdx, _ = strconv.Atoi(v)
					}
				}
			case "spPr":
				st.inSpPr = true
			case "xfrm":
				if st.inSpPr {
					shape.Rotation, shape.Flip = parseXfrm(t)
				}
			case "prstGeom":
				if st.inSpPr {
					shape.ConnectorType = attrVal(t, "prst")
					st.inPrstGeom = true
				}
			case "avLst":
				if st.inPrstGeom {
					st.inAvLst = true
				}
			case "gd":
				if st.inAvLst {
					name := attrVal(t, "name")
					fmla := attrVal(t, "fmla")
					if name != "" && strings.HasPrefix(fmla, "val ") {
						val, _ := strconv.Atoi(strings.TrimPrefix(fmla, "val "))
						if adjValues == nil {
							adjValues = make(map[string]int)
						}
						adjValues[name] = val
					}
				}
			case "txBody":
				inTxBody = true
				textParts = nil
			case "p":
				if inTxBody {
					inP = true
					currentPara.Reset()
				}
			case "r":
				// inP 内の r は CharData で inP として処理される
			}

		case xml.EndElement:
			depth--
			switch t.Name.Local {
			case "nvCxnSpPr":
				st.inNvCxnSpPr = false
			case "cNvCxnSpPr":
				st.inCxnSpPr = false
			case "spPr":
				st.inSpPr = false
			case "prstGeom":
				st.inPrstGeom = false
			case "avLst":
				st.inAvLst = false
			case "txBody":
				inTxBody = false
			case "p":
				if inP {
					textParts = append(textParts, currentPara.String())
					inP = false
				}
			case "r":
				// noop
			}
			sh.handleEndElement(t.Name.Local)

		case xml.CharData:
			if inP {
				text := string(t)
				if strings.TrimSpace(text) != "" {
					currentPara.Write(t)
				}
			}
		}
	}

	// テキスト
	shape.Label = strings.Join(textParts, "\n")

	// 調整値
	shape.Adj = adjValues

	// スタイル
	if p.includeStyle {
		shape.Line = finalizeLineStyle(sh.lineStyle)
	}

	// Excel ID マッピング
	p.registerExcelID(excelID, shape.ID)

	// コネクタの始点・終点を算出
	if pos != nil {
		shape.Start, shape.End = connectorEndpoints(pos, shape.Flip)
	}

	// 接続情報を記録（後処理で解決）
	if cr.hasStart || cr.hasEnd {
		cr.shapeIndex = len(p.shapes)
		p.connRefs = append(p.connRefs, cr)
	}

	return shape
}

// resolveConnectors はコネクタの from/to を Excel ID から連番 ID に解決する
func (p *drawingParser) resolveConnectors() {
	for _, cr := range p.connRefs {
		if cr.shapeIndex >= len(p.shapes) {
			continue
		}
		if cr.hasStart {
			if seqID, ok := p.excelIDMap[cr.startID]; ok {
				p.shapes[cr.shapeIndex].From = &seqID
				idx := cr.startIdx
				p.shapes[cr.shapeIndex].FromIdx = &idx
			}
		}
		if cr.hasEnd {
			if seqID, ok := p.excelIDMap[cr.endID]; ok {
				p.shapes[cr.shapeIndex].To = &seqID
				idx := cr.endIdx
				p.shapes[cr.shapeIndex].ToIdx = &idx
			}
		}
	}
}
