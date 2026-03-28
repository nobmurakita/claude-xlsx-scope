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
}

// parseConnector は <cxnSp> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parseConnector(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	shape := p.newShapeInfo(ShapeTypeConnector, z, cell, groupStack)

	var cr connRef
	depth := 1
	var st connectorParseState
	sh := drawingStyleHandler{p: p}
	var (
		textParts   []string
		currentPara strings.Builder
		inTxBody    bool
		inP         bool
		inR         bool

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
					shape.Name, excelID = parseCNvPr(t)
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
				}
			case "endCxn":
				if st.inCxnSpPr {
					if v := attrVal(t, "id"); v != "" {
						cr.endID, _ = strconv.Atoi(v)
						cr.hasEnd = true
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
				if inP {
					inR = true
				}
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
			case "txBody":
				inTxBody = false
			case "p":
				if inP {
					textParts = append(textParts, currentPara.String())
					inP = false
				}
			case "r":
				inR = false
			}
			sh.handleEndElement(t.Name.Local)

		case xml.CharData:
			if inR || (inP && !inR) {
				text := string(t)
				if strings.TrimSpace(text) != "" {
					currentPara.Write(t)
				}
			}
		}
	}

	// テキスト
	shape.Label = strings.Join(textParts, "\n")

	// スタイル
	if p.includeStyle {
		shape.Line = finalizeLineStyle(sh.lineStyle)
	}

	// Excel ID マッピング
	p.registerExcelID(excelID, shape.ID)

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
			}
		}
		if cr.hasEnd {
			if seqID, ok := p.excelIDMap[cr.endID]; ok {
				p.shapes[cr.shapeIndex].To = &seqID
			}
		}
	}
}
