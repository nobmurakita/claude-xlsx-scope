package excel

import (
	"encoding/xml"
	"strconv"
	"strings"
)

// connectorParseState は parseConnector の SAX パーサー状態
type connectorParseState struct {
	inNvCxnSpPr bool
	inCxnSpPr   bool // cNvCxnSpPr
	inSpPr      bool
	inLn        bool
	inFill      bool
	fillCtx     string
	inTxBody    bool
	inP         bool
	inR         bool
}

// parseConnector は <cxnSp> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parseConnector(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	shape, _ := p.newShapeInfo("connector", z, cell, groupStack)

	var cr connRef
	cr.shapeIndex = len(p.shapes)

	depth := 1
	var st connectorParseState
	var (
		textParts   []string
		currentPara strings.Builder

		lineStyle *LineStyle
		excelID   int
	)

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			p.warnings.Add("parseConnector: XMLトークン読み取りに失敗: %v", err)
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
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
			case "ln":
				if st.inSpPr {
					st.inLn = true
					if p.includeStyle {
						lineStyle = parseLineWidth(t)
					}
				}
			case "solidFill":
				st.inFill = true
				if st.inLn {
					st.fillCtx = "ln"
				} else {
					st.fillCtx = ""
				}
			case "srgbClr":
				if st.inFill {
					clr := attrVal(t, "val")
					clr = p.applyColorMods(decoder, depth, clr)
					depth--
					p.assignColor(clr, st.fillCtx, nil, lineStyle, nil, nil)
				}
			case "schemeClr":
				if st.inFill {
					clr := p.resolveSchemeColor(attrVal(t, "val"), decoder, depth)
					depth--
					p.assignColor(clr, st.fillCtx, nil, lineStyle, nil, nil)
				}
			case "prstDash":
				if st.inLn && lineStyle != nil {
					lineStyle.Style = attrVal(t, "val")
				}
			case "headEnd":
				if st.inLn {
					updateArrow(&shape.Arrow, "head", attrVal(t, "type"))
				}
			case "tailEnd":
				if st.inLn {
					updateArrow(&shape.Arrow, "tail", attrVal(t, "type"))
				}
			case "txBody":
				st.inTxBody = true
				textParts = nil
			case "p":
				if st.inTxBody {
					st.inP = true
					currentPara.Reset()
				}
			case "r":
				if st.inP {
					st.inR = true
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
			case "ln":
				st.inLn = false
			case "solidFill":
				st.inFill = false
				st.fillCtx = ""
			case "txBody":
				st.inTxBody = false
			case "p":
				if st.inP {
					textParts = append(textParts, currentPara.String())
					st.inP = false
				}
			case "r":
				st.inR = false
			}

		case xml.CharData:
			if st.inR || (st.inP && !st.inR) {
				text := string(t)
				if strings.TrimSpace(text) != "" {
					currentPara.Write(t)
				}
			}
		}
	}

	// テキスト
	shape.Label = strings.Join(textParts, "\n")
	if shape.Label == "" {
		shape.Label = ""
	}

	// スタイル
	if p.includeStyle {
		shape.Line = finalizeLineStyle(lineStyle)
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
