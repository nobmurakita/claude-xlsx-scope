package excel

import (
	"encoding/xml"
	"log"
	"strconv"
	"strings"
)

// parseConnector は <cxnSp> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parseConnector(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	shape, _ := p.newShapeInfo("connector", z, cell, groupStack)

	var cr connRef
	cr.shapeIndex = len(p.shapes)

	depth := 1
	var (
		inNvCxnSpPr bool
		inCxnSpPr   bool // cNvCxnSpPr
		inSpPr      bool
		inLn        bool
		inFill      bool
		fillCtx     string
		inTxBody    bool
		inP         bool
		inR         bool

		textParts   []string
		currentPara strings.Builder

		lineStyle *LineStyle
		excelID   int
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
			switch t.Name.Local {
			case "nvCxnSpPr":
				inNvCxnSpPr = true
			case "cNvPr":
				if inNvCxnSpPr {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "name":
							shape.Name = attr.Value
						case "id":
							excelID, _ = strconv.Atoi(attr.Value)
						}
					}
				}
			case "cNvCxnSpPr":
				if inNvCxnSpPr {
					inCxnSpPr = true
				}
			case "stCxn":
				if inCxnSpPr {
					if v := attrVal(t, "id"); v != "" {
						cr.startID, _ = strconv.Atoi(v)
						cr.hasStart = true
					}
				}
			case "endCxn":
				if inCxnSpPr {
					if v := attrVal(t, "id"); v != "" {
						cr.endID, _ = strconv.Atoi(v)
						cr.hasEnd = true
					}
				}
			case "spPr":
				inSpPr = true
			case "xfrm":
				if inSpPr {
					shape.Rotation, shape.Flip = parseXfrm(t)
				}
			case "prstGeom":
				if inSpPr {
					shape.ConnectorType = attrVal(t, "prst")
				}
			case "ln":
				if inSpPr {
					inLn = true
					if p.includeStyle {
						lineStyle = parseLineWidth(t)
					}
				}
			case "solidFill":
				inFill = true
				if inLn {
					fillCtx = "ln"
				} else {
					fillCtx = ""
				}
			case "srgbClr":
				if inFill {
					clr := attrVal(t, "val")
					clr = p.applyColorMods(decoder, depth, clr)
					depth--
					p.assignColor(clr, fillCtx, nil, lineStyle, nil, nil)
				}
			case "schemeClr":
				if inFill {
					clr := p.resolveSchemeColor(attrVal(t, "val"), decoder, depth)
					depth--
					p.assignColor(clr, fillCtx, nil, lineStyle, nil, nil)
				}
			case "prstDash":
				if inLn && lineStyle != nil {
					lineStyle.Style = attrVal(t, "val")
				}
			case "headEnd":
				if inLn {
					updateArrow(&shape.Arrow, "head", attrVal(t, "type"))
				}
			case "tailEnd":
				if inLn {
					updateArrow(&shape.Arrow, "tail", attrVal(t, "type"))
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
				inNvCxnSpPr = false
			case "cNvCxnSpPr":
				inCxnSpPr = false
			case "spPr":
				inSpPr = false
			case "ln":
				inLn = false
			case "solidFill":
				inFill = false
				fillCtx = ""
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
