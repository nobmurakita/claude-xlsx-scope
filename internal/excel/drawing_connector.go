package excel

import (
	"encoding/xml"
	"math"
	"strconv"
	"strings"
)

// parseConnector は <cxnSp> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parseConnector(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	id := p.nextID
	p.nextID++

	shape := ShapeInfo{
		ID:   id,
		Type: "connector",
		Z:    z,
		Cell: cell,
	}

	if len(groupStack) > 0 {
		parentID := groupStack[len(groupStack)-1].seqID
		shape.Parent = &parentID
	}

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
						lineStyle = &LineStyle{}
						for _, attr := range t.Attr {
							if attr.Name.Local == "w" {
								w, _ := strconv.Atoi(attr.Value)
								lineStyle.Width = math.Round(float64(w)/12700*100) / 100
							}
						}
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
					headType := attrVal(t, "type")
					if headType != "" && headType != "none" {
						if shape.Arrow == "end" {
							shape.Arrow = "both"
						} else {
							shape.Arrow = "start"
						}
					}
				}
			case "tailEnd":
				if inLn {
					tailType := attrVal(t, "type")
					if tailType != "" && tailType != "none" {
						if shape.Arrow == "start" {
							shape.Arrow = "both"
						} else {
							shape.Arrow = "end"
						}
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
	if p.includeStyle && lineStyle != nil && (lineStyle.Color != "" || lineStyle.Style != "" || lineStyle.Width > 0) {
		if lineStyle.Style == "" && (lineStyle.Color != "" || lineStyle.Width > 0) {
			lineStyle.Style = "solid"
		}
		shape.Line = lineStyle
	}

	// Excel ID マッピング
	if excelID > 0 {
		p.excelIDMap[excelID] = id
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
			}
		}
		if cr.hasEnd {
			if seqID, ok := p.excelIDMap[cr.endID]; ok {
				p.shapes[cr.shapeIndex].To = &seqID
			}
		}
	}
}
