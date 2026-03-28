package excel

import (
	"encoding/xml"
	"log"
)

// shapeParseState は parseShape の SAX パーサー状態
type shapeParseState struct {
	inNvSpPr bool
	inSpPr   bool
}

// parseShape は <sp> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parseShape(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	shape, _ := p.newShapeInfo("customShape", z, cell, groupStack)

	depth := 1
	var st shapeParseState
	var ts drawingTextState
	sh := drawingStyleHandler{p: p}
	var excelID int

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			log.Printf("[WARN] parseShape: XMLトークン読み取りに失敗: %v", err)
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++

			// テキスト処理
			ts.handleStartElement(t, p.includeStyle)

			// スタイル処理（色・線・塗り）
			if handled, adj := sh.handleStartElement(t, decoder, st.inSpPr, ts.inRPr, ts.inDefRPr, ts.currentFont, ts.shapeFont); handled {
				depth += adj
				continue
			}

			// 矢印処理
			sh.handleArrow(t, &shape.Arrow)

			switch t.Name.Local {
			case "nvSpPr":
				st.inNvSpPr = true
			case "cNvPr":
				if st.inNvSpPr {
					shape.Name, excelID = parseCNvPr(t)
				}
			case "spPr":
				st.inSpPr = true
			case "xfrm":
				if st.inSpPr {
					shape.Rotation, shape.Flip = parseXfrm(t)
				}
			case "prstGeom":
				if st.inSpPr {
					if v := attrVal(t, "prst"); v != "" {
						shape.Type = v
					}
				}
			}

		case xml.EndElement:
			depth--
			switch t.Name.Local {
			case "nvSpPr":
				st.inNvSpPr = false
			case "spPr":
				st.inSpPr = false
			}
			ts.handleEndElement(t.Name.Local, p.includeStyle, p.theme)
			sh.handleEndElement(t.Name.Local)

		case xml.CharData:
			ts.handleCharData(t)
		}
	}

	// テキスト・スタイルの組み立て
	shape.Text = ts.buildText()
	if ts.hasRuns && len(ts.runs) > 1 && p.includeStyle {
		shape.RichText = ts.runs
	}
	if p.includeStyle {
		if sh.shapeFill != "" {
			shape.Fill = sh.shapeFill
		}
		shape.Line = finalizeLineStyle(sh.lineStyle)
		if ts.shapeFont != nil {
			shape.Font = buildDrawingFontObj(ts.shapeFont, p.theme)
		}
	}
	p.registerExcelID(excelID, shape.ID)
	return shape
}

// determineFillCtx は solidFill のコンテキストを判定する
func determineFillCtx(inLn, inRPr, inDefRPr, inSpPr bool) string {
	switch {
	case inLn:
		return "ln"
	case inRPr:
		return "rPr"
	case inDefRPr:
		return "defRPr"
	case inSpPr:
		return "sp"
	default:
		return ""
	}
}

// startGroup は <grpSp> の先頭（nvGrpSpPr, grpSpPr）を読み、ShapeInfo を返す
// grpSp の EndElement は呼び出し元で処理される
func (p *drawingParser) startGroup(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	shape, _ := p.newShapeInfo("group", z, cell, groupStack)

	// nvGrpSpPr と grpSpPr を読む
	// grpSp 内の子要素はメインループで処理されるため、ここでは先頭のプロパティだけ読む
	depth := 0
	readProps := false

	for {
		tok, err := decoder.Token()
		if err != nil {
			log.Printf("[WARN] startGroup: XMLトークン読み取りに失敗: %v", err)
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "cNvPr":
				if depth <= 2 {
					name, eid := parseCNvPr(t)
					shape.Name = name
					p.registerExcelID(eid, shape.ID)
				}
			case "xfrm":
				if depth <= 2 {
					shape.Rotation, shape.Flip = parseXfrm(t)
				}
			}
		case xml.EndElement:
			depth--
			if t.Name.Local == "grpSpPr" {
				readProps = true
			}
			if depth < 0 || readProps {
				return shape
			}
		}
	}

	return shape
}
