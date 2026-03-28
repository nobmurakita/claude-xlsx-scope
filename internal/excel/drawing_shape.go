package excel

import (
	"encoding/xml"
	"math"
	"strconv"
	"strings"
)

// parseShape は <sp> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parseShape(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	id := p.nextID
	p.nextID++

	shape := ShapeInfo{
		ID:   id,
		Type: "customShape",
		Z:    z,
		Cell: cell,
	}

	if len(groupStack) > 0 {
		parentID := groupStack[len(groupStack)-1].seqID
		shape.Parent = &parentID
	}

	depth := 1
	var (
		inNvSpPr  bool
		inSpPr    bool
		inTxBody  bool
		inP       bool
		inR       bool
		inRPr     bool
		inDefRPr  bool
		inLn      bool
		inFill    bool // solidFill 直下
		fillCtx   string // "sp", "ln", "rPr", "defRPr"

		textParts   []string // 段落ごとのテキスト
		currentPara strings.Builder
		runs        []RichTextRun
		currentRunText strings.Builder
		currentFont    *parsedFont
		hasRuns        bool

		// スタイル
		shapeFill string
		lineStyle *LineStyle
		shapeFont *parsedFont

		excelID int
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
			case "nvSpPr":
				inNvSpPr = true
			case "cNvPr":
				if inNvSpPr {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "name":
							shape.Name = attr.Value
						case "id":
							excelID, _ = strconv.Atoi(attr.Value)
						}
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
					for _, attr := range t.Attr {
						if attr.Name.Local == "prst" {
							shape.Type = attr.Value
						}
					}
				}
			case "txBody":
				inTxBody = true
				textParts = nil
				runs = nil
				hasRuns = false
			case "p":
				if inTxBody {
					inP = true
					currentPara.Reset()
				}
			case "r":
				if inP {
					inR = true
					currentRunText.Reset()
					currentFont = nil
					hasRuns = true
				}
			case "rPr":
				if inR {
					inRPr = true
					currentFont = &parsedFont{}
				}
			case "defRPr":
				if inP && inTxBody && !inR {
					inDefRPr = true
					if p.includeStyle && shapeFont == nil {
						shapeFont = &parsedFont{}
					}
				}
			case "t":
				// テキスト要素（処理は CharData で）
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
				} else if inRPr {
					fillCtx = "rPr"
				} else if inDefRPr {
					fillCtx = "defRPr"
				} else if inSpPr {
					fillCtx = "sp"
				} else {
					fillCtx = ""
				}
			case "srgbClr":
				if inFill {
					clr := attrVal(t, "val")
					clr = p.applyColorMods(decoder, depth, clr)
					depth-- // applyColorMods が EndElement まで消費
					p.assignColor(clr, fillCtx, &shapeFill, lineStyle, currentFont, shapeFont)
				}
			case "schemeClr":
				if inFill {
					clr := p.resolveSchemeColor(attrVal(t, "val"), decoder, depth)
					depth-- // resolveSchemeColor が EndElement まで消費
					p.assignColor(clr, fillCtx, &shapeFill, lineStyle, currentFont, shapeFont)
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
			// rPr / defRPr 内のフォント属性
			case "latin", "ea":
				font := currentFont
				if inDefRPr {
					font = shapeFont
				}
				if font != nil {
					if v := attrVal(t, "typeface"); v != "" {
						font.Name = v
					}
				}
			case "sz":
				// DrawingML では sz は属性ではなく rPr の属性
			}

			// rPr / defRPr の属性からフォント情報取得
			if t.Name.Local == "rPr" && inR && currentFont != nil {
				parseDrawingFontAttrs(t, currentFont)
			}
			if t.Name.Local == "defRPr" && inDefRPr && shapeFont != nil {
				parseDrawingFontAttrs(t, shapeFont)
			}

		case xml.EndElement:
			depth--
			switch t.Name.Local {
			case "nvSpPr":
				inNvSpPr = false
			case "spPr":
				inSpPr = false
			case "txBody":
				inTxBody = false
			case "p":
				if inP {
					textParts = append(textParts, currentPara.String())
					inP = false
				}
			case "r":
				if inR {
					text := currentRunText.String()
					currentPara.WriteString(text)
					run := RichTextRun{Text: text}
					if currentFont != nil && p.includeStyle {
						run.Font = richTextFontDiffFromDrawing(currentFont, p.theme)
					}
					runs = append(runs, run)
					inR = false
				}
			case "rPr":
				inRPr = false
			case "defRPr":
				inDefRPr = false
			case "ln":
				inLn = false
			case "solidFill":
				inFill = false
				fillCtx = ""
			}

		case xml.CharData:
			if inP && !inR {
				// 段落直下のテキスト（<a:t> in <a:p> without <a:r>）
				text := string(t)
				if strings.TrimSpace(text) != "" {
					currentPara.Write(t)
				}
			}
			if inR {
				currentRunText.Write(t)
			}
		}
	}

	// テキストの組み立て
	shape.Text = strings.Join(textParts, "\n")
	if shape.Text == "" {
		shape.Text = ""
	}

	// リッチテキスト
	if hasRuns && len(runs) > 1 && p.includeStyle {
		shape.RichText = runs
	}

	// スタイル
	if p.includeStyle {
		if shapeFill != "" {
			shape.Fill = shapeFill
		}
		if lineStyle != nil && (lineStyle.Color != "" || lineStyle.Style != "" || lineStyle.Width > 0) {
			if lineStyle.Style == "" && (lineStyle.Color != "" || lineStyle.Width > 0) {
				lineStyle.Style = "solid"
			}
			shape.Line = lineStyle
		}
		if shapeFont != nil {
			shape.Font = buildDrawingFontObj(shapeFont, p.theme)
		}
	}

	// Excel ID マッピング
	if excelID > 0 {
		p.excelIDMap[excelID] = id
	}

	return shape
}

// startGroup は <grpSp> の先頭（nvGrpSpPr, grpSpPr）を読み、ShapeInfo を返す
// grpSp の EndElement は呼び出し元で処理される
func (p *drawingParser) startGroup(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	id := p.nextID
	p.nextID++

	shape := ShapeInfo{
		ID:   id,
		Type: "group",
		Z:    z,
		Cell: cell,
	}

	if len(groupStack) > 0 {
		parentID := groupStack[len(groupStack)-1].seqID
		shape.Parent = &parentID
	}

	// nvGrpSpPr と grpSpPr を読む
	// grpSp 内の子要素はメインループで処理されるため、ここでは先頭のプロパティだけ読む
	depth := 0
	readProps := false

	for {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "cNvPr":
				if depth <= 2 {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "name":
							shape.Name = attr.Value
						case "id":
							excelID, _ := strconv.Atoi(attr.Value)
							if excelID > 0 {
								p.excelIDMap[excelID] = id
							}
						}
					}
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
