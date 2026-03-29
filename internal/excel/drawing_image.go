package excel

import (
	"encoding/xml"
	"log"
)

// pictureParseState は parsePicture の SAX パーサー状態
type pictureParseState struct {
	inNvPicPr  bool
	inBlipFill bool
	inSpPr     bool
}

// parsePicture は <pic> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parsePicture(decoder *xml.Decoder, z int, cell string, pos *Position, groupStack []groupContext) ShapeInfo {
	shape := p.newShapeInfo(ShapeTypePicture, z, cell, groupStack)
	shape.Pos = pos

	depth := 1
	var st pictureParseState
	var (
		embedRID string
		excelID  int
	)

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			log.Printf("[WARN] parsePicture: XMLトークン読み取りに失敗: %v", err)
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "nvPicPr":
				st.inNvPicPr = true
			case "cNvPr":
				if st.inNvPicPr {
					shape.Name, excelID = parseCNvPr(t)
					if v := attrVal(t, "descr"); v != "" {
						shape.AltText = v
					}
				}
			case "blipFill":
				st.inBlipFill = true
			case "blip":
				if st.inBlipFill {
					for _, attr := range t.Attr {
						if attr.Name.Local == "embed" {
							embedRID = attr.Value
						}
					}
				}
			case "spPr":
				st.inSpPr = true
			case "xfrm":
				if st.inSpPr {
					shape.Rotation, shape.Flip = parseXfrm(t)
				}
			}

		case xml.EndElement:
			depth--
			switch t.Name.Local {
			case "nvPicPr":
				st.inNvPicPr = false
			case "blipFill":
				st.inBlipFill = false
			case "spPr":
				st.inSpPr = false
			}
		}
	}

	// Excel ID マッピング
	p.registerExcelID(excelID, shape.ID)

	// 画像の ZIP パスを解決
	shape.ImageID = p.resolveImagePath(embedRID)

	return shape
}

// resolveImagePath は embed RID から ZIP 内の画像パスを解決する
func (p *drawingParser) resolveImagePath(embedRID string) string {
	if embedRID == "" || p.drawingRels == nil {
		return ""
	}

	rel, ok := p.drawingRels[embedRID]
	if !ok {
		return ""
	}

	return resolveRelTarget(p.drawingPath, rel.Target)
}
