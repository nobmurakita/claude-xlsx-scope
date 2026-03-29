package excel

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"log"
	"os"
	"strings"
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

	// 画像の抽出
	shape.ImagePath = p.extractEmbeddedImage(embedRID)

	return shape
}

// extractEmbeddedImage は embed RID から画像を解決し、一時ディレクトリに抽出してパスを返す
func (p *drawingParser) extractEmbeddedImage(embedRID string) string {
	if embedRID == "" || p.drawingRels == nil || p.extractDir == "" {
		return ""
	}

	rel, ok := p.drawingRels[embedRID]
	if !ok {
		return ""
	}

	// 画像ファイルの ZIP パスを解決
	imagePath := resolveRelTarget(p.drawingPath, rel.Target)

	// 拡張子を取得
	ext := ""
	if dotIdx := strings.LastIndex(imagePath, "."); dotIdx >= 0 {
		ext = strings.ToLower(imagePath[dotIdx+1:])
	}

	zipEntry, ok := p.zipEntries[imagePath]
	if !ok {
		return ""
	}

	return p.extractImage(zipEntry, ext)
}

// extractImage は ZIP エントリからファイルを抽出する
func (p *drawingParser) extractImage(entry *zip.File, ext string) string {
	rc, err := entry.Open()
	if err != nil {
		log.Printf("[WARN] extractImage: ZIPエントリ %s のオープンに失敗: %v", entry.Name, err)
		return ""
	}
	defer rc.Close()

	// 一意なファイル名を自動生成
	outFile, err := os.CreateTemp(p.extractDir, "image_*."+ext)
	if err != nil {
		log.Printf("[WARN] extractImage: 一時ファイルの作成に失敗: %v", err)
		return ""
	}
	outPath := outFile.Name()
	writeOK := false
	defer func() {
		if !writeOK {
			outFile.Close()
			os.Remove(outPath)
		}
	}()

	if _, err := io.Copy(outFile, rc); err != nil {
		log.Printf("[WARN] extractImage: 画像の書き込みに失敗: %v", err)
		return ""
	}
	if err := outFile.Close(); err != nil {
		log.Printf("[WARN] extractImage: ファイルのクローズに失敗: %v", err)
		return ""
	}
	writeOK = true

	return outPath
}
