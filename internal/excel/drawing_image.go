package excel

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"log"
	"math"
	"os"
	"strconv"
	"strings"
)

// pictureParseState は parsePicture の SAX パーサー状態
type pictureParseState struct {
	inNvPicPr  bool
	inBlipFill bool
	inSpPr     bool
}

// parsePicture は <pic> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parsePicture(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	shape := p.newShapeInfo(ShapeTypePicture, z, cell, groupStack)

	depth := 1
	var st pictureParseState
	var (
		embedRID     string
		excelID      int
		extCX, extCY int // EMU
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
			case "ext":
				if st.inSpPr {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "cx":
							extCX, _ = strconv.Atoi(attr.Value)
						case "cy":
							extCY, _ = strconv.Atoi(attr.Value)
						}
					}
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

	// 画像情報の構築と抽出
	shape.Image = p.resolveAndExtractImage(embedRID, extCX, extCY)

	return shape
}

// resolveAndExtractImage は embed RID から画像を解決し、抽出する
func (p *drawingParser) resolveAndExtractImage(embedRID string, extCX, extCY int) *ImageInfo {
	if embedRID == "" || p.drawingRels == nil {
		return nil
	}

	rel, ok := p.drawingRels[embedRID]
	if !ok {
		return nil
	}

	// 画像ファイルの ZIP パスを解決
	imagePath := resolveRelTarget(p.drawingPath, rel.Target)

	// 拡張子から形式を判定
	ext := ""
	if dotIdx := strings.LastIndex(imagePath, "."); dotIdx >= 0 {
		ext = strings.ToLower(imagePath[dotIdx+1:])
	}

	format := ext
	if format == "jpg" {
		format = "jpeg"
	}

	info := &ImageInfo{
		Format: format,
	}

	// EMU → ピクセル変換（1px = 9525 EMU）
	if extCX > 0 {
		info.Width = int(math.Round(float64(extCX) / emuPerPixel))
	}
	if extCY > 0 {
		info.Height = int(math.Round(float64(extCY) / emuPerPixel))
	}

	// ZIP エントリからファイルサイズを取得
	zipEntry, ok := p.zipEntries[imagePath]
	if !ok {
		return info
	}
	info.Size = int64(zipEntry.UncompressedSize64)

	// 画像を抽出
	if p.extractDir != "" {
		outPath := p.extractImage(zipEntry, ext)
		if outPath != "" {
			info.Path = outPath
		}
	}

	return info
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
