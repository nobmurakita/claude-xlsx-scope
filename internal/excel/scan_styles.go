package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
)

// ScanResult は ScanSheet の結果
type ScanResult struct {
	UsedRange     string // dimension があればそのまま、なければセルから算出
	ValueCount     int    // 値を持つセルの数
	MergedCells   int    // 結合セルの数
	StyleVariants int    // 視覚的に意味のあるスタイルのユニーク数
}

// ScanSheet はワークシートXMLを1パスで走査し、used_range・セル数・スタイルバリエーションを取得する。
// dimension が存在する場合はそれを used_range として使用し、セル走査で used_range の算出は行わない。
// 値や数式のパースは行わない。
func (f *File) ScanSheet(sheet string, visualIDs map[int]struct{}) (*ScanResult, error) {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return nil, fmt.Errorf("シート %q の XML パスが見つかりません", sheet)
	}

	entry := findZipEntry(f.zr, xmlPath)
	if entry == nil {
		return nil, fmt.Errorf("ZIP 内に %s が見つかりません", xmlPath)
	}

	return scanSheetFromEntry(entry, visualIDs)
}

func scanSheetFromEntry(entry *zip.File, visualIDs map[int]struct{}) (*ScanResult, error) {
	result := &ScanResult{}
	foundStyles := make(map[int]struct{})
	styleRemaining := len(visualIDs)
	allStylesFound := styleRemaining == 0

	rc := NewRowCache() // used_range 算出用（dimension なしの場合に使用）
	hasDimension := false

	err := withZipXML(entry, func(decoder *xml.Decoder) error {
		inSheetData := false
		inMergeCells := false

		for {
			tok, err := decoder.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				return err
			}

			switch t := tok.(type) {
			case xml.StartElement:
				switch t.Name.Local {
				case "dimension":
					for _, attr := range t.Attr {
						if attr.Name.Local == "ref" {
							result.UsedRange = attr.Value
							hasDimension = true
						}
					}
				case "sheetData":
					inSheetData = true
				case "c":
					if !inSheetData {
						continue
					}

					var styleID int
					var col, row int

					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "r":
							col, row = parseCellRef(attr.Value)
						case "s":
							if id, err := strconv.Atoi(attr.Value); err == nil {
								styleID = id
							}
						}
					}

					if !hasDimension && col > 0 && row > 0 {
						rc.Add(col, row)
					}

					if !allStylesFound && styleID > 0 {
						if _, isVisual := visualIDs[styleID]; isVisual {
							if _, already := foundStyles[styleID]; !already {
								foundStyles[styleID] = struct{}{}
								styleRemaining--
								allStylesFound = styleRemaining == 0
							}
						}
					}

					if scanCellHasValue(decoder) {
						result.ValueCount++
					}

				case "mergeCells":
					inMergeCells = true
				case "mergeCell":
					if inMergeCells {
						result.MergedCells++
					}
				}

			case xml.EndElement:
				switch t.Name.Local {
				case "sheetData":
					inSheetData = false
				case "mergeCells":
					return nil // mergeCells 以降は不要
				}
			}
		}
		return nil
	})

	if err != nil {
		return nil, err
	}

	if !hasDimension {
		result.UsedRange = rc.CalcUsedRange()
	}
	result.StyleVariants = len(foundStyles)

	return result, nil
}

// scanCellHasValue は <c> 要素の子要素に <v> または <is> があるかを判定する。
// </c> まで読み進め、結果を返す。
func scanCellHasValue(decoder *xml.Decoder) bool {
	depth := 1
	hasValue := false
	for {
		tok, err := decoder.Token()
		if err != nil {
			return hasValue
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			if t.Name.Local == "v" || t.Name.Local == "is" {
				hasValue = true
			}
		case xml.EndElement:
			depth--
			if depth == 0 {
				return hasValue
			}
		}
	}
}
