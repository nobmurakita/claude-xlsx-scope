package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// RawCell はワークシートXMLから直接パースしたセルデータ。
// 1回のSAX走査で全属性を取得するため、excelize へのAPIコールが不要。
type RawCell struct {
	Col          int
	Row          int
	Value        string // 共有文字列は解決済み
	StyleID      int
	Formula      string
	XMLType      string // "s", "str", "inlineStr", "b", "e", "n", ""
	SharedStrIdx int    // 共有文字列のインデックス（XMLType=="s" の場合のみ有効、-1 = 無効）
}

// StreamSheet はワークシートXMLを自前でSAXパースし、全セル属性を1パスで取得する。
// excelize の Rows()/GetCellType/GetCellValue/GetCellStyle/GetCellFormula を完全に置き換える。
// needFormula が false の場合でも、XMLの型属性が formula の場合は数式を取得する。
// callback が false を返すと走査を中断する。
func (f *File) StreamSheet(sheet string, needFormula bool, callback func(cell *RawCell) bool) error {
	if err := f.initStreamData(); err != nil {
		return err
	}

	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return fmt.Errorf("シート %q の XML パスが見つかりません", sheet)
	}

	zr, err := zip.OpenReader(f.path)
	if err != nil {
		return err
	}
	defer zr.Close()

	for _, entry := range zr.File {
		if entry.Name == xmlPath {
			return streamWorksheetXML(entry, f.sharedStrings, needFormula, callback)
		}
	}
	return fmt.Errorf("ZIP 内に %s が見つかりません", xmlPath)
}

func streamWorksheetXML(entry *zip.File, ss *sharedStrings, needFormula bool, callback func(cell *RawCell) bool) error {
	rc, err := entry.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	decoder := xml.NewDecoder(rc)

	// 状態管理
	var (
		inSheetData bool
		inRow       bool
		inCell      bool
		inValue     bool
		inFormula   bool
		inInlineStr bool
		inT         bool // <is> 内の <t>

		currentRow int
		cell       RawCell
		valueBuf   strings.Builder
		formulaBuf strings.Builder
		inlineBuf  strings.Builder
	)

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
			case "sheetData":
				inSheetData = true

			case "row":
				if !inSheetData {
					continue
				}
				inRow = true
				currentRow++
				for _, attr := range t.Attr {
					if attr.Name.Local == "r" {
						if r, err := strconv.Atoi(attr.Value); err == nil {
							currentRow = r
						}
						break
					}
				}

			case "c":
				if !inRow {
					continue
				}
				inCell = true
				cell = RawCell{Row: currentRow, SharedStrIdx: -1}
				valueBuf.Reset()
				formulaBuf.Reset()
				inlineBuf.Reset()
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "r":
						cell.Col, cell.Row = parseCellRef(attr.Value)
						if cell.Row == 0 {
							cell.Row = currentRow
						}
					case "t":
						cell.XMLType = attr.Value
					case "s":
						if id, err := strconv.Atoi(attr.Value); err == nil {
							cell.StyleID = id
						}
					}
				}

			case "v":
				if inCell {
					inValue = true
					valueBuf.Reset()
				}

			case "f":
				if inCell {
					inFormula = true
					formulaBuf.Reset()
				}

			case "is":
				if inCell {
					inInlineStr = true
					inlineBuf.Reset()
				}

			case "t":
				if inInlineStr {
					inT = true
				}
			}

		case xml.EndElement:
			switch t.Name.Local {
			case "sheetData":
				return nil // sheetData 終了で走査完了

			case "row":
				inRow = false

			case "c":
				if !inCell {
					continue
				}
				inCell = false

				// 値の解決
				if cell.XMLType == "inlineStr" {
					cell.Value = inlineBuf.String()
				} else if cell.XMLType == "s" {
					// 共有文字列のインデックスを解決
					if idx, err := strconv.Atoi(valueBuf.String()); err == nil {
						cell.Value = ss.Get(idx)
						cell.SharedStrIdx = idx
					}
				} else {
					cell.Value = valueBuf.String()
				}

				// 数式
				if needFormula || cell.XMLType == "str" {
					cell.Formula = formulaBuf.String()
				} else {
					// 数式は不要だが、値が数式の結果かどうかの判定用にフラグを保持
					if formulaBuf.Len() > 0 {
						cell.Formula = formulaBuf.String()
					}
				}

				// 空セルはスキップ
				if cell.Value == "" && cell.Formula == "" {
					continue
				}

				if !callback(&cell) {
					return nil
				}

			case "v":
				inValue = false

			case "f":
				inFormula = false

			case "is":
				inInlineStr = false

			case "t":
				inT = false
			}

		case xml.CharData:
			if inValue {
				valueBuf.Write(t)
			} else if inFormula {
				formulaBuf.Write(t)
			} else if inT && inInlineStr {
				inlineBuf.Write(t)
			}
		}
	}

	return nil
}

// parseCellRef はセル参照（例: "AB123"）を列番号と行番号に分解する
func parseCellRef(ref string) (col, row int) {
	i := 0
	for i < len(ref) && ref[i] >= 'A' && ref[i] <= 'Z' {
		col = col*26 + int(ref[i]-'A') + 1
		i++
	}
	for i < len(ref) {
		row = row*10 + int(ref[i]-'0')
		i++
	}
	return col, row
}
