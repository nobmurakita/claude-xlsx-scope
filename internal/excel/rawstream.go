package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// XML セル値型（worksheet XML の c 要素の t 属性）
const (
	vtSharedString = "s"         // 共有文字列
	vtFormulaStr   = "str"       // 数式の文字列結果
	vtInlineStr    = "inlineStr" // インライン文字列
	vtBool         = "b"         // ブール値
	vtError        = "e"         // エラー値
	vtNumber       = "n"         // 数値
)

// RawCell はワークシートXMLから直接パースしたセルデータ。
// 1回のSAX走査で全属性を取得する。
type RawCell struct {
	Col          int
	Row          int
	Value        string // 共有文字列は解決済み
	StyleID      int
	Formula      string
	ValueType    string // vtSharedString, vtFormulaStr, vtInlineStr, vtBool, vtError, vtNumber, ""
	SharedStrIdx int    // 共有文字列のインデックス（ValueType==vtSharedString の場合のみ有効、-1 = 無効）
}

// StreamSheet はワークシートXMLを自前でSAXパースし、全セル属性を1パスで取得する。
// ワークシートの全セルデータを1パスで取得する。
// needFormula が false の場合でも、XMLの型属性が formula の場合は数式を取得する。
// callback が false を返すと走査を中断する。
func (f *File) StreamSheet(sheet string, needFormula bool, callback func(cell *RawCell) bool) error {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return fmt.Errorf("シート %q の XML パスが見つかりません", sheet)
	}

	entry := findZipEntry(f.zr, xmlPath)
	if entry == nil {
		return fmt.Errorf("ZIP 内に %s が見つかりません", xmlPath)
	}
	return streamWorksheetXML(entry, f.sharedStrings, needFormula, callback)
}

func streamWorksheetXML(entry *zip.File, ss *sharedStrings, needFormula bool, callback func(cell *RawCell) bool) error {
	return withZipXML(entry, func(decoder *xml.Decoder) error {
		return streamWorksheetSAX(decoder, ss, needFormula, callback)
	})
}

// worksheetSAXState は streamWorksheetSAX の SAX パーサー状態
type worksheetSAXState struct {
	inSheetData bool
	inRow       bool
	inCell      bool
	inValue     bool
	inFormula   bool
	inInlineStr bool
	inT         bool // <is> 内の <t>
}

func streamWorksheetSAX(decoder *xml.Decoder, ss *sharedStrings, needFormula bool, callback func(cell *RawCell) bool) error {
	var st worksheetSAXState
	var (
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
				st.inSheetData = true

			case "row":
				if !st.inSheetData {
					continue
				}
				st.inRow = true
				hasR := false
				for _, attr := range t.Attr {
					if attr.Name.Local == "r" {
						if r, err := strconv.Atoi(attr.Value); err == nil {
							currentRow = r
							hasR = true
						}
						break
					}
				}
				if !hasR {
					currentRow++
				}

			case "c":
				if !st.inRow {
					continue
				}
				st.inCell = true
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
						cell.ValueType = attr.Value
					case "s":
						if id, err := strconv.Atoi(attr.Value); err == nil {
							cell.StyleID = id
						}
					}
				}

			case "v":
				if st.inCell {
					st.inValue = true
					valueBuf.Reset()
				}

			case "f":
				if st.inCell {
					st.inFormula = true
					formulaBuf.Reset()
				}

			case "is":
				if st.inCell {
					st.inInlineStr = true
					inlineBuf.Reset()
				}

			case "t":
				if st.inInlineStr {
					st.inT = true
				}
			}

		case xml.EndElement:
			switch t.Name.Local {
			case "sheetData":
				return nil // sheetData 終了で走査完了

			case "row":
				st.inRow = false

			case "c":
				if !st.inCell {
					continue
				}
				st.inCell = false

				resolveCell(&cell, ss, &valueBuf, &formulaBuf, &inlineBuf)

				// スタイルのみのセルも --include-empty --style で出力するためコールバックに渡す
				if cell.Value == "" && cell.Formula == "" && cell.StyleID == 0 {
					continue
				}

				if !callback(&cell) {
					return nil
				}

			case "v":
				st.inValue = false

			case "f":
				st.inFormula = false

			case "is":
				st.inInlineStr = false

			case "t":
				st.inT = false
			}

		case xml.CharData:
			if st.inValue {
				valueBuf.Write(t)
			} else if st.inFormula {
				formulaBuf.Write(t)
			} else if st.inT && st.inInlineStr {
				inlineBuf.Write(t)
			}
		}
	}

	return nil
}

// resolveCell はセルの値と数式を各バッファから解決する
func resolveCell(cell *RawCell, ss *sharedStrings, valueBuf, formulaBuf, inlineBuf *strings.Builder) {
	// 値の解決
	if cell.ValueType == vtInlineStr {
		cell.Value = inlineBuf.String()
	} else if cell.ValueType == vtSharedString {
		// 共有文字列のインデックスを解決
		if idx, err := strconv.Atoi(valueBuf.String()); err == nil {
			cell.Value = ss.Get(idx)
			cell.SharedStrIdx = idx
		}
	} else {
		cell.Value = valueBuf.String()
	}

	// 数式（セルが数式セルかどうかの判定に必要なため、常に保持する）
	if formulaBuf.Len() > 0 {
		cell.Formula = formulaBuf.String()
	}
}

// Excel の列・行上限
const (
	maxExcelCol = 16384   // XFD
	maxExcelRow = 1048576
)

// parseCellRef はセル参照（例: "AB123"）を列番号と行番号に分解する
func parseCellRef(ref string) (col, row int) {
	i := 0
	for i < len(ref) {
		ch := ref[i]
		if ch >= 'A' && ch <= 'Z' {
			col = col*26 + int(ch-'A') + 1
		} else if ch >= 'a' && ch <= 'z' {
			col = col*26 + int(ch-'a') + 1
		} else {
			break
		}
		i++
	}
	for i < len(ref) {
		if ref[i] < '0' || ref[i] > '9' {
			break
		}
		row = row*10 + int(ref[i]-'0')
		i++
	}
	if col > maxExcelCol {
		col = maxExcelCol
	}
	if row > maxExcelRow {
		row = maxExcelRow
	}
	return col, row
}
