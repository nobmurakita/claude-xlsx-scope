package excel

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestResolveCell_SharedString(t *testing.T) {
	ss := &sharedStrings{items: []sharedStringItem{
		{Text: "hello"},
		{Text: "world"},
	}}

	cell := RawCell{ValueType: vtSharedString, SharedStrIdx: -1}
	valueBuf := &strings.Builder{}
	valueBuf.WriteString("1") // 共有文字列インデックス
	formulaBuf := &strings.Builder{}
	inlineBuf := &strings.Builder{}

	resolveCell(&cell, ss, valueBuf, formulaBuf, inlineBuf)

	if cell.Value != "world" {
		t.Errorf("Value = %q, want %q", cell.Value, "world")
	}
	if cell.SharedStrIdx != 1 {
		t.Errorf("SharedStrIdx = %d, want 1", cell.SharedStrIdx)
	}
}

func TestResolveCell_InlineStr(t *testing.T) {
	cell := RawCell{ValueType: vtInlineStr}
	valueBuf := &strings.Builder{}
	formulaBuf := &strings.Builder{}
	inlineBuf := &strings.Builder{}
	inlineBuf.WriteString("inline text")

	resolveCell(&cell, nil, valueBuf, formulaBuf, inlineBuf)

	if cell.Value != "inline text" {
		t.Errorf("Value = %q, want %q", cell.Value, "inline text")
	}
}

func TestResolveCell_Formula(t *testing.T) {
	cell := RawCell{ValueType: vtNumber}
	valueBuf := &strings.Builder{}
	valueBuf.WriteString("42")
	formulaBuf := &strings.Builder{}
	formulaBuf.WriteString("=SUM(A1:A10)")
	inlineBuf := &strings.Builder{}

	resolveCell(&cell, nil, valueBuf, formulaBuf, inlineBuf)

	if cell.Value != "42" {
		t.Errorf("Value = %q, want %q", cell.Value, "42")
	}
	if cell.Formula != "=SUM(A1:A10)" {
		t.Errorf("Formula = %q, want %q", cell.Formula, "=SUM(A1:A10)")
	}
}

func TestResolveCell_NoFormula(t *testing.T) {
	cell := RawCell{ValueType: vtNumber}
	valueBuf := &strings.Builder{}
	valueBuf.WriteString("100")
	formulaBuf := &strings.Builder{}
	inlineBuf := &strings.Builder{}

	resolveCell(&cell, nil, valueBuf, formulaBuf, inlineBuf)

	if cell.Formula != "" {
		t.Errorf("Formula = %q, want empty", cell.Formula)
	}
}

func TestStreamWorksheetSAX(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet>
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><v>42</v></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>inline</t></is></c>
    </row>
  </sheetData>
</worksheet>`

	ss := &sharedStrings{items: []sharedStringItem{
		{Text: "shared text"},
	}}

	var cells []RawCell
	err := streamWorksheetSAXFromString(xmlData, ss, false, func(cell *RawCell) bool {
		cells = append(cells, *cell)
		return true
	})
	if err != nil {
		t.Fatalf("streamWorksheetSAX error: %v", err)
	}

	if len(cells) != 3 {
		t.Fatalf("got %d cells, want 3", len(cells))
	}

	// A1: 共有文字列
	if cells[0].Col != 1 || cells[0].Row != 1 || cells[0].Value != "shared text" {
		t.Errorf("cell[0] = {Col:%d, Row:%d, Value:%q}, want {1, 1, \"shared text\"}", cells[0].Col, cells[0].Row, cells[0].Value)
	}

	// B1: 数値
	if cells[1].Col != 2 || cells[1].Row != 1 || cells[1].Value != "42" {
		t.Errorf("cell[1] = {Col:%d, Row:%d, Value:%q}, want {2, 1, \"42\"}", cells[1].Col, cells[1].Row, cells[1].Value)
	}

	// A2: インライン文字列
	if cells[2].Col != 1 || cells[2].Row != 2 || cells[2].Value != "inline" {
		t.Errorf("cell[2] = {Col:%d, Row:%d, Value:%q}, want {1, 2, \"inline\"}", cells[2].Col, cells[2].Row, cells[2].Value)
	}
}

func TestStreamWorksheetSAX_EarlyStop(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`

	count := 0
	err := streamWorksheetSAXFromString(xmlData, nil, false, func(cell *RawCell) bool {
		count++
		return count < 2 // 2つ目で停止
	})
	if err != nil {
		t.Fatalf("error: %v", err)
	}
	if count != 2 {
		t.Errorf("count = %d, want 2", count)
	}
}

func TestStreamWorksheetSAX_WithFormula(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet>
  <sheetData>
    <row r="1">
      <c r="A1" s="1"><v>100</v><f>SUM(B1:B10)</f></c>
    </row>
  </sheetData>
</worksheet>`

	var cells []RawCell
	err := streamWorksheetSAXFromString(xmlData, nil, true, func(cell *RawCell) bool {
		cells = append(cells, *cell)
		return true
	})
	if err != nil {
		t.Fatalf("error: %v", err)
	}
	if len(cells) != 1 {
		t.Fatalf("got %d cells, want 1", len(cells))
	}
	if cells[0].Formula != "SUM(B1:B10)" {
		t.Errorf("Formula = %q, want %q", cells[0].Formula, "SUM(B1:B10)")
	}
	if cells[0].StyleID != 1 {
		t.Errorf("StyleID = %d, want 1", cells[0].StyleID)
	}
}

func TestStreamWorksheetSAX_StyleOnlyCell(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet>
  <sheetData>
    <row r="1">
      <c r="A1" s="5"></c>
    </row>
  </sheetData>
</worksheet>`

	var cells []RawCell
	err := streamWorksheetSAXFromString(xmlData, nil, false, func(cell *RawCell) bool {
		cells = append(cells, *cell)
		return true
	})
	if err != nil {
		t.Fatalf("error: %v", err)
	}
	if len(cells) != 1 {
		t.Fatalf("got %d cells, want 1 (style-only cell)", len(cells))
	}
	if cells[0].StyleID != 5 {
		t.Errorf("StyleID = %d, want 5", cells[0].StyleID)
	}
}

// streamWorksheetSAXFromString はテスト用のヘルパー
func streamWorksheetSAXFromString(xmlData string, ss *sharedStrings, needFormula bool, callback func(cell *RawCell) bool) error {
	if ss == nil {
		ss = &sharedStrings{}
	}
	decoder := xml.NewDecoder(strings.NewReader(xmlData))
	return streamWorksheetSAX(decoder, ss, needFormula, callback)
}
