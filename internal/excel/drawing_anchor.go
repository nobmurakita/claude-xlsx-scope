package excel

import (
	"encoding/xml"
	"log"
	"math"
	"strings"
)

// anchorPos は anchor の from/to 内の位置情報
type anchorPos struct {
	col    int
	colOff int // EMU
	row    int
	rowOff int // EMU
}

// posCalculator はアンカー位置からポイント座標を算出する
type posCalculator struct {
	meta *SheetMeta
}

// calcX は列+オフセットからX座標（ポイント）を算出する（col は 0 始まり、off は EMU）
func (pc *posCalculator) calcX(col, off int) int {
	var x float64
	for c := 1; c <= col; c++ {
		x += pc.meta.ColWidthPt(c)
	}
	x += float64(off) / emuPerPt
	return int(math.Round(x))
}

// calcY は行+オフセットからY座標（ポイント）を算出する（row は 0 始まり、off は EMU）
func (pc *posCalculator) calcY(row, off int) int {
	var y float64
	for r := 1; r <= row; r++ {
		y += pc.meta.RowHeightPt(r)
	}
	y += float64(off) / emuPerPt
	return int(math.Round(y))
}

// cellRangeRef は from/to のセル座標（0始まり）から "A1:B2" または "A1" の形式の範囲文字列を返す
func cellRangeRef(fromCol, fromRow, toCol, toRow int) string {
	from := CellRef(fromCol+1, fromRow+1)
	to := CellRef(toCol+1, toRow+1)
	if from == to {
		return from
	}
	return from + ":" + to
}

// parseAnchorPos は <from> / <to> / <xdr:from> / <xdr:to> 内の col, colOff, row, rowOff を読む。
// 呼び出し時点で <from>/<to> の StartElement は消費済みで、EndElement まで消費する。
func parseAnchorPos(decoder *xml.Decoder) anchorPos {
	depth := 1
	var pos anchorPos
	var field string // "col", "colOff", "row", "rowOff"
	var buf strings.Builder

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			log.Printf("[WARN] parseAnchorPos: XMLトークン読み取りに失敗: %v", err)
			return anchorPos{}
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "col", "colOff", "row", "rowOff":
				field = t.Name.Local
				buf.Reset()
			}
		case xml.EndElement:
			depth--
			switch t.Name.Local {
			case "col":
				pos.col = safeAtoi(buf.String())
				field = ""
			case "colOff":
				pos.colOff = safeAtoi(buf.String())
				field = ""
			case "row":
				pos.row = safeAtoi(buf.String())
				field = ""
			case "rowOff":
				pos.rowOff = safeAtoi(buf.String())
				field = ""
			}
		case xml.CharData:
			if field != "" {
				buf.Write(t)
			}
		}
	}
	return pos
}

// twoAnchorPos は from/to アンカーからポイント座標の Position を算出する
func twoAnchorPos(pc *posCalculator, from, to anchorPos) *Position {
	x1 := pc.calcX(from.col, from.colOff)
	y1 := pc.calcY(from.row, from.rowOff)
	x2 := pc.calcX(to.col, to.colOff)
	y2 := pc.calcY(to.row, to.rowOff)
	return &Position{X: x1, Y: y1, W: x2 - x1, H: y2 - y1}
}

// buildPos はアンカー情報からポイント座標を算出する
func (p *drawingParser) buildPos(anchorType string, from, to anchorPos, hasTo bool, extCX, extCY, absX, absY int) *Position {
	if p.posCalc == nil {
		return nil
	}
	switch anchorType {
	case "two":
		if !hasTo {
			return nil
		}
		return twoAnchorPos(p.posCalc, from, to)
	case "one":
		x := p.posCalc.calcX(from.col, from.colOff)
		y := p.posCalc.calcY(from.row, from.rowOff)
		w := int(math.Round(float64(extCX) / emuPerPt))
		h := int(math.Round(float64(extCY) / emuPerPt))
		return &Position{X: x, Y: y, W: w, H: h}
	case "abs":
		x := int(math.Round(float64(absX) / emuPerPt))
		y := int(math.Round(float64(absY) / emuPerPt))
		w := int(math.Round(float64(extCX) / emuPerPt))
		h := int(math.Round(float64(extCY) / emuPerPt))
		return &Position{X: x, Y: y, W: w, H: h}
	}
	return nil
}
