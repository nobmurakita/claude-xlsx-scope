package excel

import "strings"

// MergeInfo はシート内の結合セル情報を事前取得して保持する
type MergeInfo struct {
	// key: "col,row" of top-left cell → 結合範囲文字列（例: "B3:D3"）
	topLeft map[[2]int]string
	// key: "col,row" → 結合領域に含まれるが左上ではないセル
	merged map[[2]int]bool
}

// LoadMergeInfo はシートの結合セル情報をロードする
func (f *File) LoadMergeInfo(sheet string) (*MergeInfo, error) {
	cells, err := f.File.GetMergeCells(sheet)
	if err != nil {
		return nil, err
	}
	mi := &MergeInfo{
		topLeft: make(map[[2]int]string, len(cells)),
		merged:  make(map[[2]int]bool),
	}
	for _, mc := range cells {
		startAxis := mc.GetStartAxis()
		endAxis := mc.GetEndAxis()
		rangeStr := startAxis + ":" + endAxis

		sCol, sRow := parseCellAxis(startAxis)
		eCol, eRow := parseCellAxis(endAxis)

		mi.topLeft[[2]int{sCol, sRow}] = rangeStr

		for r := sRow; r <= eRow; r++ {
			for c := sCol; c <= eCol; c++ {
				if r == sRow && c == sCol {
					continue
				}
				mi.merged[[2]int{c, r}] = true
			}
		}
	}
	return mi, nil
}

// IsTopLeft は指定セルが結合範囲の左上セルかを判定し、結合範囲を返す
func (mi *MergeInfo) IsTopLeft(col, row int) (string, bool) {
	r, ok := mi.topLeft[[2]int{col, row}]
	return r, ok
}

// IsMergedNonTopLeft は結合セルの左上以外のセルかを判定する
func (mi *MergeInfo) IsMergedNonTopLeft(col, row int) bool {
	return mi.merged[[2]int{col, row}]
}

func parseCellAxis(axis string) (col, row int) {
	axis = strings.ToUpper(axis)
	i := 0
	for i < len(axis) && axis[i] >= 'A' && axis[i] <= 'Z' {
		i++
	}
	col = colNumber(axis[:i])
	row = 0
	for _, c := range axis[i:] {
		row = row*10 + int(c-'0')
	}
	return col, row
}
