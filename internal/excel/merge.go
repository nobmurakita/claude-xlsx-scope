package excel

// cellCoord はセル座標（列, 行）を表す
type cellCoord struct {
	col int
	row int
}

// MergeInfo はシート内の結合セル情報を保持する
type MergeInfo struct {
	topLeft map[cellCoord]string
	merged  map[cellCoord]bool
}

// IsTopLeft は指定セルが結合範囲の左上セルかを判定し、結合範囲を返す
func (mi *MergeInfo) IsTopLeft(col, row int) (string, bool) {
	r, ok := mi.topLeft[cellCoord{col, row}]
	return r, ok
}

// IsMergedNonTopLeft は結合セルの左上以外のセルかを判定する
func (mi *MergeInfo) IsMergedNonTopLeft(col, row int) bool {
	return mi.merged[cellCoord{col, row}]
}
