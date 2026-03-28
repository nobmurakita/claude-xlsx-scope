package excel

// MergeInfo はシート内の結合セル情報を保持する
type MergeInfo struct {
	topLeft map[[2]int]string
	merged  map[[2]int]bool
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
