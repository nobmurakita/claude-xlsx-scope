package excel

// rowCache はシートの非空セル座標の境界情報を保持するキャッシュ。
type rowCache struct {
	minCol int
	maxCol int
	minRow int
	maxRow int
}

// newRowCache は新しい rowCache を作成する。
func newRowCache() *rowCache {
	return &rowCache{
		minCol: -1, maxCol: -1, minRow: -1, maxRow: -1,
	}
}

// add はセル座標を rowCache に追加する
func (rc *rowCache) add(col, row int) {
	if rc.minRow == -1 || row < rc.minRow {
		rc.minRow = row
	}
	if rc.maxRow == -1 || row > rc.maxRow {
		rc.maxRow = row
	}
	if rc.minCol == -1 || col < rc.minCol {
		rc.minCol = col
	}
	if rc.maxCol == -1 || col > rc.maxCol {
		rc.maxCol = col
	}
}

// calcUsedRange はキャッシュから使用範囲を算出する
func (rc *rowCache) calcUsedRange() string {
	if rc.minRow == -1 {
		return ""
	}
	return CellRef(rc.minCol, rc.minRow) + ":" + CellRef(rc.maxCol, rc.maxRow)
}
