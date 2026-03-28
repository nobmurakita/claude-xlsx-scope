package excel

// RowCache はシートの非空セル座標の境界情報を保持するキャッシュ。
type RowCache struct {
	minCol int
	maxCol int
	minRow int
	maxRow int
}

// NewRowCache は新しい RowCache を作成する。
func NewRowCache() *RowCache {
	return &RowCache{
		minCol: -1, maxCol: -1, minRow: -1, maxRow: -1,
	}
}

// Add はセル座標を RowCache に追加する
func (rc *RowCache) Add(col, row int) {
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

// CalcUsedRange はキャッシュから使用範囲を算出する
func (rc *RowCache) CalcUsedRange() string {
	if rc.minRow == -1 {
		return ""
	}
	return CellRef(rc.minCol, rc.minRow) + ":" + CellRef(rc.maxCol, rc.maxRow)
}
