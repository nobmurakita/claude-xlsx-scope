package excel

// RowCache はシートの非空セル情報を保持するキャッシュ。
type RowCache struct {
	occupied   map[[2]int]bool // full モード時のみ使用
	minCol     int
	maxCol     int
	minRow     int
	maxRow     int
	cellCount  int
	boundsOnly bool // true の場合は境界情報のみ
}

// NewRowCache は新しい RowCache を作成する。
func NewRowCache(boundsOnly bool) *RowCache {
	rc := &RowCache{
		minCol: -1, maxCol: -1, minRow: -1, maxRow: -1,
		boundsOnly: boundsOnly,
	}
	if !boundsOnly {
		rc.occupied = make(map[[2]int]bool)
	}
	return rc
}

// Add はセル座標を RowCache に追加する
func (rc *RowCache) Add(col, row int) {
	rc.cellCount++
	if !rc.boundsOnly {
		rc.occupied[[2]int{col, row}] = true
	}
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

// HasValue はキャッシュ上で指定セルに値があるか返す。
func (rc *RowCache) HasValue(col, row int) bool {
	if rc.boundsOnly {
		return true
	}
	return rc.occupied[[2]int{col, row}]
}

// CellCount は非空セルの総数を返す
func (rc *RowCache) CellCount() int {
	return rc.cellCount
}
