package excel

// RowCache はシートの非空セル情報を保持するキャッシュ。
// full モード: 全セル座標を記録（region 検出・空セルスキップ用）
// bounds モード: 行/列の境界のみ記録（軽量、大規模ファイル用）
type RowCache struct {
	occupied  map[[2]int]bool // full モード時のみ使用
	minCol    int
	maxCol    int
	minRow    int
	maxRow    int
	cellCount int
	boundsOnly bool // true の場合は境界情報のみ
}

// LoadRows はシートの全行を走査し、非空セルの座標を記録する
func (f *File) LoadRows(sheet string) (*RowCache, error) {
	return f.loadRowsInternal(sheet, false)
}

// LoadRowsBoundsOnly はシートの行/列境界と非空セル数のみ記録する（軽量版）
func (f *File) LoadRowsBoundsOnly(sheet string) (*RowCache, error) {
	return f.loadRowsInternal(sheet, true)
}

func (f *File) loadRowsInternal(sheet string, boundsOnly bool) (*RowCache, error) {
	rows, err := f.File.Rows(sheet)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	rc := &RowCache{
		minCol: -1, maxCol: -1, minRow: -1, maxRow: -1,
		boundsOnly: boundsOnly,
	}
	if !boundsOnly {
		rc.occupied = make(map[[2]int]bool)
	}

	rowIdx := 0
	for rows.Next() {
		rowIdx++
		cols, err := rows.Columns()
		if err != nil {
			continue
		}
		for c, val := range cols {
			if val != "" {
				colIdx := c + 1
				rc.cellCount++
				if !boundsOnly {
					rc.occupied[[2]int{colIdx, rowIdx}] = true
				}
				if rc.minRow == -1 || rowIdx < rc.minRow {
					rc.minRow = rowIdx
				}
				if rc.maxRow == -1 || rowIdx > rc.maxRow {
					rc.maxRow = rowIdx
				}
				if rc.minCol == -1 || colIdx < rc.minCol {
					rc.minCol = colIdx
				}
				if rc.maxCol == -1 || colIdx > rc.maxCol {
					rc.maxCol = colIdx
				}
			}
		}
	}
	return rc, nil
}

// CalcUsedRange はキャッシュから使用範囲を算出する
func (rc *RowCache) CalcUsedRange() string {
	if rc.minRow == -1 {
		return ""
	}
	return CellRef(rc.minCol, rc.minRow) + ":" + CellRef(rc.maxCol, rc.maxRow)
}

// HasValue はキャッシュ上で指定セルに値があるか返す。
// boundsOnly モードでは常に true を返す（呼び出し側で個別判定が必要）。
func (rc *RowCache) HasValue(col, row int) bool {
	if rc.boundsOnly {
		return true
	}
	return rc.occupied[[2]int{col, row}]
}

// NewRowCache は新しい RowCache を作成する。StreamSheet から構築する用途。
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

// CellCount は非空セルの総数を返す
func (rc *RowCache) CellCount() int {
	return rc.cellCount
}

// IsBoundsOnly は境界情報のみのモードかを返す
func (rc *RowCache) IsBoundsOnly() bool {
	return rc.boundsOnly
}
