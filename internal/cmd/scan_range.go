package cmd

import "github.com/nobmurakita/claude-xlsx-scope/internal/excel"

// parseScanRange は --range / --start フラグを解析して走査範囲を返す
func parseScanRange(rangeFlag, startFlag string) (scanRange *excel.CellRange, startCol, startRow int, err error) {
	if rangeFlag != "" {
		r, err := excel.ParseRange(rangeFlag, "")
		if err != nil {
			return nil, 0, 0, err
		}
		scanRange = &r
	}
	if startFlag != "" {
		startCol, startRow, err = excel.StartPosition(startFlag)
		if err != nil {
			return nil, 0, 0, err
		}
	}
	return scanRange, startCol, startRow, nil
}

// filterByRange はセルが走査範囲内かを判定する。
// skip=true: このセルをスキップ、stop=true: 走査終了
func filterByRange(col, row int, scanRange *excel.CellRange) (skip, stop bool) {
	if scanRange == nil {
		return false, false
	}
	if row < scanRange.StartRow || col < scanRange.StartCol {
		return true, false
	}
	if row > scanRange.EndRow {
		return false, true
	}
	if col > scanRange.EndCol {
		return true, false
	}
	return false, false
}

// filterByStart はセルが開始位置より前かを判定する（true=スキップ）
func filterByStart(col, row, startCol, startRow int) bool {
	if startCol > 0 {
		return row < startRow || (row == startRow && col < startCol)
	}
	return false
}

// shouldSkipCell は走査フィルタ・マージ結合のスキップ判定を行う。
// cells / search の StreamSheet コールバックで共通利用する。
func shouldSkipCell(col, row int, scanRange *excel.CellRange, startCol, startRow int, mergeInfo *excel.MergeInfo) (skip, stop bool) {
	if s, st := filterByRange(col, row, scanRange); s || st {
		return true, st
	}
	if filterByStart(col, row, startCol, startRow) {
		return true, false
	}
	if mergeInfo.IsMergedNonTopLeft(col, row) {
		return true, false
	}
	return false, false
}
