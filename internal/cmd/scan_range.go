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

// shouldSkipCell は走査フィルタ・マージ結合のスキップ判定を行う。
// cells / search の StreamSheet コールバックで共通利用する。
// skip=true: このセルをスキップ、stop=true: 走査終了
func shouldSkipCell(col, row int, scanRange *excel.CellRange, startCol, startRow int, mergeInfo *excel.MergeInfo) (skip, stop bool) {
	if scanRange != nil {
		if row > scanRange.EndRow {
			return false, true
		}
		if row < scanRange.StartRow || col < scanRange.StartCol || col > scanRange.EndCol {
			return true, false
		}
	}
	if startCol > 0 && (row < startRow || (row == startRow && col < startCol)) {
		return true, false
	}
	if mergeInfo.IsMergedNonTopLeft(col, row) {
		return true, false
	}
	return false, false
}
