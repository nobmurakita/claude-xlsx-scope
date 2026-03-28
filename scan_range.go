package main

import (
	"fmt"

	"github.com/nobmurakita/exceldump/internal/excel"
)

// parseScanRange は --range / --start フラグを解析して走査範囲を返す
func parseScanRange(rangeFlag, startFlag string) (scanRange *excel.CellRange, startCol, startRow int, err error) {
	if rangeFlag != "" && startFlag != "" {
		return nil, 0, 0, fmt.Errorf("--range と --start は同時に指定できません")
	}
	if rangeFlag != "" {
		r, err := excel.ParseRange(rangeFlag, "")
		if err != nil {
			return nil, 0, 0, err
		}
		return &r, 0, 0, nil
	}
	if startFlag != "" {
		startCol, startRow, err = excel.StartPosition(startFlag)
		if err != nil {
			return nil, 0, 0, err
		}
		return nil, startCol, startRow, nil
	}
	return nil, 0, 0, nil
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
