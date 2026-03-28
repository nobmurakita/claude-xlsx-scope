package excel

import "sort"

// defaultBandGap は領域分割時に使用するギャップ閾値。
// この行数/列数以上の空きがあれば別領域として分割する。
const defaultBandGap = 3

// Region は非空セルが密集する矩形領域
type Region struct {
	Range         string `json:"range"`
	NonEmptyCells int    `json:"non_empty_cells"`
}

// DetectRegionsFromCache は RowCache から領域を検出する（excelize 不要）
func DetectRegionsFromCache(usedRange CellRange, rowCache *RowCache) ([]Region, error) {
	if usedRange.IsEmpty() || rowCache == nil {
		return []Region{}, nil
	}

	occupiedRows := make(map[int]bool)
	occupiedCols := make(map[int]bool)
	type cell struct{ col, row int }
	var cells []cell

	for r := usedRange.StartRow; r <= usedRange.EndRow; r++ {
		for c := usedRange.StartCol; c <= usedRange.EndCol; c++ {
			if rowCache.HasValue(c, r) {
				cells = append(cells, cell{c, r})
				occupiedRows[r] = true
				occupiedCols[c] = true
			}
		}
	}

	if len(cells) == 0 {
		return []Region{}, nil
	}

	rowBands := splitIntoBands(occupiedRows, usedRange.StartRow, usedRange.EndRow, defaultBandGap)
	colBands := splitIntoBands(occupiedCols, usedRange.StartCol, usedRange.EndCol, defaultBandGap)

	var regions []Region
	for _, rb := range rowBands {
		for _, cb := range colBands {
			count := 0
			for _, c := range cells {
				if c.row >= rb[0] && c.row <= rb[1] && c.col >= cb[0] && c.col <= cb[1] {
					count++
				}
			}
			if count > 0 {
				r := CellRange{cb[0], rb[0], cb[1], rb[1]}
				regions = append(regions, Region{
					Range:         r.String(),
					NonEmptyCells: count,
				})
			}
		}
	}

	return regions, nil
}

// splitIntoBands は使用されているインデックスを、gap行/列以上の空きで分割する
func splitIntoBands(occupied map[int]bool, minIdx, maxIdx, gap int) [][2]int {
	indices := make([]int, 0, len(occupied))
	for idx := range occupied {
		indices = append(indices, idx)
	}
	sort.Ints(indices)

	if len(indices) == 0 {
		return nil
	}

	var bands [][2]int
	bandStart := indices[0]
	bandEnd := indices[0]

	for i := 1; i < len(indices); i++ {
		if indices[i]-bandEnd >= gap+1 {
			// gap以上の空きがある → バンドを分割
			bands = append(bands, [2]int{bandStart, bandEnd})
			bandStart = indices[i]
		}
		bandEnd = indices[i]
	}
	bands = append(bands, [2]int{bandStart, bandEnd})

	return bands
}
