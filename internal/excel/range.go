package excel

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// CellRange は矩形のセル範囲を表す（1始まり）
type CellRange struct {
	StartCol int // 1始まり
	StartRow int // 1始まり
	EndCol   int
	EndRow   int
}

// String はExcel記法の範囲文字列を返す（例: "A1:H20"）
func (r CellRange) String() string {
	return fmt.Sprintf("%s%d:%s%d", colName(r.StartCol), r.StartRow, colName(r.EndCol), r.EndRow)
}

var (
	reCellRange = regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	reColRange  = regexp.MustCompile(`^([A-Z]+):([A-Z]+)$`)
	reRowRange  = regexp.MustCompile(`^(\d+):(\d+)$`)
	reSingleCell = regexp.MustCompile(`^([A-Z]+)(\d+)$`)
)

// ParseRange は範囲文字列をパースする。
// colOnly/rowOnly 指定時は usedRange で補完する。
func ParseRange(rangeStr, usedRange string) (CellRange, error) {
	s := strings.ToUpper(strings.TrimSpace(rangeStr))

	if m := reCellRange.FindStringSubmatch(s); m != nil {
		return parseCellRange(m[1], m[2], m[3], m[4])
	}
	if m := reColRange.FindStringSubmatch(s); m != nil {
		return parseColRange(m[1], m[2], usedRange)
	}
	if m := reRowRange.FindStringSubmatch(s); m != nil {
		return parseRowRange(m[1], m[2], usedRange)
	}
	if m := reSingleCell.FindStringSubmatch(s); m != nil {
		col := colNumber(m[1])
		row, _ := strconv.Atoi(m[2])
		return CellRange{col, row, col, row}, nil
	}
	return CellRange{}, fmt.Errorf("範囲の形式が不正です: %q", rangeStr)
}

func parseCellRange(sc, sr, ec, er string) (CellRange, error) {
	startCol := colNumber(sc)
	startRow, _ := strconv.Atoi(sr)
	endCol := colNumber(ec)
	endRow, _ := strconv.Atoi(er)
	if endCol < startCol || endRow < startRow {
		return CellRange{}, fmt.Errorf("範囲の終端が始端より前です: %s%s:%s%s", sc, sr, ec, er)
	}
	return CellRange{startCol, startRow, endCol, endRow}, nil
}

func parseColRange(sc, ec, usedRange string) (CellRange, error) {
	startCol := colNumber(sc)
	endCol := colNumber(ec)
	if endCol < startCol {
		return CellRange{}, fmt.Errorf("範囲の終端が始端より前です: %s:%s", sc, ec)
	}
	if usedRange == "" {
		return CellRange{}, nil
	}
	ur, err := ParseRange(usedRange, "")
	if err != nil {
		return CellRange{}, err
	}
	return CellRange{startCol, ur.StartRow, endCol, ur.EndRow}, nil
}

func parseRowRange(sr, er, usedRange string) (CellRange, error) {
	startRow, _ := strconv.Atoi(sr)
	endRow, _ := strconv.Atoi(er)
	if endRow < startRow {
		return CellRange{}, fmt.Errorf("範囲の終端が始端より前です: %s:%s", sr, er)
	}
	if usedRange == "" {
		return CellRange{}, nil
	}
	ur, err := ParseRange(usedRange, "")
	if err != nil {
		return CellRange{}, err
	}
	return CellRange{ur.StartCol, startRow, ur.EndCol, endRow}, nil
}

// StartPosition は --start で指定されたセル位置をパースする
func StartPosition(startCell string) (col, row int, err error) {
	s := strings.ToUpper(strings.TrimSpace(startCell))
	m := reSingleCell.FindStringSubmatch(s)
	if m == nil {
		return 0, 0, fmt.Errorf("セル位置の形式が不正です: %q", startCell)
	}
	col = colNumber(m[1])
	row, _ = strconv.Atoi(m[2])
	return col, row, nil
}

// IsEmpty は範囲が空（ゼロ値）かどうかを返す
func (r CellRange) IsEmpty() bool {
	return r.StartCol == 0 && r.StartRow == 0 && r.EndCol == 0 && r.EndRow == 0
}

// colNumber は列名を列番号（1始まり）に変換する（例: "A"→1, "Z"→26, "AA"→27）
func colNumber(name string) int {
	n := 0
	for _, c := range name {
		n = n*26 + int(c-'A') + 1
	}
	return n
}

// ColName は列番号（1始まり）を列名に変換する（例: 1→"A", 27→"AA"）
func ColName(n int) string {
	return colName(n)
}

func colName(n int) string {
	var s string
	for n > 0 {
		n--
		s = string(rune('A'+n%26)) + s
		n /= 26
	}
	return s
}

// CellRef はセル座標文字列を返す（例: col=2, row=3 → "B3"）
func CellRef(col, row int) string {
	return fmt.Sprintf("%s%d", colName(col), row)
}

