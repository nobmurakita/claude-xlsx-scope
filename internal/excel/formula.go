package excel

// ReadFormula はセルの数式を取得する。
// 共有数式はexcelizeが展開した文字列を返す。
// 配列数式（CSE数式）は波括弧付きで返す。
func (f *File) ReadFormula(sheet string, col, row int) string {
	axis := CellRef(col, row)
	formula, err := f.File.GetCellFormula(sheet, axis)
	if err != nil {
		return ""
	}
	return formula
}
