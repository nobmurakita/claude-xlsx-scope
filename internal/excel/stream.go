package excel

// StreamRows はシートの行を Rows() イテレータで走査し、
// 非空セルごとにコールバックを呼ぶ。dimension がないファイルで
// 全行メモリ展開を避けるためのストリーミング処理。
// callback が false を返すと走査を中断する。
func (f *File) StreamRows(sheet string, callback func(col, row int, value string) bool) error {
	rows, err := f.File.Rows(sheet)
	if err != nil {
		return err
	}
	defer rows.Close()

	rowIdx := 0
	for rows.Next() {
		rowIdx++
		cols, err := rows.Columns()
		if err != nil {
			continue
		}
		for c, val := range cols {
			if val != "" {
				if !callback(c+1, rowIdx, val) {
					return nil
				}
			}
		}
	}
	return nil
}
