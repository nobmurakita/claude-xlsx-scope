package cmd

import "github.com/nobmurakita/cc-read-xlsx/internal/excel"

// openAndResolveSheet はファイルを開いてシートを解決する共通処理。
// 呼び出し元で defer f.Close() すること。
func openAndResolveSheet(path, sheetFlag string) (*excel.File, string, error) {
	f, err := excel.OpenFile(path)
	if err != nil {
		return nil, "", err
	}
	sheet, err := f.ResolveSheet(sheetFlag)
	if err != nil {
		f.Close()
		return nil, "", err
	}
	return f, sheet, nil
}
