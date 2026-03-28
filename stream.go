package main

import (
	"encoding/json"

	"github.com/nobmurakita/cc-read-excel/internal/excel"
)

// streamConfig はストリーミング走査の設定
type streamConfig struct {
	f           *excel.File
	dc          *dumpContext
	enc         *json.Encoder
	scanRange   *excel.CellRange
	startCol    int
	startRow    int
	limit       int
	showFormula bool

	// dump 固有: 空セル出力・行情報出力
	includeEmpty bool
	emitRowInfo  bool

	// search 固有: フィルタ
	filter *excel.Filter
}

// streamResult はストリーミング走査の結果
type streamResult struct {
	OutputCount   int
	TruncatedNext string
}

// runStream は dump/search 共通のストリーミング走査を実行する
func runStream(cfg *streamConfig) (*streamResult, error) {
	result := &streamResult{}
	lastRow := -1
	var encErr error

	err := cfg.f.StreamSheet(cfg.dc.sheet, cfg.showFormula, func(raw *excel.RawCell) bool {
		col, row := raw.Col, raw.Row

		if skip, stop := shouldSkipCell(col, row, cfg.scanRange, cfg.startCol, cfg.startRow, cfg.dc.mergeInfo); skip {
			return !stop
		}

		data := cfg.f.RawCellToCellData(raw)

		// 空セルの処理
		if !data.HasValue && data.Type == excel.CellTypeEmpty {
			if !cfg.includeEmpty {
				return true
			}
		}

		// search フィルタ
		if cfg.filter != nil {
			if !data.HasValue || data.Type == excel.CellTypeEmpty {
				return true
			}
			if !cfg.filter.MatchCell(data) {
				return true
			}
		}

		// limit チェック
		if cfg.limit > 0 && result.OutputCount >= cfg.limit {
			result.TruncatedNext = excel.CellRef(col, row)
			return false
		}

		// 行情報出力（dump 用）
		if cfg.emitRowInfo && row != lastRow {
			if encErr = cfg.dc.emitRowInfo(cfg.enc, row); encErr != nil {
				return false
			}
			lastRow = row
		}

		out := cfg.dc.buildCellOutput(col, row, data, raw)
		if encErr = cfg.enc.Encode(out); encErr != nil {
			return false
		}
		result.OutputCount++
		return true
	})

	if encErr != nil {
		return nil, encErr
	}
	if err != nil {
		return nil, err
	}
	return result, nil
}
