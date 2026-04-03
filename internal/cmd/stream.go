package cmd

import (
	"encoding/json"
	"reflect"

	"github.com/nobmurakita/claude-xlsx-scope/internal/excel"
)

// streamConfig はストリーミング走査の設定
type streamConfig struct {
	f           *excel.File
	dc          *cellsContext
	enc         *json.Encoder
	scanRange   *excel.CellRange
	startCol    int
	startRow    int
	limit       int
	showFormula bool

	// cells 固有: 空セル出力・行情報出力
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

// runStream は cells/search 共通のストリーミング走査を実行する
func runStream(cfg *streamConfig) (*streamResult, error) {
	result := &streamResult{}
	lastRow := -1
	var encErr error
	acc := &cellAccumulator{enc: cfg.enc}

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

		// 行情報出力（cells 用）— 行が変わったらバッファをフラッシュしてから行情報を出力
		if cfg.emitRowInfo && row != lastRow {
			if encErr = acc.flush(); encErr != nil {
				return false
			}
			if encErr = cfg.dc.emitRowInfo(cfg.enc, row); encErr != nil {
				return false
			}
			lastRow = row
		}

		out, styleDef := cfg.dc.buildCellOutput(col, row, data, raw)
		if encErr = acc.add(col, row, out, styleDef); encErr != nil {
			return false
		}
		result.OutputCount++
		return true
	})

	// 残りのバッファをフラッシュ
	if encErr == nil {
		encErr = acc.flush()
	}

	if encErr != nil {
		return nil, encErr
	}
	if err != nil {
		return nil, err
	}
	return result, nil
}

// cellAccumulator は同一行内で連続する同内容セルをまとめてトークンを節約するバッファ
type cellAccumulator struct {
	enc      *json.Encoder
	pending  *cellOutput
	styleDef *styleDefOutput
	startCol int
	endCol   int
	row      int
}

// flush はバッファされたセルを出力する
func (a *cellAccumulator) flush() error {
	if a.pending == nil {
		return nil
	}
	if a.styleDef != nil {
		if err := a.enc.Encode(a.styleDef); err != nil {
			return err
		}
	}
	// 複数セルがまとめられた場合は範囲表記にする
	if a.startCol != a.endCol {
		a.pending.Cell = excel.CellRef(a.startCol, a.row) + ":" + excel.CellRef(a.endCol, a.row)
	}
	err := a.enc.Encode(a.pending)
	a.pending = nil
	a.styleDef = nil
	return err
}

// add はセルを追加する。同一行で隣接かつ同内容なら範囲を拡張し、そうでなければフラッシュして新規バッファを開始する
func (a *cellAccumulator) add(col, row int, out cellOutput, styleDef *styleDefOutput) error {
	if a.pending != nil && row == a.row && col == a.endCol+1 && cellOutputContentEqual(a.pending, &out) {
		a.endCol = col
		return nil
	}
	if err := a.flush(); err != nil {
		return err
	}
	a.pending = &out
	a.styleDef = styleDef
	a.startCol = col
	a.endCol = col
	a.row = row
	return nil
}

// cellOutputContentEqual は cell フィールド以外の全フィールドが等しいか比較する
func cellOutputContentEqual(a, b *cellOutput) bool {
	aa, bb := *a, *b
	aa.Cell, bb.Cell = "", ""
	return reflect.DeepEqual(aa, bb)
}
