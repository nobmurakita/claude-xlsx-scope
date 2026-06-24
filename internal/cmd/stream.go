package cmd

import (
	"encoding/json"
	"reflect"

	"github.com/nobmurakita/claude-xlsx-scope/internal/excel"
)

// streamConfig はストリーミング走査の設定
type streamConfig struct {
	f         *excel.File
	dc        *cellsContext
	enc       *json.Encoder
	scanRange *excel.CellRange
	startCol  int
	startRow  int
	limit     int

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
	lastRow := 0
	var encErr error
	acc := &cellAccumulator{enc: cfg.enc}

	// flushRowsUpTo は lastRow+1 から upTo までの各行について row 情報を出力する。
	// セルが1個も無い行や、スキャン範囲外のセルしか持たない行も拾うため、
	// セル出力のコールバック発火に頼らずここで埋める。
	flushRowsUpTo := func(upTo int) bool {
		if !cfg.emitRowInfo || cfg.dc.sheetMeta == nil {
			lastRow = upTo
			return true
		}
		startEmit, endEmit := scanRowRange(cfg, lastRow+1, upTo)
		if startEmit > endEmit {
			lastRow = upTo
			return true
		}
		if encErr = acc.flush(); encErr != nil {
			return false
		}
		for r := startEmit; r <= endEmit; r++ {
			if encErr = cfg.dc.emitRowInfo(cfg.enc, r); encErr != nil {
				return false
			}
		}
		lastRow = upTo
		return true
	}

	err := cfg.f.StreamSheet(cfg.dc.sheet, func(raw *excel.RawCell) bool {
		col, row := raw.Col, raw.Row

		// セル出現に先んじて、ここまでに通り過ぎた行の row 情報を埋める
		if row != lastRow {
			if !flushRowsUpTo(row) {
				return false
			}
		}

		if skip, stop := shouldSkipCell(col, row, cfg.scanRange, cfg.startCol, cfg.startRow, cfg.dc.mergeInfo); skip || stop {
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

	// scanRange 指定時は末尾の空行も埋める（切り詰めが無い場合のみ）
	if encErr == nil && result.TruncatedNext == "" && cfg.scanRange != nil {
		flushRowsUpTo(cfg.scanRange.EndRow)
	}

	if encErr != nil {
		return nil, encErr
	}
	if err != nil {
		return nil, err
	}
	return result, nil
}

// scanRowRange は flushRowsUpTo が出力対象とする行範囲 [start, end] を返す。
// scanRange / startRow による下限と upTo による上限を反映する。
func scanRowRange(cfg *streamConfig, candidateStart, candidateEnd int) (int, int) {
	startEmit := max(candidateStart, 1)
	endEmit := candidateEnd
	if cfg.scanRange != nil {
		startEmit = max(startEmit, cfg.scanRange.StartRow)
		endEmit = min(endEmit, cfg.scanRange.EndRow)
	} else if cfg.startRow > 0 {
		startEmit = max(startEmit, cfg.startRow)
	}
	return startEmit, endEmit
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
