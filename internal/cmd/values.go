package cmd

import (
	"fmt"

	"github.com/nobmurakita/claude-xlsx-scope/internal/excel"
	"github.com/spf13/cobra"
)

// NewValuesCmd は values サブコマンドを生成する
func NewValuesCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "values <file>",
		Short: "セルの値のみを行単位でJSONL形式で出力する",
		Args:  cobra.ExactArgs(1),
		RunE:  runValues,
	}
	cmd.Flags().StringP("sheet", "s", "", "対象シート（名前 or 0始まりインデックス）")
	cmd.Flags().String("range", "", "セル範囲（例: A1:H20, A:F, 1:20）")
	cmd.Flags().Int("start", 0, "開始行番号（1始まり、0または未指定で先頭行）")
	cmd.Flags().Int("limit", defaultOutputLimit, "出力行数の上限（0で無制限）")
	return cmd
}

type rowsMetaOutput struct {
	Meta bool     `json:"meta"`
	Cols []string `json:"cols"`
}

type rowsRowOutput struct {
	Row    int   `json:"row"`
	Values []any `json:"values"`
}

type rowsTruncatedOutput struct {
	Truncated bool `json:"truncated"`
	NextRow   int  `json:"next_row"`
}

func runValues(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")
	rangeFlag, _ := cmd.Flags().GetString("range")
	startRow, _ := cmd.Flags().GetInt("start")
	limit, _ := cmd.Flags().GetInt("limit")

	f, sheet, err := openAndResolveSheet(args[0], sheetFlag)
	if err != nil {
		return err
	}
	defer f.Close()

	// 範囲の解析
	var scanRange *excel.CellRange
	if rangeFlag != "" {
		// used_range を取得して列・行範囲の補完に使用
		dim := f.LoadDimension(sheet)
		r, err := excel.ParseRange(rangeFlag, dim)
		if err != nil {
			return err
		}
		scanRange = &r
	}

	// scanRange がない場合は used_range から列範囲を決定
	startCol := 1
	endCol := 0
	if scanRange != nil {
		startCol = scanRange.StartCol
		endCol = scanRange.EndCol
	} else {
		dim := f.LoadDimension(sheet)
		if dim != "" {
			if r, err := excel.ParseRange(dim, ""); err == nil {
				startCol = r.StartCol
				endCol = r.EndCol
			}
		}
	}

	ow, err := newOutputWriter(cmd)
	if err != nil {
		return err
	}
	defer ow.cleanup()

	enc := newJSONLWriter(ow)

	// _meta 行: cols 配列を出力
	if endCol > 0 {
		cols := make([]string, 0, endCol-startCol+1)
		for c := startCol; c <= endCol; c++ {
			cols = append(cols, excel.ColName(c))
		}
		if err := enc.Encode(rowsMetaOutput{Meta: true, Cols: cols}); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	// ストリーミング走査
	rowCount := 0
	currentRow := -1
	var currentValues []any
	var encErr error

	flushRow := func() {
		if currentRow < 0 {
			return
		}
		// 末尾の null をトリム
		last := len(currentValues) - 1
		for last >= 0 && currentValues[last] == nil {
			last--
		}
		if last < 0 {
			// 全て null → 空行、スキップ
			currentRow = -1
			return
		}
		currentValues = currentValues[:last+1]
		encErr = enc.Encode(rowsRowOutput{Row: currentRow, Values: currentValues})
		rowCount++
		currentRow = -1
	}

	truncatedNextRow := 0

	err = f.StreamSheet(sheet, false, func(raw *excel.RawCell) bool {
		col, row := raw.Col, raw.Row

		// 範囲フィルタ
		if scanRange != nil {
			if row < scanRange.StartRow || col < scanRange.StartCol {
				return true
			}
			if row > scanRange.EndRow {
				return false
			}
			if col > scanRange.EndCol {
				return true
			}
		}

		// start フィルタ
		if startRow > 0 && row < startRow {
			return true
		}

		// 行が変わったらフラッシュ
		if row != currentRow {
			flushRow()
			if encErr != nil {
				return false
			}

			// limit チェック
			if limit > 0 && rowCount >= limit {
				truncatedNextRow = row
				return false
			}

			currentRow = row
			if endCol > 0 {
				currentValues = make([]any, endCol-startCol+1)
			} else {
				currentValues = nil
			}
		}

		// セルの値を取得
		data := f.RawCellToCellData(raw)
		if !data.HasValue || data.Type == excel.CellTypeEmpty {
			return true
		}

		// 値の決定: display があればそちらを優先
		var val any
		if data.Display != "" {
			val = data.Display
		} else {
			val = data.Value
		}

		// 配列に格納
		idx := col - startCol
		if endCol > 0 {
			if idx >= 0 && idx < len(currentValues) {
				currentValues[idx] = val
			}
		} else {
			// endCol 未定の場合は動的に拡張
			for idx >= len(currentValues) {
				currentValues = append(currentValues, nil)
			}
			currentValues[idx] = val
		}

		return true
	})

	// 最後の行をフラッシュ
	if encErr == nil && truncatedNextRow == 0 {
		flushRow()
	}

	if encErr != nil {
		return encErr
	}
	if err != nil {
		return err
	}

	// truncated 出力
	if truncatedNextRow > 0 {
		if err := enc.Encode(rowsTruncatedOutput{Truncated: true, NextRow: truncatedNextRow}); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	return ow.finalize()
}
