package main

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/nobmurakita/exceldump/internal/excel"
	"github.com/spf13/cobra"
)

func init() {
	dumpCmd.Flags().StringP("sheet", "s", "", "対象シート（名前 or 0始まりインデックス）")
	dumpCmd.Flags().String("range", "", "セル範囲（例: A1:H20, A:F, 1:20）")
	dumpCmd.Flags().String("start", "", "開始セル位置（例: A51）")
	dumpCmd.Flags().Bool("include-empty", false, "空セルも出力する")
	dumpCmd.Flags().Bool("style", false, "書式情報を出力する")
	dumpCmd.Flags().Bool("formula", false, "数式文字列を出力する")
	dumpCmd.Flags().Int("limit", 1000, "出力セル数の上限（0で無制限）")
	rootCmd.AddCommand(dumpCmd)
}

var dumpCmd = &cobra.Command{
	Use:   "dump <file>",
	Short: "セルの値と書式をJSONL形式でダンプする",
	Args:  cobra.ExactArgs(1),
	RunE:  runDump,
}

type metaOutput struct {
	Meta          bool               `json:"_meta"`
	DefaultWidth  float64            `json:"default_width"`
	DefaultHeight float64            `json:"default_height"`
	ColWidths     map[string]float64 `json:"col_widths,omitempty"`
}

type rowOutput struct {
	Row    int     `json:"_row"`
	Height float64 `json:"height,omitempty"`
	Hidden bool    `json:"hidden,omitempty"`
}

func (dc *dumpContext) emitRowInfo(enc *json.Encoder, row int) error {
	if dc.sheetMeta == nil {
		return nil
	}
	info, ok := dc.sheetMeta.Rows[row]
	if !ok {
		return nil
	}
	ri := rowOutput{Row: row}
	if info.Height != dc.defaultHeight {
		ri.Height = info.Height
	}
	ri.Hidden = info.Hidden
	if ri.Height == 0 && !ri.Hidden {
		return nil
	}
	return enc.Encode(ri)
}

func runDump(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")
	rangeFlag, _ := cmd.Flags().GetString("range")
	startFlag, _ := cmd.Flags().GetString("start")
	includeEmpty, _ := cmd.Flags().GetBool("include-empty")
	showStyle, _ := cmd.Flags().GetBool("style")
	showFormula, _ := cmd.Flags().GetBool("formula")
	limit, _ := cmd.Flags().GetInt("limit")

	if rangeFlag != "" && startFlag != "" {
		return fmt.Errorf("--range と --start は同時に指定できません")
	}

	f, err := excel.OpenFile(args[0])
	if err != nil {
		return err
	}
	defer f.Close()

	sheet, err := f.ResolveSheet(sheetFlag)
	if err != nil {
		return err
	}

	// 走査範囲の決定
	var scanRange *excel.CellRange
	var startCol, startRow int

	if rangeFlag != "" {
		r, err := excel.ParseRange(rangeFlag, "")
		if err != nil {
			return err
		}
		scanRange = &r
	} else if startFlag != "" {
		startCol, startRow, err = excel.StartPosition(startFlag)
		if err != nil {
			return err
		}
	}

	dc, err := newDumpContext(f, sheet, showStyle, showFormula)
	if err != nil {
		return err
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)

	// _meta 行を出力（col_widths, default_width/height）
	if dc.sheetMeta != nil {
		meta := metaOutput{
			Meta:          true,
			DefaultWidth:  dc.sheetMeta.EffectiveDefaultWidth(),
			DefaultHeight: dc.sheetMeta.DefaultHeight,
			ColWidths:     colWidthsFromMeta(dc.sheetMeta),
		}
		if err := enc.Encode(meta); err != nil {
			return err
		}
	}

	outputCount := 0
	lastRow := -1
	var truncatedNext string
	var encErr error

	err = f.StreamSheet(sheet, showFormula, func(raw *excel.RawCell) bool {
		col, row := raw.Col, raw.Row

		if skip, stop := filterByRange(col, row, scanRange); skip || stop {
			return !stop
		}

		if filterByStart(col, row, startCol, startRow) {
			return true
		}

		if dc.mergeInfo.IsMergedNonTopLeft(col, row) {
			return true
		}

		data := f.RawCellToCellData(raw)

		if !data.HasValue && data.Type == excel.CellTypeEmpty {
			if !includeEmpty {
				return true
			}
		}

		if limit > 0 && outputCount >= limit {
			truncatedNext = excel.CellRef(col, row)
			return false
		}

		// 行が変わったら行情報を出力
		if row != lastRow {
			if encErr = dc.emitRowInfo(enc, row); encErr != nil {
				return false
			}
			lastRow = row
		}

		out := dc.buildCellOutput(col, row, data, raw)
		if encErr = enc.Encode(out); encErr != nil {
			return false
		}
		outputCount++
		return true
	})

	if encErr != nil {
		return encErr
	}
	if err != nil {
		return err
	}

	if truncatedNext != "" {
		if err := enc.Encode(truncatedOutput{Truncated: true, NextCell: truncatedNext}); err != nil {
			return err
		}
	}

	return nil
}

func colWidthsFromMeta(meta *excel.SheetMeta) map[string]float64 {
	dw := meta.EffectiveDefaultWidth()
	widths := make(map[string]float64)
	for _, ci := range meta.Cols {
		if ci.Width != dw && ci.Width != 0 {
			for c := ci.Min; c <= ci.Max; c++ {
				widths[excel.ColName(c)] = ci.Width
			}
		}
	}
	if len(widths) == 0 {
		return nil
	}
	return widths
}
