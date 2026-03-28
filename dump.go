package main

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-excel/internal/excel"
	"github.com/spf13/cobra"
)

func init() {
	dumpCmd.Flags().StringP("sheet", "s", "", "対象シート（名前 or 0始まりインデックス）")
	dumpCmd.Flags().String("range", "", "セル範囲（例: A1:H20, A:F, 1:20）")
	dumpCmd.Flags().String("start", "", "開始セル位置（例: A51）")
	dumpCmd.Flags().Bool("include-empty", false, "空セルも出力する")
	dumpCmd.Flags().Bool("style", false, "書式情報を出力する")
	dumpCmd.Flags().Bool("formula", false, "数式文字列を出力する")
	dumpCmd.Flags().Int("limit", defaultOutputLimit, "出力セル数の上限（0で無制限）")
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

	scanRange, startCol, startRow, err := parseScanRange(rangeFlag, startFlag)
	if err != nil {
		return err
	}

	f, sheet, err := openAndResolveSheet(args[0], sheetFlag)
	if err != nil {
		return err
	}
	defer f.Close()

	dc, err := newDumpContext(f, sheet, showStyle, showFormula)
	if err != nil {
		return err
	}

	enc := newJSONLWriter(os.Stdout)

	// _meta 行を出力（col_widths, default_width/height）
	if err := emitMeta(enc, dc.sheetMeta); err != nil {
		return err
	}

	result, err := runStream(&streamConfig{
		f:            f,
		dc:           dc,
		enc:          enc,
		scanRange:    scanRange,
		startCol:     startCol,
		startRow:     startRow,
		limit:        limit,
		showFormula:  showFormula,
		includeEmpty: includeEmpty,
		emitRowInfo:  true,
	})
	if err != nil {
		return err
	}

	if err := emitTruncated(enc, result.TruncatedNext); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}
	return nil
}

// emitMeta は _meta 行を出力する
func emitMeta(enc *json.Encoder, meta *excel.SheetMeta) error {
	if meta == nil {
		return nil
	}
	return enc.Encode(metaOutput{
		Meta:          true,
		DefaultWidth:  meta.EffectiveDefaultWidth(),
		DefaultHeight: meta.DefaultHeight,
		ColWidths:     colWidthsFromMeta(meta),
	})
}

// emitTruncated は打ち切り行を出力する（truncatedNext が空なら何もしない）
func emitTruncated(enc *json.Encoder, truncatedNext string) error {
	if truncatedNext == "" {
		return nil
	}
	return enc.Encode(truncatedOutput{Truncated: true, NextCell: truncatedNext})
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
