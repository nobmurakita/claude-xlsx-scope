package cmd

import (
	"encoding/json"
	"fmt"
	"math"
	"os"

	"github.com/nobmurakita/claude-xlsx-scope/internal/excel"
	"github.com/spf13/cobra"
)

// NewCellsCmd は cells サブコマンドを生成する
func NewCellsCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "cells <file>",
		Short: "セルの値と書式をJSONL形式で出力する",
		Args:  cobra.ExactArgs(1),
		RunE:  runCells,
	}
	cmd.Flags().StringP("sheet", "s", "", "対象シート（名前 or 0始まりインデックス）")
	cmd.Flags().String("range", "", "セル範囲（例: A1:H20, A:F, 1:20）")
	cmd.Flags().String("start", "", "開始セル位置（例: A51）")
	cmd.Flags().Bool("include-empty", false, "空セルも出力する")
	cmd.Flags().Bool("style", false, "書式情報を出力する")
	cmd.Flags().Bool("formula", false, "数式文字列を出力する")
	cmd.Flags().Int("limit", defaultOutputLimit, "出力セル数の上限（0で無制限）")
	return cmd
}

type metaOutput struct {
	Meta          bool               `json:"_meta"`
	Origin        *originOutput      `json:"origin,omitempty"`
	DefaultWidth  float64            `json:"default_width"`
	DefaultHeight float64            `json:"default_height"`
	ColWidths     map[string]float64 `json:"col_widths,omitempty"`
}

type originOutput struct {
	X int `json:"x"`
	Y int `json:"y"`
}

type rowOutput struct {
	Row    int     `json:"_row"`
	Height float64 `json:"height,omitempty"`
	Hidden bool    `json:"hidden,omitempty"`
}

func (dc *cellsContext) emitRowInfo(enc *json.Encoder, row int) error {
	if dc.sheetMeta == nil {
		return nil
	}
	info, ok := dc.sheetMeta.Rows[row]
	if !ok {
		return nil
	}
	ri := rowOutput{Row: row}
	heightDiffers := info.Height != 0 && info.Height != dc.defaultHeight
	if heightDiffers {
		ri.Height = math.Round(info.Height*excel.RowHeightPxFactor*100) / 100
	}
	ri.Hidden = info.Hidden
	if !heightDiffers && !ri.Hidden {
		return nil
	}
	return enc.Encode(ri)
}

func runCells(cmd *cobra.Command, args []string) error {
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

	dc, err := newCellsContext(f, sheet, showStyle, showFormula)
	if err != nil {
		return err
	}

	enc := newJSONLWriter(os.Stdout)

	// _meta 行を出力（col_widths, default_width/height, origin）
	if err := emitMeta(enc, dc.sheetMeta, scanRange, startCol, startRow); err != nil {
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
func emitMeta(enc *json.Encoder, meta *excel.SheetMeta, scanRange *excel.CellRange, startCol, startRow int) error {
	if meta == nil {
		return nil
	}
	out := metaOutput{
		Meta:          true,
		DefaultWidth:  math.Round(meta.EffectiveDefaultWidth()*excel.ColWidthPxFactor*100) / 100,
		DefaultHeight: math.Round(meta.DefaultHeight*excel.RowHeightPxFactor*100) / 100,
		ColWidths:     colWidthsFromMeta(meta),
		Origin:        buildOrigin(meta, scanRange, startCol, startRow),
	}
	return enc.Encode(out)
}

// buildOrigin は出力の起点セルとそのピクセル座標を構築する
func buildOrigin(meta *excel.SheetMeta, scanRange *excel.CellRange, startCol, startRow int) *originOutput {
	col, row := 1, 1
	if scanRange != nil {
		col, row = scanRange.StartCol, scanRange.StartRow
	} else if startCol > 0 {
		col, row = startCol, startRow
	}
	x, y := meta.CellOriginPx(col, row)
	return &originOutput{X: x, Y: y}
}

// emitTruncated は打ち切り行を出力する（truncatedNext が空なら何もしない）
func emitTruncated(enc *json.Encoder, truncatedNext string) error {
	if truncatedNext == "" {
		return nil
	}
	return enc.Encode(truncatedOutput{Truncated: true, NextCell: truncatedNext})
}

func colWidthsFromMeta(meta *excel.SheetMeta) map[string]float64 {
	dwPx := math.Round(meta.EffectiveDefaultWidth()*excel.ColWidthPxFactor*100) / 100
	// デフォルトと異なる幅の列を (col, px) ペアとして収集
	type colWidth struct {
		col int
		px  float64
	}
	var entries []colWidth
	for _, ci := range meta.Cols {
		px := math.Round(ci.Width*excel.ColWidthPxFactor*100) / 100
		if px != dwPx && ci.Width != 0 {
			for c := ci.Min; c <= ci.Max; c++ {
				entries = append(entries, colWidth{c, px})
			}
		}
	}
	if len(entries) == 0 {
		return nil
	}
	// 連続する同じ幅の列をまとめてキーを生成
	widths := make(map[string]float64)
	start := entries[0]
	prev := start
	for i := 1; i < len(entries); i++ {
		e := entries[i]
		if e.col == prev.col+1 && e.px == start.px {
			prev = e
		} else {
			widths[colRangeKey(start.col, prev.col)] = start.px
			start = e
			prev = e
		}
	}
	widths[colRangeKey(start.col, prev.col)] = start.px
	return widths
}

// colRangeKey は列範囲のキー文字列を返す（単一列: "B", 範囲: "B:D"）
func colRangeKey(min, max int) string {
	if min == max {
		return excel.ColName(min)
	}
	return excel.ColName(min) + ":" + excel.ColName(max)
}
