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
	dumpCmd.Flags().Bool("no-style", false, "書式情報を省略する")
	dumpCmd.Flags().Int("limit", 1000, "出力セル数の上限（0で無制限）")
	rootCmd.AddCommand(dumpCmd)
}

var dumpCmd = &cobra.Command{
	Use:   "dump <file>",
	Short: "セルの値と書式をJSONL形式でダンプする",
	Args:  cobra.ExactArgs(1),
	RunE:  runDump,
}

// cellOutput はJSONL出力する1セルの情報
type cellOutput struct {
	Cell      string              `json:"cell"`
	Value     any                 `json:"value,omitempty"`
	Display   string              `json:"display,omitempty"`
	Type      excel.CellType      `json:"type"`
	Merge     string              `json:"merge,omitempty"`
	Formula   string              `json:"formula,omitempty"`
	Link      *excel.HyperlinkData `json:"link,omitempty"`
	HiddenRow bool                `json:"hidden_row,omitempty"`
	HiddenCol bool                `json:"hidden_col,omitempty"`
	Font      *excel.FontObj      `json:"font,omitempty"`
	Fill      *excel.FillObj      `json:"fill,omitempty"`
	Border    *excel.BorderObj    `json:"border,omitempty"`
	Alignment *excel.AlignmentObj `json:"alignment,omitempty"`
	RichText  []excel.RichTextRun `json:"rich_text,omitempty"`
}

func runDump(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")
	rangeFlag, _ := cmd.Flags().GetString("range")
	startFlag, _ := cmd.Flags().GetString("start")
	includeEmpty, _ := cmd.Flags().GetBool("include-empty")
	noStyle, _ := cmd.Flags().GetBool("no-style")
	limit, _ := cmd.Flags().GetInt("limit")

	if rangeFlag != "" && startFlag != "" {
		return fmt.Errorf("--range と --start は同時に指定できません")
	}

	f, err := excel.OpenFile(args[0])
	if err != nil {
		return err
	}
	defer f.Close()

	sheet, err := f.ResolveWorksheet(sheetFlag)
	if err != nil {
		return err
	}

	usedRangeStr, err := f.GetUsedRange(sheet)
	if err != nil {
		return err
	}

	// 走査範囲の決定
	var scanRange excel.CellRange
	var startCol, startRow int
	useStart := false

	if rangeFlag != "" {
		scanRange, err = excel.ParseRange(rangeFlag, usedRangeStr)
		if err != nil {
			return err
		}
	} else if startFlag != "" {
		startCol, startRow, err = excel.StartPosition(startFlag)
		if err != nil {
			return err
		}
		useStart = true
		if usedRangeStr != "" {
			scanRange, _ = excel.ParseRange(usedRangeStr, "")
		}
	} else {
		if usedRangeStr == "" {
			return nil // 空シート
		}
		scanRange, err = excel.ParseRange(usedRangeStr, "")
		if err != nil {
			return err
		}
	}

	if scanRange.IsEmpty() {
		return nil
	}

	// デフォルトフォント（書式差分計算用）
	var defaultFont excel.FontInfo
	if !noStyle {
		defaultFont = f.DetectDefaultFont(sheet, scanRange)
	}

	// 結合セル情報
	mergeInfo, err := f.LoadMergeInfo(sheet)
	if err != nil {
		return err
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)

	outputCount := 0
	totalCount := 0
	nextCell := ""
	limitReached := false

	// セル走査（行優先順）
	for row := scanRange.StartRow; row <= scanRange.EndRow; row++ {
		for col := scanRange.StartCol; col <= scanRange.EndCol; col++ {
			// --start: 開始位置より前のセルをスキップ
			if useStart {
				if row < startRow || (row == startRow && col < startCol) {
					continue
				}
			}

			// 結合セルの左上以外はスキップ
			if mergeInfo.IsMergedNonTopLeft(col, row) {
				continue
			}

			// セル値の読み取り
			data, err := f.ReadCell(sheet, col, row)
			if err != nil {
				return err
			}

			// 空セルの処理
			if !data.HasValue && data.Type == excel.CellTypeEmpty {
				if !includeEmpty {
					continue
				}
			}

			totalCount++

			// limit 超過: next_cell を記録してカウントを続ける
			if limitReached {
				if nextCell == "" {
					nextCell = excel.CellRef(col, row)
				}
				continue
			}
			if limit > 0 && outputCount >= limit {
				limitReached = true
				nextCell = excel.CellRef(col, row)
				continue
			}

			out := buildCellOutput(f, sheet, col, row, data, mergeInfo, noStyle, defaultFont)
			enc.Encode(out)
			outputCount++
		}
	}

	// 切り捨て通知
	if limitReached && nextCell != "" {
		trunc := excel.Truncation{
			Truncated: true,
			Total:     totalCount,
			Output:    outputCount,
			NextCell:  nextCell,
		}
		if rangeFlag != "" {
			nextCol, nextRow, _ := excel.StartPosition(nextCell)
			trunc.NextRange = excel.NextRangeFrom(nextCol, nextRow, scanRange)
		}
		enc.Encode(trunc)
	}

	return nil
}

func buildCellOutput(f *excel.File, sheet string, col, row int, data *excel.CellData, mi *excel.MergeInfo, noStyle bool, defaultFont excel.FontInfo) cellOutput {
	out := cellOutput{
		Cell: excel.CellRef(col, row),
		Type: data.Type,
	}

	switch data.Type {
	case excel.CellTypeEmpty:
		// value は省略
	case excel.CellTypeError:
		out.Value = data.Value
		out.Formula = data.Formula
	case excel.CellTypeFormula:
		out.Value = data.Value
		out.Formula = data.Formula
		out.Display = data.Display
	default:
		out.Value = data.Value
		out.Display = data.Display
	}

	// 結合セル
	if merge, ok := mi.IsTopLeft(col, row); ok {
		out.Merge = merge
	}

	// ハイパーリンク
	out.Link = f.GetHyperlink(sheet, out.Cell)

	// hidden row/col
	if f.IsHiddenRow(sheet, row) {
		out.HiddenRow = true
	}
	if f.IsHiddenCol(sheet, col) {
		out.HiddenCol = true
	}

	// 書式
	if !noStyle {
		font, fill, border, alignment, err := f.CellStyle(sheet, col, row, defaultFont)
		if err == nil {
			out.Font = font
			out.Fill = fill
			out.Border = border
			out.Alignment = alignment
		}
		out.RichText = f.GetRichText(sheet, col, row, font, defaultFont)
	}

	return out
}
