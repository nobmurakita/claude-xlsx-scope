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

type rowOutput struct {
	Row    int     `json:"_row"`
	Height float64 `json:"height,omitempty"`
	Hidden bool    `json:"hidden,omitempty"`
}

type cellOutput struct {
	Cell      string               `json:"cell"`
	Value     any                  `json:"value,omitempty"`
	Display   string               `json:"display,omitempty"`
	Type      excel.CellType       `json:"type,omitempty"`
	Merge     string               `json:"merge,omitempty"`
	Formula   string               `json:"formula,omitempty"`
	Link      *excel.HyperlinkData `json:"link,omitempty"`
	HiddenCol bool                 `json:"hidden_col,omitempty"`
	Font      *excel.FontObj       `json:"font,omitempty"`
	Fill      *excel.FillObj       `json:"fill,omitempty"`
	Border    *excel.BorderObj     `json:"border,omitempty"`
	Alignment *excel.AlignmentObj  `json:"alignment,omitempty"`
	RichText  []excel.RichTextRun  `json:"rich_text,omitempty"`
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

	// 走査範囲の決定
	var scanRange *excel.CellRange
	var startCol, startRow int

	if rangeFlag != "" {
		// --range: used_range が必要な場合は取得
		usedRange, _ := f.GetUsedRange(sheet, nil)
		r, err := excel.ParseRange(rangeFlag, usedRange)
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

	// デフォルトフォント
	var defaultFont excel.FontInfo
	if !noStyle {
		defaultFont = f.DetectDefaultFont(sheet, excel.CellRange{})
	}

	// デフォルト行高
	_, _, defaultHeight, _ := f.GetSheetMeta(sheet)

	// 結合セル情報
	mergeInfo, err := f.LoadMergeInfo(sheet)
	if err != nil {
		return err
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)

	outputCount := 0
	lastRow := -1

	err = f.StreamRows(sheet, func(col, row int, value string) bool {
		// --range フィルタ
		if scanRange != nil {
			if row < scanRange.StartRow || col < scanRange.StartCol {
				return true
			}
			if row > scanRange.EndRow {
				return false // 範囲終了、走査打ち切り
			}
			if col > scanRange.EndCol {
				return true
			}
		}

		// --start フィルタ
		if startCol > 0 {
			if row < startRow || (row == startRow && col < startCol) {
				return true
			}
		}

		// 結合セルの左上以外はスキップ
		if mergeInfo.IsMergedNonTopLeft(col, row) {
			return true
		}

		// セル値の読み取り
		data, err := f.ReadCell(sheet, col, row)
		if err != nil {
			return true
		}

		// 空セルの処理
		if !data.HasValue && data.Type == excel.CellTypeEmpty {
			if !includeEmpty {
				return true
			}
		}

		// limit チェック
		if limit > 0 && outputCount >= limit {
			return false
		}

		// 行が変わったら行情報を出力
		if row != lastRow {
			emitRowInfo(enc, f, sheet, row, defaultHeight)
			lastRow = row
		}

		out := buildCellOutput(f, sheet, col, row, data, mergeInfo, noStyle, defaultFont)
		enc.Encode(out)
		outputCount++
		return true
	})
	if err != nil {
		return err
	}

	// --include-empty で --range 指定時: ストリーミングでは空セルが来ないため、
	// 範囲内の空セルを補完する処理が必要だが、現時点では非対応
	// （include-empty + range の組み合わせは excelize の個別API呼び出しが必要）

	return nil
}

func buildCellOutput(f *excel.File, sheet string, col, row int, data *excel.CellData, mi *excel.MergeInfo, noStyle bool, defaultFont excel.FontInfo) cellOutput {
	out := cellOutput{
		Cell: excel.CellRef(col, row),
	}

	// type は date と error のみ出力（他はJSON値の型やformulaフィールドの有無から推測可能）
	switch data.Type {
	case excel.CellTypeEmpty:
		// value, type ともに省略
	case excel.CellTypeError:
		out.Type = excel.CellTypeError
		out.Value = data.Value
		out.Formula = data.Formula
	case excel.CellTypeDate:
		out.Type = excel.CellTypeDate
		out.Value = data.Value
		out.Display = data.Display
	case excel.CellTypeFormula:
		out.Value = data.Value
		out.Formula = data.Formula
		out.Display = data.Display
	default:
		out.Value = data.Value
		out.Display = data.Display
	}

	if merge, ok := mi.IsTopLeft(col, row); ok {
		out.Merge = merge
	}

	out.Link = f.GetHyperlink(sheet, out.Cell)

	if f.IsHiddenCol(sheet, col) {
		out.HiddenCol = true
	}

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

func emitRowInfo(enc *json.Encoder, f *excel.File, sheet string, row int, defaultHeight float64) {
	ri := rowOutput{Row: row}
	h, err := f.GetRowHeight(sheet, row)
	if err == nil && h != defaultHeight {
		ri.Height = h
	}
	if f.IsHiddenRow(sheet, row) {
		ri.Hidden = true
	}
	enc.Encode(ri)
}
