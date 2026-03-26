package main

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/nobmurakita/exceldump/internal/excel"
	"github.com/spf13/cobra"
)

func init() {
	searchCmd.Flags().StringP("sheet", "s", "", "対象シート（名前 or 0始まりインデックス）")
	searchCmd.Flags().String("query", "", "検索文字列（部分一致）")
	searchCmd.Flags().String("numeric", "", "数値比較（例: \">100\", \"100:200\", \"=42\"）")
	searchCmd.Flags().String("type", "", "セルの型でフィルタ（string, number, date, bool, formula）")
	searchCmd.Flags().String("range", "", "セル範囲（例: A1:H20）")
	searchCmd.Flags().String("start", "", "開始セル位置（例: A51）")
	searchCmd.Flags().Bool("no-style", false, "書式情報を省略する")
	searchCmd.Flags().Int("limit", 1000, "出力セル数の上限（0で無制限）")
	rootCmd.AddCommand(searchCmd)
}

var searchCmd = &cobra.Command{
	Use:   "search <file>",
	Short: "セル値を検索する",
	Args:  cobra.ExactArgs(1),
	RunE:  runSearch,
}

func runSearch(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")
	queryFlag, _ := cmd.Flags().GetString("query")
	numericFlag, _ := cmd.Flags().GetString("numeric")
	typeFlag, _ := cmd.Flags().GetString("type")
	rangeFlag, _ := cmd.Flags().GetString("range")
	startFlag, _ := cmd.Flags().GetString("start")
	noStyle, _ := cmd.Flags().GetBool("no-style")
	limit, _ := cmd.Flags().GetInt("limit")

	// 最低1つのフィルタが必須
	if queryFlag == "" && numericFlag == "" && typeFlag == "" {
		return fmt.Errorf("--query, --numeric, --type のうち少なくとも1つを指定してください")
	}

	if rangeFlag != "" && startFlag != "" {
		return fmt.Errorf("--range と --start は同時に指定できません")
	}

	// フィルタの構築
	filter := &excel.Filter{Query: queryFlag}
	if numericFlag != "" {
		expr, err := excel.ParseNumericExpr(numericFlag)
		if err != nil {
			return err
		}
		filter.Numeric = expr
	}
	if typeFlag != "" {
		filter.Type = excel.CellType(typeFlag)
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
			return errNoMatch
		}
		scanRange, err = excel.ParseRange(usedRangeStr, "")
		if err != nil {
			return err
		}
	}

	if scanRange.IsEmpty() {
		return errNoMatch
	}

	// デフォルトフォント
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

	for row := scanRange.StartRow; row <= scanRange.EndRow; row++ {
		for col := scanRange.StartCol; col <= scanRange.EndCol; col++ {
			if useStart {
				if row < startRow || (row == startRow && col < startCol) {
					continue
				}
			}

			if mergeInfo.IsMergedNonTopLeft(col, row) {
				continue
			}

			data, err := f.ReadCell(sheet, col, row)
			if err != nil {
				return err
			}

			if !data.HasValue && data.Type == excel.CellTypeEmpty {
				continue
			}

			// フィルタ適用
			if !filter.MatchCell(data) {
				continue
			}

			totalCount++

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

	if outputCount == 0 {
		return errNoMatch
	}
	return nil
}
