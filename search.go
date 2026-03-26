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
	searchCmd.Flags().Bool("style", false, "書式情報を出力する")
	searchCmd.Flags().Bool("formula", false, "数式文字列を出力する")
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
	showStyle, _ := cmd.Flags().GetBool("style")
	limit, _ := cmd.Flags().GetInt("limit")

	if queryFlag == "" && numericFlag == "" && typeFlag == "" {
		return fmt.Errorf("--query, --numeric, --type のうち少なくとも1つを指定してください")
	}

	if rangeFlag != "" && startFlag != "" {
		return fmt.Errorf("--range と --start は同時に指定できません")
	}

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

	// 走査範囲の決定
	var scanRange *excel.CellRange
	var startCol, startRow int

	if rangeFlag != "" {
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

	showFormula, _ := cmd.Flags().GetBool("formula")

	dc, err := newDumpContext(f, sheet, !showStyle, showFormula)
	if err != nil {
		return err
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)

	outputCount := 0

	err = f.StreamRows(sheet, func(col, row int, value string) bool {
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

		if startCol > 0 {
			if row < startRow || (row == startRow && col < startCol) {
				return true
			}
		}

		if dc.mergeInfo.IsMergedNonTopLeft(col, row) {
			return true
		}

		data, err := f.ReadCell(sheet, col, row)
		if err != nil || (!data.HasValue && data.Type == excel.CellTypeEmpty) {
			return true
		}

		if !filter.MatchCell(data) {
			return true
		}

		if limit > 0 && outputCount >= limit {
			return false
		}

		out := dc.buildCellOutput(col, row, data)
		enc.Encode(out)
		outputCount++
		return true
	})
	if err != nil {
		return err
	}

	if outputCount == 0 {
		return errNoMatch
	}
	return nil
}
