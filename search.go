package main

import (
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
	searchCmd.Flags().Int("limit", defaultOutputLimit, "出力セル数の上限（0で無制限）")
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
	showFormula, _ := cmd.Flags().GetBool("formula")
	limit, _ := cmd.Flags().GetInt("limit")

	if queryFlag == "" && numericFlag == "" && typeFlag == "" {
		return fmt.Errorf("--query, --numeric, --type のうち少なくとも1つを指定してください")
	}

	scanRange, startCol, startRow, err := parseScanRange(rangeFlag, startFlag)
	if err != nil {
		return err
	}

	filter, err := buildFilter(queryFlag, numericFlag, typeFlag)
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

	result, err := runStream(&streamConfig{
		f:           f,
		dc:          dc,
		enc:         enc,
		scanRange:   scanRange,
		startCol:    startCol,
		startRow:    startRow,
		limit:       limit,
		showFormula: showFormula,
		filter:      filter,
	})
	if err != nil {
		return err
	}

	if err := emitTruncated(enc, result.TruncatedNext); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}
	return nil
}

// buildFilter はフラグからフィルタを構築する
func buildFilter(query, numeric, typeStr string) (*excel.Filter, error) {
	filter := &excel.Filter{Query: query}
	if numeric != "" {
		expr, err := excel.ParseNumericExpr(numeric)
		if err != nil {
			return nil, err
		}
		filter.Numeric = expr
	}
	if typeStr != "" {
		switch excel.CellType(typeStr) {
		case excel.CellTypeString, excel.CellTypeNumber, excel.CellTypeBool,
			excel.CellTypeDate, excel.CellTypeFormula:
			filter.Type = excel.CellType(typeStr)
		default:
			return nil, fmt.Errorf("不明なセル型です: %q（指定可能: string, number, date, bool, formula）", typeStr)
		}
	}
	return filter, nil
}
