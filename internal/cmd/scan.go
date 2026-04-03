package cmd

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

// NewScanCmd は scan サブコマンドを生成する
func NewScanCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "scan <file>",
		Short: "シートの構造（used_range）を分析する",
		Args:  cobra.ExactArgs(1),
		RunE:  runScan,
	}
	cmd.Flags().StringP("sheet", "s", "", "対象シート（名前 or 0始まりインデックス）")
	return cmd
}

type scanOutput struct {
	Sheet         string `json:"sheet"`
	UsedRange     string `json:"used_range,omitempty"`
	ValueCount    int    `json:"value_count"`
	MergedCells   int    `json:"merged_cells"`
	StyleVariants int    `json:"style_variants"`
	HasShapes     bool   `json:"has_shapes,omitempty"`
}

func runScan(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")

	f, sheet, err := openAndResolveSheet(args[0], sheetFlag)
	if err != nil {
		return err
	}
	defer f.Close()

	out := scanOutput{Sheet: sheet}
	out.HasShapes = f.HasShapes(sheet)

	// 1パスで used_range・セル数・スタイルバリエーションを取得
	visualIDs := f.VisualStyleIDs()
	result, err := f.ScanSheet(sheet, visualIDs)
	if err != nil {
		return err
	}
	out.UsedRange = result.UsedRange
	out.ValueCount = result.ValueCount
	out.MergedCells = result.MergedCells
	out.StyleVariants = result.StyleVariants

	enc := newJSONLWriter(os.Stdout)
	if err := enc.Encode(out); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}
	return nil
}
