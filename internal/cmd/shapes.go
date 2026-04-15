package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

// NewShapesCmd は shapes サブコマンドを生成する
func NewShapesCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "shapes <file>",
		Short: "シート上の図形をJSONL形式で出力する",
		Args:  cobra.ExactArgs(1),
		RunE:  runShapes,
	}
	cmd.Flags().StringP("sheet", "s", "", "対象シート（名前 or 0始まりインデックス）")
	cmd.Flags().Int("limit", defaultOutputLimit, "出力図形数の上限（0で無制限）")
	return cmd
}

func runShapes(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")
	limit, _ := cmd.Flags().GetInt("limit")

	f, sheet, err := openAndResolveSheet(args[0], sheetFlag)
	if err != nil {
		return err
	}
	defer f.Close()

	result, err := f.LoadDrawing(sheet)
	if err != nil {
		return err
	}

	ow, err := newOutputWriter(cmd)
	if err != nil {
		return err
	}
	defer ow.cleanup()

	enc := newJSONLWriter(ow)

	// _meta 行
	if err := enc.Encode(result.Meta); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}

	// 図形
	for i, shape := range result.Shapes {
		if limit > 0 && i >= limit {
			break
		}
		if err := enc.Encode(shape); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	return ow.finalize()
}
