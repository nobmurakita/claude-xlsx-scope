package cmd

import (
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-xlsx/internal/excel"
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
	cmd.Flags().Bool("style", false, "書式情報を出力する")
	return cmd
}

func runShapes(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")
	limit, _ := cmd.Flags().GetInt("limit")
	showStyle, _ := cmd.Flags().GetBool("style")

	f, sheet, err := openAndResolveSheet(args[0], sheetFlag)
	if err != nil {
		return err
	}
	defer f.Close()

	result, err := f.LoadDrawing(sheet, excel.DrawingOptions{
		IncludeStyle: showStyle,
	})
	if err != nil {
		return err
	}

	enc := newJSONLWriter(os.Stdout)

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

	return nil
}
