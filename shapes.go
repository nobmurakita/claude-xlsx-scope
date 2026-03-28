package main

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/nobmurakita/exceldump/internal/excel"
	"github.com/spf13/cobra"
)

func init() {
	shapesCmd.Flags().StringP("sheet", "s", "", "対象シート（名前 or 0始まりインデックス）")
	shapesCmd.Flags().Int("limit", 1000, "出力図形数の上限（0で無制限）")
	shapesCmd.Flags().Bool("style", false, "書式情報を出力する")
	rootCmd.AddCommand(shapesCmd)
}

var shapesCmd = &cobra.Command{
	Use:   "shapes <file>",
	Short: "シート上の図形をJSONL形式で出力する",
	Args:  cobra.ExactArgs(1),
	RunE:  runShapes,
}

func runShapes(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")
	limit, _ := cmd.Flags().GetInt("limit")
	showStyle, _ := cmd.Flags().GetBool("style")

	f, err := excel.OpenFile(args[0])
	if err != nil {
		return err
	}

	sheet, err := f.ResolveSheet(sheetFlag)
	if err != nil {
		return err
	}

	result, err := f.LoadDrawing(sheet, showStyle)
	if err != nil {
		return err
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)

	// _meta 行
	if err := enc.Encode(result.Meta); err != nil {
		return fmt.Errorf("JSON出力に失敗しました: %w", err)
	}

	// 図形
	for i, shape := range result.Shapes {
		if limit > 0 && i >= limit {
			break
		}
		if err := enc.Encode(shape); err != nil {
			return fmt.Errorf("JSON出力に失敗しました: %w", err)
		}
	}

	return nil
}
