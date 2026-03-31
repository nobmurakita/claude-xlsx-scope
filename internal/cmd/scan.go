package cmd

import (
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-xlsx/internal/excel"
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
	Sheet     string `json:"sheet"`
	UsedRange string `json:"used_range,omitempty"`
	HasShapes bool   `json:"has_shapes,omitempty"`
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

	// dimension があればそのまま使用、なければフルスキャン
	dim := f.LoadDimension(sheet)
	if dim != "" {
		out.UsedRange = dim
	} else {
		rc := excel.NewRowCache()
		f.StreamSheet(sheet, false, func(raw *excel.RawCell) bool {
			rc.Add(raw.Col, raw.Row)
			return true
		})
		out.UsedRange = rc.CalcUsedRange()
	}

	enc := newJSONLWriter(os.Stdout)
	if err := enc.Encode(out); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}
	return nil
}
