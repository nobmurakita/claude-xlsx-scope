package main

import (
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-excel/internal/excel"
	"github.com/spf13/cobra"
)

func init() {
	scanCmd.Flags().StringP("sheet", "s", "", "対象シート（名前 or 0始まりインデックス）")
	rootCmd.AddCommand(scanCmd)
}

var scanCmd = &cobra.Command{
	Use:   "scan <file>",
	Short: "シートの構造（used_range）を分析する",
	Args:  cobra.ExactArgs(1),
	RunE:  runScan,
}

// singleCellDimension は dimension が実質的に空（データ範囲なし）であることを示す特殊値
const singleCellDimension = "A1:A1"

type scanOutput struct {
	Sheet       string `json:"sheet"`
	UsedRange   string `json:"used_range,omitempty"`
	HasDrawings bool   `json:"has_drawings,omitempty"`
}

func runScan(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")

	f, sheet, err := openAndResolveSheet(args[0], sheetFlag)
	if err != nil {
		return err
	}
	defer f.Close()

	out := scanOutput{Sheet: sheet}
	out.HasDrawings = f.HasDrawings(sheet)

	// dimension があればそのまま使用、なければフルスキャン
	dim := f.LoadDimension(sheet)
	if dim != "" && dim != singleCellDimension {
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
