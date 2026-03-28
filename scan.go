package main

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/nobmurakita/exceldump/internal/excel"
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

type scanOutput struct {
	Sheet     string `json:"sheet"`
	UsedRange string `json:"used_range,omitempty"`
}

func runScan(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")

	f, err := excel.OpenFileLite(args[0])
	if err != nil {
		return err
	}

	sheet, err := f.ResolveSheetLite(sheetFlag)
	if err != nil {
		return err
	}

	out := scanOutput{Sheet: sheet}

	// dimension があればそのまま使用、なければフルスキャン
	dim := f.LoadDimensionLite(sheet)
	if dim != "" && dim != "A1:A1" {
		out.UsedRange = dim
	} else {
		rc := excel.NewRowCache(true)
		f.StreamSheet(sheet, false, func(raw *excel.RawCell) bool {
			rc.Add(raw.Col, raw.Row)
			return true
		})
		out.UsedRange = rc.CalcUsedRange()
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)
	if err := enc.Encode(out); err != nil {
		return fmt.Errorf("JSON出力に失敗しました: %w", err)
	}
	return nil
}
