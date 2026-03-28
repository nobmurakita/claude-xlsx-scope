package main

import (
	"encoding/json"
	"fmt"
	"os"
	"strconv"

	"github.com/nobmurakita/exceldump/internal/excel"
	"github.com/spf13/cobra"
)

func init() {
	scanCmd.Flags().StringP("sheet", "s", "", "対象シート（名前 or 0始まりインデックス）")
	rootCmd.AddCommand(scanCmd)
}

var scanCmd = &cobra.Command{
	Use:   "scan <file>",
	Short: "シートの構造（データ領域の分布）を分析する",
	Args:  cobra.ExactArgs(1),
	RunE:  runScan,
}

type scanOutput struct {
	Sheet         string             `json:"sheet"`
	UsedRange     string             `json:"used_range,omitempty"`
	TabColor      string             `json:"tab_color,omitempty"`
	DefaultFont   excel.FontInfo     `json:"default_font"`
	DefaultWidth  float64            `json:"default_width"`
	DefaultHeight float64            `json:"default_height"`
	ColWidths     map[string]float64 `json:"col_widths,omitempty"`
	RowHeights    map[string]float64 `json:"row_heights,omitempty"`
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

	meta, err := f.LoadSheetMetaLite(sheet)
	if err != nil {
		return err
	}

	out := scanOutput{
		Sheet:         sheet,
		DefaultWidth:  meta.EffectiveDefaultWidth(),
		DefaultHeight: meta.DefaultHeight,
	}

	// タブ色
	out.TabColor = f.ResolveTabColor(meta)

	// デフォルトフォント
	out.DefaultFont = f.DetectDefaultFontLite()

	// 列幅（デフォルトと異なるもののみ）
	out.ColWidths = collectColWidthsFromMeta(meta, meta.EffectiveDefaultWidth())

	// used_range: dimension があればそのまま使用、なければフルスキャン
	dim := meta.Dimension
	if dim != "" && dim != "A1:A1" {
		out.UsedRange = dim
	} else {
		// dimension がないファイル: 全セル走査で used_range を算出
		rowCache := buildRowCacheFromStream(f, sheet)
		out.UsedRange = rowCache.CalcUsedRange()
	}

	// 行高（used_range 内でデフォルトと異なるもののみ）
	if out.UsedRange != "" {
		usedRange, _ := excel.ParseRange(out.UsedRange, "")
		if !usedRange.IsEmpty() {
			out.RowHeights = collectRowHeightsFromMeta(meta, usedRange, out.DefaultHeight)
		}
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)
	if err := enc.Encode(out); err != nil {
		return fmt.Errorf("JSON出力に失敗しました: %w", err)
	}
	return nil
}

// collectColWidthsFromMeta は SheetMeta の列情報からデフォルトと異なる列幅を取得する
func collectColWidthsFromMeta(meta *excel.SheetMeta, defaultWidth float64) map[string]float64 {
	widths := make(map[string]float64)
	for _, ci := range meta.Cols {
		if ci.Width != defaultWidth && ci.Width != 0 {
			for c := ci.Min; c <= ci.Max; c++ {
				widths[excel.ColName(c)] = ci.Width
			}
		}
	}
	if len(widths) == 0 {
		return nil
	}
	return widths
}

// collectRowHeightsFromMeta は SheetMeta の行情報からデフォルトと異なる行高を取得する
func collectRowHeightsFromMeta(meta *excel.SheetMeta, usedRange excel.CellRange, defaultHeight float64) map[string]float64 {
	heights := make(map[string]float64)
	for row, ri := range meta.Rows {
		if row >= usedRange.StartRow && row <= usedRange.EndRow {
			if ri.Height != defaultHeight && ri.Height != 0 {
				heights[strconv.Itoa(row)] = ri.Height
			}
		}
	}
	if len(heights) == 0 {
		return nil
	}
	return heights
}

// buildRowCacheFromStream は StreamSheet を使って RowCache を構築する
func buildRowCacheFromStream(f *excel.File, sheet string) *excel.RowCache {
	rc := excel.NewRowCache(true) // boundsOnly で十分（used_range 算出のみ）
	f.StreamSheet(sheet, false, func(raw *excel.RawCell) bool {
		rc.Add(raw.Col, raw.Row)
		return true
	})
	return rc
}
