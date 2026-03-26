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
	Sheet         string                   `json:"sheet"`
	UsedRange     string                   `json:"used_range,omitempty"`
	TabColor      string                   `json:"tab_color,omitempty"`
	DefaultFont   excel.FontInfo           `json:"default_font"`
	DefaultWidth  float64                  `json:"default_width"`
	DefaultHeight float64                  `json:"default_height"`
	ColWidths     map[string]float64       `json:"col_widths,omitempty"`
	RowHeights    map[string]float64       `json:"row_heights,omitempty"`
	Regions       []excel.Region           `json:"regions,omitempty"`
}

func runScan(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")

	f, err := excel.OpenFile(args[0])
	if err != nil {
		return err
	}
	defer f.Close()

	sheet, err := f.ResolveWorksheet(sheetFlag)
	if err != nil {
		return err
	}

	// dimension の有無で全行走査の要否を判定
	dim, _ := f.GetSheetDimension(sheet)
	fullScan := dim != "" && dim != "A1:A1"

	out := scanOutput{
		Sheet: sheet,
	}

	// タブ色・デフォルト幅高
	tabColor, defaultWidth, defaultHeight, err := f.GetSheetMeta(sheet)
	if err == nil {
		out.TabColor = tabColor
		out.DefaultWidth = defaultWidth
		out.DefaultHeight = defaultHeight
	}

	// デフォルトフォント
	out.DefaultFont = f.DetectDefaultFont(sheet, excel.CellRange{})

	if fullScan {
		// 全行走査: used_range, regions, col_widths, row_heights を算出
		rowCache, err := f.LoadRows(sheet)
		if err != nil {
			return err
		}

		usedRangeStr, err := f.GetUsedRange(sheet, rowCache)
		if err != nil {
			return err
		}
		out.UsedRange = usedRangeStr

		var usedRange excel.CellRange
		if usedRangeStr != "" {
			usedRange, _ = excel.ParseRange(usedRangeStr, "")
		}

		if !usedRange.IsEmpty() {
			out.ColWidths = collectColWidths(f, sheet, usedRange, out.DefaultWidth)
			out.RowHeights = collectRowHeights(f, sheet, usedRange, out.DefaultHeight)

			regions, err := f.DetectRegions(sheet, usedRange, rowCache)
			if err != nil {
				return err
			}
			out.Regions = regions
		}
	} else {
		// 軽量走査: col_widths のみ（列定義はシートXMLから取得可能）
		out.ColWidths = collectAllColWidths(f, sheet, out.DefaultWidth)
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)
	if err := enc.Encode(out); err != nil {
		return fmt.Errorf("JSON出力に失敗しました: %w", err)
	}
	return nil
}

// collectAllColWidths は used_range なしでデフォルトと異なる列幅を取得する。
// 列A から順に走査し、連続してデフォルト幅が続いたら打ち切る。
func collectAllColWidths(f *excel.File, sheet string, defaultWidth float64) map[string]float64 {
	widths := make(map[string]float64)
	consecutive := 0
	for c := 1; c <= 16384; c++ { // Excel最大列数
		colStr := excel.ColName(c)
		w, err := f.GetColWidth(sheet, colStr)
		if err != nil {
			break
		}
		if w != defaultWidth {
			widths[colStr] = w
			consecutive = 0
		} else {
			consecutive++
			if consecutive > 10 {
				break
			}
		}
	}
	if len(widths) == 0 {
		return nil
	}
	return widths
}

func collectColWidths(f *excel.File, sheet string, r excel.CellRange, defaultWidth float64) map[string]float64 {
	widths := make(map[string]float64)
	for c := r.StartCol; c <= r.EndCol; c++ {
		colStr := excel.ColName(c)
		w, err := f.GetColWidth(sheet, colStr)
		if err != nil {
			continue
		}
		if w != defaultWidth {
			widths[colStr] = w
		}
	}
	if len(widths) == 0 {
		return nil
	}
	return widths
}

func collectRowHeights(f *excel.File, sheet string, r excel.CellRange, defaultHeight float64) map[string]float64 {
	heights := make(map[string]float64)
	for row := r.StartRow; row <= r.EndRow; row++ {
		h, err := f.GetRowHeight(sheet, row)
		if err != nil {
			continue
		}
		if h != defaultHeight {
			heights[strconv.Itoa(row)] = h
		}
	}
	if len(heights) == 0 {
		return nil
	}
	return heights
}
