package main

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/nobmurakita/exceldump/internal/excel"
	"github.com/spf13/cobra"
)

func init() {
	dumpCmd.Flags().StringP("sheet", "s", "", "対象シート（名前 or 0始まりインデックス）")
	dumpCmd.Flags().String("range", "", "セル範囲（例: A1:H20, A:F, 1:20）")
	dumpCmd.Flags().String("start", "", "開始セル位置（例: A51）")
	dumpCmd.Flags().Bool("include-empty", false, "空セルも出力する")
	dumpCmd.Flags().Bool("style", false, "書式情報を出力する")
	dumpCmd.Flags().Bool("formula", false, "数式文字列を出力する")
	dumpCmd.Flags().Int("limit", 1000, "出力セル数の上限（0で無制限）")
	rootCmd.AddCommand(dumpCmd)
}

var dumpCmd = &cobra.Command{
	Use:   "dump <file>",
	Short: "セルの値と書式をJSONL形式でダンプする",
	Args:  cobra.ExactArgs(1),
	RunE:  runDump,
}

type rowOutput struct {
	Row    int     `json:"_row"`
	Height float64 `json:"height,omitempty"`
	Hidden bool    `json:"hidden,omitempty"`
}

type cellOutput struct {
	Cell      string               `json:"cell"`
	Value     any                  `json:"value,omitempty"`
	Display   string               `json:"display,omitempty"`
	Type      excel.CellType       `json:"type,omitempty"`
	Merge     string               `json:"merge,omitempty"`
	Formula   string               `json:"formula,omitempty"`
	Link      *excel.HyperlinkData `json:"link,omitempty"`
	HiddenCol bool                 `json:"hidden_col,omitempty"`
	Font      *excel.FontObj       `json:"font,omitempty"`
	Fill      *excel.FillObj       `json:"fill,omitempty"`
	Border    *excel.BorderObj     `json:"border,omitempty"`
	Alignment *excel.AlignmentObj  `json:"alignment,omitempty"`
	RichText  []excel.RichTextRun  `json:"rich_text,omitempty"`
}

// dumpContext は dump/search の走査で共有するコンテキスト
type dumpContext struct {
	f             *excel.File
	sheet         string
	defaultFont   excel.FontInfo
	defaultHeight float64
	mergeInfo     *excel.MergeInfo
	hyperlinks    excel.HyperlinkMap
	sheetMeta     *excel.SheetMeta // lite モード用
	showStyle     bool
	showFormula   bool
	hiddenColCache map[int]bool // 列の非表示キャッシュ
	styleCache     map[int]*styleResult // スタイルIDのキャッシュ
}

type styleResult struct {
	font      *excel.FontObj
	fill      *excel.FillObj
	border    *excel.BorderObj
	alignment *excel.AlignmentObj
}

func newDumpContext(f *excel.File, sheet string, showStyle, showFormula bool) (*dumpContext, error) {
	dc := &dumpContext{
		f:              f,
		sheet:          sheet,
		showStyle:      showStyle,
		showFormula:    showFormula,
		hiddenColCache: make(map[int]bool),
		styleCache:     make(map[int]*styleResult),
	}

	if f.IsLite() {
		// lite モード: SheetMeta からメタデータを取得
		meta, err := f.LoadSheetMetaLite(sheet)
		if err != nil {
			return nil, err
		}
		dc.sheetMeta = meta
		dc.defaultHeight = meta.DefaultHeight
		dc.mergeInfo = meta.BuildMergeInfo()
		dc.hyperlinks = meta.BuildHyperlinkMap(f.LoadSheetRelsLite(sheet))
		if showStyle {
			dc.defaultFont = f.DetectDefaultFontLite()
		}
	} else {
		// excelize モード
		_, _, defaultHeight, _ := f.GetSheetMeta(sheet)
		dc.defaultHeight = defaultHeight
		mergeInfo, err := f.LoadMergeInfo(sheet)
		if err != nil {
			return nil, err
		}
		dc.mergeInfo = mergeInfo
		dc.hyperlinks = f.LoadHyperlinks(sheet)
		if showStyle {
			dc.defaultFont = f.DetectDefaultFont(sheet, excel.CellRange{})
		}
	}

	return dc, nil
}

func (dc *dumpContext) isHiddenCol(col int) bool {
	if v, ok := dc.hiddenColCache[col]; ok {
		return v
	}
	hidden := false
	if dc.sheetMeta != nil {
		for _, ci := range dc.sheetMeta.Cols {
			if col >= ci.Min && col <= ci.Max {
				hidden = ci.Hidden
				break
			}
		}
	} else {
		hidden = dc.f.IsHiddenCol(dc.sheet, col)
	}
	dc.hiddenColCache[col] = hidden
	return hidden
}

func (dc *dumpContext) getCellStyleByID(styleID int) *styleResult {
	if styleID == 0 {
		return nil
	}
	if cached, ok := dc.styleCache[styleID]; ok {
		return cached
	}
	var result *styleResult
	if dc.f.IsLite() {
		font, fill, border, alignment := dc.f.StyleByIDLite(styleID, dc.defaultFont)
		result = &styleResult{font: font, fill: fill, border: border, alignment: alignment}
	} else {
		font, fill, border, alignment, _ := dc.f.StyleByID(styleID, dc.defaultFont)
		result = &styleResult{font: font, fill: fill, border: border, alignment: alignment}
	}
	dc.styleCache[styleID] = result
	return result
}

func (dc *dumpContext) emitRowInfo(enc *json.Encoder, row int) {
	ri := rowOutput{Row: row}
	if dc.sheetMeta != nil {
		if info, ok := dc.sheetMeta.Rows[row]; ok {
			if info.Height != dc.defaultHeight {
				ri.Height = info.Height
			}
			ri.Hidden = info.Hidden
		}
	} else {
		h, err := dc.f.GetRowHeight(dc.sheet, row)
		if err == nil && h != dc.defaultHeight {
			ri.Height = h
		}
		if dc.f.IsHiddenRow(dc.sheet, row) {
			ri.Hidden = true
		}
	}
	// height も hidden もデフォルトなら行情報を省略
	if ri.Height == 0 && !ri.Hidden {
		return
	}
	enc.Encode(ri)
}

func (dc *dumpContext) buildCellOutput(col, row int, data *excel.CellData) cellOutput {
	out := cellOutput{
		Cell: excel.CellRef(col, row),
	}

	switch data.Type {
	case excel.CellTypeEmpty:
		// value, type ともに省略
	case excel.CellTypeError:
		out.Type = excel.CellTypeError
		out.Value = data.Value
		if dc.showFormula {
			out.Formula = data.Formula
		}
	case excel.CellTypeDate:
		out.Type = excel.CellTypeDate
		out.Value = data.Value
		out.Display = data.Display
	case excel.CellTypeFormula:
		out.Value = data.Value
		if dc.showFormula {
			out.Formula = data.Formula
		}
		out.Display = data.Display
	default:
		out.Value = data.Value
		out.Display = data.Display
	}

	if merge, ok := dc.mergeInfo.IsTopLeft(col, row); ok {
		out.Merge = merge
	}

	out.Link = dc.hyperlinks[out.Cell]

	if dc.isHiddenCol(col) {
		out.HiddenCol = true
	}

	if dc.showStyle {
		sr := dc.getCellStyleByID(data.StyleID)
		if sr != nil {
			out.Font = sr.font
			out.Fill = sr.fill
			out.Border = sr.border
			out.Alignment = sr.alignment
		}
		if !dc.f.IsLite() {
			out.RichText = dc.f.GetRichText(dc.sheet, col, row, out.Font, dc.defaultFont)
		}
	}

	return out
}

func runDump(cmd *cobra.Command, args []string) error {
	sheetFlag, _ := cmd.Flags().GetString("sheet")
	rangeFlag, _ := cmd.Flags().GetString("range")
	startFlag, _ := cmd.Flags().GetString("start")
	includeEmpty, _ := cmd.Flags().GetBool("include-empty")
	showStyle, _ := cmd.Flags().GetBool("style")
	showFormula, _ := cmd.Flags().GetBool("formula")
	limit, _ := cmd.Flags().GetInt("limit")

	if rangeFlag != "" && startFlag != "" {
		return fmt.Errorf("--range と --start は同時に指定できません")
	}

	f, err := excel.OpenFileLite(args[0])
	if err != nil {
		return err
	}

	sheet, err := f.ResolveSheetLite(sheetFlag)
	if err != nil {
		return err
	}

	// 走査範囲の決定
	var scanRange *excel.CellRange
	var startCol, startRow int

	if rangeFlag != "" {
		r, err := excel.ParseRange(rangeFlag, "")
		if err != nil {
			return err
		}
		scanRange = &r
	} else if startFlag != "" {
		startCol, startRow, err = excel.StartPosition(startFlag)
		if err != nil {
			return err
		}
	}

	dc, err := newDumpContext(f, sheet, showStyle, showFormula)
	if err != nil {
		return err
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)

	outputCount := 0
	lastRow := -1

	err = f.StreamSheet(sheet, showFormula, func(raw *excel.RawCell) bool {
		col, row := raw.Col, raw.Row

		// --range フィルタ
		if scanRange != nil {
			if row < scanRange.StartRow || col < scanRange.StartCol {
				return true
			}
			if row > scanRange.EndRow {
				return false
			}
			if col > scanRange.EndCol {
				return true
			}
		}

		// --start フィルタ
		if startCol > 0 {
			if row < startRow || (row == startRow && col < startCol) {
				return true
			}
		}

		if dc.mergeInfo.IsMergedNonTopLeft(col, row) {
			return true
		}

		data := f.RawCellToCellData(raw)

		if !data.HasValue && data.Type == excel.CellTypeEmpty {
			if !includeEmpty {
				return true
			}
		}

		if limit > 0 && outputCount >= limit {
			return false
		}

		// 行が変わったら行情報を出力
		if row != lastRow {
			dc.emitRowInfo(enc, row)
			lastRow = row
		}

		out := dc.buildCellOutput(col, row, data)
		enc.Encode(out)
		outputCount++
		return true
	})

	return err
}
