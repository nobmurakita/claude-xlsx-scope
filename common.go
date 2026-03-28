package main

import (
	"github.com/nobmurakita/exceldump/internal/excel"
)

type truncatedOutput struct {
	Truncated bool   `json:"_truncated"`
	NextCell  string `json:"next_cell"`
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
	Comment   *excel.CommentData   `json:"comment,omitempty"`
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
	comments      excel.CommentMap
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
	meta, err := f.LoadSheetMeta(sheet)
	if err != nil {
		return nil, err
	}

	dc := &dumpContext{
		f:              f,
		sheet:          sheet,
		sheetMeta:      meta,
		defaultHeight:  meta.DefaultHeight,
		mergeInfo:      meta.BuildMergeInfo(),
		hyperlinks:     meta.BuildHyperlinkMap(f.LoadSheetRels(sheet)),
		comments:       f.LoadComments(sheet),
		showStyle:      showStyle,
		showFormula:    showFormula,
		hiddenColCache: make(map[int]bool),
		styleCache:     make(map[int]*styleResult),
	}

	if showStyle {
		dc.defaultFont = f.DetectDefaultFont()
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
	font, fill, border, alignment := dc.f.StyleByID(styleID, dc.defaultFont)
	result := &styleResult{font: font, fill: fill, border: border, alignment: alignment}
	dc.styleCache[styleID] = result
	return result
}

func (dc *dumpContext) buildCellOutput(col, row int, data *excel.CellData, raw *excel.RawCell) cellOutput {
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

	if dc.comments != nil {
		out.Comment = dc.comments[out.Cell]
	}

	if dc.showStyle {
		sr := dc.getCellStyleByID(data.StyleID)
		if sr != nil {
			out.Font = sr.font
			out.Fill = sr.fill
			out.Border = sr.border
			out.Alignment = sr.alignment
		}
		if raw != nil {
			out.RichText = dc.f.GetRichText(raw.SharedStrIdx, out.Font, dc.defaultFont)
		}
	}

	return out
}

// filterByRange はセルが走査範囲内かを判定する。
// skip=true: このセルをスキップ、stop=true: 走査終了
func filterByRange(col, row int, scanRange *excel.CellRange) (skip, stop bool) {
	if scanRange == nil {
		return false, false
	}
	if row < scanRange.StartRow || col < scanRange.StartCol {
		return true, false
	}
	if row > scanRange.EndRow {
		return false, true
	}
	if col > scanRange.EndCol {
		return true, false
	}
	return false, false
}

// filterByStart はセルが開始位置より前かを判定する（true=スキップ）
func filterByStart(col, row, startCol, startRow int) bool {
	if startCol > 0 {
		return row < startRow || (row == startRow && col < startCol)
	}
	return false
}
