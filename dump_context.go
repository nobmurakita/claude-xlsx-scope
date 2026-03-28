package main

import (
	"github.com/nobmurakita/exceldump/internal/excel"
)

// cellStyler はセルのスタイル・リッチテキスト取得機能を抽象化する
type cellStyler interface {
	StyleByID(styleID int, defaultFont excel.FontInfo) (*excel.FontObj, *excel.FillObj, *excel.BorderObj, *excel.AlignmentObj)
	GetRichText(sharedStrIdx int, cellFont *excel.FontObj, defaultFont excel.FontInfo) []excel.RichTextRun
}

// dumpContext は dump/search の走査で共有するコンテキスト
type dumpContext struct {
	styler        cellStyler
	sheet         string
	defaultFont   excel.FontInfo
	defaultHeight float64
	mergeInfo     *excel.MergeInfo
	hyperlinks    excel.HyperlinkMap
	comments      excel.CommentMap
	sheetMeta     *excel.SheetMeta // lite モード用
	showStyle     bool
	showFormula   bool
	hiddenColCache map[int]bool         // 列の非表示キャッシュ
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
		styler:         f,
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
	font, fill, border, alignment := dc.styler.StyleByID(styleID, dc.defaultFont)
	result := &styleResult{font: font, fill: fill, border: border, alignment: alignment}
	dc.styleCache[styleID] = result
	return result
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
			out.RichText = dc.styler.GetRichText(raw.SharedStrIdx, out.Font, dc.defaultFont)
		}
	}

	return out
}
