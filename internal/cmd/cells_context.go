package cmd

import (
	"github.com/nobmurakita/claude-xlsx-scope/internal/excel"
)

// cellStyler はセルのスタイル・リッチテキスト取得機能を抽象化する
type cellStyler interface {
	StyleByID(styleID int, defaultFont excel.FontInfo) (*excel.FontObj, *excel.FillObj, *excel.BorderObj, *excel.AlignmentObj)
	GetRichText(sharedStrIdx int, cellFont *excel.FontObj, defaultFont excel.FontInfo) []excel.RichTextRun
}

// cellsContext は cells/search の走査で共有するコンテキスト
type cellsContext struct {
	styler         cellStyler
	sheet          string
	defaultFont    excel.FontInfo
	defaultHeight  float64
	mergeInfo      *excel.MergeInfo
	hyperlinks     excel.HyperlinkMap
	comments       excel.CommentMap
	sheetMeta      *excel.SheetMeta // lite モード用
	showStyle      bool
	showFormula    bool
	hiddenColCache map[int]bool         // 列の非表示キャッシュ
	styleCache     map[int]*styleResult // スタイルIDのキャッシュ

	// スタイル参照化: ExcelのstyleID → 出力用インデックス
	styleRefMap    map[int]int
	nextStyleIdx   int
}

type styleResult struct {
	font      *excel.FontObj
	fill      *excel.FillObj
	border    *excel.BorderObj
	alignment *excel.AlignmentObj
}

func (sr *styleResult) isEmpty() bool {
	return sr.font == nil && sr.fill == nil && sr.border == nil && sr.alignment == nil
}

// styleDefOutput はスタイル定義行の出力構造体
type styleDefOutput struct {
	StyleDef  int                 `json:"_style"`
	Font      *excel.FontObj      `json:"font,omitempty"`
	Fill      *excel.FillObj      `json:"fill,omitempty"`
	Border    *excel.BorderObj    `json:"border,omitempty"`
	Alignment *excel.AlignmentObj `json:"alignment,omitempty"`
}

func newCellsContext(f *excel.File, sheet string, showStyle, showFormula bool) (*cellsContext, error) {
	meta, err := f.LoadSheetMeta(sheet)
	if err != nil {
		return nil, err
	}

	dc := &cellsContext{
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
		styleRefMap:    make(map[int]int),
	}

	if showStyle {
		dc.defaultFont = f.DetectDefaultFont()
	}

	return dc, nil
}

func (dc *cellsContext) isHiddenCol(col int) bool {
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

func (dc *cellsContext) getCellStyleByID(styleID int) *styleResult {
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

// resolveStyleRef はスタイル参照のインデックスを返す。
// 初出のスタイルの場合は定義行（styleDefOutput）も返す。
// スタイルなし、または全フィールドが空の場合は (0, nil) を返す。
func (dc *cellsContext) resolveStyleRef(styleID int) (int, *styleDefOutput) {
	sr := dc.getCellStyleByID(styleID)
	if sr == nil || sr.isEmpty() {
		return 0, nil
	}
	if idx, ok := dc.styleRefMap[styleID]; ok {
		return idx, nil
	}
	dc.nextStyleIdx++
	idx := dc.nextStyleIdx
	dc.styleRefMap[styleID] = idx
	return idx, &styleDefOutput{
		StyleDef:  idx,
		Font:      sr.font,
		Fill:      sr.fill,
		Border:    sr.border,
		Alignment: sr.alignment,
	}
}

type cellOutput struct {
	Cell      string               `json:"cell"`
	Value     any                  `json:"value,omitempty"`
	Display   string               `json:"display,omitempty"`
	Type      excel.CellType       `json:"type,omitempty"`
	Fmt       string               `json:"fmt,omitempty"`
	Error     bool                 `json:"error,omitempty"`
	Merge     string               `json:"merge,omitempty"`
	Formula   string               `json:"formula,omitempty"`
	Link      *excel.HyperlinkData `json:"link,omitempty"`
	HiddenCol bool                 `json:"hidden_col,omitempty"`
	Comment   *excel.CommentData   `json:"comment,omitempty"`
	// スタイル参照（--style 時、スタイルがある場合のみ出力）
	StyleRef  *int                 `json:"s,omitempty"`
	RichText  []excel.RichTextRun  `json:"rich_text,omitempty"`
}

// buildCellOutput はセルデータから出力構造体を生成する。
// --style 使用時、初出のスタイルがあれば styleDefOutput も返す。
func (dc *cellsContext) buildCellOutput(col, row int, data *excel.CellData, raw *excel.RawCell) (cellOutput, *styleDefOutput) {
	out := cellOutput{
		Cell: excel.CellRef(col, row),
	}

	out.Error = data.Error

	switch data.Type {
	case excel.CellTypeEmpty:
		// value, type ともに省略
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

	// 数値セルにフォーマット文字列があれば付与
	if data.Type == excel.CellTypeNumber && data.NumFmtStr != "" {
		out.Fmt = data.NumFmtStr
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

	var styleDef *styleDefOutput
	if dc.showStyle {
		idx, def := dc.resolveStyleRef(data.StyleID)
		if idx > 0 {
			out.StyleRef = &idx
			styleDef = def
		}

		// rich_text はセル固有（共有文字列依存）なのでインラインのまま
		var cellFont *excel.FontObj
		sr := dc.getCellStyleByID(data.StyleID)
		if sr != nil {
			cellFont = sr.font
		}
		if raw != nil {
			out.RichText = dc.styler.GetRichText(raw.SharedStrIdx, cellFont, dc.defaultFont)
		}
	}

	return out, styleDef
}
