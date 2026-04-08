package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"strconv"
	"strings"
)

// DefaultColWidth は XML に defaultColWidth が未指定の場合に使用する
// Excel の標準デフォルト列幅（標準フォント8文字幅 + パディング）
const DefaultColWidth = 9.140625

// ポイント変換係数
const (
	ColWidthPtFactor = 5.625 // Excel 列幅単位 → ポイント（標準フォント近似値）
	DefaultRowHeight = 15.0  // デフォルト行高（ポイント）
)

// SheetMeta はワークシートXMLから直接パースしたシートメタデータ。
// LoadSheetMeta で取得する。
type SheetMeta struct {
	Dimension     string  // dimension 属性（例: "A1:N27653"）
	DefaultWidth  float64 // デフォルト列幅（XML未指定時は 0、EffectiveDefaultWidth で標準値を取得）
	DefaultHeight float64 // デフォルト行高（ポイント単位）

	// sheetPr のタブ色（テーマ参照・tint 付き）
	TabColorRGB   string
	TabColorTheme *int
	TabColorTint  float64

	Cols       []ColInfo        // 列幅・非表示等の列定義
	Rows       map[int]RowInfo  // 行高・非表示（デフォルトと異なる行のみ）
	MergeCells []MergeCellRange // マージセル定義
	Hyperlinks []HyperlinkEntry // ハイパーリンク（外部リンクは rels で URL 解決）
}

// ColInfo はワークシートの列定義
type ColInfo struct {
	Min     int // 開始列番号（1始まり）
	Max     int // 終了列番号（1始まり）
	Width   float64
	StyleID int
	Hidden  bool
}

// RowInfo はワークシートの行属性
type RowInfo struct {
	Height float64
	Hidden bool
}

// MergeCellRange はマージセルの範囲定義
type MergeCellRange struct {
	Ref string // "A1:C3"
}

// HyperlinkEntry はワークシートのハイパーリンク定義
type HyperlinkEntry struct {
	Ref      string // セル参照 "A1"
	RID      string // 外部リンク用のリレーションID
	Location string // 内部リンク先
}

// LoadSheetMeta はワークシートXMLから sheetData 以外のメタデータを読み取る。
// SAX パースで sheetData 内の行属性も取得する。
func LoadSheetMeta(zr *zip.ReadCloser, xmlPath string) (*SheetMeta, error) {
	entry := findZipEntry(zr, xmlPath)
	if entry == nil {
		return nil, fmt.Errorf("ZIP 内に %s が見つかりません", xmlPath)
	}
	return parseSheetMetaFull(entry)
}

// sheetMetaSection は parseSheetMetaFull のセクション状態（相互排他）
type sheetMetaSection int

const (
	sectionNone       sheetMetaSection = iota
	sectionSheetPr                     // <sheetPr>
	sectionCols                        // <cols>
	sectionSheetData                   // <sheetData>
	sectionMergeCells                  // <mergeCells>
	sectionHyperlinks                  // <hyperlinks>
)

// sheetMetaFullState は parseSheetMetaFull の SAX パーサー状態
type sheetMetaFullState struct {
	section   sheetMetaSection
	skipDepth int // sheetData 内のネスト深さ（スキップ用）
}

func newSheetMeta() *SheetMeta {
	return &SheetMeta{
		Rows: make(map[int]RowInfo),
	}
}

// parseSheetMetaFull はワークシートXML全体をパースし、全メタデータを取得する。
func parseSheetMetaFull(entry *zip.File) (*SheetMeta, error) {
	meta := newSheetMeta()
	err := withZipXML(entry, func(decoder *xml.Decoder) error {
		var st sheetMetaFullState

		for {
			tok, err := decoder.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				return err
			}

			switch t := tok.(type) {
			case xml.StartElement:
				if st.section == sectionSheetData {
					st.skipDepth++
					// sheetData 内では行の属性だけ取得する
					if t.Name.Local == "row" {
						parseRowMeta(t, meta)
					}
					continue
				}

				switch t.Name.Local {
				case "dimension":
					parseMetaDimension(t, meta)

				case "sheetPr":
					st.section = sectionSheetPr

				case "tabColor":
					if st.section == sectionSheetPr {
						parseTabColor(t, meta)
					}

				case "sheetFormatPr":
					parseMetaFormatPr(t, meta)

				case "cols":
					st.section = sectionCols

				case "col":
					if st.section == sectionCols {
						parseColInfo(t, meta)
					}

				case "sheetData":
					st.section = sectionSheetData
					st.skipDepth = 0

				case "mergeCells":
					st.section = sectionMergeCells

				case "mergeCell":
					if st.section == sectionMergeCells {
						for _, attr := range t.Attr {
							if attr.Name.Local == "ref" {
								meta.MergeCells = append(meta.MergeCells, MergeCellRange{Ref: attr.Value})
							}
						}
					}

				case "hyperlinks":
					st.section = sectionHyperlinks

				case "hyperlink":
					if st.section == sectionHyperlinks {
						parseHyperlink(t, meta)
					}
				}

			case xml.EndElement:
				if st.section == sectionSheetData {
					if t.Name.Local == "sheetData" {
						st.section = sectionNone
					} else {
						st.skipDepth--
					}
					continue
				}

				switch t.Name.Local {
				case "sheetPr", "cols", "mergeCells", "hyperlinks":
					st.section = sectionNone
				}
			}
		}

		return nil
	})
	if err != nil {
		return nil, err
	}
	return meta, nil
}

// parseMetaDimension は dimension 要素からレンジ文字列を取得する
func parseMetaDimension(t xml.StartElement, meta *SheetMeta) {
	for _, attr := range t.Attr {
		if attr.Name.Local == "ref" {
			meta.Dimension = attr.Value
		}
	}
}

// parseMetaFormatPr は sheetFormatPr 要素からデフォルト幅・高さを取得する
func parseMetaFormatPr(t xml.StartElement, meta *SheetMeta) {
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "defaultColWidth":
			meta.DefaultWidth, _ = strconv.ParseFloat(attr.Value, 64)
		case "defaultRowHeight":
			meta.DefaultHeight, _ = strconv.ParseFloat(attr.Value, 64)
		}
	}
}

func parseRowMeta(t xml.StartElement, meta *SheetMeta) {
	var rowNum int
	var ri RowInfo
	hasCustom := false

	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "r":
			rowNum, _ = strconv.Atoi(attr.Value)
		case "ht":
			ri.Height, _ = strconv.ParseFloat(attr.Value, 64)
			hasCustom = true
		case "hidden":
			ri.Hidden = attr.Value == "1"
			hasCustom = true
		}
	}
	if rowNum > 0 && hasCustom {
		meta.Rows[rowNum] = ri
	}
}

func parseTabColor(t xml.StartElement, meta *SheetMeta) {
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "rgb":
			meta.TabColorRGB = attr.Value
		case "theme":
			v, _ := strconv.Atoi(attr.Value)
			meta.TabColorTheme = &v
		case "tint":
			meta.TabColorTint, _ = strconv.ParseFloat(attr.Value, 64)
		}
	}
}

func parseColInfo(t xml.StartElement, meta *SheetMeta) {
	ci := ColInfo{}
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "min":
			ci.Min, _ = strconv.Atoi(attr.Value)
		case "max":
			ci.Max, _ = strconv.Atoi(attr.Value)
		case "width":
			ci.Width, _ = strconv.ParseFloat(attr.Value, 64)
		case "style":
			ci.StyleID, _ = strconv.Atoi(attr.Value)
		case "hidden":
			ci.Hidden = attr.Value == "1"
		}
	}
	if ci.Min > 0 {
		meta.Cols = append(meta.Cols, ci)
	}
}

func parseHyperlink(t xml.StartElement, meta *SheetMeta) {
	hl := HyperlinkEntry{}
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "ref":
			hl.Ref = attr.Value
		case "location":
			hl.Location = attr.Value
		}
		// r:id のための名前空間付き属性
		if attr.Name.Local == "id" && strings.HasSuffix(attr.Name.Space, "relationships") {
			hl.RID = attr.Value
		}
	}
	if hl.Ref != "" {
		meta.Hyperlinks = append(meta.Hyperlinks, hl)
	}
}

// EffectiveDefaultWidth は実効デフォルト列幅を返す。
// XML に defaultColWidth がなければ Excel 標準値を返す。
func (sm *SheetMeta) EffectiveDefaultWidth() float64 {
	if sm.DefaultWidth > 0 {
		return sm.DefaultWidth
	}
	return DefaultColWidth
}

// ColWidthPt は指定列（1始まり）のポイント幅を返す。
func (sm *SheetMeta) ColWidthPt(col int) float64 {
	w := sm.EffectiveDefaultWidth()
	for _, ci := range sm.Cols {
		if col >= ci.Min && col <= ci.Max {
			if ci.Hidden {
				return 0
			}
			w = ci.Width
			break
		}
	}
	return w * ColWidthPtFactor
}

// RowHeightPt は指定行（1始まり）のポイント高さを返す。
func (sm *SheetMeta) RowHeightPt(row int) float64 {
	h := sm.DefaultHeight
	if h <= 0 {
		h = DefaultRowHeight
	}
	if ri, ok := sm.Rows[row]; ok {
		if ri.Hidden {
			return 0
		}
		h = ri.Height
	}
	return h
}

// CellOriginPt は指定セル（col, row: 1始まり）の左上ポイント座標を返す。
func (sm *SheetMeta) CellOriginPt(col, row int) (int, int) {
	var x float64
	for c := 1; c < col; c++ {
		x += sm.ColWidthPt(c)
	}
	var y float64
	for r := 1; r < row; r++ {
		y += sm.RowHeightPt(r)
	}
	return int(math.Round(x)), int(math.Round(y))
}

// LoadDimensionOnly はワークシートXMLから dimension 属性のみを高速に取得する。
// XML 先頭付近の <dimension> 要素を見つけた時点で即座に返す。
// dimension が見つからない場合やパースエラー時は空文字列を返す（警告のみ出力）。
func LoadDimensionOnly(zr *zip.ReadCloser, xmlPath string) string {
	entry := findZipEntry(zr, xmlPath)
	if entry == nil {
		return ""
	}

	rc, err := entry.Open()
	if err != nil {
		return ""
	}
	defer rc.Close()

	decoder := xml.NewDecoder(rc)
	for {
		tok, err := decoder.Token()
		if err != nil {
			return ""
		}
		if se, ok := tok.(xml.StartElement); ok {
			switch se.Name.Local {
			case "dimension":
				for _, attr := range se.Attr {
					if attr.Name.Local == "ref" {
						return attr.Value
					}
				}
				return ""
			case "sheetData":
				return ""
			}
		}
	}
}

// BuildMergeInfo は SheetMeta のマージセル情報から MergeInfo を構築する
func (sm *SheetMeta) BuildMergeInfo() *MergeInfo {
	mi := &MergeInfo{
		topLeft: make(map[cellCoord]string, len(sm.MergeCells)),
		merged:  make(map[cellCoord]bool),
	}
	for _, mc := range sm.MergeCells {
		parts := strings.SplitN(mc.Ref, ":", 2)
		if len(parts) != 2 {
			continue
		}
		sCol, sRow := parseCellRef(parts[0])
		eCol, eRow := parseCellRef(parts[1])
		if sCol == 0 || sRow == 0 || eCol == 0 || eRow == 0 {
			continue
		}

		mi.topLeft[cellCoord{sCol, sRow}] = mc.Ref
		for r := sRow; r <= eRow; r++ {
			for c := sCol; c <= eCol; c++ {
				if r == sRow && c == sCol {
					continue
				}
				mi.merged[cellCoord{c, r}] = true
			}
		}
	}
	return mi
}

// BuildHyperlinkMap は SheetMeta のハイパーリンク情報から HyperlinkMap を構築する。
// 外部リンク（RID付き）は sheetRels から URL を解決する。
func (sm *SheetMeta) BuildHyperlinkMap(sheetRels map[string]string) HyperlinkMap {
	m := make(HyperlinkMap)
	for _, hl := range sm.Hyperlinks {
		if hl.RID != "" {
			// 外部リンク: rels から URL を解決
			if target, ok := sheetRels[hl.RID]; ok {
				m[hl.Ref] = parseHyperlinkTarget(target)
			}
		} else if hl.Location != "" {
			m[hl.Ref] = &HyperlinkData{Location: hl.Location}
		}
	}
	return m
}

// LoadSheetRelsFromZip はシートのリレーションファイルを読み、rId → target のマップを返す。
// 主にハイパーリンクの外部URL解決に使用。
func LoadSheetRelsFromZip(zr *zip.ReadCloser, sheetXMLPath string) map[string]string {
	rels := loadSheetRelsAll(zr, sheetXMLPath)
	if len(rels) == 0 {
		return nil
	}
	m := make(map[string]string, len(rels))
	for _, r := range rels {
		m[r.ID] = r.Target
	}
	return m
}
