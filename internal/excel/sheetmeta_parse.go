package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// DefaultColWidth は XML に defaultColWidth が未指定の場合に使用する
// Excel の標準デフォルト列幅（標準フォント8文字幅 + パディング）
const DefaultColWidth = 9.140625

// SheetMeta はワークシートXMLから直接パースしたシートメタデータ。
// StreamSheet の前に1パスで取得し、excelize の各種メタデータAPIを置き換える。
type SheetMeta struct {
	Dimension     string  // "A1:N27653"
	DefaultWidth  float64 // デフォルト列幅（XML未指定時は 0）
	DefaultHeight float64 // デフォルト行高

	// sheetPr
	TabColorRGB   string
	TabColorTheme *int
	TabColorTint  float64

	// 列情報
	Cols []ColInfo

	// 行情報（高さ・非表示）
	Rows map[int]RowInfo

	// マージセル
	MergeCells []MergeCellRange

	// ハイパーリンク（内部リンク = location のみ。外部リンクは rels で解決）
	Hyperlinks []HyperlinkEntry
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
// SAX パースで sheetData はスキップし、メタデータのみを高速に取得する。
func LoadSheetMeta(zr *zip.ReadCloser, xmlPath string) (*SheetMeta, error) {
	for _, entry := range zr.File {
		if entry.Name == xmlPath {
			return parseSheetMeta(entry)
		}
	}
	return nil, fmt.Errorf("ZIP 内に %s が見つかりません", xmlPath)
}

func parseSheetMeta(entry *zip.File) (*SheetMeta, error) {
	rc, err := entry.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	meta := &SheetMeta{
		Rows: make(map[int]RowInfo),
	}
	decoder := xml.NewDecoder(rc)

	var (
		inSheetData  bool
		inMergeCells bool
		inHyperlinks bool
		inCols       bool
		inSheetPr    bool
		skipDepth    int // sheetData 内のネスト深さ（スキップ用）
	)

	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if inSheetData {
				skipDepth++
				// sheetData 内では行の属性だけ取得する
				if t.Name.Local == "row" {
					parseRowMeta(t, meta)
				}
				continue
			}

			switch t.Name.Local {
			case "dimension":
				for _, attr := range t.Attr {
					if attr.Name.Local == "ref" {
						meta.Dimension = attr.Value
					}
				}

			case "sheetPr":
				inSheetPr = true

			case "tabColor":
				if inSheetPr {
					parseTabColor(t, meta)
				}

			case "sheetFormatPr":
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "defaultColWidth":
						meta.DefaultWidth, _ = strconv.ParseFloat(attr.Value, 64)
					case "defaultRowHeight":
						meta.DefaultHeight, _ = strconv.ParseFloat(attr.Value, 64)
					}
				}

			case "cols":
				inCols = true

			case "col":
				if inCols {
					parseColInfo(t, meta)
				}

			case "sheetData":
				inSheetData = true
				skipDepth = 0

			case "mergeCells":
				inMergeCells = true

			case "mergeCell":
				if inMergeCells {
					for _, attr := range t.Attr {
						if attr.Name.Local == "ref" {
							meta.MergeCells = append(meta.MergeCells, MergeCellRange{Ref: attr.Value})
						}
					}
				}

			case "hyperlinks":
				inHyperlinks = true

			case "hyperlink":
				if inHyperlinks {
					parseHyperlink(t, meta)
				}
			}

		case xml.EndElement:
			if inSheetData {
				if t.Name.Local == "sheetData" {
					inSheetData = false
				} else {
					skipDepth--
				}
				continue
			}

			switch t.Name.Local {
			case "sheetPr":
				inSheetPr = false
			case "cols":
				inCols = false
			case "mergeCells":
				inMergeCells = false
			case "hyperlinks":
				inHyperlinks = false
			}
		}
	}

	return meta, nil
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

// BuildMergeInfo は SheetMeta のマージセル情報から MergeInfo を構築する
func (sm *SheetMeta) BuildMergeInfo() *MergeInfo {
	mi := &MergeInfo{
		topLeft: make(map[[2]int]string, len(sm.MergeCells)),
		merged:  make(map[[2]int]bool),
	}
	for _, mc := range sm.MergeCells {
		parts := strings.SplitN(mc.Ref, ":", 2)
		if len(parts) != 2 {
			continue
		}
		sCol, sRow := parseCellRef(parts[0])
		eCol, eRow := parseCellRef(parts[1])

		mi.topLeft[[2]int{sCol, sRow}] = mc.Ref
		for r := sRow; r <= eRow; r++ {
			for c := sCol; c <= eCol; c++ {
				if r == sRow && c == sCol {
					continue
				}
				mi.merged[[2]int{c, r}] = true
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

// LoadSheetRels はシートのリレーションファイルを読み、rId → target のマップを返す。
// 主にハイパーリンクの外部URL解決に使用。
func LoadSheetRels(zr *zip.ReadCloser, sheetXMLPath string) map[string]string {
	// xl/worksheets/sheet1.xml → xl/worksheets/_rels/sheet1.xml.rels
	dir := sheetXMLPath[:strings.LastIndex(sheetXMLPath, "/")+1]
	base := sheetXMLPath[strings.LastIndex(sheetXMLPath, "/")+1:]
	relsPath := dir + "_rels/" + base + ".rels"

	data, err := readZipFileFromReader(zr, relsPath)
	if err != nil {
		return nil
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(data, &rels); err != nil {
		return nil
	}

	m := make(map[string]string, len(rels.Rels))
	for _, r := range rels.Rels {
		m[r.ID] = r.Target
	}
	return m
}

// readZipFileFromReader は zip.ReadCloser から指定パスのファイルを読む
func readZipFileFromReader(zr *zip.ReadCloser, name string) ([]byte, error) {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("not found: %s", name)
}
