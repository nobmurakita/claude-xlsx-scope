package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"strconv"
	"strings"
)

// DefaultColWidth は XML に defaultColWidth が未指定の場合に使用する
// Excel の標準デフォルト列幅（標準フォント8文字幅 + パディング）
const DefaultColWidth = 9.140625

// SheetMeta はワークシートXMLから直接パースしたシートメタデータ。
// LoadSheetMeta / LoadSheetMetaQuick で取得する。
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

// LoadSheetMetaQuick はワークシートXMLの先頭部分（sheetData の前）のみを読む軽量版。
// dimension, sheetFormatPr, cols, sheetPr を取得し、行属性・マージ・ハイパーリンクは取得しない。
func LoadSheetMetaQuick(zr *zip.ReadCloser, xmlPath string) (*SheetMeta, error) {
	entry := findZipEntry(zr, xmlPath)
	if entry == nil {
		return nil, fmt.Errorf("ZIP 内に %s が見つかりません", xmlPath)
	}
	return parseSheetMetaQuick(entry)
}

// parseSheetMetaQuick は sheetData の前の要素のみを読む軽量パーサー。
// sheetData に到達した時点で即座に返す。
func parseSheetMetaQuick(entry *zip.File) (*SheetMeta, error) {
	meta := &SheetMeta{
		Rows: make(map[int]RowInfo),
	}
	err := withZipXML(entry, func(decoder *xml.Decoder) error {
		var (
			inCols    bool
			inSheetPr bool
		)

		for {
			tok, err := decoder.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				return err
			}

			se, ok := tok.(xml.StartElement)
			if !ok {
				if ee, ok := tok.(xml.EndElement); ok {
					switch ee.Name.Local {
					case "sheetPr":
						inSheetPr = false
					case "cols":
						inCols = false
					}
				}
				continue
			}

			switch se.Name.Local {
			case "sheetData":
				// sheetData に到達したら終了
				return nil
			case "dimension":
				parseMetaDimension(se, meta)
			case "sheetPr":
				inSheetPr = true
			case "tabColor":
				if inSheetPr {
					parseTabColor(se, meta)
				}
			case "sheetFormatPr":
				parseMetaFormatPr(se, meta)
			case "cols":
				inCols = true
			case "col":
				if inCols {
					parseColInfo(se, meta)
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

// sheetMetaFullState は parseSheetMetaFull の SAX パーサー状態
type sheetMetaFullState struct {
	inSheetData  bool
	inMergeCells bool
	inHyperlinks bool
	inCols       bool
	inSheetPr    bool
	skipDepth    int // sheetData 内のネスト深さ（スキップ用）
}

// parseSheetMetaFull はワークシートXML全体をパースし、全メタデータを取得する。
func parseSheetMetaFull(entry *zip.File) (*SheetMeta, error) {
	meta := &SheetMeta{
		Rows: make(map[int]RowInfo),
	}
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
				if st.inSheetData {
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
					st.inSheetPr = true

				case "tabColor":
					if st.inSheetPr {
						parseTabColor(t, meta)
					}

				case "sheetFormatPr":
					parseMetaFormatPr(t, meta)

				case "cols":
					st.inCols = true

				case "col":
					if st.inCols {
						parseColInfo(t, meta)
					}

				case "sheetData":
					st.inSheetData = true
					st.skipDepth = 0

				case "mergeCells":
					st.inMergeCells = true

				case "mergeCell":
					if st.inMergeCells {
						for _, attr := range t.Attr {
							if attr.Name.Local == "ref" {
								meta.MergeCells = append(meta.MergeCells, MergeCellRange{Ref: attr.Value})
							}
						}
					}

				case "hyperlinks":
					st.inHyperlinks = true

				case "hyperlink":
					if st.inHyperlinks {
						parseHyperlink(t, meta)
					}
				}

			case xml.EndElement:
				if st.inSheetData {
					if t.Name.Local == "sheetData" {
						st.inSheetData = false
					} else {
						st.skipDepth--
					}
					continue
				}

				switch t.Name.Local {
				case "sheetPr":
					st.inSheetPr = false
				case "cols":
					st.inCols = false
				case "mergeCells":
					st.inMergeCells = false
				case "hyperlinks":
					st.inHyperlinks = false
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
		log.Printf("[WARN] LoadDimensionOnly: ZIPエントリ %s のオープンに失敗: %v", xmlPath, err)
		return ""
	}
	defer rc.Close()

	decoder := xml.NewDecoder(rc)
	for {
		tok, err := decoder.Token()
		if err != nil {
			if err != io.EOF {
				log.Printf("[WARN] LoadDimensionOnly: XMLトークン読み取りに失敗: %v", err)
			}
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

