package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

// LineStyle は図形の枠線情報
type LineStyle struct {
	Color string  `json:"color,omitempty"`
	Style string  `json:"style,omitempty"`
	Width float64 `json:"width,omitempty"`
}

// ImageInfo は画像のメタデータ
type ImageInfo struct {
	Format string `json:"format"`
	Width  int    `json:"width,omitempty"`
	Height int    `json:"height,omitempty"`
	Size   int64  `json:"size,omitempty"`
	Path   string `json:"path,omitempty"`
}

// ShapeInfo は出力用の図形情報
type ShapeInfo struct {
	ID            int           `json:"id"`
	Type          string        `json:"type"`
	Name          string        `json:"name"`
	Text          string        `json:"text,omitempty"`
	Cell          string        `json:"cell,omitempty"`
	Z             int           `json:"z"`
	Rotation      float64       `json:"rotation,omitempty"`
	Flip          string        `json:"flip,omitempty"`
	RichText      []RichTextRun `json:"rich_text,omitempty"`
	From          *int          `json:"from,omitempty"`
	To            *int          `json:"to,omitempty"`
	ConnectorType string        `json:"connector_type,omitempty"`
	Arrow         string        `json:"arrow,omitempty"`
	Label         string        `json:"label,omitempty"`
	Children      []int         `json:"children,omitempty"`
	Parent        *int          `json:"parent,omitempty"`
	AltText       string        `json:"alt_text,omitempty"`
	Image         *ImageInfo    `json:"image,omitempty"`
	Fill          string        `json:"fill,omitempty"`
	Line          *LineStyle    `json:"line,omitempty"`
	Font          *FontObj      `json:"font,omitempty"`
}

// ShapesMeta は shapes コマンドのメタ情報行
type ShapesMeta struct {
	Meta           bool `json:"_meta"`
	ShapeCount     int  `json:"shape_count"`
	ConnectorCount int  `json:"connector_count"`
}

// DrawingResult は drawing XML のパース結果
type DrawingResult struct {
	Meta   ShapesMeta
	Shapes []ShapeInfo
}

// HasDrawings はシートに drawing リレーションが存在するかを返す
func (f *File) HasDrawings(sheet string) bool {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return false
	}
	zr, err := zip.OpenReader(f.path)
	if err != nil {
		return false
	}
	defer zr.Close()
	return findDrawingTarget(zr, xmlPath) != ""
}

// LoadDrawing はシートの drawing XML をパースして図形情報を返す。
// extractDir が空でない場合、画像を指定ディレクトリに抽出する。
func (f *File) LoadDrawing(sheet string, includeStyle bool, extractDir string) (*DrawingResult, error) {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return nil, fmt.Errorf("シート %q が見つかりません", sheet)
	}

	zr, err := zip.OpenReader(f.path)
	if err != nil {
		return nil, err
	}
	defer zr.Close()

	target := findDrawingTarget(zr, xmlPath)
	if target == "" {
		// 図形なし
		return &DrawingResult{
			Meta: ShapesMeta{Meta: true},
		}, nil
	}

	// drawing XML パスを解決
	drawingPath := resolveDrawingPath(xmlPath, target)

	// drawing の .rels を読む（画像パス解決用）
	drawingRels := loadDrawingRels(zr, drawingPath)

	// ZIP エントリのマップを構築
	zipEntries := make(map[string]*zip.File, len(zr.File))
	for _, entry := range zr.File {
		zipEntries[entry.Name] = entry
	}

	entry, ok := zipEntries[drawingPath]
	if !ok {
		return nil, fmt.Errorf("ZIP 内に %s が見つかりません", drawingPath)
	}

	return parseDrawingXML(entry, f.theme, includeStyle, drawingPath, drawingRels, zipEntries, extractDir)
}

// loadDrawingRels は drawing の .rels を読み、rId → (type, target) のマップを返す
func loadDrawingRels(zr *zip.ReadCloser, drawingPath string) map[string]xmlRelationship {
	dir := drawingPath[:strings.LastIndex(drawingPath, "/")+1]
	base := drawingPath[strings.LastIndex(drawingPath, "/")+1:]
	relsPath := dir + "_rels/" + base + ".rels"

	data, err := readZipFileFromReader(zr, relsPath)
	if err != nil {
		return nil
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(data, &rels); err != nil {
		return nil
	}

	m := make(map[string]xmlRelationship, len(rels.Rels))
	for _, r := range rels.Rels {
		m[r.ID] = r
	}
	return m
}

// findDrawingTarget はシートの .rels から drawing リレーションのターゲットを探す
func findDrawingTarget(zr *zip.ReadCloser, sheetXMLPath string) string {
	dir := sheetXMLPath[:strings.LastIndex(sheetXMLPath, "/")+1]
	base := sheetXMLPath[strings.LastIndex(sheetXMLPath, "/")+1:]
	relsPath := dir + "_rels/" + base + ".rels"

	data, err := readZipFileFromReader(zr, relsPath)
	if err != nil {
		return ""
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(data, &rels); err != nil {
		return ""
	}

	for _, r := range rels.Rels {
		if strings.Contains(r.Type, "drawing") {
			return r.Target
		}
	}
	return ""
}

// resolveDrawingPath は drawing ターゲットを ZIP 内の絶対パスに変換する
func resolveDrawingPath(sheetXMLPath, target string) string {
	if strings.HasPrefix(target, "/") {
		return target[1:]
	}
	// 相対パス: シートのディレクトリからの相対
	dir := sheetXMLPath[:strings.LastIndex(sheetXMLPath, "/")+1]
	resolved := dir + target
	// "../drawings/drawing1.xml" のような相対パスを解決
	parts := strings.Split(resolved, "/")
	var result []string
	for _, p := range parts {
		if p == ".." {
			if len(result) > 0 {
				result = result[:len(result)-1]
			}
		} else if p != "" && p != "." {
			result = append(result, p)
		}
	}
	return strings.Join(result, "/")
}

// drawingParser は drawing XML の SAX パーサー
type drawingParser struct {
	theme        *themeColors
	includeStyle bool

	// 結果
	shapes    []ShapeInfo
	connCount int
	picCount  int

	// Excel図形ID → 連番IDのマッピング
	excelIDMap map[int]int
	nextID     int

	// アンカーの状態
	topZ int // トップレベルの z-order カウンタ

	// コネクタの接続先（後処理用）
	connRefs []connRef

	// 画像対応
	drawingPath string
	drawingRels map[string]xmlRelationship
	zipEntries  map[string]*zip.File
	extractDir  string // 空なら画像スキップ
}

type connRef struct {
	shapeIndex int // shapes スライス内のインデックス
	startID    int // Excel の接続元ID
	endID      int // Excel の接続先ID
	hasStart   bool
	hasEnd     bool
}

// グループのコンテキスト
type groupContext struct {
	seqID    int // グループの連番ID
	childZ   int // グループ内の z-order カウンタ
	children []int
}

func parseDrawingXML(entry *zip.File, theme *themeColors, includeStyle bool, drawingPath string, drawingRels map[string]xmlRelationship, zipEntries map[string]*zip.File, extractDir string) (*DrawingResult, error) {
	rc, err := entry.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	p := &drawingParser{
		theme:        theme,
		includeStyle: includeStyle,
		excelIDMap:   make(map[int]int),
		nextID:       1,
		drawingPath:  drawingPath,
		drawingRels:  drawingRels,
		zipEntries:   zipEntries,
		extractDir:   extractDir,
	}

	if err := p.parse(rc); err != nil {
		return nil, err
	}

	// コネクタの接続先を解決
	p.resolveConnectors()

	return &DrawingResult{
		Meta: ShapesMeta{
			Meta:           true,
			ShapeCount:     len(p.shapes),
			ConnectorCount: p.connCount,
		},
		Shapes: p.shapes,
	}, nil
}

func (p *drawingParser) parse(r io.Reader) error {
	decoder := xml.NewDecoder(r)

	var (
		// アンカー状態
		inAnchor     bool
		anchorType   string // "two", "one", "abs"
		anchorFromCol, anchorFromRow int
		anchorToCol, anchorToRow     int
		hasTo        bool

		// スキップ状態（pic, graphicFrame）
		skipDepth int

		// グループスタック
		groupStack []groupContext
	)

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
			if skipDepth > 0 {
				skipDepth++
				continue
			}

			switch t.Name.Local {
			case "twoCellAnchor":
				inAnchor = true
				anchorType = "two"
				anchorFromCol, anchorFromRow = 0, 0
				anchorToCol, anchorToRow = 0, 0
				hasTo = false

			case "oneCellAnchor":
				inAnchor = true
				anchorType = "one"
				anchorFromCol, anchorFromRow = 0, 0
				hasTo = false

			case "absoluteAnchor":
				inAnchor = true
				anchorType = "abs"
				hasTo = false

			case "from":
				if inAnchor && len(groupStack) == 0 {
					anchorFromCol, anchorFromRow = p.parseAnchorPos(decoder)
				} else {
					skipElement(decoder)
				}

			case "to":
				if inAnchor && len(groupStack) == 0 && anchorType == "two" {
					anchorToCol, anchorToRow = p.parseAnchorPos(decoder)
					hasTo = true
				} else {
					skipElement(decoder)
				}

			case "sp":
				if !inAnchor {
					continue
				}
				z := p.currentZ(groupStack)
				cell := p.buildCell(anchorType, anchorFromCol, anchorFromRow, anchorToCol, anchorToRow, hasTo)
				shape := p.parseShape(decoder, z, cell, groupStack)
				p.incrementZ(groupStack)
				idx := len(p.shapes)
				p.shapes = append(p.shapes, shape)
				if len(groupStack) > 0 {
					groupStack[len(groupStack)-1].children = append(groupStack[len(groupStack)-1].children, shape.ID)
				}
				_ = idx

			case "cxnSp":
				if !inAnchor {
					continue
				}
				z := p.currentZ(groupStack)
				cell := p.buildCell(anchorType, anchorFromCol, anchorFromRow, anchorToCol, anchorToRow, hasTo)
				shape := p.parseConnector(decoder, z, cell, groupStack)
				p.incrementZ(groupStack)
				p.connCount++
				p.shapes = append(p.shapes, shape)
				if len(groupStack) > 0 {
					groupStack[len(groupStack)-1].children = append(groupStack[len(groupStack)-1].children, shape.ID)
				}

			case "grpSp":
				if !inAnchor {
					continue
				}
				z := p.currentZ(groupStack)
				cell := p.buildCell(anchorType, anchorFromCol, anchorFromRow, anchorToCol, anchorToRow, hasTo)
				grpShape := p.startGroup(decoder, z, cell, groupStack)
				p.incrementZ(groupStack)
				groupStack = append(groupStack, groupContext{
					seqID: grpShape.ID,
				})
				p.shapes = append(p.shapes, grpShape)
				if len(groupStack) > 1 {
					groupStack[len(groupStack)-2].children = append(groupStack[len(groupStack)-2].children, grpShape.ID)
				}

			case "pic":
				if !inAnchor {
					continue
				}
				if p.extractDir == "" {
					// 画像抽出なし: スキップ
					skipDepth = 1
					continue
				}
				z := p.currentZ(groupStack)
				cell := p.buildCell(anchorType, anchorFromCol, anchorFromRow, anchorToCol, anchorToRow, hasTo)
				shape := p.parsePicture(decoder, z, cell, groupStack)
				p.incrementZ(groupStack)
				p.picCount++
				p.shapes = append(p.shapes, shape)
				if len(groupStack) > 0 {
					groupStack[len(groupStack)-1].children = append(groupStack[len(groupStack)-1].children, shape.ID)
				}

			case "graphicFrame":
				// スキップ対象
				skipDepth = 1

			default:
				// アンカー内のその他の要素はスキップ
			}

		case xml.EndElement:
			if skipDepth > 0 {
				skipDepth--
				continue
			}

			switch t.Name.Local {
			case "twoCellAnchor", "oneCellAnchor", "absoluteAnchor":
				inAnchor = false
				p.topZ++

			case "grpSp":
				if len(groupStack) > 0 {
					top := groupStack[len(groupStack)-1]
					groupStack = groupStack[:len(groupStack)-1]
					// グループの children を設定
					for i := range p.shapes {
						if p.shapes[i].ID == top.seqID {
							p.shapes[i].Children = top.children
							break
						}
					}
				}
			}
		}
	}

	return nil
}

func (p *drawingParser) currentZ(groupStack []groupContext) int {
	if len(groupStack) > 0 {
		return groupStack[len(groupStack)-1].childZ
	}
	return p.topZ
}

func (p *drawingParser) incrementZ(groupStack []groupContext) {
	if len(groupStack) > 0 {
		groupStack[len(groupStack)-1].childZ++
	}
}

func (p *drawingParser) buildCell(anchorType string, fromCol, fromRow, toCol, toRow int, hasTo bool) string {
	switch anchorType {
	case "two":
		if hasTo {
			from := CellRef(fromCol+1, fromRow+1)
			to := CellRef(toCol+1, toRow+1)
			if from == to {
				return from
			}
			return from + ":" + to
		}
		return CellRef(fromCol+1, fromRow+1)
	case "one":
		return CellRef(fromCol+1, fromRow+1)
	default:
		return ""
	}
}

// parseAnchorPos は <from> または <to> 内の col, row を読む
func (p *drawingParser) parseAnchorPos(decoder *xml.Decoder) (col, row int) {
	depth := 1
	var inCol, inRow bool
	var buf strings.Builder

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			return 0, 0
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "col":
				inCol = true
				buf.Reset()
			case "row":
				inRow = true
				buf.Reset()
			}
		case xml.EndElement:
			depth--
			if t.Name.Local == "col" {
				col, _ = strconv.Atoi(buf.String())
				inCol = false
			} else if t.Name.Local == "row" {
				row, _ = strconv.Atoi(buf.String())
				inRow = false
			}
		case xml.CharData:
			if inCol || inRow {
				buf.Write(t)
			}
		}
	}
	return col, row
}

// parseShape は <sp> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parseShape(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	id := p.nextID
	p.nextID++

	shape := ShapeInfo{
		ID:   id,
		Type: "customShape",
		Z:    z,
		Cell: cell,
	}

	if len(groupStack) > 0 {
		parentID := groupStack[len(groupStack)-1].seqID
		shape.Parent = &parentID
	}

	depth := 1
	var (
		inNvSpPr  bool
		inSpPr    bool
		inTxBody  bool
		inP       bool
		inR       bool
		inRPr     bool
		inDefRPr  bool
		inLn      bool
		inFill    bool // solidFill 直下
		fillCtx   string // "sp", "ln", "rPr", "defRPr"

		textParts   []string // 段落ごとのテキスト
		currentPara strings.Builder
		runs        []RichTextRun
		currentRunText strings.Builder
		currentFont    *parsedFont
		hasRuns        bool

		// スタイル
		shapeFill string
		lineStyle *LineStyle
		shapeFont *parsedFont

		excelID int
	)

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "nvSpPr":
				inNvSpPr = true
			case "cNvPr":
				if inNvSpPr {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "name":
							shape.Name = attr.Value
						case "id":
							excelID, _ = strconv.Atoi(attr.Value)
						}
					}
				}
			case "spPr":
				inSpPr = true
			case "xfrm":
				if inSpPr {
					shape.Rotation, shape.Flip = parseXfrm(t)
				}
			case "prstGeom":
				if inSpPr {
					for _, attr := range t.Attr {
						if attr.Name.Local == "prst" {
							shape.Type = attr.Value
						}
					}
				}
			case "txBody":
				inTxBody = true
				textParts = nil
				runs = nil
				hasRuns = false
			case "p":
				if inTxBody {
					inP = true
					currentPara.Reset()
				}
			case "r":
				if inP {
					inR = true
					currentRunText.Reset()
					currentFont = nil
					hasRuns = true
				}
			case "rPr":
				if inR {
					inRPr = true
					currentFont = &parsedFont{}
				}
			case "defRPr":
				if inP && inTxBody && !inR {
					inDefRPr = true
					if p.includeStyle && shapeFont == nil {
						shapeFont = &parsedFont{}
					}
				}
			case "t":
				// テキスト要素（処理は CharData で）
			case "ln":
				if inSpPr {
					inLn = true
					if p.includeStyle {
						lineStyle = &LineStyle{}
						for _, attr := range t.Attr {
							if attr.Name.Local == "w" {
								w, _ := strconv.Atoi(attr.Value)
								lineStyle.Width = math.Round(float64(w)/12700*100) / 100
							}
						}
					}
				}
			case "solidFill":
				inFill = true
				if inLn {
					fillCtx = "ln"
				} else if inRPr {
					fillCtx = "rPr"
				} else if inDefRPr {
					fillCtx = "defRPr"
				} else if inSpPr {
					fillCtx = "sp"
				} else {
					fillCtx = ""
				}
			case "srgbClr":
				if inFill {
					clr := attrVal(t, "val")
					clr = p.applyColorMods(decoder, depth, clr)
					depth-- // applyColorMods が EndElement まで消費
					p.assignColor(clr, fillCtx, &shapeFill, lineStyle, currentFont, shapeFont)
				}
			case "schemeClr":
				if inFill {
					clr := p.resolveSchemeColor(attrVal(t, "val"), decoder, depth)
					depth-- // resolveSchemeColor が EndElement まで消費
					p.assignColor(clr, fillCtx, &shapeFill, lineStyle, currentFont, shapeFont)
				}
			case "prstDash":
				if inLn && lineStyle != nil {
					lineStyle.Style = attrVal(t, "val")
				}
			case "headEnd":
				if inLn {
					headType := attrVal(t, "type")
					if headType != "" && headType != "none" {
						if shape.Arrow == "end" {
							shape.Arrow = "both"
						} else {
							shape.Arrow = "start"
						}
					}
				}
			case "tailEnd":
				if inLn {
					tailType := attrVal(t, "type")
					if tailType != "" && tailType != "none" {
						if shape.Arrow == "start" {
							shape.Arrow = "both"
						} else {
							shape.Arrow = "end"
						}
					}
				}
			// rPr / defRPr 内のフォント属性
			case "latin", "ea":
				font := currentFont
				if inDefRPr {
					font = shapeFont
				}
				if font != nil {
					if v := attrVal(t, "typeface"); v != "" {
						font.Name = v
					}
				}
			case "sz":
				// DrawingML では sz は属性ではなく rPr の属性
			}

			// rPr / defRPr の属性からフォント情報取得
			if t.Name.Local == "rPr" && inR && currentFont != nil {
				parseDrawingFontAttrs(t, currentFont)
			}
			if t.Name.Local == "defRPr" && inDefRPr && shapeFont != nil {
				parseDrawingFontAttrs(t, shapeFont)
			}

		case xml.EndElement:
			depth--
			switch t.Name.Local {
			case "nvSpPr":
				inNvSpPr = false
			case "spPr":
				inSpPr = false
			case "txBody":
				inTxBody = false
			case "p":
				if inP {
					textParts = append(textParts, currentPara.String())
					inP = false
				}
			case "r":
				if inR {
					text := currentRunText.String()
					currentPara.WriteString(text)
					run := RichTextRun{Text: text}
					if currentFont != nil && p.includeStyle {
						run.Font = richTextFontDiffFromDrawing(currentFont, p.theme)
					}
					runs = append(runs, run)
					inR = false
				}
			case "rPr":
				inRPr = false
			case "defRPr":
				inDefRPr = false
			case "ln":
				inLn = false
			case "solidFill":
				inFill = false
				fillCtx = ""
			}

		case xml.CharData:
			if inP && !inR {
				// 段落直下のテキスト（<a:t> in <a:p> without <a:r>）
				text := string(t)
				if strings.TrimSpace(text) != "" {
					currentPara.Write(t)
				}
			}
			if inR {
				currentRunText.Write(t)
			}
		}
	}

	// テキストの組み立て
	shape.Text = strings.Join(textParts, "\n")
	if shape.Text == "" {
		shape.Text = ""
	}

	// リッチテキスト
	if hasRuns && len(runs) > 1 && p.includeStyle {
		shape.RichText = runs
	}

	// スタイル
	if p.includeStyle {
		if shapeFill != "" {
			shape.Fill = shapeFill
		}
		if lineStyle != nil && (lineStyle.Color != "" || lineStyle.Style != "" || lineStyle.Width > 0) {
			if lineStyle.Style == "" && (lineStyle.Color != "" || lineStyle.Width > 0) {
				lineStyle.Style = "solid"
			}
			shape.Line = lineStyle
		}
		if shapeFont != nil {
			shape.Font = buildDrawingFontObj(shapeFont, p.theme)
		}
	}

	// Excel ID マッピング
	if excelID > 0 {
		p.excelIDMap[excelID] = id
	}

	return shape
}

// parseConnector は <cxnSp> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parseConnector(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	id := p.nextID
	p.nextID++

	shape := ShapeInfo{
		ID:   id,
		Type: "connector",
		Z:    z,
		Cell: cell,
	}

	if len(groupStack) > 0 {
		parentID := groupStack[len(groupStack)-1].seqID
		shape.Parent = &parentID
	}

	var cr connRef
	cr.shapeIndex = len(p.shapes)

	depth := 1
	var (
		inNvCxnSpPr bool
		inCxnSpPr   bool // cNvCxnSpPr
		inSpPr      bool
		inLn        bool
		inFill      bool
		fillCtx     string
		inTxBody    bool
		inP         bool
		inR         bool

		textParts   []string
		currentPara strings.Builder

		lineStyle *LineStyle
		excelID   int
	)

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "nvCxnSpPr":
				inNvCxnSpPr = true
			case "cNvPr":
				if inNvCxnSpPr {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "name":
							shape.Name = attr.Value
						case "id":
							excelID, _ = strconv.Atoi(attr.Value)
						}
					}
				}
			case "cNvCxnSpPr":
				if inNvCxnSpPr {
					inCxnSpPr = true
				}
			case "stCxn":
				if inCxnSpPr {
					if v := attrVal(t, "id"); v != "" {
						cr.startID, _ = strconv.Atoi(v)
						cr.hasStart = true
					}
				}
			case "endCxn":
				if inCxnSpPr {
					if v := attrVal(t, "id"); v != "" {
						cr.endID, _ = strconv.Atoi(v)
						cr.hasEnd = true
					}
				}
			case "spPr":
				inSpPr = true
			case "xfrm":
				if inSpPr {
					shape.Rotation, shape.Flip = parseXfrm(t)
				}
			case "prstGeom":
				if inSpPr {
					shape.ConnectorType = attrVal(t, "prst")
				}
			case "ln":
				if inSpPr {
					inLn = true
					if p.includeStyle {
						lineStyle = &LineStyle{}
						for _, attr := range t.Attr {
							if attr.Name.Local == "w" {
								w, _ := strconv.Atoi(attr.Value)
								lineStyle.Width = math.Round(float64(w)/12700*100) / 100
							}
						}
					}
				}
			case "solidFill":
				inFill = true
				if inLn {
					fillCtx = "ln"
				} else {
					fillCtx = ""
				}
			case "srgbClr":
				if inFill {
					clr := attrVal(t, "val")
					clr = p.applyColorMods(decoder, depth, clr)
					depth--
					p.assignColor(clr, fillCtx, nil, lineStyle, nil, nil)
				}
			case "schemeClr":
				if inFill {
					clr := p.resolveSchemeColor(attrVal(t, "val"), decoder, depth)
					depth--
					p.assignColor(clr, fillCtx, nil, lineStyle, nil, nil)
				}
			case "prstDash":
				if inLn && lineStyle != nil {
					lineStyle.Style = attrVal(t, "val")
				}
			case "headEnd":
				if inLn {
					headType := attrVal(t, "type")
					if headType != "" && headType != "none" {
						if shape.Arrow == "end" {
							shape.Arrow = "both"
						} else {
							shape.Arrow = "start"
						}
					}
				}
			case "tailEnd":
				if inLn {
					tailType := attrVal(t, "type")
					if tailType != "" && tailType != "none" {
						if shape.Arrow == "start" {
							shape.Arrow = "both"
						} else {
							shape.Arrow = "end"
						}
					}
				}
			case "txBody":
				inTxBody = true
				textParts = nil
			case "p":
				if inTxBody {
					inP = true
					currentPara.Reset()
				}
			case "r":
				if inP {
					inR = true
				}
			}

		case xml.EndElement:
			depth--
			switch t.Name.Local {
			case "nvCxnSpPr":
				inNvCxnSpPr = false
			case "cNvCxnSpPr":
				inCxnSpPr = false
			case "spPr":
				inSpPr = false
			case "ln":
				inLn = false
			case "solidFill":
				inFill = false
				fillCtx = ""
			case "txBody":
				inTxBody = false
			case "p":
				if inP {
					textParts = append(textParts, currentPara.String())
					inP = false
				}
			case "r":
				inR = false
			}

		case xml.CharData:
			if inR || (inP && !inR) {
				text := string(t)
				if strings.TrimSpace(text) != "" {
					currentPara.Write(t)
				}
			}
		}
	}

	// テキスト
	shape.Label = strings.Join(textParts, "\n")
	if shape.Label == "" {
		shape.Label = ""
	}

	// スタイル
	if p.includeStyle && lineStyle != nil && (lineStyle.Color != "" || lineStyle.Style != "" || lineStyle.Width > 0) {
		if lineStyle.Style == "" && (lineStyle.Color != "" || lineStyle.Width > 0) {
			lineStyle.Style = "solid"
		}
		shape.Line = lineStyle
	}

	// Excel ID マッピング
	if excelID > 0 {
		p.excelIDMap[excelID] = id
	}

	// 接続情報を記録（後処理で解決）
	if cr.hasStart || cr.hasEnd {
		cr.shapeIndex = len(p.shapes)
		p.connRefs = append(p.connRefs, cr)
	}

	return shape
}

// startGroup は <grpSp> の先頭（nvGrpSpPr, grpSpPr）を読み、ShapeInfo を返す
// grpSp の EndElement は呼び出し元で処理される
func (p *drawingParser) startGroup(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	id := p.nextID
	p.nextID++

	shape := ShapeInfo{
		ID:   id,
		Type: "group",
		Z:    z,
		Cell: cell,
	}

	if len(groupStack) > 0 {
		parentID := groupStack[len(groupStack)-1].seqID
		shape.Parent = &parentID
	}

	// nvGrpSpPr と grpSpPr を読む
	// grpSp 内の子要素はメインループで処理されるため、ここでは先頭のプロパティだけ読む
	depth := 0
	readProps := false

	for {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "cNvPr":
				if depth <= 2 {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "name":
							shape.Name = attr.Value
						case "id":
							excelID, _ := strconv.Atoi(attr.Value)
							if excelID > 0 {
								p.excelIDMap[excelID] = id
							}
						}
					}
				}
			case "xfrm":
				if depth <= 2 {
					shape.Rotation, shape.Flip = parseXfrm(t)
				}
			}
		case xml.EndElement:
			depth--
			if t.Name.Local == "grpSpPr" {
				readProps = true
			}
			if depth < 0 || readProps {
				return shape
			}
		}
	}

	return shape
}

// parsePicture は <pic> 要素を末尾まで読み、ShapeInfo を返す
func (p *drawingParser) parsePicture(decoder *xml.Decoder, z int, cell string, groupStack []groupContext) ShapeInfo {
	id := p.nextID
	p.nextID++

	shape := ShapeInfo{
		ID:   id,
		Type: "picture",
		Z:    z,
		Cell: cell,
	}

	if len(groupStack) > 0 {
		parentID := groupStack[len(groupStack)-1].seqID
		shape.Parent = &parentID
	}

	depth := 1
	var (
		inNvPicPr  bool
		inBlipFill bool
		inSpPr     bool
		embedRID   string
		excelID    int
		extCX, extCY int // EMU
	)

	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "nvPicPr":
				inNvPicPr = true
			case "cNvPr":
				if inNvPicPr {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "name":
							shape.Name = attr.Value
						case "descr":
							shape.AltText = attr.Value
						case "id":
							excelID, _ = strconv.Atoi(attr.Value)
						}
					}
				}
			case "blipFill":
				inBlipFill = true
			case "blip":
				if inBlipFill {
					for _, attr := range t.Attr {
						if attr.Name.Local == "embed" {
							embedRID = attr.Value
						}
					}
				}
			case "spPr":
				inSpPr = true
			case "xfrm":
				if inSpPr {
					shape.Rotation, shape.Flip = parseXfrm(t)
				}
			case "ext":
				if inSpPr {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "cx":
							extCX, _ = strconv.Atoi(attr.Value)
						case "cy":
							extCY, _ = strconv.Atoi(attr.Value)
						}
					}
				}
			}

		case xml.EndElement:
			depth--
			switch t.Name.Local {
			case "nvPicPr":
				inNvPicPr = false
			case "blipFill":
				inBlipFill = false
			case "spPr":
				inSpPr = false
			}
		}
	}

	// Excel ID マッピング
	if excelID > 0 {
		p.excelIDMap[excelID] = id
	}

	// 画像情報の構築と抽出
	shape.Image = p.resolveAndExtractImage(embedRID, extCX, extCY)

	return shape
}

// resolveAndExtractImage は embed RID から画像を解決し、抽出する
func (p *drawingParser) resolveAndExtractImage(embedRID string, extCX, extCY int) *ImageInfo {
	if embedRID == "" || p.drawingRels == nil {
		return nil
	}

	rel, ok := p.drawingRels[embedRID]
	if !ok {
		return nil
	}

	// 画像ファイルの ZIP パスを解決
	imagePath := resolveDrawingPath(p.drawingPath, rel.Target)

	// 拡張子から形式を判定
	ext := ""
	if dotIdx := strings.LastIndex(imagePath, "."); dotIdx >= 0 {
		ext = strings.ToLower(imagePath[dotIdx+1:])
	}

	info := &ImageInfo{
		Format: ext,
	}

	// EMU → ピクセル変換（1px = 9525 EMU）
	if extCX > 0 {
		info.Width = int(math.Round(float64(extCX) / 9525))
	}
	if extCY > 0 {
		info.Height = int(math.Round(float64(extCY) / 9525))
	}

	// ZIP エントリからファイルサイズを取得
	zipEntry, ok := p.zipEntries[imagePath]
	if !ok {
		return info
	}
	info.Size = int64(zipEntry.UncompressedSize64)

	// 画像を抽出
	if p.extractDir != "" {
		outPath := p.extractImage(zipEntry, ext)
		if outPath != "" {
			info.Path = outPath
		}
	}

	return info
}

// extractImage は ZIP エントリからファイルを抽出する
func (p *drawingParser) extractImage(entry *zip.File, ext string) string {
	rc, err := entry.Open()
	if err != nil {
		return ""
	}
	defer rc.Close()

	// ファイル名: image_1.png, image_2.jpg, ...
	filename := fmt.Sprintf("image_%d.%s", p.picCount+1, ext)
	outPath := filepath.Join(p.extractDir, filename)

	outFile, err := os.Create(outPath)
	if err != nil {
		return ""
	}
	defer outFile.Close()

	if _, err := io.Copy(outFile, rc); err != nil {
		os.Remove(outPath)
		return ""
	}

	return outPath
}

// resolveConnectors はコネクタの from/to を Excel ID から連番 ID に解決する
func (p *drawingParser) resolveConnectors() {
	for _, cr := range p.connRefs {
		if cr.shapeIndex >= len(p.shapes) {
			continue
		}
		if cr.hasStart {
			if seqID, ok := p.excelIDMap[cr.startID]; ok {
				p.shapes[cr.shapeIndex].From = &seqID
			}
		}
		if cr.hasEnd {
			if seqID, ok := p.excelIDMap[cr.endID]; ok {
				p.shapes[cr.shapeIndex].To = &seqID
			}
		}
	}
}

// parseXfrm は xfrm の属性から回転と反転を取得する
func parseXfrm(t xml.StartElement) (rotation float64, flip string) {
	var flipH, flipV bool
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "rot":
			rot, _ := strconv.Atoi(attr.Value)
			rotation = math.Round(float64(rot)/60000*100) / 100
		case "flipH":
			flipH = attr.Value == "1"
		case "flipV":
			flipV = attr.Value == "1"
		}
	}
	if flipH && flipV {
		flip = "hv"
	} else if flipH {
		flip = "h"
	} else if flipV {
		flip = "v"
	}
	return rotation, flip
}

// schemeColorIndex はスキームカラー名をテーマインデックスにマッピングする
var schemeColorIndex = map[string]int{
	"dk1":      0,
	"lt1":      1,
	"dk2":      2,
	"lt2":      3,
	"accent1":  4,
	"accent2":  5,
	"accent3":  6,
	"accent4":  7,
	"accent5":  8,
	"accent6":  9,
	"hlink":    10,
	"folHlink": 11,
}

// resolveSchemeColor はスキームカラーを解決し、子の色変換要素まで消費する
func (p *drawingParser) resolveSchemeColor(scheme string, decoder *xml.Decoder, startDepth int) string {
	idx, ok := schemeColorIndex[scheme]
	base := ""
	if ok && p.theme != nil {
		base = p.theme.Get(idx)
	}

	// 子要素から lumMod, lumOff, tint, shade を収集
	var lumMod, lumOff float64
	lumMod = 1.0 // デフォルト
	var tint float64
	hasTint := false

	depth := 1
	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "lumMod":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					lumMod = float64(n) / 100000.0
				}
			case "lumOff":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					lumOff = float64(n) / 100000.0
				}
			case "tint":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					tint = float64(n) / 100000.0
					hasTint = true
				}
			case "shade":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					tint = -(1.0 - float64(n)/100000.0)
					hasTint = true
				}
			}
		case xml.EndElement:
			depth--
		}
	}

	if base == "" {
		return ""
	}

	// tint を適用
	if hasTint {
		return applyTint(base, tint)
	}

	// lumMod/lumOff を適用
	if lumMod != 1.0 || lumOff != 0 {
		return applyLuminance(base, lumMod, lumOff)
	}

	return base
}

// applyColorMods は srgbClr の子要素（alpha 等）を消費し、色を返す
func (p *drawingParser) applyColorMods(decoder *xml.Decoder, startDepth int, color string) string {
	clr := normalizeHexColor(color)

	var lumMod, lumOff float64
	lumMod = 1.0
	var tint float64
	hasTint := false

	depth := 1
	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "lumMod":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					lumMod = float64(n) / 100000.0
				}
			case "lumOff":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					lumOff = float64(n) / 100000.0
				}
			case "tint":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					tint = float64(n) / 100000.0
					hasTint = true
				}
			case "shade":
				if v := attrVal(t, "val"); v != "" {
					n, _ := strconv.Atoi(v)
					tint = -(1.0 - float64(n)/100000.0)
					hasTint = true
				}
			}
		case xml.EndElement:
			depth--
		}
	}

	if hasTint {
		return applyTint(clr, tint)
	}
	if lumMod != 1.0 || lumOff != 0 {
		return applyLuminance(clr, lumMod, lumOff)
	}
	return clr
}

// applyLuminance は lumMod/lumOff を適用する
func applyLuminance(hex string, lumMod, lumOff float64) string {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return "#" + strings.ToUpper(hex)
	}
	r, _ := strconv.ParseInt(hex[0:2], 16, 32)
	g, _ := strconv.ParseInt(hex[2:4], 16, 32)
	b, _ := strconv.ParseInt(hex[4:6], 16, 32)

	// HSL に変換して luminance を調整
	h, s, l := rgbToHSL(float64(r)/255, float64(g)/255, float64(b)/255)
	l = l*lumMod + lumOff
	if l < 0 {
		l = 0
	}
	if l > 1 {
		l = 1
	}
	rr, gg, bb := hslToRGB(h, s, l)
	return fmt.Sprintf("#%02X%02X%02X", int(math.Round(rr*255)), int(math.Round(gg*255)), int(math.Round(bb*255)))
}

// assignColor は解決済み色を適切なターゲットに割り当てる
func (p *drawingParser) assignColor(color, ctx string, shapeFill *string, lineStyle *LineStyle, runFont, defFont *parsedFont) {
	if color == "" {
		return
	}
	switch ctx {
	case "sp":
		if shapeFill != nil {
			*shapeFill = color
		}
	case "ln":
		if lineStyle != nil {
			lineStyle.Color = color
		}
	case "rPr":
		if runFont != nil {
			runFont.Color = color
		}
	case "defRPr":
		if defFont != nil {
			defFont.Color = color
		}
	}
}

// attrVal は StartElement から指定属性の値を返す
func attrVal(t xml.StartElement, name string) string {
	for _, attr := range t.Attr {
		if attr.Name.Local == name {
			return attr.Value
		}
	}
	return ""
}

// skipElement は現在の要素を末尾まで読み飛ばす
func skipElement(decoder *xml.Decoder) {
	depth := 1
	for depth > 0 {
		tok, err := decoder.Token()
		if err != nil {
			return
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
}

// parseDrawingFontAttrs は DrawingML の rPr/defRPr 属性からフォント情報を取得する
func parseDrawingFontAttrs(t xml.StartElement, font *parsedFont) {
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "sz":
			// 100分の1ポイント単位
			sz, _ := strconv.Atoi(attr.Value)
			font.Size = float64(sz) / 100
		case "b":
			font.Bold = attr.Value == "1"
		case "i":
			font.Italic = attr.Value == "1"
		case "strike":
			if attr.Value != "" && attr.Value != "noStrike" {
				font.Strike = true
			}
		case "u":
			if attr.Value != "" && attr.Value != "none" {
				font.Underline = attr.Value
			}
		}
	}
}

// buildDrawingFontObj は DrawingML の parsedFont から FontObj を構築する
func buildDrawingFontObj(font *parsedFont, theme *themeColors) *FontObj {
	if font == nil {
		return nil
	}
	obj := &FontObj{
		Name:          font.Name,
		Bold:          font.Bold,
		Italic:        font.Italic,
		Strikethrough: font.Strike,
		Underline:     font.Underline,
	}
	if font.Size != 0 {
		obj.Size = font.Size
	}
	if font.Color != "" {
		obj.Color = font.Color
	} else if font.ColorTheme != nil {
		color := resolveColorLite("", font.ColorTheme, font.ColorTint, theme)
		if color != "" && color != "#000000" {
			obj.Color = color
		}
	}
	if obj.IsEmpty() {
		return nil
	}
	return obj
}

// richTextFontDiffFromDrawing は DrawingML の parsedFont から差分フォントを構築する
func richTextFontDiffFromDrawing(font *parsedFont, theme *themeColors) *FontObj {
	if font == nil {
		return nil
	}
	obj := &FontObj{}
	if font.Name != "" {
		obj.Name = font.Name
	}
	if font.Size != 0 {
		obj.Size = font.Size
	}
	obj.Bold = font.Bold
	obj.Italic = font.Italic
	obj.Strikethrough = font.Strike
	obj.Underline = font.Underline

	if font.Color != "" {
		obj.Color = font.Color
	} else if font.ColorTheme != nil {
		color := resolveColorLite("", font.ColorTheme, font.ColorTint, theme)
		if color != "" && color != "#000000" {
			obj.Color = color
		}
	}
	if obj.IsEmpty() {
		return nil
	}
	return obj
}

// rgbToHSL, hslToRGB, hueToRGB は styles_parse.go で定義済み
