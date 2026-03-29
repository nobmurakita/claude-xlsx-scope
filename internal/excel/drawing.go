package excel

import (
	"archive/zip"
	"fmt"
	"strings"
)

// 図形タイプ（ShapeInfo.Type に出力される値）
const (
	ShapeTypeCustom    = "customShape"
	ShapeTypeGroup     = "group"
	ShapeTypePicture   = "picture"
	ShapeTypeConnector = "connector"
)

// LineStyle は図形の枠線スタイル（色・線種・太さ）
type LineStyle struct {
	Color string  `json:"color,omitempty"` // #RRGGBB
	Style string  `json:"style,omitempty"` // "solid", "dash", "dot" 等
	Width float64 `json:"width,omitempty"` // ポイント単位
}

// ImageInfo は埋め込み画像のメタデータ
type ImageInfo struct {
	Format string `json:"format"`           // 拡張子（"png", "jpg" 等）
	Width  int    `json:"width,omitempty"`  // ピクセル
	Height int    `json:"height,omitempty"` // ピクセル
	Size   int64  `json:"size,omitempty"`   // バイト数
	Path   string `json:"path,omitempty"`   // 抽出先パス
}

// ShapeInfo は Drawing XML から取得した図形情報。
// shape, connector, group, picture を統一的に表現する。
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
	return getDrawingTarget(f.zr, xmlPath) != ""
}

// DrawingOptions は LoadDrawing の動作を制御するオプション
type DrawingOptions struct {
	IncludeStyle bool   // true: fill/line/font 等の書式情報を出力に含める
	ExtractDir   string // 非空: 画像を一時ディレクトリに抽出する
}

// LoadDrawing はシートの drawing XML をパースして図形情報を返す。
func (f *File) LoadDrawing(sheet string, opts DrawingOptions) (*DrawingResult, error) {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return nil, fmt.Errorf("シート %q が見つかりません", sheet)
	}

	zr := f.zr
	target := getDrawingTarget(zr, xmlPath)
	if target == "" {
		// 図形なし
		return &DrawingResult{
			Meta: ShapesMeta{Meta: true},
		}, nil
	}

	// drawing XML パスを解決
	drawingPath := resolveRelTarget(xmlPath, target)

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

	return parseDrawingXML(entry, drawingParserConfig{
		theme:        f.getTheme(),
		includeStyle: opts.IncludeStyle,
		drawingPath:  drawingPath,
		drawingRels:  drawingRels,
		zipEntries:   zipEntries,
		extractDir:   opts.ExtractDir,
	})
}

// loadDrawingRels は drawing の .rels を読み、rId → (type, target) のマップを返す
func loadDrawingRels(zr *zip.ReadCloser, drawingPath string) map[string]xmlRelationship {
	rels := loadSheetRelsAll(zr, drawingPath)
	if len(rels) == 0 {
		return nil
	}
	m := make(map[string]xmlRelationship, len(rels))
	for _, r := range rels {
		m[r.ID] = r
	}
	return m
}

// getDrawingTarget はシートの .rels から drawing リレーションのターゲットを探す
func getDrawingTarget(zr *zip.ReadCloser, sheetXMLPath string) string {
	rels := loadSheetRelsAll(zr, sheetXMLPath)
	for _, r := range rels {
		if strings.Contains(r.Type, relKeywordDrawing) {
			return r.Target
		}
	}
	return ""
}
