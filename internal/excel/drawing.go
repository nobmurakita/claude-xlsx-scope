package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
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
	return getDrawingTarget(f.zr, xmlPath) != ""
}

// LoadDrawing はシートの drawing XML をパースして図形情報を返す。
// extractDir が空でない場合、画像を指定ディレクトリに抽出する。
func (f *File) LoadDrawing(sheet string, includeStyle bool, extractDir string) (*DrawingResult, error) {
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

	return parseDrawingXML(entry, f.theme, includeStyle, drawingPath, drawingRels, zipEntries, extractDir)
}

// loadDrawingRels は drawing の .rels を読み、rId → (type, target) のマップを返す
func loadDrawingRels(zr *zip.ReadCloser, drawingPath string) map[string]xmlRelationship {
	data, err := readZipFile(zr, relsPathFor(drawingPath))
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

// getDrawingTarget はシートの .rels から drawing リレーションのターゲットを探す
func getDrawingTarget(zr *zip.ReadCloser, sheetXMLPath string) string {
	data, err := readZipFile(zr, relsPathFor(sheetXMLPath))
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


