package excel

import (
	"fmt"
	"io"
	"strings"
)

// 図形タイプ（ShapeInfo.Type に出力される値）
const (
	ShapeTypeCustom    = "customShape"
	ShapeTypeGroup     = "group"
	ShapeTypePicture   = "picture"
	ShapeTypeConnector = "connector"

	// VMLフォームコントロール（xl/drawings/vmlDrawing*.vml + xl/ctrlProps/ctrlProp*.xml）
	ShapeTypeCheckbox = "checkbox"
	ShapeTypeRadio    = "radio"
	ShapeTypeDrop     = "drop"
	ShapeTypeList     = "list"
	ShapeTypeSpin     = "spin"
	ShapeTypeScroll   = "scroll"
	ShapeTypeButton   = "button"
	ShapeTypeGroupBox = "gbox"
	ShapeTypeLabel    = "label"
)

// LineStyle は図形の枠線スタイル（色・線種・太さ）
type LineStyle struct {
	Color string  `json:"color,omitempty"` // #RRGGBB
	Style string  `json:"style,omitempty"` // "solid", "dash", "dot" 等
	Width float64 `json:"width,omitempty"` // ポイント単位
}

// Position は図形のポイント座標（左上原点）
type Position struct {
	X int `json:"x"`
	Y int `json:"y"`
	W int `json:"w"`
	H int `json:"h"`
}

// Point はポイント座標の点
type Point struct {
	X int `json:"x"`
	Y int `json:"y"`
}

// ShapeInfo は Drawing XML から取得した図形情報。
// shape, connector, group, picture を統一的に表現する。
type ShapeInfo struct {
	ID            int           `json:"id"`
	Type          string        `json:"type"`
	Name          string        `json:"name"`
	Text          string        `json:"text,omitempty"`
	Cell          string        `json:"cell,omitempty"`
	Pos           *Position     `json:"pos,omitempty"`
	Z             int           `json:"z"`
	Rotation      float64       `json:"rotation,omitempty"`
	Flip          string        `json:"flip,omitempty"`
	RichText      []RichTextRun `json:"rich_text,omitempty"`
	From          *int          `json:"from,omitempty"`
	To            *int          `json:"to,omitempty"`
	FromIdx       *int          `json:"from_idx,omitempty"`
	ToIdx         *int          `json:"to_idx,omitempty"`
	ConnectorType string        `json:"connector_type,omitempty"`
	Adj           map[string]int `json:"adj,omitempty"`
	Arrow         string        `json:"arrow,omitempty"`
	Start         *Point        `json:"start,omitempty"`
	End           *Point        `json:"end,omitempty"`
	Label         string        `json:"label,omitempty"`
	Children      []int         `json:"children,omitempty"`
	Parent        *int          `json:"parent,omitempty"`
	CalloutTarget *Point        `json:"callout_target,omitempty"`
	AltText       string        `json:"alt_text,omitempty"`
	ImageID       string        `json:"image_id,omitempty"`
	Fill          string        `json:"fill,omitempty"`
	Line          *LineStyle    `json:"line,omitempty"`
	Font          *FontObj      `json:"font,omitempty"`

	// VMLフォームコントロール用フィールド
	Checked       *bool  `json:"checked,omitempty"`        // checkbox, radio
	LinkedCell    string `json:"linked_cell,omitempty"`    // 値を書き出すセル参照（fmlaLink）
	ListRange     string `json:"list_range,omitempty"`     // drop, list の候補範囲（fmlaRange）
	SelectedIndex int    `json:"selected_index,omitempty"` // drop, list の選択インデックス（1始まり）
	DropLines     int    `json:"drop_lines,omitempty"`     // drop の表示行数
	SelType       string `json:"sel_type,omitempty"`       // list の選択種別（single/multi/extend）
	Min           *int   `json:"min,omitempty"`            // spin, scroll の最小値
	Max           *int   `json:"max,omitempty"`            // spin, scroll の最大値
	Val           *int   `json:"val,omitempty"`            // spin, scroll の現在値
	Inc           *int   `json:"inc,omitempty"`            // spin, scroll の増分
	Page          *int   `json:"page,omitempty"`           // scroll のページ増分
	Macro         string `json:"macro,omitempty"`          // button のマクロ名（fmlaMacro）

	// 内部フラグ（JSON 出力しない）。フォームコントロールの DrawingML 描画代替など、
	// Excel が互換用に自動生成する非表示図形を除外するために使う。
	Hidden bool `json:"-"`
}

// ShapesMeta は shapes コマンドのメタ情報行
type ShapesMeta struct {
	Meta           bool `json:"meta"`
	ShapeCount     int  `json:"shape_count"`
	ConnectorCount int  `json:"connector_count"`
}

// DrawingResult は drawing XML のパース結果
type DrawingResult struct {
	Meta   ShapesMeta
	Shapes []ShapeInfo
}

// HasShapes はシートに drawing リレーションが存在するかを返す
func (f *File) HasShapes(sheet string) bool {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return false
	}
	return getDrawingTarget(f.zi, xmlPath) != ""
}

// LoadDrawing はシートの drawing XML をパースして図形情報を返す。
// DrawingML 図形に加え、VML フォームコントロール（チェックボックス等）も
// 続きの z-order で統合して返す。
func (f *File) LoadDrawing(sheet string) (*DrawingResult, error) {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return nil, fmt.Errorf("シート %q が見つかりません", sheet)
	}

	// シートメタデータを読み込む（図形の座標計算用）
	sheetMeta, metaErr := LoadSheetMeta(f.zi, xmlPath)
	if metaErr != nil {
		sheetMeta = newSheetMeta()
	}

	result := &DrawingResult{
		Meta: ShapesMeta{Meta: true},
	}

	target := getDrawingTarget(f.zi, xmlPath)
	if target != "" {
		drawingPath := resolveRelTarget(xmlPath, target)
		drawingRels := loadDrawingRels(f.zi, drawingPath)
		entry := f.zi.lookup(drawingPath)
		if entry == nil {
			return nil, fmt.Errorf("ZIP 内に %s が見つかりません", drawingPath)
		}
		dr, err := parseDrawingXML(entry, drawingParserConfig{
			theme:        f.getTheme(),
			includeStyle: true,
			drawingPath:  drawingPath,
			drawingRels:  drawingRels,
			sheetMeta:    sheetMeta,
		})
		if err != nil {
			return nil, err
		}
		result = dr
	}

	// VML フォームコントロールを DrawingML の後ろに追加する。
	// z は DrawingML トップレベル最大値+1 から、ID は既存最大+1 から採番。
	nextZ, nextID := nextTopZ(result.Shapes), nextShapeID(result.Shapes)
	ctrlShapes, _, _ := loadFormControls(f.zi, xmlPath, sheetMeta, nextZ, nextID)
	if len(ctrlShapes) > 0 {
		result.Shapes = append(result.Shapes, ctrlShapes...)
		result.Meta.ShapeCount = len(result.Shapes)
	}

	return result, nil
}

// nextTopZ はトップレベル図形（parent なし）の最大 Z+1 を返す。
// 図形が無ければ 0。
func nextTopZ(shapes []ShapeInfo) int {
	next := 0
	for i := range shapes {
		if shapes[i].Parent != nil {
			continue
		}
		if shapes[i].Z+1 > next {
			next = shapes[i].Z + 1
		}
	}
	return next
}

// nextShapeID は既存 shapes の最大 ID+1 を返す。shapes が空の場合は 1。
func nextShapeID(shapes []ShapeInfo) int {
	next := 1
	for i := range shapes {
		if shapes[i].ID >= next {
			next = shapes[i].ID + 1
		}
	}
	return next
}

// ExtractImage は ZIP 内の画像を w に書き出す。
func (f *File) ExtractImage(mediaPath string, w io.Writer) error {
	entry := f.zi.lookup(mediaPath)
	if entry == nil {
		return fmt.Errorf("ZIP 内に %s が見つかりません", mediaPath)
	}
	rc, err := entry.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	_, err = io.Copy(w, rc)
	return err
}

// loadDrawingRels は drawing の .rels を読み、rId → (type, target) のマップを返す
func loadDrawingRels(zi *zipIndex, drawingPath string) map[string]xmlRelationship {
	rels := loadSheetRelsAll(zi, drawingPath)
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
func getDrawingTarget(zi *zipIndex, sheetXMLPath string) string {
	rels := loadSheetRelsAll(zi, sheetXMLPath)
	for _, r := range rels {
		if strings.Contains(r.Type, relKeywordDrawing) {
			return r.Target
		}
	}
	return ""
}
