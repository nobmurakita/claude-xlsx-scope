package excel

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// CellType はセルの値の型
type CellType string

const (
	CellTypeString  CellType = "string"
	CellTypeNumber  CellType = "number"
	CellTypeBool    CellType = "bool"
	CellTypeDate    CellType = "date"
	CellTypeFormula CellType = "formula"
	CellTypeError   CellType = "error"
	CellTypeEmpty   CellType = "empty"
)

// CellData はセルから読み取った値情報
type CellData struct {
	Type      CellType
	Value     any
	Display   string
	Formula   string
	HasValue  bool
	NumFmtID  int
	NumFmtStr string
	StyleID   int // getCellStyle でのキャッシュキーとして再利用
}

// ReadCellOpts は ReadCell のオプション
type ReadCellOpts struct {
	Value       string // StreamRows から取得済みの値
	HasValue    bool   // Value が設定済みか（true なら GetCellValue を省略）
	NeedFormula bool   // 数式文字列を取得するか
}

// ReadCell はセルの値を読み取る
func (f *File) ReadCell(sheet string, col, row int, opts ReadCellOpts) (*CellData, error) {
	axis := CellRef(col, row)

	cellType, err := f.File.GetCellType(sheet, axis)
	if err != nil {
		return nil, err
	}

	// StreamRows から値を受け取れる場合は GetCellValue を省略
	rawValue := opts.Value
	if !opts.HasValue {
		rawValue, err = f.File.GetCellValue(sheet, axis)
		if err != nil {
			return nil, err
		}
	}

	styleID, err := f.File.GetCellStyle(sheet, axis)
	if err != nil {
		return nil, err
	}

	numFmtID, numFmtStr := f.getNumFormat(styleID)

	// 数式の取得: 型が Formula の場合、または呼び出し元が要求した場合のみ
	var formula string
	if cellType == excelize.CellTypeFormula || opts.NeedFormula {
		formula, _ = f.File.GetCellFormula(sheet, axis)
	}

	data := &CellData{
		NumFmtID:  numFmtID,
		NumFmtStr: numFmtStr,
		StyleID:   styleID,
	}

	// 数式セル
	if formula != "" {
		data.Formula = formula
		data.Type = CellTypeFormula
		data.Display = rawValue
		data.HasValue = true
		data.Value = parseCachedValue(rawValue, cellType, numFmtID, sheet, f)
		adjustDisplay(data)
		return data, nil
	}

	switch cellType {
	case excelize.CellTypeSharedString, excelize.CellTypeInlineString:
		data.Type = CellTypeString
		data.Value = rawValue
		data.HasValue = rawValue != ""
		data.Display = rawValue

	case excelize.CellTypeNumber:
		data.HasValue = true
		if isDateFormat(numFmtID, numFmtStr) {
			data.Type = CellTypeDate
			data.Value = parseDate(rawValue)
			// RawCellValue モードではシリアル値が返るため、
			// Display は変換後の日付文字列から adjustDisplay に任せる
			if s, ok := data.Value.(string); ok {
				data.Display = s
			}
		} else {
			data.Type = CellTypeNumber
			data.Value = parseNumber(rawValue)
			data.Display = rawValue
		}

	case excelize.CellTypeBool:
		data.Type = CellTypeBool
		data.HasValue = true
		data.Value = rawValue == "1" || strings.EqualFold(rawValue, "true")
		if data.Value.(bool) {
			data.Display = "TRUE"
		} else {
			data.Display = "FALSE"
		}

	case excelize.CellTypeFormula:
		// GetCellFormula が空だが型がformula（キャッシュ値のみ）
		data.Type = CellTypeFormula
		data.Formula = ""
		data.Display = rawValue
		data.HasValue = true
		data.Value = parseCachedValue(rawValue, cellType, numFmtID, sheet, f)
		adjustDisplay(data)
		return data, nil

	default:
		// 空セルまたは未知の型
		if rawValue == "" {
			data.Type = CellTypeEmpty
			data.HasValue = false
			return data, nil
		}
		// CellTypeUnset でも数値+日付フォーマットの場合がある（RawCellValue モード）
		if _, err := strconv.ParseFloat(rawValue, 64); err == nil {
			if isDateFormat(numFmtID, numFmtStr) {
				data.Type = CellTypeDate
				data.HasValue = true
				data.Value = parseDate(rawValue)
				if s, ok := data.Value.(string); ok {
					data.Display = s
				}
			} else {
				data.Type = CellTypeNumber
				data.HasValue = true
				data.Value = parseNumber(rawValue)
				data.Display = rawValue
			}
		} else {
			data.Type = CellTypeString
			data.Value = rawValue
			data.HasValue = true
			data.Display = rawValue
		}
	}

	// エラー値の判定
	if data.Type == CellTypeString && isErrorValue(rawValue) {
		data.Type = CellTypeError
	}

	adjustDisplay(data)
	return data, nil
}

// adjustDisplay はdisplayをvalueのJSON表現と比較し、同一なら空にする
func adjustDisplay(data *CellData) {
	if data.Value == nil {
		data.Display = ""
		return
	}
	jsonRepr := valueToJSONString(data.Value)
	if data.Display == jsonRepr {
		data.Display = ""
	}
}

func valueToJSONString(v any) string {
	switch val := v.(type) {
	case string:
		return val
	case float64:
		if val == math.Trunc(val) && !math.IsInf(val, 0) && !math.IsNaN(val) {
			if val >= -1e15 && val <= 1e15 {
				return strconv.FormatFloat(val, 'f', -1, 64)
			}
		}
		return strconv.FormatFloat(val, 'f', -1, 64)
	case bool:
		if val {
			return "true"
		}
		return "false"
	default:
		return fmt.Sprintf("%v", v)
	}
}

func parseNumber(s string) any {
	if f, err := strconv.ParseFloat(s, 64); err == nil {
		return f
	}
	return s
}

func parseDate(rawValue string) any {
	f, err := strconv.ParseFloat(rawValue, 64)
	if err != nil {
		return rawValue
	}
	t, err := excelize.ExcelDateToTime(f, false)
	if err != nil {
		return rawValue
	}
	return formatDateTime(t, f)
}

func formatDateTime(t time.Time, serial float64) string {
	// 時刻のみ（シリアル値が1未満）
	if serial > 0 && serial < 1 {
		return t.Format("15:04:05")
	}
	// 日付のみ（時分秒が0）
	if t.Hour() == 0 && t.Minute() == 0 && t.Second() == 0 {
		return t.Format("2006-01-02")
	}
	// 日時
	return t.Format("2006-01-02T15:04:05")
}

func parseCachedValue(rawValue string, cellType excelize.CellType, numFmtID int, sheet string, f *File) any {
	if rawValue == "" {
		return nil
	}
	if isErrorValue(rawValue) {
		return rawValue
	}
	if num, err := strconv.ParseFloat(rawValue, 64); err == nil {
		if isDateFormat(numFmtID, "") {
			return parseDate(rawValue)
		}
		return num
	}
	if rawValue == "TRUE" || rawValue == "true" {
		return true
	}
	if rawValue == "FALSE" || rawValue == "false" {
		return false
	}
	return rawValue
}

func isErrorValue(s string) bool {
	switch s {
	case "#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!":
		return true
	}
	return false
}

// isDateFormat は数値フォーマットが日付系かどうかを判定する
func isDateFormat(numFmtID int, numFmtStr string) bool {
	// excelize の組み込み日付フォーマットID
	switch numFmtID {
	case 14, 15, 16, 17, 18, 19, 20, 21, 22,
		27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
		45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58:
		return true
	}
	if numFmtStr == "" {
		return false
	}
	// カスタムフォーマットの簡易判定
	lower := strings.ToLower(numFmtStr)
	// 日付系キーワードの存在をチェック（"0.00" のような数値フォーマットを除外）
	dateTokens := []string{"yy", "mm", "dd", "d", "h", "ss", "am/pm", "yyyy", "gg"}
	for _, tok := range dateTokens {
		if strings.Contains(lower, tok) {
			// "mm" は分にも使われるので、"h" や "s" と共に使われる場合は時刻と判断
			return true
		}
	}
	return false
}

func (f *File) getNumFormat(styleID int) (int, string) {
	style, err := f.File.GetStyle(styleID)
	if err != nil || style == nil {
		return 0, ""
	}
	customFmt := ""
	if style.CustomNumFmt != nil {
		customFmt = *style.CustomNumFmt
	}
	return style.NumFmt, customFmt
}

// IsHiddenRow は指定行が非表示かどうかを返す
func (f *File) IsHiddenRow(sheet string, row int) bool {
	visible, err := f.File.GetRowVisible(sheet, row)
	if err != nil {
		return false
	}
	return !visible
}

// IsHiddenCol は指定列が非表示かどうかを返す
func (f *File) IsHiddenCol(sheet string, col int) bool {
	colStr := colName(col)
	visible, err := f.File.GetColVisible(sheet, colStr)
	if err != nil {
		return false
	}
	return !visible
}

// HyperlinkData はハイパーリンク情報
type HyperlinkData struct {
	URL      string `json:"url,omitempty"`
	Location string `json:"location,omitempty"`
}

// GetHyperlink はセルのハイパーリンク情報を返す
func (f *File) GetHyperlink(sheet, axis string) *HyperlinkData {
	hasLink, target, err := f.File.GetCellHyperLink(sheet, axis)
	if err != nil || !hasLink {
		return nil
	}
	return parseHyperlinkTarget(target)
}

// HyperlinkMap はシート内の全ハイパーリンクを保持するマップ
type HyperlinkMap map[string]*HyperlinkData

// LoadHyperlinks はシートの全ハイパーリンクを一括取得する。
// セルごとの GetCellHyperLink 呼び出しを不要にする。
func (f *File) LoadHyperlinks(sheet string) HyperlinkMap {
	m := make(HyperlinkMap)
	cells, err := f.File.GetHyperLinkCells(sheet, "")
	if err != nil {
		return m
	}
	for _, axis := range cells {
		link := f.GetHyperlink(sheet, axis)
		if link != nil {
			m[axis] = link
		}
	}
	return m
}

func parseHyperlinkTarget(target string) *HyperlinkData {
	link := &HyperlinkData{}
	if strings.HasPrefix(target, "http://") || strings.HasPrefix(target, "https://") || strings.HasPrefix(target, "mailto:") {
		link.URL = target
	} else if target != "" {
		link.Location = target
	}
	return link
}
