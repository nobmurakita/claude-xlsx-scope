package excel

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"
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

// maxExactIntFloat は float64 で整数として正確に表現できる最大値。
// この値を超える場合、FormatFloat の結果が不正確になりうる。
const maxExactIntFloat = 1e15

// excelLeapYearBugSerial は Excel の 1900年うるう年バグの閾値。
// シリアル値がこの値以下の場合、基準日を 1899-12-31 にする必要がある。
// Excel は 1900-02-29 を存在する日として扱うため（実際は存在しない）。
const excelLeapYearBugSerial = 60

// 時刻計算の定数
const (
	secondsPerDay    = 86400
	secondsPerHour   = 3600
	secondsPerMinute = 60
)

// CellData はセルから読み取った値情報。
// RawCell を RawCellToCellData() で変換して得る。
type CellData struct {
	Type      CellType // セル値の型
	Value     any      // パース済みの値（string, float64, bool, nil）
	Display   string   // 表示文字列（Value の JSON 表現と同一なら空）
	Formula   string   // 数式文字列（数式セルの場合のみ）
	HasValue  bool     // true: セルに値がある
	NumFmtID  int      // 数値フォーマットID
	NumFmtStr string   // カスタム数値フォーマット文字列
	StyleID   int      // スタイルID
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
			if val >= -maxExactIntFloat && val <= maxExactIntFloat {
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
	t, err := excelDateToTime(f)
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

// excelDateToTime はExcelのシリアル値を time.Time に変換する（1900年基準）
func excelDateToTime(serial float64) (time.Time, error) {
	if serial < 0 {
		return time.Time{}, fmt.Errorf("negative serial: %f", serial)
	}
	// Excel の 1900年基準: シリアル値 1 = 1900-01-01
	// シリアル値 excelLeapYearBugSerial 以下は 1899-12-31 基準、それより上は 1899-12-30 基準
	base := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	if serial <= excelLeapYearBugSerial {
		base = time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)
	}
	days := int(serial)
	fraction := serial - float64(days)
	t := base.AddDate(0, 0, days)
	// 時刻部分（小数部）
	totalSeconds := fraction * secondsPerDay
	hours := int(totalSeconds / secondsPerHour)
	minutes := int(math.Mod(totalSeconds, secondsPerHour) / secondsPerMinute)
	seconds := int(math.Mod(totalSeconds, secondsPerMinute))
	t = t.Add(time.Duration(hours)*time.Hour + time.Duration(minutes)*time.Minute + time.Duration(seconds)*time.Second)
	return t, nil
}

// parseCachedValue は数式セルのキャッシュ値をパースする。
// エラー値・数値（日付判定含む）・ブール・文字列の順で判定する。
func parseCachedValue(rawValue string, numFmtID int, numFmtStr string) any {
	if rawValue == "" {
		return nil
	}
	if isErrorValue(rawValue) {
		return rawValue
	}
	if _, err := strconv.ParseFloat(rawValue, 64); err == nil {
		if isDateFormat(numFmtID, numFmtStr) {
			return parseDate(rawValue)
		}
		return parseNumber(rawValue)
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

// builtinDateFormatIDs は ECMA-376 で定義された組み込み日付フォーマットIDの集合。
// numFmtId がこの集合に含まれる場合、そのセルは日付型として扱う。
var builtinDateFormatIDs = map[int]bool{
	14: true, 15: true, 16: true, 17: true, 18: true, 19: true, 20: true, 21: true, 22: true,
	27: true, 28: true, 29: true, 30: true, 31: true, 32: true, 33: true, 34: true, 35: true, 36: true,
	45: true, 46: true, 47: true, 50: true, 51: true, 52: true, 53: true, 54: true, 55: true, 56: true, 57: true, 58: true,
}

// isDateFormat は数値フォーマットが日付系かどうかを判定する
func isDateFormat(numFmtID int, numFmtStr string) bool {
	if builtinDateFormatIDs[numFmtID] {
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
	if f.styles != nil {
		return f.styles.GetNumFmt(styleID)
	}
	return 0, ""
}

// HyperlinkData はハイパーリンク情報
type HyperlinkData struct {
	URL      string `json:"url,omitempty"`
	Location string `json:"location,omitempty"`
}

// HyperlinkMap はシート内の全ハイパーリンクを保持するマップ
type HyperlinkMap map[string]*HyperlinkData

func parseHyperlinkTarget(target string) *HyperlinkData {
	link := &HyperlinkData{}
	if strings.HasPrefix(target, "http://") || strings.HasPrefix(target, "https://") || strings.HasPrefix(target, "mailto:") {
		link.URL = target
	} else if target != "" {
		link.Location = target
	}
	return link
}

// RawCellToCellData は RawCell（SAXストリーミングで取得）を CellData に変換する。
// excelize への API コールを行わず、getNumFormat（キャッシュ済み）のみ使用。
func (f *File) RawCellToCellData(raw *RawCell) *CellData {
	numFmtID, numFmtStr := f.getNumFormat(raw.StyleID)

	data := &CellData{
		NumFmtID:  numFmtID,
		NumFmtStr: numFmtStr,
		StyleID:   raw.StyleID,
	}

	// 数式セル
	if raw.Formula != "" {
		data.Formula = raw.Formula
		data.Type = CellTypeFormula
		data.HasValue = true
		data.Value = parseCachedValue(raw.Value, numFmtID, numFmtStr)
		data.Display = displayFromCachedValue(data.Value, raw.Value)
		adjustDisplay(data)
		return data
	}

	switch raw.ValueType {
	case "s", "str", "inlineStr":
		data.Type = CellTypeString
		data.Value = raw.Value
		data.HasValue = raw.Value != ""
		data.Display = raw.Value

	case "b":
		data.Type = CellTypeBool
		data.HasValue = true
		data.Value = raw.Value == "1" || strings.EqualFold(raw.Value, "true")
		if data.Value.(bool) {
			data.Display = "TRUE"
		} else {
			data.Display = "FALSE"
		}

	case "e":
		data.Type = CellTypeError
		data.HasValue = true
		data.Value = raw.Value
		data.Display = raw.Value

	case "n", "":
		data.HasValue = true
		fillNumericOrDate(data, raw.Value, numFmtID, numFmtStr)

	default:
		if raw.Value == "" {
			data.Type = CellTypeEmpty
			data.HasValue = false
			return data
		}
		// 未知の型: 数値ならフォーマット判定、それ以外は文字列
		if _, err := strconv.ParseFloat(raw.Value, 64); err == nil {
			data.HasValue = true
			fillNumericOrDate(data, raw.Value, numFmtID, numFmtStr)
		} else {
			data.Type = CellTypeString
			data.Value = raw.Value
			data.HasValue = true
			data.Display = raw.Value
		}
	}

	// エラー値の判定
	if data.Type == CellTypeString && isErrorValue(raw.Value) {
		data.Type = CellTypeError
	}

	adjustDisplay(data)
	return data
}

// fillNumericOrDate は数値セルの型・値・Displayを設定する（日付フォーマットなら日付、それ以外は数値）
func fillNumericOrDate(data *CellData, rawValue string, numFmtID int, numFmtStr string) {
	if isDateFormat(numFmtID, numFmtStr) {
		data.Type = CellTypeDate
		data.Value = parseDate(rawValue)
		if s, ok := data.Value.(string); ok {
			data.Display = s
		}
	} else {
		data.Type = CellTypeNumber
		data.Value = parseNumber(rawValue)
		data.Display = rawValue
	}
}

// displayFromCachedValue はキャッシュ値から Display 文字列を決定する
func displayFromCachedValue(value any, rawValue string) string {
	if s, ok := value.(string); ok {
		return s
	}
	return rawValue
}

