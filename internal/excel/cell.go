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
	CellTypeFormula CellType = "formula"
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
	Error     bool     // true: 値がExcelエラー（#N/A, #REF! 等）
	HasValue  bool     // true: セルに値がある
	NumFmtStr string   // 数値フォーマット文字列
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
// エラー値・数値・ブール・文字列の順で判定する。
func parseCachedValue(rawValue string) any {
	if rawValue == "" {
		return nil
	}
	if isErrorValue(rawValue) {
		return rawValue
	}
	if _, err := strconv.ParseFloat(rawValue, 64); err == nil {
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

func (f *File) getNumFormat(styleID int) string {
	if f.styles != nil {
		return f.styles.GetNumFmt(styleID)
	}
	return ""
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
func (f *File) RawCellToCellData(raw *RawCell) *CellData {
	numFmtStr := f.getNumFormat(raw.StyleID)

	data := &CellData{
		NumFmtStr: numFmtStr,
		StyleID:   raw.StyleID,
	}

	// 数式セル
	if raw.Formula != "" {
		data.Formula = raw.Formula
		data.Type = CellTypeFormula
		data.HasValue = true
		data.Error = raw.ValueType == vtError || isErrorValue(raw.Value)
		data.Value = parseCachedValue(raw.Value)
		data.Display = displayFromCachedValue(data.Value, raw.Value)
		adjustDisplay(data)
		return data
	}

	switch raw.ValueType {
	case vtSharedString, vtFormulaStr, vtInlineStr:
		data.Type = CellTypeString
		data.Value = raw.Value
		data.HasValue = raw.Value != ""
		data.Display = raw.Value

	case vtBool:
		data.Type = CellTypeBool
		data.HasValue = true
		data.Value = raw.Value == "1" || strings.EqualFold(raw.Value, "true")
		if data.Value.(bool) {
			data.Display = "TRUE"
		} else {
			data.Display = "FALSE"
		}

	case vtError:
		data.Type = CellTypeString
		data.HasValue = true
		data.Error = true
		data.Value = raw.Value
		data.Display = raw.Value

	case vtNumber, "":
		data.HasValue = true
		fillNumeric(data, raw.Value, numFmtStr)

	default:
		if raw.Value == "" {
			data.Type = CellTypeEmpty
			data.HasValue = false
			return data
		}
		// 未知の型: 数値ならフォーマット判定、それ以外は文字列
		if _, err := strconv.ParseFloat(raw.Value, 64); err == nil {
			data.HasValue = true
			fillNumeric(data, raw.Value, numFmtStr)
		} else {
			data.Type = CellTypeString
			data.Value = raw.Value
			data.HasValue = true
			data.Display = raw.Value
		}
	}

	adjustDisplay(data)
	return data
}

// fillNumeric は数値セルの型・値・Displayを設定する。
// フォーマットエンジンでフォーマット文字列に沿った表示文字列を生成する。
func fillNumeric(data *CellData, rawValue string, numFmtStr string) {
	data.Type = CellTypeNumber
	data.Value = parseNumber(rawValue)
	if numFmtStr != "" {
		if f, ok := data.Value.(float64); ok {
			if display := FormatNumericValue(numFmtStr, f); display != "" {
				data.Display = display
				return
			}
		}
	}
	data.Display = rawValue
}

// displayFromCachedValue はキャッシュ値から Display 文字列を決定する
func displayFromCachedValue(value any, rawValue string) string {
	if s, ok := value.(string); ok {
		return s
	}
	return rawValue
}

