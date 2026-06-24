package excel

import (
	"encoding/xml"
	"strings"
)

// drawingStyleHandler は shape / connector パーサーで共通のスタイル処理を担う
type drawingStyleHandler struct {
	p         *drawingParser
	inLn      bool
	inFill    bool
	fillCtx   string // "sp", "ln", "rPr", "defRPr"
	shapeFill string
	lineStyle *LineStyle

	// <xdr:style> の参照（spPr に明示指定が無い場合のフォールバック）
	inFillRef      bool   // fillRef 要素内（idx="0" の noFill 時は false のまま）
	inLnRef        bool   // lnRef 要素内（idx="0" の noLine 時は false のまま）
	fillRefIdx     int    // fillRef の idx（テーマ fillStyleLst の色変換参照に使用）
	lnRefIdx       int    // lnRef の idx
	styleFill      string // fillRef から解決した塗り色（テーマ色変換適用後）
	styleLineColor string // lnRef から解決した線色（テーマ色変換適用後）
	spFillNone     bool   // spPr に明示の <a:noFill/>（塗りつぶしなし）。style fillRef を抑止する
	lnFillNone     bool   // ln 内に明示の <a:noFill/>（線なし）。style lnRef を抑止する
}

// handleStartElement はスタイル関連の StartElement を処理する。
// 処理した場合は true を返し、depth の調整値を返す（通常 0、色変換で要素を消費した場合は -1）。
func (h *drawingStyleHandler) handleStartElement(t xml.StartElement, decoder *xml.Decoder, inSpPr, inRPr, inDefRPr bool, currentFont, shapeFont *parsedFont) (handled bool, depthAdj int) {
	switch t.Name.Local {
	case "ln":
		if inSpPr {
			h.inLn = true
			if h.p.includeStyle {
				h.lineStyle = parseLineWidth(t)
			}
			return true, 0
		}
	case "solidFill":
		h.inFill = true
		h.fillCtx = determineFillCtx(h.inLn, inRPr, inDefRPr, inSpPr)
		return true, 0
	case "noFill":
		// 明示的な「塗りつぶしなし／線なし」。style の fillRef/lnRef フォールバックを抑止する。
		switch determineFillCtx(h.inLn, inRPr, inDefRPr, inSpPr) {
		case "sp":
			h.spFillNone = true
		case "ln":
			h.lnFillNone = true
		}
		return true, 0
	case "fillRef":
		// idx="0" は noFill。色子要素を持つが塗りは無いので参照対象にしない。
		h.fillRefIdx = safeAtoi(attrVal(t, "idx"))
		h.inFillRef = h.fillRefIdx != 0
		return true, 0
	case "lnRef":
		// idx="0" は noLine。
		h.lnRefIdx = safeAtoi(attrVal(t, "idx"))
		h.inLnRef = h.lnRefIdx != 0
		return true, 0
	case "srgbClr":
		if h.inFill {
			clr := attrVal(t, "val")
			clr = h.p.applyColorMods(decoder, 0, clr)
			h.p.assignColor(clr, h.fillCtx, &h.shapeFill, h.lineStyle, currentFont, shapeFont)
			return true, -1 // applyColorMods が EndElement まで消費
		}
		if h.inFillRef || h.inLnRef {
			clr := h.p.applyColorMods(decoder, 0, attrVal(t, "val"))
			h.assignStyleRefColor(clr)
			return true, -1
		}
	case "schemeClr":
		if h.inFill {
			clr := h.p.resolveSchemeColor(attrVal(t, "val"), decoder, 0)
			h.p.assignColor(clr, h.fillCtx, &h.shapeFill, h.lineStyle, currentFont, shapeFont)
			return true, -1 // resolveSchemeColor が EndElement まで消費
		}
		if h.inFillRef || h.inLnRef {
			clr := h.p.resolveSchemeColor(attrVal(t, "val"), decoder, 0)
			h.assignStyleRefColor(clr)
			return true, -1
		}
	case "prstDash":
		if h.inLn && h.lineStyle != nil {
			h.lineStyle.Style = attrVal(t, "val")
			return true, 0
		}
	}
	return false, 0
}

// handleEndElement はスタイル関連の EndElement を処理する
func (h *drawingStyleHandler) handleEndElement(name string) {
	switch name {
	case "ln":
		h.inLn = false
	case "solidFill":
		h.inFill = false
		h.fillCtx = ""
	case "fillRef":
		h.inFillRef = false
	case "lnRef":
		h.inLnRef = false
	}
}

// assignStyleRefColor は fillRef / lnRef の phClr 色にテーマ fmtScheme の色変換を適用して格納する
func (h *drawingStyleHandler) assignStyleRefColor(clr string) {
	if clr == "" {
		return
	}
	switch {
	case h.inFillRef:
		h.styleFill = h.p.theme.ApplyFillStyle(h.fillRefIdx, clr)
	case h.inLnRef:
		h.styleLineColor = h.p.theme.ApplyLineStyle(h.lnRefIdx, clr)
	}
}

// resolvedFill は塗り色を返す。spPr の明示塗りを優先し、明示の noFill 時は塗りなし、
// いずれも無ければ style の fillRef を使う。
func (h *drawingStyleHandler) resolvedFill() string {
	if h.shapeFill != "" {
		return h.shapeFill
	}
	if h.spFillNone {
		return ""
	}
	return h.styleFill
}

// resolvedLine は線スタイルを最終化する。ln 内に明示の noFill がある場合は線なし、
// spPr の線色が無い場合は style の lnRef 色で補う。
func (h *drawingStyleHandler) resolvedLine() *LineStyle {
	if h.lnFillNone {
		return nil
	}
	ls := h.lineStyle
	if h.styleLineColor != "" {
		if ls == nil {
			ls = &LineStyle{}
		}
		if ls.Color == "" {
			ls.Color = h.styleLineColor
		}
	}
	return finalizeLineStyle(ls)
}

// handleArrow は矢印関連の StartElement を処理する
func (h *drawingStyleHandler) handleArrow(t xml.StartElement, arrow *string) {
	switch t.Name.Local {
	case "headEnd":
		if h.inLn {
			updateArrow(arrow, "head", attrVal(t, "type"))
		}
	case "tailEnd":
		if h.inLn {
			updateArrow(arrow, "tail", attrVal(t, "type"))
		}
	}
}

// drawingTextState は DrawingML テキスト（txBody）の SAX パース状態
type drawingTextState struct {
	inTxBody bool
	inP      bool
	inR      bool
	inRPr    bool
	inDefRPr bool

	textParts      []string
	currentPara    strings.Builder
	runs           []RichTextRun
	currentRunText strings.Builder
	currentFont    *parsedFont
	hasRuns        bool
	shapeFont      *parsedFont
}

// handleStartElement はテキスト関連の StartElement を処理する
func (ts *drawingTextState) handleStartElement(t xml.StartElement, includeStyle bool) {
	switch t.Name.Local {
	case "txBody":
		ts.inTxBody = true
		ts.textParts = nil
		ts.runs = nil
		ts.hasRuns = false
	case "p":
		if ts.inTxBody {
			ts.inP = true
			ts.currentPara.Reset()
		}
	case "r":
		if ts.inP {
			ts.inR = true
			ts.currentRunText.Reset()
			ts.currentFont = nil
			ts.hasRuns = true
		}
	case "rPr":
		if ts.inR {
			ts.inRPr = true
			ts.currentFont = &parsedFont{}
			parseDrawingFontAttrs(t, ts.currentFont)
		}
	case "defRPr":
		if ts.inP && ts.inTxBody && !ts.inR {
			ts.inDefRPr = true
			if includeStyle && ts.shapeFont == nil {
				ts.shapeFont = &parsedFont{}
			}
			if ts.shapeFont != nil {
				parseDrawingFontAttrs(t, ts.shapeFont)
			}
		}
	case "latin", "ea":
		font := ts.currentFont
		if ts.inDefRPr {
			font = ts.shapeFont
		}
		if font != nil {
			if v := attrVal(t, "typeface"); v != "" {
				font.Name = v
			}
		}
	}
}

// handleEndElement はテキスト関連の EndElement を処理する
func (ts *drawingTextState) handleEndElement(name string, includeStyle bool, theme *themeColors) {
	switch name {
	case "txBody":
		ts.inTxBody = false
	case "p":
		if ts.inP {
			ts.textParts = append(ts.textParts, ts.currentPara.String())
			ts.inP = false
		}
	case "r":
		if ts.inR {
			text := ts.currentRunText.String()
			ts.currentPara.WriteString(text)
			run := RichTextRun{Text: text}
			if ts.currentFont != nil && includeStyle {
				run.Font = buildDrawingFontObj(ts.currentFont, theme)
			}
			ts.runs = append(ts.runs, run)
			ts.inR = false
		}
	case "rPr":
		ts.inRPr = false
	case "defRPr":
		ts.inDefRPr = false
	}
}

// handleCharData はテキスト関連の CharData を処理する
func (ts *drawingTextState) handleCharData(data []byte) {
	if ts.inP && !ts.inR {
		text := string(data)
		if strings.TrimSpace(text) != "" {
			ts.currentPara.Write(data)
		}
	}
	if ts.inR {
		ts.currentRunText.Write(data)
	}
}

// buildText はテキストパーツを結合して返す
func (ts *drawingTextState) buildText() string {
	return strings.Join(ts.textParts, "\n")
}
