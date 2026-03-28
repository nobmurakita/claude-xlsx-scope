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
	case "srgbClr":
		if h.inFill {
			clr := attrVal(t, "val")
			clr = h.p.applyColorMods(decoder, 0, clr)
			h.p.assignColor(clr, h.fillCtx, &h.shapeFill, h.lineStyle, currentFont, shapeFont)
			return true, -1 // applyColorMods が EndElement まで消費
		}
	case "schemeClr":
		if h.inFill {
			clr := h.p.resolveSchemeColor(attrVal(t, "val"), decoder, 0)
			h.p.assignColor(clr, h.fillCtx, &h.shapeFill, h.lineStyle, currentFont, shapeFont)
			return true, -1 // resolveSchemeColor が EndElement まで消費
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
	}
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
