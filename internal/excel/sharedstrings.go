package excel

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"strconv"
	"strings"
)

// sharedStringItem は共有文字列テーブルのエントリ
type sharedStringItem struct {
	Text string
	Runs []richTextRunRaw // リッチテキストの場合のみ非nil（2ラン以上）
}

// richTextRunRaw は sharedStrings.xml から取得したリッチテキストラン
type richTextRunRaw struct {
	Text string
	Font *parsedFont // nil = フォント指定なし
}

// sharedStrings は共有文字列テーブル
type sharedStrings struct {
	items []sharedStringItem
}

// Get はインデックスからテキストを返す
func (ss *sharedStrings) Get(index int) string {
	if ss == nil || index < 0 || index >= len(ss.items) {
		return ""
	}
	return ss.items[index].Text
}

// GetRichTextRuns はインデックスからリッチテキストランを返す。
// リッチテキストでない場合（ラン数が1以下）は nil を返す。
func (ss *sharedStrings) GetRichTextRuns(index int) []richTextRunRaw {
	if ss == nil || index < 0 || index >= len(ss.items) {
		return nil
	}
	return ss.items[index].Runs
}

// parseSharedStringsFromZip は ZIP 内の sharedStrings.xml を SAX パースする
func parseSharedStringsFromZip(zr *zip.ReadCloser) (*sharedStrings, error) {
	if entry := findZipEntry(zr, "xl/sharedStrings.xml"); entry != nil {
		return parseSharedStringsEntry(entry)
	}
	// sharedStrings.xml がないファイルもある（数値のみ等）
	return &sharedStrings{}, nil
}

func parseSharedStringsEntry(f *zip.File) (*sharedStrings, error) {
	ss := &sharedStrings{}
	err := withZipXML(f, func(decoder *xml.Decoder) error {
		return parseSharedStringsSAX(decoder, ss)
	})
	if err != nil {
		return nil, err
	}
	return ss, nil
}

func parseSharedStringsSAX(decoder *xml.Decoder, ss *sharedStrings) error {

	type saxState struct {
		inSI  bool
		inR   bool // <r> リッチテキストラン内
		inRPr bool // <rPr> フォントプロパティ内
		inRPh bool // <rPh> ルビ内
		inT   bool // <t> テキスト内
	}
	var st saxState
	var textBuf strings.Builder
	var runs []richTextRunRaw
	var currentFont *parsedFont
	var runText strings.Builder

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
			switch t.Name.Local {
			case "si":
				st.inSI = true
				textBuf.Reset()
				runs = nil

			case "r":
				if st.inSI && !st.inRPh {
					st.inR = true
					currentFont = nil
					runText.Reset()
				}

			case "rPr":
				if st.inR {
					st.inRPr = true
					currentFont = &parsedFont{}
				}

			case "rPh":
				st.inRPh = true

			case "t":
				if st.inSI && !st.inRPh {
					st.inT = true
				}

			// rPr 内のフォント属性
			case "rFont":
				if st.inRPr && currentFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							currentFont.Name = attr.Value
						}
					}
				}
			case "sz":
				if st.inRPr && currentFont != nil {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							currentFont.Size, _ = strconv.ParseFloat(attr.Value, 64)
						}
					}
				}
			case "b":
				if st.inRPr && currentFont != nil {
					currentFont.Bold = true
				}
			case "i":
				if st.inRPr && currentFont != nil {
					currentFont.Italic = true
				}
			case "strike":
				if st.inRPr && currentFont != nil {
					currentFont.Strikethrough = true
				}
			case "u":
				if st.inRPr && currentFont != nil {
					val := "single"
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							val = attr.Value
						}
					}
					currentFont.Underline = val
				}
			case "color":
				if st.inRPr && currentFont != nil {
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "rgb":
							currentFont.Color = normalizeHexColor(attr.Value)
						case "theme":
							v, _ := strconv.Atoi(attr.Value)
							currentFont.ColorTheme = &v
						case "tint":
							currentFont.ColorTint, _ = strconv.ParseFloat(attr.Value, 64)
						}
					}
				}
			}

		case xml.EndElement:
			switch t.Name.Local {
			case "si":
				item := sharedStringItem{Text: textBuf.String()}
				if len(runs) > 0 {
					item.Runs = runs
				}
				ss.items = append(ss.items, item)
				st.inSI = false

			case "r":
				if st.inR {
					text := runText.String()
					runs = append(runs, richTextRunRaw{
						Text: text,
						Font: currentFont,
					})
					st.inR = false
				}

			case "rPr":
				st.inRPr = false

			case "rPh":
				st.inRPh = false

			case "t":
				st.inT = false
			}

		case xml.CharData:
			if st.inT {
				textBuf.Write(t)
				if st.inR {
					runText.Write(t)
				}
			}
		}
	}

	return nil
}
