package excel

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"strings"
)

// sharedStrings は共有文字列テーブル
type sharedStrings struct {
	items []string
}

// Get はインデックスから共有文字列を返す
func (ss *sharedStrings) Get(index int) string {
	if ss == nil || index < 0 || index >= len(ss.items) {
		return ""
	}
	return ss.items[index]
}

// parseSharedStringsFromZip は ZIP 内の sharedStrings.xml を SAX パースする
func parseSharedStringsFromZip(zr *zip.ReadCloser) (*sharedStrings, error) {
	for _, f := range zr.File {
		if f.Name == "xl/sharedStrings.xml" {
			return parseSharedStringsEntry(f)
		}
	}
	// sharedStrings.xml がないファイルもある（数値のみ等）
	return &sharedStrings{}, nil
}

func parseSharedStringsEntry(f *zip.File) (*sharedStrings, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	ss := &sharedStrings{}
	decoder := xml.NewDecoder(rc)

	var inSI, inRPh, inT bool
	var buf strings.Builder

	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "si":
				inSI = true
				buf.Reset()
			case "rPh":
				// ルビ（ふりがな）要素: 内部の <t> は無視する
				inRPh = true
			case "t":
				if inSI && !inRPh {
					inT = true
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "si":
				ss.items = append(ss.items, buf.String())
				inSI = false
			case "rPh":
				inRPh = false
			case "t":
				inT = false
			}
		case xml.CharData:
			if inT {
				buf.Write(t)
			}
		}
	}

	return ss, nil
}
