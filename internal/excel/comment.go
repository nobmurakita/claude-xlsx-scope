package excel

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"log"
	"strconv"
	"strings"
)

// CommentData はセルに付与されたコメント情報
type CommentData struct {
	Author string         `json:"author,omitempty"`
	Text   string         `json:"text"`
	Thread []ThreadEntry  `json:"thread,omitempty"`
}

// ThreadEntry はスレッドコメントの1エントリ（返信）
type ThreadEntry struct {
	Author string `json:"author,omitempty"`
	Text   string `json:"text"`
	Date   string `json:"date,omitempty"`
	Done   bool   `json:"done,omitempty"`
}

// CommentMap はセル参照 → コメントデータのマップ
type CommentMap map[string]*CommentData

// LoadComments はシートのコメントを読み込む
func (f *File) LoadComments(sheet string) CommentMap {
	xmlPath, ok := f.sheetPaths[sheet]
	if !ok {
		return nil
	}

	rels := loadSheetRelsAll(f.zr, xmlPath)
	if len(rels) == 0 {
		return nil
	}

	// レガシーコメントを読む
	comments := make(CommentMap)
	for _, rel := range rels {
		if strings.Contains(rel.Type, relKeywordComments) {
			commentsPath := resolveRelTarget(xmlPath, rel.Target)
			parseComments(f.zr, commentsPath, comments)
		}
	}

	// スレッドコメントを読む
	for _, rel := range rels {
		if strings.Contains(strings.ToLower(rel.Type), relKeywordThreadedComments) {
			threadPath := resolveRelTarget(xmlPath, rel.Target)
			parseThreadedComments(f.zr, threadPath, comments)
		}
	}

	if len(comments) == 0 {
		return nil
	}
	return comments
}

// loadSheetRelsAll はシートの .rels から全リレーションを返す
func loadSheetRelsAll(zr *zip.ReadCloser, sheetXMLPath string) []xmlRelationship {
	data, err := readZipFile(zr, relsPathFor(sheetXMLPath))
	if err != nil {
		return nil
	}

	var rels xmlRelationships
	if err := xml.Unmarshal(data, &rels); err != nil {
		return nil
	}
	return rels.Rels
}

// resolveRelTarget はリレーションターゲットをZIP内の絶対パスに解決する
func resolveRelTarget(sheetXMLPath, target string) string {
	if strings.HasPrefix(target, "/") {
		return target[1:]
	}
	dir := sheetXMLPath[:strings.LastIndex(sheetXMLPath, "/")+1]
	resolved := dir + target
	parts := strings.Split(resolved, "/")
	var result []string
	for _, p := range parts {
		if p == ".." {
			if len(result) > 0 {
				result = result[:len(result)-1]
			}
		} else if p != "" && p != "." {
			result = append(result, p)
		}
	}
	return strings.Join(result, "/")
}

// parseComments はレガシーコメント（comments.xml）をパースする
func parseComments(zr *zip.ReadCloser, path string, comments CommentMap) {
	if entry := findZipEntry(zr, path); entry != nil {
		parseCommentsEntry(entry, comments)
	}
}

// commentParseState は parseCommentsEntry の SAX パーサー状態
type commentParseState struct {
	inAuthors bool
	inAuthor  bool
	inComment bool
	inText    bool
	inT       bool
}

func parseCommentsEntry(entry *zip.File, comments CommentMap) {
	_ = withZipXML(entry, func(decoder *xml.Decoder) error {
		parseCommentsSAX(decoder, comments)
		return nil
	})
}

func parseCommentsSAX(decoder *xml.Decoder, comments CommentMap) {
	var st commentParseState
	var (
		authors    []string
		commentRef string
		authorID   int
		textBuf    strings.Builder
		authorBuf  strings.Builder
	)

	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Printf("[WARN] parseCommentsEntry: XMLトークン読み取りに失敗: %v", err)
			return
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "authors":
				st.inAuthors = true
			case "author":
				if st.inAuthors {
					st.inAuthor = true
					authorBuf.Reset()
				}
			case "comment":
				st.inComment = true
				commentRef = ""
				authorID = 0
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "ref":
						commentRef = attr.Value
					case "authorId":
						authorID = safeAtoi(attr.Value)
					}
				}
			case "text":
				if st.inComment {
					st.inText = true
					textBuf.Reset()
				}
			case "t":
				if st.inText {
					st.inT = true
				}
			}

		case xml.EndElement:
			switch t.Name.Local {
			case "authors":
				st.inAuthors = false
			case "author":
				if st.inAuthor {
					authors = append(authors, authorBuf.String())
					st.inAuthor = false
				}
			case "comment":
				if st.inComment && commentRef != "" {
					cd := &CommentData{
						Text: textBuf.String(),
					}
					if authorID >= 0 && authorID < len(authors) {
						cd.Author = authors[authorID]
					}
					comments[commentRef] = cd
				}
				st.inComment = false
			case "text":
				st.inText = false
			case "t":
				st.inT = false
			}

		case xml.CharData:
			if st.inAuthor {
				authorBuf.Write(t)
			}
			if st.inT && st.inText {
				textBuf.Write(t)
			}
		}
	}
}

// parseThreadedComments はスレッドコメント（threadedComment.xml）をパースする
func parseThreadedComments(zr *zip.ReadCloser, path string, comments CommentMap) {
	if entry := findZipEntry(zr, path); entry != nil {
		parseThreadedCommentsEntry(entry, comments)
	}
}

// threadedCommentState は parseThreadedCommentsEntry の SAX パーサー状態
type threadedCommentState struct {
	inComment bool
	inText    bool
}

// threadedCommentRaw はパース時の中間データ
type threadedCommentRaw struct {
	ref      string
	parentID string
	id       string
	personID string
	date     string
	text     string
	done     bool
}

func parseThreadedCommentsEntry(entry *zip.File, comments CommentMap) {
	var items []threadedCommentRaw
	_ = withZipXML(entry, func(decoder *xml.Decoder) error {
		items = parseThreadedCommentsSAX(decoder)
		return nil
	})
	resolveThreadedComments(items, comments)
}

func parseThreadedCommentsSAX(decoder *xml.Decoder) []threadedCommentRaw {
	var items []threadedCommentRaw
	var current threadedCommentRaw
	var st threadedCommentState
	var textBuf strings.Builder

	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return items
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "threadedComment" {
				st.inComment = true
				current = threadedCommentRaw{}
				textBuf.Reset()
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "ref":
						current.ref = attr.Value
					case "parentId":
						current.parentID = attr.Value
					case "id":
						current.id = attr.Value
					case "personId":
						current.personID = attr.Value
					case "dT":
						current.date = attr.Value
					case "done":
						current.done = attr.Value == "1"
					}
				}
			} else if t.Name.Local == "text" && st.inComment {
				st.inText = true
				textBuf.Reset()
			}

		case xml.EndElement:
			if t.Name.Local == "threadedComment" {
				current.text = textBuf.String()
				items = append(items, current)
				st.inComment = false
			} else if t.Name.Local == "text" {
				st.inText = false
			}

		case xml.CharData:
			if st.inText {
				textBuf.Write(t)
			}
		}
	}
	return items
}

// resolveThreadedComments はパース済みスレッドコメントをレガシーコメントに統合する
func resolveThreadedComments(items []threadedCommentRaw, comments CommentMap) {
	if len(items) == 0 {
		return
	}

	// 親コメントのIDマップ
	idMap := make(map[string]*threadedCommentRaw, len(items))
	for i := range items {
		if items[i].id != "" {
			idMap[items[i].id] = &items[i]
		}
	}

	// 親コメント（parentID なし）→ レガシーコメントのテキストを上書き
	// 返信（parentID あり）→ thread に追加
	for i := range items {
		item := &items[i]

		if item.parentID == "" {
			cd, ok := comments[item.ref]
			if !ok {
				cd = &CommentData{}
				comments[item.ref] = cd
			}
			cd.Text = item.text
		} else {
			parent, ok := idMap[item.parentID]
			if !ok {
				continue
			}
			cd, ok := comments[parent.ref]
			if !ok {
				continue
			}
			te := ThreadEntry{
				Text: item.text,
				Date: item.date,
			}
			if item.done {
				te.Done = true
			}
			cd.Thread = append(cd.Thread, te)
		}
	}
}

// safeAtoi は文字列を int に変換する（エラー時は 0）
func safeAtoi(s string) int {
	n, _ := strconv.Atoi(s)
	return n
}
