package excel

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"log"
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
		if strings.Contains(rel.Type, "/comments") {
			commentsPath := resolveRelTarget(xmlPath, rel.Target)
			parseComments(f.zr, commentsPath, comments)
		}
	}

	// スレッドコメントを読む
	for _, rel := range rels {
		if strings.Contains(rel.Type, "threadedcomments") || strings.Contains(rel.Type, "threadedComments") {
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

func parseCommentsEntry(entry *zip.File, comments CommentMap) {
	rc, err := entry.Open()
	if err != nil {
		log.Printf("[WARN] parseCommentsEntry: ZIPエントリ %s のオープンに失敗: %v", entry.Name, err)
		return
	}
	defer rc.Close()

	decoder := xml.NewDecoder(rc)

	var (
		authors     []string
		inAuthors   bool
		inAuthor    bool
		inComment   bool
		inText      bool
		inT         bool
		commentRef  string
		authorID    int
		textBuf     strings.Builder
		authorBuf   strings.Builder
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
				inAuthors = true
			case "author":
				if inAuthors {
					inAuthor = true
					authorBuf.Reset()
				}
			case "comment":
				inComment = true
				commentRef = ""
				authorID = 0
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "ref":
						commentRef = attr.Value
					case "authorId":
						authorID = atoi(attr.Value)
					}
				}
			case "text":
				if inComment {
					inText = true
					textBuf.Reset()
				}
			case "t":
				if inText {
					inT = true
				}
			}

		case xml.EndElement:
			switch t.Name.Local {
			case "authors":
				inAuthors = false
			case "author":
				if inAuthor {
					authors = append(authors, authorBuf.String())
					inAuthor = false
				}
			case "comment":
				if inComment && commentRef != "" {
					cd := &CommentData{
						Text: textBuf.String(),
					}
					if authorID >= 0 && authorID < len(authors) {
						cd.Author = authors[authorID]
					}
					comments[commentRef] = cd
				}
				inComment = false
			case "text":
				inText = false
			case "t":
				inT = false
			}

		case xml.CharData:
			if inAuthor {
				authorBuf.Write(t)
			}
			if inT && inText {
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
	rc, err := entry.Open()
	if err != nil {
		log.Printf("[WARN] parseThreadedCommentsEntry: ZIPエントリ %s のオープンに失敗: %v", entry.Name, err)
		return
	}
	defer rc.Close()

	decoder := xml.NewDecoder(rc)

	var items []threadedCommentRaw
	var current threadedCommentRaw
	var inComment, inText bool
	var textBuf strings.Builder

	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Printf("[WARN] parseThreadedCommentsEntry: XMLトークン読み取りに失敗: %v", err)
			return
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "threadedComment" {
				inComment = true
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
			} else if t.Name.Local == "text" && inComment {
				inText = true
				textBuf.Reset()
			}

		case xml.EndElement:
			if t.Name.Local == "threadedComment" {
				current.text = textBuf.String()
				items = append(items, current)
				inComment = false
			} else if t.Name.Local == "text" {
				inText = false
			}

		case xml.CharData:
			if inText {
				textBuf.Write(t)
			}
		}
	}

	// personId → 名前のマップを構築（レガシーコメントの著者名を流用）
	// スレッドコメントの personId は GUID なので、レガシーコメントの著者名と直接対応しない
	// ただし、同じセルのレガシーコメントの著者名を使う

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
			// 親コメント: レガシーコメントが既にあればそのauthorを保持
			cd, ok := comments[item.ref]
			if !ok {
				cd = &CommentData{}
				comments[item.ref] = cd
			}
			// スレッドコメントのテキストで上書き（より新しい）
			cd.Text = item.text
			if item.date != "" {
				// 日付は親コメントには付けない（thread エントリにのみ）
			}
		} else {
			// 返信: 親コメントのセルを特定
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

// atoi は文字列を int に変換する（エラー時は 0）
func atoi(s string) int {
	n := 0
	for _, c := range s {
		if c >= '0' && c <= '9' {
			n = n*10 + int(c-'0')
		}
	}
	return n
}
