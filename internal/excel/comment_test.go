package excel

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestParseCommentsSAX(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>田中太郎</author>
    <author>佐藤花子</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text><r><t>最初のコメント</t></r></text>
    </comment>
    <comment ref="B2" authorId="1">
      <text><r><t>二番目のコメント</t></r></text>
    </comment>
  </commentList>
</comments>`

	comments := make(CommentMap)
	decoder := xml.NewDecoder(strings.NewReader(xmlData))
	parseCommentsSAX(decoder, comments)

	if len(comments) != 2 {
		t.Fatalf("got %d comments, want 2", len(comments))
	}

	a1 := comments["A1"]
	if a1 == nil || a1.Author != "田中太郎" || a1.Text != "最初のコメント" {
		t.Errorf("A1 = %+v, want {Author:田中太郎, Text:最初のコメント}", a1)
	}

	b2 := comments["B2"]
	if b2 == nil || b2.Author != "佐藤花子" || b2.Text != "二番目のコメント" {
		t.Errorf("B2 = %+v, want {Author:佐藤花子, Text:二番目のコメント}", b2)
	}
}

func TestParseThreadedCommentsSAX(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<ThreadedComments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:x18tc="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" dT="2025-01-01T10:00:00.00" personId="{person-1}" id="{tc-1}">
    <text>親コメント</text>
  </threadedComment>
  <threadedComment ref="A1" dT="2025-01-02T12:00:00.00" personId="{person-2}" id="{tc-2}" parentId="{tc-1}">
    <text>返信コメント</text>
  </threadedComment>
</ThreadedComments>`

	decoder := xml.NewDecoder(strings.NewReader(xmlData))
	items := parseThreadedCommentsSAX(decoder)

	if len(items) != 2 {
		t.Fatalf("got %d items, want 2", len(items))
	}

	if items[0].ref != "A1" || items[0].personID != "{person-1}" || items[0].text != "親コメント" {
		t.Errorf("items[0] = %+v", items[0])
	}
	if items[0].parentID != "" {
		t.Errorf("items[0].parentID = %q, want empty", items[0].parentID)
	}

	if items[1].parentID != "{tc-1}" || items[1].personID != "{person-2}" || items[1].text != "返信コメント" {
		t.Errorf("items[1] = %+v", items[1])
	}
}

func TestResolveThreadedComments(t *testing.T) {
	comments := CommentMap{
		"A1": {Author: "legacy author", Text: "legacy text"},
	}

	persons := map[string]string{
		"{person-1}": "平井理夫",
		"{person-2}": "森田航平",
	}

	items := []threadedCommentRaw{
		{ref: "A1", id: "{tc-1}", personID: "{person-1}", text: "新しいコメント"},
		{ref: "A1", id: "{tc-2}", parentID: "{tc-1}", personID: "{person-2}", text: "返信です", date: "2025-01-02T12:00:00.00"},
	}

	resolveThreadedComments(items, comments, persons)

	cd := comments["A1"]
	if cd == nil {
		t.Fatal("A1 comment is nil")
	}

	// 親コメントのテキストと著者がスレッドコメントで上書きされる
	if cd.Text != "新しいコメント" {
		t.Errorf("Text = %q, want %q", cd.Text, "新しいコメント")
	}
	if cd.Author != "平井理夫" {
		t.Errorf("Author = %q, want %q", cd.Author, "平井理夫")
	}

	// 返信
	if len(cd.Thread) != 1 {
		t.Fatalf("Thread has %d entries, want 1", len(cd.Thread))
	}
	if cd.Thread[0].Author != "森田航平" {
		t.Errorf("Thread[0].Author = %q, want %q", cd.Thread[0].Author, "森田航平")
	}
	if cd.Thread[0].Text != "返信です" {
		t.Errorf("Thread[0].Text = %q, want %q", cd.Thread[0].Text, "返信です")
	}
	if cd.Thread[0].Date != "2025-01-02T12:00:00.00" {
		t.Errorf("Thread[0].Date = %q", cd.Thread[0].Date)
	}
}

func TestResolveThreadedComments_NoPersons(t *testing.T) {
	comments := CommentMap{
		"B1": {Author: "original", Text: "original text"},
	}

	items := []threadedCommentRaw{
		{ref: "B1", id: "{tc-1}", personID: "{unknown}", text: "updated text"},
	}

	resolveThreadedComments(items, comments, nil)

	cd := comments["B1"]
	// persons が nil の場合、著者名は上書きされない
	if cd.Author != "original" {
		t.Errorf("Author = %q, want %q (should not be overwritten)", cd.Author, "original")
	}
	if cd.Text != "updated text" {
		t.Errorf("Text = %q, want %q", cd.Text, "updated text")
	}
}

func TestResolveThreadedComments_NewCell(t *testing.T) {
	comments := make(CommentMap)

	persons := map[string]string{
		"{p1}": "Author Name",
	}

	items := []threadedCommentRaw{
		{ref: "C1", id: "{tc-1}", personID: "{p1}", text: "new comment"},
	}

	resolveThreadedComments(items, comments, persons)

	cd := comments["C1"]
	if cd == nil {
		t.Fatal("C1 should be created")
	}
	if cd.Author != "Author Name" || cd.Text != "new comment" {
		t.Errorf("C1 = %+v", cd)
	}
}
