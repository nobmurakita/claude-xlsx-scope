package excel

import (
	"fmt"
	"log"
)

// ParseWarnings はパース中の警告を収集する。
// 各警告は即座にログ出力せず、呼び出し元が Flush() で一括出力できる。
type ParseWarnings struct {
	items []string
}

// Add は警告メッセージを追加する
func (w *ParseWarnings) Add(format string, args ...any) {
	w.items = append(w.items, fmt.Sprintf(format, args...))
}

// Flush は収集した警告をログに出力する
func (w *ParseWarnings) Flush() {
	for _, msg := range w.items {
		log.Printf("[WARN] %s", msg)
	}
}

// HasWarnings は警告が存在するかを返す
func (w *ParseWarnings) HasWarnings() bool {
	return len(w.items) > 0
}
