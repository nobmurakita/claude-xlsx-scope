package main

import (
	"encoding/json"
	"io"
)

// defaultOutputLimit は dump/search/shapes のデフォルト出力上限
const defaultOutputLimit = 1000

type truncatedOutput struct {
	Truncated bool   `json:"_truncated"`
	NextCell  string `json:"next_cell"`
}

// newJSONLWriter は JSONL 出力用のエンコーダを生成する
func newJSONLWriter(w io.Writer) *json.Encoder {
	enc := json.NewEncoder(w)
	enc.SetEscapeHTML(false)
	return enc
}
