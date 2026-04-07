package cmd

import (
	"bufio"
	"encoding/json"
	"fmt"
	"io"
	"os"

	"github.com/spf13/cobra"
)

// defaultOutputLimit は cells/search/shapes のデフォルト出力上限
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

// outputResult は一時ファイルモード時に stdout に出力する結果JSON
type outputResult struct {
	File  string `json:"file"`
	Lines *int   `json:"lines,omitempty"`
}

// outputWriter は一時ファイルへの書き込みと行数カウントを管理する
type outputWriter struct {
	w         *bufio.Writer
	file      *os.File
	lineCount int
	useStdout bool
}

// newOutputWriter はコマンドの出力先を生成する。
// --stdout フラグが指定されていれば標準出力、なければ一時ファイルに書き出す。
func newOutputWriter(cmd *cobra.Command) (*outputWriter, error) {
	useStdout, _ := cmd.Root().PersistentFlags().GetBool("stdout")
	if useStdout {
		return &outputWriter{
			w:         bufio.NewWriter(os.Stdout),
			useStdout: true,
		}, nil
	}
	f, err := os.CreateTemp("", "xlsx-scope-*.jsonl")
	if err != nil {
		return nil, fmt.Errorf("一時ファイルの作成エラー: %w", err)
	}
	return &outputWriter{
		w:    bufio.NewWriter(f),
		file: f,
	}, nil
}

func (ow *outputWriter) Write(p []byte) (n int, err error) {
	n, err = ow.w.Write(p)
	for _, b := range p[:n] {
		if b == '\n' {
			ow.lineCount++
		}
	}
	return
}

// cleanup は一時ファイルを閉じる（結果JSONは出力しない）
func (ow *outputWriter) cleanup() {
	if ow.file != nil {
		ow.file.Close()
	}
}

// finalize はバッファをフラッシュし、一時ファイルモードなら結果JSONを stdout に出力する。
// 成功時のみ呼び出すこと。
func (ow *outputWriter) finalize() error {
	if err := ow.w.Flush(); err != nil {
		return err
	}
	if ow.useStdout {
		return nil
	}
	name := ow.file.Name()
	if err := ow.file.Close(); err != nil {
		return err
	}
	ow.file = nil // cleanup での二重 Close を防止
	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)
	lines := ow.lineCount
	return enc.Encode(outputResult{File: name, Lines: &lines})
}
