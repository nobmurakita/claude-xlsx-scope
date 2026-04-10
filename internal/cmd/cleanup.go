package cmd

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"
)

// tmpFilePrefix は cleanup の安全確認に使う一時ファイルのプレフィックス
const tmpFilePrefix = "xlsx-scope-tmp-"

// NewCleanupCmd は cleanup サブコマンドを生成する
func NewCleanupCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "cleanup <file> [file...]",
		Short: "xlsx-scope が生成した一時ファイルを削除する",
		Args:  cobra.MinimumNArgs(1),
		RunE:  runCleanup,
	}
}

type cleanupOutput struct {
	Deleted int `json:"deleted"`
}

func runCleanup(cmd *cobra.Command, args []string) error {
	// os.CreateTemp("", ...) で生成された一時ファイルのみを削除対象とする。
	// 他コマンドは os.TempDir() 直下に xlsx-scope-tmp-* を作成するため、
	// 親ディレクトリがそれと一致するかを文字列比較で確認する。
	tmpDir := filepath.Clean(os.TempDir())

	deleted := 0
	for _, path := range args {
		abs, err := filepath.Abs(path)
		if err != nil {
			return fmt.Errorf("パスの解決エラー: %w", err)
		}
		if !strings.HasPrefix(filepath.Base(abs), tmpFilePrefix) {
			return fmt.Errorf("xlsx-scope が生成した一時ファイルではありません: %s", path)
		}
		if filepath.Dir(abs) != tmpDir {
			return fmt.Errorf("一時ディレクトリ配下ではありません: %s", path)
		}
		if err := os.Remove(abs); err != nil {
			if os.IsNotExist(err) {
				continue
			}
			return fmt.Errorf("ファイルの削除エラー: %w", err)
		}
		deleted++
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)
	return enc.Encode(cleanupOutput{Deleted: deleted})
}
