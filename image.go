package main

import (
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-xlsx/internal/excel"
	"github.com/spf13/cobra"
)

func init() {
	rootCmd.AddCommand(imageCmd)
}

var imageCmd = &cobra.Command{
	Use:   "image <file> <image_id> <output>",
	Short: "画像をファイルに保存する",
	Args:  cobra.ExactArgs(3),
	RunE:  runImage,
}

func runImage(cmd *cobra.Command, args []string) error {
	f, err := excel.OpenFile(args[0])
	if err != nil {
		return err
	}
	defer f.Close()

	out, err := os.Create(args[2])
	if err != nil {
		return fmt.Errorf("出力ファイルの作成エラー: %w", err)
	}
	defer out.Close()

	return f.ExtractImage(args[1], out)
}
