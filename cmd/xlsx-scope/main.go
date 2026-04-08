package main

import (
	"fmt"
	"os"

	"github.com/nobmurakita/claude-xlsx-scope/internal/cmd"
	"github.com/spf13/cobra"
)

func main() {
	rootCmd := &cobra.Command{
		Use:           "xlsx-scope",
		Short:         "Excel ファイル（.xlsx / .xlsm）の内容をAIエージェント向けに読み取るツール",
		SilenceUsage:  true,
		SilenceErrors: true,
	}
	rootCmd.PersistentFlags().Bool("stdout", false, "出力を標準出力に直接書き出す（デバッグ用）")
	rootCmd.AddCommand(
		cmd.NewCellsCmd(),
		cmd.NewSearchCmd(),
		cmd.NewScanCmd(),
		cmd.NewInfoCmd(),
		cmd.NewShapesCmd(),
		cmd.NewImageCmd(),
		cmd.NewValuesCmd(),
	)

	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintf(os.Stderr, "xlsx-scope: %s\n", err)
		os.Exit(1)
	}
}
