package main

import (
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-xlsx/internal/cmd"
	"github.com/spf13/cobra"
)

const (
	exitOK    = 0
	exitError = 1
)

var rootCmd = &cobra.Command{
	Use:           "read-xlsx",
	Short:         "Excel ファイル（.xlsx / .xlsm）の内容をAIエージェント向けに読み取るツール",
	SilenceUsage:  true,
	SilenceErrors: true,
}

func init() {
	rootCmd.AddCommand(
		cmd.NewCellsCmd(),
		cmd.NewSearchCmd(),
		cmd.NewScanCmd(),
		cmd.NewInfoCmd(),
		cmd.NewShapesCmd(),
		cmd.NewImageCmd(),
	)
}

func execute() int {
	err := rootCmd.Execute()
	if err == nil {
		return exitOK
	}
	fmt.Fprintf(os.Stderr, "read-xlsx: %s\n", err)
	return exitError
}
