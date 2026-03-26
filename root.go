package main

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

const (
	exitOK    = 0
	exitError = 1
)

var rootCmd = &cobra.Command{
	Use:           "exceldump",
	Short:         "Excel ファイル（.xlsx / .xlsm）の内容をCLIからダンプするツール",
	SilenceUsage:  true,
	SilenceErrors: true,
}

func execute() int {
	err := rootCmd.Execute()
	if err == nil {
		return exitOK
	}
	fmt.Fprintf(os.Stderr, "exceldump: %s\n", err)
	return exitError
}
