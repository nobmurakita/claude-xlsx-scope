package main

import (
	"errors"
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

const (
	exitOK      = 0
	exitNoMatch = 1
	exitError   = 2
)

var errNoMatch = errors.New("no match")

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
	if errors.Is(err, errNoMatch) {
		return exitNoMatch
	}
	fmt.Fprintf(os.Stderr, "exceldump: %s\n", err)
	return exitError
}
