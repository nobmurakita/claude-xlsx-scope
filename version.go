package main

import (
	"fmt"

	"github.com/spf13/cobra"
)

var version = "dev"

func init() {
	rootCmd.AddCommand(versionCmd)
}

var versionCmd = &cobra.Command{
	Use:   "version",
	Short: "バージョン情報を表示する",
	Run: func(cmd *cobra.Command, args []string) {
		fmt.Printf("cc-read-excel version %s\n", version)
	},
}
