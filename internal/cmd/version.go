package cmd

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

// Version はビルド時に -ldflags="-X ..." で埋め込まれるバージョン文字列
var Version = "latest"

// NewVersionCmd は version サブコマンドを生成する
func NewVersionCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "version",
		Short: "バージョン情報を表示する",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			_, err := fmt.Fprintln(os.Stdout, Version)
			return err
		},
	}
}
