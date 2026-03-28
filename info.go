package main

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/nobmurakita/exceldump/internal/excel"
	"github.com/spf13/cobra"
)

func init() {
	rootCmd.AddCommand(infoCmd)
}

var infoCmd = &cobra.Command{
	Use:   "info <file>",
	Short: "ファイルの概要（シート一覧、定義名）を表示する",
	Args:  cobra.ExactArgs(1),
	RunE:  runInfo,
}

type infoOutput struct {
	File         string           `json:"file"`
	DefinedNames []definedNameOut `json:"defined_names"`
	Sheets       []sheetOut       `json:"sheets"`
}

type definedNameOut struct {
	Name  string `json:"name"`
	Scope string `json:"scope"`
	Refer string `json:"refer"`
}

type sheetOut struct {
	Index     int    `json:"index"`
	Name      string `json:"name"`
	Type      string `json:"type"`
	Hidden    bool   `json:"hidden,omitempty"`
	Dimension string `json:"dimension,omitempty"`
}

func runInfo(cmd *cobra.Command, args []string) error {
	result, err := excel.QuickInfo(args[0])
	if err != nil {
		return err
	}

	sheetOuts := make([]sheetOut, len(result.Sheets))
	for i, s := range result.Sheets {
		sheetOuts[i] = sheetOut{
			Index:     s.Index,
			Name:      s.Name,
			Type:      s.Type,
			Hidden:    s.Hidden,
			Dimension: s.Dimension,
		}
	}

	dnOuts := make([]definedNameOut, len(result.DefinedNames))
	for i, dn := range result.DefinedNames {
		scope := dn.Scope
		if scope == "" {
			scope = "Workbook"
		}
		dnOuts[i] = definedNameOut{
			Name:  dn.Name,
			Scope: scope,
			Refer: dn.RefersTo,
		}
	}

	out := infoOutput{
		File:         result.FileName,
		DefinedNames: dnOuts,
		Sheets:       sheetOuts,
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)
	if err := enc.Encode(out); err != nil {
		return fmt.Errorf("JSON出力に失敗しました: %w", err)
	}
	return nil
}
