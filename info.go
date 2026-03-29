package main

import (
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-xlsx/internal/excel"
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
	File         string              `json:"file"`
	DefinedNames []definedNameOutput `json:"defined_names"`
	Sheets       []sheetOutput       `json:"sheets"`
}

type definedNameOutput struct {
	Name  string `json:"name"`
	Scope string `json:"scope"`
	Refer string `json:"refer"`
}

type sheetOutput struct {
	Index  int    `json:"index"`
	Name   string `json:"name"`
	Type   string `json:"type"`
	Hidden bool   `json:"hidden,omitempty"`
}

func runInfo(cmd *cobra.Command, args []string) error {
	result, err := excel.BookInfo(args[0])
	if err != nil {
		return err
	}

	sheetOutputs := make([]sheetOutput, len(result.Sheets))
	for i, s := range result.Sheets {
		sheetOutputs[i] = sheetOutput{
			Index:  s.Index,
			Name:   s.Name,
			Type:   s.Type,
			Hidden: s.Hidden,
		}
	}

	dnOutputs := make([]definedNameOutput, len(result.DefinedNames))
	for i, dn := range result.DefinedNames {
		scope := dn.Scope
		if scope == "" {
			scope = "Workbook"
		}
		dnOutputs[i] = definedNameOutput{
			Name:  dn.Name,
			Scope: scope,
			Refer: dn.RefersTo,
		}
	}

	out := infoOutput{
		File:         result.FileName,
		DefinedNames: dnOutputs,
		Sheets:       sheetOutputs,
	}

	enc := newJSONLWriter(os.Stdout)
	if err := enc.Encode(out); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}
	return nil
}
