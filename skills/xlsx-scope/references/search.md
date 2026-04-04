# search コマンド

```bash
xlsx-scope search [options] <file>
```

フィルタ（少なくとも1つ必須、複数指定はAND条件）:

| オプション | 説明 |
|-----------|------|
| `--text <text>` | 部分一致検索（大文字小文字無視。display と value の両方を検索） |
| `--numeric <expr>` | 数値比較。`">100"`, `">=50"`, `"<10"`, `"100:200"`（範囲）, `"=42"`（等値） |
| `--type <type>` | 型フィルタ: `string`, `number`, `bool`, `formula` |

その他: `--sheet`, `--range`, `--start`, `--limit`, `--style`, `--formula`

出力例:
```jsonl
{"cell":"A10","value":"合計"}
{"cell":"A25","value":"小合計"}
```

- `--numeric` と `--type` は formula セルのキャッシュ値にもヒットする
- 結果なしでも正常終了（終了コード 0）
- 出力フィールドは cells コマンドと同じ（[cells.md](cells.md) 参照）
