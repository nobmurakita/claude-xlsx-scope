# cells コマンド

```bash
xlsx-scope cells [options] <file>
```

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet <name\|index>` | 対象シート | 最初のシート |
| `--range <range>` | セル範囲（`A1:H20`, `A:F`, `1:20`） | 全体 |
| `--start <cell>` | 開始セル位置（`--range` と併用可） | 先頭 |
| `--limit <n>` | 出力セル数の上限（0で無制限） | 1000 |
| `--style` | 書式情報を出力 | OFF |
| `--formula` | 数式文字列を出力 | OFF |
| `--include-empty` | 空セルも出力。表形式で空セルの位置（どの列が空か）が重要な場合に使用 | OFF |

出力例:
```jsonl
{"_meta":true,"default_width":63.23,"default_height":20,"col_widths":{"B:D":183.75,"H":225},"origin":{"x":0,"y":0}}
{"_row":1,"height":32}
{"cell":"A1","value":"項目名"}
{"cell":"B1","value":"数量"}
{"cell":"C1","value":"単価"}
{"cell":"A2","value":"商品A","merge":"A2:A3"}
{"cell":"B2","value":100}
```

## _meta 行（最初の行）

| フィールド | 説明 |
|-----------|------|
| `default_width` | デフォルト列幅（ピクセル） |
| `default_height` | デフォルト行高（ピクセル） |
| `col_widths` | デフォルトと異なる列幅のマップ（ピクセル）。連続する同じ幅の列は `"B:D"` のように範囲表記 |
| `origin` | 起点セルとそのピクセル座標。`shapes` の `pos` と同じ座標系 |

## _row 行

行高がデフォルトと異なる、または非表示の行でのみ、セル出力の前に挿入される。

```jsonl
{"_row":1,"height":40}
{"_row":5,"hidden":true}
```

## セルの範囲まとめ

同一行内で隣接する同内容のセル（`cell` 以外の全フィールドが同一）は `"cell":"A1:C1"` のように範囲表記でまとめられる。

## セルの型の判定

`type` フィールドは通常省略。JSON 値から判定する:

| 条件 | 型 |
|------|-----|
| `value` が JSON 文字列 | string（エラー値 `#N/A`, `#REF!` 等も文字列として出力） |
| `value` が JSON 数値 | number |
| `value` が true/false | bool |
| `formula` フィールドあり | formula（`value` はキャッシュ値） |
| `value` なし | empty |

日付セルは独立した型を持たず、数値セルとして出力される。`value` はシリアル値（数値）、`display` はフォーマット文字列に沿った表示文字列（例: `"2025/3/19"`）、`fmt` にフォーマット文字列（例: `"yyyy/m/d"`）が入る。

## セルの追加フィールド

| フィールド | 出力条件 | 説明 |
|-----------|---------|------|
| `display` | 表示文字列がvalueと異なる場合 | フォーマット文字列に沿った表示文字列（例: `"2025/3/19"`, `"1,234,567"`, `"15%"`） |
| `fmt` | 数値セルにフォーマットがある場合 | 数値フォーマット文字列（例: `"yyyy/m/d"`, `"#,##0"`, `"0%"`） |
| `error` | 値がExcelエラーの場合 | `true`。`#N/A`, `#REF!` 等のエラー値を文字列と区別する |
| `merge` | 結合セルの場合 | 結合範囲（例: `"B4:B5"`）。左上セルのみ出力される |
| `link` | ハイパーリンクがある場合 | `{"url":"https://..."}` または `{"location":"Sheet2!A1"}` |
| `hidden_col` | 列が非表示の場合 | `true` |
| `comment` | コメント/メモがある場合 | `{"author":"著者","text":"本文","thread":[...]}` |

## --style 指定時（スタイル参照化）

書式情報はスタイル定義行（`_style`）として初出時に1回だけ出力され、以降のセルはインデックス `s` で参照する。

```jsonl
{"_style":1,"font":{"bold":true,"color":"#FFFFFF"},"fill":{"color":"#4A86E8"},"alignment":{"horizontal":"center"}}
{"cell":"A1","value":"項目名","s":1}
{"cell":"B1","value":"数量","s":1}
```

- `_style` 行はそのスタイルを使う最初のセルの直前に出力される
- `s` フィールドの値は `_style` 行の値と対応する
- `rich_text` はセル固有の情報のため、スタイル定義ではなくセル行にインライン出力される
- スタイルが全て空のセル（デフォルト書式のみ）は `s` フィールドを持たない

**_style 行のフィールド:**

- `font`: name, size, bold, italic, strikethrough, underline, color
- `fill`: color（ソリッド塗りつぶしの前景色のみ）
- `border`: top, bottom, left, right, diagonal_up, diagonal_down（各 style + color）
- `alignment`: horizontal, vertical, wrap, indent, text_rotation, shrink_to_fit

## --formula 指定時

```jsonl
{"cell":"D2","value":50000,"formula":"B2*C2"}
```

## 続きの取得

`--limit` で打ち切られた場合、最終行に切り捨て通知が出力される。`next_cell` をそのまま `--start` に渡す。`--range` と `--start` は併用できるため、範囲内でのページングも可能。

`_truncated` 行が出力されなければ、残りのデータはない。
