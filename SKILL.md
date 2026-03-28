---
name: exceldump
description: Excelファイル（.xlsx/.xlsm）を読み取る。Excelの内容確認、データ抽出、Excel方眼紙の解析時に使用する。
user-invocable: false
allowed-tools: Bash(exceldump *)
---

# exceldump

Excelファイル（.xlsx/.xlsm）の内容をCLIからダンプするツール。

## 利用フロー

```
1. info  → シート一覧を確認し対象シートを特定
2. scan  → used_range を取得し dump の --range を決定（任意）
3. dump  → セルデータを取得（先頭に _meta でレイアウト情報を出力）
4. search → 特定値の検索（dump より効率的）
```

scan は used_range の取得に特化。dump の `_meta` 行で列幅・行高を取得できるため、scan を省略して info → dump で直接データ取得も可能。

## コマンドリファレンス

### info

```bash
exceldump info <file>
```

出力例:
```json
{"file":"example.xlsx","defined_names":[],"sheets":[{"index":0,"name":"データ一覧","type":"worksheet"},{"index":1,"name":"設定","type":"worksheet","hidden":true}]}
```

- `sheets[].type`: `worksheet` / `chartsheet`。dump/search は worksheet のみ対応
- `sheets[].hidden`: 非表示シートの場合のみ出力
- `defined_names`: 名前付き範囲の一覧（`name`, `scope`, `refer`）

### scan

```bash
exceldump scan --sheet <name|index> <file>
```

出力例:
```json
{"sheet":"機能一覧","used_range":"A1:H200"}
```

- `used_range`: シートのデータ使用範囲
- dimension（XML属性）があれば即座に返す（数十ms）。なければ全セル走査で算出
- dimension なし（Google Sheets 由来等）で空シートの場合は `used_range` を省略

### dump

```bash
exceldump dump [options] <file>
```

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet <name\|index>` | 対象シート | 最初のシート |
| `--range <range>` | セル範囲（`A1:H20`, `A:F`, `1:20`） | 全体 |
| `--start <cell>` | 開始セル位置（`--range` と排他） | 先頭 |
| `--limit <n>` | 出力セル数の上限（0で無制限） | 1000 |
| `--style` | 書式情報を出力 | OFF |
| `--formula` | 数式文字列を出力 | OFF |
| `--include-empty` | 空セルも出力 | OFF |

出力例:
```jsonl
{"_meta":true,"default_width":8.43,"default_height":15,"col_widths":{"B":24.5,"H":30}}
{"_row":1,"height":24}
{"cell":"A1","value":"項目名"}
{"cell":"B1","value":"数量"}
{"cell":"C1","value":"単価"}
{"_row":2}
{"cell":"A2","value":"商品A","merge":"A2:A3"}
{"cell":"B2","value":100}
```

**`_meta` 行（最初の行）:**

| フィールド | 説明 |
|-----------|------|
| `default_width` | デフォルト列幅（未指定時は Excel 標準値 9.14） |
| `default_height` | デフォルト行高 |
| `col_widths` | デフォルトと異なる列幅のマップ |

`--style` 指定時:
```jsonl
{"cell":"A1","value":"項目名","font":{"bold":true,"color":"#FFFFFF"},"fill":{"color":"#4A86E8"},"alignment":{"horizontal":"center"}}
```

`--formula` 指定時:
```jsonl
{"cell":"D2","value":50000,"formula":"B2*C2"}
```

**続きの取得:**

`--limit` で打ち切られた場合、最後に出力されたセルの次のセルを `--start` に渡す。

```bash
# 最初の1000セル
exceldump dump --sheet 0 example.xlsx
# 出力の最後が {"cell":"A501",...} なら次は B501 から（行優先順）
exceldump dump --sheet 0 --start B501 example.xlsx
```

### search

```bash
exceldump search [options] <file>
```

フィルタ（少なくとも1つ必須、複数指定はAND条件）:

| オプション | 説明 |
|-----------|------|
| `--query <text>` | 部分一致検索（大文字小文字無視。display と value の両方を検索） |
| `--numeric <expr>` | 数値比較。`">100"`, `">=50"`, `"<10"`, `"100:200"`（範囲）, `"=42"`（等値） |
| `--type <type>` | 型フィルタ: `string`, `number`, `date`, `bool`, `formula` |

その他: `--sheet`, `--range`, `--start`, `--limit`, `--style`, `--formula`

出力例:
```bash
exceldump search --query "合計" --sheet 0 example.xlsx
```
```jsonl
{"cell":"A10","value":"合計"}
{"cell":"A25","value":"小合計"}
```

`--numeric` と `--type` は formula セルのキャッシュ値にもヒットする。結果なしでも正常終了（終了コード 0）する。

## 出力形式の詳細

### メタ情報（`_meta`）

dump の最初の行に出力。シートのレイアウト基準値を含む。

### 行情報（`_row`）

行高がデフォルトと異なる、または非表示の行でのみ、セル出力の前に挿入される。デフォルト行高の行では出力されない。

```jsonl
{"_row":1,"height":30}
{"_row":5,"hidden":true}
```

### セルの型の判定

`type` フィールドは `date` と `error` のみ出力。他は JSON 値から判定する:

| 条件 | 型 |
|------|-----|
| `value` が JSON 文字列 | string |
| `value` が JSON 数値 | number |
| `value` が true/false | bool |
| `formula` フィールドあり | formula（`value` はキャッシュ値） |
| `value` なし | empty |
| `type: "date"` | 日付（`value` は ISO 8601: `"2025-03-15"`, `"2025-03-15T10:30:00"`, `"10:30:00"`） |
| `type: "error"` | エラー（`value` は `#N/A`, `#REF!`, `#VALUE!` 等） |

### セルの追加フィールド

| フィールド | 出力条件 | 説明 |
|-----------|---------|------|
| `merge` | 結合セルの場合 | 結合範囲（例: `"B4:B5"`）。左上セルのみ出力される |
| `link` | ハイパーリンクがある場合 | `{"url":"https://..."}` または `{"location":"Sheet2!A1"}` |
| `hidden_col` | 列が非表示の場合 | `true` |
| `display` | 表示文字列がvalueと異なる場合 | 通貨（`¥1,000`）、パーセント（`50%`）、日付の和暦表示等 |

### 書式フィールド（`--style` 指定時）

デフォルトフォント（`_meta` の基準値）との差分のみ出力される。

```jsonl
{"cell":"A1","value":"項目名","font":{"bold":true,"color":"#FFFFFF"},"fill":{"color":"#4A86E8"},"border":{"bottom":{"style":"thin"}},"alignment":{"horizontal":"center","vertical":"center","wrap":true}}
```

- `font`: name, size, bold, italic, strikethrough, underline, color
- `fill`: color（ソリッド塗りつぶしの前景色のみ）
- `border`: top, bottom, left, right, diagonal_up, diagonal_down（各 style + color）
- `alignment`: horizontal, vertical, wrap, indent, text_rotation, shrink_to_fit
- `rich_text`: セル内の一部が異なる書式を持つ場合のラン配列
