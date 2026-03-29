---
name: cc-read-xlsx
description: Excelファイル（.xlsx/.xlsm）を読み取る。Excelの内容確認、データ抽出、Excel方眼紙の解析時に使用する。
user-invocable: false
allowed-tools: Bash(cc-read-xlsx *)
---

# cc-read-xlsx

Excelファイル（.xlsx/.xlsm）の内容をCLIから出力するツール。

## 利用フロー

```
1. info   → シート一覧を確認し対象シートを特定
2. scan   → used_range と has_drawings を取得（任意）
3. cells  → セルデータを取得（先頭に _meta でレイアウト情報を出力）
4. shapes → 図形・フローチャート・画像を取得（has_drawings: true のシートに対して）
5.          → 画像があれば出力の image.path を Read で確認し、一時ディレクトリを削除
6. search → 特定値の検索（cells より効率的）
```

scan は used_range の取得に特化。cells の `_meta` 行で列幅・行高を取得できるため、scan を省略して info → cells で直接データ取得も可能。
図形がある場合は shapes で構造を把握する。画像は自動的に一時ディレクトリに抽出される。

**画像の確認手順:** `scan` の結果で `has_drawings: true` のシートがある場合、`shapes` で画像を含む図形情報を取得できる。画像がある場合は以下の手順で確認すること:

1. `shapes` コマンドを実行する（画像は一時ディレクトリに自動抽出される）
2. 出力の `image.path` を Read ツールで読み、画像の内容を確認する
3. 確認が終わったら `image.path` の親ディレクトリを削除する

**書式情報（`--style`）の取得判断:**

罫線・背景色・フォント等の書式情報はデフォルトでは出力されない。書式がレイアウトの理解に必要かどうかは実際に見るまで判断できないため、各シートの初回 cells 取得時は `--style` をつけて取得し、書式の有無を確認する。書式が構造の把握に不要と判断したシートでは、以降は `--style` を外して取得する。

**大量データの取得戦略:**

used_range が広いシートでは、まず `--range` で必要な領域を絞って取得する。全体が必要な場合は `--limit`（デフォルト1000）で分割し、`--start` でページングする。

**空セルの扱い（`--include-empty`）:**

デフォルトでは空セルは出力されない。表形式のデータで空セルの位置（どの列が空か）が重要な場合は `--include-empty` をつけて取得する。

## コマンドリファレンス

### info

```bash
cc-read-xlsx info <file>
```

出力例:
```json
{"file":"example.xlsx","defined_names":[],"sheets":[{"index":0,"name":"データ一覧","type":"worksheet"},{"index":1,"name":"設定","type":"worksheet","hidden":true}]}
```

- `sheets[].type`: `worksheet` / `chartsheet`。cells/search は worksheet のみ対応
- `sheets[].hidden`: 非表示シートの場合のみ出力
- `defined_names`: 名前付き範囲の一覧（`name`, `scope`, `refer`）

### scan

```bash
cc-read-xlsx scan --sheet <name|index> <file>
```

出力例:
```json
{"sheet":"機能一覧","used_range":"A1:H200","has_drawings":true}
```

- `used_range`: シートのデータ使用範囲。空シートの場合は省略
- `has_drawings`: 図形が存在する場合のみ `true`。`shapes` コマンドを使うべきか判断に使う

### cells

```bash
cc-read-xlsx cells [options] <file>
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
{"cell":"A2","value":"商品A","merge":"A2:A3"}
{"cell":"B2","value":100}
```

**`_meta` 行（最初の行）:**

| フィールド | 説明 |
|-----------|------|
| `default_width` | デフォルト列幅 |
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

`--limit` で打ち切られた場合、最終行に切り捨て通知が出力される。`next_cell` をそのまま `--start` に渡す。

```bash
# 最初の1000セル
cc-read-xlsx cells --sheet 0 example.xlsx
# 最終行: {"_truncated":true,"next_cell":"B501"}
# 続きを取得
cc-read-xlsx cells --sheet 0 --start B501 example.xlsx
```

`_truncated` 行が出力されなければ、残りのデータはない。

### shapes

```bash
cc-read-xlsx shapes [options] <file>
```

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet <name\|index>` | 対象シート | 最初のシート |
| `--limit <n>` | 出力図形数の上限（0で無制限） | 1000 |
| `--style` | 書式情報を出力 | OFF |

出力例:
```jsonl
{"_meta":true,"shape_count":8,"connector_count":3}
{"id":1,"type":"flowChartProcess","text":"処理A","cell":"B2:D4","z":0}
{"id":2,"type":"flowChartDecision","text":"条件分岐","cell":"B6:D8","z":1}
{"id":3,"type":"connector","from":1,"to":2,"connector_type":"straightConnector1","arrow":"end","z":2}
```

画像は自動的に一時ディレクトリに抽出される:
```jsonl
{"id":10,"type":"picture","name":"図 1","cell":"B2:F8","z":5,"alt_text":"構成図","image":{"format":"png","width":640,"height":480,"size":45230,"path":"/tmp/cc-read-xlsx-images-xxx/image_abc.png"}}
```

**図形種別:**

- シェイプ: `rect`, `roundRect`, `ellipse`, `flowChartProcess`, `flowChartDecision`, `flowChartTerminator` 等
- コネクタ: `type` は常に `"connector"`。`from`/`to` で接続先の図形IDを参照。`connector_type` でコネクタ形状
- グループ: `type` は `"group"`。`children` に子要素ID配列。子要素は `parent` で親を参照
- 画像: `type` は `"picture"`。一時ディレクトリに自動抽出される

**図形内テキスト:**

- `text`: プレーンテキスト（複数段落は `\n` で結合）
- `rich_text`: 書式の異なるランがある場合のみ出力（`--style` 指定時）
- コネクタのテキストは `label` フィールド

**画像の確認方法:**

画像は自動抽出されるため、出力の `image.path` を Read ツールで読むことで画像の中身を視覚的に確認できる。

### search

```bash
cc-read-xlsx search [options] <file>
```

フィルタ（少なくとも1つ必須、複数指定はAND条件）:

| オプション | 説明 |
|-----------|------|
| `--text <text>` | 部分一致検索（大文字小文字無視。display と value の両方を検索） |
| `--numeric <expr>` | 数値比較。`">100"`, `">=50"`, `"<10"`, `"100:200"`（範囲）, `"=42"`（等値） |
| `--type <type>` | 型フィルタ: `string`, `number`, `bool`, `formula` |

その他: `--sheet`, `--range`, `--start`, `--limit`, `--style`, `--formula`

出力例:
```bash
cc-read-xlsx search --text "合計" --sheet 0 example.xlsx
```
```jsonl
{"cell":"A10","value":"合計"}
{"cell":"A25","value":"小合計"}
```

`--numeric` と `--type` は formula セルのキャッシュ値にもヒットする。結果なしでも正常終了（終了コード 0）する。

## 出力形式の詳細

### メタ情報（`_meta`）

cells の最初の行に出力。シートのレイアウト基準値を含む。

### 行情報（`_row`）

行高がデフォルトと異なる、または非表示の行でのみ、セル出力の前に挿入される。デフォルト行高の行では出力されない。

```jsonl
{"_row":1,"height":30}
{"_row":5,"hidden":true}
```

### セルの型の判定

`type` フィールドは通常省略。JSON 値から判定する:

| 条件 | 型 |
|------|-----|
| `value` が JSON 文字列 | string（エラー値 `#N/A`, `#REF!` 等も文字列として出力） |
| `value` が JSON 数値 | number |
| `value` が true/false | bool |
| `formula` フィールドあり | formula（`value` はキャッシュ値） |
| `value` なし | empty |

日付セルは独立した型を持たず、数値セルとして出力される。`value` はシリアル値（数値）、`display` はフォーマット文字列に沿った表示文字列（例: `"2025/3/19"`）、`fmt` にフォーマット文字列（例: `"yyyy/m/d"`）が入る。`fmt` から日付かどうかを判断できる。

### セルの追加フィールド

| フィールド | 出力条件 | 説明 |
|-----------|---------|------|
| `display` | 表示文字列がvalueと異なる場合 | フォーマット文字列に沿った表示文字列（例: `"2025/3/19"`, `"1,234,567"`, `"15%"`） |
| `fmt` | 数値セルにフォーマットがある場合 | 数値フォーマット文字列（例: `"yyyy/m/d"`, `"#,##0"`, `"0%"`） |
| `error` | 値がExcelエラーの場合 | `true`。`#N/A`, `#REF!` 等のエラー値を文字列と区別する |
| `merge` | 結合セルの場合 | 結合範囲（例: `"B4:B5"`）。左上セルのみ出力される |
| `link` | ハイパーリンクがある場合 | `{"url":"https://..."}` または `{"location":"Sheet2!A1"}` |
| `hidden_col` | 列が非表示の場合 | `true` |
| `comment` | コメント/メモがある場合 | `{"author":"著者","text":"本文","thread":[...]}` |

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
