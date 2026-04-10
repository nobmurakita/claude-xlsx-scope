---
name: xlsx-scope
description: |-
  Excelファイル(.xlsx/.xlsm)を読み取り値/式/書式/図/画像などの情報をJSONLで出力する。
  データ抽出、Excel方眼紙の仕様書・設計書の解析、図形・フローチャートの構造把握に使用する。
allowed-tools:
  - Bash(*/.claude/skills/xlsx-scope/scripts/xlsx-scope *)
  - Read(*/xlsx-scope-tmp-*)
---

# xlsx-scope

Excelファイル（.xlsx/.xlsm）の内容をCLIから出力するツール。

実行ファイル: `bash ${CLAUDE_SKILL_DIR}/scripts/xlsx-scope <command> [options] <file>`

## 出力の読み取り方

全コマンドの出力は自動的に一時ファイル（プレフィックス `xlsx-scope-tmp-`）に保存され、stdout にはファイルパスと行数のみが返る。

```bash
$ bash ${CLAUDE_SKILL_DIR}/scripts/xlsx-scope info example.xlsx
{"file":"$TMPDIR/xlsx-scope-tmp-abc123","lines":1}

$ bash ${CLAUDE_SKILL_DIR}/scripts/xlsx-scope cells --sheet 0 example.xlsx
{"file":"$TMPDIR/xlsx-scope-tmp-abc456","lines":3482}
```

返された `file` パスを Read で読む（offset: 0始まり行番号, limit: 読む行数）。読み終わったら都度 `cleanup` サブコマンドで削除する。

## 利用フロー

```
1. info   → シート一覧を確認し対象シートを特定
2. scan   → used_range, value_count, merged_cells, has_shapes, style_variants を取得
3. cells  → セルデータを取得（書式・結合セル等のレイアウト情報が必要な場合）
4. values → 値のみを行単位で取得（書式不要と判断したデータシート向け）
5. shapes → 図形・フローチャート・画像を取得（has_shapes: true のシートで必ず実行）
6. image  → image_id がある図形の画像を取得して確認
7. search → 特定値の検索（cells より効率的）
```

scan は各シートに対して基本的に実行する。各指標を総合してシートの性質（データ/ドキュメント/図）を推測し、後続のコマンドとオプションを判断する。

**書式情報（`--style`）の取得判断:**

罫線・背景色・フォント等の書式情報はデフォルトでは出力されない。scan の結果を総合して判断する:
- `style_variants` が多い、`merged_cells` が多い → 書式がレイアウトの理解に重要な可能性が高い
- `style_variants: 0` → 視覚的書式なし。`--style` 不要

**図形・画像の確認:**

`has_shapes: true` のシートでは shapes を必ず実行し、図形の構造を把握する。shapes 出力に `image_id` がある場合、内容の把握に役立つ可能性が高いため積極的に `image` サブコマンドで取得して確認する。

**大量データの取得戦略:**

used_range が広いシートでは、まず `--range` で必要な領域を絞って取得する。全体が必要な場合は `--limit`（デフォルト1000）で分割し、`--start` でページングする。

## コマンド詳細

### info

`xlsx-scope info <file>` — シート一覧・名前付き範囲を出力。

出力例:
```json
{"file":"基本設計書.xlsx","defined_names":[{"name":"マスタ","scope":"Workbook","refer":"Sheet1!$A$1:$D$100"}],"sheets":[{"index":0,"name":"表紙","type":"worksheet"},{"index":1,"name":"機能一覧","type":"worksheet"},{"index":2,"name":"非表示データ","type":"worksheet","hidden":true}]}
```

### scan

`xlsx-scope scan --sheet <name|index> <file>` — シートの構造概要を出力。

出力例:
```json
{"sheet":"機能一覧","used_range":"A1:H200","value_count":1520,"merged_cells":42,"style_variants":8,"has_shapes":true}
```

### cells

```
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
| `--include-empty` | 空セルも出力 | OFF |

出力例:
```jsonl
{"meta":true,"default_width":63.23,"default_height":20,"col_widths":{"B:D":183.75,"H":225},"origin":{"x":0,"y":0}}
{"row":1,"height":32}
{"cell":"A1","value":"項目名"}
{"cell":"B1","value":"数量"}
{"cell":"A2","value":"商品A","merge":"A2:A3"}
{"cell":"B2","value":100,"display":"100.00","fmt":"0.00"}
```

**meta 行:** レイアウト情報。`default_width`/`default_height`（ポイント）、`col_widths`（デフォルトと異なる列幅）、`origin`（起点座標、`shapes` の `pos` と同じ座標系）。

**row 行:** 行高がデフォルトと異なる、または非表示の行でのみ出力。`height`（行高、ポイント。デフォルトと異なる場合のみ）、`hidden`（非表示の場合のみ `true`）。

**セルのフィールド:**

| フィールド | 説明 |
|-----------|------|
| `cell` | セル位置。隣接する同内容セルは `"A1:C1"` のように範囲表記 |
| `value` | セル値（文字列/数値/真偽値）。日付はシリアル値 |
| `display` | フォーマット済み表示文字列（value と異なる場合のみ） |
| `type` | 通常は省略（JSONの型や `formula` の有無から推測可能）。`--include-empty` 時に `"empty"` |
| `fmt` | 数値フォーマット文字列（例: `"yyyy/m/d"`, `"#,##0"`） |
| `formula` | 数式文字列（`--formula` 指定時のみ。`value` はキャッシュ値） |
| `error` | `true`（値が `#N/A`, `#REF!` 等のエラーの場合） |
| `merge` | 結合範囲（左上セルのみ出力） |
| `link` | `{"url":"..."}` または `{"location":"Sheet2!A1"}` |
| `hidden_col` | `true`（列が非表示の場合） |
| `comment` | `{"author":"著者","text":"本文","thread":[...]}` |

**--style 指定時:** あるスタイルが初めて登場したとき `style` 定義行が出力され、以降の同一スタイルのセルは `s` で参照番号のみを出力する。

`style` 定義行のフィールド: `font`, `fill`, `border`, `alignment`（デフォルトと異なるもののみ）。

```jsonl
{"style":1,"font":{"bold":true,"color":"#FFFFFF"},"fill":{"color":"#4A86E8"},"alignment":{"horizontal":"center"}}
{"cell":"A1","value":"項目名","s":1}
{"cell":"A2","value":"データ","s":1}
```

`rich_text` はセル固有のため `style` には含まれず、セル行にインライン出力される。

**続きの取得:** `{"truncated":true,"next_cell":"..."}` が出力されたら `next_cell` を `--start` に渡す。

### values

```
xlsx-scope values [options] <file>
```

値のみを行単位で出力。書式・レイアウト情報を含まないデータシート向け。

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet <name\|index>` | 対象シート | 最初のシート |
| `--range <range>` | セル範囲（`A1:H20`, `A:F`, `1:20`） | 全体 |
| `--start <row>` | 開始行番号（1始まり） | 先頭 |
| `--limit <n>` | 出力行数の上限（0で無制限） | 1000 |

出力例:
```jsonl
{"meta":true,"cols":["A","B","C","D"]}
{"row":1,"values":["ID","名前","部署","入社日"]}
{"row":2,"values":[1,"田中太郎","営業部","2025/4/1"]}
{"row":3,"values":[2,"鈴木花子",null,"2025/4/15"]}
```

`meta.cols` は values 配列の各インデックスに対応する列名。空行はスキップ、末尾の null はトリム。

**続きの取得:** `{"truncated":true,"next_row":101}` の `next_row` を `--start` に渡す。

### shapes

```
xlsx-scope shapes [options] <file>
```

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet <name\|index>` | 対象シート | 最初のシート |
| `--limit <n>` | 出力図形数の上限（0で無制限） | 1000 |
| `--style` | 書式情報を出力 | OFF |

座標（`pos`, `start`, `end`, `callout_target`）の単位はすべてpt。`cells` の `meta.origin` と同じ座標系。

出力例:
```jsonl
{"meta":true,"shape_count":8,"connector_count":3}
{"id":1,"type":"roundRect","text":"処理A","cell":"B2:D4","pos":{"x":120,"y":80,"w":200,"h":60},"adj":{"adj1":16667},"z":0}
{"id":2,"type":"flowChartDecision","text":"条件分岐","cell":"B6:D8","pos":{"x":120,"y":200,"w":200,"h":80},"z":1}
{"id":3,"type":"connector","cell":"B4:B6","pos":{"x":220,"y":140,"w":0,"h":60},"from":1,"to":2,"from_idx":2,"to_idx":0,"connector_type":"bentConnector3","adj":{"adj1":50000},"arrow":"end","start":{"x":220,"y":140},"end":{"x":220,"y":200},"z":2}
```

**図形種別:**
- シェイプ: `rect`, `roundRect`, `ellipse`, `flowChartProcess`, `flowChartDecision` 等
- 吹き出し: `wedgeRectCallout` 等。`callout_target` でポインタ先を出力
- コネクタ: `type` は `"connector"`。`from`/`to` で接続先図形ID、`start`/`end` で両端座標（pt）
- グループ: `type` は `"group"`。`children` に子要素ID配列
- 画像: `type` は `"picture"`。`image_id` で `image` サブコマンドにより取得可能

**図形内テキスト:** `text` にプレーンテキスト（複数段落は `\n` 結合）。コネクタのテキストは `label`。

### image

`xlsx-scope image <file> <image_id>` — 画像を一時ファイルに保存。

shapes 出力の `image_id` を指定する。stdout に `{"file":"$TMPDIR/xlsx-scope-tmp-abc123.png"}` が返る。

### search

```
xlsx-scope search [options] <file>
```

フィルタ（少なくとも1つ必須、複数指定はAND条件）:

| オプション | 説明 |
|-----------|------|
| `--text <text>` | 部分一致検索（大文字小文字無視。display と value の両方を検索） |
| `--numeric <expr>` | 数値比較。`">100"`, `">=50"`, `"<10"`, `"100:200"`（範囲）, `"=42"`（等値） |
| `--type <type>` | 型フィルタ: `string`, `number`, `bool`, `formula` |

その他: `--sheet`, `--range`, `--start`, `--limit`, `--style`, `--formula`

出力フィールドは cells と同じ。結果なしでも正常終了（終了コード 0）。

### cleanup

`xlsx-scope cleanup <file> [file...]` — xlsx-scope が生成した一時ファイルを削除する。

```bash
$ bash ${CLAUDE_SKILL_DIR}/scripts/xlsx-scope cleanup /tmp/xlsx-scope-tmp-abc123 /tmp/xlsx-scope-tmp-def456
{"deleted":2}
```
