# exceldump

Excel ファイル（.xlsx / .xlsm）の内容を CLI からダンプするGoツール。
Claude Code が Excel 資料（通常の表、Excel方眼紙の仕様書など）を読み取る用途を主眼とする。

## インストール

```bash
go install github.com/nobmurakita/exceldump@latest
```

ビルド時にバージョンを埋め込む場合:

```bash
go build -ldflags "-X main.version=0.1.0" -o exceldump .
```

### Claude Code スキルのインストール

GitHub から直接インストール:

```bash
mkdir -p ~/.claude/skills/exceldump
curl -fsSL https://raw.githubusercontent.com/nobmurakita/exceldump/main/SKILL.md -o ~/.claude/skills/exceldump/SKILL.md
```

またはローカルからコピー:

```bash
mkdir -p ~/.claude/skills/exceldump
cp SKILL.md ~/.claude/skills/exceldump/SKILL.md
```

インストール後、Claude Code が Excel ファイルの読み取りが必要な場面で自動的に exceldump を使用する。

## コマンド

### info — ファイルの概要を表示

```bash
exceldump info 基本設計書.xlsx
```

```json
{"file":"基本設計書.xlsx","defined_names":[],"sheets":[{"index":0,"name":"表紙","type":"worksheet"},{"index":1,"name":"機能一覧","type":"worksheet"}]}
```

### scan — シートの使用範囲を取得

```bash
exceldump scan --sheet 0 基本設計書.xlsx
```

```json
{"sheet":"表紙","used_range":"A1:CD55","has_drawings":true}
```

dimension（XMLのシート範囲属性）があれば即座に返す。なければ全セル走査で算出する。`has_drawings` は図形が存在するシートでのみ `true` を出力する。

### dump — セルデータをダンプ

```bash
exceldump dump --sheet 0 --limit 5 見積計算.xlsx
```

```jsonl
{"_meta":true,"default_width":8.43,"default_height":15,"col_widths":{"B":24.5,"H":30}}
{"cell":"A1","value":"項目名"}
{"cell":"B1","value":"数量"}
{"cell":"C1","value":"単価"}
{"_row":2,"height":22.5}
{"cell":"A2","value":"商品A","merge":"A2:A3"}
{"_truncated":true,"next_cell":"B2"}
```

最初の行に `_meta`（レイアウト情報）を出力し、その後にセルデータが続く。`--limit` で打ち切られた場合は最終行に `_truncated` が出力され、`next_cell` を `--start` に渡して続きを取得できる。

書式付き:

```bash
exceldump dump --sheet 0 --style --range "B3:K4" --limit 3 見積計算.xlsx
```

```jsonl
{"_meta":true,"default_width":8.43,"default_height":15,"col_widths":{"B":24.5}}
{"_row":3,"height":22.5}
{"cell":"B3","value":"工程","font":{"color":"#FFFFFF"},"fill":{"color":"#4A86E8"},"alignment":{"vertical":"center"}}
{"cell":"C3","value":"作業内容","font":{"color":"#FFFFFF"},"fill":{"color":"#4A86E8"},"alignment":{"vertical":"center"}}
{"cell":"D3","value":"成果物","font":{"color":"#FFFFFF"},"fill":{"color":"#4A86E8"},"alignment":{"vertical":"center"}}
```

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet` | 対象シート（名前 or 0始まりインデックス） | 最初のシート |
| `--range` | セル範囲（例: `A1:H20`, `A:F`, `1:20`） | 全体 |
| `--start` | 開始セル位置（例: `A51`）。`--range` と排他 | 先頭 |
| `--include-empty` | 空セルも出力する | OFF |
| `--style` | 書式情報を出力する | OFF |
| `--formula` | 数式文字列を出力する | OFF |
| `--limit` | 出力セル数の上限（0で無制限） | 1000 |

### shapes — 図形・フローチャート・画像を取得

```bash
exceldump shapes --sheet 0 処理フロー.xlsx
```

```jsonl
{"_meta":true,"shape_count":8,"connector_count":3}
{"id":1,"type":"flowChartProcess","text":"処理A","cell":"B2:D4","z":0}
{"id":2,"type":"flowChartDecision","text":"条件分岐","cell":"B6:D8","z":1}
{"id":3,"type":"connector","from":1,"to":2,"connector_type":"straightConnector1","arrow":"end","z":2}
```

画像を抽出する場合:

```bash
exceldump shapes --sheet 0 --extract-images /tmp/images 設計書.xlsx
```

```jsonl
{"id":10,"type":"picture","name":"図 1","cell":"B2:F8","z":5,"alt_text":"構成図","image":{"format":"png","width":640,"height":480,"size":45230,"path":"/tmp/images/image_1.png"}}
```

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet` | 対象シート（名前 or 0始まりインデックス） | 最初のシート |
| `--limit` | 出力図形数の上限（0で無制限） | 1000 |
| `--style` | 書式情報を出力する | OFF |
| `--extract-images` | 画像を指定ディレクトリに抽出する | OFF（画像スキップ） |

### search — セル値を検索

```bash
exceldump search --query "マスタ" 運用シナリオ.xlsx
```

```jsonl
{"cell":"B2","value":"マスタを登録する"}
{"cell":"D2","value":"マスタファイル"}
```

```bash
exceldump search --numeric ">100" 見積計算.xlsx
```

```jsonl
{"cell":"K4","value":800}
{"cell":"K6","value":800}
{"cell":"K8","value":400}
```

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--query` | 検索文字列（部分一致、大文字小文字無視） | — |
| `--numeric` | 数値比較（`">100"`, `"100:200"`, `"=42"`） | — |
| `--type` | 型フィルタ（`string`, `number`, `date`, `bool`, `formula`） | — |
| `--sheet` | 対象シート | 最初のシート |
| `--range` | セル範囲 | 全体 |
| `--start` | 開始セル位置。`--range` と排他 | 先頭 |
| `--style` | 書式情報を出力する | OFF |
| `--formula` | 数式文字列を出力する | OFF |
| `--limit` | 出力セル数の上限（0で無制限） | 1000 |

`--query`, `--numeric`, `--type` のうち少なくとも1つが必須。複数指定時は AND 条件。

出力形式やフィールドの詳細は [DESIGN.md](DESIGN.md) を参照。
