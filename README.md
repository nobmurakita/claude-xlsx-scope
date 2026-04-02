# read-xlsx

Excel ファイル（.xlsx / .xlsm）の内容を CLI から読み取り、JSONL形式で出力するGoツール。
Claude Code 等の AI エージェントが Excel 資料（通常の表、Excel方眼紙の仕様書など）を構造的に読み取る用途を主眼とする。

## インストール

### Claude Code スキルとして使う（推奨）

ワンライナーでインストール:

```bash
curl -fsSL https://raw.githubusercontent.com/nobmurakita/cc-read-xlsx/main/install.sh | bash
```

macOS / Linux / Git Bash（Windows）で動作する。インストール後、Claude Code が Excel ファイルの読み取りが必要な場面で自動的に read-xlsx を使用する。

### Claude Desktop で使う

1. [Releases](https://github.com/nobmurakita/cc-read-xlsx/releases) から最新の zip をダウンロード
2. Claude の設定で「コード実行とファイル作成」が有効になっていることを確認
3. [カスタマイズ > スキル](https://claude.ai/customize/skills) を開く
4. 「+」→「スキルをアップロード」で zip をアップロード

スキルリストに追加されたら、Claude が Excel ファイルの読み取りが必要な場面で自動的に read-xlsx を使用する。

### Go ツールとして使う

```bash
go install github.com/nobmurakita/cc-read-xlsx/cmd/read-xlsx@latest
```

## コマンド

### info — ファイルの概要を表示

```bash
read-xlsx info 基本設計書.xlsx
```

```json
{"file":"基本設計書.xlsx","defined_names":[],"sheets":[{"index":0,"name":"表紙","type":"worksheet"},{"index":1,"name":"機能一覧","type":"worksheet"}]}
```

### scan — シートの使用範囲を取得

```bash
read-xlsx scan --sheet 0 基本設計書.xlsx
```

```json
{"sheet":"表紙","used_range":"A1:CD55","has_shapes":true}
```

dimension（XMLのシート範囲属性）があれば即座に返す。なければ全セル走査で算出する。`has_shapes` は図形が存在するシートでのみ `true` を出力する。

### cells — セルデータを出力

```bash
read-xlsx cells --sheet 0 --limit 5 見積計算.xlsx
```

```jsonl
{"_meta":true,"default_width":63.23,"default_height":20,"col_widths":{"B:D":183.75,"H":225},"origin":{"x":0,"y":0}}
{"cell":"A1","value":"項目名"}
{"cell":"B1","value":"数量"}
{"cell":"C1","value":"単価"}
{"_row":2,"height":30}
{"cell":"A2","value":"商品A","merge":"A2:A3"}
{"_truncated":true,"next_cell":"B2"}
```

最初の行に `_meta`（レイアウト情報）を出力し、その後にセルデータが続く。幅・高さの値はすべてピクセル単位（96 DPI 基準）。`origin` は出力の起点セルとそのピクセル座標で、`shapes` コマンドの `pos` と同じ座標系。同一行内で隣接する同内容セルは `"cell":"A1:C1"` のように範囲表記でまとめられる。`--limit` で打ち切られた場合は最終行に `_truncated` が出力され、`next_cell` を `--start` に渡して続きを取得できる。

書式付き:

```bash
read-xlsx cells --sheet 0 --style --range "B3:K4" --limit 3 見積計算.xlsx
```

```jsonl
{"_meta":true,"default_width":63.23,"default_height":20,"col_widths":{"B:D":183.75},"origin":{"x":63,"y":40}}
{"_row":3,"height":30}
{"cell":"B3","value":"工程","font":{"color":"#FFFFFF"},"fill":{"color":"#4A86E8"},"alignment":{"vertical":"center"}}
{"cell":"C3","value":"作業内容","font":{"color":"#FFFFFF"},"fill":{"color":"#4A86E8"},"alignment":{"vertical":"center"}}
{"cell":"D3","value":"成果物","font":{"color":"#FFFFFF"},"fill":{"color":"#4A86E8"},"alignment":{"vertical":"center"}}
```

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet` | 対象シート（名前 or 0始まりインデックス） | 最初のシート |
| `--range` | セル範囲（例: `A1:H20`, `A:F`, `1:20`） | 全体 |
| `--start` | 開始セル位置（例: `A51`）。`--range` と併用可 | 先頭 |
| `--include-empty` | 空セルも出力する | OFF |
| `--style` | 書式情報を出力する | OFF |
| `--formula` | 数式文字列を出力する | OFF |
| `--limit` | 出力セル数の上限（0で無制限） | 1000 |

### shapes — 図形・フローチャート・画像を取得

```bash
read-xlsx shapes --sheet 0 処理フロー.xlsx
```

```jsonl
{"_meta":true,"shape_count":8,"connector_count":3}
{"id":1,"type":"flowChartProcess","text":"処理A","cell":"B2:D4","pos":{"x":120,"y":80,"w":200,"h":60},"z":0}
{"id":2,"type":"flowChartDecision","text":"条件分岐","cell":"B6:D8","pos":{"x":120,"y":200,"w":200,"h":80},"z":1}
{"id":3,"type":"connector","cell":"B4:B6","pos":{"x":220,"y":140,"w":0,"h":60},"from":1,"to":2,"from_idx":2,"to_idx":0,"connector_type":"bentConnector3","adj":{"adj1":50000},"arrow":"end","start":{"x":220,"y":140},"end":{"x":220,"y":200},"z":2}
```

`pos` はピクセル座標（96 DPI 基準、左上原点）。コネクタは `start`/`end` で両端座標を出力する。吹き出し形状は `callout_target` でポインタ先座標を出力する。

画像は `image_id` フィールドで参照される:

```jsonl
{"id":10,"type":"picture","name":"図 1","cell":"B2:F8","pos":{"x":120,"y":80,"w":640,"h":480},"z":5,"alt_text":"構成図","image_id":"xl/media/image1.png"}
```

`image_id` を `image` サブコマンドに渡すと画像を取得できる:

```bash
read-xlsx image 処理フロー.xlsx xl/media/image1.png
# stdout: /var/folders/.../read-xlsx-1234567.png
```

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet` | 対象シート（名前 or 0始まりインデックス） | 最初のシート |
| `--limit` | 出力図形数の上限（0で無制限） | 1000 |
| `--style` | 書式情報を出力する | OFF |

### image — 画像を取得

```bash
read-xlsx image <file> <image_id>
```

`shapes` 出力の `image_id` を指定して画像を一時ファイルに保存する。パスが stdout に出力される。

```bash
read-xlsx image 処理フロー.xlsx xl/media/image1.png
# stdout: /var/folders/.../read-xlsx-1234567.png
```

### search — セル値を検索

```bash
read-xlsx search --text "マスタ" 運用シナリオ.xlsx
```

```jsonl
{"cell":"B2","value":"マスタを登録する"}
{"cell":"D2","value":"マスタファイル"}
```

```bash
read-xlsx search --numeric ">100" 見積計算.xlsx
```

```jsonl
{"cell":"K4","value":800}
{"cell":"K6","value":800}
{"cell":"K8","value":400}
```

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--text` | 検索文字列（部分一致、大文字小文字無視） | — |
| `--numeric` | 数値比較（`">100"`, `"100:200"`, `"=42"`） | — |
| `--type` | 型フィルタ（`string`, `number`, `bool`, `formula`） | — |
| `--sheet` | 対象シート | 最初のシート |
| `--range` | セル範囲 | 全体 |
| `--start` | 開始セル位置。`--range` と併用可 | 先頭 |
| `--style` | 書式情報を出力する | OFF |
| `--formula` | 数式文字列を出力する | OFF |
| `--limit` | 出力セル数の上限（0で無制限） | 1000 |

`--text`, `--numeric`, `--type` のうち少なくとも1つが必須。複数指定時は AND 条件。

出力形式やフィールドの詳細は [DESIGN.md](DESIGN.md) を参照。
