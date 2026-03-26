---
name: exceldump
description: Excelファイル（.xlsx/.xlsm）を読み取る。Excelの内容確認、データ抽出、Excel方眼紙の解析時に使用する。
allowed-tools: Bash(exceldump *)
---

# exceldump

Excelファイル（.xlsx/.xlsm）の内容をCLIからダンプするツール。

## コマンド

### 1. info — ファイルの概要を把握する

```bash
exceldump info <file>
```

シート一覧と定義名を JSON で出力する。まずこのコマンドでファイルの構成を把握し、対象シートを特定する。

### 2. scan — シートの構造を分析する

```bash
exceldump scan --sheet <name|index> <file>
```

デフォルトフォント、列幅等のメタ情報を JSON で出力する。dimension がある場合は使用範囲（`used_range`）とデータ領域（`regions`）も返す。`regions` がある場合は `dump --range` で領域ごとに効率的に取得できる。

### 3. dump — セルデータを取得する

```bash
exceldump dump [options] <file>
```

セルの値をJSONL形式（1行1セル）で出力する。行が変わるタイミングで行情報（`_row`）を出力する。

**オプション:**
- `--sheet <name|index>` — 対象シート（デフォルト: 最初のシート）
- `--range <range>` — セル範囲（例: `A1:H20`, `A:F`, `1:20`）
- `--start <cell>` — 開始セル位置（例: `A51`）。`--range` と排他
- `--limit <n>` — 出力セル数の上限（デフォルト: 1000、0で無制限）
- `--style` — 書式情報（font, fill, border, alignment）を出力
- `--formula` — 数式文字列を出力
- `--include-empty` — 空セルも出力

### 4. search — セル値を検索する

```bash
exceldump search [options] <file>
```

条件に合うセルをJSONL形式で出力する。以下のフィルタのうち少なくとも1つが必須。複数指定時はAND条件。

- `--query <text>` — 部分一致検索（大文字小文字無視）
- `--numeric <expr>` — 数値比較（`">100"`, `"100:200"`, `"=42"`）
- `--type <type>` — 型フィルタ（`string`, `number`, `date`, `bool`, `formula`）

その他: `--sheet`, `--range`, `--start`, `--limit`, `--style`, `--formula`

検索結果が0件の場合は終了コード 1。

## 出力の読み方

### 型の判定

`type` フィールドは `date` と `error` の場合のみ出力される:
- `value` が文字列 → string
- `value` が数値 → number
- `value` が true/false → bool
- `formula` フィールドあり → formula
- `value` なし → empty
- `type: "date"` → 日付（ISO 8601）
- `type: "error"` → エラー値（`#N/A` 等）

### 行情報

行高がデフォルトと異なる、または非表示の行では `_row` 行が挿入される:
```jsonl
{"_row":1,"height":30}
{"cell":"A1","value":"タイトル"}
```

### 書式（`--style` 指定時）

font, fill, border, alignment はデフォルトとの差分のみ出力。`default_font` は `scan` で確認できる。

## 利用フロー

1. `info` でシート一覧を確認
2. `scan` でシート構造を把握（`regions` があれば `--range` で効率的に取得可能）
3. `dump` でデータ取得。`--limit` のデフォルトは1000。続きは最後のセルの次を `--start` に指定
4. 必要に応じて `search` で特定値を検索

`scan` の `regions` がない場合（dimension なしのファイル）は `dump` を直接実行する。
