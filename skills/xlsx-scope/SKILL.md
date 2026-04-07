---
name: xlsx-scope
description: Excelファイル（.xlsx/.xlsm）を読み取る。Excelの内容確認、データ抽出、Excel方眼紙の解析時に使用する。
user-invocable: false
allowed-tools:
  - Bash
  - Read
---

# xlsx-scope

Excelファイル（.xlsx/.xlsm）の内容をCLIから出力するツール。
各コマンドの詳細は [references/](references/) 内のファイルを参照。

実行ファイル: `bash ${CLAUDE_SKILL_DIR}/scripts/xlsx-scope <command> [options] <file>`

## 出力の読み取り方

全コマンドの出力は自動的に一時ファイルに保存され、stdout にはファイルパスと行数のみが返る。

```bash
$ bash ${CLAUDE_SKILL_DIR}/scripts/xlsx-scope info example.xlsx
{"file":"/tmp/xlsx-scope-abc123","lines":1}

$ bash ${CLAUDE_SKILL_DIR}/scripts/xlsx-scope cells --sheet 0 example.xlsx
{"file":"/tmp/xlsx-scope-abc456","lines":3482}
```

返された `file` パスを Read で読む（offset: 0始まり行番号, limit: 読む行数）。読み終わったら削除する。

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

`has_shapes: true` のシートでは shapes を必ず実行し、図形の構造を把握する。shapes 出力に `image_id` がある場合、内容の把握に役立つ可能性が高いため積極的に確認する。

1. `image` サブコマンドで画像を取得: `xlsx-scope image <file> <image_id>`
2. 返された `file` パスを Read で画像の内容を確認する
3. 確認が終わったら削除する

**大量データの取得戦略:**

used_range が広いシートでは、まず `--range` で必要な領域を絞って取得する。全体が必要な場合は `--limit`（デフォルト1000）で分割し、`--start` でページングする。

## コマンド一覧

| コマンド | 説明 | 詳細 |
|---------|------|------|
| `info <file>` | シート一覧・名前付き範囲 | |
| `scan --sheet <s> <file>` | シート概要（used_range, value_count 等） | |
| `cells <file>` | セルデータ（書式・結合セル含む） | [cells.md](references/cells.md) |
| `values <file>` | 値のみ行単位出力 | [values.md](references/values.md) |
| `shapes <file>` | 図形・画像 | [shapes.md](references/shapes.md) |
| `image <file> <image_id>` | 画像を一時ファイルに保存 | |
| `search <file>` | セル値の検索 | [search.md](references/search.md) |
