# exceldump 実装計画

## ディレクトリ構成

```
exceldump/
├── DESIGN.md
├── PLAN.md
├── go.mod
├── go.sum
├── main.go                          # エントリーポイント（Execute()を呼ぶだけ）
├── root.go                          # ルートコマンド定義（cobra）、エラー処理、終了コード制御
├── info.go                          # info サブコマンド
├── scan.go                          # scan サブコマンド
├── dump.go                          # dump サブコマンド
├── search.go                        # search サブコマンド
├── version.go                       # version サブコマンド
├── internal/excel/
│   ├── open.go                      # ファイルオープン・バリデーション（拡張子チェック）
│   ├── sheet.go                     # シート情報取得（一覧、種別判定、表示状態、シート解決）
│   ├── cell.go                      # セル値の読み取り・型判定・display生成
│   ├── range.go                     # 範囲パース・走査（A1:H20, A:F, 1:20, 単一セル）、--start処理
│   ├── merge.go                     # 結合セル情報の取得・左上セル判定
│   ├── region.go                    # scan用の領域分割ロジック（行バンド×列バンド→矩形候補）
│   ├── formula.go                   # 数式取得（共有数式、配列数式対応）
│   ├── style.go                     # font/fill/border/alignment オブジェクト生成・差分計算
│   ├── richtext.go                  # リッチテキストラン処理
│   ├── defaultfont.go               # デフォルトフォント検出ロジック（列スタイル頻度カウント）
│   ├── color.go                     # テーマカラー→HEX RGB変換、tint適用（HSL明度調整）
│   ├── filter.go                    # search用フィルタ（--query, --numeric, --type）+ numericパーサー
│   └── truncation.go               # 切り捨て通知構造体
└── testdata/                        # テスト用Excelファイル
```

## 実装フェーズ

### フェーズ 1: プロジェクト骨格

**目的:** ビルド可能な状態を作り、最小限のコマンドを動作させる。

**実装順序:**

1. `go.mod` — モジュール初期化。excelize v2 と cobra を依存追加
2. `main.go` — エントリーポイント
3. `root.go` — cobra ルートコマンド定義（`SilenceUsage: true`, `SilenceErrors: true`）、エラー処理（`exceldump: <メッセージ>` 形式の stderr 出力）、終了コード定数（0/1/2）
4. `version.go` — version サブコマンド。`-ldflags` でバージョン埋め込み、未設定時 `dev`

**完了条件:** `go build` が成功し `exceldump version` が動作する。

---

### フェーズ 2: Excel操作基盤

**目的:** excelize を使ったファイル読み込み、シート解決、範囲パース等の共通操作層を構築する。

**実装順序:**

1. `internal/excel/open.go` — ファイルオープン。拡張子チェック（.xlsx/.xlsm のみ）、`excelize.OpenFile()` ラッパー、パスワード保護・破損ファイルのエラーハンドリング
2. `internal/excel/sheet.go` — シート一覧取得、名前/インデックスによるシート解決（`--sheet` 共通処理）、シート種別判定、表示状態取得、ワークシート以外指定時のエラー
3. `internal/excel/range.go` — 範囲文字列パーサー（`A1:H20`, `A:F`, `1:20`, `B5`）、`used_range` による行/列補完、`--start` の走査開始位置計算、`--range`と`--start`の排他チェック、不正範囲バリデーション
4. `internal/excel/merge.go` — 結合セル情報取得、左上セル判定、結合範囲文字列取得
5. `internal/excel/cell.go` — セル値読み取り、型判定（excelize の型情報＋数値フォーマットによる日付判定）、日付のISO 8601変換、display フィールド生成、エラー値判定、bool の display
6. `internal/excel/formula.go` — 数式文字列取得、共有数式の展開、配列数式の波括弧付き出力、キャッシュ値なし時の null 処理

**完了条件:** 共通Excel操作のユニットテストが通る。

---

### フェーズ 3: 書式処理

**目的:** セルの書式情報（font, fill, border, alignment, rich_text）の取得とデフォルト差分出力を実装する。

**実装順序:**

1. `internal/excel/color.go` — テーマカラーからHEX RGB変換、tint適用（HSL明度調整）、変換不能時の生値フォールバック
2. `internal/excel/defaultfont.go` — デフォルトフォント検出。`used_range` 内の列スタイルからフォント頻度カウント、最頻フォント採用（同数時は列インデックスが小さい方優先）、列スタイル未設定時はブックデフォルトにフォールバック
3. `internal/excel/style.go` — font/fill/border/alignment オブジェクト生成。デフォルトフォントとの差分計算、デフォルト値フィールド省略、color は HEX RGB に変換。ソリッド塗りつぶしの前景色のみ。6エッジの罫線。alignmentのデフォルト値省略
4. `internal/excel/richtext.go` — リッチテキストラン処理。各ランの font はセルレベル font との差分

**完了条件:** 書式関連のユニットテストが通る。

---

### フェーズ 4: info コマンド

**目的:** 最もシンプルなコマンドを完成させ、E2Eの動作確認を行う。

**実装内容:**

- `info.go` — ファイル引数受け取り、オープン、シート一覧取得、定義名一覧取得（`excelize.GetDefinedName`）、JSON出力。hidden フィールドの omitempty 制御

**完了条件:** 実際の .xlsx ファイルに対して `exceldump info` が正しいJSONを出力する。

**備考:** 書式不要のためフェーズ 2 完了直後に着手可能（フェーズ 3 と並行可）。

---

### フェーズ 5: scan コマンド

**目的:** シート構造分析機能を完成させる。

**実装順序:**

1. `internal/excel/region.go` — 領域分割ロジック。非空セル座標の全取得→行バンド分割（3行以上の空行）→列バンド分割（3列以上の空列）→直積で矩形候補→空矩形除外→走査順ソート
2. `scan.go` — `--sheet` オプション、ファイルオープン、シート解決、`used_range` 取得、`tab_color` 取得、デフォルトフォント検出、デフォルト列幅/行高取得、`col_widths`/`row_heights`（差分のみ）、領域分割、JSON出力。空シート対応

**完了条件:** `exceldump scan` が設計通りのJSON出力を生成する。

---

### フェーズ 6: dump コマンド

**目的:** セルデータのJSONLダンプ機能を完成させる。

**実装内容:**

`dump.go` — 以下の処理を組み立てる:

1. オプション定義（`--sheet`, `--range`, `--start`, `--include-empty`, `--no-style`, `--limit`）
2. ファイルオープン、シート解決
3. 範囲決定（`--range` / `--start` / デフォルト）
4. 結合セル情報の事前取得
5. セル走査ループ（行優先順）:
   - 空セルスキップ（`--include-empty` 時は出力）
   - 結合セルの左上判定
   - セル値・型取得
   - 書式取得（`--no-style` でなければ: font差分, fill, border, alignment, rich_text）
   - ハイパーリンク取得
   - hidden_row / hidden_col 判定
   - display フィールド生成
   - JSONL 1行出力
6. `--limit` カウントと切り捨て通知出力（`next_cell`, `next_range` 計算）

**完了条件:** `exceldump dump` が設計通りのJSONLを出力する。切り捨て通知が正しく機能する。

**備考:** dump のセル走査ロジックは search でも再利用するため、コールバックパターンで共通化する。

---

### フェーズ 7: search コマンド

**目的:** 検索機能を完成させる。

**実装順序:**

1. `internal/excel/filter.go` — `--numeric` 式パーサー（`>100`, `>=100`, `<50`, `100:200`, `=42`、浮動小数点誤差 1e-9）。検索フィルタ（`--query`: display/value の大文字小文字無視部分一致、`--numeric`: number 型のみ・date は対象外、`--type`: 型フィルタ）。AND結合、最低1つ指定必須
2. `search.go` — dump と同様のセル走査ループ＋検索フィルタ。結果0件時の終了コード1。JSONL＋切り捨て通知

**完了条件:** `exceldump search` が設計通りに動作する。終了コード 0/1/2 が正しい。

---

## フェーズ依存関係

```
フェーズ1 (骨格)
    │
    ▼
フェーズ2 (Excel操作基盤)
    │
    ├──────────────────┐
    ▼                  ▼
フェーズ3 (書式処理)   フェーズ4 (info)  ※並行可
    │
    ├─────────┐
    ▼         ▼
フェーズ5   フェーズ6 (dump)  ※並行可
(scan)        │
              ▼
           フェーズ7 (search)
```

## 実装上の注意点

### tint 適用アルゴリズム
Excel のテーマカラー tint は HSL 色空間の明度 (L) に対して適用する。tint > 0 の場合 `L' = L * (1 - tint) + tint`、tint < 0 の場合 `L' = L * (1 + tint)` で計算する。RGB↔HSL 変換が必要。

### デフォルトフォント検出
`used_range` 内の各列に対して `excelize.GetColStyle` で列スタイルを取得し、フォント情報を集計する。列スタイルが未設定（スタイルID 0）の列はカウント対象外とし、全列が未設定の場合は `excelize.GetDefaultFont` にフォールバック。

### dump/search のセル走査の共通化
dump のセル走査ロジックをコールバックパターンで設計し、search から再利用可能にする。「範囲内のセルを行優先順で走査し、各セルに対して処理関数を呼ぶ」共通関数を `internal/excel` パッケージに置く。

### `--range` の列/行のみ指定時の補完
`A:F` 指定時は `used_range` から行の上下限を補完する。空シート（`used_range` が空）の場合は空出力で正常終了。

### JSON の omitempty 制御
Go の `json:",omitempty"` で `false` や `0` が省略されるため、「false時は省略」「0時は省略」のフィールドには `omitempty` が適合する。map 型は nil なら省略される。

### 終了コード制御
cobra の `RunE` が返すエラーの種類に応じて `root.go` の `Execute()` で終了コードを決定する。search の「結果なし」は sentinel error で表現し、終了コード 1 にマッピング。
