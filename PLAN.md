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
│   ├── sheet.go                     # シート情報取得（一覧、種別判定、表示状態、シート解決、使用範囲）
│   ├── sheetmeta.go                 # シートメタ情報（タブ色、デフォルト列幅・行高）
│   ├── cell.go                      # セル値の読み取り・型判定・display生成
│   ├── range.go                     # 範囲パース・走査（A1:H20, A:F, 1:20, 単一セル）、--start処理
│   ├── merge.go                     # 結合セル情報の取得・左上セル判定
│   ├── region.go                    # scan用の領域分割ロジック（行バンド×列バンド→矩形候補）
│   ├── rowcache.go                  # scan用の行キャッシュ（領域分割の高速化）
│   ├── stream.go                    # dump/search用のRows()ストリーミング走査
│   ├── formula.go                   # 数式取得（共有数式、配列数式対応）
│   ├── style.go                     # font/fill/border/alignment オブジェクト生成・差分計算
│   ├── richtext.go                  # リッチテキストラン処理
│   ├── defaultfont.go               # デフォルトフォント検出ロジック（列スタイル頻度カウント）
│   ├── color.go                     # テーマカラー→HEX RGB変換（excelize のGetBaseColor/ThemeColor利用）
│   ├── filter.go                    # search用フィルタ（--query, --numeric, --type）+ numericパーサー
│   └── truncation.go               # 切り捨て通知構造体
└── testdata/                        # テスト用Excelファイル
```

## 実装状況

全フェーズ完了済み。以下は実装時の設計判断の記録。

### ストリーミング処理

dump/search は excelize の `Rows()` イテレータによるストリーミング処理に統一。
- `--limit` 到達時に即座に走査を打ち切る（大規模ファイルでも高速）
- `--range` は走査中にフィルタ、範囲外の行で打ち切り
- `--start` は開始位置まで読み飛ばし
- RowCache は不使用（scan のみ使用）

### dimension の有無による分岐

シートXMLに `<dimension>` 要素がないファイル（Google Sheets 由来等）では全行走査のコストが高い。

**scan:**
- dimension あり: RowCache で全行走査 → used_range, regions, col_widths, row_heights を出力
- dimension なし: 全行走査をスキップ → default_font, default_width, default_height, col_widths のみ出力

**dump/search:**
- dimension の有無に関わらず同一のストリーミング処理。パフォーマンスへの影響なし

行高は dump の行情報（`_row`）で取得できるため、scan で row_heights が省略されても解析精度に差はない。

### 行情報の出力

dump で行が変わるタイミングで `{"_row": N, "height": H}` 形式の行情報を出力。
- 行高（デフォルトと異なる場合）と非表示フラグを含む
- セル出力から `hidden_row` を削除し、行情報に集約

### search のフィルタ拡張

`--numeric` と `--type` は formula 型セルのキャッシュ値も考慮する。
- `--numeric ">100"`: number 型 + formula のキャッシュ値が数値のセルにヒット
- `--type number`: number 型 + formula の数値キャッシュにヒット

### テーマカラー変換

excelize の `GetBaseColor` + `ThemeColor` を利用。自前の HSL 変換は不要。

### 終了コード制御

cobra の `RunE` が返すエラーの種類に応じて `root.go` の `Execute()` で終了コードを決定。
- 0: 成功
- 1: 検索結果なし（search の sentinel error `errNoMatch`）
- 2: エラー
